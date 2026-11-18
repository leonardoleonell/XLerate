/* global Excel, Office */

import { TracePortLive, type TracePortMetric } from "../adapters/tracePortLive";
import { buildTraceLoadingPolicy } from "../core/tracePolicy";
import type { TraceTarget } from "../core/traceGraph";
import { parseDialogToParentMessage, type ParentToDialogMessage } from "../core/traceProtocol";
import {
  sanitizeTraceDepth,
  sanitizeTraceSafetyLimit,
  type TraceDirection,
} from "../core/traceUtils";
import { TraceRequestScheduler } from "../services/traceScheduler";
import { TraceSessionService } from "../services/traceSession.service";

type PendingOpen = {
  sessionId: string;
  direction: TraceDirection;
  maxDepth: number;
  safetyLimit: number;
};

let activeDialog: Office.Dialog | null = null;
let activeDialogSessionId: string | null = null;
let activeSession: TraceSessionService | null = null;
let pendingOpen: PendingOpen | null = null;
let sessionSequence = 0;
const traceScheduler = new TraceRequestScheduler();

function isCurrentSession(dialog: Office.Dialog, sessionId: string): boolean {
  return activeDialog === dialog && activeDialogSessionId === sessionId;
}

function logTraceMetric(sessionId: string, metric: TracePortMetric): void {
  // eslint-disable-next-line no-console
  console.log("[trace-metric]", { sessionId, ...metric });
}

function logFocus(step: string, detail: Record<string, unknown> = {}): void {
  // eslint-disable-next-line no-console
  console.log("[trace-focus]", step, {
    at: Math.round(performance.now()),
    paneHasFocus: typeof document.hasFocus === "function" ? document.hasFocus() : "n/a",
    ...detail,
  });
}

/**
 * Ask Excel to give keyboard focus back to the grid. Each call runs in its
 * own batch: when activate/select/focus shared one batch, a single
 * "InvalidOperation" (seen in sideload, 2026-11-18) sank all of them and
 * we could not tell which call Excel refused.
 */
async function pullFocusToGrid(step: string): Promise<void> {
  const supported = Office.context.requirements.isSetSupported("ExcelApiDesktop", "1.1");
  const result: Record<string, unknown> = { workbookFocusApi: supported };

  if (supported) {
    try {
      await Excel.run(async (context) => {
        context.workbook.focus();
        await context.sync();
      });
      result.focus = "ok";
    } catch (error) {
      result.focus = String(error);
    }
  }

  try {
    await Excel.run(async (context) => {
      context.workbook.getActiveCell().select();
      await context.sync();
    });
    result.select = "ok";
  } catch (error) {
    result.select = String(error);
  }

  logFocus(step, result);
}

// Whether the XLerate pane is on screen. The runtime keeps running when the
// user closes the pane (shortcuts still work), but after a dialog closes
// the host hands keyboard focus to that hidden pane and refuses
// workbook.focus() from it, leaving arrow keys dead until a mouse click.
let paneVisible = true;
try {
  Office.onReady(() => {
    try {
      void Office.addin.onVisibilityModeChanged((args) => {
        paneVisible = args.visibilityMode === Office.VisibilityMode.taskpane;
        logFocus("pane visibility", { paneVisible });
      });
    } catch {
      // Older hosts: assume visible, the normal pulls still run.
    }
  });
} catch {
  // Office not present (tests).
}

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

/**
 * With the pane hidden, hiding it again makes the host move focus off the
 * invisible webview. If the pane still holds focus, show and hide it once
 * (a brief flash) as a last resort.
 */
async function releaseFocusFromHiddenPane(): Promise<void> {
  if (paneVisible) return;
  try {
    await Office.addin.hide();
    await sleep(100);
    logFocus("hidden pane: hide()");
    if (document.hasFocus()) {
      await Office.addin.showAsTaskpane();
      await Office.addin.hide();
      await sleep(100);
      logFocus("hidden pane: show()+hide()");
    }
  } catch (error) {
    logFocus("hidden pane: failed", { error: String(error) });
  }
}

// The host returns focus to the add-in webview some time after the dialog
// window closes, and that can land after a single early pull. Releasing the
// webview's own focus and pulling again later keeps arrow keys on the grid.
const FOCUS_PULL_DELAYS_MS = [50, 250, 600, 1200, 2000];

function schedulePullFocusToGrid(): void {
  for (const delay of FOCUS_PULL_DELAYS_MS) {
    setTimeout(() => {
      if (activeDialog !== null) return;
      try {
        (document.activeElement as HTMLElement | null)?.blur();
        window.blur();
      } catch {
        // Best effort only.
      }
      void pullFocusToGrid(`pull+${delay}ms`);
    }, delay);
  }
}

function disposeActiveSession(): void {
  activeSession?.cancel();
  traceScheduler.cancelAll();
  activeSession = null;
}

/**
 * Close the trace dialog. When the user closes it (Enter/Esc), hand focus to
 * the grid BEFORE closing: Windows then returns the keyboard to where Excel
 * already had it instead of to the add-in pane, which left arrow keys stuck
 * until a mouse click. Opening a new trace closes synchronously instead,
 * because Office allows only one dialog at a time.
 */
function closeActiveDialog(options: { focusGridFirst?: boolean } = {}): void {
  disposeActiveSession();
  pendingOpen = null;
  const dialog = activeDialog;
  if (!dialog) return;
  activeDialog = null;
  activeDialogSessionId = null;

  const close = (): void => {
    try {
      dialog.close();
    } catch {
      // The dialog may already be closing through the host UI.
    }
    logFocus("dialog closed", { paneVisible });
    void sleep(50)
      .then(releaseFocusFromHiddenPane)
      .finally(() => schedulePullFocusToGrid());
  };

  if (!options.focusGridFirst) {
    close();
    return;
  }
  void pullFocusToGrid("before close").finally(close);
}

function restoreRootAndClose(dialog: Office.Dialog, sessionId: string): void {
  const session = activeSession;
  if (!session) {
    closeActiveDialog({ focusGridFirst: true });
    return;
  }
  void traceScheduler
    .enqueueNavigation("restore-root", () => session.restoreRootSelection())
    .catch(() => undefined)
    .finally(() => {
      if (isCurrentSession(dialog, sessionId)) closeActiveDialog({ focusGridFirst: true });
    });
}

function sendToDialog(
  dialog: Office.Dialog,
  sessionId: string,
  message: ParentToDialogMessage
): void {
  if (!isCurrentSession(dialog, sessionId)) return;
  try {
    dialog.messageChild(JSON.stringify(message));
  } catch {
    // The dialog closed between the identity check and message send.
  }
}

function sendFailure(
  dialog: Office.Dialog,
  sessionId: string,
  requestId: string,
  nodeId: string,
  error: unknown
): void {
  if (!isCurrentSession(dialog, sessionId)) return;
  const message = error instanceof Error ? error.message : String(error);
  sendToDialog(dialog, sessionId, {
    action: "nodeFailed",
    sessionId,
    requestId,
    nodeId,
    message,
  });
}

function assertDirectionSupported(direction: TraceDirection): void {
  const minimumVersion = direction === "dependents" ? "1.13" : "1.12";
  if (!Office.context.requirements.isSetSupported("ExcelApi", minimumVersion)) {
    throw new Error(
      `Trace ${direction} requires ExcelApi ${minimumVersion} or later on this Excel host.`
    );
  }
}

function traceTargetKey(target: TraceTarget): string {
  if ("address" in target) return target.address;
  if (target.kind === "namedRange") {
    return `name:${target.formulaWorksheetName}:${target.worksheetName ?? "workbook"}:${target.name}`;
  }
  return `table:${target.tableName}:${target.section}:${target.columnStart ?? ""}:${target.columnEnd ?? ""}`;
}

function queuePrefetch(
  dialog: Office.Dialog,
  sessionId: string,
  session: TraceSessionService,
  scheduler: TraceRequestScheduler,
  nodes: Parameters<TraceSessionService["getPrefetchCandidates"]>[0]
): void {
  let remainingPrefetchBudget = session.policy.prefetchMaxCells;
  for (const candidate of session.getPrefetchCandidates(nodes)) {
    void scheduler
      .enqueue("prefetch", candidate.id, async () => {
        if (
          !isCurrentSession(dialog, sessionId) ||
          session.isCancelled() ||
          remainingPrefetchBudget <= 0
        ) {
          return;
        }
        const loaded = await session.prefetch(candidate.id, remainingPrefetchBudget);
        remainingPrefetchBudget = Math.max(0, remainingPrefetchBudget - loaded);
      })
      .catch(() => undefined);
  }
}

function startInitialLoad(request: PendingOpen, dialog: Office.Dialog): void {
  try {
    assertDirectionSupported(request.direction);
  } catch (error) {
    sendFailure(dialog, request.sessionId, "initial", "root", error);
    return;
  }

  const port = new TracePortLive((metric) => logTraceMetric(request.sessionId, metric));
  const policy = buildTraceLoadingPolicy({ safetyLimit: request.safetyLimit });
  const session = new TraceSessionService(port, request.direction, policy, request.maxDepth);
  activeSession = session;

  void traceScheduler
    .enqueue("initial", "root", async () => {
      const result = await session.open();
      if (!isCurrentSession(dialog, request.sessionId) || session.isCancelled()) return;
      sendToDialog(dialog, request.sessionId, {
        action: "rootLoaded",
        sessionId: request.sessionId,
        direction: request.direction,
        root: result.root,
        children: result.children,
        truncated: result.truncated,
        totalNodeCount: result.totalNodeCount,
      });
      queuePrefetch(dialog, request.sessionId, session, traceScheduler, result.children);
    })
    .catch((error) => sendFailure(dialog, request.sessionId, "initial", "root", error));
}

function handleExpand(
  dialog: Office.Dialog,
  sessionId: string,
  requestId: string,
  nodeId: string
): void {
  const session = activeSession;
  if (!session) return;

  void traceScheduler
    .enqueue("expansion", nodeId, async () => {
      const result = await session.expand(nodeId);
      if (!isCurrentSession(dialog, sessionId) || session.isCancelled()) return;
      sendToDialog(dialog, sessionId, {
        action: "childrenLoaded",
        sessionId,
        requestId,
        ...result,
      });
      queuePrefetch(dialog, sessionId, session, traceScheduler, result.children);
    })
    .catch((error) => sendFailure(dialog, sessionId, requestId, nodeId, error));
}

function handleRangePage(
  dialog: Office.Dialog,
  sessionId: string,
  requestId: string,
  nodeId: string,
  page: number
): void {
  const session = activeSession;
  if (!session) return;

  void traceScheduler
    .enqueue("expansion", `${nodeId}:page:${page}`, async () => {
      const result = await session.loadRangePage(nodeId, page);
      if (!isCurrentSession(dialog, sessionId) || session.isCancelled()) return;
      sendToDialog(dialog, sessionId, {
        action: "rangePageLoaded",
        sessionId,
        requestId,
        ...result,
      });
    })
    .catch((error) => sendFailure(dialog, sessionId, requestId, nodeId, error));
}

function handleDialogMessage(
  dialog: Office.Dialog,
  sessionId: string,
  arg: { message?: string; origin?: string | undefined } | { error: number }
): void {
  if (!("message" in arg) || typeof arg.message !== "string") return;
  const message = parseDialogToParentMessage(arg.message);
  if (!message || message.sessionId !== sessionId || !isCurrentSession(dialog, sessionId)) {
    return;
  }

  if (message.action === "ready") {
    if (pendingOpen?.sessionId === sessionId) {
      const request = pendingOpen;
      pendingOpen = null;
      startInitialLoad(request, dialog);
    }
    return;
  }
  if (message.action === "expand") {
    handleExpand(dialog, sessionId, message.requestId, message.nodeId);
    return;
  }
  if (message.action === "loadRangePage") {
    handleRangePage(dialog, sessionId, message.requestId, message.nodeId, message.page);
    return;
  }
  if (message.action === "navigate") {
    const session = activeSession;
    if (session) {
      void traceScheduler
        .enqueueNavigation(traceTargetKey(message.target), () => session.navigate(message.target))
        .catch(() => undefined);
    }
    return;
  }
  if (message.restoreRoot === true) restoreRootAndClose(dialog, sessionId);
  else closeActiveDialog({ focusGridFirst: true });
}

export type OpenTraceDialogOptions = {
  maxDepth?: number;
  safetyLimit?: number;
  height?: number;
  width?: number;
};

export async function openTraceDialog(
  direction: TraceDirection,
  options: OpenTraceDialogOptions = {}
): Promise<void> {
  closeActiveDialog();
  const sessionId = `trace-${Date.now()}-${++sessionSequence}`;
  pendingOpen = {
    sessionId,
    direction,
    maxDepth: sanitizeTraceDepth(options.maxDepth),
    safetyLimit: sanitizeTraceSafetyLimit(options.safetyLimit),
  };

  const url = new URL("traceDialog.html", window.location.href);
  url.searchParams.set("direction", direction);
  url.searchParams.set("sessionId", sessionId);
  const height = typeof options.height === "number" ? options.height : 30;
  const width = typeof options.width === "number" ? options.width : 20;

  return new Promise<void>((resolve) => {
    Office.context.ui.displayDialogAsync(
      url.toString(),
      { height, width, displayInIframe: true },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          pendingOpen = null;
          resolve();
          return;
        }
        const dialog = result.value;
        activeDialog = dialog;
        activeDialogSessionId = sessionId;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) =>
          handleDialogMessage(dialog, sessionId, arg)
        );
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          if (!isCurrentSession(dialog, sessionId)) return;
          disposeActiveSession();
          pendingOpen = null;
          activeDialog = null;
          activeDialogSessionId = null;
          schedulePullFocusToGrid();
        });
        resolve();
      }
    );
  });
}
