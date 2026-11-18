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

// The host returns focus to the add-in webview some time after the dialog
// window closes, and that can land after a single early pull, so pull again
// later. (These pulls succeed when the pane is on screen.)
const FOCUS_PULL_DELAYS_MS = [50, 250, 600, 1200, 2000];

function schedulePullFocusToGrid(): void {
  for (const delay of FOCUS_PULL_DELAYS_MS) {
    setTimeout(() => {
      if (activeDialog !== null) return;
      void pullFocusToGrid(`pull+${delay}ms`);
    }, delay);
  }
}

/**
 * Key rescue. With the pane closed, the keyboard lands in the hidden pane
 * after the trace dialog closes and Excel refuses workbook.focus() from
 * there (user-confirmed, 2026-11-18). The keys still reach this document,
 * so after a trace we act on them ourselves: arrows/Enter/Tab move the
 * selection like the grid would, and each key retries workbook.focus().
 * Rescue ends when the pane loses focus (the grid has it again) or the
 * user starts typing in one of the pane's own fields.
 */
let keyRescueArmed = false;
let keyRescueQueue: Promise<void> = Promise.resolve();

type Move = { rows: number; cols: number; edge?: Excel.KeyboardDirection };

function moveForKey(event: KeyboardEvent): Move | null {
  const ctrl = event.ctrlKey;
  switch (event.key) {
    case "ArrowDown":
      return ctrl ? { rows: 0, cols: 0, edge: Excel.KeyboardDirection.down } : { rows: 1, cols: 0 };
    case "ArrowUp":
      return ctrl ? { rows: 0, cols: 0, edge: Excel.KeyboardDirection.up } : { rows: -1, cols: 0 };
    case "ArrowRight":
      return ctrl
        ? { rows: 0, cols: 0, edge: Excel.KeyboardDirection.right }
        : { rows: 0, cols: 1 };
    case "ArrowLeft":
      return ctrl
        ? { rows: 0, cols: 0, edge: Excel.KeyboardDirection.left }
        : { rows: 0, cols: -1 };
    case "Enter":
      return { rows: event.shiftKey ? -1 : 1, cols: 0 };
    case "Tab":
      return { rows: 0, cols: event.shiftKey ? -1 : 1 };
    default:
      return null;
  }
}

async function rescueKey(move: Move | null): Promise<void> {
  const result: Record<string, unknown> = {};
  if (move) {
    try {
      await Excel.run(async (context) => {
        const cell = context.workbook.getActiveCell();
        if (move.edge) {
          context.workbook.getSelectedRange().getRangeEdge(move.edge, cell).select();
        } else {
          cell.load(["rowIndex", "columnIndex"]);
          await context.sync();
          const row = Math.min(1048575, Math.max(0, cell.rowIndex + move.rows));
          const col = Math.min(16383, Math.max(0, cell.columnIndex + move.cols));
          cell.getOffsetRange(row - cell.rowIndex, col - cell.columnIndex).select();
        }
        await context.sync();
      });
      result.move = "ok";
    } catch (error) {
      result.move = String(error);
    }
  }
  await pullFocusToGrid("key rescue");
  logFocus("key rescue", result);
}

function handleRescueKeydown(event: KeyboardEvent): void {
  if (!keyRescueArmed || activeDialog !== null) return;
  const target = event.target as HTMLElement | null;
  if (target && /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName)) {
    keyRescueArmed = false;
    return;
  }
  if (["Shift", "Control", "Alt", "Meta"].includes(event.key)) return;
  // Shift+arrow (extend selection) is not emulated; just hand focus back.
  const move = event.shiftKey && event.key.startsWith("Arrow") ? null : moveForKey(event);
  event.preventDefault();
  keyRescueQueue = keyRescueQueue.then(() => rescueKey(move)).catch(() => undefined);
}

try {
  document.addEventListener("keydown", handleRescueKeydown, true);
  // A click inside the pane means the user is working in the pane itself.
  document.addEventListener("mousedown", () => (keyRescueArmed = false), true);
  // Losing focus means the grid (or another window) has the keyboard again.
  window.addEventListener("blur", () => {
    if (keyRescueArmed) logFocus("key rescue off (pane lost focus)");
    keyRescueArmed = false;
  });
} catch {
  // No DOM (tests).
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
    keyRescueArmed = true;
    logFocus("dialog closed", { paneVisible });
    schedulePullFocusToGrid();
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
