import type {
  DirectNeighborResult,
  TraceCellData,
  TraceDynamicResult,
  TraceRangeData,
  TraceRangePageResult,
  TraceTarget,
} from "../core/traceGraph";
import type { TraceLoadingPolicy } from "../core/tracePolicy";
import type { TraceDirection } from "../core/traceUtils";

export type TraceLoadRequest = {
  remainingNodes: number;
  policy: TraceLoadingPolicy;
};

export interface TracePort {
  getRoot(): Promise<TraceCellData>;
  getDirectNeighbors(
    node: TraceCellData,
    direction: TraceDirection,
    request: TraceLoadRequest
  ): Promise<DirectNeighborResult>;
  getRangePage(
    node: TraceRangeData,
    page: number,
    request: TraceLoadRequest
  ): Promise<TraceRangePageResult>;
  select(target: TraceTarget): Promise<void>;
  /** Resolve OFFSET/INDEX/SUMIF-style references in a cell's formula. */
  resolveDynamic(node: TraceCellData, request: TraceLoadRequest): Promise<TraceDynamicResult[]>;
  /** Mark the traced cell on the sheet while the dialog is open. */
  highlightRoot(node: TraceCellData): Promise<void>;
  /** Put back the traced cell's original fill. */
  clearRootHighlight(): Promise<void>;
}
