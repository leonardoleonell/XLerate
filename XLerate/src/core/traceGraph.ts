export type TraceNodeLoadState = "unloaded" | "loading" | "loaded" | "leaf" | "error";

export type TraceCellData = {
  kind: "cell";
  /** Stable workbook-session identity, normally worksheet + row + column. */
  key: string;
  worksheetName: string;
  rowIndex: number;
  columnIndex: number;
  address: string;
  value: string;
  formula: string;
};

export type TraceRangeData = {
  kind: "range";
  /** Stable identity for the rectangular block. */
  key: string;
  worksheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  cellCount: number;
  address: string;
};

export type TraceNodeData = TraceCellData | TraceRangeData;

type TraceNodePlacement = {
  /** Tree-instance identity. References to the same cell get distinct IDs. */
  id: string;
  parentId: string | null;
  level: number;
  loadState: TraceNodeLoadState;
};

export type TraceCellNode = TraceCellData & TraceNodePlacement;
export type TraceRangeNode = TraceRangeData & TraceNodePlacement;

export type TraceReferenceNode = TraceNodePlacement & {
  kind: "reference";
  key: string;
  targetId: string;
  worksheetName: string;
  address: string;
  value: string;
  formula: string;
};

/**
 * A dynamic reference resolved from the parent's formula: OFFSET/INDEX/
 * INDIRECT/lookups (the cell actually used) or a SUMIF-style filter (the
 * cells that pass the criteria). `address` may hold several areas.
 */
export type TraceDynamicData = {
  kind: "dynamic";
  key: string;
  dynamicKind: "reference" | "filter";
  functionName: string;
  /** Source text of the call, e.g. OFFSET(I$140,0,-1). */
  label: string;
  worksheetName: string;
  /** Sheet-qualified target, e.g. Calcs!H140 or Calcs!I17,I19,I21. */
  address: string;
  /** Short summary shown in the value column. */
  value: string;
  formula: string;
};

export type TraceDynamicResult = {
  data: TraceDynamicData;
  /** Resolved cells (or a summarized range) shown when the node expands. */
  children: TraceNodeData[];
};

export type TraceDynamicNode = TraceDynamicData & TraceNodePlacement;

export type TraceNode = TraceCellNode | TraceRangeNode | TraceReferenceNode | TraceDynamicNode;

export type TraceAddressTarget = {
  kind?: "address";
  address: string;
};

export type TraceNamedRangeTarget = {
  kind: "namedRange";
  name: string;
  /** Worksheet containing the formula, used for Excel's local-name precedence. */
  formulaWorksheetName: string;
  /** Present only when the formula explicitly qualifies the name as Sheet!Name. */
  worksheetName?: string;
};

export type TraceTableSection = "all" | "data" | "headers" | "totals";

export type TraceTableTarget = {
  kind: "table";
  tableName: string;
  section: TraceTableSection;
  columnStart?: string;
  columnEnd?: string;
};

export type TraceTarget = TraceAddressTarget | TraceNamedRangeTarget | TraceTableTarget;

export type DirectNeighborResult = {
  nodes: TraceNodeData[];
  truncated: boolean;
  materializedCellCount: number;
};

export type TraceRangePageResult = {
  nodes: TraceCellData[];
  page: number;
  hasMore: boolean;
  truncated: boolean;
};

export function buildTraceRangeKey(
  worksheetName: string,
  rowIndex: number,
  columnIndex: number,
  rowCount: number,
  columnCount: number
): string {
  return `${worksheetName}!R${rowIndex}C${columnIndex}:${rowCount}x${columnCount}`;
}

export function isTraceNodeExpandable(node: TraceNode): boolean {
  return node.kind !== "reference" && node.loadState !== "leaf";
}
