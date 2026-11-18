import type {
  DirectNeighborResult,
  TraceCellData,
  TraceDynamicResult,
  TraceNodeData,
  TraceRangeData,
  TraceRangePageResult,
  TraceTableSection,
  TraceTableTarget,
  TraceTarget,
} from "../core/traceGraph";
import {
  materializeTraceArea,
  materializeTraceCells,
  type TraceAreaSnapshot,
} from "../core/traceAreaMaterializer";
import { getRangePageBounds } from "../core/tracePolicy";
import { planTraceRangePageBlocks } from "../core/traceRangePaging";
import {
  buildTraceCellKey,
  formatTraceFormula,
  formatTraceValue,
  parseWorksheetScopedAddress,
  scalarFromMatrix,
  type TraceDirection,
} from "../core/traceUtils";
import type { TraceLoadRequest, TracePort } from "./tracePort";
import { isStillPainted, planRestoreFill, type OriginalFill } from "../core/formulaMap";
import {
  quoteSheet,
  rectA1,
  type RefValue,
  type Scalar,
  type UsedBounds,
} from "../core/formulaEval";
import {
  areasAddress,
  resolveDynamicReferences,
  type CellRect,
  type DynamicDataProvider,
} from "../core/traceDynamic";

const ROOT_MARK_COLOR = "#FF99FF";
const MAX_DYNAMIC_CELLS = 500;

/** Reads what the dynamic-reference resolver asks for, one sync per batch. */
function excelDynamicDataProvider(): DynamicDataProvider {
  return {
    async readRects(rects: CellRect[]): Promise<unknown[][][]> {
      return Excel.run(async (context) => {
        const ranges = rects.map((rect) => {
          const range = context.workbook.worksheets
            .getItem(rect.sheet)
            .getRangeByIndexes(rect.r1, rect.c1, rect.r2 - rect.r1 + 1, rect.c2 - rect.c1 + 1);
          range.load("values");
          return range;
        });
        await context.sync();
        return ranges.map((range) => range.values as unknown[][]);
      });
    },
    async readUsedRanges(sheets: string[]): Promise<UsedBounds[]> {
      return Excel.run(async (context) => {
        const used = sheets.map((sheet) => {
          const range = context.workbook.worksheets.getItem(sheet).getUsedRangeOrNullObject(true);
          range.load(["isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
          return range;
        });
        await context.sync();
        return used.map((range) =>
          range.isNullObject
            ? null
            : {
                r1: range.rowIndex,
                c1: range.columnIndex,
                r2: range.rowIndex + range.rowCount - 1,
                c2: range.columnIndex + range.columnCount - 1,
              }
        );
      });
    },
    async resolveNames(names, formulaSheet): Promise<Array<RefValue | Scalar>> {
      return Excel.run(async (context) => {
        const lookups = names.map((entry) => {
          const local = context.workbook.worksheets
            .getItem(entry.sheet ?? formulaSheet)
            .names.getItemOrNullObject(entry.name);
          const global = entry.sheet
            ? null
            : context.workbook.names.getItemOrNullObject(entry.name);
          local.load(["type", "value"]);
          global?.load(["type", "value"]);
          return { local, global };
        });
        await context.sync();
        const ranges = lookups.map(({ local, global }) => {
          const item = !local.isNullObject ? local : global && !global.isNullObject ? global : null;
          if (!item) return { item: null, range: null };
          if (item.type !== "Range") return { item, range: null };
          const range = item.getRangeOrNullObject();
          range.load([
            "isNullObject",
            "rowIndex",
            "columnIndex",
            "rowCount",
            "columnCount",
            "worksheet/name",
          ]);
          return { item, range };
        });
        await context.sync();
        return ranges.map(({ item, range }): RefValue | Scalar => {
          if (!item) return { error: "#NAME?" };
          if (range && !range.isNullObject) {
            return {
              kind: "ref",
              sheet: range.worksheet.name,
              r1: range.rowIndex,
              c1: range.columnIndex,
              r2: range.rowIndex + range.rowCount - 1,
              c2: range.columnIndex + range.columnCount - 1,
            };
          }
          const raw = item.value as unknown;
          if (typeof raw === "number" || typeof raw === "boolean") return raw;
          if (typeof raw === "string") {
            const text = raw.replace(/^=/, "");
            const n = Number(text);
            if (text.trim() !== "" && !Number.isNaN(n)) return n;
            return /^".*"$/.test(text) ? text.slice(1, -1) : { error: "#NAME?" };
          }
          return { error: "#NAME?" };
        });
      });
    },
  };
}

const RANGE_COLLECTION_LOAD_SPEC =
  "items/address,rowIndex,columnIndex,rowCount,columnCount,worksheet/name";

export type TracePortMetric = {
  operation: "root" | "neighbors" | "rangePage" | "select";
  syncCount: number;
  returnedAreaCount?: number;
  returnedNodeCount?: number;
  materializedCellCount?: number;
};

export type TracePortMetricSink = (metric: TracePortMetric) => void;

function isItemNotFoundError(error: unknown): boolean {
  if (!error || typeof error !== "object") return false;
  const candidate = error as { code?: unknown; message?: unknown };
  if (candidate.code === "ItemNotFound") return true;
  return typeof candidate.message === "string" && candidate.message.includes("ItemNotFound");
}

function snapshotCell(
  worksheetName: string,
  rowIndex: number,
  columnIndex: number,
  address: string,
  values: unknown,
  formulas: unknown
): TraceCellData {
  return {
    kind: "cell",
    key: buildTraceCellKey(worksheetName, rowIndex, columnIndex),
    worksheetName,
    rowIndex,
    columnIndex,
    address,
    value: formatTraceValue(scalarFromMatrix(values)),
    formula: formatTraceFormula(scalarFromMatrix(formulas)),
  };
}

function getTableRangeForSection(table: Excel.Table, section: TraceTableSection): Excel.Range {
  if (section === "all") return table.getRange();
  if (section === "headers") return table.getHeaderRowRange();
  if (section === "totals") return table.getTotalRowRange();
  return table.getDataBodyRange();
}

function getTableColumnRangeForSection(
  column: Excel.TableColumn,
  section: TraceTableSection
): Excel.Range {
  if (section === "all") return column.getRange();
  if (section === "headers") return column.getHeaderRowRange();
  if (section === "totals") return column.getTotalRowRange();
  return column.getDataBodyRange();
}

function getStructuredTableRange(
  context: Excel.RequestContext,
  target: TraceTableTarget
): Excel.Range {
  const table = context.workbook.tables.getItem(target.tableName);
  if (!target.columnStart) return getTableRangeForSection(table, target.section);

  const start = getTableColumnRangeForSection(
    table.columns.getItem(target.columnStart),
    target.section
  );
  if (!target.columnEnd) return start;
  const end = getTableColumnRangeForSection(
    table.columns.getItem(target.columnEnd),
    target.section
  );
  return start.getBoundingRect(end);
}

export class TracePortLive implements TracePort {
  constructor(private readonly onMetric: TracePortMetricSink = () => undefined) {}

  async getRoot(): Promise<TraceCellData> {
    let syncCount = 0;
    const result = await Excel.run(async (context) => {
      const cell = context.workbook.getActiveCell();
      cell.load(["address", "worksheet/name", "rowIndex", "columnIndex", "values", "formulas"]);
      await context.sync();
      syncCount += 1;
      return snapshotCell(
        cell.worksheet.name,
        cell.rowIndex,
        cell.columnIndex,
        cell.address,
        cell.values,
        cell.formulas
      );
    });
    this.onMetric({ operation: "root", syncCount, returnedNodeCount: 1 });
    return result;
  }

  async getDirectNeighbors(
    node: TraceCellData,
    direction: TraceDirection,
    request: TraceLoadRequest
  ): Promise<DirectNeighborResult> {
    let syncCount = 0;
    const result = await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(node.worksheetName);
      const source = worksheet.getRangeByIndexes(node.rowIndex, node.columnIndex, 1, 1);
      const links =
        direction === "precedents" ? source.getDirectPrecedents() : source.getDirectDependents();
      const ranges = links.ranges;
      ranges.load(RANGE_COLLECTION_LOAD_SPEC);

      try {
        await context.sync();
        syncCount += 1;
      } catch (error) {
        syncCount += 1;
        if (isItemNotFoundError(error)) {
          return {
            nodes: [],
            truncated: false,
            materializedCellCount: 0,
            areaCount: 0,
          };
        }
        throw error;
      }

      let remainingNodes = request.remainingNodes;
      const smallAreas: Excel.Range[] = [];
      for (const area of ranges.items) {
        const cellCount = area.rowCount * area.columnCount;
        if (cellCount <= request.policy.rangeSummaryThreshold && remainingNodes > 0) {
          area.load(["values", "formulas"]);
          smallAreas.push(area);
        }
        remainingNodes -=
          cellCount > request.policy.rangeSummaryThreshold
            ? 1
            : Math.min(cellCount, Math.max(remainingNodes, 0));
        remainingNodes = Math.max(0, remainingNodes);
      }

      if (smallAreas.length > 0) {
        await context.sync();
        syncCount += 1;
      }

      const nodes: DirectNeighborResult["nodes"] = [];
      let truncated = false;
      let materializedCellCount = 0;
      remainingNodes = request.remainingNodes;
      for (const area of ranges.items) {
        if (remainingNodes <= 0) {
          truncated = true;
          break;
        }
        const snapshot: TraceAreaSnapshot = {
          worksheetName: area.worksheet.name,
          address: area.address,
          rowIndex: area.rowIndex,
          columnIndex: area.columnIndex,
          rowCount: area.rowCount,
          columnCount: area.columnCount,
          values:
            area.rowCount * area.columnCount <= request.policy.rangeSummaryThreshold
              ? (area.values as unknown[][])
              : undefined,
          formulas:
            area.rowCount * area.columnCount <= request.policy.rangeSummaryThreshold
              ? (area.formulas as unknown[][])
              : undefined,
        };
        const materialized = materializeTraceArea(snapshot, request.policy, remainingNodes);
        nodes.push(...materialized.nodes);
        remainingNodes -= materialized.nodes.length;
        materializedCellCount += materialized.materializedCellCount;
        truncated = truncated || materialized.truncated;
      }

      return {
        nodes,
        truncated,
        materializedCellCount,
        areaCount: ranges.items.length,
      };
    });

    const measured = result as DirectNeighborResult & { areaCount?: number };
    this.onMetric({
      operation: "neighbors",
      syncCount,
      returnedAreaCount: measured.areaCount,
      returnedNodeCount: measured.nodes.length,
      materializedCellCount: measured.materializedCellCount,
    });
    return {
      nodes: measured.nodes,
      truncated: measured.truncated,
      materializedCellCount: measured.materializedCellCount,
    };
  }

  async getRangePage(
    node: TraceRangeData,
    page: number,
    request: TraceLoadRequest
  ): Promise<TraceRangePageResult> {
    const bounds = getRangePageBounds(node.cellCount, page, request.policy);
    const count = Math.min(bounds.count, request.remainingNodes);
    if (count === 0) {
      return {
        nodes: [],
        page,
        hasMore: bounds.hasMore,
        truncated: bounds.count > 0,
      };
    }

    const blocks = planTraceRangePageBlocks(node.rowCount, node.columnCount, bounds.start, count);

    let syncCount = 0;
    const result = await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(node.worksheetName);
      const ranges = blocks.map((block) => {
        const range = worksheet.getRangeByIndexes(
          node.rowIndex + block.rowOffset,
          node.columnIndex + block.columnOffset,
          block.rowCount,
          block.columnCount
        );
        range.load([
          "address",
          "rowIndex",
          "columnIndex",
          "rowCount",
          "columnCount",
          "values",
          "formulas",
        ]);
        return range;
      });
      if (ranges.length === 0) {
        return [];
      }
      await context.sync();
      syncCount += 1;

      const nodes: TraceCellData[] = [];
      let remainingNodes = count;
      for (const range of ranges) {
        const materialized = materializeTraceCells(
          {
            worksheetName: node.worksheetName,
            address: range.address,
            rowIndex: range.rowIndex,
            columnIndex: range.columnIndex,
            rowCount: range.rowCount,
            columnCount: range.columnCount,
            values: range.values as unknown[][],
            formulas: range.formulas as unknown[][],
          },
          remainingNodes
        );
        nodes.push(...materialized.nodes);
        remainingNodes -= materialized.nodes.length;
      }
      return nodes;
    });

    this.onMetric({
      operation: "rangePage",
      syncCount,
      returnedAreaCount: blocks.length,
      returnedNodeCount: result.length,
      materializedCellCount: result.length,
    });
    return {
      nodes: result,
      page,
      hasMore: bounds.hasMore || count < bounds.count,
      truncated: count < bounds.count,
    };
  }

  async resolveDynamic(
    node: TraceCellData,
    request: TraceLoadRequest
  ): Promise<TraceDynamicResult[]> {
    const row = node.rowIndex;
    const col = node.columnIndex;
    const items = await resolveDynamicReferences(
      node.formula,
      { sheet: node.worksheetName, row, col },
      excelDynamicDataProvider()
    );
    if (items.length === 0) return [];

    return Excel.run(async (context) => {
      let remaining = Math.max(0, request.remainingNodes);
      const planned = items.map((item, index) => {
        const worksheet = context.workbook.worksheets.getItem(item.sheet);
        const areas: Array<{ area: (typeof item.areas)[number]; range?: Excel.Range }> = [];
        let cellsLeft = Math.min(remaining, MAX_DYNAMIC_CELLS);
        for (const area of item.areas) {
          const count = (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1);
          const small = count <= request.policy.rangeSummaryThreshold && cellsLeft > 0;
          if (small) {
            const range = worksheet.getRangeByIndexes(
              area.r1,
              area.c1,
              area.r2 - area.r1 + 1,
              area.c2 - area.c1 + 1
            );
            range.load(["values", "formulas"]);
            areas.push({ area, range });
            cellsLeft -= count;
          } else {
            areas.push({ area });
          }
        }
        return { item, index, areas };
      });
      await context.sync();

      return planned.map(({ item, index, areas }): TraceDynamicResult => {
        const children: TraceNodeData[] = [];
        for (const { area, range } of areas) {
          if (remaining <= 0) break;
          const materialized = materializeTraceArea(
            {
              worksheetName: item.sheet,
              address: `${quoteSheet(item.sheet)}!${rectA1(area)}`,
              rowIndex: area.r1,
              columnIndex: area.c1,
              rowCount: area.r2 - area.r1 + 1,
              columnCount: area.c2 - area.c1 + 1,
              values: range ? (range.values as unknown[][]) : undefined,
              formulas: range ? (range.formulas as unknown[][]) : undefined,
            },
            request.policy,
            remaining
          );
          children.push(...materialized.nodes);
          remaining -= materialized.nodes.length;
        }
        // Multi-area selection addresses are capped (~255 chars) by Excel;
        // keep as many areas as fit so selecting the node never fails.
        let fitting = item.areas.length;
        while (fitting > 1 && areasAddress(item.sheet, item.areas.slice(0, fitting)).length > 250) {
          fitting -= 1;
        }
        const address = areasAddress(item.sheet, item.areas.slice(0, fitting));
        const first = children[0];
        const value =
          item.kind === "filter"
            ? `${item.cellCount} de ${item.candidateCount ?? item.cellCount} células`
            : item.cellCount === 1 && first?.kind === "cell"
              ? first.value
              : `${item.cellCount} células`;
        return {
          data: {
            kind: "dynamic",
            key: `${node.key}::dyn:${index}`,
            dynamicKind: item.kind,
            functionName: item.functionName,
            label: item.label,
            worksheetName: item.sheet,
            address,
            value,
            formula: "",
          },
          children,
        };
      });
    });
  }

  private rootMark: { sheet: string; row: number; col: number; original: OriginalFill } | null =
    null;

  async highlightRoot(node: TraceCellData): Promise<void> {
    await Excel.run(async (context) => {
      const cell = context.workbook.worksheets
        .getItem(node.worksheetName)
        .getRangeByIndexes(node.rowIndex, node.columnIndex, 1, 1);
      const props = cell.getCellProperties({
        format: { fill: { color: true, pattern: true, patternColor: true } },
      });
      await context.sync();
      const fill = props.value[0]?.[0]?.format?.fill;
      this.rootMark = {
        sheet: node.worksheetName,
        row: node.rowIndex,
        col: node.columnIndex,
        original: {
          color: fill?.color ?? null,
          pattern: (fill?.pattern as string | undefined) ?? null,
          patternColor: fill?.patternColor ?? null,
        },
      };
      // Pattern before colour, or the fill may not render (CLAUDE.md).
      cell.format.fill.pattern = "Solid";
      cell.format.fill.color = ROOT_MARK_COLOR;
      await context.sync();
    });
  }

  async clearRootHighlight(): Promise<void> {
    const mark = this.rootMark;
    if (!mark) return;
    this.rootMark = null;
    await Excel.run(async (context) => {
      const cell = context.workbook.worksheets
        .getItem(mark.sheet)
        .getRangeByIndexes(mark.row, mark.col, 1, 1);
      const props = cell.getCellProperties({ format: { fill: { color: true } } });
      await context.sync();
      // Leave it alone if the user recoloured the cell meanwhile.
      if (
        !isStillPainted(
          { color: props.value[0]?.[0]?.format?.fill?.color ?? null },
          ROOT_MARK_COLOR
        )
      )
        return;
      const plan = planRestoreFill(mark.original);
      if (plan.kind === "clear") {
        cell.format.fill.clear();
      } else {
        cell.format.fill.pattern = plan.fill.pattern as Excel.FillPattern;
        cell.format.fill.color = plan.fill.color;
        if (plan.fill.patternColor) cell.format.fill.patternColor = plan.fill.patternColor;
      }
      await context.sync();
    });
  }

  async select(target: TraceTarget): Promise<void> {
    let syncCount = 0;
    await Excel.run(async (context) => {
      if ("address" in target) {
        const parsed = parseWorksheetScopedAddress(target.address);
        if (!parsed) return;
        const worksheet = context.workbook.worksheets.getItem(parsed.worksheetName);
        worksheet.getRanges(parsed.rangeAddress).select();
      } else if (target.kind === "namedRange") {
        if (target.worksheetName) {
          const worksheet = context.workbook.worksheets.getItem(target.worksheetName);
          worksheet.names.getItem(target.name).getRange().select();
        } else {
          const formulaWorksheet = context.workbook.worksheets.getItem(target.formulaWorksheetName);
          const localName = formulaWorksheet.names.getItemOrNullObject(target.name);
          const workbookName = context.workbook.names.getItemOrNullObject(target.name);
          localName.load("name");
          workbookName.load("name");
          await context.sync();
          syncCount += 1;

          const resolvedName = localName.isNullObject ? workbookName : localName;
          if (resolvedName.isNullObject) {
            throw new Error(`Named range "${target.name}" was not found.`);
          }
          resolvedName.getRange().select();
        }
      } else {
        getStructuredTableRange(context, target).select();
      }
      await context.sync();
      syncCount += 1;
    });
    this.onMetric({ operation: "select", syncCount });
  }
}
