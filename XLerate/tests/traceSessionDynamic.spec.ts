import { describe, expect, it } from "vitest";
import { TracePortFake } from "../src/adapters/tracePortFake";
import type { TraceCellData } from "../src/core/traceGraph";
import { buildTraceLoadingPolicy } from "../src/core/tracePolicy";
import { parseParentToDialogMessage } from "../src/core/traceProtocol";
import { TraceSessionService } from "../src/services/traceSession.service";

const cell = (key: string, address: string, formula = "", value = "1"): TraceCellData => ({
  kind: "cell",
  key,
  worksheetName: "Calcs",
  rowIndex: 0,
  columnIndex: 0,
  address,
  value,
  formula,
});

function setup(direction: "precedents" | "dependents" = "precedents") {
  const root = cell("root", "Calcs!I153", "=IF(I$3>1,OFFSET(I$140,0,-1),0)", "5294");
  const port = new TracePortFake(root);
  port.setNeighbors("root", [
    cell("i3", "Calcs!I3", "", "2"),
    cell("i140", "Calcs!I140", "", "7467"),
  ]);
  port.setDynamic("root", [
    {
      data: {
        kind: "dynamic",
        key: "root::dyn:0",
        dynamicKind: "reference",
        functionName: "OFFSET",
        label: "OFFSET(I$140,0,-1)",
        worksheetName: "Calcs",
        address: "Calcs!H140",
        value: "5294.12",
        formula: "",
      },
      children: [cell("h140", "Calcs!H140", "=H138*H5", "5294.12")],
    },
  ]);
  const service = new TraceSessionService(port, direction, buildTraceLoadingPolicy());
  return { port, service };
}

describe("trace session with dynamic references", () => {
  it("lists resolved references first, then Excel's direct precedents", async () => {
    const { service } = setup();
    const opened = await service.open();
    expect(opened.children.map((n) => [n.kind, n.address])).toEqual([
      ["dynamic", "Calcs!H140"],
      ["cell", "Calcs!I3"],
      ["cell", "Calcs!I140"],
    ]);
  });

  it("expands a dynamic node into the cells it resolved to", async () => {
    const { service } = setup();
    const opened = await service.open();
    const expansion = await service.expand(opened.children[0].id);
    expect(expansion.children.map((n) => [n.kind, n.address, n.loadState])).toEqual([
      ["cell", "Calcs!H140", "unloaded"],
    ]);
  });

  it("does not resolve dynamic references when tracing dependents", async () => {
    const { service } = setup("dependents");
    const opened = await service.open();
    expect(opened.children.every((n) => n.kind !== "dynamic")).toBe(true);
  });

  it("dynamic nodes survive the parent-to-dialog message check", async () => {
    const { service } = setup();
    const opened = await service.open();
    const parsed = parseParentToDialogMessage(
      JSON.stringify({
        action: "rootLoaded",
        sessionId: "s",
        direction: "precedents",
        root: opened.root,
        children: opened.children,
        truncated: false,
        totalNodeCount: opened.totalNodeCount,
      })
    );
    expect(parsed).not.toBeNull();
  });
});
