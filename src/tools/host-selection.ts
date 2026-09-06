/** Host-specific core tool implementation selection. */

import type { AgentTool } from "@earendil-works/pi-agent-core";
import type { TSchema } from "typebox";

import type { SpreadsheetHostKind } from "../host/index.js";
import { CORE_TOOL_NAMES, type CoreToolName } from "./names.js";

export type AnyHostSelectableTool = AgentTool<TSchema, DynamicValue>;

export function selectCoreToolForHost(
  name: CoreToolName,
  tool: AnyHostSelectableTool,
  _hostKind: SpreadsheetHostKind,
): AnyHostSelectableTool {
  void name;
  return tool;
}

/**
 * Compose the core tool list for a host: `CORE_TOOL_NAMES` drives ordering,
 * `selectCoreToolForHost` applies any host-specific swaps. Kept free of
 * tool-implementation imports so the composition behavior is directly
 * unit-testable.
 */
export function composeCoreToolsForHost(
  createTool: (name: CoreToolName) => AnyHostSelectableTool,
  hostKind: SpreadsheetHostKind,
): AnyHostSelectableTool[] {
  return CORE_TOOL_NAMES.map((name) => selectCoreToolForHost(name, createTool(name), hostKind));
}

/**
 * Non-core tools that drive Office.js/Excel directly. Currently a pass-through
 * because this build only targets Office hosts (plus browser/UI-test fallback).
 */
export function selectOfficeCoupledToolForHost(
  tool: AnyHostSelectableTool,
  _hostKind: SpreadsheetHostKind,
): AnyHostSelectableTool {
  return tool;
}
