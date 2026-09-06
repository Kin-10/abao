/** Host global detection for taskpane boot. */

import type { SpreadsheetHostKind } from "./types.js";

function getGlobalMember(scope: object, key: string): DynamicValue {
  return Reflect.get(scope, key);
}

function hasObjectOrFunction(value: DynamicValue): boolean {
  return (typeof value === "object" && value !== null) || typeof value === "function";
}

export function hasOfficeJsGlobal(scope: object = globalThis): boolean {
  return hasObjectOrFunction(getGlobalMember(scope, "Office"));
}

/**
 * Synchronous host detection from globals. Boot resolution may still downgrade
 * `office` to `browser` if Office.onReady never fires within the UI-test timeout.
 */
export function detectSpreadsheetHost(scope: object = globalThis): SpreadsheetHostKind {
  if (hasOfficeJsGlobal(scope)) return "office";
  return "browser";
}
