import assert from "node:assert/strict";
import { test } from "node:test";

import {
  detectSpreadsheetHost,
  resetCurrentSpreadsheetHostForTests,
  resolveSpreadsheetHostForBoot,
} from "../src/host/index.ts";

type GlobalKey = "Office";
type MaybePromise<T> = T | Promise<T>;

function readGlobal(key: string): DynamicValue {
  return Reflect.get(globalThis, key) as DynamicValue;
}

function writeGlobal(key: string, value: DynamicValue): void {
  Reflect.set(globalThis, key, value);
}

function deleteGlobal(key: string): void {
  Reflect.deleteProperty(globalThis, key);
}

async function withGlobal<T>(key: GlobalKey, value: DynamicValue, fn: () => MaybePromise<T>): Promise<T> {
  const hadValue = Reflect.has(globalThis, key);
  const previousValue = readGlobal(key);
  writeGlobal(key, value);

  try {
    return await fn();
  } finally {
    if (hadValue) {
      writeGlobal(key, previousValue);
    } else {
      deleteGlobal(key);
    }
    resetCurrentSpreadsheetHostForTests();
  }
}

async function withoutHostGlobals<T>(fn: () => MaybePromise<T>): Promise<T> {
  const snapshots = ["Office", "wps", "Application"].map((key) => ({
    key,
    hadValue: Reflect.has(globalThis, key),
    previousValue: readGlobal(key),
  }));

  for (const snapshot of snapshots) {
    deleteGlobal(snapshot.key);
  }

  try {
    return await fn();
  } finally {
    for (const snapshot of snapshots) {
      if (snapshot.hadValue) {
        writeGlobal(snapshot.key, snapshot.previousValue);
      } else {
        deleteGlobal(snapshot.key);
      }
    }
    resetCurrentSpreadsheetHostForTests();
  }
}

void test("detectSpreadsheetHost ignores WPS globals and stays browser without Office", async () => {
  await withoutHostGlobals(() => {
    writeGlobal("wps", { Application: {} });
    writeGlobal("Application", {});
    try {
      assert.equal(detectSpreadsheetHost(), "browser");
    } finally {
      deleteGlobal("wps");
      deleteGlobal("Application");
    }
  });
});

void test("detectSpreadsheetHost detects Office and browser hosts", async () => {
  await withoutHostGlobals(async () => {
    assert.equal(detectSpreadsheetHost(), "browser");

    await withGlobal("Office", { onReady: () => Promise.resolve({ host: "Excel", platform: "PC" }) }, () => {
      assert.equal(detectSpreadsheetHost(), "office");
    });
  });
});

void test("resolveSpreadsheetHostForBoot resolves Office when Office.onReady fires", async () => {
  await withoutHostGlobals(async () => {
    const info = { host: "Excel", platform: "PC" };
    const officeGlobal = {
      onReady: (callback: (readyInfo: typeof info) => void) => {
        callback(info);
        return Promise.resolve(info);
      },
    };

    await withGlobal("Office", officeGlobal, async () => {
      const result = await resolveSpreadsheetHostForBoot({ officeReadyTimeoutMs: 50 });
      assert.equal(result.host.kind, "office");
      assert.equal(result.readyInfo.reason, "office-ready");
      assert.equal(result.readyInfo.nativeHost, "Excel");
    });
  });
});

void test("resolveSpreadsheetHostForBoot keeps the Office host after Office.onReady timeout", async () => {
  await withoutHostGlobals(async () => {
    const officeGlobal = {
      onReady: () => new Promise(() => {}),
    };

    await withGlobal("Office", officeGlobal, async () => {
      const result = await resolveSpreadsheetHostForBoot({ officeReadyTimeoutMs: 5 });
      assert.equal(result.host.kind, "office");
      assert.equal(result.readyInfo.reason, "office-timeout");
    });
  });
});

void test("resolveSpreadsheetHostForBoot keeps the Office host when Office.onReady rejects", async () => {
  await withoutHostGlobals(async () => {
    const officeGlobal = {
      onReady: () => Promise.reject(new Error("boom")),
    };

    await withGlobal("Office", officeGlobal, async () => {
      const result = await resolveSpreadsheetHostForBoot({ officeReadyTimeoutMs: 5 });
      assert.equal(result.host.kind, "office");
      assert.equal(result.readyInfo.reason, "office-timeout");
    });
  });
});

void test("resolveSpreadsheetHostForBoot uses the browser host when Office.js is absent", async () => {
  await withoutHostGlobals(async () => {
    const result = await resolveSpreadsheetHostForBoot({ officeReadyTimeoutMs: 5 });
    assert.equal(result.host.kind, "browser");
    assert.equal(result.readyInfo.reason, "office-unavailable");
  });
});
