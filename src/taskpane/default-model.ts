/**
 * Default model selection for the taskpane.
 *
 * This build prefers DeepSeek, then a configured custom gateway default,
 * then any other still-enabled available provider.
 */

import type { Api, Model, Models } from "@earendil-works/pi-ai";

import { modelRecencyScore } from "../models/model-ordering.js";

function getProviderModels(modelsRuntime: Models, provider: string): Model<Api>[] {
  return [...modelsRuntime.getModels(provider)];
}

function pickLatestProviderModel(modelsRuntime: Models, provider: string): Model<Api> | null {
  const models = getProviderModels(modelsRuntime, provider);
  return (
    models
      .slice()
      .sort((a, b) => {
        const recency = modelRecencyScore(b.id) - modelRecencyScore(a.id);
        if (recency !== 0) return recency;
        return a.id.localeCompare(b.id);
      })[0] ?? null
  );
}

function pickPreferredDeepSeekModel(modelsRuntime: Models): Model<Api> | null {
  const models = getProviderModels(modelsRuntime, "deepseek");
  if (models.length === 0) return null;

  // Prefer Pro-quality ids when recency ties (v4-flash vs v4-pro).
  const proModels = models.filter((model) => /\bpro\b|[-_.]pro(?:$|[-_.])/i.test(model.id));
  const pool = proModels.length > 0 ? proModels : models;
  return (
    pool
      .slice()
      .sort((a, b) => {
        const recency = modelRecencyScore(b.id) - modelRecencyScore(a.id);
        if (recency !== 0) return recency;
        return a.id.localeCompare(b.id);
      })[0] ?? null
  );
}

export function pickDefaultModel(
  modelsRuntime: Models,
  availableProviders: string[],
  customDefaultModel?: Model<Api> | null,
): Model<Api> {
  if (availableProviders.includes("deepseek")) {
    const deepseek = pickPreferredDeepSeekModel(modelsRuntime);
    if (deepseek) return deepseek;
  }

  if (customDefaultModel) {
    return customDefaultModel;
  }

  for (const provider of availableProviders) {
    const best = pickLatestProviderModel(modelsRuntime, provider);
    if (best) return best;
  }

  const preferredFallback = pickPreferredDeepSeekModel(modelsRuntime);
  if (preferredFallback) return preferredFallback;

  const firstKnownModel = modelsRuntime.getModels()[0];
  if (firstKnownModel) return firstKnownModel;
  throw new Error("No models are registered.");
}
