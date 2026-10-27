// Minimal entrypoint: resolve the profile and show what it maps to.
// The full `ingest -> review -> approve` loop lives in src/workflow.ts and is wired next.

import { resolveProfile } from './model.js';

const provider = resolveProfile(process.env.LLM_PROFILE);
console.log(
  `sheetflow: profile=${provider.profile} model=${provider.model} fallbacks=[${provider.fallbacks.join(', ') || 'none'}]`,
);
