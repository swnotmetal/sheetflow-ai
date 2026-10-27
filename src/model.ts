// One prompt, one schema, one acceptance rule — for every provider.
//
// sheetflow is framework-agnostic about the *model*: the workflow (Mastra) never knows which
// provider is behind the call. It only ever sees a typed Extraction and the grounding guarantee
// below. This mirrors `src/llm.ts` in langorder-ai, but the provider is a user-facing profile
// ("what cloud / whose key do you have") rather than a hard-coded vendor.

import { z } from 'zod';

export const FACT_KEYS = ['customer', 'sender', 'location', 'po', 'date', 'address'] as const;
export type FactKey = (typeof FACT_KEYS)[number];

// Plain string arrays, one entry per occurrence. No nested "evidence spans" — a small local model
// is never asked for character-exact citations; grounding happens in `groundFacts`, not in the prompt.
export const EXTRACT_SCHEMA = z.object({
  customer: z.array(z.string()).optional(),
  sender: z.array(z.string()).optional(),
  location: z.array(z.string()).optional(),
  po: z.array(z.string()).optional(),
  date: z.array(z.string()).optional(),
  address: z.array(z.string()).optional(),
  lines: z.array(
    z.object({
      description: z.string(),
      quantity: z.string().nullable().optional(),
      unit: z.string().nullable().optional(),
    }),
  ),
});

export type Extraction = z.infer<typeof EXTRACT_SCHEMA>;

export const EXTRACT_SYSTEM = [
  'You extract order facts from an email or a pasted row of text.',
  'Copy every value exactly as it appears — never paraphrase, never infer, never complete a partial value.',
  'Put each distinct value in the matching array. Leave a key absent when the source has no such value.',
  'Report an address only when the source explicitly states a delivery address; a signature is not an address.',
].join('\n');

// ------------------------------------------------------------------------------------------------
// Profiles. The user answers "what do I use", not "what model id".
// ------------------------------------------------------------------------------------------------

export const PROFILES = ['aws', 'openai', 'anthropic', 'groq', 'deepseek', 'local'] as const;
export type Profile = (typeof PROFILES)[number];

export interface ResolvedProvider {
  profile: Profile;
  /** "provider/model" form Mastra understands, e.g. "openai/gpt-4o-mini". */
  model: string;
  /** Ordered fallbacks. Only transport/availability failures fall back; never a bad extraction. */
  fallbacks: string[];
}

const table: Record<Profile, ResolvedProvider> = {
  aws: {
    profile: 'aws',
    model: 'amazon.nova-micro-v1:0',
    fallbacks: ['amazon.nova-lite-v1:0'],
  },
  openai: {
    profile: 'openai',
    model: 'openai/gpt-4o-mini',
    fallbacks: ['openai/gpt-4o'],
  },
  anthropic: {
    profile: 'anthropic',
    model: 'anthropic/claude-3-5-haiku-20241022',
    fallbacks: ['anthropic/claude-3-5-sonnet-20241022'],
  },
  groq: {
    profile: 'groq',
    model: 'groq/llama-3.3-70b-versatile',
    fallbacks: [],
  },
  deepseek: {
    profile: 'deepseek',
    model: 'deepseek/deepseek-chat',
    fallbacks: [],
  },
  local: {
    profile: 'local',
    model: 'ollama/qwen2.5:14b',
    fallbacks: [], // a local failure is a real failure — never silently hop to a paid API
  },
};

export function resolveProfile(raw?: string): ResolvedProvider {
  const p = (raw ?? 'aws').toLowerCase();
  const hit = PROFILES.find((x) => x === p);
  if (!hit) throw new Error(`LLM_PROFILE "${raw}" is not one of ${PROFILES.join(' | ')}`);
  return table[hit];
}

// ------------------------------------------------------------------------------------------------
// Grounding. A fact is kept only if it literally appears in the source text. This is the
// cross-model quality stabiliser: it does not care how the model produced structured output,
// it independently re-checks that the value has a source.
// ------------------------------------------------------------------------------------------------

export interface GroundResult {
  kept: Partial<Extraction>;
  dropped: { fact: FactKey | 'line'; value: string; reason: 'not-in-source' }[];
}

export function groundFacts(extraction: Extraction, sourceText: string): GroundResult {
  const src = sourceText.toLowerCase();
  const kept: Partial<Extraction> = { lines: [] };
  const dropped: GroundResult['dropped'] = [];

  const keep = (fact: FactKey, value: string) => {
    const arr = (kept[fact] ??= []) as string[];
    if (src.includes(value.toLowerCase())) arr.push(value);
    else dropped.push({ fact, value, reason: 'not-in-source' });
  };

  for (const k of FACT_KEYS) for (const v of extraction[k] ?? []) keep(k, v);

  for (const line of extraction.lines ?? []) {
    if (src.includes(line.description.toLowerCase())) {
      (kept.lines as NonNullable<Extraction['lines']>).push(line);
    } else {
      dropped.push({ fact: 'line', value: line.description, reason: 'not-in-source' });
    }
  }

  return { kept, dropped };
}
