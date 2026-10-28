// The transport seam: how a profile becomes an actual model call.
//
// `src/model.ts` owns the prompt, the schema and the grounding rule. This file owns *talking to a
// provider* — and the one policy that must stay provider-independent: a bad extraction is never a
// reason to ask a second model (invariant 2). Only transport/availability failures fall back;
// a validation failure surfaces to the human as-is.

import { Agent } from '@mastra/core/agent';
import type { MastraModelConfig } from '@mastra/core/llm';
import {
  catalogueBlock,
  EXTRACT_SCHEMA,
  EXTRACT_SYSTEM,
  groundFacts,
  type Extraction,
  type GroundResult,
  type Profile,
  type ResolvedProvider,
} from './model.js';

/** One provider endpoint attempt. Fallbacks share the profile, so they share credentials. */
export interface CallTarget {
  model: string;
  profile: Profile;
  /** Set for local OpenAI-compatible servers (Ollama); never for cloud profiles. */
  baseUrl?: string;
  /** The sender's own catalogue, read from the workbook. Empty when it could not be read. */
  catalogue?: { sku: string; name: string }[];
}

/**
 * The only thing this module needs from a provider. Tests pass a mock; production passes
 * {@link mastraExtractionCall}. Swapping it out cannot change the fallback or grounding policy.
 */
export type ExtractionCall = (target: CallTarget, text: string) => Promise<Extraction>;

export interface ExtractOutcome {
  kept: Partial<Extraction>;
  dropped: GroundResult['dropped'];
  /** What the model decided the mail is. Not grounded: it is a classification, not a quotation. */
  intent: Extraction['intent'];
  /** The model whose output passed validation — not necessarily `provider.model`. */
  model: string;
  /** Transport failures that happened before this model answered. */
  attempts: { model: string; error: string }[];
}

/** The model answered, but the answer cannot be used. Never retried against another model. */
export class ExtractionInvalidError extends Error {
  constructor(
    readonly model: string,
    readonly cause: unknown,
  ) {
    super(
      `model ${model} produced no usable extraction; refusing to retry with another model ` +
        `(only an unreachable model may fall back, never an unusable answer or a bad credential)`,
      { cause: cause instanceof Error ? cause : undefined },
    );
    this.name = 'ExtractionInvalidError';
  }
}

/** Every model on the profile was unreachable. Distinct from "the model answered badly". */
export class ExtractionUnavailableError extends Error {
  constructor(readonly attempts: { model: string; error: string }[]) {
    super(
      `every model in the profile was unavailable: ` +
        attempts.map((a) => `${a.model} (${a.error})`).join(', '),
    );
    this.name = 'ExtractionUnavailableError';
  }
}

const NETWORK_CODES = new Set([
  'ECONNREFUSED',
  'ECONNRESET',
  'ETIMEDOUT',
  'ENOTFOUND',
  'EAI_AGAIN',
  'EHOSTUNREACH',
  'ENETUNREACH',
  'EPIPE',
  'UND_ERR_CONNECT_TIMEOUT',
  'UND_ERR_SOCKET',
  'UND_ERR_HEADERS_TIMEOUT',
  'UND_ERR_BODY_TIMEOUT',
]);

// Never a transport problem: the model produced something the acceptance rule rejects.
const NOT_TRANSPORT_NAME = /(zod|noobjectgenerated|typevalidation|invalidprompt|invalidargument|invalidtoolinput)/i;
const TRANSPORT_MESSAGE =
  /fetch failed|network error|socket hang up|connection refused|econnrefused|etimedout|enotfound|timed out|timeout|temporarily unavailable/i;

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return typeof value === 'object' && value !== null ? (value as Record<string, unknown>) : undefined;
}

function statusOf(error: Record<string, unknown>): number | undefined {
  const raw = error.statusCode ?? error.status;
  return typeof raw === 'number' ? raw : undefined;
}

/**
 * Is this a transport/availability failure — the only kind that may fall back?
 *
 * Fail-closed on purpose: an error we cannot classify is treated as *not* transport, so an
 * unexplained failure stops the run instead of quietly producing a second, ungrounded extraction.
 */
export function isTransportError(error: unknown): boolean {
  const e = asRecord(error);
  if (!e) return false;
  if (NOT_TRANSPORT_NAME.test(String(e.name ?? ''))) return false;

  const status = statusOf(e);
  if (status !== undefined) return status === 408 || status === 425 || status === 429 || status >= 500;

  if (typeof e.code === 'string' && NETWORK_CODES.has(e.code)) return true;

  const cause = asRecord(e.cause);
  if (cause && typeof cause.code === 'string' && NETWORK_CODES.has(cause.code)) return true;

  if (e.isRetryable === true) return true;

  // Undici surfaces a dead socket as `TypeError: fetch failed` with the real code on `cause`.
  return typeof e.message === 'string' && TRANSPORT_MESSAGE.test(e.message);
}

/**
 * Resolve the profile's model into what Mastra needs. Credentials stay explicit: `LLM_API_KEY` is
 * passed through only where a bearer token is the actual auth mechanism.
 *
 * Bedrock is the one profile Mastra's model router cannot build from a string — `aws-bedrock` is
 * not a routable provider there — so it is constructed from the AI SDK provider instead, signed
 * with the same ambient AWS credential chain the AWS CLI and boto3 use.
 */
export async function modelConfigFor(
  target: CallTarget,
  env: NodeJS.ProcessEnv = process.env,
): Promise<MastraModelConfig> {
  const apiKey = env.LLM_API_KEY?.trim();
  // Every entry in `src/model.ts` is a `provider/model` router id; the cast only tells TS what
  // `resolveProfile` already guarantees.
  const id = target.model as `${string}/${string}`;

  if (target.profile === 'aws') {
    const [{ createAmazonBedrock }, { defaultProvider }] = await Promise.all([
      import('@ai-sdk/amazon-bedrock'),
      import('@aws-sdk/credential-provider-node'),
    ]);
    const bedrock = createAmazonBedrock({
      region: env.AWS_REGION?.trim() || 'eu-north-1',
      credentialProvider: defaultProvider(),
    });
    return bedrock(target.model.replace(/^aws-bedrock\//, '')) as unknown as MastraModelConfig;
  }

  // A local server ignores the key but the OpenAI-compatible transport still wants one.
  if (target.baseUrl) return { id, url: target.baseUrl, apiKey: apiKey || 'local' };

  return apiKey ? { id, apiKey } : target.model;
}

/**
 * The production transport: one Mastra agent, one prompt, one schema. Native response format when
 * the provider has one, prompt injection otherwise — the acceptance rule never changes, so the
 * caller never learns which path was taken.
 */
export const mastraExtractionCall: ExtractionCall = async (target, text) => {
  const agent = new Agent({
    id: 'sheetflow-extract',
    name: 'sheetflow-extract',
    instructions: [EXTRACT_SYSTEM, catalogueBlock(target.catalogue ?? [])].filter(Boolean).join('\n\n'),
    model: await modelConfigFor(target),
  });

  const result = await agent.generate(text, {
    structuredOutput: {
      schema: EXTRACT_SCHEMA,
      // Bedrock has native structured output (tool mode). The router's capability table does not
      // know this vendor id, so 'auto' would inject a schema dump that Nova Lite happily echoes
      // back as `$schema`/`properties` instead of answering. Measured 2026-10-27: native 737 ms,
      // valid; injected schema dump, invalid.
      jsonPromptInjection: target.profile === 'aws' ? false : 'auto',
    },
  });

  return result.object;
};

function describe(error: unknown): string {
  if (error instanceof Error) return error.message;
  return String(error);
}

/**
 * Extract, ground, and hand back what survived. The model is chosen by `LLM_PROFILE`; a fallback
 * is tried only when the current model was unreachable, and the result always names the model that
 * actually answered.
 *
 * `sourceText` defaults to `text`, but a mail supplies subject + sender + body: a customer PO number
 * very often exists only in the subject, and grounding must be able to cite it.
 */
export async function extractFacts(
  text: string,
  provider: ResolvedProvider,
  call: ExtractionCall = mastraExtractionCall,
  sourceText: string = text,
  catalogue: { sku: string; name: string }[] = [],
): Promise<ExtractOutcome> {
  const attempts: { model: string; error: string }[] = [];

  for (const model of [provider.model, ...provider.fallbacks]) {
    let raw: unknown;
    try {
      raw = await call({ model, profile: provider.profile, baseUrl: provider.baseUrl, catalogue }, text);
    } catch (error) {
      if (!isTransportError(error)) throw new ExtractionInvalidError(model, error);
      attempts.push({ model, error: describe(error) });
      continue;
    }

    // Even a hand-written transport cannot smuggle a shape the contract does not allow.
    const parsed = EXTRACT_SCHEMA.safeParse(raw);
    if (!parsed.success) throw new ExtractionInvalidError(model, parsed.error);

    const { kept, dropped } = groundFacts(parsed.data, sourceText);
    return { kept, dropped, intent: parsed.data.intent, model, attempts };
  }

  throw new ExtractionUnavailableError(attempts);
}
