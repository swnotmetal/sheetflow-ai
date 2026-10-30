// Direction A acceptance: the extract step really calls a provider, and one thing above all is
// testable without a network — what happens when the model misbehaves.
//
// The mock is the transport (one function), so the fallback and grounding policies are exercised
// exactly as they run in production.

import assert from 'node:assert/strict';
import test from 'node:test';
import { resolveModelConfig } from '@mastra/core/llm';
import { EXTRACT_SCHEMA, PROFILES, catalogueBlock, resolveProfile, type Extraction } from '../src/model.js';
import {
  ExtractionInvalidError,
  ExtractionUnavailableError,
  extractFacts,
  isTransportError,
  modelConfigFor,
  type ExtractionCall,
} from '../src/extract.js';

const BODY = [
  'From: Acer Industrial <orders@acer.example>',
  'Reference PO 4500',
  'Please ship 2 x steel bracket to our depot.',
].join('\n');

const GOOD: Extraction = {
  intent: 'order',
  customer: ['Acer Industrial'],
  po: ['4500'],
  lines: [{ description: 'steel bracket', quantity: '2' }],
};

/** A provider that is reachable but returns something invalid — never a fallback trigger. */
const badOutput = () => Object.assign(new Error('could not parse response'), { name: 'AI_NoObjectGeneratedError' });
/** A provider that cannot be reached — the only thing that may fall back. */
const unreachable = () => Object.assign(new Error('upstream unavailable'), { name: 'AI_APICallError', statusCode: 503 });

test('grounding drops an invented fact but keeps the real ones', async () => {
  const call: ExtractionCall = async () => ({ ...GOOD, address: ['12 Nowhere Street'] });

  const outcome = await extractFacts(BODY, resolveProfile('openai'), call);

  assert.equal(outcome.model, 'openai/gpt-4o-mini');
  assert.deepEqual(outcome.kept.address, undefined);
  assert.deepEqual(outcome.kept.customer, ['Acer Industrial']);
  assert.deepEqual(outcome.dropped, [
    { fact: 'address', value: '12 Nowhere Street', reason: 'not-in-source' },
  ]);
});

test('a transport failure falls back within the same explicit profile', async () => {
  const seen: string[] = [];
  const call: ExtractionCall = async (target) => {
    seen.push(target.model);
    if (target.model === 'openai/gpt-4o-mini') throw unreachable();
    return GOOD;
  };

  const outcome = await extractFacts(BODY, resolveProfile('openai'), call);

  assert.deepEqual(seen, ['openai/gpt-4o-mini', 'openai/gpt-4o']);
  assert.equal(outcome.model, 'openai/gpt-4o');
  assert.equal(outcome.attempts.length, 1);
  assert.equal(outcome.attempts[0].model, 'openai/gpt-4o-mini');
});

test('a validation failure never falls back — no laundering a bad extraction into a second model', async () => {
  const seen: string[] = [];
  const call: ExtractionCall = async (target) => {
    seen.push(target.model);
    throw badOutput();
  };

  await assert.rejects(
    () => extractFacts(BODY, resolveProfile('openai'), call),
    (error: unknown) => error instanceof ExtractionInvalidError && error.model === 'openai/gpt-4o-mini',
  );
  assert.deepEqual(seen, ['openai/gpt-4o-mini']);
});

test('a structurally wrong object is a validation failure, not a transport failure', async () => {
  let calls = 0;
  const call: ExtractionCall = (async () => {
    calls++;
    return { lines: 'not-an-array' } as unknown as Extraction;
  }) as ExtractionCall;

  await assert.rejects(() => extractFacts(BODY, resolveProfile('openai'), call), ExtractionInvalidError);
  assert.equal(calls, 1);
});

test('an unclassifiable error fails closed instead of trying another model', async () => {
  let calls = 0;
  const call: ExtractionCall = async () => {
    calls++;
    throw new TypeError('something we do not understand');
  };

  await assert.rejects(() => extractFacts(BODY, resolveProfile('openai'), call), ExtractionInvalidError);
  assert.equal(calls, 1);
});

test('exhausting every model in the profile is an availability error, not a silent empty draft', async () => {
  const call: ExtractionCall = async () => {
    throw unreachable();
  };

  await assert.rejects(
    () => extractFacts(BODY, resolveProfile('openai'), call),
    (error: unknown) =>
      error instanceof ExtractionUnavailableError && error.attempts.length === 2,
  );
});

test('the local profile never falls back — a dead local server must not quietly spend money', async () => {
  const local = resolveProfile('local');
  const call: ExtractionCall = async () => {
    throw unreachable();
  };

  await assert.rejects(() => extractFacts(BODY, local, call), ExtractionUnavailableError);
  assert.deepEqual(local.fallbacks, []);
});

test('isTransportError classifies fail-closed', () => {
  assert.equal(isTransportError(Object.assign(new Error('x'), { statusCode: 429 })), true);
  assert.equal(isTransportError(Object.assign(new Error('x'), { statusCode: 503 })), true);
  assert.equal(isTransportError(Object.assign(new Error('x'), { statusCode: 400 })), false);
  assert.equal(isTransportError(Object.assign(new Error('x'), { code: 'ECONNREFUSED' })), true);
  assert.equal(isTransportError(new TypeError('fetch failed')), true);
  assert.equal(isTransportError(Object.assign(new Error('x'), { name: 'AI_NoObjectGeneratedError' })), false);
  assert.equal(isTransportError(new Error('cannot read properties of undefined')), false);
  assert.equal(isTransportError('nope'), false);
});

test('every profile resolves to a model Mastra can actually build', async () => {
  for (const profile of PROFILES) {
    const provider = resolveProfile(profile);
    const model = await modelConfigFor({ ...provider });

    if (typeof model === 'string') {
      // Router path: the string must resolve, and to the provider the table named.
      const resolved = await resolveModelConfig(model);
      assert.equal(resolved.provider, provider.model.split('/')[0], profile);
    } else if ('modelId' in model) {
      // Provider instance (Bedrock): Mastra takes the model object as-is.
      assert.equal((model as { modelId: string }).modelId, provider.model.split('/')[1], profile);
    } else {
      // Config object (local server): the router id still has to be the one the table named.
      assert.equal((model as { id: string }).id, provider.model, profile);
    }
  }
});

test('LLM_API_KEY is only used where a bearer token is the real auth mechanism', async () => {
  const env = { LLM_API_KEY: 'sk-test' };

  assert.deepEqual(await modelConfigFor({ model: 'openai/gpt-4o-mini', profile: 'openai' }, env), {
    id: 'openai/gpt-4o-mini',
    apiKey: 'sk-test',
  });

  // Bedrock signs with SigV4 from the ambient chain — never a bearer token.
  const aws = (await modelConfigFor(
    { model: 'aws-bedrock/eu.amazon.nova-lite-v1:0', profile: 'aws' },
    { ...env, AWS_REGION: 'eu-north-1' },
  )) as unknown as { provider: string; modelId: string };
  assert.equal(aws.provider, 'amazon-bedrock');
  assert.equal(aws.modelId, 'eu.amazon.nova-lite-v1:0');

  assert.equal(
    await modelConfigFor({ model: 'openai/gpt-4o-mini', profile: 'openai' }, {}),
    'openai/gpt-4o-mini',
  );
  assert.deepEqual(
    await modelConfigFor(
      { model: 'ollama/phi4', profile: 'local', baseUrl: 'http://localhost:11434/v1' },
      {},
    ),
    { id: 'ollama/phi4', url: 'http://localhost:11434/v1', apiKey: 'local' },
  );
});

test('the local endpoint and model come from the environment, not the table', () => {
  const provider = resolveProfile('local', {
    OLLAMA_MODEL: 'phi4',
    OLLAMA_BASE_URL: 'http://box:11434/',
  });

  assert.equal(provider.model, 'ollama/phi4');
  assert.equal(provider.baseUrl, 'http://box:11434/v1');
  assert.deepEqual(provider.fallbacks, []);
});

test('an unknown profile is an error, never a silent default', () => {
  assert.throws(() => resolveProfile('gpt-5-please'), /is not one of/);
});

test('a mail with nothing to order may omit `lines` entirely', () => {
  assert.equal(
    EXTRACT_SCHEMA.safeParse({ intent: 'order', lines: [{ description: 'x' }] }).success,
    true,
  );
  // Nova Lite answers {"intent":"question"} for a pitch with no items; that answer is correct.
  const noLines = EXTRACT_SCHEMA.safeParse({ intent: 'question' });
  assert.equal(noLines.success, true);
  assert.deepEqual(noLines.success && noLines.data.lines, undefined);
  assert.equal(EXTRACT_SCHEMA.safeParse({ intent: 'order', lines: 'none' }).success, false);
});

test('intent must be one of the three labels — never an invented one', () => {
  assert.equal(EXTRACT_SCHEMA.safeParse({ intent: 'question', lines: [] }).success, true);
  assert.equal(EXTRACT_SCHEMA.safeParse({ intent: 'other', lines: [] }).success, true);
  assert.equal(EXTRACT_SCHEMA.safeParse({ intent: 'purchase', lines: [] }).success, false);
  assert.equal(EXTRACT_SCHEMA.safeParse({ lines: [] }).success, false, 'intent is required');
});

test('the catalogue asks for a SKU only when the workbook has one', () => {
  const skus = catalogueBlock([
    { sku: 'AP-TSH-BLK-M', name: 'Oversized Tee' },
    { sku: 'ACC-SOC-BLK-OS', name: 'Ribbed Socks' },
  ]);
  assert.match(skus, /copy its `sku` exactly/);
  assert.match(skus, /- AP-TSH-BLK-M \| Oversized Tee/);

  // A shop that keeps names only: nothing to copy into `sku`, so the name becomes the key, verbatim.
  const names = catalogueBlock([
    { sku: '', name: 'Oversized Tee' },
    { sku: '', name: 'Ribbed Socks' },
  ]);
  assert.match(names, /no SKU column/);
  assert.match(names, /copy its name into `description` exactly/);
  assert.match(names, /^CATALOGUE \(product name\):$/m);
  assert.match(names, /^- Oversized Tee$/m);
  assert.doesNotMatch(names, /\|/, 'no empty SKU column to confuse the model');

  assert.equal(catalogueBlock([]), '', 'no catalogue is better than an empty heading');
});
