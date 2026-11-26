import assert from 'node:assert/strict';
import test from 'node:test';
import { OAuth2Client } from 'google-auth-library';

import { classifyApiProbeError, runApiProbe } from '../../src/tools/apiProbe.js';
import { apiErrorFixture as fixture, gaxiosErrorFor } from '../helpers/gaxios-error.js';

// The classifier only ever sees errors thrown by googleapis, so every case here
// is a real gaxios error (see gaxiosErrorFor).
//
// Fixtures are in test/fixtures/api-errors/, with project ids and numbers replaced by
// placeholders:
// - service-disabled-errorinfo.json: the current disabled-API shape. SERVICE_DISABLED
//   in a google.rpc.ErrorInfo with metadata.activationUrl, plus accessNotConfigured in
//   errors[]. The body is from a Google API that is not a Workspace one (Gemini for
//   Google Cloud), copied from a public report. No body in this shape from a Workspace
//   API has been captured yet.
// - service-disabled-legacy.json: Sheets API, the older shape. No reason code anywhere,
//   only the message and a google.rpc.Help link. Copied from a public report.
// - insufficient-permission.json: a 403 for a missing scope, which is not a disabled
//   API. Copied from a public report.
// - not-found.json: Docs documents.get on a nonexistent id, captured live.

test('current ErrorInfo shape is disabled, with the activation URL from ErrorInfo metadata', async () => {
  const url = 'https://console.developers.google.com/apis/api/cloudaicompanion.googleapis.com/overview?project=example-project';
  const body = fixture('service-disabled-errorinfo.json');
  const asSent = classifyApiProbeError(await gaxiosErrorFor(403, body));
  assert.equal(asSent.status, 'disabled');
  assert.equal(asSent.activationUrl, url);

  // The message is free text, so the reason codes must carry the verdict on
  // their own: the same body with every message and Help link removed is
  // still disabled, and the URL still comes from the ErrorInfo.
  const parsed = JSON.parse(body);
  delete parsed.error.message;
  for (const e of parsed.error.errors) delete e.message;
  parsed.error.details = parsed.error.details.filter((d: { '@type': string }) => d['@type'].endsWith('ErrorInfo'));
  const reasonsOnly = classifyApiProbeError(await gaxiosErrorFor(403, JSON.stringify(parsed)));
  assert.equal(reasonsOnly.status, 'disabled');
  assert.equal(reasonsOnly.activationUrl, url);
});

test('legacy message-only shape is disabled, with the activation URL from the Help link', async () => {
  const err = await gaxiosErrorFor(403, fixture('service-disabled-legacy.json'));
  const result = classifyApiProbeError(err);
  assert.equal(result.status, 'disabled');
  assert.equal(
    result.activationUrl,
    'https://console.developers.google.com/apis/api/sheets.googleapis.com/overview?project=123456789012',
  );
});

test('a 403 that is not a disabled API is unknown, with Google\'s message', async () => {
  const err = await gaxiosErrorFor(403, fixture('insufficient-permission.json'));
  const result = classifyApiProbeError(err);
  assert.equal(result.status, 'unknown');
  assert.equal(result.activationUrl, undefined);
  assert.match(result.message ?? '', /Insufficient Permission/);
});

test('a 404 or 400 from the method means the API answered, so it is enabled', async () => {
  const notFound = await gaxiosErrorFor(404, fixture('not-found.json'));
  assert.deepEqual(classifyApiProbeError(notFound), { status: 'enabled' });
  const badRequest = await gaxiosErrorFor(400, '{"error":{"code":400,"message":"Invalid value","status":"INVALID_ARGUMENT"}}');
  assert.deepEqual(classifyApiProbeError(badRequest), { status: 'enabled' });
});

test('a server error or a network failure is unknown, never enabled', async () => {
  const serverError = await gaxiosErrorFor(500, '{"error":{"code":500,"message":"Internal error encountered.","status":"INTERNAL"}}');
  const r500 = classifyApiProbeError(serverError);
  assert.equal(r500.status, 'unknown');
  assert.match(r500.message ?? '', /Internal error encountered/);

  // Nothing listens on port 1, so this is a connection failure with no response.
  const client = new OAuth2Client();
  client.setCredentials({ access_token: 'test-token-not-a-secret' });
  const networkErr = await client.request({ url: 'http://127.0.0.1:1/probe' }).then(
    () => { throw new Error('expected a connection failure'); },
    (e: unknown) => e,
  );
  const rNet = classifyApiProbeError(networkErr);
  assert.equal(rNet.status, 'unknown');
  assert.ok(rNet.message, 'the failure message is kept');
});

test('runApiProbe reports a probe that never answers as unknown once the timeout passes', { timeout: 2000 }, async () => {
  // The request ignores the abort signal and never settles, the worst case.
  const result = await runApiProbe(() => new Promise(() => {}), 20);
  assert.equal(result.status, 'unknown');
  assert.match(result.message ?? '', /timed out after 20ms/);
});
