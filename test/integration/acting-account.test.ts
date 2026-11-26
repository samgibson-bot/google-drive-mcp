import assert from 'node:assert/strict';
import { describe, it, before, after, beforeEach } from 'node:test';
import { setupTestServer, callTool, type TestContext } from '../helpers/setup-server.js';

// With more than one account connected, a successful write names the account
// that took it, so a write that landed on the default instead of the intended
// account is visible in the reply. With one account there is nothing to tell
// apart, and the reply is unchanged.

describe('Acting account on write replies', () => {
  let ctx: TestContext;
  let serverModule: any;

  before(async () => {
    ctx = await setupTestServer();
    serverModule = await import('../../src/index.js');
  });

  after(async () => {
    // Back to the single 'test' account for any later test in this process.
    serverModule._setAuthClientForTesting({});
    await ctx.cleanup();
  });

  beforeEach(() => {
    ctx.mocks.drive.tracker.reset();
  });

  describe('single account', () => {
    it('leaves a successful write reply unchanged', async () => {
      const res = await callTool(ctx.client, 'createFolder', { name: 'Reports' });
      assert.equal(res.isError, false);
      assert.ok(!res.content.some((b) => b.text?.includes('[acting as')), res.content[0].text);
    });
  });

  describe('several accounts', () => {
    before(async () => {
      // 'test' stays the default; 'work' and 'home' are the extra accounts.
      await serverModule._addSyntheticAccountForTesting('work', {});
      await serverModule._addSyntheticAccountForTesting('home', {});
    });

    it('ends a successful write reply with the account the call was routed to, in its own block', async () => {
      const res = await callTool(ctx.client, 'createFolder', { name: 'Reports', account: 'work' });
      assert.equal(res.isError, false);
      // The handler's own block comes first and is left untouched; the stamp is a block after it.
      assert.ok(res.content.length > 1, 'the stamp is a block of its own');
      assert.ok(!res.content[0].text!.includes('[acting as'), res.content[0].text);
      assert.equal(res.content[res.content.length - 1].text, '[acting as work@example.com (work)]');
      assert.equal(res.content.filter((b) => b.text?.includes('[acting as')).length, 1, 'stamped once');
    });

    it('names the default account when the write gave none', async () => {
      const res = await callTool(ctx.client, 'createFolder', { name: 'Reports' });
      assert.equal(res.isError, false);
      assert.equal(res.content[res.content.length - 1].text, '[acting as test@example.com (test)]');
    });

    it('keeps a JSON write reply parseable', async () => {
      // updateGoogleSheetIfUnchanged is classed as a write and its whole reply is
      // one JSON.stringify block; its caller feeds postFingerprint back to roll back.
      ctx.mocks.sheets.service.spreadsheets.get._setImpl(async () => ({
        data: {
          properties: { title: 'Model' },
          sheets: [{
            properties: { sheetId: 0, title: 'Sheet1' },
            data: [{ startRow: 0, startColumn: 0, rowData: [{ values: [{ userEnteredValue: { stringValue: 'old' } }] }] }],
          }],
        },
      }));
      try {
        const res = await callTool(ctx.client, 'updateGoogleSheetIfUnchanged', {
          spreadsheetId: 'sheet-1',
          updates: [{ range: 'Sheet1!A1', values: [['1']] }],
          dryRun: true,
          account: 'work',
        });
        assert.equal(res.isError, false, res.content[0].text);
        const payload = JSON.parse(res.content[0].text!);
        assert.match(payload.fingerprint, /^v1:[0-9a-f]{64}$/);
        assert.equal(res.content[res.content.length - 1].text, '[acting as work@example.com (work)]');
      } finally {
        ctx.mocks.sheets.service.spreadsheets.get._resetImpl();
      }
    });

    it('prints only the alias while the account identity is still pending', async () => {
      // An account whose email lookup failed at add time is stored as 'unknown'.
      const store = serverModule._getAuthSystemForTesting().store;
      const original = store.get('home');
      await store.upsert({ ...original, email: 'unknown', pendingIdentity: true });
      try {
        const res = await callTool(ctx.client, 'createFolder', { name: 'Reports', account: 'home' });
        assert.equal(res.isError, false);
        assert.equal(res.content[res.content.length - 1].text, '[acting as home]');
      } finally {
        await store.upsert(original);
      }
    });

    it('leaves a read reply unchanged', async () => {
      const res = await callTool(ctx.client, 'search', { query: 'q', account: 'work' });
      assert.equal(res.isError, false);
      assert.ok(!res.content.some((b) => b.text?.includes('[acting as')), res.content[0].text);
    });

    it('leaves a failed write reply unchanged', async () => {
      // restoreRevision returns its API failure as an isError reply rather than
      // throwing, so this reaches the stamping step instead of the outer catch.
      ctx.mocks.drive.service.files.get._setImpl(async () => { throw new Error('Insufficient Permission'); });
      try {
        const res = await callTool(ctx.client, 'restoreRevision', {
          fileId: 'file-1', revisionId: '1', confirm: true, account: 'work',
        });
        assert.equal(res.isError, true);
        assert.equal(res.content[0].text, 'Error: Failed to restore revision: Insufficient Permission');
      } finally {
        ctx.mocks.drive.service.files.get._resetImpl();
      }
    });

    it('leaves an admin tool reply unchanged', async () => {
      const res = await callTool(ctx.client, 'manage_accounts', { action: 'list' });
      assert.notEqual(res.isError, true);
      // The reply is JSON; a stamp would make it unparseable.
      const payload = JSON.parse(res.content[0].text!);
      assert.deepEqual(payload.accounts.map((a: { alias: string }) => a.alias).sort(), ['home', 'test', 'work']);
    });
  });
});
