/**
 * The logic of the MembershipRequests service binding (src/rpc.js), apart from the
 * `cloudflare:workers` entrypoint so that it runs under `node --test`.
 */
import { extractEmails } from './parser.js';
import { submitRequests, replyText } from './requests.js';
import { getTeamName } from './graph.js';
import { replyToTeams } from './teams.js';

// The Bot Connector of the public cloud, where a conversation id from elsewhere is answered.
const SERVICE_URL = 'https://smba.trafficmanager.net/teams/';
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
// A Teams channel, or a thread in one: 19:…@thread.tacv2[;messageid=…]
const CONVERSATION = /^19:[^;\s]+@thread\.[a-z0-9]+(;messageid=\d+)?$/i;
const MAX_EMAILS = 20;

export class MembershipRequests {
  constructor(env, deps = {}) {
    this.env = env;
    this.deps = {
      teamName: getTeamName, submit: submitRequests, reply: replyToTeams, ...deps,
    };
  }

  /**
   * Asks the admins to add people to a Teams team, or remove them, as an @admin-bot mention
   * would. `input`:
   *
   * - `action`: `add` or `remove`.
   * - `emails`: the people's addresses (at most 20).
   * - `teamId`: the team's Microsoft 365 group id (a GUID).
   * - `conversationId`: the channel, or the thread in it (`19:…@thread.tacv2;messageid=<root>`),
   *   where admin-bot answers now and once each request is reviewed.
   * - `requester`: `{ name, email, aadObjectId }`, the person who asked; the card names them.
   * - `via`: what asks for them, such as `opspile`; the card says so.
   * - `message`: the request in words, kept with it. Optional.
   * - `serviceUrl`: the Bot Connector to answer through. Optional.
   *
   * Returns `{ ok, created: [{ id, email }], skipped: [email], failed: [{ email, error }],
   * replied }`, `replied` telling whether admin-bot's answer reached the thread (it must be
   * installed in the team), or `{ ok: false, error }` when the input is wrong.
   */
  async submit(input = {}) {
    const { env } = this;
    const action = input.action === 'remove' ? 'remove' : (input.action === 'add' ? 'add' : null);
    if (!action) {
      return { ok: false, error: 'action must be add or remove' };
    }
    const emails = extractEmails([].concat(input.emails || []).join(' '));
    if (!emails.length) {
      return { ok: false, error: 'no email address given' };
    }
    if (emails.length > MAX_EMAILS) {
      return { ok: false, error: `at most ${MAX_EMAILS} email addresses at a time` };
    }
    const teamId = String(input.teamId || '');
    if (!GUID.test(teamId)) {
      return { ok: false, error: 'teamId must be the team\'s group id' };
    }
    const conversationId = String(input.conversationId || '');
    if (!CONVERSATION.test(conversationId)) {
      return { ok: false, error: 'conversationId must be a Teams channel or thread id' };
    }
    const requester = input.requester || {};
    const via = String(input.via || '').trim().slice(0, 40) || null;
    if (!requester.name && !requester.email && !via) {
      return { ok: false, error: 'say who asks: requester or via' };
    }
    let teamName;
    try {
      teamName = await this.deps.teamName(env, teamId);
    } catch (err) {
      return { ok: false, error: `team ${teamId} could not be read: ${err.message}` };
    }
    const serviceUrl = String(input.serviceUrl || SERVICE_URL);
    const result = await this.deps.submit(env, {
      action,
      emails,
      teamId,
      teamName: teamName || 'This team',
      teamsChannelId: conversationId.split(';')[0],
      requesterName: String(requester.name || via || 'Unknown'),
      requesterAadId: String(requester.aadObjectId || via || 'unknown'),
      requesterEmail: requester.email ? String(requester.email).toLowerCase() : null,
      originalMessage: input.message ? String(input.message).slice(0, 2000) : null,
      conversationId,
      serviceUrl,
      via,
    });
    const text = replyText(action, result);
    let replied = false;
    if (text) {
      replied = await this.deps.reply({ serviceUrl, conversation: { id: conversationId } }, env, text)
        .catch((err) => {
          console.error('replyToTeams failed:', err);
          return false;
        });
    }
    console.log(JSON.stringify({
      event: 'rpc-submit', via, action, created: result.created.length,
      skipped: result.skipped.length, failed: result.failed.length, replied,
    }));
    return {
      ok: true,
      created: result.created.map((r) => ({ id: r.id, email: r.member_email })),
      skipped: result.skipped,
      failed: result.failed,
      replied,
    };
  }
}
