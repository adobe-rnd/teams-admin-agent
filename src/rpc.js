/**
 * Member requests from other Workers of the Cloudflare account, through a service binding.
 *
 * Teams never delivers a message written by a bot to another bot, so a bot (opspile's thread
 * agents, for one) cannot ask by mentioning @admin-bot. It binds this entrypoint instead:
 *
 *   services = [{ binding = "TEAMS_ADMIN", service = "teams-admin-agent",
 *                 entrypoint = "MembershipRequests" }]
 *
 * and calls `env.TEAMS_ADMIN.submit({ ... })`. An entrypoint other than the default one is
 * reachable only through a binding, never over HTTP, so there is no secret to keep. The request
 * takes the same path as a mention (src/requests.js): a D1 row and an approval card in
 * SLACK_ADMIN_CHANNEL_ID per email, and admin-bot's reply in the Teams thread. Nothing changes
 * in Teams until an admin approves.
 */
import { WorkerEntrypoint } from 'cloudflare:workers';
import { MembershipRequests as Requests } from './membership.js';

export class MembershipRequests extends WorkerEntrypoint {
  /** See `submit` in src/membership.js. */
  async submit(input) {
    return new Requests(this.env).submit(input);
  }
}
