import { handleTeamsActivity } from './teams.js';
import { handleSlackInteraction } from './slack.js';

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    if (request.method === 'POST' && url.pathname === '/api/messages') {
      return handleTeamsActivity(request, env, ctx);
    }

    if (request.method === 'POST' && url.pathname === '/api/slack/interactions') {
      return handleSlackInteraction(request, env, ctx);
    }

    if (url.pathname === '/health') {
      return Response.json({ status: 'ok' });
    }

    return new Response('Not Found', { status: 404 });
  },
};
