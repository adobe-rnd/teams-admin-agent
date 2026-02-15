import { handleTeamsActivity } from './teams.js';
import { handleSlackInteraction } from './slack.js';
import { handleAuthMicrosoft, handleAuthMicrosoftCallback } from './auth.js';

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    try {
      if (request.method === 'POST' && url.pathname === '/api/messages') {
        return await handleTeamsActivity(request, env, ctx);
      }

      if (request.method === 'POST' && url.pathname === '/api/slack/interactions') {
        return await handleSlackInteraction(request, env, ctx);
      }

      if (request.method === 'GET' && url.pathname === '/auth/microsoft') {
        return handleAuthMicrosoft(request, env);
      }

      if (request.method === 'GET' && url.pathname === '/auth/microsoft/callback') {
        return await handleAuthMicrosoftCallback(request, env);
      }

      if (url.pathname === '/health') {
        return Response.json({ status: 'ok' });
      }

      return new Response('Not Found', { status: 404 });
    } catch (err) {
      console.error('Unhandled error:', err.message, err.stack);
      return new Response('Internal Server Error', { status: 500 });
    }
  },
};
