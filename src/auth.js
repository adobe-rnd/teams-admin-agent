/**
 * One-time OAuth2 flow for delegated permissions (add guests to Teams).
 * A team owner signs in; we get a refresh token and store it as DELEGATED_REFRESH_TOKEN.
 */

const DELEGATED_SCOPES = 'offline_access https://graph.microsoft.com/TeamMember.ReadWrite.All https://graph.microsoft.com/User.Read';

export function handleAuthMicrosoft(request, env) {
  const origin = new URL(request.url).origin;
  const redirectUri = `${origin}/auth/microsoft/callback`;
  const params = new URLSearchParams({
    client_id: env.MS_CLIENT_ID,
    response_type: 'code',
    redirect_uri: redirectUri,
    response_mode: 'query',
    scope: DELEGATED_SCOPES,
    state: '',
  });
  const url = `https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/authorize?${params}`;
  return Response.redirect(url, 302);
}

export async function handleAuthMicrosoftCallback(request, env) {
  const url = new URL(request.url);
  const code = url.searchParams.get('code');
  const error = url.searchParams.get('error');
  if (error) {
    return new Response(
      `Authorization failed: ${error}. ${url.searchParams.get('error_description') || ''}`,
      { status: 400, headers: { 'Content-Type': 'text/plain' } },
    );
  }
  if (!code) {
    return new Response('Missing code', { status: 400, headers: { 'Content-Type': 'text/plain' } });
  }

  const origin = new URL(request.url).origin;
  const redirectUri = `${origin}/auth/microsoft/callback`;

  const res = await fetch(`https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: env.MS_CLIENT_ID,
      client_secret: env.MS_CLIENT_SECRET,
      grant_type: 'authorization_code',
      code,
      redirect_uri: redirectUri,
      scope: DELEGATED_SCOPES,
    }),
  });
  if (!res.ok) {
    const text = await res.text();
    return new Response(`Token exchange failed: ${text}`, { status: 400, headers: { 'Content-Type': 'text/plain' } });
  }

  const data = await res.json();
  const refreshToken = data.refresh_token;
  if (!refreshToken) {
    return new Response('No refresh_token in response (ensure offline_access was requested if needed).', {
      status: 400,
      headers: { 'Content-Type': 'text/plain' },
    });
  }

  const html = `<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Teams Admin Agent — Link account</title></head>
<body style="font-family: system-ui; max-width: 600px; margin: 2rem auto; padding: 1rem;">
  <h1>Account linked</h1>
  <p>A team owner (or admin) has signed in. To allow the bot to add <strong>guests</strong> to teams, set this refresh token as a secret:</p>
  <pre style="background: #f4f4f4; padding: 1rem; overflow-x: auto; word-break: break-all;">${escapeHtml(refreshToken)}</pre>
  <p>Run:</p>
  <pre style="background: #f4f4f4; padding: 1rem;">wrangler secret put DELEGATED_REFRESH_TOKEN</pre>
  <p>and paste the token when prompted.</p>
  <p><strong>Redirect URI</strong> to register in your Azure app if not already: <code>${escapeHtml(redirectUri)}</code></p>
  <p><a href="${escapeHtml(origin)}/auth/microsoft">Link again</a></p>
</body>
</html>`;

  return new Response(html, {
    headers: { 'Content-Type': 'text/html; charset=utf-8' },
  });
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
