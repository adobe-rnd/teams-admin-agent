/**
 * Microsoft Graph API helpers — app-only client-credentials flow.
 * Pure fetch, no SDK.
 */

let _graphToken = { value: null, expiresAt: 0 };

async function getGraphToken(env) {
  if (_graphToken.value && Date.now() < _graphToken.expiresAt - 60_000) {
    return _graphToken.value;
  }
  const res = await fetch(
    `https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: env.MS_CLIENT_ID,
        client_secret: env.MS_CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials',
      }),
    },
  );
  if (!res.ok) throw new Error(`Graph token error ${res.status}: ${await res.text()}`);
  const data = await res.json();
  _graphToken = { value: data.access_token, expiresAt: Date.now() + data.expires_in * 1000 };
  return _graphToken.value;
}

async function graphApi(env, path, method = 'GET', body = null) {
  const token = await getGraphToken(env);
  const opts = {
    method,
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
  };
  if (body) opts.body = JSON.stringify(body);

  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, opts);
  if (!res.ok) throw new Error(`Graph ${method} ${path} (${res.status}): ${await res.text()}`);
  if (res.status === 204) return null;
  return res.json();
}

export async function getTeamName(env, teamId) {
  const data = await graphApi(env, `/teams/${teamId}?$select=displayName`);
  return data.displayName;
}

export async function resolveUser(env, email) {
  return graphApi(env, `/users/${email}?$select=id,displayName,mail`);
}

export async function addTeamMember(env, teamId, email) {
  const user = await resolveUser(env, email);
  await graphApi(env, `/teams/${teamId}/members`, 'POST', {
    '@odata.type': '#microsoft.graph.aadUserConversationMember',
    roles: [],
    'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${user.id}')`,
  });
  return user;
}
