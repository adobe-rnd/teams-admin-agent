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
  return graphApiWithToken(token, path, method, body);
}

/** Call Graph with a specific token (e.g. delegated). */
async function graphApiWithToken(accessToken, path, method = 'GET', body = null) {
  const opts = {
    method,
    headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
  };
  if (body) opts.body = JSON.stringify(body);

  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, opts);
  const resText = await res.text();
  if (!res.ok) {
    if (res.status === 403 && path.includes('/members')) {
      throw new Error(
        '403 Forbidden — the account from /auth/microsoft must be an owner of this team. If the invitee is a guest, your tenant may block adding guests via Graph; add them manually in Teams if needed.',
      );
    }
    if (res.status === 403 && path.includes('/invitations')) {
      const lower = resText.toLowerCase();
      if (lower.includes('guest invitations not allowed') || lower.includes('not allowed for your company')) {
        throw new Error(
          'Guest invitations are disabled by your organization. A Microsoft 365 admin must enable B2B guest invitations in **Microsoft Entra ID** → External ID → External collaboration settings.',
        );
      }
    }
    throw new Error(`Graph ${method} ${path} (${res.status}): ${resText}`);
  }
  if (res.status === 204) return null;
  return resText ? JSON.parse(resText) : null;
}

let _delegatedToken = { value: null, expiresAt: 0 };

/** Get access token from delegated refresh token (for adding guests; app-only cannot add guests). */
async function getDelegatedToken(env) {
  if (!env.DELEGATED_REFRESH_TOKEN) return null;
  if (_delegatedToken.value && Date.now() < _delegatedToken.expiresAt - 120_000) {
    return _delegatedToken.value;
  }
  const res = await fetch(
    `https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: env.MS_CLIENT_ID,
        client_secret: env.MS_CLIENT_SECRET,
        grant_type: 'refresh_token',
        refresh_token: env.DELEGATED_REFRESH_TOKEN,
        scope: 'https://graph.microsoft.com/.default',
      }),
    },
  );
  if (!res.ok) {
    console.error('Delegated token refresh failed:', await res.text());
    return null;
  }
  const data = await res.json();
  _delegatedToken = { value: data.access_token, expiresAt: Date.now() + (data.expires_in ?? 3600) * 1000 };
  return _delegatedToken.value;
}

export async function getTeamName(env, teamId) {
  const data = await graphApi(env, `/teams/${teamId}?$select=displayName`);
  return data.displayName;
}

/** Set of member emails (lowercase) for the team. Handles pagination. */
export async function getTeamMemberEmails(env, teamId) {
  const emails = new Set();
  let path = `/teams/${teamId}/members`;
  while (path) {
    const data = await graphApi(env, path);
    const members = data.value ?? [];
    for (const m of members) {
      if (m.email) emails.add(m.email.toLowerCase().trim());
    }
    const nextLink = data['@odata.nextLink'];
    path = nextLink ? nextLink.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/, '') : null;
  }
  return emails;
}

/** Get user mail by AAD object ID. Returns null if not found or no mail. */
export async function getRequesterEmail(env, aadObjectId) {
  if (!aadObjectId) return null;
  try {
    const data = await graphApi(env, `/users/${aadObjectId}?$select=mail`);
    return data?.mail ?? null;
  } catch {
    return null;
  }
}

/**
 * Resolve user by email (UPN or mail). GET /users/{id} only accepts objectId or userPrincipalName;
 * if the address is in mail but not UPN (e.g. rofe@adobe.com), we fall back to $filter by mail.
 */
export async function resolveUser(env, email) {
  try {
    return await graphApi(env, `/users/${encodeURIComponent(email)}?$select=id,displayName,mail`);
  } catch (err) {
    if (!err.message?.includes('404')) throw err;
  }
  const filter = `mail eq '${email.replace(/'/g, "''")}' or userPrincipalName eq '${email.replace(/'/g, "''")}'`;
  const data = await graphApi(env, `/users?$filter=${encodeURIComponent(filter)}&$select=id,displayName,mail&$top=1`);
  const user = data?.value?.[0];
  if (!user) throw new Error(`User not found: ${email}`);
  return user;
}

/**
 * Send a B2B invitation to an external email. Requires delegated token with User.Invite.All.
 * Redirect URL sends the user to Teams after they accept.
 * Returns the invitation response (includes invitedUser.id for adding to team pro forma).
 * @param {object} [options] - Optional. { displayName } sets invitedUserDisplayName on the guest.
 */
export async function sendInvitation(env, email, options = {}) {
  const token = await getDelegatedToken(env);
  if (!token) {
    throw new Error(
      'DELEGATED_REFRESH_TOKEN is required to send invitations. Visit GET /auth/microsoft, sign in, then set the refresh token.',
    );
  }
  const body = {
    invitedUserEmailAddress: email,
    inviteRedirectUrl: 'https://teams.microsoft.com',
    sendInvitationMessage: true,
  };
  if (options.displayName?.trim()) {
    body.invitedUserDisplayName = options.displayName.trim();
  }
  return graphApiWithToken(token, '/invitations', 'POST', body);
}

/**
 * Add the user to the team, or if they are not in the tenant, send a B2B invitation and add them to the team pro forma (they get access once they accept).
 * Returns { user } when added to team; { user, invited: true } when an invite was sent and they were added to the team.
 * @param {object} [options] - Optional. { displayName } sets the guest's display name when sending an invitation.
 */
export async function addTeamMember(env, teamId, email, options = {}) {
  let user;
  try {
    user = await resolveUser(env, email);
  } catch (err) {
    if (!err.message?.includes('User not found')) throw err;
    const delegatedToken = await getDelegatedToken(env);
    if (!delegatedToken) {
      throw new Error(
        'DELEGATED_REFRESH_TOKEN is required. Visit GET /auth/microsoft, sign in as a team owner, then set the returned refresh token as DELEGATED_REFRESH_TOKEN.',
      );
    }
    const invitation = await sendInvitation(env, email, { displayName: options.displayName });
    const invitedUserId = invitation?.invitedUser?.id;
    if (!invitedUserId) {
      throw new Error('Invitation did not return invited user id');
    }
    const body = {
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: [],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${invitedUserId}')`,
    };
    await graphApiWithToken(delegatedToken, `/teams/${teamId}/members`, 'POST', body);
    return { user: { id: invitedUserId }, invited: true };
  }

  const delegatedToken = await getDelegatedToken(env);
  if (!delegatedToken) {
    throw new Error(
      'DELEGATED_REFRESH_TOKEN is required. Visit GET /auth/microsoft, sign in as a team owner, then set the returned refresh token as DELEGATED_REFRESH_TOKEN.',
    );
  }
  const body = {
    '@odata.type': '#microsoft.graph.aadUserConversationMember',
    roles: [],
    'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${user.id}')`,
  };
  await graphApiWithToken(delegatedToken, `/teams/${teamId}/members`, 'POST', body);
  return { user };
}
