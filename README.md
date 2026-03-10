# teams-admin-agent

Cloudflare Worker that bridges Microsoft Teams and Slack to automate member-addition requests with admin approval.

## Data Flow

There are two independent flows — the **request flow** (Teams → Worker → Slack) and the **approval flow** (Slack → Worker → Teams). The Cloudflare Worker is the only component you deploy; everything else is a managed service that calls it or gets called by it.

### Flow 1: Submitting a request (Teams → Slack)

```mermaid
sequenceDiagram
    participant User as User in Teams
    participant Bot as Azure Bot Service
    participant Worker as Cloudflare Worker
    participant Slack as Slack API

    User->>Bot: "@admin-bot add alice@co.com and bob@co.com"
    Note over User,Bot: Teams detects @mention, forwards message
    Bot->>Worker: POST /api/messages (JSON activity + Bearer JWT)
    Note over Worker: 1. Validate JWT against Bot Framework JWKS<br/>2. Extract emails via regex<br/>3. Look up team via Graph API<br/>4. For each email: save in D1, POST approval card to Slack<br/>5. Reply to requester via Bot REST API
    Worker->>Slack: chat.postMessage (approval cards with Approve/Reject)
    Worker->>Bot: POST /v3/conversations/activities
    Bot->>User: "2 requests submitted"
```

**How does Teams reach the Worker?** When you register an [Azure Bot](https://portal.azure.com/#create/Microsoft.AzureBot), you set a **messaging endpoint** URL — you point this at `https://<your-worker>.workers.dev/api/messages`. From then on, any time a user @mentions the bot in Teams, Azure Bot Service delivers the message as an HTTP POST to that URL. The Worker never polls — it just receives webhooks.

**How does the Worker post to Slack?** It calls the [Slack Web API](https://api.slack.com/methods/chat.postMessage) directly via `fetch`. No Slack SDK, no Socket Mode — just `POST https://slack.com/api/chat.postMessage` with the bot token and a JSON body containing the approval card blocks.

### Flow 2: Approving a request (Slack → Teams)

```mermaid
sequenceDiagram
    participant Admin as Admin in Slack
    participant Slack as Slack Platform
    participant Worker as Cloudflare Worker
    participant Graph as Microsoft Graph
    participant Bot as Azure Bot Service
    participant User as User in Teams

    Admin->>Slack: clicks [Approve]
    Slack->>Worker: POST /api/slack/interactions (x-slack-signature + form-encoded payload)
    Note over Worker: 1. Verify HMAC-SHA256 signature<br/>2. Look up request in D1<br/>3. Add member (or invite+add if guest) via Graph<br/>4. Update request status in D1<br/>5. Update Slack card (replace buttons with Approved)<br/>6. Notify requester in Teams via Bot REST API
    Worker->>Graph: POST /teams/{id}/members (or POST /invitations then add)
    Worker->>Slack: chat.update (card shows Approved)
    Worker->>Bot: POST /v3/conversations/activities
    Bot->>User: "alice@co.com has been added to Engineering"
```

### Flow 3: Rejecting a request (Slack → modal → Teams)

Rejection is a two-step interaction: the button click opens a modal for an optional reason, then the modal submission completes the rejection.

```mermaid
sequenceDiagram
    participant Admin as Admin in Slack
    participant Slack as Slack Platform
    participant Worker as Cloudflare Worker
    participant Bot as Azure Bot Service
    participant User as User in Teams

    Admin->>Slack: clicks [Reject]
    Slack->>Worker: POST /api/slack/interactions (block_actions, reject_request, trigger_id)
    Note over Worker: 1. Verify signature<br/>2. Look up request in D1, confirm still pending<br/>3. Call views.open with trigger_id (modal with reason field)
    Worker->>Slack: views.open (Reject Request modal)
    Slack->>Admin: modal appears (Reason: "Not part of this project", [Cancel] [Reject])

    Admin->>Slack: clicks [Reject] in modal
    Slack->>Worker: POST /api/slack/interactions (view_submission, reject_reason_modal, values.reason)
    Note over Worker: 4. Verify signature<br/>5. Update request status → rejected in D1, save reason<br/>6. Update Slack card (replace buttons with Rejected + reason)<br/>7. Notify requester in Teams via Bot REST API
    Worker->>Slack: chat.update (card shows Rejected)
    Worker->>Bot: POST /v3/conversations/activities
    Bot->>User: "Request #1 rejected — alice@co.com was not added to Engineering. > Not part of this project"
```

**Why two round-trips for reject?** The first interaction (button click) carries a `trigger_id` — a short-lived token that Slack requires to open a modal. The Worker uses it to call `views.open`, which pops the reason form on the admin's screen. When the admin submits the modal, Slack sends a second POST (`view_submission`) to the same Worker endpoint. Only then does the Worker update D1, rewrite the Slack card, and notify Teams. If the admin clicks Cancel, nothing happens — the request stays pending.

**How does Slack reach the Worker?** When you configure [Interactivity](https://api.slack.com/interactivity) in your Slack app, you set a **Request URL** — you point this at `https://<your-worker>.workers.dev/api/slack/interactions`. Every button click and modal submission is delivered as an HTTP POST to that URL. Both Flow 2 and Flow 3 use this same endpoint.

**How does the Worker add the member?** For users already in the tenant, it calls the [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/api/team-post-members) (`POST /teams/{team-id}/members`) using a delegated token (or app-only where allowed). For users not yet in the tenant (guests), it sends a B2B invitation (`POST /invitations`), gets `invitedUser.id` from the response, then adds that user to the team in the same step — so once they accept the invite, they already have access (no second “add to team” action). This only happens on approval; rejection skips this step entirely.

**How does the Worker notify back in Teams?** It calls the [Bot Framework REST API](https://learn.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-connector-send-and-receive-messages) (`POST {serviceUrl}/v3/conversations/{id}/activities`) using a bot token. The `serviceUrl` and `conversationId` were saved in D1 when the original request came in. Both approval and rejection send a notification.

### What you deploy vs. what's managed

| Component | Who runs it | Role |
|---|---|---|
| **Cloudflare Worker** | You deploy via `wrangler deploy` | The only code you own — routes webhooks, calls APIs |
| **Cloudflare D1** | Cloudflare (edge SQLite) | Stores request state so approval can find the original context |
| **Azure Bot Service** | Microsoft (managed) | Routes Teams @mentions to your Worker URL as HTTP POSTs |
| **Slack Platform** | Slack (managed) | Routes button clicks to your Worker URL as HTTP POSTs |
| **Microsoft Graph API** | Microsoft (managed) | Called by the Worker to add members to Teams |

## Prerequisites

- A Cloudflare account (Workers + D1)
- A Microsoft 365 tenant with admin access
- A Slack workspace

## Setup

### 1. Azure: App Registration + Bot

A single script creates everything — Entra ID app registration, client secret, Graph API permissions with admin consent, Azure Bot resource, and Teams channel:

```bash
# Prerequisites: az cli (logged in), jq
az login

./infra/setup.sh --worker-url https://teams-admin-agent.workers.dev
```

The script:

| Step | What it creates | Why |
|---|---|---|
| App registration | Entra ID app (multi-tenant) | Identity for the bot and Graph API calls |
| Client secret | 2-year credential | `BOT_PASSWORD` and `MS_CLIENT_SECRET` |
| API permissions | `User.Read.All`, `Team.ReadBasic.All`, `TeamMember.ReadWrite.All` | Graph API access to resolve users, read teams, add members |
| Admin consent | Tenant-wide grant | Application permissions require admin consent |
| Resource group | `teams-admin-agent-rg` | Container for Azure resources |
| Azure Bot | Bot resource with messaging endpoint | Routes Teams @mentions to your Worker URL |
| Teams channel | Enables the Teams channel on the bot | Allows the bot to receive messages from Teams |

When finished it prints the exact `wrangler secret put` commands to run.

To tear down all Azure resources later:

```bash
./infra/teardown.sh
```

#### Install the bot in Teams

After running `setup.sh`, package the Teams app manifest:

```bash
./infra/package-manifest.sh
```

This generates `dist/teams-admin-agent.zip` containing `manifest.json` (with your BOT_ID substituted) and placeholder icons. Upload it via **Teams Admin Center → Manage apps → Upload** or sideload in **Teams → Apps → Manage your apps → Upload a custom app**.

The manifest sets the bot's short name to `admin-bot`, so users will @mention it as `@admin-bot` in Teams channels.

To use custom icons, replace `dist/manifest/color.png` (192×192) and `dist/manifest/outline.png` (32×32) before zipping, or replace the placeholder PNGs and re-run the script.

### 2. Slack App

1. [api.slack.com/apps](https://api.slack.com/apps) → **Create New App → From scratch**.
2. **Interactivity & Shortcuts** → toggle **On** → Request URL:
   ```
   https://<your-worker>.workers.dev/api/slack/interactions
   ```
3. **OAuth & Permissions → Bot Token Scopes**: add `chat:write`.
4. **Install to Workspace** → copy the `xoxb-…` token → `SLACK_BOT_TOKEN`.
5. **Basic Information** → copy **Signing Secret** → `SLACK_SIGNING_SECRET`.
6. Create a channel (e.g. `#teams-requests`), invite the bot, copy the channel ID → `SLACK_ADMIN_CHANNEL_ID`.

### 3. Deploy the Worker

```bash
npm install

# Create the D1 database
npm run db:create
# Paste the returned database_id into wrangler.toml

# Apply the schema
npm run db:migrate:prod

# Set secrets
wrangler secret put BOT_ID
wrangler secret put BOT_PASSWORD
wrangler secret put SLACK_BOT_TOKEN
wrangler secret put SLACK_SIGNING_SECRET
wrangler secret put SLACK_ADMIN_CHANNEL_ID
wrangler secret put MS_TENANT_ID
wrangler secret put MS_CLIENT_ID
wrangler secret put MS_CLIENT_SECRET

# Deploy
npm run deploy
```

#### Adding guest users to teams

Microsoft Graph does not allow **app-only** (client credentials) to add **guest** users to a team. To add guests, a **team owner** (or admin) must link their account once so the Worker can use **delegated** permissions for invitations and add-member calls:

1. In **Azure Portal** → your app registration → **Authentication** → add a **Redirect URI** (Web):
   `https://<your-worker>.workers.dev/auth/microsoft/callback`
   Under **API permissions**, ensure **Delegated** permissions include `TeamMember.ReadWrite.All`, `User.Read`, and `User.Invite.All` (Add permission → Microsoft Graph → Delegated).
2. Open in a browser: `https://<your-worker>.workers.dev/auth/microsoft`
3. Sign in with an account that is **owner** (or team admin) of the teams you want to add guests to.
4. On the success page, copy the **refresh token** and run:
   `wrangler secret put DELEGATED_REFRESH_TOKEN`
   Paste the token when prompted.
5. Redeploy (or the next approve will use the new secret).

After that, when the requested user is not in the tenant, the Worker sends a B2B invitation and **adds them to the team in the same step** using the `invitedUser.id` from the invitation response. Once they accept the invite, they already have access to the team — no second approval or manual add needed.

### Local Development

```bash
# Create a .dev.vars file with your secrets (same keys as .env.example)
npm run db:migrate   # applies migrations locally
npm run dev          # starts local Worker
```

Use [Cloudflare Tunnel](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/) or ngrok to expose the local Worker for the Bot Framework and Slack endpoints.

## API Payloads

Every HTTP call the Worker receives and makes, with the exact JSON shapes.

### Inbound: Teams → Worker

Azure Bot Service delivers activities to `POST /api/messages`.

**Headers:**

```
Authorization: Bearer eyJhbG…  (JWT signed by Microsoft, validated against Bot Framework JWKS)
Content-Type: application/json
```

**Body** (Bot Framework Activity — only the fields we use):

```json
{
  "type": "message",
  "text": "<at>admin</at> please add alice@company.com and bob@company.com",
  "from": {
    "id": "29:1abc…",
    "name": "Jane Smith",
    "aadObjectId": "00000000-0000-0000-0000-000000000001"
  },
  "conversation": {
    "id": "19:abc123…@thread.tacv2"
  },
  "channelData": {
    "team": {
      "id": "19:xyz789…@thread.tacv2"
    },
    "teamsChannelId": "19:abc123…@thread.tacv2"
  },
  "serviceUrl": "https://smba.trafficmanager.net/teams/",
  "entities": [
    {
      "type": "mention",
      "mentioned": { "id": "28:bot-id", "name": "admin" },
      "text": "<at>admin</at>"
    }
  ]
}
```

**Worker response:** `200` (empty body). All processing happens asynchronously via `ctx.waitUntil()`.

### Outbound: Worker → Slack (`chat.postMessage`)

Posts one approval card per email to the admin channel.

```
POST https://slack.com/api/chat.postMessage
Authorization: Bearer xoxb-…
Content-Type: application/json
```

```json
{
  "channel": "C0123456789",
  "text": "Request #1: add alice@company.com to Engineering (from Jane Smith)",
  "blocks": [
    {
      "type": "header",
      "text": { "type": "plain_text", "text": "📋  Request #1" }
    },
    {
      "type": "section",
      "fields": [
        { "type": "mrkdwn", "text": "*Requested by:*\nJane Smith" },
        { "type": "mrkdwn", "text": "*Date:*\n2026-02-12 14:30:00" }
      ]
    },
    {
      "type": "section",
      "fields": [
        { "type": "mrkdwn", "text": "*Microsoft Team:*\nEngineering" },
        { "type": "mrkdwn", "text": "*Email to add:*\nalice@company.com" }
      ]
    },
    {
      "type": "section",
      "text": {
        "type": "mrkdwn",
        "text": "*Original message:*\n> please add alice@company.com and bob@company.com"
      }
    },
    { "type": "divider" },
    {
      "type": "actions",
      "block_id": "approval_actions",
      "elements": [
        {
          "type": "button",
          "text": { "type": "plain_text", "text": "✅ Approve" },
          "style": "primary",
          "action_id": "approve_request",
          "value": "1",
          "confirm": {
            "title": { "type": "plain_text", "text": "Confirm Approval" },
            "text": { "type": "mrkdwn", "text": "Add *alice@company.com* to *Engineering*?" },
            "confirm": { "type": "plain_text", "text": "Approve" },
            "deny": { "type": "plain_text", "text": "Cancel" }
          }
        },
        {
          "type": "button",
          "text": { "type": "plain_text", "text": "❌ Reject" },
          "style": "danger",
          "action_id": "reject_request",
          "value": "1"
        }
      ]
    }
  ]
}
```

**Slack response:**

```json
{
  "ok": true,
  "ts": "1707745800.001234",
  "channel": "C0123456789"
}
```

The `ts` is saved in D1 so the card can be updated later via `chat.update`.

### Outbound: Worker → Teams (bot reply)

Sends a confirmation or notification back to the Teams channel.

**Token acquisition:**

```
POST https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token
Content-Type: application/x-www-form-urlencoded

grant_type=client_credentials
&client_id={BOT_ID}
&client_secret={BOT_PASSWORD}
&scope=https://api.botframework.com/.default
```

**Response:**

```json
{
  "access_token": "eyJhbG…",
  "expires_in": 3600,
  "token_type": "Bearer"
}
```

**Send reply:**

```
POST {serviceUrl}/v3/conversations/{conversationId}/activities
Authorization: Bearer eyJhbG…
Content-Type: application/json
```

```json
{
  "type": "message",
  "text": "**Submitted 2 request(s)** for admin approval:\n- `alice@company.com` → request #1\n- `bob@company.com` → request #2\n\nYou'll be notified here once each is approved or rejected."
}
```

### Inbound: Slack → Worker (approve button click)

Slack delivers interaction payloads to `POST /api/slack/interactions`.

**Headers:**

```
Content-Type: application/x-www-form-urlencoded
X-Slack-Request-Timestamp: 1707745900
X-Slack-Signature: v0=a1b2c3d4…  (HMAC-SHA256 of v0:{timestamp}:{body} using signing secret)
```

**Body** (form-encoded, `payload` field contains JSON):

```json
{
  "type": "block_actions",
  "trigger_id": "7890.1234.abcd",
  "user": {
    "id": "U0SLACKADMIN",
    "username": "adminuser",
    "name": "Admin User"
  },
  "channel": { "id": "C0123456789" },
  "message": { "ts": "1707745800.001234" },
  "actions": [
    {
      "action_id": "approve_request",
      "block_id": "approval_actions",
      "type": "button",
      "value": "1"
    }
  ]
}
```

**Worker response:** `200` (empty body). Processing happens via `ctx.waitUntil()`.

### Outbound: Worker → Graph API (add team member)

When the user is **not** in the tenant (guest), the Worker first creates a B2B invitation, then adds the invited user to the team:

1. **Create invitation:** `POST https://graph.microsoft.com/v1.0/invitations` (delegated token, body: `invitedUserEmailAddress`, `inviteRedirectUrl`, `sendInvitationMessage: true`). Response includes `invitedUser.id`.
2. **Add to team:** Same as below, using `invitedUser.id` from the response. The guest gets access once they accept the invite.

For users **already** in the tenant, the Worker resolves the user by email, then adds them:

**Token acquisition:**

```
POST https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token
Content-Type: application/x-www-form-urlencoded

grant_type=client_credentials
&client_id={MS_CLIENT_ID}
&client_secret={MS_CLIENT_SECRET}
&scope=https://graph.microsoft.com/.default
```

**Resolve user by email:**

```
GET https://graph.microsoft.com/v1.0/users/alice@company.com?$select=id,displayName,mail
Authorization: Bearer eyJhbG…
```

```json
{
  "id": "00000000-0000-0000-0000-000000000042",
  "displayName": "Alice Johnson",
  "mail": "alice@company.com"
}
```

**Add member to team:**

```
POST https://graph.microsoft.com/v1.0/teams/{teamId}/members
Authorization: Bearer eyJhbG…
Content-Type: application/json
```

```json
{
  "@odata.type": "#microsoft.graph.aadUserConversationMember",
  "roles": [],
  "user@odata.bind": "https://graph.microsoft.com/v1.0/users('00000000-0000-0000-0000-000000000042')"
}
```

### Outbound: Worker → Slack (`chat.update` — after approval)

Replaces the buttons with a static "Approved" card.

```
POST https://slack.com/api/chat.update
Authorization: Bearer xoxb-…
Content-Type: application/json
```

```json
{
  "channel": "C0123456789",
  "ts": "1707745800.001234",
  "text": "Request #1 approved by Admin User",
  "blocks": [
    {
      "type": "header",
      "text": { "type": "plain_text", "text": "✅  Request #1 — Approved" }
    },
    {
      "type": "section",
      "fields": [
        { "type": "mrkdwn", "text": "*Requested by:*\nJane Smith" },
        { "type": "mrkdwn", "text": "*Reviewed by:*\nAdmin User" }
      ]
    },
    {
      "type": "section",
      "fields": [
        { "type": "mrkdwn", "text": "*Team:*\nEngineering" },
        { "type": "mrkdwn", "text": "*Member:*\nalice@company.com" }
      ]
    },
    {
      "type": "context",
      "elements": [
        { "type": "mrkdwn", "text": "Approved on 2026-02-12 14:35:00" }
      ]
    }
  ]
}
```

### Inbound: Slack → Worker (reject button click)

Same endpoint, same headers/signature as approve. The payload differs in `action_id`:

```json
{
  "type": "block_actions",
  "trigger_id": "7890.5678.efgh",
  "user": {
    "id": "U0SLACKADMIN",
    "username": "adminuser",
    "name": "Admin User"
  },
  "channel": { "id": "C0123456789" },
  "actions": [
    {
      "action_id": "reject_request",
      "type": "button",
      "value": "1"
    }
  ]
}
```

**Worker response:** `200`. The Worker then calls `views.open` to show the reason modal.

### Outbound: Worker → Slack (`views.open` — rejection reason modal)

```
POST https://slack.com/api/views.open
Authorization: Bearer xoxb-…
Content-Type: application/json
```

```json
{
  "trigger_id": "7890.5678.efgh",
  "view": {
    "type": "modal",
    "callback_id": "reject_reason_modal",
    "private_metadata": "{\"requestId\":1,\"reviewerId\":\"U0SLACKADMIN\",\"reviewerName\":\"Admin User\"}",
    "title": { "type": "plain_text", "text": "Reject Request" },
    "submit": { "type": "plain_text", "text": "Reject" },
    "close": { "type": "plain_text", "text": "Cancel" },
    "blocks": [
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "Rejecting request *#1* — add *alice@company.com* to *Engineering*."
        }
      },
      {
        "type": "input",
        "block_id": "reject_reason",
        "optional": true,
        "label": { "type": "plain_text", "text": "Reason" },
        "element": {
          "type": "plain_text_input",
          "action_id": "reason",
          "multiline": true,
          "placeholder": { "type": "plain_text", "text": "Optional reason…" }
        }
      }
    ]
  }
}
```

### Inbound: Slack → Worker (reject modal submission)

When the admin clicks "Reject" in the modal:

```json
{
  "type": "view_submission",
  "user": {
    "id": "U0SLACKADMIN",
    "username": "adminuser",
    "name": "Admin User"
  },
  "view": {
    "callback_id": "reject_reason_modal",
    "private_metadata": "{\"requestId\":1,\"reviewerId\":\"U0SLACKADMIN\",\"reviewerName\":\"Admin User\"}",
    "state": {
      "values": {
        "reject_reason": {
          "reason": {
            "type": "plain_text_input",
            "value": "Not part of this project"
          }
        }
      }
    }
  }
}
```

**Worker response:**

```json
{ "response_action": "clear" }
```

This closes the modal. The Worker then updates the Slack card via `chat.update` (same shape as the approval update but with "Rejected" and the reason) and notifies the requester in Teams.

### Outbound: Worker → Teams (rejection notification)

Same endpoint and auth as the approval notification:

```
POST {serviceUrl}/v3/conversations/{conversationId}/activities
Authorization: Bearer eyJhbG…
Content-Type: application/json
```

```json
{
  "type": "message",
  "text": "🚫 Request #1 rejected — **alice@company.com** was *not* added to **Engineering**.\n> Not part of this project"
}
```

### Outbound: Worker → Graph API (get team name)

Called during the request flow to resolve the team's display name from its ID.

```
GET https://graph.microsoft.com/v1.0/teams/{teamId}?$select=displayName
Authorization: Bearer eyJhbG…
```

```json
{
  "displayName": "Engineering"
}
```

## Security

Every arrow between components is authenticated. There is no unauthenticated path into or out of the Worker.

```mermaid
flowchart LR
    subgraph Inbound["Inbound"]
        Azure[Azure Bot Service]
        SlackIn[Slack Platform]
    end

    subgraph Worker["Cloudflare Worker"]
        W[Routes + D1 + Secrets]
    end

    subgraph Outbound["Outbound"]
        SlackAPI[Slack Web API]
        Graph[Graph API]
        BotREST[Bot Framework REST]
    end

    Azure -->|"POST /api/messages<br/>JWT (RSA), JWKS, audience, issuer, expiry"| W
    SlackIn -->|"POST /api/slack/interactions<br/>HMAC-SHA256, timestamp, x-slack-signature"| W
    W -->|"Bearer xoxb-…"| SlackAPI
    W -->|"OAuth2 client credentials"| Graph
    W -->|"BOT_ID + BOT_PASSWORD"| BotREST
```

**Inbound:** `/api/messages` validates JWT (Bearer, JWKS, audience, issuer, expiry). `/api/slack/interactions` validates HMAC-SHA256 of `v0:{timestamp}:{body}` and timestamp &lt; 5 min. **Outbound:** Slack (bot token), Graph (client credentials), Bot Framework (client credentials). **Storage:** All credentials in Worker secrets; D1 bound via wrangler, not exposed over HTTP.

### Trust boundary breakdown

| Boundary | Direction | Auth mechanism | What it proves | Code location |
|---|---|---|---|---|
| Azure Bot Service → Worker | Inbound | JWT (RSA signature, JWKS verification) | The request genuinely came from Microsoft's Bot Framework, not a spoofed POST. Audience check confirms it's intended for *this* bot. Issuer check confirms the token source. | `src/teams.js` lines 26-29 |
| Slack → Worker | Inbound | HMAC-SHA256 signature + timestamp | The request genuinely came from Slack (only Slack and the Worker know the signing secret). Timestamp check within 5 minutes rejects replay attacks. Constant-time comparison prevents timing attacks. | `src/slack.js` lines 94-102 |
| Worker → Slack API | Outbound | Bearer token (`xoxb-…`) over HTTPS | The Worker is authorized to post/update messages in the workspace. Token is scoped to `chat:write` only — minimal privilege. | `src/slack.js` lines 8-15 |
| Worker → Graph API | Outbound | OAuth2 client credentials → Bearer token, over HTTPS | The Worker is authorized to read teams, resolve users, and add members. Scoped to `User.Read.All`, `Team.ReadBasic.All`, `TeamMember.ReadWrite.All` — no broader access. | `src/graph.js` lines 8-29 |
| Worker → Bot Framework REST | Outbound | OAuth2 client credentials → Bearer token, over HTTPS | The Worker is authorized to send messages as the bot back to Teams conversations. | `src/teams.js` lines 78-93 |
| Worker → D1 | Internal | Cloudflare binding (not network-accessible) | D1 is only reachable from this Worker via the `DB` binding. There is no HTTP endpoint for the database. | `wrangler.toml` binding config |
| Secrets | At rest | Cloudflare Worker secrets (encrypted) | Credentials are never in source code, `wrangler.toml`, or environment files. Set via `wrangler secret put`, encrypted at rest by Cloudflare. | `.env.example` (reference only) |

### What's NOT covered (and how to address it)

| Gap | Risk | Mitigation |
|---|---|---|
| No per-user authorization on Slack side | Any member of `#teams-requests` can click Approve/Reject | Restrict channel membership to admins only. Slack's channel permissions are the access control layer. |
| No rate limiting on `/api/messages` | A compromised token could flood the Worker | Add Cloudflare rate limiting rules in the dashboard, or implement per-IP/per-conversation throttling in code. |
| Bot Framework JWKS is cached in-memory | If the isolate lives long, stale keys could be used | `jose`'s `createRemoteJWKSet` handles key rotation by re-fetching when verification fails with an unknown `kid`. |
| `serviceUrl` from the activity is trusted | A crafted activity with a malicious `serviceUrl` could redirect bot replies | The JWT validation ensures only Microsoft-signed activities are accepted, so `serviceUrl` is trustworthy. |
| Graph API permissions are broad | `User.Read.All` can read all users in the tenant | This is the minimum needed to resolve emails to object IDs. Cannot be further scoped in the current Graph API. |

## File Overview

| File | Purpose |
|---|---|
| `src/index.js` | Worker entry — routes `/api/messages` and `/api/slack/interactions` |
| `src/teams.js` | Validates Bot Framework JWT, parses @admin-bot messages, replies in Teams |
| `src/slack.js` | Posts approval cards, handles Approve/Reject buttons and modals |
| `src/graph.js` | Graph API via fetch — token acquisition, resolve users, add members |
| `src/parser.js` | Extracts email addresses from natural-language messages |
| `src/db.js` | D1 CRUD for request tracking |
| `migrations/` | D1 schema migrations |
| `infra/setup.sh` | Creates all Azure resources (app registration, bot, permissions) |
| `infra/teardown.sh` | Deletes all Azure resources |
| `infra/manifest.json` | Teams app manifest template (`{{BOT_ID}}` placeholder) |
| `infra/package-manifest.sh` | Substitutes BOT_ID, generates icons, produces ZIP for upload |

## Example

In a Teams channel:

> **@admin-bot** please add alice@company.com, bob@company.com, and carol@company.com

Bot replies:

> **Submitted 3 request(s)** for admin approval:
> - `alice@company.com` → request #1
> - `bob@company.com` → request #2
> - `carol@company.com` → request #3
>
> You'll be notified here once each is approved or rejected.

Three separate cards appear in Slack `#teams-requests`, each with Approve / Reject buttons.

## Troubleshooting

| Symptom | Fix |
|---|---|
| Bot doesn't respond in Teams | Check the messaging endpoint in Azure Bot points to your Worker URL |
| 401 on `/api/messages` | Verify `BOT_ID` matches the app registration; check JWT clock skew |
| No cards in Slack | Verify `SLACK_ADMIN_CHANNEL_ID` is correct and the bot is in the channel |
| "Failed to add member" | Check Graph permissions are admin-consented; verify the email exists in AAD |
| Slack buttons do nothing | Verify the Interactivity Request URL points to your Worker |

## License

MIT
