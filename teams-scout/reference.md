# Teams Scout — Graph API Reference

## Authentication

Teams Scout extracts the user's access token from the MSAL (Microsoft Authentication Library) cache in the Teams web app's `sessionStorage`. This is a delegated token with the scopes the Teams app was granted.

### MSAL Token Cache Structure

MSAL v2 stores tokens in `sessionStorage` with composite keys:

```
<homeAccountId>-<environment>-accesstoken-<clientId>-<realm>-<scopes>
```

The value is JSON:

```json
{
  "credential_type": "AccessToken",
  "secret": "<the actual bearer token>",
  "home_account_id": "...",
  "environment": "login.microsoftonline.com",
  "client_id": "...",
  "target": "openid profile User.Read ...",
  "realm": "<tenant-id>",
  "token_type": "Bearer",
  "expires_on": "1713200000",
  "extended_expires_on": "1713203600"
}
```

The `teams-scout auth` command searches for keys containing both `accesstoken` and `graph.microsoft.com`, then extracts the `secret` field.

### Required Scopes

The Teams web app requests broad scopes. Teams Scout uses:

| Scope | Used by |
|---|---|
| `User.Read` | `auth` (verify token via `/me`) |
| `Team.ReadBasic.All` | `teams` |
| `Channel.ReadBasic.All` | `channels` |
| `ChannelMessage.Read.All` | `messages`, `unanswered`, `digest` |
| `Chat.Read` | `search`, `mentions` (via Search API) |

If the token lacks a scope, the Graph API returns 403. The user may need to consent via the Azure portal or by visiting `https://login.microsoftonline.com/common/adminconsent?client_id=<teams-client-id>`.

### Token Lifetime

Tokens typically expire after 60–90 minutes. The Teams web app silently refreshes them. When a token expires:

1. `teams-scout` commands fail with 401
2. Re-run `teams-scout auth` — the Teams app will have already refreshed the token in `sessionStorage`

## Graph API Endpoints

### User Profile

```
GET /me
```

Returns `displayName`, `mail`, `userPrincipalName`, `id`. Used by `teams-scout auth` to verify the token.

### List Joined Teams

```
GET /me/joinedTeams
```

Returns teams the user is a member of. Response `value` array contains objects with `id`, `displayName`, `description`.

### List Channels

```
GET /teams/{team-id}/channels
```

Returns channels in a team. Each channel has `id`, `displayName`, `description`, `membershipType` (`standard`, `private`, `shared`).

### List Channel Messages

```
GET /teams/{team-id}/channels/{channel-id}/messages
```

Returns top-level messages (not replies). Key query parameters:

| Parameter | Example | Notes |
|---|---|---|
| `$top` | `$top=50` | Messages per page (max 50) |
| `$expand` | `$expand=replies($top=1)` | Include replies inline |
| `$filter` | `$filter=lastModifiedDateTime gt 2024-01-01T00:00:00Z` | Time filter (limited support) |
| `$orderby` | `$orderby=createdDateTime desc` | Sort order |

**Pagination**: If more results exist, the response includes `@odata.nextLink` — a full URL to fetch the next page. Follow it until no `nextLink` is returned.

Each message contains:

```json
{
  "id": "...",
  "messageType": "message",
  "createdDateTime": "2024-03-15T10:30:00Z",
  "from": {
    "user": { "displayName": "Jane Doe", "id": "..." }
  },
  "body": { "contentType": "html", "content": "<p>Hello</p>" },
  "importance": "normal",
  "mentions": [
    { "id": 0, "mentionText": "John", "mentioned": { "user": { "id": "...", "displayName": "John" } } }
  ],
  "reactions": [
    { "reactionType": "like", "user": { "displayName": "Bob" } }
  ],
  "attachments": [],
  "replies": []
}
```

### Get Message Replies

```
GET /teams/{team-id}/channels/{channel-id}/messages/{message-id}/replies
```

Returns replies to a specific message. Same structure as messages.

### Search Messages (Beta)

```
POST https://graph.microsoft.com/beta/search/query
```

Request body:

```json
{
  "requests": [
    {
      "entityTypes": ["chatMessage"],
      "query": { "queryString": "deployment issue" },
      "from": 0,
      "size": 25
    }
  ]
}
```

Response contains `hitsContainers[0].hits[]`, each with a `resource` (the chatMessage) and `summary` (highlighted snippet).

Search covers the signed-in user's accessible Teams messages. Supports KQL-style queries.

### Get All Messages Across Channels (Application-only)

```
GET /teams/{team-id}/channels/getAllMessages
```

Requires application permissions (`ChannelMessage.Read.All`). Supports `$top` and `$filter` on `lastModifiedDateTime`. Not available with delegated tokens from the browser.

## Rate Limiting

Microsoft Graph applies per-app and per-tenant throttling. When throttled:

- Response status: `429 Too Many Requests`
- `Retry-After` header indicates seconds to wait

Teams Scout paginates conservatively (max 5–10 pages per request) to stay within limits. If you hit throttling, wait and retry.

## Common Error Codes

| Status | Meaning | Resolution |
|---|---|---|
| 401 | Token expired or invalid | Run `teams-scout auth` |
| 403 | Insufficient permissions | Token lacks required scopes |
| 404 | Resource not found | Team/channel ID is invalid |
| 429 | Throttled | Wait per `Retry-After` header |
| 503 | Service unavailable | Transient — retry after a few seconds |
