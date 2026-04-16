---
name: teams-scout
description: >-
  Scour Microsoft Teams channels via Graph API using the browser session for auth.
  Use when the user asks about Teams messages, mentions, unanswered questions,
  channel activity, or wants a digest of what is happening across their teams.
allowed-tools: bash
---

# Teams Scout

Scan Microsoft Teams channels via the Microsoft Graph API. Auth is zero-config: the skill extracts the user's access token from a live Teams browser session via CDP.

## Prerequisites

The user must be logged into Microsoft Teams in the browser (`teams.microsoft.com`). If Teams is not open, open it:

```bash
open https://teams.microsoft.com
```

Wait for the page to fully load before extracting a token.

## Authentication

Run `teams-scout auth` to extract and store a Graph API token from the Teams browser tab. The command:

1. Finds the Teams tab via `playwright-cli tab-list`
2. Reads the MSAL token cache from **`localStorage`** via `playwright-cli eval`
3. Stores the token at `/workspace/.teams-token`
4. Prints the authenticated user's name and ID

```bash
teams-scout auth
```

If the token expires (you get 401 errors), re-run `teams-scout auth` to refresh it. The browser session handles token renewal automatically -- just re-extract.

> **Implementation note:** The Teams web client (v2, `teams.microsoft.com/v2/`) stores MSAL tokens in `localStorage`, not `sessionStorage`. The auth command searches `localStorage` for the key matching `accesstoken` + `graph.microsoft.com`, picks the entry with the latest `expiresOn`, and extracts its `secret` field. The `sessionStorage` path is a dead end on modern Teams.

## API Endpoint Note

Message reads use the **beta** Graph endpoint (`https://graph.microsoft.com/beta/...`), not v1.0. The delegated token from the Teams browser session does not include `ChannelMessage.Read.All`, so the v1.0 messages endpoint returns 403. The beta endpoint works with the scopes the Teams session token actually provides. Channel listing and user profile calls use v1.0 as normal.

## Commands

All commands output JSON to stdout. Parse the output to answer the user's question.

```bash
teams-scout auth                                        # Extract token, print user info
teams-scout teams                                       # List joined teams
teams-scout channels <teamNameOrId>                     # List channels in a team
teams-scout messages <teamNameOrId> <channelNameOrId>   # Recent messages (default: 24h)
teams-scout messages <team> <channel> --since=7d        # Messages from last 7 days
teams-scout messages <team> <channel> --top=50          # Limit to 50 messages
teams-scout mentions                                    # Messages mentioning me (default: 7d)
teams-scout mentions --since=30d                        # Mentions in last 30 days
teams-scout search "deployment issue"                   # Full-text search across Teams
teams-scout search "deployment issue" --since=14d       # Search with time filter
teams-scout unanswered <team> <channel>                 # Messages with 0 replies (48h)
teams-scout unanswered <team> <channel> --since=7d      # Unanswered in last 7 days
teams-scout digest                                      # Activity digest across all teams (24h)
teams-scout digest --since=7d                           # Weekly digest
```

Team and channel arguments accept either display names (case-insensitive partial match) or IDs.

## Workflows

### "What tagged me?"

```bash
teams-scout auth
teams-scout mentions --since=7d
```

Parse the output and summarize: who mentioned the user, in which channel, with message previews.

### "Any unanswered questions in #general?"

```bash
teams-scout auth
teams-scout unanswered "My Team" "General" --since=7d
```

Look for messages that are questions (contain `?`, start with interrogative words, or have a question-like tone). Summarize each with author, timestamp, and content.

### "What are the top issues this week?"

```bash
teams-scout auth
teams-scout digest --since=7d
```

Group messages by topic/theme, count occurrences, and rank by frequency. Highlight threads with the most replies or reactions.

### "Search for a topic across all channels"

```bash
teams-scout auth
teams-scout search "outage" --since=30d
```

Summarize matching messages grouped by channel with timestamps and authors.

## Troubleshooting

| Problem | Fix |
|---|---|
| `No Teams tab found` | Open Teams: `open https://teams.microsoft.com`, wait for load, retry |
| `No MSAL token found` | Modern Teams (v2) stores tokens in `localStorage`, not `sessionStorage`. If auth fails, the script may be checking the wrong storage. Ensure the skill's `cmdAuth()` searches `localStorage` for keys containing both `accesstoken` and `graph.microsoft.com`, selecting the entry with the highest `expiresOn`. |
| `401 Unauthorized` | Token expired. Run `teams-scout auth` again |
| `403 Forbidden on message reads` | The v1.0 messages endpoint requires `ChannelMessage.Read.All` which the delegated Teams token does not have. Use the beta endpoint instead: `https://graph.microsoft.com/beta/teams/{id}/channels/{id}/messages` |
| `Team/channel not found` | Run `teams-scout teams` or `teams-scout channels <team>` to list available names/IDs |

For Graph API details, scopes, pagination, and MSAL token structure, see [reference.md](reference.md).
