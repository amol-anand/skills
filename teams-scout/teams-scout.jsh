// teams-scout.jsh — Microsoft Teams channel scanner via Graph API
// Auto-discovered as `teams-scout` shell command in SLICC.
//
// Usage: teams-scout <subcommand> [args] [--since=<duration>] [--top=<n>]
// Subcommands: auth, teams, channels, messages, mentions, search, unanswered, digest

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const GRAPH_BETA = 'https://graph.microsoft.com/beta';
const TOKEN_PATH = '/workspace/.teams-token';
const TEAMS_CACHE_PATH = '/workspace/.teams-cache.json';

// ---------------------------------------------------------------------------
// Argument parsing
// ---------------------------------------------------------------------------

const args = process.argv.slice(2);
const subcommand = args[0] || '';
const positional = [];
const flags = {};

for (let i = 1; i < args.length; i++) {
  const arg = args[i];
  if (arg.startsWith('--')) {
    const eq = arg.indexOf('=');
    if (eq !== -1) {
      flags[arg.slice(2, eq)] = arg.slice(eq + 1);
    } else {
      flags[arg.slice(2)] = true;
    }
  } else {
    positional.push(arg);
  }
}

const sinceDuration = flags.since || null;
const topN = flags.top ? parseInt(flags.top, 10) : null;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function parseDuration(dur) {
  if (!dur) return null;
  const match = dur.match(/^(\d+)(m|h|d|w)$/);
  if (!match) return null;
  const n = parseInt(match[1], 10);
  const unit = match[2];
  const ms = { m: 60000, h: 3600000, d: 86400000, w: 604800000 };
  return ms[unit] * n;
}

function sinceDate(dur, fallbackHours) {
  const ms = dur ? parseDuration(dur) : fallbackHours * 3600000;
  if (!ms) {
    console.error(`Invalid duration: ${dur}. Use format like 24h, 7d, 2w`);
    process.exit(1);
  }
  return new Date(Date.now() - ms).toISOString();
}

function die(msg) {
  console.error(msg);
  process.exit(1);
}

function out(data) {
  console.log(JSON.stringify(data, null, 2));
}

// ---------------------------------------------------------------------------
// Token management
// ---------------------------------------------------------------------------

async function readToken() {
  try {
    const token = (await fs.readFile(TOKEN_PATH)).trim();
    if (!token) throw new Error('empty');
    return token;
  } catch {
    die(
      'No auth token found. Run `teams-scout auth` first to extract a token from your Teams browser session.'
    );
  }
}

async function saveToken(token) {
  await fs.writeFile(TOKEN_PATH, token);
}

// ---------------------------------------------------------------------------
// Graph API client
// ---------------------------------------------------------------------------

async function graphGet(token, path, params) {
  let url = path.startsWith('http') ? path : `${GRAPH_BASE}${path}`;
  if (params) {
    const qs = new URLSearchParams(params).toString();
    url += (url.includes('?') ? '&' : '?') + qs;
  }
  const resp = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
  });
  if (resp.status === 401) {
    die('401 Unauthorized — token expired. Run `teams-scout auth` to refresh.');
  }
  if (resp.status === 403) {
    die(
      '403 Forbidden — insufficient permissions. The token may lack required Graph API scopes. See reference.md.'
    );
  }
  if (!resp.ok) {
    const body = await resp.text();
    die(`Graph API error ${resp.status}: ${body}`);
  }
  return resp.json();
}

async function graphPost(token, path, body) {
  const url = path.startsWith('http') ? path : `${GRAPH_BETA}${path}`;
  const resp = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      Accept: 'application/json',
    },
    body: JSON.stringify(body),
  });
  if (resp.status === 401) {
    die('401 Unauthorized — token expired. Run `teams-scout auth` to refresh.');
  }
  if (!resp.ok) {
    const text = await resp.text();
    die(`Graph API error ${resp.status}: ${text}`);
  }
  return resp.json();
}

async function graphGetAllPages(token, path, params, maxPages) {
  maxPages = maxPages || 10;
  const results = [];
  let url = path.startsWith('http') ? path : `${GRAPH_BASE}${path}`;
  if (params) {
    const qs = new URLSearchParams(params).toString();
    url += (url.includes('?') ? '&' : '?') + qs;
  }
  let pages = 0;
  while (url && pages < maxPages) {
    const data = await graphGet(token, url);
    if (data.value) results.push(...data.value);
    url = data['@odata.nextLink'] || null;
    pages++;
  }
  return results;
}

// ---------------------------------------------------------------------------
// Teams/channel resolution (name → ID)
// ---------------------------------------------------------------------------

async function getTeams(token) {
  return graphGetAllPages(token, '/me/joinedTeams');
}

async function resolveTeam(token, nameOrId) {
  const teams = await getTeams(token);
  const lower = nameOrId.toLowerCase();
  const exact = teams.find((t) => t.id === nameOrId);
  if (exact) return exact;
  const match = teams.find((t) => t.displayName.toLowerCase().includes(lower));
  if (!match) die(`Team not found: "${nameOrId}". Run \`teams-scout teams\` to list available teams.`);
  return match;
}

async function getChannels(token, teamId) {
  return graphGetAllPages(token, `/teams/${teamId}/channels`);
}

async function resolveChannel(token, teamId, nameOrId) {
  const channels = await getChannels(token, teamId);
  const lower = nameOrId.toLowerCase();
  const exact = channels.find((c) => c.id === nameOrId);
  if (exact) return exact;
  const match = channels.find((c) => c.displayName.toLowerCase().includes(lower));
  if (!match)
    die(
      `Channel not found: "${nameOrId}". Run \`teams-scout channels ${teamId}\` to list available channels.`
    );
  return match;
}

// ---------------------------------------------------------------------------
// Auth subcommand — extract MSAL token from Teams browser tab
// ---------------------------------------------------------------------------

async function cmdAuth() {
  const tabId = await findTeamsTab();

  // Write the token-extraction script to a temp VFS file so we avoid
  // shell-quoting headaches with the long JS expression.
  const extractScript = [
    '(function(){',
    'for(var i=0;i<sessionStorage.length;i++){',
    'var k=sessionStorage.key(i);',
    'if(k&&k.toLowerCase().indexOf("accesstoken")!==-1&&k.toLowerCase().indexOf("graph.microsoft.com")!==-1){',
    'try{var e=JSON.parse(sessionStorage.getItem(k));',
    'if(e&&e.secret)return JSON.stringify({token:e.secret,expiresOn:e.expires_on||e.expiresOn})}catch(x){}}',
    '}',
    'for(var j=0;j<sessionStorage.length;j++){',
    'var k2=sessionStorage.key(j);',
    'if(k2&&k2.toLowerCase().indexOf("accesstoken")!==-1){',
    'try{var e2=JSON.parse(sessionStorage.getItem(k2));',
    'if(e2&&e2.secret&&k2.toLowerCase().indexOf("microsoft")!==-1)',
    'return JSON.stringify({token:e2.secret,expiresOn:e2.expires_on||e2.expiresOn})}catch(x2){}}',
    '}',
    'return null})()',
  ].join('');

  await fs.writeFile('/tmp/.teams-scout-eval.js', extractScript);
  const scriptContent = await fs.readFile('/tmp/.teams-scout-eval.js');

  // Pass the single-line expression through exec; JSON.stringify adds safe quoting
  const evalResult = await exec(
    'playwright-cli eval --tab=' + tabId + ' ' + JSON.stringify(scriptContent)
  );
  const evalOutput = evalResult.stdout.trim();

  if (!evalOutput || evalOutput === 'null' || evalOutput === 'undefined') {
    die(
      'No MSAL token found in Teams session storage. Make sure Teams is fully loaded and you are logged in. Try refreshing the page.'
    );
  }

  let tokenData;
  try {
    let parsed = evalOutput;
    // The eval output may be double-stringified
    if (parsed.startsWith('"') && parsed.endsWith('"')) {
      parsed = JSON.parse(parsed);
    }
    tokenData = typeof parsed === 'string' ? JSON.parse(parsed) : parsed;
  } catch (e) {
    die('Failed to parse token data: ' + evalOutput);
  }

  if (!tokenData || !tokenData.token) {
    die('Token extraction returned empty data. Teams may not be fully loaded.');
  }

  await saveToken(tokenData.token);

  // Verify token by fetching user profile
  const me = await graphGet(tokenData.token, '/me');
  out({
    status: 'authenticated',
    user: me.displayName,
    email: me.mail || me.userPrincipalName,
    id: me.id,
    expiresOn: tokenData.expiresOn || 'unknown',
  });
}

async function findTeamsTab() {
  const tabListResult = await exec('playwright-cli tab-list');
  const lines = tabListResult.stdout.split('\n');
  const teamsLine = lines.find(
    (l) => l.includes('teams.microsoft.com') || l.includes('teams.live.com')
  );

  if (!teamsLine) {
    die(
      'No Teams tab found. Open Teams first:\n  open https://teams.microsoft.com\nWait for it to load, then retry `teams-scout auth`.'
    );
  }

  const idMatch = teamsLine.match(/\[targetId:\s*([^\]]+)\]/) || teamsLine.match(/^(\S+)/);
  if (!idMatch) die('Could not parse Teams tab ID from tab-list output.');
  return idMatch[1].trim();
}

// ---------------------------------------------------------------------------
// Teams subcommand
// ---------------------------------------------------------------------------

async function cmdTeams() {
  const token = await readToken();
  const teams = await getTeams(token);
  out(
    teams.map((t) => ({
      id: t.id,
      name: t.displayName,
      description: t.description || '',
    }))
  );
}

// ---------------------------------------------------------------------------
// Channels subcommand
// ---------------------------------------------------------------------------

async function cmdChannels() {
  if (!positional[0]) die('Usage: teams-scout channels <teamNameOrId>');
  const token = await readToken();
  const team = await resolveTeam(token, positional[0]);
  const channels = await getChannels(token, team.id);
  out(
    channels.map((c) => ({
      id: c.id,
      name: c.displayName,
      description: c.description || '',
      membershipType: c.membershipType,
      team: team.displayName,
    }))
  );
}

// ---------------------------------------------------------------------------
// Messages subcommand
// ---------------------------------------------------------------------------

async function cmdMessages() {
  if (positional.length < 2) die('Usage: teams-scout messages <team> <channel> [--since=24h] [--top=50]');
  const token = await readToken();
  const team = await resolveTeam(token, positional[0]);
  const channel = await resolveChannel(token, team.id, positional[1]);
  const since = sinceDate(sinceDuration, 24);
  const top = topN || 50;

  const messages = await graphGetAllPages(
    token,
    `/teams/${team.id}/channels/${channel.id}/messages`,
    { $top: String(top) },
    5
  );

  const cutoff = new Date(since).getTime();
  const filtered = messages.filter((m) => {
    const ts = new Date(m.createdDateTime).getTime();
    return ts >= cutoff && m.messageType === 'message';
  });

  out(
    filtered.map((m) => ({
      id: m.id,
      from: m.from?.user?.displayName || m.from?.application?.displayName || 'unknown',
      date: m.createdDateTime,
      body: m.body?.content ? stripHtml(m.body.content).slice(0, 500) : '',
      replyCount: m.replies?.length || 0,
      hasAttachments: (m.attachments || []).length > 0,
      importance: m.importance,
      reactions: (m.reactions || []).map((r) => r.reactionType),
      team: team.displayName,
      channel: channel.displayName,
    }))
  );
}

function stripHtml(html) {
  return html
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

// ---------------------------------------------------------------------------
// Mentions subcommand
// ---------------------------------------------------------------------------

async function cmdMentions() {
  const token = await readToken();
  const since = sinceDate(sinceDuration, 168); // default 7 days

  // Use the Search API (beta) to find messages mentioning the current user
  const me = await graphGet(token, '/me');
  const displayName = me.displayName;

  const searchBody = {
    requests: [
      {
        entityTypes: ['chatMessage'],
        query: { queryString: displayName },
        from: 0,
        size: topN || 25,
      },
    ],
  };

  try {
    const result = await graphPost(token, '/search/query', searchBody);
    const hits = result.value?.[0]?.hitsContainers?.[0]?.hits || [];

    const mentions = hits
      .map((hit) => {
        const resource = hit.resource || {};
        return {
          summary: hit.summary || '',
          from: resource.from?.emailAddress?.name || 'unknown',
          date: resource.createdDateTime || resource.lastModifiedDateTime || '',
          body: resource.body?.content ? stripHtml(resource.body.content).slice(0, 500) : hit.summary || '',
          channelName: resource.channelIdentity?.channelId || '',
          teamName: resource.channelIdentity?.teamId || '',
          webUrl: resource.webUrl || '',
        };
      })
      .filter((m) => {
        if (!sinceDuration && !m.date) return true;
        if (!m.date) return true;
        return new Date(m.date).getTime() >= new Date(since).getTime();
      });

    out(mentions);
  } catch (e) {
    // Fallback: if search API is not available, scan channels manually
    console.error(
      'Search API failed (may require additional permissions). Falling back to channel scan...'
    );
    await cmdMentionsFallback(token, me, since);
  }
}

async function cmdMentionsFallback(token, me, since) {
  const teams = await getTeams(token);
  const mentions = [];

  for (const team of teams.slice(0, 10)) {
    const channels = await getChannels(token, team.id);
    for (const channel of channels.slice(0, 10)) {
      try {
        const messages = await graphGetAllPages(
          token,
          `/teams/${team.id}/channels/${channel.id}/messages`,
          { $top: '50' },
          2
        );
        const cutoff = new Date(since).getTime();
        for (const m of messages) {
          if (m.messageType !== 'message') continue;
          if (new Date(m.createdDateTime).getTime() < cutoff) continue;
          const hasMention = (m.mentions || []).some(
            (mention) => mention.mentioned?.user?.id === me.id
          );
          const bodyText = m.body?.content ? stripHtml(m.body.content) : '';
          if (hasMention || bodyText.toLowerCase().includes(me.displayName.toLowerCase())) {
            mentions.push({
              from: m.from?.user?.displayName || 'unknown',
              date: m.createdDateTime,
              body: bodyText.slice(0, 500),
              team: team.displayName,
              channel: channel.displayName,
            });
          }
        }
      } catch {
        // Skip channels we can't read
      }
    }
  }

  out(mentions);
}

// ---------------------------------------------------------------------------
// Search subcommand
// ---------------------------------------------------------------------------

async function cmdSearch() {
  if (!positional[0]) die('Usage: teams-scout search <query> [--since=7d]');
  const token = await readToken();
  const query = positional.join(' ');

  const searchBody = {
    requests: [
      {
        entityTypes: ['chatMessage'],
        query: { queryString: query },
        from: 0,
        size: topN || 25,
      },
    ],
  };

  const result = await graphPost(token, '/search/query', searchBody);
  const hits = result.value?.[0]?.hitsContainers?.[0]?.hits || [];

  const since = sinceDuration ? sinceDate(sinceDuration, 168) : null;

  const results = hits
    .map((hit) => {
      const resource = hit.resource || {};
      return {
        summary: hit.summary || '',
        from: resource.from?.emailAddress?.name || 'unknown',
        date: resource.createdDateTime || resource.lastModifiedDateTime || '',
        body: resource.body?.content ? stripHtml(resource.body.content).slice(0, 500) : hit.summary || '',
        webUrl: resource.webUrl || '',
      };
    })
    .filter((m) => {
      if (!since || !m.date) return true;
      return new Date(m.date).getTime() >= new Date(since).getTime();
    });

  out(results);
}

// ---------------------------------------------------------------------------
// Unanswered subcommand
// ---------------------------------------------------------------------------

async function cmdUnanswered() {
  if (positional.length < 2) die('Usage: teams-scout unanswered <team> <channel> [--since=48h]');
  const token = await readToken();
  const team = await resolveTeam(token, positional[0]);
  const channel = await resolveChannel(token, team.id, positional[1]);
  const since = sinceDate(sinceDuration, 48);

  const messages = await graphGetAllPages(
    token,
    `/teams/${team.id}/channels/${channel.id}/messages`,
    { $top: '50', $expand: 'replies($top=1)' },
    5
  );

  const cutoff = new Date(since).getTime();
  const unanswered = messages.filter((m) => {
    if (m.messageType !== 'message') return false;
    if (new Date(m.createdDateTime).getTime() < cutoff) return false;
    const replyCount = m.replies?.length || 0;
    return replyCount === 0;
  });

  out(
    unanswered.map((m) => ({
      id: m.id,
      from: m.from?.user?.displayName || 'unknown',
      date: m.createdDateTime,
      body: m.body?.content ? stripHtml(m.body.content).slice(0, 500) : '',
      importance: m.importance,
      team: team.displayName,
      channel: channel.displayName,
    }))
  );
}

// ---------------------------------------------------------------------------
// Digest subcommand
// ---------------------------------------------------------------------------

async function cmdDigest() {
  const token = await readToken();
  const since = sinceDate(sinceDuration, 24);
  const cutoff = new Date(since).getTime();
  const teams = await getTeams(token);
  const digest = [];

  for (const team of teams) {
    let channels;
    try {
      channels = await getChannels(token, team.id);
    } catch {
      continue;
    }

    for (const channel of channels) {
      try {
        const messages = await graphGetAllPages(
          token,
          `/teams/${team.id}/channels/${channel.id}/messages`,
          { $top: '50' },
          2
        );

        const recent = messages.filter(
          (m) => m.messageType === 'message' && new Date(m.createdDateTime).getTime() >= cutoff
        );

        if (recent.length === 0) continue;

        const authors = new Set(recent.map((m) => m.from?.user?.displayName || 'unknown'));
        const hasAttachments = recent.some((m) => (m.attachments || []).length > 0);
        const allReactions = recent.flatMap((m) => (m.reactions || []).map((r) => r.reactionType));
        const topMessages = recent.slice(0, 3).map((m) => ({
          from: m.from?.user?.displayName || 'unknown',
          date: m.createdDateTime,
          preview: m.body?.content ? stripHtml(m.body.content).slice(0, 200) : '',
        }));

        digest.push({
          team: team.displayName,
          channel: channel.displayName,
          messageCount: recent.length,
          uniqueAuthors: authors.size,
          authors: [...authors],
          hasAttachments,
          reactionSummary: countOccurrences(allReactions),
          topMessages,
        });
      } catch {
        // Skip channels we can't access
      }
    }
  }

  // Sort by message count descending
  digest.sort((a, b) => b.messageCount - a.messageCount);
  out(digest);
}

function countOccurrences(arr) {
  const counts = {};
  for (const item of arr) {
    counts[item] = (counts[item] || 0) + 1;
  }
  return counts;
}

// ---------------------------------------------------------------------------
// Help
// ---------------------------------------------------------------------------

function showHelp() {
  console.log(`teams-scout — Scan Microsoft Teams channels via Graph API

Usage: teams-scout <command> [args] [--since=<duration>] [--top=<n>]

Commands:
  auth                              Extract auth token from Teams browser session
  teams                             List joined teams
  channels <team>                   List channels in a team
  messages <team> <channel>         Fetch recent messages (default: --since=24h)
  mentions                          Messages mentioning me (default: --since=7d)
  search <query>                    Full-text search across Teams messages
  unanswered <team> <channel>       Messages with no replies (default: --since=48h)
  digest                            Activity summary across all teams (default: --since=24h)

Duration format: <number><unit> where unit is m(inutes), h(ours), d(ays), w(eeks)
  Examples: 30m, 24h, 7d, 2w

Team and channel arguments accept display names (case-insensitive partial match) or IDs.`);
}

// ---------------------------------------------------------------------------
// Router
// ---------------------------------------------------------------------------

switch (subcommand) {
  case 'auth':
    await cmdAuth();
    break;
  case 'teams':
    await cmdTeams();
    break;
  case 'channels':
    await cmdChannels();
    break;
  case 'messages':
  case 'msgs':
    await cmdMessages();
    break;
  case 'mentions':
    await cmdMentions();
    break;
  case 'search':
    await cmdSearch();
    break;
  case 'unanswered':
    await cmdUnanswered();
    break;
  case 'digest':
    await cmdDigest();
    break;
  case '--help':
  case '-h':
  case 'help':
  case '':
    showHelp();
    break;
  default:
    console.error(`Unknown command: ${subcommand}`);
    showHelp();
    process.exit(1);
}
