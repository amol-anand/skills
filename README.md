# skills

Standalone SLICC skills installable via `upskill`.

## Available Skills

| Skill | Description | Install |
|---|---|---|
| [handoff](handoff/) | Hand off tasks to the SLICC agent | `upskill amol-anand/skills --skill handoff` |
| [teams-scout](teams-scout/) | Scour Microsoft Teams channels via Graph API using the browser session for auth | `upskill amol-anand/skills --skill teams-scout` |

## Usage

Install a skill into SLICC:

```bash
upskill amol-anand/skills
```

To install a specific skill when the repo contains multiple:

```bash
upskill amol-anand/skills --skill teams-scout
```
