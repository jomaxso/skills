# Copilot CLI Marketplace

Personal GitHub Copilot CLI marketplace with custom plugins, agents, and skills.

## Register this marketplace

```bash
copilot plugin marketplace add jomaxso/skills
```

## Available plugins

| Plugin | Description | Version |
|--------|-------------|---------|
| [`atlassian`](./plugins/atlassian) | Manage Jira Cloud work items via `acli` and reusable Atlassian agents for review workflows | 1.2.0 |
| [`database`](./plugins/database) | Manage MariaDB databases from the terminal using the native `mariadb` CLI client | 1.0.0 |
| [`engineering`](./plugins/engineering) | Engineering planning and decision-support skills, including `grilly` for stress-testing ideas and requirements | 1.0.0 |
| [`presentation`](./plugins/presentation) | Build PowerPoint decks as code with `deck-as-code` — a scripted source of truth driving PowerPoint COM (Windows + PowerPoint required) | 1.0.0 |

## Install a plugin

```bash
# From marketplace
copilot plugin install atlassian@jomaxso
copilot plugin install engineering@jomaxso
copilot plugin install presentation@jomaxso

# Directly from GitHub
copilot plugin install jomaxso/skills:plugins/atlassian
copilot plugin install jomaxso/skills:plugins/engineering
copilot plugin install jomaxso/skills:plugins/presentation
```

## Add a new plugin

1. Create `plugins/<name>/` with a `plugin.json` manifest
2. Add `skills/`, `agents/`, `.mcp.json` as needed
3. Register the entry in `.github/plugin/marketplace.json`
4. Push — done.

## Repository structure

```
skills/
├── .github/plugin/
│   └── marketplace.json          # Marketplace manifest (required)
└── plugins/
    └── atlassian/                 # Atlassian tools plugin
        ├── plugin.json
        ├── README.md
        ├── agents/
        │   └── jira.agent.md
        └── skills/
            └── jira/              # Jira management skill
                ├── SKILL.md
                ├── references/
                │   └── REFERENCE.md
                └── scripts/
                    ├── install-windows.ps1
                    ├── install-linux.sh
                    └── install-macos.sh
```
