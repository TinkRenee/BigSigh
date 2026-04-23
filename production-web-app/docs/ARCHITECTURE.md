# Architecture Notes

## Current goal

Preserve the working dashboard while giving it a clearer repository layout that is easier to update over the next few months.

## App layers

### `public/`

Contains the browser entrypoint and static assets.

### `src/client/`

Contains the dashboard UI logic.

Recommended future split:

- `src/client/core/`
  App bootstrapping, config, shared helpers.
- `src/client/state/`
  App state, persistence, selectors.
- `src/client/features/accounts/`
  Account list, account detail, summary table.
- `src/client/features/jira/`
  Jira cards, aging, turnaround, CX request modal.
- `src/client/features/onboarding/`
  Onboarding workspace and conversion flow.
- `src/client/features/pipelines/`
  Pipelines, routing, Eve opportunity signals.
- `src/client/features/integrations/`
  Microsoft, NetSuite, Genesys, future ZoomInfo.

### `src/server/`

Contains the local server and API endpoints.

Recommended future split:

- `src/server/config/`
  Environment variables and integration config.
- `src/server/routes/`
  `dashboard-state`, `jira`, `microsoft`, `netsuite`, `genesys`.
- `src/server/lib/`
  HTTP helpers, static file serving, auth utilities.
- `src/server/integrations/`
  Jira, Microsoft Graph, NetSuite, Genesys, ZoomInfo service clients.

### `data/`

Temporary local persistence for development.

Recommended future destination:

- Postgres / Supabase for shared state
- object storage for attachments and exports

## Why this is better than the prototype root

- clearer repository layout
- easier onboarding for future developers
- a cleaner path to typed modules and tests
- less confusion about what is static, client, server, or persisted data

## Production migration path

1. Keep this folder as the new source of truth.
2. Create a fresh Git repository from this folder.
3. Deploy this repository to a fresh Vercel project.
4. Replace local persistence with a cloud database.
5. Migrate large client files into feature modules one section at a time.
