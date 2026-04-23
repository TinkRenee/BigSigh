# AM Dashboard Production Foundation

This folder is the cleaner app foundation for moving the AM dashboard from a fast local prototype into a maintainable web app.

## Structure

- `public/`
  Browser entry files and downloadable assets.
- `src/client/`
  Client-side dashboard code and snapshot data.
- `src/server/`
  Local preview and API server.
- `data/`
  Persisted local dashboard state.
- `docs/`
  Architecture and migration notes.
- `scripts/`
  Snapshot import helpers.

## Run locally

From this folder:

```powershell
node src/server/index.js
```

Then open:

```text
http://127.0.0.1:4173
```

## What this gives you

- a clean repository root
- separated client and server code
- a place for persisted app data
- a place for docs and migration notes
- the current dashboard behavior preserved as a starting point

## What should happen next

This is the transition base, not the final architecture. The next steps should be:

1. Split `src/client/app.js` into focused feature modules.
2. Move server routes in `src/server/index.js` into separate route files.
3. Replace local JSON persistence with a real database.
4. Replace scaffolded integrations with real authenticated services.
5. Add a real build system and typed code once the repo is stable.
