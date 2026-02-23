# justcal.ai

A single-page infinite scrolling calendar built with Vite and vanilla JavaScript.

## Features

- Infinite month timeline with lazy loading in both directions.
- Present-day focused startup with quick "Back to current month/day" floating action.
- Day-cell selection flow with smooth pan+zoom, table row/column expansion, and animated deselection.
- Today indicator styling and per-day hover/selection states.
- Fast smooth scrolling helper used for navigation-to-target day.
- Full keyboard `Esc` handling to close selected day and open popovers.

- Multiple calendar types:
- `Semaphore` (`signal-3`): 4-state day controls (`X`, `Red`, `Yellow`, `Green`) with compact hover controls and top-right state dot.
- `Score`: per-day slider (`-1` to `10`) with themed styling, hover-only controls, and animated in-cell score badge.
- `Check`: click-to-toggle checkmark day state.
- `Notes`: per-day text notes with dedicated day editor, auto-focus on open, and note indicators.

- Score display modes per calendar:
- `Number`
- `Heatmap`
- `Number + Heatmap`
- Heatmap mode shows numeric score on hover while keeping cell intensity rendering.

- Notes UX:
- In-day note editor fills the selected day area, no outline, non-resizable, smaller text.
- Small inline note preview on cell hover.
- Delayed larger note preview popup near hovered cell.

- Calendar management:
- Add unlimited calendars with Name, Type, and Color.
- Score calendars support conditional `Display` property in Add/Edit forms.
- Edit calendar name/color/display.
- Delete calendar with typed-name confirmation.
- Duplicate calendar names are blocked.
- Calendar data is stored per calendar id.

- Calendar pinning:
- Toggle pin per calendar from the menu row.
- Reorder behavior:
- Pinning moves that calendar to top.
- Unpinning moves it to first unpinned position.
- Smooth list-reorder animation for pin/unpin.
- Max 3 pinned calendars.
- When pin cap is reached, unpinned calendars no longer show pin action.
- Non-active pinned calendars are also shown in the header row before the top-right calendar button.

- Themes and visuals:
- Theme switcher with 5 themes: `Dark`, `Tokyo Night Storm` (default), `Solarized Dark`, `Solarized Light`, `Light`.
- Themed score slider colors and theme-aware component styling.
- Custom tooltip system and polished floating controls.

- Developer controls panel:
- Sliders for `Cell Zoom`, `Expand X`, `Expand Y`, and `Fade Delta`.
- Values persist in `localStorage`.
- Toggle with keyboard shortcut `P`.
- Mobile debug toggle button included.

- Misc:
- Header calendar switcher with active calendar chip and calendar menu.
- Pinned calendar chips rendered in header.
- Floating GitHub shortcut button.

## Tech Stack

- Vite 5
- Vanilla JavaScript (ES modules)
- HTML + CSS (single page UI)

## Requirements

- Node.js (recommended: current LTS; this project is running on Node 22 in production).
- npm

## Quick Start

```bash
npm install
npm run dev
```

App scripts:

- `npm run dev` -> start Vite dev server
- `npm run build` -> production build to `dist/`
- `npm run preview` -> preview built app

## HTTPS and Cloudflare (Full Mode)

This project is configured to run HTTPS directly from Vite on port `443`.

Current Vite configuration:

- `host: 0.0.0.0`
- `port: 443`
- `strictPort: true`
- `allowedHosts`: `justcal.ai`, `www.justcal.ai` (plus internal hosts)
- TLS certificate/key loaded from:
  - `certs/justcal.ai.crt`
  - `certs/justcal.ai.key`

Certificate files are intentionally ignored by git (`certs/` in `.gitignore`).

### Generate a self-signed cert (compatible with Cloudflare `Full`)

```bash
mkdir -p certs
openssl req -x509 -nodes -newkey rsa:2048 -sha256 -days 3650 \
  -keyout certs/justcal.ai.key \
  -out certs/justcal.ai.crt \
  -subj "/CN=justcal.ai" \
  -addext "subjectAltName=DNS:justcal.ai,DNS:www.justcal.ai"
```

Cloudflare SSL/TLS mode:

- Use **Full** (not Full strict) when using self-signed origin certs.

## Browser Storage

The app persists state in `localStorage` using:

- `justcal-calendars`
- `justcal-calendar-day-states`
- `justcal-day-states` (legacy key read for migration)
- `justcal-theme`
- `justcal-camera-zoom`
- `justcal-cell-expansion-x`
- `justcal-cell-expansion-y`
- `justcal-fade-delta`
- `justcal-cell-expansion` (legacy key read for migration)
- `justcal-selection-expansion` (legacy key read for migration)

## Project Structure

```text
.
├── index.html
├── src/
│   ├── main.js
│   ├── calendar.js
│   ├── calendars.js
│   ├── theme.js
│   └── tweak-controls.js
├── vite.config.js
└── package.json
```

## Notes for VPS Deployment

- Port `443` usually requires elevated privileges or a reverse proxy.
- If you run the app as a non-root user, either:
  - terminate TLS at Nginx/Caddy and proxy to a higher local port, or
  - grant Node permission to bind low ports.
