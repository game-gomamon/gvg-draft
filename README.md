# GVG Team Counter Viewer

Look at the enemy GVG teams, pick the Animus to send against them, mark the result.

Pure static site — no backend, no build step, no database.

```
/
├── index.html
├── style.css
├── app.js
├── vendor/
│   └── xlsx.full.min.js     SheetJS, bundled so nothing loads from a CDN
├── data/
│   └── animus.xlsx          two columns: Animus | Profile
└── assets/
    └── animus/              79 portraits (.webp)
```

## Deploy

Push the folder to a repo, then turn on **Settings → Pages → Deploy from branch**.
All paths are relative, so `https://USERNAME.github.io/REPOSITORY/` works as-is.

## Run it locally

The spreadsheet is fetched over HTTP, so opening `index.html` straight from disk
will not work. Serve the folder instead:

```bash
python3 -m http.server 8000
# then open http://localhost:8000
```

## Set this week's enemy teams

Open `app.js` and edit the array at the top. Names must match the `Animus`
column in the spreadsheet; `null` leaves a slot blank.

```js
const enemyTeams = [
  { id: 1, enemies: ["Airon", "Areal", "Asal"] },
  { id: 2, enemies: ["Batsby", "Beyontin", "Borgne"] },
  ...
];
```

You can also edit enemy slots straight from the page — tap one and pick from the
library, same as your own team. Those edits live in the tab only; put anything
you want to keep into `enemyTeams`.

## Add or change an Animus

1. Drop the portrait into `assets/animus/` (square, `.webp` or `.png`).
2. Add a row to `data/animus.xlsx`: the name in **Animus**, the filename in **Profile**.

The column names are read exactly as written — keep them as `Animus` and `Profile`.

## How to use

Tap a slot → pick an Animus from the library → it fills and jumps to the next
empty slot in that trio. `×` on a portrait clears it. An Animus already in a trio
is dimmed in the library so you cannot double it up. After the fight, hit **WIN**
or **Loss**; the counter in the header keeps score. Drag-and-drop from the library
onto a slot works too.
