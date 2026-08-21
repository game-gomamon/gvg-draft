/* ============================================================
   GVG Team Counter Viewer
   Everything runs in the browser. Paths are relative so this
   works from https://USERNAME.github.io/REPOSITORY/
   ============================================================ */

const DATA_FILE  = "./data/animus.xlsx";
const IMAGE_PATH = "./assets/animus/";

/* ------------------------------------------------------------
   EDIT ME: the enemy teams you are scouting this week.
   Names must match the "Animus" column in data/animus.xlsx.
   Use null for an unknown enemy slot.
   ------------------------------------------------------------ */
const enemyTeams = [
  { id: 1, enemies: ["Airon", "Areal", "Asal"] },
  { id: 2, enemies: ["Batsby", "Beyontin", "Borgne"] },
  { id: 3, enemies: ["Cachi", "Celince", "Chiaki"] },
  { id: 4, enemies: ["Danvery", "Dinah", "Diting"] },
  { id: 5, enemies: ["Freya", "Fuqiu", "Gianc"] }
];

/* ---------------- state ---------------- */

const animus = [];              // [{ name, file }]
const byKey  = new Map();       // lowercase name -> record

const state = {
  teams: enemyTeams.map((t, i) => ({
    id: t.id ?? i + 1,
    enemies: normalise(t.enemies),
    mine: [null, null, null],
    result: null                // null | "win" | "loss"
  })),
  target: null                  // { team, side, slot }
};

function normalise(list) {
  const out = [null, null, null];
  (list || []).slice(0, 3).forEach((n, i) => { out[i] = n || null; });
  return out;
}

/* ---------------- elements ---------------- */

const el = {
  rows:        document.getElementById("rows"),
  grid:        document.getElementById("libraryGrid"),
  status:      document.getElementById("libraryStatus"),
  search:      document.getElementById("search"),
  targetNote:  document.getElementById("targetNote"),
  library:     document.getElementById("library"),
  drawerToggle:document.getElementById("drawerToggle"),
  drawerClose: document.getElementById("drawerClose"),
  scrim:       document.getElementById("scrim"),
  tallyWin:    document.getElementById("tallyWin"),
  tallyLoss:   document.getElementById("tallyLoss"),
  addTeam:     document.getElementById("addTeamBtn"),
  reset:       document.getElementById("resetBtn"),
  hint:        document.getElementById("hint")
};

/* ---------------- data loading ---------------- */

async function loadAnimus() {
  const res = await fetch(DATA_FILE);
  if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
  const book  = XLSX.read(await res.arrayBuffer(), { type: "array" });
  const sheet = book.Sheets[book.SheetNames[0]];
  const rows  = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  rows.forEach(row => {
    const name = String(row["Animus"] || "").trim();
    const file = String(row["Profile"] || "").trim();
    if (!name || !file) return;
    const record = { name, file };
    animus.push(record);
    byKey.set(name.toLowerCase(), record);
  });

  animus.sort((a, b) => a.name.localeCompare(b.name));
}

function lookup(name) {
  return name ? byKey.get(String(name).trim().toLowerCase()) || null : null;
}

/* ---------------- board ---------------- */

function renderBoard() {
  el.rows.replaceChildren(...state.teams.map(rowFor));
  updateTally();
}

function rowFor(team, index) {
  const row = document.createElement("div");
  row.className = "team-row";

  const name = document.createElement("div");
  name.className = "team-name";
  name.append(`Team # ${index + 1}`);
  if (state.teams.length > 1) {
    const kill = document.createElement("button");
    kill.type = "button";
    kill.className = "team-remove";
    kill.textContent = "×";
    kill.title = "Remove this team";
    kill.setAttribute("aria-label", `Remove team ${index + 1}`);
    kill.addEventListener("click", () => {
      state.teams.splice(index, 1);
      if (state.target && state.target.team === team) clearTarget();
      renderBoard();
      renderLibrary();
    });
    name.append(kill);
  }
  row.append(name);

  row.append(sideFor(team, "enemy", index));
  row.append(sideFor(team, "mine", index));

  const result = document.createElement("div");
  result.className = "result";
  result.append(resultButton(team, "win", "WIN"), resultButton(team, "loss", "Loss"));
  row.append(result);

  return row;
}

function sideFor(team, side, index) {
  const wrap = document.createElement("div");
  wrap.className = "side";

  const label = document.createElement("span");
  label.className = `side-label ${side === "enemy" ? "e" : "m"}`;
  label.textContent = side === "enemy" ? "Enemy" : "Mine";
  wrap.append(label);

  const slots = document.createElement("div");
  slots.className = "slots";
  for (let i = 0; i < 3; i++) slots.append(slotFor(team, side, i, index));
  wrap.append(slots);
  return wrap;
}

function slotFor(team, side, i, index) {
  const picked = team[side === "enemy" ? "enemies" : "mine"][i];
  const record = lookup(picked);

  const slot = document.createElement("button");
  slot.type = "button";
  slot.className = `slot ${side}`;
  const who = side === "enemy" ? "enemy" : "my team";
  slot.setAttribute("aria-label",
    picked ? `${picked}, ${who} slot ${i + 1}, team ${index + 1}. Change`
           : `Empty ${who} slot ${i + 1}, team ${index + 1}. Choose an Animus`);

  if (isTarget(team, side, i)) slot.classList.add("active");

  if (record) {
    slot.classList.add("filled");
    const img = document.createElement("img");
    img.src = IMAGE_PATH + record.file;
    img.alt = "";
    img.loading = "lazy";
    img.draggable = false;
    const tag = document.createElement("span");
    tag.className = "slot-name";
    tag.textContent = record.name;
    const clear = document.createElement("button");
    clear.type = "button";
    clear.className = "slot-clear";
    clear.textContent = "×";
    clear.setAttribute("aria-label", `Remove ${record.name}`);
    clear.addEventListener("click", e => {
      e.stopPropagation();
      setSlot(team, side, i, null);
    });
    slot.append(img, tag, clear);
    slot.title = record.name;
  } else if (picked) {
    slot.classList.add("filled", "missing");
    slot.textContent = `${picked}?`;
    slot.title = `${picked} is not in animus.xlsx`;
  } else {
    const plus = document.createElement("span");
    plus.className = "plus";
    plus.textContent = "+";
    slot.append(plus);
  }

  slot.addEventListener("click", () => setTarget(team, side, i));

  slot.addEventListener("dragover", e => {
    if (!dragged) return;
    e.preventDefault();
    slot.classList.add("dropping");
  });
  slot.addEventListener("dragleave", () => slot.classList.remove("dropping"));
  slot.addEventListener("drop", e => {
    e.preventDefault();
    slot.classList.remove("dropping");
    if (dragged) setSlot(team, side, i, dragged);
  });

  return slot;
}

function resultButton(team, kind, label) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = `res-btn ${kind}${team.result === kind ? " on" : ""}`;
  btn.textContent = label;
  btn.setAttribute("aria-pressed", String(team.result === kind));
  btn.addEventListener("click", () => {
    team.result = team.result === kind ? null : kind;
    renderBoard();
  });
  return btn;
}

function updateTally() {
  el.tallyWin.textContent  = state.teams.filter(t => t.result === "win").length;
  el.tallyLoss.textContent = state.teams.filter(t => t.result === "loss").length;
}

/* ---------------- placing ---------------- */

function isTarget(team, side, slot) {
  const t = state.target;
  return !!t && t.team === team && t.side === side && t.slot === slot;
}

function setTarget(team, side, slot) {
  state.target = isTarget(team, side, slot) ? null : { team, side, slot };
  renderBoard();
  renderLibrary();
  updateTargetNote();
  if (state.target && window.matchMedia("(max-width:900px)").matches) openDrawer();
}

function clearTarget() {
  state.target = null;
  updateTargetNote();
}

function updateTargetNote() {
  const t = state.target;
  el.targetNote.classList.remove("enemy-target", "mine-target");
  if (!t) {
    el.targetNote.textContent = "No slot selected";
    return;
  }
  const index = state.teams.indexOf(t.team) + 1;
  el.targetNote.textContent =
    `${t.side === "enemy" ? "Enemy" : "My team"} · Team ${index} · slot ${t.slot + 1}`;
  el.targetNote.classList.add(t.side === "enemy" ? "enemy-target" : "mine-target");
}

function setSlot(team, side, slot, name) {
  const list = team[side === "enemy" ? "enemies" : "mine"];
  if (name && list.some((n, i) => i !== slot && sameName(n, name))) return;
  list[slot] = name;
  advance(team, side, slot, name);
  renderBoard();
  renderLibrary();
  updateTargetNote();
}

function sameName(a, b) {
  return a && b && String(a).toLowerCase() === String(b).toLowerCase();
}

/* after filling a slot, move to the next empty one in the same trio */
function advance(team, side, slot, name) {
  if (!name) { state.target = { team, side, slot }; return; }
  const list = team[side === "enemy" ? "enemies" : "mine"];
  for (let step = 1; step <= 2; step++) {
    const next = (slot + step) % 3;
    if (!list[next]) { state.target = { team, side, slot: next }; return; }
  }
  state.target = null;
}

function place(name) {
  const t = state.target;
  if (!t) {
    el.hint.textContent = "Pick a slot on the board first, then choose an Animus.";
    return;
  }
  el.hint.textContent = "Tap a slot, then pick an Animus from the library.";
  setSlot(t.team, t.side, t.slot, name);
}

/* ---------------- library ---------------- */

let dragged = null;

function renderLibrary() {
  const query = el.search.value.trim().toLowerCase();
  const t = state.target;
  const used = t ? t.team[t.side === "enemy" ? "enemies" : "mine"] : [];

  el.grid.replaceChildren(...animus.map(record => {
    const card = document.createElement("div");
    card.className = "pick";
    card.setAttribute("role", "button");
    if (query && !record.name.toLowerCase().includes(query)) card.classList.add("hidden");

    const taken = used.some((n, i) => sameName(n, record.name) && (!t || i !== t.slot));
    if (taken) {
      card.classList.add("taken");
      card.setAttribute("aria-disabled", "true");
      card.title = "Already in this team";
    } else {
      card.tabIndex = 0;
    }

    const img = document.createElement("img");
    img.src = IMAGE_PATH + record.file;
    img.alt = "";
    img.loading = "lazy";
    img.draggable = false;

    const label = document.createElement("span");
    label.textContent = record.name;

    card.append(img, label);
    card.setAttribute("aria-label", record.name);

    if (!taken) {
      card.draggable = true;
      card.addEventListener("dragstart", e => {
        dragged = record.name;
        e.dataTransfer.effectAllowed = "copy";
        e.dataTransfer.setData("text/plain", record.name);
      });
      card.addEventListener("dragend", () => { dragged = null; });
      const choose = () => {
        place(record.name);
        if (!state.target && window.matchMedia("(max-width:900px)").matches) closeDrawer();
      };
      card.addEventListener("click", choose);
      card.addEventListener("keydown", e => {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); choose(); }
      });
    }
    return card;
  }));

  const visible = el.grid.querySelectorAll(".pick:not(.hidden)").length;
  if (animus.length && !visible) {
    el.status.hidden = false;
    el.status.classList.remove("error");
    el.status.textContent = `No Animus matches “${el.search.value.trim()}”.`;
  } else {
    el.status.hidden = true;
  }
}

/* ---------------- drawer ---------------- */

function openDrawer() {
  el.library.classList.add("open");
  el.scrim.hidden = false;
}
function closeDrawer() {
  el.library.classList.remove("open");
  el.scrim.hidden = true;
}

el.drawerToggle.addEventListener("click", openDrawer);
el.drawerClose.addEventListener("click", closeDrawer);
el.scrim.addEventListener("click", closeDrawer);
document.addEventListener("keydown", e => {
  if (e.key !== "Escape") return;
  closeDrawer();
  if (state.target) { clearTarget(); renderBoard(); renderLibrary(); }
});

/* ---------------- controls ---------------- */

el.search.addEventListener("input", renderLibrary);

el.addTeam.addEventListener("click", () => {
  state.teams.push({
    id: state.teams.length + 1,
    enemies: [null, null, null],
    mine: [null, null, null],
    result: null
  });
  renderBoard();
});

el.reset.addEventListener("click", () => {
  state.teams.forEach(team => {
    team.mine = [null, null, null];
    team.result = null;
  });
  clearTarget();
  renderBoard();
  renderLibrary();
});

/* ---------------- start ---------------- */

renderBoard();

loadAnimus()
  .then(() => {
    el.status.hidden = true;
    renderBoard();
    renderLibrary();
  })
  .catch(err => {
    el.status.hidden = false;
    el.status.classList.add("error");
    el.status.textContent =
      "Could not read data/animus.xlsx (" + err.message +
      "). Open the site over http — a local file:// page cannot load the spreadsheet.";
  });
