const RANKINGS = [
  ["1.0", "First timer", "Bagong-bago; hindi pa marunong ng proper stance, bridge, aiming, cue-ball control, rules."],
  ["1.5", "Baguhan", "Marunong nang tumira pero madalas sablay; walang consistent stroke at positioning."],
  ["2.0", "Beginner", "Nakakapagpasok ng easy shots; basic rules alam na; cue ball madalas walang control."],
  ["2.5", "Advanced beginner", "Kaya nang mag-run ng 1–2 balls minsan; may konting aiming at safety awareness."],
  ["3.0", "Low amateur", "May basic pattern play; kaya nang manalo sa casual players; inconsistent pa sa long shots at cue-ball position."],
  ["3.5", "Amateur", "Mas consistent sa pocketing; marunong na sa basic english/spin, stop shot, follow, draw."],
  ["4.0", "Club player", "Regular na player sa bilyaran; kaya nang mag-run ng ilang balls; may defensive shots."],
  ["4.5", "Strong club player", "Mahirap talunin ng casual players; marunong sa pattern, safety, kick/bank basics."],
  ["5.0", "Intermediate competitive", "Kaya nang sumali sa local tournaments; may run-out ability kapag open table."],
  ["5.5", "Strong intermediate", "Consistent sa 8-ball/9-ball; kaya nang magbigay ng handicap sa lower players."],
  ["6.0", "Advanced local player", "Madalas nasa top spots sa local tournaments; may strong break, safety, at cue-ball control."],
  ["6.5", "Strong advanced", "Kaya nang lumaban sa kilalang local hustlers/tournament players; may regular run-outs."],
  ["7.0", "Semi-pro / elite local", "Isa sa malalakas sa area/city; kaya nang lumaban sa regional tournaments."],
  ["7.5", "Regional top player", "Top-level provincial/city player; may laban sa national-level players."],
  ["8.0", "National-level player", "Kaya nang lumaban sa big Philippine tournaments; very consistent sa pressure games."],
  ["8.5", "National elite", "Kabilang sa mga pinakamalalakas sa bansa; kaya makipagsabayan sa international-caliber players."],
  ["9.0", "Top amateur / lower pro", "Halos pro level; puwedeng manalo sa high-level events; bihira ang unforced errors."],
  ["9.5", "Pro / international level", "Professional-caliber; may mataas na run-out percentage at strong match strategy."],
  ["10.0", "World-class pro", "Pinakamataas na level; world champion / world tour caliber, gaya ng level ng elite Filipino pros."]
];

const defaultDb = {
  Users: [],
  PlayerProfiles: [],
  MatchRequests: [],
  BilliardHalls: [],
  Rankings: RANKINGS.map(([Rating, Level, Description]) => ({ Rating, Level, Description }))
};
let db = structuredClone(defaultDb);
let activeWorkbook = null;

const $ = (id) => document.getElementById(id);
const today = () => new Date().toISOString().slice(0, 10);
const safe = (value) => String(value ?? "").replace(/[&<>'"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;","\"":"&quot;"}[c]));
const rankInfo = (rating) => db.Rankings.find(r => String(r.Rating) === String(rating)) || db.Rankings[0];

function toast(message) {
  const el = $("toast");
  el.textContent = message;
  el.classList.add("show");
  setTimeout(() => el.classList.remove("show"), 2800);
}
function scrollToSection(id) { document.getElementById(id)?.scrollIntoView({ behavior: "smooth" }); }
function makeId(prefix) { return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`; }

function normalizeSheetRows(rows) {
  return rows.map(row => Object.fromEntries(Object.entries(row).map(([k, v]) => [String(k).trim(), v])));
}

function workbookToDb(workbook) {
  const names = workbook.SheetNames;
  const next = structuredClone(defaultDb);
  for (const name of names) {
    const sheet = workbook.Sheets[name];
    const rows = normalizeSheetRows(XLSX.utils.sheet_to_json(sheet, { defval: "" }));
    const normalizedName = name.replace(/\s+/g, "");
    if (next[normalizedName]) next[normalizedName] = rows;
  }
  if (!next.Rankings.length) next.Rankings = defaultDb.Rankings;
  return next;
}

async function loadInitialWorkbook() {
  const stored = localStorage.getItem("emmp-db");
  if (stored) {
    db = JSON.parse(stored);
    render();
    toast("Loaded saved browser database.");
    return;
  }
  try {
    const res = await fetch("data/BilliardMatchDatabase.xlsx", { cache: "no-store" });
    if (!res.ok) throw new Error("Workbook not found");
    const buffer = await res.arrayBuffer();
    activeWorkbook = XLSX.read(buffer, { type: "array" });
    db = workbookToDb(activeWorkbook);
  } catch (err) {
    console.warn(err);
    seedDemoData();
  }
  render();
}

function seedDemoData() {
  db = structuredClone(defaultDb);
  db.Users = [
    { UserId: "U-1001", FullName: "Juan Dela Cruz", Monicker: "Dodong Break", Email: "juan@example.com", MobileNumber: "09170000001", FacebookUrl: "https://facebook.com/", CreatedAt: today() },
    { UserId: "U-1002", FullName: "Mark Santos", Monicker: "Bank Shot Mark", Email: "mark@example.com", MobileNumber: "09170000002", FacebookUrl: "https://facebook.com/", CreatedAt: today() }
  ];
  db.PlayerProfiles = [
    { ProfileId: "P-1001", UserId: "U-1001", Rating: "4.5", RatingLevel: "Strong club player", PreferredBilliardHall: "Star Billiards", HallLocation: "Quezon City", PreferredGame: "9-ball", PreferredStake: "₱1,000 race to 5", Negotiable: "Yes", Status: "Looking for match", Bio: "Weekends only." },
    { ProfileId: "P-1002", UserId: "U-1002", Rating: "5.5", RatingLevel: "Strong intermediate", PreferredBilliardHall: "Break Room", HallLocation: "Manila", PreferredGame: "10-ball", PreferredStake: "Negotiable", Negotiable: "Yes", Status: "Available today", Bio: "Open to handicap games." }
  ];
  db.BilliardHalls = [
    { HallId: "H-1001", Name: "Star Billiards", Location: "Quezon City" },
    { HallId: "H-1002", Name: "Break Room", Location: "Manila" }
  ];
}

function saveBrowserDb() { localStorage.setItem("emmp-db", JSON.stringify(db)); }
function resetBrowserDb() { localStorage.removeItem("emmp-db"); location.reload(); }

function setupRankingControls() {
  const options = db.Rankings.map(r => `<option value="${safe(r.Rating)}">${safe(r.Rating)} - ${safe(r.Level)}</option>`).join("");
  $("ratingSelect").innerHTML = options;
  $("ratingMin").innerHTML = `<option value="">Min rating</option>${options}`;
  $("ratingMax").innerHTML = `<option value="">Max rating</option>${options}`;
  $("rankingList").innerHTML = db.Rankings.map(r => `<div class="rank-item"><strong>${safe(r.Rating)} - ${safe(r.Level)}</strong><small>${safe(r.Description)}</small></div>`).join("");
}

function joinedPlayers() {
  return db.PlayerProfiles.map(profile => {
    const user = db.Users.find(u => String(u.UserId) === String(profile.UserId)) || {};
    return { ...profile, ...user };
  });
}

function renderPlayers() {
  const q = $("searchInput").value.toLowerCase().trim();
  const min = parseFloat($("ratingMin").value || "0");
  const max = parseFloat($("ratingMax").value || "10");
  const game = $("gameFilter").value;
  const negotiable = $("negotiableOnly").checked;
  const players = joinedPlayers().filter(p => {
    const rating = parseFloat(p.Rating || "0");
    const haystack = `${p.FullName} ${p.Monicker} ${p.PreferredBilliardHall} ${p.HallLocation} ${p.PreferredGame}`.toLowerCase();
    return (!q || haystack.includes(q)) && rating >= min && rating <= max && (!game || p.PreferredGame === game) && (!negotiable || String(p.Negotiable).toLowerCase() === "yes");
  });
  $("playerCards").innerHTML = players.length ? players.map(playerCard).join("") : `<p class="meta">No players match the current filters.</p>`;
  $("opponentSelect").innerHTML = joinedPlayers().map(p => `<option value="${safe(p.UserId)}">${safe(p.Monicker || p.FullName)} - ${safe(p.Rating)} ${safe(p.RatingLevel || "")}</option>`).join("");
}

function playerCard(p) {
  const fb = p.FacebookUrl || p.facebook || "#";
  return `<article class="player-card">
    <h3>${safe(p.Monicker || p.FullName || "Player")}</h3>
    <div class="chips"><span class="chip">${safe(p.Rating)} - ${safe(p.RatingLevel || rankInfo(p.Rating)?.Level)}</span><span class="chip">${safe(p.Status || "Looking for match")}</span></div>
    <p class="meta"><strong>Hall:</strong> ${safe(p.PreferredBilliardHall || "Not specified")}<br><strong>Location:</strong> ${safe(p.HallLocation || "Not specified")}<br><strong>Game:</strong> ${safe(p.PreferredGame || "Any")}<br><strong>Stake:</strong> ${safe(p.PreferredStake || "Negotiable")} ${String(p.Negotiable).toLowerCase() === "yes" ? "(negotiable)" : ""}</p>
    <p class="meta">${safe(p.Bio || "")}</p>
    <div class="card-actions"><a class="btn primary" href="${safe(fb)}" target="_blank" rel="noopener">Message on Facebook</a><button class="btn" onclick="preselectOpponent('${safe(p.UserId)}')">Challenge</button></div>
  </article>`;
}

function preselectOpponent(userId) { $("opponentSelect").value = userId; scrollToSection("requests"); }

function renderStats() {
  $("statPlayers").textContent = db.PlayerProfiles.length;
  $("statLooking").textContent = db.PlayerProfiles.filter(p => /looking|available/i.test(p.Status || "")).length;
  $("statHalls").textContent = new Set(db.PlayerProfiles.map(p => `${p.PreferredBilliardHall}|${p.HallLocation}`).filter(Boolean)).size;
  $("statRequests").textContent = db.MatchRequests.length;
  const featured = joinedPlayers()[0];
  if (featured) {
    $("featuredName").textContent = featured.Monicker || featured.FullName;
    $("featuredRating").textContent = `${featured.Rating} - ${featured.RatingLevel || rankInfo(featured.Rating)?.Level}`;
    $("featuredChips").innerHTML = [featured.PreferredBilliardHall, featured.HallLocation, featured.PreferredGame, featured.PreferredStake].filter(Boolean).map(x => `<span class="chip">${safe(x)}</span>`).join("");
  } else {
    $("featuredName").textContent = "No players yet";
    $("featuredRating").textContent = "Create the first profile";
    $("featuredChips").innerHTML = `<span class="chip">Excel only</span><span class="chip">Free hosting ready</span>`;
  }
}

function renderRequests() {
  $("requestList").innerHTML = db.MatchRequests.length ? db.MatchRequests.slice().reverse().map(r => {
    const opponent = joinedPlayers().find(p => String(p.UserId) === String(r.OpponentUserId));
    return `<div class="request"><strong>${safe(r.GameType || "Match")} vs ${safe(opponent?.Monicker || opponent?.FullName || "Opponent")}</strong><p class="meta">${safe(r.RaceFormat || "")} · ${safe(r.ProposedLocation || "No location yet")} · ${safe(r.ProposedStake || "Negotiable")}<br>${safe(r.Message || "")}</p></div>`;
  }).join("") : `<p class="meta">No match requests yet.</p>`;
}

function render() { setupRankingControls(); renderPlayers(); renderStats(); renderRequests(); }

function handleSignup(event) {
  event.preventDefault();
  const form = new FormData(event.target);
  const rating = form.get("rating");
  const rank = rankInfo(rating);
  const userId = makeId("U");
  db.Users.push({
    UserId: userId,
    FullName: form.get("fullName"),
    Monicker: form.get("monicker"),
    Email: form.get("email"),
    MobileNumber: form.get("mobile"),
    FacebookUrl: form.get("facebook"),
    CreatedAt: today()
  });
  db.PlayerProfiles.push({
    ProfileId: makeId("P"),
    UserId: userId,
    Rating: rating,
    RatingLevel: rank?.Level || "",
    RatingDescription: rank?.Description || "",
    PreferredBilliardHall: form.get("hall"),
    HallLocation: form.get("location"),
    PreferredGame: form.get("game"),
    PreferredStake: form.get("stake") || "Negotiable",
    Negotiable: form.get("negotiable") ? "Yes" : "No",
    Status: form.get("status"),
    Bio: form.get("bio")
  });
  db.BilliardHalls.push({ HallId: makeId("H"), Name: form.get("hall"), Location: form.get("location") });
  saveBrowserDb();
  event.target.reset();
  render();
  toast("Player profile saved. Export Excel to publish it.");
}

function handleRequest(event) {
  event.preventDefault();
  const form = new FormData(event.target);
  db.MatchRequests.push({
    RequestId: makeId("M"),
    ChallengerUserId: "ManualEntry",
    OpponentUserId: form.get("opponent"),
    GameType: form.get("game"),
    RaceFormat: form.get("race"),
    ProposedLocation: form.get("location"),
    ProposedDateTime: form.get("datetime"),
    ProposedStake: form.get("stake"),
    Negotiable: form.get("negotiable") ? "Yes" : "No",
    Message: form.get("message"),
    Status: "Pending",
    CreatedAt: new Date().toISOString()
  });
  saveBrowserDb();
  event.target.reset();
  render();
  toast("Match request saved to browser database.");
}

function exportExcel() {
  const wb = XLSX.utils.book_new();
  const sheetOrder = ["Users", "PlayerProfiles", "MatchRequests", "BilliardHalls", "Rankings"];
  sheetOrder.forEach(name => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(db[name] || []), name));
  XLSX.writeFile(wb, "BilliardMatchDatabase.xlsx");
}

function downloadBlankTemplate() {
  const oldDb = db;
  db = structuredClone(defaultDb);
  exportExcel();
  db = oldDb;
}

function importExcel(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const wb = XLSX.read(e.target.result, { type: "array" });
    db = workbookToDb(wb);
    saveBrowserDb();
    render();
    toast("Excel database imported.");
  };
  reader.readAsArrayBuffer(file);
}

["searchInput", "ratingMin", "ratingMax", "gameFilter", "negotiableOnly"].forEach(id => $(id).addEventListener("input", renderPlayers));
$("signupForm").addEventListener("submit", handleSignup);
$("requestForm").addEventListener("submit", handleRequest);
$("excelFile").addEventListener("change", (e) => e.target.files[0] && importExcel(e.target.files[0]));
loadInitialWorkbook();
