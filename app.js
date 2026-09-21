"use strict";

// ═══════════ Araç Tipleri — iç ölçüler cm ═══════════
const VEHICLE_TYPES = [
  { id: "tir",      name: "TIR 13.6m",          w: 245, l: 1360, h: 270, capacity: 25000 },
  { id: "c40hc",    name: "40HC Konteyner",     w: 235, l: 1203, h: 269, capacity: 26000 },
  { id: "c20dc",    name: "20DC Konteyner",     w: 235, l: 589,  h: 239, capacity: 21000 },
  { id: "c45hc",    name: "45HC Konteyner",     w: 235, l: 1355, h: 269, capacity: 27000 },
  { id: "kamyon10", name: "Kamyon 10 Teker",    w: 245, l: 720,  h: 250, capacity: 10000 },
  { id: "parsiyel", name: "Parsiyel",           w: 0, l: 0, h: 0, capacity: 0 },
  { id: "custom",   name: "Özel Araç",          w: 0, l: 0, h: 0, capacity: 0 },
];

const DEMO = [
  { UrunKod: "Yatak",      En: 160, Boy: 200, Yukseklik: 30, Adet: 50, Kural: "STANDART", Istif: "", Yukleme: "BOYUNA" },
  { UrunKod: "BazaBaslik", En: 90,  Boy: 200, Yukseklik: 20, Adet: 10, Kural: "STANDART", Istif: "", Yukleme: "ENINE" },
  { UrunKod: "Komodin",    En: 50,  Boy: 50,  Yukseklik: 40, Adet: 20, Kural: "STANDART", Istif: "", Yukleme: "" },
  { UrunKod: "Markiz",     En: 40,  Boy: 120, Yukseklik: 40, Adet: 10, Kural: "STANDART", Istif: "", Yukleme: "" },
  { UrunKod: "Ped",        En: 180, Boy: 200, Yukseklik: 8,  Adet: 20, Kural: "DIK",      Istif: "", Yukleme: "" },
];

const state = { packages: [], result: null, activeVehicle: 0 };
let editingIndex = -1;
const $ = (id) => document.getElementById(id);

function capVolume() {
  const s = state.result?.spec;
  return s ? (s.w * s.h * s.l) / 1e6 : 0;
}

// ═══════════ Manuel giriş ═══════════
function resetForm() {
  editingIndex = -1;
  $("mUrunKod").value = "";
  $("mEn").value = ""; $("mBoy").value = ""; $("mYukseklik").value = "";
  $("mAdet").value = 1;
  $("mKural").value = "STANDART";
  $("mIstif").value = "";
  $("mYukleme").value = "";
  $("addBtn").textContent = "+ Ekle";
  $("cancelEditBtn").classList.add("hidden");
  $("mUrunKod").focus();
}

function startEdit(i) {
  const p = state.packages[i];
  $("mUrunKod").value = p.UrunKod;
  $("mEn").value = p.En;
  $("mBoy").value = p.Boy;
  $("mYukseklik").value = p.Yukseklik;
  $("mAdet").value = p.Adet;
  $("mKural").value = p.Kural || "STANDART";
  $("mIstif").value = p.Istif || "";
  $("mYukleme").value = p.Yukleme || "";
  editingIndex = i;
  $("addBtn").textContent = "✔ Güncelle";
  $("cancelEditBtn").classList.remove("hidden");
  $("mUrunKod").focus();
}

 $("addBtn").addEventListener("click", () => {
  const code = $("mUrunKod").value.trim();
  const w = num($("mEn").value), l = num($("mBoy").value), h = num($("mYukseklik").value);
  if (!code || !w || !l || !h) { alert("Ürün kodu ile En/Boy/Yükseklik alanları zorunludur."); return; }
  const qty = Math.max(1, Math.round(num($("mAdet").value) || 1));
  const pkg = {
    UrunKod: code, En: w, Boy: l, Yukseklik: h, Adet: qty,
    Kural: $("mKural").value, Istif: $("mIstif").value,
    Yukleme: $("mYukleme").value,
  };
  if (editingIndex >= 0) state.packages[editingIndex] = pkg;
  else state.packages.push(pkg);
  resetForm();
  renderTable();
});

 $("cancelEditBtn").addEventListener("click", resetForm);

document.querySelector(".mform").addEventListener("keydown", (e) => {
  if (e.key === "Enter" && e.target.tagName !== "SELECT") { e.preventDefault(); $("addBtn").click(); }
});

// ═══════════ Excel şablonu indir ═══════════
 $("tplBtn").addEventListener("click", () => {
  if (typeof XLSX === "undefined") { alert("Excel kütüphanesi yüklenemedi (internet engeli olabilir)."); return; }
  const wb = XLSX.utils.book_new();
  const data = [
    ["UrunKod", "En", "Boy", "Yukseklik", "Adet", "Kural", "Istif", "Yukleme"],
    ["Yatak", 160, 200, 30, 5, "", "", "BOYUNA"],
    ["BazaBaslik", 90, 200, 20, 5, "", "", "ENINE"],
    ["Komodin", 180, 110, 16, 10, "", "", ""],
    ["Markiz", 40, 40, 40, 5, "KIRILGAN", "", ""],
  ];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [{ wch: 14 }, { wch: 8 }, { wch: 8 }, { wch: 11 }, { wch: 7 }, { wch: 12 }, { wch: 8 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, ws, "YuklemeListesi");
  const help = [
    ["ALAN", "AÇIKLAMA"],
    ["UrunKod", "Paket kodu (zorunlu)"],
    ["En / Boy / Yukseklik", "Paket ölçüleri cm (zorunlu)"],
    ["Adet", "Kaç adet yüklenecek (boşsa 1)"],
    ["Kural", "STANDART / KIRILGAN / DIK / YERDE (boşsa STANDART)"],
    ["Istif", "Üst üste maksimum katman sayısı (boşsa OTOMATİK: araç yüksekliğine göre hesaplanır)"],
    ["Yukleme", "ENINE / BOYUNA (boşsa serbest dönüş)"],
    [],
    ["KURAL / YÖN", "ANLAMI"],
    ["STANDART", "Serbest döner, istiflenebilir"],
    ["KIRILGAN", "Üstüne hiçbir paket konamaz"],
    ["DIK", "Döndürülemez, sadece dik durur"],
    ["YERDE", "En alt katmanda durur"],
    ["ENINE", "Paketin BOY ölçüsü aracın genişliğine yatar"],
    ["BOYUNA", "Paketin BOY ölçüsü aracın uzunluğuna paralel durur"],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(help);
  ws2["!cols"] = [{ wch: 22 }, { wch: 52 }];
  XLSX.utils.book_append_sheet(wb, ws2, "Aciklama");
  XLSX.writeFile(wb, "yukleme-listesi-sablon.xlsx");
});

// ═══════════ Excel okuma ═══════════
const COLS = {
  UrunKod: ["urunkod", "urunkodu", "urun", "urunadi", "kod"],
  En: ["en", "genislik"],
  Boy: ["boy", "uzunluk"],
  Yukseklik: ["yukseklik", "yukseklikcm", "yuk"],
  Adet: ["adet", "miktar", "qty"],
  Kural: ["kural", "yuklemekurali"],
  Istif: ["istif", "maxistif", "istifsayisi"],
  Yukleme: ["yukleme", "yuklemeyonu", "yukyonu", "yuklemesekli", "yon"],
};

// Türkçe karakterleri ayıkla: "Yükleme" → "yukleme" gibi eşleşsin
function normKey(k) {
  return k.toLowerCase()
    .replace(/ı/g, "i").replace(/İ/g, "i").replace(/ş/g, "s").replace(/Ş/g, "s")
    .replace(/ğ/g, "g").replace(/Ğ/g, "g").replace(/ü/g, "u").replace(/Ü/g, "u")
    .replace(/ö/g, "o").replace(/Ö/g, "o").replace(/ç/g, "c").replace(/Ç/g, "c")
    .replace(/[\s_.]/g, "");
}

function pickCol(row, candidates) {
  const keys = Object.keys(row);
  for (const c of candidates) {
    const k = keys.find((key) => normKey(key) === c);
    if (k !== undefined && String(row[k]).trim() !== "") return String(row[k]).trim();
  }
  return "";
}

function num(v) {
  const n = parseFloat(String(v).replace(",", ".")); // TR ondalık virgülü
  return isNaN(n) ? 0 : n;
}

function normDir(v) {
  const s = String(v).trim().toUpperCase().replace(/İ/g, "I").replace(/Ş/g, "S");
  if (!s) return "";
  if (s.includes("ENINE") || s === "EN") return "ENINE";
  if (s.includes("BOY")) return "BOYUNA";
  return "";
}

function parseRows(rows) {
  const good = [], errors = [];
  rows.forEach((row, idx) => {
    const code = pickCol(row, COLS.UrunKod);
    const w = num(pickCol(row, COLS.En));
    const l = num(pickCol(row, COLS.Boy));
    const h = num(pickCol(row, COLS.Yukseklik));
    const qty = Math.round(num(pickCol(row, COLS.Adet)) || 1);
    let rule = pickCol(row, COLS.Kural).toUpperCase().replace(/İ/g, "I").replace(/Ş/g, "S");
    if (!["STANDART", "KIRILGAN", "DIK", "YERDE"].includes(rule)) rule = "STANDART";
    const stack = pickCol(row, COLS.Istif);
    const yukleme = normDir(pickCol(row, COLS.Yukleme));
    if (!code || !w || !l || !h) {
      errors.push(`Satır ${idx + 2}: eksik/hatalı veri (Ürün kodu + En/Boy/Yükseklik gerekli)`);
      return;
    }
    good.push({ UrunKod: code, En: w, Boy: l, Yukseklik: h, Adet: qty, Kural: rule, Istif: stack, Yukleme: yukleme });
  });
  return { good, errors };
}

function handleFile(file) {
  if (typeof XLSX === "undefined") { alert("Excel kütüphanesi yüklenemedi (internet engeli olabilir)."); return; }
  const r = new FileReader();
  r.onload = (ev) => {
    try {
      const wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
      const { good, errors } = parseRows(rows);
      if (!good.length) {
        showErrors([...errors, "Geçerli satır yok. İlk satır başlık olmalı: UrunKod, En, Boy, Yukseklik, Adet, Kural, Istif, Yukleme"]);
        return;
      }
      state.packages = state.packages.concat(good);
      resetForm();
      renderTable();
      showErrors(errors);
    } catch (err) { alert("Dosya okunamadı: " + err.message); }
  };
  r.readAsArrayBuffer(file);
}

function showErrors(list) {
  $("parseErrors").innerHTML = list.map((e) => `<div class="parseErr">⚠ ${e}</div>`).join("");
}

// ═══════════ Tablo ═══════════
function renderTable() {
  const tb = $("pkgBody");
  tb.innerHTML = "";
  let total = 0;
  state.packages.forEach((p, i) => {
    total += p.Adet;
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${p.UrunKod}</td><td>${p.En}</td><td>${p.Boy}</td><td>${p.Yukseklik}</td>` +
      `<td>${p.Adet}</td><td>${p.Kural}</td><td>${p.Istif || "oto"}</td><td>${p.Yukleme || "-"}</td>` +
      `<td><button class="edit" data-i="${i}" title="Düzenle">✎</button> ` +
      `<button class="del" data-i="${i}" title="Sil">✕</button></td>`;
    tb.appendChild(tr);
  });
  $("pkgCount").textContent = state.packages.length
    ? `${state.packages.length} kalem • toplam ${total} paket` : "";
  $("tableWrap").classList.toggle("hidden", !state.packages.length);
}

 $("pkgBody").addEventListener("click", (e) => {
  const i = +e.target.dataset.i;
  if (e.target.classList.contains("del")) {
    state.packages.splice(i, 1);
    if (editingIndex === i) resetForm();
    renderTable();
  } else if (e.target.classList.contains("edit")) {
    startEdit(i);
    $("tableWrap").scrollIntoView({ behavior: "smooth", block: "nearest" });
  }
});

// ═══════════ Simülasyon ═══════════
function runSim() {
  if (!state.packages.length) { alert("Önce yükleme listesine ürün ekleyin (manuel veya Excel)."); return; }
  const vi = +$("vehicleSelect").value;
  let spec = { ...VEHICLE_TYPES[vi] };
  if (spec.id === "parsiyel" || spec.id === "custom") {
    spec = {
      id: spec.id,
      name: spec.id === "custom" ? ($("cname").value.trim() || "Özel Araç") : "Parsiyel",
      w: +$("cw").value, l: +$("cl").value, h: +$("ch").value, capacity: 0,
    };
    if (!spec.w || !spec.l || !spec.h) { alert("Lütfen araç iç ölçülerini (En/Boy/Yükseklik) girin."); return; }
  }
  const count = Math.min(20, Math.max(1, Math.round(+$("vCount").value || 1)));
  $("vCount").value = count;

  const totalItems = state.packages.reduce((s, p) => s + p.Adet, 0);
  if (totalItems > 2000 && !confirm(`${totalItems} paket hesaplanacak, tarayıcı yavaşlayabilir. Devam edilsin mi?`)) return;

  state.result = PackingEngine.simulate(state.packages, spec, count);
  state.activeVehicle = 0;
  renderResults();
  buildScene();
}

function renderResults() {
  const res = state.result;
  const tabs = $("vehicleTabs");
  tabs.innerHTML = "";
  res.vehicles.forEach((v, i) => {
    const b = document.createElement("button");
    b.className = "tab" + (i === state.activeVehicle ? " active" : "");
    b.textContent = `Araç ${i + 1} • %${res.stats[i].fillRate}`;
    b.onclick = () => { state.activeVehicle = i; renderResults(); buildScene(); };
    tabs.appendChild(b);
  });

  const s = res.stats[state.activeVehicle];
  const placed = res.vehicles.reduce((a, v) => a + v.placements.length, 0);
  $("capStat").textContent = capVolume().toFixed(1);
  $("vehInfo").textContent =
    `${res.spec.name} — iç ölçü ${res.spec.w}×${res.spec.l}×${res.spec.h} cm — ` +
    `dolum: ön sütunlar yukarı istiflenir, arkaya (kapıya) doğru dolar`;
  $("statsRest").innerHTML =
    `<span>🚛 Araç ${state.activeVehicle + 1}: <b>${s.count}</b> paket</span>` +
    `<span>Toplam yüklendi: <b>${placed}</b></span>` +
    `<span>Yüklenemeyen: <b>${res.unplaced.length}</b></span>`;

  const names = [...new Set([
    ...res.vehicles.flatMap((v) => v.placements.map((p) => p.pkg.name)),
    ...res.unplaced.map((u) => u.name),
  ])];
  $("legend").innerHTML = names.map((n) =>
    `<span class="chip"><span class="swatch" style="background:#${colorFor(n).getHexString()}"></span>${n}</span>`
  ).join("");

  const ub = $("unplacedBox");
  if (res.unplaced.length) {
    ub.classList.remove("hidden");
    const g = {};
    res.unplaced.forEach((u) => (g[u.name] = (g[u.name] || 0) + 1));
    $("unplacedList").innerHTML =
      Object.entries(g).map(([n, c]) => `<b>${n}</b> × ${c}`).join(" • ") +
      `<br><span style="color:#fca5a5">Bu ürünler hiçbir araca sığmadı — araç adedini artırmayı deneyin.</span>`;
  } else ub.classList.add("hidden");
}

// ═══════════ 3D Sahne (Three.js) ═══════════
let renderer = null, scene, camera, boxGroup, sceneGroup, vehLabel = null;
let meshes = [];
let camCtl = { theta: Math.PI / 4, phi: 1.05, radius: 800, center: { x: 0, y: 0, z: 0 } };
const raycaster = new THREE.Raycaster();
const mouseV = new THREE.Vector2();
const colorCache = {};

function colorFor(name) {
  if (colorCache[name]) return colorCache[name];
  let h = 0;
  for (const c of name) h = (h * 31 + c.charCodeAt(0)) % 360;
  return (colorCache[name] = new THREE.Color().setHSL(h / 360, 0.65, 0.55));
}

function initThree() {
  const wrap = $("canvasWrap");
  try { renderer = new THREE.WebGLRenderer({ antialias: true }); }
  catch (e) {
    wrap.innerHTML = '<div style="padding:30px;color:#f87171">3D başlatılamadı — tarayıcıda WebGL kapalı olabilir.</div>';
    return false;
  }
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
  renderer.setSize(wrap.clientWidth, wrap.clientHeight);
  wrap.appendChild(renderer.domElement);

  scene = new THREE.Scene();
  scene.background = new THREE.Color(0x0b1220);
  camera = new THREE.PerspectiveCamera(50, wrap.clientWidth / wrap.clientHeight, 1, 200000);
  scene.add(new THREE.HemisphereLight(0xffffff, 0x334155, 0.9));
  const dl = new THREE.DirectionalLight(0xffffff, 0.7);
  dl.position.set(1, 2, 1);
  scene.add(dl);
  sceneGroup = new THREE.Group(); scene.add(sceneGroup);
  boxGroup = new THREE.Group(); scene.add(boxGroup);

  bindControls(wrap);
  (function loop() { requestAnimationFrame(loop); renderer.render(scene, camera); })();

  window.addEventListener("resize", () => {
    renderer.setSize(wrap.clientWidth, wrap.clientHeight);
    camera.aspect = wrap.clientWidth / wrap.clientHeight;
    camera.updateProjectionMatrix();
  });
  return true;
}

function disposeGroup(g) {
  while (g.children.length) {
    const c = g.children.pop();
    c.traverse((o) => {
      if (o.geometry) o.geometry.dispose();
      if (o.material) {
        if (Array.isArray(o.material)) o.material.forEach((m) => { if (m.map) m.map.dispose(); m.dispose(); });
        else { if (o.material.map) o.material.map.dispose(); o.material.dispose(); }
      }
    });
  }
}

function textSprite(text, color, w, h) {
  const c = document.createElement("canvas");
  c.width = 512; c.height = 128;
  const g = c.getContext("2d");
  g.fillStyle = "rgba(11,18,32,0.85)";
  g.fillRect(0, 0, 512, 128);
  g.strokeStyle = color; g.lineWidth = 6; g.strokeRect(3, 3, 506, 122);
  g.font = "bold 60px system-ui"; g.fillStyle = color;
  g.textAlign = "center"; g.textBaseline = "middle";
  g.fillText(text, 256, 68);
  const sp = new THREE.Sprite(new THREE.SpriteMaterial({ map: new THREE.CanvasTexture(c), transparent: true }));
  sp.scale.set(w, h, 1);
  return sp;
}

function makeVehLabel(spec) {
  const c = document.createElement("canvas");
  c.width = 1200; c.height = 180;
  const tex = new THREE.CanvasTexture(c);
  const sp = new THREE.Sprite(new THREE.SpriteMaterial({ map: tex, transparent: true }));
  sp.scale.set(spec.w * 2.1, spec.w * 2.1 * (180 / 1200), 1);
  sp.position.set(spec.w / 2, spec.h + spec.w * 0.28, spec.l / 2);
  sp.userData = { canvas: c, tex };
  sceneGroup.add(sp);
  return sp;
}

function updateVehLabel(visibleCount) {
  if (!vehLabel) return;
  const { canvas, tex } = vehLabel.userData;
  const g = canvas.getContext("2d");
  g.clearRect(0, 0, canvas.width, canvas.height);
  g.fillStyle = "rgba(17,26,46,0.92)";
  g.fillRect(0, 0, canvas.width, canvas.height);
  g.strokeStyle = "#3b82f6"; g.lineWidth = 8;
  g.strokeRect(4, 4, canvas.width - 8, canvas.height - 8);

  let vol = 0;
  for (let i = 0; i < visibleCount && i < meshes.length; i++) vol += meshes[i].userData.vol;
  const cap = capVolume();
  const pct = cap ? ((vol / cap) * 100).toFixed(1) : "0";

  g.font = "bold 58px system-ui";
  g.textAlign = "center"; g.textBaseline = "middle";
  g.fillStyle = "#4ade80";
  g.fillText(`${vol.toFixed(1)} / ${cap.toFixed(1)} m³`, canvas.width / 2 - 200, canvas.height / 2);
  g.fillStyle = "#60a5fa";
  g.fillText(`Doluluk %${pct}`, canvas.width / 2 + 330, canvas.height / 2);
  g.fillStyle = "#e2e8f0";
  g.font = "bold 48px system-ui";
  g.fillText(`Araç ${state.activeVehicle + 1}`, 190, canvas.height / 2);
    // imza
  g.font = "bold 30px system-ui";
  g.textAlign = "right"; g.textBaseline = "alphabetic";
  g.fillStyle = "rgba(147,197,253,0.75)";
  g.fillText("Design by Sait", canvas.width - 30, canvas.height - 26);
  tex.needsUpdate = true;
}

function buildScene() {
  if (!renderer && !initThree()) return;
  const res = state.result;
  if (!res) return;
  const spec = res.spec;
  const v = res.vehicles[state.activeVehicle];

  disposeGroup(sceneGroup);
  disposeGroup(boxGroup);
  meshes = [];

  const floor = new THREE.Mesh(
    new THREE.PlaneGeometry(spec.w, spec.l),
    new THREE.MeshLambertMaterial({ color: 0x16223b })
  );
  floor.rotation.x = -Math.PI / 2;
  floor.position.set(spec.w / 2, 0, spec.l / 2);
  sceneGroup.add(floor);
  const grid = new THREE.GridHelper(Math.max(spec.w, spec.l), 10, 0x334155, 0x1e293b);
  grid.position.set(spec.w / 2, 0.5, spec.l / 2);
  sceneGroup.add(grid);

  const bgeo = new THREE.BoxGeometry(spec.w, spec.h, spec.l);
  const outline = new THREE.LineSegments(
    new THREE.EdgesGeometry(bgeo),
    new THREE.LineBasicMaterial({ color: 0x3b82f6, transparent: true, opacity: 0.6 })
  );
  outline.position.set(spec.w / 2, spec.h / 2, spec.l / 2);
  sceneGroup.add(outline);
  bgeo.dispose();

  const startStrip = new THREE.Mesh(
    new THREE.PlaneGeometry(spec.w, Math.min(8, spec.l * 0.05)),
    new THREE.MeshBasicMaterial({ color: 0x22c55e, transparent: true, opacity: 0.55, side: THREE.DoubleSide })
  );
  startStrip.rotation.x = -Math.PI / 2;
  startStrip.position.set(spec.w / 2, 0.8, Math.min(4, spec.l * 0.03));
  sceneGroup.add(startStrip);

  const doorGeo = new THREE.PlaneGeometry(spec.w, spec.h);
  const door = new THREE.LineSegments(
    new THREE.EdgesGeometry(doorGeo),
    new THREE.LineBasicMaterial({ color: 0xf97316 })
  );
  door.position.set(spec.w / 2, spec.h / 2, spec.l);
  sceneGroup.add(door);
  doorGeo.dispose();

  const arrowLen = Math.min(spec.l * 0.35, 250);
  sceneGroup.add(new THREE.ArrowHelper(
    new THREE.Vector3(0, 0, 1),
    new THREE.Vector3(spec.w / 2, 25, 15),
    arrowLen, 0x22c55e, arrowLen * 0.25, arrowLen * 0.12
  ));
  const lblW = spec.w * 0.85, lblH = spec.w * 0.22;
  const sFront = textSprite("ÖN • dolum başlar", "#4ade80", lblW, lblH);
  sFront.position.set(spec.w / 2, spec.h * 0.9, -lblW * 0.6);
  sceneGroup.add(sFront);
  const sBack = textSprite("ARKA • KAPI", "#fb923c", lblW, lblH);
  sBack.position.set(spec.w / 2, spec.h * 0.9, spec.l + lblW * 0.6);
  sceneGroup.add(sBack);

  vehLabel = makeVehLabel(spec);

  const edgeMat = new THREE.LineBasicMaterial({ color: 0x0b1220, transparent: true, opacity: 0.4 });
  v.placements.forEach((p, idx) => {
    const geo = new THREE.BoxGeometry(p.w, p.h, p.l);
    const mesh = new THREE.Mesh(geo, new THREE.MeshLambertMaterial({ color: colorFor(p.pkg.name) }));
    mesh.position.set(p.x + p.w / 2, p.y + p.h / 2, p.z + p.l / 2);
    mesh.userData = {
      name: p.pkg.name,
      dims: `${p.w}×${p.l}×${p.h} cm`,
      rule: p.pkg.rule,
      dir: p.pkg.loadDir === "ENINE" ? "Enine" : p.pkg.loadDir === "BOYUNA" ? "Boyuna" : "Serbest",
      level: p.stackLevel + 1,
      order: idx + 1,
      vol: (p.w * p.h * p.l) / 1e6,
    };
    mesh.add(new THREE.LineSegments(new THREE.EdgesGeometry(geo), edgeMat));
    boxGroup.add(mesh);
    meshes.push(mesh);
  });

  camCtl.center = { x: spec.w / 2, y: spec.h * 0.4, z: spec.l / 2 };
  camCtl.radius = Math.hypot(spec.w, spec.h, spec.l) * 1.15;
  camCtl.theta = Math.PI / 4;
  camCtl.phi = 1.05;
  updateCamera();

  const sl = $("loadSlider");
  sl.max = meshes.length;
  sl.value = meshes.length;
  applyVisibility(meshes.length);
  stopPlay();
}

// ─── Kamera kontrolleri ───
function updateCamera() {
  const { theta, phi, radius, center } = camCtl;
  camera.position.set(
    center.x + radius * Math.sin(phi) * Math.cos(theta),
    center.y + radius * Math.cos(phi),
    center.z + radius * Math.sin(phi) * Math.sin(theta)
  );
  camera.lookAt(center.x, center.y, center.z);
}

function bindControls(el) {
  let drag = false, lx = 0, ly = 0;
  el.addEventListener("mousedown", (e) => { drag = true; lx = e.clientX; ly = e.clientY; });
  window.addEventListener("mouseup", () => (drag = false));
  window.addEventListener("mousemove", (e) => {
    if (drag) {
      camCtl.theta -= (e.clientX - lx) * 0.006;
      camCtl.phi = Math.min(1.52, Math.max(0.08, camCtl.phi - (e.clientY - ly) * 0.006));
      lx = e.clientX; ly = e.clientY;
      updateCamera();
    }
    hover(e, el);
  });
  el.addEventListener("wheel", (e) => {
    e.preventDefault();
    camCtl.radius = Math.min(100000, Math.max(100, camCtl.radius * (1 + Math.sign(e.deltaY) * 0.1)));
    updateCamera();
  }, { passive: false });
}

function hover(e, el) {
  if (!camera || !state.result || !meshes.length) return;
  const rect = el.getBoundingClientRect();
  const tip = $("tooltip");
  if (e.clientX < rect.left || e.clientX > rect.right || e.clientY < rect.top || e.clientY > rect.bottom) {
    tip.style.display = "none";
    return;
  }
  mouseV.x = ((e.clientX - rect.left) / rect.width) * 2 - 1;
  mouseV.y = -((e.clientY - rect.top) / rect.height) * 2 + 1;
  raycaster.setFromCamera(mouseV, camera);
  const hits = raycaster.intersectObjects(meshes.filter((m) => m.visible), false);
  if (hits.length) {
    const d = hits[0].object.userData;
    tip.innerHTML = `<b>${d.name}</b> • ${d.order}. yüklenen • ${d.vol.toFixed(3)} m³<br>${d.dims}<br>Kural: ${d.rule} • Yön: ${d.dir} • Katman ${d.level}`;
    tip.style.display = "block";
    tip.style.left = Math.min(rect.width - 230, e.clientX - rect.left + 14) + "px";
    tip.style.top = e.clientY - rect.top + 14 + "px";
  } else tip.style.display = "none";
}

// ─── Animasyon + hız kontrolü (0.25x – 2x) ───
let playTimer = null;
function stopPlay() {
  if (playTimer) { clearInterval(playTimer); playTimer = null; $("playBtn").textContent = "⏵ Yükleme animasyonu"; }
}

function applyVisibility(n) {
  meshes.forEach((m, i) => (m.visible = i < n));
  $("sliderLabel").textContent = `${n} / ${meshes.length} paket`;
  let vol = 0;
  for (let i = 0; i < n && i < meshes.length; i++) vol += meshes[i].userData.vol;
  $("volStat").textContent = vol.toFixed(1);
  const cap = capVolume();
  $("fillStat").textContent = cap ? ((vol / cap) * 100).toFixed(1) : "0";
  updateVehLabel(n);
}

function startPlay(keepPos) {
  if (!meshes.length) return;
  stopPlay();
  const sl = $("loadSlider");
  if (!keepPos) { sl.value = 0; applyVisibility(0); }
  const speed = +$("speedSelect").value || 1;
  const step = Math.max(1, Math.ceil((meshes.length / 150) * speed));
  const interval = Math.max(12, 30 / speed);
  $("playBtn").textContent = "⏸ Durdur";
  playTimer = setInterval(() => {
    const nv = Math.min(meshes.length, +sl.value + step);
    sl.value = nv;
    applyVisibility(nv);
    if (nv >= meshes.length) stopPlay();
  }, interval);
}

 $("playBtn").addEventListener("click", () => (playTimer ? stopPlay() : startPlay(false)));
 $("speedSelect").addEventListener("change", () => {
  if (playTimer) startPlay(true);
});
 $("loadSlider").addEventListener("input", (e) => { stopPlay(); applyVisibility(+e.target.value); });

// ═══════════ Olay bağlama + başlangıç ═══════════
VEHICLE_TYPES.forEach((v, i) => {
  const o = document.createElement("option");
  o.value = i;
  o.textContent = v.w ? `${v.name} (${v.w}×${v.l}×${v.h} cm)` : `${v.name} — ölçüleri sen belirle`;
  $("vehicleSelect").appendChild(o);
});
 $("vehicleSelect").addEventListener("change", () => {
  const id = VEHICLE_TYPES[$("vehicleSelect").value].id;
  const custom = id === "parsiyel" || id === "custom";
  $("customDims").classList.toggle("hidden", !custom);
  $("customNameWrap").classList.toggle("hidden", id !== "custom");
});

 $("runBtn").addEventListener("click", runSim);
 $("Btn").addEventListener("click", () => { state.packages = .map((d) => ({ ...d })); resetForm(); renderTable(); });

 $("fileInput").addEventListener("change", (e) => {
  if (e.target.files[0]) handleFile(e.target.files[0]);
  e.target.value = "";
});
const dz = $("dropzone");
dz.addEventListener("click", () => $("fileInput").click());
dz.addEventListener("dragover", (e) => { e.preventDefault(); dz.classList.add("drag"); });
dz.addEventListener("dragleave", () => dz.classList.remove("drag"));
dz.addEventListener("drop", (e) => {
  e.preventDefault();
  dz.classList.remove("drag");
  if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
});

if (typeof THREE === "undefined") {
  document.body.insertAdjacentHTML("afterbegin",
    '<div style="background:#7f1d1d;color:#fff;padding:10px;text-align:center;font-size:14px">⚠ 3D kütüphanesi yüklenemedi — internet bağlantınızı / proxy ayarlarınızı kontrol edin.</div>');
}
