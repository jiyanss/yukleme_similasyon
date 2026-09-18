"use strict";

// ═══════════ Araç Tipleri — iç ölçüler cm, kendi filona göre buradan düzelt ═══════════
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
  { UrunKod: "Buzdolabi", En: 70, Boy: 75, Yukseklik: 180, Adet: 10, Kural: "DIK", Istif: "" },
  { UrunKod: "TV-55", En: 140, Boy: 25, Yukseklik: 85, Adet: 15, Kural: "KIRILGAN", Istif: "" },
  { UrunKod: "Koli-A", En: 40, Boy: 60, Yukseklik: 50, Adet: 100, Kural: "", Istif: 4 },
  { UrunKod: "Koli-B", En: 50, Boy: 50, Yukseklik: 40, Adet: 80, Kural: "", Istif: 5 },
  { UrunKod: "Sandik", En: 80, Boy: 80, Yukseklik: 60, Adet: 20, Kural: "YERDE", Istif: "" },
  { UrunKod: "CamasirMak", En: 60, Boy: 60, Yukseklik: 85, Adet: 12, Kural: "DIK", Istif: "" },
];

const state = { packages: [], result: null, activeVehicle: 0 };
const $ = (id) => document.getElementById(id);

// ═══════════ Excel şablonu indir ═══════════
 $("tplBtn").addEventListener("click", () => {
  if (typeof XLSX === "undefined") { alert("Excel kütüphanesi yüklenemedi (internet engeli olabilir)."); return; }
  const wb = XLSX.utils.book_new();

  const data = [
    ["UrunKod", "En", "Boy", "Yukseklik", "Adet", "Kural", "Istif"],
    ["TV-55", 140, 25, 85, 4, "KIRILGAN", ""],
    ["Buzdolabi", 70, 75, 180, 2, "DIK", ""],
    ["Koli-A", 40, 60, 50, 100, "", 4],
    ["Sandik", 80, 80, 60, 10, "YERDE", ""],
  ];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [{ wch: 14 }, { wch: 8 }, { wch: 8 }, { wch: 11 }, { wch: 7 }, { wch: 12 }, { wch: 8 }];
  XLSX.utils.book_append_sheet(wb, ws, "YuklemeListesi");

  const help = [
    ["ALAN", "AÇIKLAMA"],
    ["UrunKod", "Paket kodu (zorunlu)"],
    ["En / Boy / Yukseklik", "Paket ölçüleri cm (zorunlu)"],
    ["Adet", "Kaç adet yüklenecek (boşsa 1)"],
    ["Kural", "STANDART / KIRILGAN / DIK / YERDE (boşsa STANDART)"],
    ["Istif", "Üstüne maksimum kaç paket konabilir (boşsa 5)"],
    [],
    ["KURAL", "ANLAMI"],
    ["STANDART", "Serbest döner, istiflenebilir"],
    ["KIRILGAN", "Üstüne hiçbir paket konamaz"],
    ["DIK", "Döndürülemez, sadece dik durur"],
    ["YERDE", "En alt katmanda durur"],
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
};

function pickCol(row, candidates) {
  const keys = Object.keys(row);
  for (const c of candidates) {
    const k = keys.find((key) => key.toLowerCase().replace(/[\s_.]/g, "") === c);
    if (k !== undefined && String(row[k]).trim() !== "") return String(row[k]).trim();
  }
  return "";
}

function num(v) {
  const n = parseFloat(String(v).replace(",", ".")); // TR ondalık virgülü
  return isNaN(n) ? 0 : n;
}

function parseRows(rows) {
  const good = [], errors = [];
  rows.forEach((row, idx) => {
    const code = pickCol(row, COLS.UrunKod);
    const w = num(pickCol(row, COLS.En));
    const l = num(pickCol(row, COLS.Boy));
    const h = num(pickCol(row, COLS.Yukseklik));
    const qty = Math.round(num(pickCol(row, COLS.Adet)) || 1);
    let rule = pickCol(row, COLS.Kural).toUpperCase();
    if (!["STANDART", "KIRILGAN", "DIK", "YERDE"].includes(rule)) rule = "STANDART";
    const stack = pickCol(row, COLS.Istif);
    if (!code || !w || !l || !h) {
      errors.push(`Satır ${idx + 2}: eksik/hatalı veri (Ürün kodu + En/Boy/Yükseklik gerekli)`);
      return;
    }
    good.push({ UrunKod: code, En: w, Boy: l, Yukseklik: h, Adet: qty, Kural: rule, Istif: stack });
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
        showErrors([...errors, "Geçerli satır yok. İlk satır başlık olmalı: UrunKod, En, Boy, Yukseklik, Adet, Kural, Istif"]);
        return;
      }
      state.packages = good;
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
      `<td>${p.Adet}</td><td>${p.Kural}</td><td>${p.Istif || "-"}</td>` +
      `<td><button class="del" data-i="${i}" title="Sil">✕</button></td>`;
    tb.appendChild(tr);
  });
  $("pkgCount").textContent = state.packages.length
    ? `${state.packages.length} kalem • toplam ${total} paket` : "";
  $("tableWrap").classList.toggle("hidden", !state.packages.length);
}

 $("pkgBody").addEventListener("click", (e) => {
  if (e.target.classList.contains("del")) {
    state.packages.splice(+e.target.dataset.i, 1);
    renderTable();
  }
});

// ═══════════ Simülasyon ═══════════
function runSim() {
  if (!state.packages.length) { alert("Önce yükleme listesi ekleyin (Excel veya örnek veri)."); return; }
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
  const capVol = ((res.spec.w * res.spec.h * res.spec.l) / 1e6).toFixed(1);
  $("vehInfo").textContent =
    `${res.spec.name} — iç ölçü ${res.spec.w}×${res.spec.l}×${res.spec.h} cm — ` +
    `dolum: önden (yeşil) arkaya, kapıya (turuncu) doğru`;
  $("statsBar").innerHTML =
    `<span>🚛 Araç ${state.activeVehicle + 1}: <b>${s.count}</b> paket</span>` +
    `<span>Doluluk: <b>%${s.fillRate}</b></span>` +
    `<span>Hacim: <b>${s.usedVolume}</b> / ${capVol} m³</span>` +
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
let renderer = null, scene, camera, boxGroup, sceneGroup;
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

// Yazı etiketi (sprite) — ÖN / KAPI işaretleri için
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

function buildScene() {
  if (!renderer && !initThree()) return;
  const res = state.result;
  if (!res) return;
  const spec = res.spec;
  const v = res.vehicles[state.activeVehicle];

  disposeGroup(sceneGroup);
  disposeGroup(boxGroup);
  meshes = [];

  // zemin + grid
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

  // araç konturu
  const bgeo = new THREE.BoxGeometry(spec.w, spec.h, spec.l);
  const outline = new THREE.LineSegments(
    new THREE.EdgesGeometry(bgeo),
    new THREE.LineBasicMaterial({ color: 0x3b82f6, transparent: true, opacity: 0.6 })
  );
  outline.position.set(spec.w / 2, spec.h / 2, spec.l / 2);
  sceneGroup.add(outline);
  bgeo.dispose();

  // ── Dolum yönü işaretleri ──
  // ÖN: yeşil şerit (dolum buradan başlar, z=0 ucu)
  const startStrip = new THREE.Mesh(
    new THREE.PlaneGeometry(spec.w, Math.min(8, spec.l * 0.05)),
    new THREE.MeshBasicMaterial({ color: 0x22c55e, transparent: true, opacity: 0.55, side: THREE.DoubleSide })
  );
  startStrip.rotation.x = -Math.PI / 2;
  startStrip.position.set(spec.w / 2, 0.8, Math.min(4, spec.l * 0.03));
  sceneGroup.add(startStrip);

  // KAPI: turuncu çerçeve (z = boy ucu)
  const doorGeo = new THREE.PlaneGeometry(spec.w, spec.h);
  const door = new THREE.LineSegments(
    new THREE.EdgesGeometry(doorGeo),
    new THREE.LineBasicMaterial({ color: 0xf97316 })
  );
  door.position.set(spec.w / 2, spec.h / 2, spec.l);
  sceneGroup.add(door);
  doorGeo.dispose();

  // Yön oku + etiketler
  const arrowLen = Math.min(spec.l * 0.35, 250);
  const arrow = new THREE.ArrowHelper(
    new THREE.Vector3(0, 0, 1),
    new THREE.Vector3(spec.w / 2, 25, 15),
    arrowLen, 0x22c55e, arrowLen * 0.25, arrowLen * 0.12
  );
  sceneGroup.add(arrow);

  const lblW = spec.w * 0.85, lblH = spec.w * 0.22;
  const sFront = textSprite("ÖN • dolum başlar", "#4ade80", lblW, lblH);
  sFront.position.set(spec.w / 2, spec.h * 0.9, -lblW * 0.6);
  sceneGroup.add(sFront);
  const sBack = textSprite("ARKA • KAPI", "#fb923c", lblW, lblH);
  sBack.position.set(spec.w / 2, spec.h * 0.9, spec.l + lblW * 0.6);
  sceneGroup.add(sBack);

  // ── Paketler (yerleşim sırası = dolum sırası: önden arkaya) ──
  const edgeMat = new THREE.LineBasicMaterial({ color: 0x0b1220, transparent: true, opacity: 0.4 });
  v.placements.forEach((p) => {
    const geo = new THREE.BoxGeometry(p.w, p.h, p.l);
    const mesh = new THREE.Mesh(geo, new THREE.MeshLambertMaterial({ color: colorFor(p.pkg.name) }));
    mesh.position.set(p.x + p.w / 2, p.y + p.h / 2, p.z + p.l / 2);
    mesh.userData = {
      name: p.pkg.name,
      dims: `${p.w}×${p.l}×${p.h} cm`,
      rule: p.pkg.rule,
      level: p.stackLevel + 1,
      order: v.placements.indexOf(p) + 1,
    };
    mesh.add(new THREE.LineSegments(new THREE.EdgesGeometry(geo), edgeMat));
    boxGroup.add(mesh);
    meshes.push(mesh);
  });

  // kamerayı araca göre konumlandır
  camCtl.center = { x: spec.w / 2, y: spec.h * 0.4, z: spec.l / 2 };
  camCtl.radius = Math.hypot(spec.w, spec.h, spec.l) * 1.15;
  camCtl.theta = Math.PI / 4;
  camCtl.phi = 1.05;
  updateCamera();

  const sl = $("loadSlider");
  sl.max = meshes.length;
  sl.value = meshes.length;
  $("sliderLabel").textContent = meshes.length ? `${meshes.length} / ${meshes.length} paket` : "paket yok";
  stopPlay();
}

// ─── Kamera kontrolleri (döndür + zoom + hover bilgi) ───
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
    tip.innerHTML = `<b>${d.name}</b> • ${d.order}. yüklenen<br>${d.dims}<br>Kural: ${d.rule} • Katman ${d.level}`;
    tip.style.display = "block";
    tip.style.left = Math.min(rect.width - 190, e.clientX - rect.left + 14) + "px";
    tip.style.top = e.clientY - rect.top + 14 + "px";
  } else tip.style.display = "none";
}

// ─── Animasyon ───
let playTimer = null;
function stopPlay() {
  if (playTimer) { clearInterval(playTimer); playTimer = null; $("playBtn").textContent = "⏵ Yükleme animasyonu"; }
}
function applyVisibility(n) {
  meshes.forEach((m, i) => (m.visible = i < n));
  $("sliderLabel").textContent = `${n} / ${meshes.length} paket`;
}

 $("playBtn").addEventListener("click", () => {
  if (playTimer) { stopPlay(); return; }
  if (!meshes.length) return;
  const sl = $("loadSlider");
  sl.value = 0;
  applyVisibility(0);
  const step = Math.max(1, Math.ceil(meshes.length / 150));
  $("playBtn").textContent = "⏸ Durdur";
  playTimer = setInterval(() => {
    const nv = Math.min(meshes.length, +sl.value + step);
    sl.value = nv;
    applyVisibility(nv);
    if (nv >= meshes.length) stopPlay();
  }, 30);
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
 $("demoBtn").addEventListener("click", () => { state.packages = DEMO.map((d) => ({ ...d })); renderTable(); });

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
