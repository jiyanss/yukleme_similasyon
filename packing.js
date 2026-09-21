// ═══════════════════════════════════════════════════════
//  Yükleme Simülasyon Motoru — 3D Bin Packing
//  Yöntem: First-Fit Decreasing + Extreme Points
//  Ölçüler: cm | simulate(paketler, aracOlcu, aracAdedi, fireYuzde)
//  Yön: "" serbest | "ENINE" boy→araç eni | "BOYUNA" boy→araç boyu
//  Öncelik: küçük sayı önce yüklenir
//  Fire: kullanılabilir hacim = brüt × (1 - fire/100), aşılırsa paket o araca yüklenmez
//  İstif: boşsa otomatik = floor(araç yüksekliği / paketin o duruştaki yüksekliği)
// ═══════════════════════════════════════════════════════
var PackingEngine = (function () {
  "use strict";

  var RULE_ORDER = { YERDE: 0, STANDART: 1, DIK: 1, KIRILGAN: 2 };

  function prepareItems(packages) {
    var items = [];
    packages.forEach(function (p) {
      var qty = Math.max(1, parseInt(p.Adet, 10) || 1);
      var dir = String(p.Yukleme || "").trim().toUpperCase();
      if (dir !== "ENINE" && dir !== "BOYUNA") dir = "";
      var stackRaw = String(p.Istif === undefined || p.Istif === null ? "" : p.Istif).trim();
      var stackInput = stackRaw === "" ? null : parseInt(stackRaw, 10);
      if (stackInput !== null && (isNaN(stackInput) || stackInput < 1)) stackInput = null;
      var prio = parseInt(p.Oncelik, 10);
      if (isNaN(prio)) prio = 99;
      for (var i = 0; i < qty; i++) {
        items.push({
          id: p.UrunKod + "-" + (i + 1),
          name: p.UrunKod,
          w: +p.En, l: +p.Boy, h: +p.Yukseklik,
          rule: p.Kural || "STANDART",
          maxStackInput: stackInput,
          loadDir: dir,
          oncelik: prio,
        });
      }
    });
    items.sort(function (a, b) {
      if (a.oncelik !== b.oncelik) return a.oncelik - b.oncelik;
      var ro = RULE_ORDER[a.rule] - RULE_ORDER[b.rule];
      if (ro !== 0) return ro;
      return b.w * b.l * b.h - a.w * a.l * a.h;
    });
    return items;
  }

  // w = araç genişliği (x), l = araç uzunluğu (z), h = dikey (y)
  function getRotations(item) {
    if (item.rule === "DIK") return [{ w: item.w, l: item.l, h: item.h }];

    var d = [item.w, item.l, item.h];
    var perms;

    if (item.loadDir === "BOYUNA") {
      perms = [
        [d[0], d[1], d[2]],
        [d[2], d[1], d[0]],
      ];
    } else if (item.loadDir === "ENINE") {
      perms = [
        [d[1], d[0], d[2]],
        [d[1], d[2], d[0]],
      ];
    } else {
      perms = [
        [d[0], d[1], d[2]], [d[0], d[2], d[1]],
        [d[1], d[0], d[2]], [d[1], d[2], d[0]],
        [d[2], d[0], d[1]], [d[2], d[1], d[0]],
      ];
    }

    var seen = {}, out = [];
    perms.forEach(function (pm) {
      var key = pm.join("|");
      if (!seen[key]) { seen[key] = 1; out.push({ w: pm[0], l: pm[1], h: pm[2] }); }
    });

    // Akıllı sıralama: aracın enini en çok dolduran önce → yan şerit minik kalır,
    // açılan şeriğe sonraki parçalar dikine yerleşebilir
    out.sort(function (a, b) {
      if (b.w !== a.w) return b.w - a.w;
      return (b.w * b.l) - (a.w * a.l);
    });
    return out;
  }

  function collides(placements, box) {
    for (var i = 0; i < placements.length; i++) {
      var p = placements[i];
      if (box.x < p.x + p.w && box.x + box.w > p.x &&
          box.y < p.y + p.h && box.y + box.h > p.y &&
          box.z < p.z + p.l && box.z + box.l > p.z) return true;
    }
    return false;
  }

  // Alt yüzeyin en az %75'i desteklenmeli
  function supported(placements, box) {
    if (box.y < 0.01) return true;
    var support = 0, total = box.w * box.l;
    for (var i = 0; i < placements.length; i++) {
      var p = placements[i];
      if (Math.abs(p.y + p.h - box.y) >= 0.01) continue;
      var ox = Math.min(box.x + box.w, p.x + p.w) - Math.max(box.x, p.x);
      var oz = Math.min(box.z + box.l, p.z + p.l) - Math.max(box.z, p.z);
      if (ox > 0 && oz > 0) support += ox * oz;
    }
    return support / total >= 0.75;
  }

  // Kırılğanın üstüne yasak + katman limiti
  function stackingOk(placements, box, item, maxStack) {
    var level = 0;
    for (var i = 0; i < placements.length; i++) {
      var p = placements[i];
      if (Math.abs(p.y + p.h - box.y) >= 0.01) continue;
      var ox = Math.min(box.x + box.w, p.x + p.w) - Math.max(box.x, p.x);
      var oz = Math.min(box.z + box.l, p.z + p.l) - Math.max(box.z, p.z);
      if (ox > 0 && oz > 0) {
        if (p.pkg.rule === "KIRILGAN") return false;
        level = Math.max(level, p.stackLevel + 1);
      }
    }
    if (level >= maxStack) return false;
    box._stackLevel = level;
    return true;
  }

  function tryPlace(vehicle, item, spec, usableVol) {
    var rotations = getRotations(item);
    var itemVol = (item.w * item.l * item.h) / 1e6;

    // FIRE kontrolü: kullanılabilir hacmi aşacak paket bu araca girmez
    if (vehicle.usedVol + itemVol > usableVol + 1e-6) return false;

    var points = vehicle.points.slice().sort(function (a, b) {
      return a.z - b.z || a.x - b.x || a.y - b.y;
    });
    for (var pi = 0; pi < points.length; pi++) {
      var pt = points[pi];
      for (var ri = 0; ri < rotations.length; ri++) {
        var r = rotations[ri];
        var box = { x: pt.x, y: pt.y, z: pt.z, w: r.w, l: r.l, h: r.h };
        if (box.x + box.w > spec.w + 0.01) continue;
        if (box.y + box.h > spec.h + 0.01) continue;
        if (box.z + box.l > spec.l + 0.01) continue;

        // Otomatik istif: bu duruşta fiziksel olarak kaç katman sığar?
        var autoStack = Math.max(1, Math.floor(spec.h / box.h));
        var effMaxStack = item.maxStackInput !== null
          ? Math.min(item.maxStackInput, autoStack)
          : autoStack;

        if (collides(vehicle.placements, box)) continue;
        if (!supported(vehicle.placements, box)) continue;
        if (!stackingOk(vehicle.placements, box, item, effMaxStack)) continue;

        vehicle.placements.push({
          x: box.x, y: box.y, z: box.z, w: box.w, l: box.l, h: box.h,
          stackLevel: box._stackLevel, pkg: item,
        });
        vehicle.usedVol += itemVol;
        vehicle.points.push(
          { x: box.x + box.w, y: box.y, z: box.z },
          { x: box.x, y: box.y + box.h, z: box.z },
          { x: box.x, y: box.y, z: box.z + box.l }
        );
        return true;
      }
    }
    return false;
  }

  function simulate(packages, vehicleSpec, vehicleCount, firePct) {
    firePct = Math.min(50, Math.max(0, +firePct || 0));
    var items = prepareItems(packages);
    var grossVol = (vehicleSpec.w * vehicleSpec.h * vehicleSpec.l) / 1e6;
    var usableVol = grossVol * (1 - firePct / 100);

    var vehicles = [];
    for (var i = 0; i < vehicleCount; i++)
      vehicles.push({ placements: [], points: [{ x: 0, y: 0, z: 0 }], usedVol: 0 });
    var unplaced = [];

    items.forEach(function (item) {
      for (var vi = 0; vi < vehicles.length; vi++) {
        if (tryPlace(vehicles[vi], item, vehicleSpec, usableVol)) return;
      }
      unplaced.push(item);
    });

    var stats = vehicles.map(function (v) {
      return {
        count: v.placements.length,
        usedVolume: +v.usedVol.toFixed(2),
        fillRate: usableVol > 0 ? +((v.usedVol / usableVol) * 100).toFixed(1) : 0,
      };
    });
    return {
      vehicles: vehicles, unplaced: unplaced, stats: stats, spec: vehicleSpec,
      firePct: firePct, grossVolume: +grossVol.toFixed(2), usableVolume: +usableVol.toFixed(2),
    };
  }

  return { simulate: simulate };
})();
