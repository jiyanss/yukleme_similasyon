// ═══════════════════════════════════════════════════════
//  Yükleme Simülasyon Motoru — 3D Bin Packing
//  Yöntem: First-Fit Decreasing + Extreme Points
//  Ölçüler: cm | Kullanım: PackingEngine.simulate(paketler, aracOlcu, aracAdedi)
// ═══════════════════════════════════════════════════════
var PackingEngine = (function () {
  "use strict";

  var RULE_ORDER = { YERDE: 0, STANDART: 1, DIK: 1, KIRILGAN: 2 };

  // Adetleri tekil paketlere aç, kural + hacme göre sırala
  function prepareItems(packages) {
    var items = [];
    packages.forEach(function (p) {
      var qty = Math.max(1, parseInt(p.Adet, 10) || 1);
      for (var i = 0; i < qty; i++) {
        items.push({
          id: p.UrunKod + "-" + (i + 1),
          name: p.UrunKod,
          w: +p.En, l: +p.Boy, h: +p.Yukseklik,
          rule: p.Kural || "STANDART",
          maxStack: parseInt(p.Istif, 10) || 5,
        });
      }
    });
    items.sort(function (a, b) {
      var ro = RULE_ORDER[a.rule] - RULE_ORDER[b.rule];
      if (ro !== 0) return ro;
      return b.w * b.l * b.h - a.w * a.l * a.h; // büyükten küçüğe
    });
    return items;
  }

  // İzinli dönüşler: DIK sadece dik, diğerleri 6 permütasyon (tekiller elenir)
  function getRotations(item) {
    if (item.rule === "DIK") return [{ w: item.w, l: item.l, h: item.h }];
    var d = [item.w, item.l, item.h];
    var perms = [
      [d[0], d[1], d[2]], [d[0], d[2], d[1]],
      [d[1], d[0], d[2]], [d[1], d[2], d[0]],
      [d[2], d[0], d[1]], [d[2], d[1], d[0]],
    ];
    var seen = {}, out = [];
    perms.forEach(function (pm) {
      var key = pm.join("|");
      if (!seen[key]) { seen[key] = 1; out.push({ w: pm[0], l: pm[1], h: pm[2] }); }
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

  // Alt yüzeyin en az %75'i desteklenmeli (zemin veya alttaki paket)
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

  // Kırılğanın üstüne yasak + istif katman limiti
  function stackingOk(placements, box, item) {
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
    if (level > item.maxStack) return false;
    box._stackLevel = level;
    return true;
  }

  function tryPlace(vehicle, item, spec) {
    var rotations = getRotations(item);
    // alttan ve önden başla → taban dolu ve doğal görünüm
    var points = vehicle.points.slice().sort(function (a, b) {
      return a.y - b.y || a.z - b.z || a.x - b.x;
    });
    for (var pi = 0; pi < points.length; pi++) {
      var pt = points[pi];
      for (var ri = 0; ri < rotations.length; ri++) {
        var r = rotations[ri];
        var box = { x: pt.x, y: pt.y, z: pt.z, w: r.w, l: r.l, h: r.h };
        if (box.x + box.w > spec.w + 0.01) continue;
        if (box.y + box.h > spec.h + 0.01) continue;
        if (box.z + box.l > spec.l + 0.01) continue;
        if (collides(vehicle.placements, box)) continue;
        if (!supported(vehicle.placements, box)) continue;
        if (!stackingOk(vehicle.placements, box, item)) continue;

        vehicle.placements.push({
          x: box.x, y: box.y, z: box.z, w: box.w, l: box.l, h: box.h,
          stackLevel: box._stackLevel, pkg: item,
        });
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

  function simulate(packages, vehicleSpec, vehicleCount) {
    var items = prepareItems(packages);
    var vehicles = [];
    for (var i = 0; i < vehicleCount; i++)
      vehicles.push({ placements: [], points: [{ x: 0, y: 0, z: 0 }] });
    var unplaced = [];

    items.forEach(function (item) {
      for (var vi = 0; vi < vehicles.length; vi++) {
        if (tryPlace(vehicles[vi], item, vehicleSpec)) return;
      }
      unplaced.push(item);
    });

    var totalVol = (vehicleSpec.w * vehicleSpec.h * vehicleSpec.l) / 1e6;
    var stats = vehicles.map(function (v) {
      var used = v.placements.reduce(function (s, p) {
        return s + (p.w * p.h * p.l) / 1e6;
      }, 0);
      return {
        count: v.placements.length,
        usedVolume: +used.toFixed(2),
        fillRate: +((used / totalVol) * 100).toFixed(1),
      };
    });
    return { vehicles: vehicles, unplaced: unplaced, stats: stats, spec: vehicleSpec };
  }

  return { simulate: simulate };
})();
