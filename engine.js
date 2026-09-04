/* STEN_MASTER v3.8 — pure calculation engine, synchronized with Google Sheets v3.7.
 * This file intentionally contains no DOM code so it can be regression-tested in Node.js.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.StenEngine = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const FINISH_NONE = 'Без отделки';

  function num(v, fallback = 0) {
    const n = typeof v === 'number' ? v : Number(String(v ?? '').replace(/\s/g, '').replace(',', '.'));
    return Number.isFinite(n) ? n : fallback;
  }

  function ceilPositive(v) {
    return v > 0 ? Math.ceil(v - 1e-12) : 0;
  }

  function getConstruction(model, category, name) {
    const list = model?.constructions?.[category] || [];
    return list.find(x => x.name === name) || list[0] || null;
  }

  function isFrame(name) {
    return String(name || '').includes('Каркасная');
  }

  function getFinishSubtype(materialName) {
    return isFrame(materialName) ? 'ГВЛВ' : 'Кирпич/блок';
  }

  function getFinishRate(model, finishType, materialName) {
    if (!finishType || finishType === FINISH_NONE) return 0;
    const subtype = getFinishSubtype(materialName);
    const item = (model.finish || []).find(x => x.type === finishType && x.subtype === subtype);
    return item ? num(item.sheetRate ?? item.rate ?? item.modelRate) : 0;
  }

  function getEffectiveThickness(mat, finishType) {
    if (!mat) return 0;
    return finishType === FINISH_NONE ? num(mat.thick) : num(mat.thickFin);
  }

  function getDistance(model, objectCity, productionCity) {
    const d = model?.delivery;
    if (!d) return 0;
    const idx = (d.productionCities || d.cities || []).indexOf(productionCity);
    const row = d.distances?.[objectCity];
    if (!row || idx < 0) return 0;
    return num(row[idx]);
  }

  function calcWall(category, mat, finishType, inputs, model) {
    if (!mat) return emptyCalc(category);
    const isExt = category === 'ext';
    const perim = num(isExt ? inputs.perimExt : inputs.perimInter);
    const floorH = num(inputs.floorH);
    const floors = num(inputs.floors);
    const delivery = model.delivery || {};
    const rateKm = num(delivery.rateKm);
    const truckVol = num(delivery.truckVol);
    const truckKg = num(delivery.truckKg);
    const thick = num(mat.thick);
    const thickFin = getEffectiveThickness(mat, finishType);
    const density = num(mat.density);
    const openingPct = isExt ? num(model.constants?.extOpeningPct, 0.15) : num(model.constants?.interOpeningPct, 0.07);
    const wasteFactor = num(model.constants?.unitWasteFactor, 1.07);

    // Exact calculation basis of the current Google Sheets model.
    const footprintArea = perim * thickFin / 1000;
    const volumeFloor = perim * thick / 1000 * floorH;
    const densityT = density;
    const massFloor = volumeFloor * densityT / 1000;
    const unitDen = num(mat.L) * num(mat.W) * num(mat.H);
    const unitsFloor = unitDen > 0 ? (volumeFloor / unitDen) * wasteFactor : 0;
    const unitPrice = num(mat.sheetPrice ?? mat.rsc);
    const workUnit = num(mat.workCostSheet ?? mat.workCostModel);
    const materialUnit = num(mat.materialCostSheet ?? mat.materialCostModel, unitPrice - workUnit);
    const costFloor = volumeFloor * unitPrice;
    const costObj = costFloor * floors;
    const distance = getDistance(model, inputs.city, mat.city);
    const volObj = volumeFloor * floors;
    const massObjKg = massFloor * floors * 1000;
    const byWeight = truckKg > 0 ? ceilPositive(massObjKg / truckKg) : 0;
    const byVolume = truckVol > 0 ? ceilPositive(volObj / truckVol) : 0;
    const trucks = volumeFloor <= 0 ? 0 : Math.max(byWeight, byVolume, 1);
    const deliveryCost = trucks * distance * rateKm;
    const finishAreaFloor = perim * floorH * (1 - openingPct);
    const finishRate = getFinishRate(model, finishType, mat.name);
    // Exterior is only the inner face; finish base rate is built for two faces.
    const finishCost = isExt
      ? finishRate * finishAreaFloor / 2 * floors
      : finishRate * finishAreaFloor * floors;
    const total = costObj + finishCost;

    return {
      category,
      material: mat.name,
      city: mat.city,
      finishType,
      thick,
      thickFin,
      footprintArea,
      density,
      volumeFloor,
      unitsFloor,
      massFloor,
      unitPrice,
      workUnit,
      materialUnit,
      costFloor,
      costObj,
      distance,
      trucks,
      deliveryCost,
      finishAreaFloor,
      finishRate,
      finishCost,
      total,
      insulation: isExt ? mat.reference : null,
      noise: !isExt ? mat.reference : null,
      components: mat.components || []
    };
  }

  function calcPartition(mat, finishType, inputs, model) {
    if (!mat) return emptyCalc('part');
    const perim = num(inputs.perimPart);
    const floorH = num(inputs.floorH);
    const floors = num(inputs.floors);
    const delivery = model.delivery || {};
    const rateKm = num(delivery.rateKm);
    const truckVol = num(delivery.truckVol);
    const truckKg = num(delivery.truckKg);
    const thick = num(mat.thick);
    const thickFin = getEffectiveThickness(mat, finishType);
    const density = num(mat.density);
    const openingPct = num(model.constants?.partOpeningPct, 0.10);
    const wasteFactor = num(model.constants?.unitWasteFactor, 1.07);

    const footprintArea = perim * thickFin / 1000;
    const wallAreaFloor = perim * floorH * (1 - openingPct);
    const unitDen = num(mat.L) * num(mat.H);
    const unitsFloor = unitDen > 0 ? (wallAreaFloor / unitDen) * wasteFactor : 0;
    const massFloor = wallAreaFloor * thick / 1000 * density / 1000;
    const unitPrice = num(mat.sheetPrice ?? mat.rsc);
    const workUnit = num(mat.workCostSheet ?? mat.workCostModel);
    const materialUnit = num(mat.materialCostSheet ?? mat.materialCostModel, unitPrice - workUnit);
    const costFloor = wallAreaFloor * unitPrice;
    const costObj = costFloor * floors;

    let deliveryVolumeFloor;
    if (isFrame(mat.name)) {
      // Exact v3.5 formula. GVL uses 10 mm sheets, the other frame variant 12.5 mm.
      const boardThickness = String(mat.name).includes('ГВЛ') ? 0.01 : 0.0125;
      deliveryVolumeFloor = wallAreaFloor * 4 * (2.5 * 1.2 * boardThickness) + wallAreaFloor * 0.05;
    } else {
      deliveryVolumeFloor = wallAreaFloor * thick / 1000;
    }

    const distance = getDistance(model, inputs.city, mat.city);
    const massObjKg = massFloor * floors * 1000;
    const volObj = deliveryVolumeFloor * floors;
    const byWeight = truckKg > 0 ? ceilPositive(massObjKg / truckKg) : 0;
    const byVolume = truckVol > 0 ? ceilPositive(volObj / truckVol) : 0;
    const trucks = wallAreaFloor <= 0 ? 0 : Math.max(byWeight, byVolume, 1);
    const trucksFloor = wallAreaFloor <= 0 ? 0 : Math.max(
      truckKg > 0 ? ceilPositive(massFloor * 1000 / truckKg) : 0,
      truckVol > 0 ? ceilPositive(deliveryVolumeFloor / truckVol) : 0,
      1
    );
    const deliveryCost = trucks * distance * rateKm;
    const finishAreaFloor = wallAreaFloor;
    const finishRate = getFinishRate(model, finishType, mat.name);
    const finishCost = finishAreaFloor * floors * finishRate;
    const total = costObj + finishCost;

    return {
      category: 'part', material: mat.name, city: mat.city, finishType,
      thick, thickFin, footprintArea, density,
      wallAreaFloor, unitsFloor, massFloor, unitPrice, workUnit, materialUnit,
      costFloor, costObj, distance, trucks, trucksFloor, deliveryVolumeFloor,
      deliveryCost, finishAreaFloor, finishRate, finishCost, total,
      noise: mat.reference, components: mat.components || []
    };
  }

  function calcCategory(category, selection, inputs, model) {
    const mat = getConstruction(model, category, selection?.material);
    const finishType = selection?.finish || FINISH_NONE;
    return category === 'part'
      ? calcPartition(mat, finishType, inputs, model)
      : calcWall(category, mat, finishType, inputs, model);
  }

  function projectThickness(category, projectSelection, model) {
    const mat = getConstruction(model, category, projectSelection?.material);
    return getEffectiveThickness(mat, projectSelection?.finish || FINISH_NONE);
  }

  function calcScenario(scenario, project, inputs, model) {
    const ext = calcCategory('ext', scenario.ext, inputs, model);
    const inter = calcCategory('inter', scenario.inter, inputs, model);
    const part = calcCategory('part', scenario.part, inputs, model);
    const cats = { ext, inter, part };
    const perims = { ext: num(inputs.perimExt), inter: num(inputs.perimInter), part: num(inputs.perimPart) };

    let deltaFootprint = 0;
    const changedByCategory = {};
    for (const category of ['ext', 'inter', 'part']) {
      const selected = scenario[category] || {};
      const base = project[category] || {};
      const selectedFinish = selected.finish || FINISH_NONE;
      const baseFinish = base.finish || FINISH_NONE;
      const sameProjectSolution = Boolean(base.material) && selected.material === base.material && selectedFinish === baseFinish;

      // v3.7 rule: the selected solution always has a cost. Area delta is zero only
      // when both the material and the finish match the project solution.
      changedByCategory[category] = cats[category].total;
      if (!base.material || sameProjectSolution) continue;

      const scenarioThickness = cats[category].thickFin;
      const baseThickness = projectThickness(category, base, model);
      deltaFootprint += perims[category] * (scenarioThickness - baseThickness) / 1000;
    }

    const floors = num(inputs.floors);
    const deltaSaleFloor = -deltaFootprint;
    const deltaSaleObj = deltaSaleFloor * floors;
    const deltaSaleValue = deltaSaleObj * num(inputs.salePrice);
    const totalArea = num(inputs.totalArea);
    const saleArea = num(inputs.saleArea);

    const fullTotal = ext.total + inter.total + part.total;
    const changedTotal = fullTotal;

    return {
      categories: cats,
      fullTotal,
      deltaFootprint,
      deltaSaleFloor,
      deltaSaleObj,
      deltaSaleValue,
      changedByCategory,
      changedTotal,
      changedPerTotalArea: totalArea ? changedTotal / totalArea : 0,
      changedPerSaleArea: saleArea ? changedTotal / saleArea : 0,
      // This KPI is intentionally additional; it is not used by spreadsheet formulas.
      netEconomicEffect: deltaSaleValue - changedTotal
    };
  }

  function calcAll(state, model) {
    return (state.scenarios || []).map(sc => calcScenario(sc, state.project || {}, state.inputs || {}, model));
  }

  function emptyCalc(category) {
    return { category, total: 0, unitPrice: 0, workUnit: 0, materialUnit: 0, components: [] };
  }

  function almostEqual(a, b, tolerance = 1e-6) {
    return Math.abs(num(a) - num(b)) <= tolerance * Math.max(1, Math.abs(num(b)));
  }

  function runGoldenTests(model) {
    const d = model.defaults;
    const state = { inputs: {
      city:d.city, saleArea:d.saleArea, totalArea:d.totalArea, perimExt:d.perimExt,
      perimInter:d.perimInter, perimPart:d.perimPart, floorH:d.floorH,
      floors:d.floors, salePrice:d.salePrice
    }, project:d.project, scenarios:d.scenarios };
    const results = calcAll(state, model);
    const checks = [];
    const baseline = model.baseline || {};

    const mapping = {
      ext: ['thick','thickFin','footprintArea:footprint','density','unitsFloor:quantity','massFloor','unitPrice','costFloor','costObj','trucks','deliveryCost:delivery','finishAreaFloor','finishCost','total','workUnit','materialUnit'],
      inter: ['thick','thickFin','footprintArea:footprint','density','unitsFloor:quantity','massFloor','unitPrice','costFloor','costObj','trucks','deliveryCost:delivery','finishAreaFloor','finishRate','finishCost','total','workUnit','materialUnit'],
      part: ['thick','thickFin','footprintArea:footprint','density','wallAreaFloor:areaFloor','unitsFloor:quantity','massFloor','unitPrice','costFloor','costObj','trucks','deliveryVolumeFloor','trucksFloor','deliveryCost:delivery','finishAreaFloor','finishRate','finishCost','total','workUnit','materialUnit']
    };
    for (const cat of ['ext','inter','part']) {
      for (let i=0;i<3;i++) {
        const calc = results[i].categories[cat];
        const base = baseline[cat]?.[i] || {};
        for (const spec of mapping[cat]) {
          const [calcKey, baseKey] = spec.split(':');
          const bk = baseKey || calcKey;
          const expected = base[bk];
          if (expected === undefined) continue;
          checks.push({name:`${cat}[${i}].${calcKey}`, actual:calc[calcKey], expected, pass:almostEqual(calc[calcKey],expected,1e-7)});
        }
      }
    }
    for (let i=0;i<3;i++) {
      const calc=results[i], base=baseline.summary?.[i] || {};
      const pairs=[
        ['deltaFootprint','deltaFootprint'],['deltaSaleFloor','deltaSaleFloor'],['deltaSaleObj','deltaSaleObj'],['deltaSaleValue','deltaSaleValue'],
        ['changedTotal','changedTotal'],['changedPerTotalArea','changedPerTotalArea'],['changedPerSaleArea','changedPerSaleArea'],
      ];
      for (const [ck,bk] of pairs) if(base[bk]!==undefined) checks.push({name:`summary[${i}].${ck}`,actual:calc[ck],expected:base[bk],pass:almostEqual(calc[ck],base[bk],1e-7)});
      for (const [cat,bk] of [['ext','changedExt'],['inter','changedInter'],['part','changedPart']]) if(base[bk]!==undefined) checks.push({name:`summary[${i}].${bk}`,actual:calc.changedByCategory[cat],expected:base[bk],pass:almostEqual(calc.changedByCategory[cat],base[bk],1e-7)});
    }
    return {pass:checks.every(x=>x.pass), passed:checks.filter(x=>x.pass).length, total:checks.length, failed:checks.filter(x=>!x.pass)};
  }

  return {
    FINISH_NONE,
    num,
    getConstruction,
    getFinishSubtype,
    getFinishRate,
    getEffectiveThickness,
    getDistance,
    calcWall,
    calcPartition,
    calcCategory,
    calcScenario,
    calcAll,
    runGoldenTests
  };
});
