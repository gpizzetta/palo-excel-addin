/* global Office, Excel, OfficeRuntime */
(function commandsBootstrap() {
  var PALO_PLUGIN_VERSION = "1.0.3.24";
  var PICKER_STORAGE_KEY = "palo_ename_picker_v1";
  var DATAC_WRITE_STORAGE_KEY = "palo_datac_write_v1";
  var NUMERIC_CELL_TEXT_MESSAGE = "Cette mesure est numérique mais la cellule est du texte : la consolidation restera à 0. Resaisir en nombre.";
  var DATAC_WRITE_MAX_CELLS = 200;

  function complete(event) {
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  }

  /**
   * Excel Web / shared runtime : window.alert n'est pas supporte.
   * Met a jour le volet (#status-log) si present, sinon mini-dialogue Office.
   */
  function paloUserNotify(message, kind, title) {
    var text = String(message || "");
    var state = kind === "error" ? "error" : kind === "ok" ? "ok" : "neutral";
    var heading = title || (state === "error" ? "Palo — erreur" : "Palo");

    var statusLog = document.getElementById("status-log");
    if (statusLog) {
      statusLog.textContent = text;
      var banner = document.getElementById("status-banner");
      if (banner) {
        banner.classList.remove("status-banner--ok", "status-banner--error");
        if (state === "ok") {
          banner.classList.add("status-banner--ok");
        } else if (state === "error") {
          banner.classList.add("status-banner--error");
        }
      }
      return Promise.resolve();
    }

    return new Promise(function (resolve) {
      if (!Office.context || !Office.context.ui || typeof Office.context.ui.displayDialogAsync !== "function") {
        console.warn("[Palo]", heading, text);
        resolve();
        return;
      }
      var notifyHref;
      try {
        var notifyUrl = new URL("palo-action-notify.html", window.location.href);
        notifyUrl.searchParams.set("t", heading);
        notifyUrl.searchParams.set("m", text);
        notifyHref = notifyUrl.href;
      } catch (_urlErr) {
        notifyHref = "https://gpizzetta.github.io/palo-excel-addin/palo-action-notify.html?t="
          + encodeURIComponent(heading) + "&m=" + encodeURIComponent(text);
      }
      Office.context.ui.displayDialogAsync(
        notifyHref,
        { height: 35, width: 40, displayInIframe: true },
        function (asyncResult) {
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            resolve();
            return;
          }
          var dlg = asyncResult.value;
          var settled = false;
          function finish() {
            if (settled) {
              return;
            }
            settled = true;
            try {
              dlg.close();
            } catch (_close) {
              // ignore
            }
            resolve();
          }
          dlg.addEventHandler(Office.EventType.DialogMessageReceived, function () {
            finish();
          });
          dlg.addEventHandler(Office.EventType.DialogEventReceived, function () {
            finish();
          });
        }
      );
    });
  }

  async function openTaskpane(event) {
    try {
      if (Office.addin && typeof Office.addin.showAsTaskpane === "function") {
        await Office.addin.showAsTaskpane();
      }
    } catch (_e) {
      // Keep ribbon action resilient.
    }
    complete(event);
  }

  async function testConnection(event) {
    try {
      var manager = window.PaloOffice.createConnectionManager();
      var active = manager.getActiveConnectionName();
      var resultText = "Aucune connexion active.";
      if (active) {
        var result = await manager.testConnection(active);
        resultText = result.details;
      }
      await Excel.run(async function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange("A1");
        range.values = [["Palo test: " + resultText]];
        range.format.autofitColumns();
        await context.sync();
      });
    } catch (_error) {
      // Avoid blocking ribbon action if workbook is not available.
    }
    complete(event);
  }

  function escapeDoubleQuotesForFormula(s) {
    return String(s).replace(/"/g, '""');
  }

  async function insertPaloDataFormula(event) {
    try {
      if (!window.PaloOffice || typeof window.PaloOffice.createConnectionManager !== "function") {
        complete(event);
        return;
      }
      var manager = window.PaloOffice.createConnectionManager();
      var active = manager.getActiveConnectionName();
      if (!active) {
        await paloUserNotify("Aucune connexion active. Ouvrez le volet Palo et selectionnez une connexion dans la liste.", "error");
        complete(event);
        return;
      }
      var servdb = active + "/DWH";
      var safe = escapeDoubleQuotesForFormula(servdb);
      var formula = '=PALO.DATAC("' + safe + '";"Sales";"Actual";"2024";"Jan";"Total Products";"Local")';
      await Excel.run(async function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getActiveCell();
        range.formulasLocal = [[formula]];
        await context.sync();
      });
    } catch (_error) {
      // Ignore; command should not crash Office callbacks.
    }
    complete(event);
  }

  async function insertPaloSetDataFormula(event) {
    try {
      if (!window.PaloOffice || typeof window.PaloOffice.createConnectionManager !== "function") {
        complete(event);
        return;
      }
      var manager = window.PaloOffice.createConnectionManager();
      var active = manager.getActiveConnectionName();
      if (!active) {
        await paloUserNotify("Aucune connexion active. Ouvrez le volet Palo et selectionnez une connexion dans la liste.", "error");
        complete(event);
        return;
      }
      var servdb = active + "/DWH";
      var safe = escapeDoubleQuotesForFormula(servdb);
      var formula = '=PALO_SETDATA(100;0;"' + safe + '";"Sales";"Actual";"2024";"Jan";"Total Products";"Local")';
      await Excel.run(async function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getActiveCell();
        range.formulasLocal = [[formula]];
        await context.sync();
      });
    } catch (_error) {
      // Ignore; command should not crash Office callbacks.
    }
    complete(event);
  }

  function storageSetJson(key, obj) {
    var json = JSON.stringify(obj);
    try {
      window.localStorage.setItem(key, json);
    } catch (_e) {
      // Quota / mode prive : le parent enverra aussi les donnees via messageChild.
    }
    var storage = typeof OfficeRuntime !== "undefined" ? OfficeRuntime.storage : null;
    if (storage && typeof storage.setItem === "function") {
      return storage.setItem(key, json);
    }
    return Promise.resolve();
  }

  function storageRemove(key) {
    try {
      window.localStorage.removeItem(key);
    } catch (_e) {
      // ignore
    }
    var storage = typeof OfficeRuntime !== "undefined" ? OfficeRuntime.storage : null;
    if (storage && typeof storage.removeItem === "function") {
      return storage.removeItem(key);
    }
    return Promise.resolve();
  }

  /**
   * Ecrit la formule locale puis force le recalcul. Les fonctions personnalisees (PALO.ENAME)
   * restent souvent figees si application.calculate() est appele avant le sync qui applique la formule.
   */
  async function applyFormulaLocalAndRecalculate(context, cell, newFormulaLocal) {
    cell.formulasLocal = [[newFormulaLocal]];
    await context.sync();
    try {
      if (typeof cell.calculate === "function") {
        cell.calculate();
        await context.sync();
        return;
      }
    } catch (_rangeCalc) {
      // Continuer avec le recalcul global.
    }
    try {
      if (
        typeof context.workbook.application.calculate === "function"
        && typeof Excel !== "undefined"
        && Excel.CalculationType
        && Excel.CalculationType.fullRebuild
      ) {
        context.workbook.application.calculate(Excel.CalculationType.fullRebuild);
        await context.sync();
        return;
      }
    } catch (_full) {
      // Recalculate leger.
    }
    try {
      if (
        typeof context.workbook.application.calculate === "function"
        && typeof Excel !== "undefined"
        && Excel.CalculationType
      ) {
        context.workbook.application.calculate(Excel.CalculationType.recalculate);
        await context.sync();
      }
    } catch (_recalc) {
      // La formule est au moins persistee.
    }
  }

  function normalizeExcelAddress(address) {
    return String(address || "").replace(/\$/g, "").toUpperCase();
  }

  function qualifyExcelAddress(sheetName, address) {
    var raw = String(address || "").trim();
    if (!raw || raw.indexOf("!") >= 0) {
      return raw;
    }
    var name = String(sheetName || "").trim();
    if (!name) {
      return raw;
    }
    if (/[^A-Za-z0-9_]/.test(name)) {
      name = "'" + name.replace(/'/g, "''") + "'";
    }
    return name + "!" + raw;
  }

  function logDatacPerf(step, fields) {
    try {
      if (window.PaloOffice && typeof window.PaloOffice.datacPerf === "function") {
        window.PaloOffice.datacPerf(step, fields || null);
      }
    } catch (_log) {
      // Le journal ne bloque pas le recalcul.
    }
  }

  function prepareDatacArgument(context, workbook, worksheet, rawArgText) {
    var trimmed = String(rawArgText || "").trim();
    var lit = window.paloEnameRibbon.tryUnquoteStringLiteral(trimmed);
    if (lit !== null) {
      return { kind: "literal", value: lit };
    }
    if (!trimmed) {
      return { kind: "literal", value: "" };
    }
    if (/^[-+]?\d+(?:[.,]\d+)?$/.test(trimmed)) {
      var n = Number(trimmed.replace(",", "."));
      return { kind: "literal", value: Number.isFinite(n) ? String(n) : trimmed };
    }
    try {
      var range = trimmed.indexOf("!") === -1 ? worksheet.getRange(trimmed) : workbook.getRange(trimmed);
      range.load(["values"]);
      return { kind: "range", range: range };
    } catch (_e) {
      return { kind: "bad" };
    }
  }

  function readDatacPrepared(prepared) {
    if (!prepared || prepared.kind === "bad") {
      return null;
    }
    if (prepared.kind === "literal") {
      return prepared.value;
    }
    var vals = prepared.range.values;
    if (!vals || !vals[0] || vals[0].length === 0) {
      return "";
    }
    var v = vals[0][0];
    if (v == null || v === "") {
      return "";
    }
    if (typeof v === "number") {
      return Number.isFinite(v) ? String(v) : null;
    }
    if (typeof v === "boolean") {
      return v ? "TRUE" : "FALSE";
    }
    if (typeof v === "object") {
      return null;
    }
    var text = String(v).trim();
    if (text.charAt(0) === "#") {
      return null;
    }
    return text;
  }

  /**
   * Cellules PALO.DATAC de la selection, arguments resolus (litteraux ou references).
   * Meme principe que l'ancien add-in : la formule reste en place, l'ecriture passe par le cube.
   */
  async function collectDatacWriteCells(context, workbook, sheet, activeCell, selection) {
    var scan = window.paloEnameRibbon.scanRangeForPaloDatacHits(selection, 5000);
    var hits = scan.hits.slice();
    var truncated = scan.truncated;
    if (hits.length > DATAC_WRITE_MAX_CELLS) {
      hits = hits.slice(0, DATAC_WRITE_MAX_CELLS);
      truncated = true;
    }
    var located = hits.map(function (hit) {
      var range = selection.getCell(hit.row, hit.col);
      range.load("address");
      return { hit: hit, range: range };
    });
    activeCell.load("address");
    sheet.load("name");
    await context.sync();

    var sheetName = String(sheet.name || "");
    var activeAddr = normalizeExcelAddress(qualifyExcelAddress(sheetName, activeCell.address));
    var hasActive = false;
    located.forEach(function (item) {
      item.address = qualifyExcelAddress(sheetName, item.range.address);
      if (normalizeExcelAddress(item.address) === activeAddr) {
        hasActive = true;
      }
    });
    var activeFormula = window.paloEnameRibbon.readFormulaBestForPalo(activeCell);
    if (!hasActive && window.paloEnameRibbon.mentionsPaloDatacInText(activeFormula)) {
      located.unshift({
        hit: { formula: activeFormula, row: 0, col: 0 },
        address: qualifyExcelAddress(sheetName, activeCell.address),
        range: activeCell
      });
    }

    var preparedCells = [];
    located.forEach(function (item) {
      var parsed = window.paloEnameRibbon.parsePaloDatacCall(item.hit.formula);
      if (!parsed || parsed.args.length < 2) {
        preparedCells.push({ address: item.address, error: true });
        return;
      }
      preparedCells.push({
        address: item.address,
        preps: parsed.args.map(function (arg) {
          return prepareDatacArgument(context, workbook, sheet, arg.raw);
        })
      });
    });
    await context.sync();

    var cells = [];
    var skipped = 0;
    preparedCells.forEach(function (item) {
      if (item.error) {
        skipped += 1;
        return;
      }
      var values = item.preps.map(readDatacPrepared);
      if (values.some(function (v) { return v == null; })) {
        skipped += 1;
        return;
      }
      var servdb = String(values[0] || "").trim();
      var cube = String(values[1] || "").trim();
      var coordinates = values.slice(2).map(function (v) { return String(v).trim(); });
      if (!servdb || !cube || coordinates.some(function (c) { return !c; })) {
        skipped += 1;
        return;
      }
      cells.push({
        address: item.address,
        servdb: servdb,
        cube: cube,
        coordinates: coordinates,
        active: normalizeExcelAddress(item.address) === activeAddr
      });
    });
    if (cells.length && !cells.some(function (c) { return c.active; })) {
      cells[0].active = true;
    }
    return {
      found: located.length > 0,
      cells: cells,
      truncated: truncated,
      skipped: skipped
    };
  }

  async function elementTypeForWrite(ctx, dimensionName, elementName, cache) {
    var key = ctx.database + "\0" + dimensionName + "\0" + elementName;
    if (Object.prototype.hasOwnProperty.call(cache, key)) {
      return cache[key];
    }
    var info = null;
    try {
      info = await ctx.client.elementInfo(ctx.sid, ctx.database, dimensionName, elementName);
    } catch (err) {
      if (isTransportError(err)) {
        throw err;
      }
      throw new Error(missingElementMessage(elementName, dimensionName));
    }
    if (!elementInfoIsUsable(info)) {
      throw new Error(missingElementMessage(elementName, dimensionName));
    }
    cache[key] = Number(info.type);
    return cache[key];
  }

  /**
   * Action du popup. splash/compute : /cell/replace. copy : /cell/copy.
   * literal = le nombre saisi. Les autres calculs lisent la cellule d'abord.
   */
  var WRITE_ACTIONS = {
    replace: { splash: 0, compute: "literal" },
    add: { splash: 2, compute: "literal" },
    "add-pct": { splash: 2, compute: "delta-pct" },
    "set-pct": { splash: 0, compute: "ratio" },
    spread: { splash: 1, compute: "literal" },
    "add-total": { splash: 1, compute: "add-current" },
    "add-total-pct": { splash: 1, compute: "grow-pct" },
    "set-total-pct": { splash: 1, compute: "ratio" },
    "set-bases": { splash: 3, compute: "literal" },
    "add-bases": { splash: 2, compute: "literal" },
    like: { copy: "value" },
    "like-add": { copy: "add" },
    "like-pct": { copy: "pct" },
    "like-add-pct": { copy: "add-pct" },
    copy: { copy: "plain" }
  };

  function requireWriteAmount(value) {
    var n = window.PaloOffice.parsePaloNumericInput(value);
    if (n == null) {
      throw new Error(NUMERIC_CELL_TEXT_MESSAGE);
    }
    return n;
  }

  async function readPlanNumber(plan, idPath) {
    var storedType = 0;
    try {
      storedType = await plan.ctx.client.cellStoredTypeByIds(
        plan.ctx.sid,
        plan.ctx.database,
        plan.cell.cube,
        idPath
      );
    } catch (errRead) {
      if (isTransportError(errRead)) {
        throw errRead;
      }
    }
    if (storedType === 2) {
      throw new Error(NUMERIC_CELL_TEXT_MESSAGE);
    }
    var raw = 0;
    try {
      raw = await plan.ctx.client.cellValueByIds(
        plan.ctx.sid,
        plan.ctx.database,
        plan.cell.cube,
        idPath
      );
    } catch (errValue) {
      if (isTransportError(errValue)) {
        throw errValue;
      }
    }
    if (typeof raw === "number" && Number.isFinite(raw)) {
      return raw;
    }
    var parsed = window.PaloOffice.parsePaloNumericInput(raw);
    return parsed == null ? 0 : parsed;
  }

  /**
   * Type d'element avant ecriture : 2 → string quotee. 1 ou 4 → nombre non quote.
   * Une cellule Palo deja string sur une mesure numerique est videe puis reecrite en numeric.
   */
  async function executeDatacWrite(cells, mode, action, value, pathOverrides) {
    var spec = WRITE_ACTIONS[action];
    if (!spec) {
      throw new Error("Action inconnue.");
    }
    var targets = cells;
    if (mode !== "bulk") {
      targets = cells.filter(function (c) { return c.active; });
      if (!targets.length) {
        targets = cells.slice(0, 1);
      }
    }
    var groups = {};
    targets.forEach(function (cell) {
      var key = cell.servdb + "\0" + cell.cube;
      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(cell);
    });
    var manager = window.PaloOffice.createConnectionManager();
    var keys = Object.keys(groups);
    var typeCache = {};
    var plans = [];
    var amount = spec.copy === "plain" ? null : requireWriteAmount(value);
    var i;
    var j;
    var k;
    for (i = 0; i < keys.length; i += 1) {
      var group = groups[keys[i]];
      var ctx = await manager.getClientAndContext(group[0].servdb);
      var dimNames = await manager.getCubeDimensionNamesOrdered(
        ctx.connectionName,
        ctx.sid,
        ctx.client,
        ctx.database,
        group[0].cube
      );
      var groupPlans = [];
      for (j = 0; j < group.length; j += 1) {
        var cell = group[j];
        if (cell.coordinates.length !== dimNames.length) {
          throw new Error(
            "Nombre de coordonnees (" + cell.coordinates.length
            + ") different du nombre de dimensions du cube " + cell.cube
            + " (" + dimNames.length + ")."
          );
        }
        var types = [];
        for (k = 0; k < dimNames.length; k += 1) {
          types.push(await elementTypeForWrite(ctx, dimNames[k], cell.coordinates[k], typeCache));
        }
        var kind = window.PaloOffice.paloPathStorageKind(types);
        if (kind !== "numeric") {
          throw new Error("Cette action s'applique a une mesure numerique.");
        }
        var sourceCoordinates = cell.coordinates.slice();
        if (pathOverrides) {
          for (k = 0; k < dimNames.length; k += 1) {
            if (pathOverrides[dimNames[k]]) {
              sourceCoordinates[k] = String(pathOverrides[dimNames[k]]);
            }
          }
        }
        groupPlans.push({
          ctx: ctx,
          cell: cell,
          kind: kind,
          sourceCoordinates: sourceCoordinates
        });
      }
      if (mode === "bulk") {
        var idPaths = await manager.buildCellIdPathsListFromSegments(
          ctx.connectionName,
          ctx.sid,
          ctx.client,
          ctx.database,
          group[0].cube,
          group.map(function (c) { return c.coordinates; })
        );
        for (j = 0; j < groupPlans.length; j += 1) {
          groupPlans[j].idPath = idPaths[j];
        }
      } else {
        for (j = 0; j < groupPlans.length; j += 1) {
          groupPlans[j].idPath = await manager.buildCellIdPathFromSegments(
            ctx.connectionName,
            ctx.sid,
            ctx.client,
            ctx.database,
            groupPlans[j].cell.cube,
            groupPlans[j].cell.coordinates
          );
        }
      }
      plans = plans.concat(groupPlans);
    }

    if (spec.copy) {
      var copied = 0;
      for (i = 0; i < plans.length; i += 1) {
        var sourcePath = await manager.buildCellIdPathFromSegments(
          plans[i].ctx.connectionName,
          plans[i].ctx.sid,
          plans[i].ctx.client,
          plans[i].ctx.database,
          plans[i].cell.cube,
          plans[i].sourceCoordinates
        );
        var copyValue = null;
        if (spec.copy === "value") {
          copyValue = amount;
        } else if (spec.copy !== "plain") {
          var sourceNumber = await readPlanNumber(plans[i], sourcePath);
          if (spec.copy === "add") {
            copyValue = sourceNumber + amount;
          } else if (spec.copy === "pct") {
            copyValue = sourceNumber * amount / 100;
          } else {
            copyValue = sourceNumber * (100 + amount) / 100;
          }
        }
        await plans[i].ctx.client.cellCopyByIds(
          plans[i].ctx.sid,
          plans[i].ctx.database,
          plans[i].cell.cube,
          sourcePath,
          plans[i].idPath,
          copyValue
        );
        copied += 1;
      }
      return {
        written: copied,
        addresses: targets.map(function (c) { return c.address; })
      };
    }

    for (i = 0; i < plans.length; i += 1) {
      if (spec.compute === "literal") {
        plans[i].writeValue = amount;
      } else {
        var currentNumber = await readPlanNumber(plans[i], plans[i].idPath);
        if (spec.compute === "add-current") {
          plans[i].writeValue = currentNumber + amount;
        } else if (spec.compute === "grow-pct") {
          plans[i].writeValue = currentNumber * (100 + amount) / 100;
        } else {
          plans[i].writeValue = currentNumber * amount / 100;
        }
      }
    }

    if (spec.compute === "literal") {
      for (i = 0; i < plans.length; i += 1) {
        var storedType = 0;
        try {
          storedType = await plans[i].ctx.client.cellStoredTypeByIds(
            plans[i].ctx.sid,
            plans[i].ctx.database,
            plans[i].cell.cube,
            plans[i].idPath
          );
        } catch (errType) {
          if (isTransportError(errType)) {
            throw errType;
          }
          storedType = 0;
        }
        if (storedType !== 2) {
          continue;
        }
        var cleared = false;
        try {
          cleared = await plans[i].ctx.client.cellClearByIds(
            plans[i].ctx.sid,
            plans[i].ctx.database,
            plans[i].cell.cube,
            plans[i].idPath
          );
        } catch (errClear) {
          if (isTransportError(errClear)) {
            throw errClear;
          }
          cleared = false;
        }
        if (!cleared) {
          throw new Error(NUMERIC_CELL_TEXT_MESSAGE);
        }
      }
    }

    var written = 0;
    var splashMode = spec.splash;
    if (mode === "bulk" && spec.compute === "literal") {
      for (i = 0; i < keys.length; i += 1) {
        var bulkPlans = plans.filter(function (plan) {
          return plan.cell.servdb + "\0" + plan.cell.cube === keys[i];
        });
        if (!bulkPlans.length) {
          continue;
        }
        await bulkPlans[0].ctx.client.cellReplaceBulk(
          bulkPlans[0].ctx.sid,
          bulkPlans[0].ctx.database,
          bulkPlans[0].cell.cube,
          bulkPlans.map(function (plan) { return plan.idPath; }),
          bulkPlans.map(function (plan) { return plan.writeValue; }),
          splashMode,
          bulkPlans.map(function () { return "numeric"; })
        );
        written += bulkPlans.length;
      }
    } else {
      for (i = 0; i < plans.length; i += 1) {
        var ok = await plans[i].ctx.client.cellReplaceByIds(
          plans[i].ctx.sid,
          plans[i].ctx.database,
          plans[i].cell.cube,
          plans[i].idPath,
          plans[i].writeValue,
          splashMode,
          true,
          "numeric"
        );
        if (!ok) {
          throw new Error("cell/replace a refuse l'ecriture (" + plans[i].cell.address + ").");
        }
        written += 1;
      }
    }
    return {
      written: written,
      addresses: targets.map(function (c) { return c.address; })
    };
  }

  function isConsolidatedElement(info) {
    if (!info) {
      return false;
    }
    if (Number(info.type) === 4) {
      return true;
    }
    return Array.isArray(info.childIds) && info.childIds.length > 0;
  }

  function missingElementMessage(elementName, dimensionName) {
    return "Element " + elementName + " introuvable dans la dimension " + dimensionName;
  }

  function isTransportError(err) {
    var msg = err && err.message ? String(err.message) : String(err || "");
    return msg.indexOf("Timeout HTTP") === 0
      || msg.indexOf("Impossible de joindre") !== -1
      || msg.indexOf("Failed to fetch") !== -1
      || msg.indexOf("NetworkError") !== -1
      || msg.indexOf("Load failed") !== -1;
  }

  function elementInfoIsUsable(info) {
    if (!info) {
      return false;
    }
    if (!/^\d+$/.test(String(info.id || "").trim())) {
      return false;
    }
    return Boolean(String(info.name || "").trim());
  }

  /**
   * Avant le popup : un element consolide (regroupement) sur le chemin de la cellule active
   * selectionne cell/replace_bulk. Sinon, cell/replace uniquement.
   */
  async function classifyDatacWriteMode(cells) {
    var active = null;
    var i;
    for (i = 0; i < cells.length; i += 1) {
      if (cells[i].active) {
        active = cells[i];
      }
    }
    if (!active) {
      active = cells[0];
    }
    var manager = window.PaloOffice.createConnectionManager();
    var ctx = await manager.getClientAndContext(active.servdb);
    var dimNames = await manager.getCubeDimensionNamesOrdered(
      ctx.connectionName,
      ctx.sid,
      ctx.client,
      ctx.database,
      active.cube
    );
    if (active.coordinates.length !== dimNames.length) {
      throw new Error(
        "Nombre de coordonnees (" + active.coordinates.length
        + ") different du nombre de dimensions du cube " + active.cube
        + " (" + dimNames.length + ")."
      );
    }
    var consolidated = [];
    var missing = [];
    for (i = 0; i < dimNames.length; i += 1) {
      var elementName = active.coordinates[i];
      var dimensionName = dimNames[i];
      var info = null;
      try {
        info = await ctx.client.elementInfo(
          ctx.sid,
          ctx.database,
          dimensionName,
          elementName
        );
      } catch (err) {
        if (isTransportError(err)) {
          throw err;
        }
        missing.push({
          element: elementName,
          dimension: dimensionName,
          message: missingElementMessage(elementName, dimensionName)
        });
        continue;
      }
      if (!elementInfoIsUsable(info)) {
        missing.push({
          element: elementName,
          dimension: dimensionName,
          message: missingElementMessage(elementName, dimensionName)
        });
        continue;
      }
      if (isConsolidatedElement(info)) {
        consolidated.push(dimensionName + " / " + String(info.name || elementName));
      }
    }
    return {
      mode: consolidated.length ? "bulk" : "replace",
      consolidated: consolidated,
      missing: missing
    };
  }

  function wrapDatacFormula(formula) {
    if (typeof formula !== "string" || formula.charAt(0) !== "=") {
      return formula;
    }
    return "=IF(TRUE," + formula.slice(1) + ")";
  }

  async function recalculateDatacAddresses(addresses) {
    if (!addresses || !addresses.length) {
      return;
    }
    logDatacPerf("recalcul", { phase: "demande", n: addresses.length });
    await Excel.run(async function (context) {
      var ranges = [];
      var i;
      for (i = 0; i < addresses.length; i += 1) {
        ranges.push(context.workbook.getRange(addresses[i]));
      }
      for (i = 0; i < ranges.length; i += 1) {
        ranges[i].load("formulas");
      }
      await context.sync();
      var originals = [];
      for (i = 0; i < ranges.length; i += 1) {
        originals.push(ranges[i].formulas);
        ranges[i].formulas = ranges[i].formulas.map(function (row) {
          return row.map(wrapDatacFormula);
        });
      }
      await context.sync();
      for (i = 0; i < ranges.length; i += 1) {
        ranges[i].formulas = originals[i];
      }
      await context.sync();
      logDatacPerf("recalcul", { phase: "ok", api: "formule", n: ranges.length });
    });
  }

  async function loadDatacLikeDimensions(cells) {
    var active = null;
    var i;
    for (i = 0; i < cells.length; i += 1) {
      if (cells[i].active) {
        active = cells[i];
      }
    }
    if (!active) {
      active = cells[0];
    }
    var manager = window.PaloOffice.createConnectionManager();
    var ctx = await manager.getClientAndContext(active.servdb);
    var dimNames = await manager.getCubeDimensionNamesOrdered(
      ctx.connectionName,
      ctx.sid,
      ctx.client,
      ctx.database,
      active.cube
    );
    var dimensions = [];
    for (i = 0; i < dimNames.length; i += 1) {
      var elems = await ctx.client.dimensionElements(ctx.sid, ctx.database, dimNames[i]);
      var names = [];
      var seen = {};
      var j;
      for (j = 0; j < elems.length; j += 1) {
        var name = String(elems[j].name || "").trim();
        if (!name || seen[name]) {
          continue;
        }
        seen[name] = true;
        names.push(name);
      }
      var current = String(active.coordinates[i] || "");
      if (current && !seen[current]) {
        names.unshift(current);
      }
      dimensions.push({
        name: dimNames[i],
        current: current,
        elements: names
      });
    }
    return dimensions;
  }

  function openDatacWriteDialog(event, pickContext) {
    var payload = {
      cells: pickContext.cells,
      truncated: Boolean(pickContext.truncated),
      skipped: pickContext.skipped || 0,
      writeMode: pickContext.writeMode === "bulk" ? "bulk" : "replace",
      consolidated: pickContext.consolidated || [],
      missing: pickContext.missing || [],
      dimensions: pickContext.dimensions || []
    };
    return storageRemove(DATAC_WRITE_STORAGE_KEY).then(function () {
      return storageSetJson(DATAC_WRITE_STORAGE_KEY, payload);
    }).then(function () {
      return new Promise(function (resolve) {
        var dialogUrl = new URL("palo-datac-write.html", window.location.href);
        dialogUrl.searchParams.set("v", PALO_PLUGIN_VERSION);
        Office.context.ui.displayDialogAsync(
          dialogUrl.href,
          { height: 72, width: 40, displayInIframe: true },
          function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
              paloUserNotify(
                "Impossible d'ouvrir le menu d'ecriture PALO.DATAC."
                  + (asyncResult.errorMessage ? " " + asyncResult.errorMessage : ""),
                "error"
              );
              storageRemove(DATAC_WRITE_STORAGE_KEY);
              complete(event);
              resolve();
              return;
            }
            var dlg = asyncResult.value;
            var cleaned = false;
            var ribbonEventCompleted = false;
            var refreshAddresses = (payload.cells || []).map(function (cell) {
              return cell && cell.address ? String(cell.address) : "";
            }).filter(Boolean);
            var reloadAddresses = refreshAddresses.slice();
            var reloadStarted = false;
            function launchReload() {
              if (reloadStarted || !reloadAddresses || !reloadAddresses.length) {
                return;
              }
              reloadStarted = true;
              var addresses = reloadAddresses.slice();
              var attempt = 0;
              function run() {
                attempt += 1;
                recalculateDatacAddresses(addresses).catch(function (err) {
                  var message = err && err.message ? err.message : String(err);
                  if (attempt < 3) {
                    logDatacPerf("recalcul", { phase: "retry", attempt: attempt, error: message });
                    setTimeout(run, 500);
                    return;
                  }
                  logDatacPerf("recalcul", { phase: "erreur", error: message });
                  paloUserNotify("Excel n'a pas relu les formules PALO.DATAC. " + message, "error");
                });
              }
              setTimeout(run, 400);
            }
            function finishRibbonUiEvent() {
              if (ribbonEventCompleted || !event) {
                return;
              }
              ribbonEventCompleted = true;
              complete(event);
            }
            setTimeout(finishRibbonUiEvent, 200);

            function cleanup() {
              if (cleaned) {
                return;
              }
              cleaned = true;
              storageRemove(DATAC_WRITE_STORAGE_KEY);
            }

            function postResult(ok, message) {
              if (typeof dlg.messageChild !== "function") {
                paloUserNotify(message, ok ? "ok" : "error");
                return;
              }
              try {
                dlg.messageChild(JSON.stringify({ type: "result", ok: ok, message: message }));
              } catch (_e) {
                paloUserNotify(message, ok ? "ok" : "error");
              }
            }

            dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
              var msg = String(arg.message || "");
              if (msg === "palo-datac-write-ready") {
                if (typeof dlg.messageChild === "function") {
                  try {
                    dlg.messageChild(JSON.stringify({
                      type: "init",
                      cells: payload.cells,
                      truncated: payload.truncated,
                      skipped: payload.skipped,
                      writeMode: payload.writeMode,
                      consolidated: payload.consolidated,
                      missing: payload.missing,
                      dimensions: payload.dimensions
                    }));
                  } catch (_e) {
                    // Le dialogue relit le stockage.
                  }
                }
                return;
              }
              var o = null;
              try {
                o = JSON.parse(msg);
              } catch (_parse) {
                o = null;
              }
              if (o && o.type === "cancel") {
                cleanup();
                try {
                  dlg.close();
                } catch (_cancelClose) {
                  // Le dialogue est deja ferme.
                }
                return;
              }
              if (!o || o.type !== "write") {
                return;
              }
              if (payload.missing && payload.missing.length) {
                postResult(false, payload.missing.map(function (item) {
                  return item.message;
                }).join("\n"));
                return;
              }
              var mode = payload.writeMode === "bulk" ? "bulk" : "replace";
              executeDatacWrite(
                payload.cells,
                mode,
                o.action,
                o.value == null ? "" : String(o.value),
                o.path || null
              )
                .then(function (result) {
                  if (result && result.addresses && result.addresses.length) {
                    reloadAddresses = result.addresses;
                  }
                  cleanup();
                  try {
                    dlg.close();
                  } catch (_closeOk) {
                    launchReload();
                  }
                  setTimeout(launchReload, 700);
                })
                .catch(function (err) {
                  postResult(false, err && err.message ? err.message : String(err));
                });
            });

            dlg.addEventHandler(Office.EventType.DialogEventReceived, function (ev) {
              if (ev && (ev.error === 12004 || ev.error === 12006)) {
                cleanup();
                launchReload();
              }
            });
            resolve();
          }
        );
      });
    });
  }

  async function paloRibbonAction(event) {
    try {
      if (!window.paloEnameRibbon || !window.PaloOffice) {
        await paloUserNotify(
          "Palo action indisponible : script ruban incomplet (palo-ename-ribbon ou connexion Palo). Rechargez le volet ou reimportez le manifeste.",
          "error"
        );
        complete(event);
        return;
      }

      async function resolvePickContextOnce() {
        return Excel.run(async function (context) {
          var workbook = context.workbook;
          var sheet = workbook.worksheets.getActiveWorksheet();
          var cell = workbook.getActiveCell();
          var sel = workbook.getSelectedRange();
          cell.load(["formulasLocal", "formulas"]);
          sel.load(["formulasLocal", "formulas"]);
          await context.sync();

          function readFormula(range) {
            return window.paloEnameRibbon.readFormulaBestForPalo(range);
          }

          function buildResultFromFormula(formulaLocal) {
            var fl = window.paloEnameRibbon.normalizeRibbonFormula(formulaLocal);
            if (!fl || fl.charAt(0) !== "=") {
              return { skip: true, reason: "empty_formula" };
            }
            var parsed = window.paloEnameRibbon.parsePaloEnameCall(fl);
            if (!parsed || parsed.args.length < 3) {
              return { skip: true, reason: "not_palo_ename" };
            }
            var p1 = window.paloEnameRibbon.prepareResolveArgument(context, workbook, sheet, parsed.args[0].raw);
            var p2 = window.paloEnameRibbon.prepareResolveArgument(context, workbook, sheet, parsed.args[1].raw);
            return { skip: false, p1: p1, p2: p2 };
          }

          // Cellule active PALO.ENAME : selecteur d'element (comportement inchange).
          var activeFormula = readFormula(cell);
          var fromActive = buildResultFromFormula(activeFormula);
          if (!fromActive.skip) {
            await context.sync();
            return {
              skip: false,
              kind: "ename",
              servdb: window.paloEnameRibbon.readPrepared(fromActive.p1).trim(),
              dimension: window.paloEnameRibbon.readPrepared(fromActive.p2).trim()
            };
          }

          // Ecriture seulement si la cellule active est PALO.DATAC.
          // Une selection mixte ne doit pas prendre la place du selecteur PALO.ENAME.
          if (window.paloEnameRibbon.mentionsPaloDatacInText(activeFormula)) {
            var datac = await collectDatacWriteCells(context, workbook, sheet, cell, sel);
            if (datac.cells.length) {
              return {
                skip: false,
                kind: "datac",
                cells: datac.cells,
                truncated: datac.truncated,
                skipped: datac.skipped
              };
            }
            return { skip: true, reason: "datac_unresolved" };
          }

          var fromSel = buildResultFromFormula(readFormula(sel));
          if (fromSel.skip) {
            return fromActive;
          }
          await context.sync();
          return {
            skip: false,
            kind: "ename",
            servdb: window.paloEnameRibbon.readPrepared(fromSel.p1).trim(),
            dimension: window.paloEnameRibbon.readPrepared(fromSel.p2).trim()
          };
        });
      }

      // Sur certains hots (ruban/menu contextuel), le premier tick peut encore etre stale.
      // Retry auto: evite d'exiger un 2e clic utilisateur.
      var pickContext = await resolvePickContextOnce();
      if (pickContext.skip && (pickContext.reason === "empty_formula" || pickContext.reason === "not_palo_ename")) {
        await new Promise(function (resolve) { setTimeout(resolve, 120); });
        var retryContext = await resolvePickContextOnce();
        if (!retryContext.skip) {
          pickContext = retryContext;
        }
      }
      if (pickContext.skip && (pickContext.reason === "empty_formula" || pickContext.reason === "not_palo_ename")) {
        await new Promise(function (resolve) { setTimeout(resolve, 280); });
        var retryContext2 = await resolvePickContextOnce();
        if (!retryContext2.skip) {
          pickContext = retryContext2;
        }
      }

      if (!pickContext.skip && pickContext.kind === "datac") {
        var classified = await classifyDatacWriteMode(pickContext.cells);
        pickContext.writeMode = classified.mode;
        pickContext.consolidated = classified.consolidated;
        pickContext.missing = classified.missing;
        pickContext.dimensions = [];
        if (!(classified.missing && classified.missing.length)) {
          pickContext.dimensions = await loadDatacLikeDimensions(pickContext.cells);
        }
        await openDatacWriteDialog(event, pickContext);
        return;
      }

      if (pickContext.skip) {
        await paloUserNotify(
          pickContext.reason === "datac_unresolved"
            ? "PALO.DATAC trouvee, mais les coordonnees n'ont pas pu etre lues (reference vide ou formule illisible)."
            : "Palo action fonctionne sur une cellule PALO.ENAME(...) ou PALO.DATAC(...).",
          "error"
        );
        complete(event);
        return;
      }

      if (!pickContext.servdb || !pickContext.dimension) {
        await paloUserNotify("PALO.ENAME: servdb ou dimension vide apres resolution des references.", "error");
        complete(event);
        return;
      }

      var manager = window.PaloOffice.createConnectionManager();
      var ctx = await manager.getClientAndContext(pickContext.servdb);

      function mapDimensionElementsToPickerItems(els) {
        return els
          .map(function (e) {
            var hasChildren = Array.isArray(e.childIds) && e.childIds.length > 0;
            var isConsolidated = Number(e.type || 0) === 4 || hasChildren;
            return {
              name: String(e.name || ""),
              isConsolidated: isConsolidated
            };
          })
          .filter(function (e) {
            return Boolean(e.name);
          })
          .sort(function (a, b) {
            return a.name.localeCompare(b.name, "fr", { sensitivity: "base" });
          });
      }

      var dimRows = await ctx.client.databaseDimensions(ctx.sid, ctx.database);
      var dimNames = dimRows
        .map(function (d) {
          return String(d.name || "");
        })
        .filter(Boolean)
        .sort(function (a, b) {
          return a.localeCompare(b, "fr", { sensitivity: "base" });
        });
      if (pickContext.dimension && dimNames.indexOf(pickContext.dimension) === -1) {
        dimNames.push(pickContext.dimension);
        dimNames.sort(function (a, b) {
          return a.localeCompare(b, "fr", { sensitivity: "base" });
        });
      }

      if (!dimNames.length) {
        await paloUserNotify("Aucune dimension pour la base courante.", "error");
        complete(event);
        return;
      }

      var els = await ctx.client.dimensionElements(ctx.sid, ctx.database, pickContext.dimension);
      var items = mapDimensionElementsToPickerItems(els);

      if (!items.length) {
        await paloUserNotify("Aucun element pour la dimension \"" + pickContext.dimension + "\".", "error");
        complete(event);
        return;
      }

      var pickerPayload = {
        dimensions: dimNames,
        currentDimension: pickContext.dimension,
        elements: items
      };

      await storageRemove(PICKER_STORAGE_KEY);
      await storageSetJson(PICKER_STORAGE_KEY, pickerPayload);

      var dialogUrl = new URL("palo-ename-picker.html", window.location.href);
      dialogUrl.searchParams.set("v", PALO_PLUGIN_VERSION);

      Office.context.ui.displayDialogAsync(
        dialogUrl.href,
        { height: 62, width: 32, displayInIframe: true },
        function (asyncResult) {
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            var errDetail = asyncResult.errorMessage || "";
            var errCode = asyncResult.errorCode != null ? String(asyncResult.errorCode) : "";
            paloUserNotify(
              "Impossible d'ouvrir la fenetre de selection."
                + (errCode ? " Code: " + errCode + "." : "")
                + (errDetail ? " " + errDetail : ""),
              "error"
            );
            storageRemove(PICKER_STORAGE_KEY);
            complete(event);
            return;
          }

          var dlg = asyncResult.value;
          var cleaned = false;
          var ribbonEventCompleted = false;
          function finishRibbonUiEvent() {
            if (ribbonEventCompleted || !event) {
              return;
            }
            ribbonEventCompleted = true;
            complete(event);
          }
          setTimeout(finishRibbonUiEvent, 200);

          function cleanupPickerStorage() {
            if (cleaned) {
              return;
            }
            cleaned = true;
            storageRemove(PICKER_STORAGE_KEY);
          }

          dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
            var msg = String(arg.message || "");
            if (msg === "palo-ename-ready") {
              if (typeof dlg.messageChild === "function") {
                try {
                  dlg.messageChild(
                    JSON.stringify({
                      type: "init",
                      dimensions: dimNames,
                      currentDimension: pickContext.dimension,
                      elements: items
                    })
                  );
                } catch (_e) {
                  // Le dialogue retombe sur localStorage / OfficeRuntime.
                }
              }
              return;
            }
            var o = null;
            try {
              o = JSON.parse(msg);
            } catch (_e) {
              o = null;
            }
            if (o && o.type === "loadDimension" && o.dimension) {
              manager
                .getClientAndContext(pickContext.servdb)
                .then(function (ctx2) {
                  return ctx2.client.dimensionElements(
                    ctx2.sid,
                    ctx2.database,
                    String(o.dimension || "").trim()
                  );
                })
                .then(function (els2) {
                  var items2 = mapDimensionElementsToPickerItems(els2);
                  if (typeof dlg.messageChild === "function") {
                    dlg.messageChild(
                      JSON.stringify({
                        type: "elements",
                        dimension: String(o.dimension || "").trim(),
                        elements: items2
                      })
                    );
                  }
                })
                .catch(function (err) {
                  if (typeof dlg.messageChild === "function") {
                    dlg.messageChild(
                      JSON.stringify({
                        type: "loadDimensionError",
                        message: err && err.message ? err.message : String(err)
                      })
                    );
                  }
                });
              return;
            }
            var chosen = "";
            var chosenDim = "";
            if (o && o.element) {
              chosen = String(o.element);
              chosenDim = o.dimension != null ? String(o.dimension).trim() : "";
            }
            if (!chosen) {
              try {
                chosen = String(msg || "");
              } catch (_e2) {
                chosen = "";
              }
            }
            if (!chosen) {
              try {
                dlg.close();
              } catch (_c) {
                // ignore
              }
              cleanupPickerStorage();
              return;
            }

            // Fermer tout de suite : sinon Excel considere souvent le dialogue encore ouvert
            // et displayDialogAsync echoue au clic suivant (ex. erreur 12007).
            try {
              dlg.close();
            } catch (_closeEarly) {
              // ignore
            }
            cleanupPickerStorage();

            Excel.run(async function (context) {
              var cell = context.workbook.getActiveCell();
              var sel = context.workbook.getSelectedRange();
              cell.load(["formulasLocal", "formulas"]);
              sel.load(["formulasLocal", "formulas"]);
              await context.sync();
              var metaC = window.paloEnameRibbon.readFormulaBestMetaForPalo(cell);
              var targetRange = cell.getCell(metaC.row, metaC.col);
              var parsed2 = window.paloEnameRibbon.parsePaloEnameCall(metaC.formula);
              if (!parsed2 || parsed2.args.length < 3) {
                var metaS = window.paloEnameRibbon.readFormulaBestMetaForPalo(sel);
                parsed2 = window.paloEnameRibbon.parsePaloEnameCall(metaS.formula);
                targetRange = sel.getCell(metaS.row, metaS.col);
              }
              if (!parsed2 || parsed2.args.length < 3) {
                paloUserNotify("Cellule PALO.ENAME introuvable pour appliquer le choix.", "error");
                return;
              }
              var newLit = window.paloEnameRibbon.buildExcelStringLiteral(chosen);
              var base = parsed2.formula;
              var updated = base.slice(0, parsed2.args[2].start) + newLit + base.slice(parsed2.args[2].end);
              if (chosenDim && chosenDim !== String(pickContext.dimension || "").trim()) {
                var dimLit = window.paloEnameRibbon.buildExcelStringLiteral(chosenDim);
                updated = updated.slice(0, parsed2.args[1].start) + dimLit + updated.slice(parsed2.args[1].end);
              }
              await applyFormulaLocalAndRecalculate(context, targetRange, updated);
            }).catch(function (err) {
              paloUserNotify(err && err.message ? err.message : String(err), "error");
            });
          });

          dlg.addEventHandler(Office.EventType.DialogEventReceived, function (ev) {
            if (ev && (ev.error === 12004 || ev.error === 12006)) {
              cleanupPickerStorage();
            }
          });
        }
      );
    } catch (err) {
      await paloUserNotify(err && err.message ? err.message : String(err), "error");
      storageRemove(PICKER_STORAGE_KEY);
      complete(event);
    }
  }

  function pad2(n) {
    return n < 10 ? "0" + n : String(n);
  }

  function sanitizeFileNamePart(name) {
    return String(name || "")
      .replace(/[<>:"/\\|?*\x00-\x1f]/g, "_")
      .replace(/\s+/g, " ")
      .trim();
  }

  function getDocumentBaseName() {
    var url = "";
    try {
      url = Office.context && Office.context.document ? String(Office.context.document.url || "") : "";
    } catch (_e) {
      url = "";
    }
    if (!url) {
      return "Workbook";
    }
    try {
      var decoded = decodeURIComponent(url);
      var parts = decoded.split(/[/\\]/);
      var file = parts.length ? parts[parts.length - 1].split("?")[0] : "";
      if (!file) {
        return "Workbook";
      }
      return sanitizeFileNamePart(file.replace(/\.(xlsx|xlsm|xlsb|xls|csv)$/i, ""));
    } catch (_e2) {
      return "Workbook";
    }
  }

  function buildSnapshotFileName() {
    var d = new Date();
    var dateStr = d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
    return getDocumentBaseName() + "_snapshot_" + dateStr + ".xlsx";
  }

  /** Conversion formules→valeurs autorisee seulement si le nom contient "snapshot" (casse ignoree). */
  function nameAllowsSnapshotConvert(name) {
    return String(name || "").toLowerCase().indexOf("snapshot") !== -1;
  }

  async function getWorkbookDisplayName(workbookId) {
    var name = "";
    await Excel.run(async function (context) {
      var wb = workbookId ? getWorkbookById(context, workbookId) : context.workbook;
      wb.load("name");
      await context.sync();
      name = String(wb.name || "");
    });
    if (name) {
      return name;
    }
    return getDocumentBaseName();
  }

  function assertSnapshotNameOrThrow(name) {
    if (nameAllowsSnapshotConvert(name)) {
      return;
    }
    throw new Error(
      "Conversion refusee : le nom du fichier doit contenir \"snapshot\" "
      + "(ex. MonFichier_snapshot.xlsx). Nom actuel : \"" + String(name || "") + "\". "
      + "Sur Excel Online : Fichier > Enregistrer une copie avec \"snapshot\" dans le nom, puis relancer Snapshot."
    );
  }

  function arrayBufferToBase64(buffer) {
    var bytes = new Uint8Array(buffer);
    var chunk = 8192;
    var binary = "";
    var i;
    for (i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return window.btoa(binary);
  }

  function sliceDataToUint8(data) {
    if (!data) {
      return new Uint8Array(0);
    }
    if (data instanceof ArrayBuffer) {
      return new Uint8Array(data);
    }
    if (typeof ArrayBuffer !== "undefined" && ArrayBuffer.isView(data)) {
      return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
    }
    return new Uint8Array(data);
  }

  function closeDocumentFileAsync(file) {
    return new Promise(function (resolve, reject) {
      if (!file || typeof file.closeAsync !== "function") {
        resolve();
        return;
      }
      file.closeAsync(function (closeResult) {
        if (closeResult.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }
        reject(new Error(
          closeResult.error && closeResult.error.message
            ? closeResult.error.message
            : "closeAsync impossible."
        ));
      });
    });
  }

  /** Export binaire du classeur hote (lecture seule — ne modifie pas le fichier ouvert). */
  function getWorkbookDocumentBase64() {
    return new Promise(function (resolve, reject) {
      if (!Office.context || !Office.context.document || typeof Office.context.document.getFileAsync !== "function") {
        reject(new Error("Export du classeur non disponible sur ce client."));
        return;
      }
      var fileHandle = null;

      function failWithClose(message) {
        if (!fileHandle) {
          reject(new Error(message));
          return;
        }
        closeDocumentFileAsync(fileHandle).then(function () {
          reject(new Error(message));
        }).catch(function () {
          reject(new Error(message));
        });
      }

      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 65536 },
        function (result) {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error(result.error && result.error.message ? result.error.message : "Lecture du fichier impossible."));
            return;
          }
          fileHandle = result.value;
          var sliceCount = fileHandle.sliceCount || 0;
          if (sliceCount === 0) {
            closeDocumentFileAsync(fileHandle).then(function () {
              resolve("");
            }).catch(reject);
            return;
          }

          var parts = new Array(sliceCount);

          function readSlice(index) {
            fileHandle.getSliceAsync(index, function (sliceResult) {
              if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
                failWithClose(
                  sliceResult.error && sliceResult.error.message
                    ? sliceResult.error.message
                    : "Lecture du fragment " + index + " impossible."
                );
                return;
              }
              parts[index] = sliceDataToUint8(sliceResult.value && sliceResult.value.data);
              if (index + 1 < sliceCount) {
                readSlice(index + 1);
                return;
              }
              var total = 0;
              var j;
              for (j = 0; j < sliceCount; j += 1) {
                total += parts[j].length;
              }
              var merged = new Uint8Array(total);
              var offset = 0;
              for (j = 0; j < sliceCount; j += 1) {
                merged.set(parts[j], offset);
                offset += parts[j].length;
              }
              closeDocumentFileAsync(fileHandle).then(function () {
                fileHandle = null;
                try {
                  resolve(arrayBufferToBase64(merged.buffer));
                } catch (encodeErr) {
                  reject(encodeErr);
                }
              }).catch(function () {
                failWithClose("Fermeture du fichier exporte impossible.");
              });
            });
          }

          readSlice(0);
        }
      );
    });
  }

  function getWorkbookById(context, workbookId) {
    return context.application.workbooks.getItem(workbookId);
  }

  function paloConfirmSnapshotDialog(mode, fileName) {
    var dialogMode = mode === "inplace" ? "inplace" : "copy";
    return new Promise(function (resolve) {
      if (!Office.context || !Office.context.ui || typeof Office.context.ui.displayDialogAsync !== "function") {
        resolve(false);
        return;
      }
      var confirmHref;
      try {
        var dialogUrl = new URL("palo-snapshot-confirm.html", window.location.href);
        dialogUrl.searchParams.set("mode", dialogMode);
        dialogUrl.searchParams.set("v", PALO_PLUGIN_VERSION);
        if (fileName) {
          dialogUrl.searchParams.set("file", fileName);
        }
        confirmHref = dialogUrl.href;
      } catch (_urlErr) {
        confirmHref = "https://gpizzetta.github.io/palo-excel-addin/palo-snapshot-confirm.html?mode="
          + dialogMode
          + "&v=" + encodeURIComponent(PALO_PLUGIN_VERSION)
          + (fileName ? "&file=" + encodeURIComponent(fileName) : "");
      }
      Office.context.ui.displayDialogAsync(
        confirmHref,
        { height: 52, width: 40, displayInIframe: true },
        function (asyncResult) {
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            resolve(false);
            return;
          }
          var dlg = asyncResult.value;
          var settled = false;
          function finish(value) {
            if (settled) {
              return;
            }
            settled = true;
            try {
              dlg.close();
            } catch (_close) {
              // ignore
            }
            resolve(value);
          }
          dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
            finish(String(arg && arg.message ? arg.message : "") === "yes");
          });
          dlg.addEventHandler(Office.EventType.DialogEventReceived, function () {
            finish(false);
          });
        }
      );
    });
  }

  async function snapshotSupportsSeparateWorkbook() {
    var supported = false;
    await Excel.run(async function (context) {
      supported = Boolean(
        context.application
        && context.application.workbooks
        && typeof context.application.workbooks.add === "function"
      );
    });
    return supported;
  }

  async function applySnapshotValuesInPlace() {
    await Excel.run(async function (context) {
      if (context.application && typeof context.application.suspendScreenUpdatingUntilNextSync === "function") {
        context.application.suspendScreenUpdatingUntilNextSync();
      }
      await convertAllSheetsToValues(context, context.workbook);
    });
    try {
      await Excel.run(async function (context) {
        if (
          context.workbook
          && context.workbook.application
          && typeof context.workbook.application.calculate === "function"
          && typeof Excel !== "undefined"
          && Excel.CalculationType
        ) {
          context.workbook.application.calculate(Excel.CalculationType.recalculate);
          await context.sync();
        }
      });
    } catch (_recalc) {
      // ignore
    }
  }

  async function convertAllSheetsToValues(context, workbook) {
    var app = context.application;
    var previousCalcMode = null;
    if (app && Excel.CalculationMode) {
      app.load("calculationMode");
      await context.sync();
      previousCalcMode = app.calculationMode;
      app.calculationMode = Excel.CalculationMode.manual;
      await context.sync();
    }
    try {
      var sheets = workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      var copyValuesOnly = Excel.RangeCopyType && Excel.RangeCopyType.values;
      var i;
      for (i = 0; i < sheets.items.length; i += 1) {
        var sheet = sheets.items[i];
        var used = sheet.getUsedRangeOrNullObject();
        used.load("values");
        await context.sync();
        if (!used.isNullObject) {
          if (copyValuesOnly) {
            used.copyFrom(used, Excel.RangeCopyType.values, false, false);
          } else {
            used.values = used.values;
          }
          await context.sync();
        }
      }
    } finally {
      if (app && previousCalcMode != null && Excel.CalculationMode) {
        app.calculationMode = previousCalcMode;
        await context.sync();
      }
    }
  }

  async function removeDefaultBlankSheetAfterInsert(context, workbook) {
    var sheets = workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    if (sheets.items.length <= 1) {
      return;
    }
    var defaultNames = { Sheet1: true, Feuil1: true, Tabelle1: true };
    var i;
    for (i = 0; i < sheets.items.length; i += 1) {
      if (defaultNames[sheets.items[i].name]) {
        sheets.items[i].delete();
        await context.sync();
        return;
      }
    }
  }

  /**
   * Copie dans la meme instance Excel. context.workbook reste le classeur hote :
   * toutes les ecritures ciblent snapshotWbId explicitement.
   */
  async function createSnapshotWorkbookFromBase64(base64) {
    var snapshotWbId = null;
    await Excel.run(async function (context) {
      if (!context.application || !context.application.workbooks
        || typeof context.application.workbooks.add !== "function") {
        throw new Error("Snapshot: creation de classeur indisponible sur ce client.");
      }
      var snapshotWb = context.application.workbooks.add();
      snapshotWb.load("id");
      await context.sync();
      snapshotWbId = snapshotWb.id;
      if (typeof snapshotWb.insertWorksheetsFromBase64 !== "function") {
        snapshotWb.close(Excel.CloseBehavior.skipSave);
        await context.sync();
        throw new Error("Snapshot: ExcelApi insertWorksheetsFromBase64 requis (1.8+).");
      }
      snapshotWb.insertWorksheetsFromBase64(base64);
      await context.sync();
      if (typeof snapshotWb.activate === "function") {
        snapshotWb.activate();
        await context.sync();
      }
      await removeDefaultBlankSheetAfterInsert(context, snapshotWb);
    });
    if (!snapshotWbId) {
      throw new Error("Snapshot: id du classeur copie introuvable.");
    }
    return snapshotWbId;
  }

  async function saveSnapshotWorkbookAs(fileName, snapshotWbId) {
    await Excel.run(async function (context) {
      var wb = getWorkbookById(context, snapshotWbId);
      var baseName = String(fileName || "").replace(/\.xlsx$/i, "");
      wb.name = baseName;
      await context.sync();
      try {
        wb.save(Excel.SaveBehavior.saveAs);
        await context.sync();
        return;
      } catch (_saveAs) {
        // Repli save / prompt.
      }
      try {
        wb.save(Excel.SaveBehavior.save);
        await context.sync();
      } catch (_save) {
        wb.save(Excel.SaveBehavior.prompt);
        await context.sync();
      }
    });
  }

  async function saveSnapshotWorkbookPersist(snapshotWbId) {
    await Excel.run(async function (context) {
      var wb = getWorkbookById(context, snapshotWbId);
      try {
        wb.save(Excel.SaveBehavior.save);
        await context.sync();
      } catch (_save) {
        try {
          wb.save(Excel.SaveBehavior.prompt);
          await context.sync();
        } catch (_prompt) {
          // Autosave peut deja avoir persisté la copie.
        }
      }
    });
  }

  async function closeSnapshotWorkbookWithoutPrompt(snapshotWbId) {
    await Excel.run(async function (context) {
      getWorkbookById(context, snapshotWbId).close(Excel.CloseBehavior.skipSave);
      await context.sync();
    });
  }

  async function paloSnapshotWorkbookValues(event) {
    var fileName = buildSnapshotFileName();
    var snapshotWbId = null;
    try {
      if (typeof Excel === "undefined" || typeof Excel.run !== "function") {
        throw new Error("Excel JavaScript API indisponible.");
      }

      var canCopy = await snapshotSupportsSeparateWorkbook();
      var targetName = canCopy ? fileName : await getWorkbookDisplayName(null);
      var confirmed = await paloConfirmSnapshotDialog(canCopy ? "copy" : "inplace", targetName);
      if (!confirmed) {
        await paloUserNotify("Snapshot annule.", "neutral", "Snapshot");
        complete(event);
        return;
      }

      if (!canCopy) {
        var openName = await getWorkbookDisplayName(null);
        assertSnapshotNameOrThrow(openName);
        await paloUserNotify("Conversion en valeurs du classeur ouvert…", "neutral", "Snapshot");
        await applySnapshotValuesInPlace();
        await paloUserNotify(
          "Formules remplacees par les valeurs affichees dans \"" + openName + "\". "
          + "Le nom contenait bien \"snapshot\".",
          "ok",
          "Snapshot"
        );
        complete(event);
        return;
      }

      await paloUserNotify("Creation du snapshot en cours…", "neutral", "Snapshot");

      var base64 = await getWorkbookDocumentBase64();
      snapshotWbId = await createSnapshotWorkbookFromBase64(base64);

      await saveSnapshotWorkbookAs(fileName, snapshotWbId);

      var copyName = await getWorkbookDisplayName(snapshotWbId);
      if (!nameAllowsSnapshotConvert(copyName) && !nameAllowsSnapshotConvert(fileName)) {
        throw new Error(
          "La copie n'a pas \"snapshot\" dans le nom (\"" + copyName + "\" / \"" + fileName + "\")."
        );
      }
      assertSnapshotNameOrThrow(nameAllowsSnapshotConvert(copyName) ? copyName : fileName);

      await Excel.run(async function (context) {
        if (context.application && typeof context.application.suspendScreenUpdatingUntilNextSync === "function") {
          context.application.suspendScreenUpdatingUntilNextSync();
        }
        await convertAllSheetsToValues(context, getWorkbookById(context, snapshotWbId));
      });

      await saveSnapshotWorkbookPersist(snapshotWbId);

      try {
        await closeSnapshotWorkbookWithoutPrompt(snapshotWbId);
        snapshotWbId = null;
      } catch (_close) {
        // La copie peut rester ouverte.
      }

      await paloUserNotify(
        "Snapshot enregistre : " + fileName + ". "
        + "Copie en valeurs (sans formules). "
        + "Le classeur d'origine avec formules n'a pas ete modifie.",
        "ok",
        "Snapshot"
      );
    } catch (err) {
      if (snapshotWbId) {
        try {
          await closeSnapshotWorkbookWithoutPrompt(snapshotWbId);
        } catch (_cleanup) {
          // ignore
        }
      }
      await paloUserNotify(
        "Snapshot impossible : " + (err && err.message ? err.message : String(err))
        + " — nom prevu : " + fileName,
        "error",
        "Snapshot"
      );
    }
    complete(event);
  }

  window.paloSnapshotWorkbookValues = paloSnapshotWorkbookValues;

  Office.onReady(function () {
    if (Office.actions && typeof Office.actions.associate === "function") {
      Office.actions.associate("openTaskpane", openTaskpane);
      Office.actions.associate("testConnection", testConnection);
      Office.actions.associate("insertPaloDataFormula", insertPaloDataFormula);
      Office.actions.associate("insertPaloSetDataFormula", insertPaloSetDataFormula);
      Office.actions.associate("paloRibbonAction", paloRibbonAction);
      Office.actions.associate("paloSnapshotWorkbookValues", paloSnapshotWorkbookValues);
    }
  });
})();
