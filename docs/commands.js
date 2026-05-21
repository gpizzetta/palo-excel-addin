/* global Office, Excel, OfficeRuntime */
(function commandsBootstrap() {
  var PICKER_STORAGE_KEY = "palo_ename_picker_v1";

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
        function () {
          resolve();
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

          // Tentative 1: cellule active.
          var fromActive = buildResultFromFormula(readFormula(cell));
          // Tentative 2: cellule en haut-gauche de la selection.
          var picked = fromActive.skip ? buildResultFromFormula(readFormula(sel)) : fromActive;
          if (picked.skip) {
            return picked;
          }

          await context.sync();
          return {
            skip: false,
            servdb: window.paloEnameRibbon.readPrepared(picked.p1).trim(),
            dimension: window.paloEnameRibbon.readPrepared(picked.p2).trim()
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

      if (pickContext.skip) {
        await paloUserNotify("Palo action fonctionne sur une cellule contenant PALO.ENAME(...).", "error");
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
      dialogUrl.searchParams.set("v", "1.0.2.1");

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
    return String(name || "Workbook")
      .replace(/[\\/:*?"<>|]/g, "_")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 180) || "Workbook";
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

  function getWorkbookDocumentBase64() {
    return new Promise(function (resolve, reject) {
      if (!Office.context || !Office.context.document || typeof Office.context.document.getFileAsync !== "function") {
        reject(new Error("Export du classeur non disponible sur ce client."));
        return;
      }
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 4194304 },
        function (result) {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error(result.error && result.error.message ? result.error.message : "Lecture du fichier impossible."));
            return;
          }
          var file = result.value;
          var sliceCount = file.sliceCount;
          var slices = [];
          var received = 0;

          function onSliceError(err) {
            reject(new Error(err && err.message ? err.message : "Lecture d'un fragment impossible."));
          }

          function getSlice(index) {
            file.getSliceAsync(index, function (sliceResult) {
              if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
                onSliceError(sliceResult.error);
                return;
              }
              slices[index] = sliceResult.value.data;
              received += 1;
              if (received >= sliceCount) {
                var total = 0;
                var j;
                for (j = 0; j < sliceCount; j += 1) {
                  total += slices[j].byteLength;
                }
                var merged = new Uint8Array(total);
                var offset = 0;
                for (j = 0; j < sliceCount; j += 1) {
                  merged.set(new Uint8Array(slices[j]), offset);
                  offset += slices[j].byteLength;
                }
                resolve(arrayBufferToBase64(merged.buffer));
                return;
              }
              getSlice(received);
            });
          }

          getSlice(0);
        }
      );
    });
  }

  async function convertAllSheetsToValues(context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    var i;
    for (i = 0; i < sheets.items.length; i += 1) {
      var sheet = sheets.items[i];
      var used = sheet.getUsedRangeOrNullObject();
      used.load("values");
      await context.sync();
      if (!used.isNullObject) {
        used.values = used.values;
        await context.sync();
      }
    }
  }

  async function saveSnapshotWorkbook(fileName) {
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var baseName = String(fileName || "").replace(/\.xlsx$/i, "");
      wb.name = baseName;
      try {
        wb.save(Excel.SaveBehavior.save);
        await context.sync();
      } catch (_save) {
        wb.save(Excel.SaveBehavior.prompt);
        await context.sync();
      }
    });
  }

  async function closeActiveWorkbookWithoutPrompt() {
    await Excel.run(async function (context) {
      context.workbook.close(Excel.CloseBehavior.skipSave);
      await context.sync();
    });
  }

  async function paloSnapshotWorkbookValues(event) {
    var fileName = buildSnapshotFileName();
    try {
      if (typeof Excel === "undefined" || typeof Excel.createWorkbook !== "function") {
        throw new Error("Snapshot non disponible (ExcelApi 1.8 requis).");
      }
      if (typeof Excel.run !== "function") {
        throw new Error("Excel JavaScript API indisponible.");
      }

      var base64 = await getWorkbookDocumentBase64();
      await Excel.createWorkbook(base64);

      await Excel.run(async function (context) {
        if (context.application && typeof context.application.suspendScreenUpdatingUntilNextSync === "function") {
          context.application.suspendScreenUpdatingUntilNextSync();
        }
        await convertAllSheetsToValues(context);
      });

      await saveSnapshotWorkbook(fileName);

      try {
        await closeActiveWorkbookWithoutPrompt();
      } catch (_close) {
        // Le classeur snapshot peut rester ouvert si Excel refuse la fermeture.
      }

      await paloUserNotify(
        "Snapshot enregistre : " + fileName + ". "
        + "Tous les onglets en valeurs (sans formules). "
        + "Le classeur d'origine avec formules Palo est reste ouvert.",
        "ok",
        "Snapshot"
      );
    } catch (err) {
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
