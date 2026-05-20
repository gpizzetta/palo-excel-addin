/* global Office, Excel, OfficeRuntime */
(function commandsBootstrap() {
  var PICKER_STORAGE_KEY = "palo_ename_picker_v1";

  function complete(event) {
    if (event && typeof event.completed === "function") {
      event.completed();
    }
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
        window.alert("Aucune connexion active. Ouvrez le volet Palo et selectionnez une connexion dans la liste.");
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
        window.alert("Aucune connexion active. Ouvrez le volet Palo et selectionnez une connexion dans la liste.");
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
        window.alert(
          "Palo action indisponible : script ruban incomplet (palo-ename-ribbon ou connexion Palo). Rechargez le volet ou reimportez le manifeste."
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
        window.alert("Palo action fonctionne sur une cellule contenant PALO.ENAME(...).");
        complete(event);
        return;
      }

      if (!pickContext.servdb || !pickContext.dimension) {
        window.alert("PALO.ENAME: servdb ou dimension vide apres resolution des references.");
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
        window.alert("Aucune dimension pour la base courante.");
        complete(event);
        return;
      }

      var els = await ctx.client.dimensionElements(ctx.sid, ctx.database, pickContext.dimension);
      var items = mapDimensionElementsToPickerItems(els);

      if (!items.length) {
        window.alert("Aucun element pour la dimension \"" + pickContext.dimension + "\".");
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
      dialogUrl.searchParams.set("v", "1.0.1.122");

      Office.context.ui.displayDialogAsync(
        dialogUrl.href,
        { height: 62, width: 32, displayInIframe: true },
        function (asyncResult) {
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            var errDetail = asyncResult.errorMessage || "";
            var errCode = asyncResult.errorCode != null ? String(asyncResult.errorCode) : "";
            window.alert(
              "Impossible d'ouvrir la fenetre de selection."
                + (errCode ? " Code: " + errCode + "." : "")
                + (errDetail ? " " + errDetail : "")
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
                window.alert("Cellule PALO.ENAME introuvable pour appliquer le choix.");
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
              window.alert(err && err.message ? err.message : String(err));
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
      window.alert(err && err.message ? err.message : String(err));
      storageRemove(PICKER_STORAGE_KEY);
      complete(event);
    }
  }

  Office.onReady(function () {
    if (Office.actions && typeof Office.actions.associate === "function") {
      Office.actions.associate("openTaskpane", openTaskpane);
      Office.actions.associate("testConnection", testConnection);
      Office.actions.associate("insertPaloDataFormula", insertPaloDataFormula);
      Office.actions.associate("insertPaloSetDataFormula", insertPaloSetDataFormula);
      Office.actions.associate("paloRibbonAction", paloRibbonAction);
    }
  });
})();
