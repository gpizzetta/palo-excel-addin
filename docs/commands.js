/* global Office, Excel */

(function () {
	function escapeHtml(s) {
		return String(s == null ? "" : s)
			.replace(/&/g, "&amp;")
			.replace(/</g, "&lt;")
			.replace(/>/g, "&gt;")
			.replace(/"/g, "&quot;")
			.replace(/'/g, "&#39;");
	}

	function detectFunctionName(formula) {
		var f = String(formula == null ? "" : formula).trim();
		if (!f) {
			return "";
		}
		var m = f.match(/^=\s*([A-Za-z_][A-Za-z0-9_.]*)\s*\(/);
		return m ? String(m[1]).toUpperCase() : "";
	}

	function isEnameRibbonFunction(name) {
		var u = String(name || "").toUpperCase();
		return u === "PALO.ENAME" || u === "ENAME" || /\.ENAME$/i.test(String(name || ""));
	}

	function buildActionDialogUrl(payload) {
		var base = window.location.href.split("#")[0];
		base = base.slice(0, base.lastIndexOf("/") + 1);
		var q = new URLSearchParams({
			func: payload.functionName || "",
			formula: payload.formula || "",
			value: payload.value || "",
			address: payload.address || "",
		});
		if (payload.enameDb) {
			q.set("ename_db", payload.enameDb);
		}
		if (payload.enameDim) {
			q.set("ename_dim", payload.enameDim);
		}
		if (payload.enameEl) {
			q.set("ename_el", payload.enameEl);
		}
		return base + "action-popup.html?v=1.0.54.0&" + q.toString();
	}

	function isLocalA1Notation(loc) {
		return /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$/.test(String(loc || "").trim());
	}

	/**
	 * Référence une cellule : même classeur — A1, Feuille!B2, 'Ma feuille'!C3.
	 * Autre classeur — [fichier.xlsx]Feuil!A1 : détecté mais non résolu ici (voir message dans le popup).
	 */
	function parseSingleCellReference(expr) {
		var t = String(expr || "").trim();
		if (!t) {
			return null;
		}
		if (t.indexOf(":") >= 0) {
			return null;
		}
		/** Excel : [Classeur.xlsx]Feuille!$A$1 — le moteur de calcul Excel y accède ; Office.js ne lit que le classeur hôte. */
		var extM = t.match(/^\[([^\]]+)\]([^!]+)!(.+)$/);
		if (extM) {
			var locExt = extM[3].trim();
			if (!isLocalA1Notation(locExt)) {
				return null;
			}
			return { external: true };
		}
		if (/[\[\]#@]/.test(t)) {
			return null;
		}
		if (/^'/.test(t)) {
			var m = t.match(/^'((?:''|[^'])*)'!(.+)$/);
			if (!m) {
				return null;
			}
			var sheet = m[1].replace(/''/g, "'");
			var loc = m[2].trim();
			if (!isLocalA1Notation(loc)) {
				return null;
			}
			return { sheet: sheet, local: loc };
		}
		var bang = t.indexOf("!");
		if (bang >= 0) {
			var sheetPart = t.slice(0, bang).trim();
			var loc = t.slice(bang + 1).trim();
			if (!sheetPart || !isLocalA1Notation(loc)) {
				return null;
			}
			return { sheet: sheetPart, local: loc };
		}
		if (!isLocalA1Notation(t)) {
			return null;
		}
		return { sheet: null, local: t };
	}


	/**
	 * Résout base / dimension / élément : chaînes littérales ou valeurs des cellules référencées
	 * (équivalent pratique à « remonter à la source » pour des références simples).
	 */
	function resolvePaloEnameArgsFromFormula(ctx, cell, formula) {
		if (typeof parsePaloEnameFirstThreeArgExpressions !== "function") {
			return Promise.resolve(null);
		}
		var exprs = parsePaloEnameFirstThreeArgExpressions(formula);
		if (!exprs) {
			return Promise.resolve(null);
		}
		var resolved = [null, null, null];
		var pending = [];
		for (var i = 0; i < 3; i++) {
			var lit =
				typeof tryExcelFormulaStringLiteral === "function"
					? tryExcelFormulaStringLiteral(exprs[i])
					: null;
			if (lit !== null) {
				resolved[i] = lit;
			} else {
				pending.push({ i: i, expr: exprs[i] });
			}
		}
		if (!pending.length) {
			return resolved[0] && resolved[1] && resolved[2]
				? Promise.resolve(resolved)
				: Promise.resolve(null);
		}
		for (var p = 0; p < pending.length; p++) {
			var addr = parseSingleCellReference(pending[p].expr);
			if (!addr) {
				return Promise.resolve(null);
			}
			if (addr.external) {
				return Promise.resolve(null);
			}
			pending[p]._addr = addr;
			if (addr.sheet) {
				if (typeof ctx.workbook.worksheets.getItemOrNullObject === "function") {
					pending[p]._ws = ctx.workbook.worksheets.getItemOrNullObject(addr.sheet);
					pending[p]._ws.load("isNullObject");
				} else {
					pending[p]._ws = ctx.workbook.worksheets.getItem(addr.sheet);
					pending[p]._wsReady = true;
				}
			} else {
				pending[p]._ws = cell.worksheet;
				pending[p]._wsReady = true;
			}
		}
		return ctx.sync().then(function () {
			for (var pi = 0; pi < pending.length; pi++) {
				var pobj = pending[pi];
				var ws = pobj._ws;
				if (!pobj._wsReady && ws.isNullObject) {
					return null;
				}
				var rng = ws.getRange(pobj._addr.local);
				rng.load("values");
				pobj.rng = rng;
			}
			return ctx.sync().then(function () {
				for (var q = 0; q < pending.length; q++) {
					var it = pending[q];
					var v =
						it.rng.values &&
						it.rng.values[0] &&
						it.rng.values[0][0] !== undefined &&
						it.rng.values[0][0] !== null
							? String(it.rng.values[0][0]).trim()
							: "";
					resolved[it.i] = v;
				}
				return resolved[0] && resolved[1] && resolved[2] ? resolved : null;
			});
		});
	}

	function applyFormulaToAddress(address, formula) {
		return Excel.run(function (ctx) {
			var parts = String(address || "").trim().lastIndexOf("!");
			if (parts < 0) {
				var cell = ctx.workbook.getActiveCell();
				cell.formulas = [[formula]];
				return ctx.sync();
			}
			var sheetRaw = String(address || "").trim().slice(0, parts);
			var rangeA1 = String(address || "").trim().slice(parts + 1);
			var sheetName = sheetRaw.replace(/^'|'$/g, "").replace(/''/g, "'");
			var sheet = ctx.workbook.worksheets.getItem(sheetName);
			var rng = sheet.getRange(rangeA1);
			rng.formulas = [[formula]];
			return ctx.sync();
		});
	}

	function openActionDialog(payload, event) {
		var dialogUrl = buildActionDialogUrl(payload);
		var size = isEnameRibbonFunction(payload.functionName)
			? { height: 55, width: 38, displayInIframe: true }
			: { height: 40, width: 40, displayInIframe: true };
		Office.context.ui.displayDialogAsync(dialogUrl, size, function (asyncResult) {
			if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
				var dialog = asyncResult.value;
				dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
					var raw =
						arg && typeof arg.message === "string"
							? arg.message
							: typeof arg === "string"
								? arg
								: "";
					if (raw === "close") {
						try {
							dialog.close();
						} catch (e) {}
						return;
					}
					try {
						var o = JSON.parse(raw);
						if (o && o.action === "updateFormula" && o.formula && o.address) {
							applyFormulaToAddress(o.address, o.formula)
								.then(function () {
									try {
										dialog.close();
									} catch (e2) {}
								})
								.catch(function () {
									try {
										dialog.close();
									} catch (e3) {}
								});
							return;
						}
					} catch (e) {
						/* ignore */
					}
					try {
						dialog.close();
					} catch (e4) {}
				});
			}
			event.completed();
		});
	}

	function showActionPopup(event) {
		Excel.run(function (ctx) {
			var cell = ctx.workbook.getActiveCell();
			cell.load(["address", "formulas", "values"]);
			cell.worksheet.load("name");
			return ctx.sync().then(function () {
				var formula = "";
				if (cell.formulas && cell.formulas[0] && cell.formulas[0][0] != null) {
					formula = String(cell.formulas[0][0]);
				}
				var value = "";
				if (cell.values && cell.values[0] && cell.values[0][0] != null) {
					value = String(cell.values[0][0]);
				}
				var fn = detectFunctionName(formula);
				var payload = {
					address: cell.address || "",
					formula: formula,
					value: value,
					functionName: fn,
				};
				if (!isEnameRibbonFunction(fn) || !formula) {
					return payload;
				}
				return resolvePaloEnameArgsFromFormula(ctx, cell, formula)
					.then(function (triple) {
						if (triple && triple[0] && triple[1] && triple[2]) {
							payload.enameDb = triple[0];
							payload.enameDim = triple[1];
							payload.enameEl = triple[2];
						}
						return payload;
					})
					.catch(function () {
						return payload;
					});
			});
		})
			.catch(function (err) {
				return {
					address: "",
					formula: "",
					value: "",
					functionName: "",
					error: err && err.message ? err.message : String(err),
				};
			})
			.then(function (payload) {
				if (payload.error) {
					payload.formula = "Erreur lecture cellule: " + escapeHtml(payload.error);
				}
				openActionDialog(payload, event);
			});
	}

	Office.onReady(function () {
		Office.actions.associate("showActionPopup", showActionPopup);
	});
})();
