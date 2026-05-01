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
		return base + "action-popup.html?v=1.0.49.0&" + q.toString();
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
				return {
					address: cell.address || "",
					formula: formula,
					value: value,
					functionName: fn,
				};
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
