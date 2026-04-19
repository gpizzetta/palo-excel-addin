(function () {
	var ctx = {
		apiBase: "",
		sid: "",
		name_database: "",
		name_cube: "",
		mode: "create",
		rule: "",
	};

	function parseQuery() {
		var p = new URLSearchParams(window.location.search);
		ctx.apiBase = (p.get("apiBase") || "").trim();
		ctx.sid = p.get("sid") || "";
		ctx.name_database = p.get("name_database") || "";
		ctx.name_cube = p.get("name_cube") || "";
		ctx.mode = (p.get("mode") || "create").trim();
		ctx.rule = (p.get("rule") || "").trim();
	}

	function showErr(msg) {
		var el = document.getElementById("err");
		if (el) {
			el.textContent = msg || "Erreur";
			el.className = "show";
		}
	}

	function stripBom(text) {
		return String(text).replace(/^\uFEFF/, "");
	}

	function stripPaloCsvField(s) {
		var t = String(s).trim();
		if (t.length >= 2 && t.charAt(0) === '"' && t.charAt(t.length - 1) === '"') {
			return t.slice(1, -1).replace(/""/g, '"');
		}
		return t;
	}

	function isNumericId(s) {
		return s !== undefined && s !== null && /^\d+$/.test(String(s).trim());
	}

	function parseCsvLines(text) {
		return stripBom(text)
			.split(/\r?\n/)
			.map(function (line) {
				return line.replace(/\s+$/, "");
			})
			.filter(function (line) {
				return line.length;
			});
	}

	function splitPaloRuleLine(line) {
		var parts = [];
		var i = 0;
		var cur = "";
		var inQuote = false;
		while (i < line.length) {
			var c = line.charAt(i);
			if (c === '"') {
				if (inQuote && line.charAt(i + 1) === '"') {
					cur += '"';
					i += 2;
					continue;
				}
				inQuote = !inQuote;
				cur += c;
				i++;
				continue;
			}
			if (c === ";" && !inQuote) {
				parts.push(cur.trim());
				cur = "";
				i++;
				continue;
			}
			cur += c;
			i++;
		}
		parts.push(cur.trim());
		return parts;
	}

	function findRuleInCsv(text, ruleId) {
		var lines = parseCsvLines(text);
		for (var i = 0; i < lines.length; i++) {
			var cells = splitPaloRuleLine(lines[i]);
			if (!cells.length || !isNumericId(cells[0])) {
				continue;
			}
			if (String(cells[0]).trim() !== String(ruleId).trim()) {
				continue;
			}
			var def = cells.length > 1 ? stripPaloCsvField(cells[1]) : "";
			var comment = cells.length > 3 ? stripPaloCsvField(cells[3]) : "";
			return { definition: def, comment: comment };
		}
		return null;
	}

	function loadExistingRule() {
		var q = new URLSearchParams({
			sid: ctx.sid,
			name_database: ctx.name_database,
			name_cube: ctx.name_cube,
		});
		var url = ctx.apiBase + "/cube/rules?" + q.toString();
		return fetch(url, {
			method: "GET",
			mode: "cors",
			cache: "no-store",
			credentials: "omit",
		}).then(function (res) {
			return res.text().then(function (text) {
				if (!res.ok) {
					throw new Error("HTTP " + res.status + " — " + text.slice(0, 400));
				}
				var found = findRuleInCsv(text, ctx.rule);
				if (!found) {
					throw new Error("Règle id " + ctx.rule + " introuvable dans /cube/rules.");
				}
				return found;
			});
		});
	}

	function save() {
		parseQuery();
		var def = document.getElementById("areaDef").value;
		var comment = document.getElementById("inputComment").value.trim();
		if (!def.trim()) {
			showErr("Indiquez une définition de règle.");
			return;
		}
		var btn = document.getElementById("btnOk");
		btn.disabled = true;

		function finish() {
			if (Office && Office.context && Office.context.ui && Office.context.ui.messageParent) {
				Office.context.ui.messageParent("refresh", { targetOrigin: "*" });
			}
		}

		function fail(err) {
			showErr(err && err.message ? err.message : String(err));
			btn.disabled = false;
		}

		if (ctx.mode === "edit") {
			var qe = new URLSearchParams({
				sid: ctx.sid,
				name_database: ctx.name_database,
				name_cube: ctx.name_cube,
				rule: ctx.rule,
				definition: def,
			});
			if (comment) {
				qe.set("comment", comment);
			}
			var urlE = ctx.apiBase + "/rule/modify?" + qe.toString();
			fetch(urlE, { method: "GET", mode: "cors", cache: "no-store", credentials: "omit" })
				.then(function (res) {
					return res.text().then(function (text) {
						if (!res.ok) {
							throw new Error("HTTP " + res.status + " — " + text.slice(0, 500));
						}
						finish();
					});
				})
				.catch(fail);
			return;
		}

		var qc = new URLSearchParams({
			sid: ctx.sid,
			name_database: ctx.name_database,
			name_cube: ctx.name_cube,
			definition: def,
			activate: "1",
		});
		if (comment) {
			qc.set("comment", comment);
		}
		var urlC = ctx.apiBase + "/rule/create?" + qc.toString();
		fetch(urlC, { method: "GET", mode: "cors", cache: "no-store", credentials: "omit" })
			.then(function (res) {
				return res.text().then(function (text) {
					if (!res.ok) {
						throw new Error("HTTP " + res.status + " — " + text.slice(0, 500));
					}
					finish();
				});
			})
			.catch(fail);
	}

	function cancel() {
		if (Office && Office.context && Office.context.ui && Office.context.ui.messageParent) {
			Office.context.ui.messageParent("cancel", { targetOrigin: "*" });
		}
	}

	Office.onReady(function () {
		parseQuery();
		document.getElementById("dlgSub").textContent =
			ctx.name_database + " / " + ctx.name_cube + (ctx.mode === "edit" ? " — règle " + ctx.rule : "");
		document.getElementById("dlgTitle").textContent =
			ctx.mode === "edit" ? "Modifier la règle" : "Nouvelle règle";
		document.getElementById("btnCancel").addEventListener("click", cancel);
		document.getElementById("btnOk").addEventListener("click", save);
		if (ctx.mode === "edit" && ctx.rule) {
			document.getElementById("btnOk").disabled = true;
			loadExistingRule()
				.then(function (r) {
					document.getElementById("areaDef").value = r.definition || "";
					document.getElementById("inputComment").value = r.comment || "";
					document.getElementById("btnOk").disabled = false;
				})
				.catch(function (err) {
					showErr(err && err.message ? err.message : String(err));
					document.getElementById("btnOk").disabled = false;
				});
		} else {
			document.getElementById("btnOk").disabled = false;
			document.getElementById("areaDef").focus();
		}
	});
})();
