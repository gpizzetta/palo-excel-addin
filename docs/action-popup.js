/* global Office */
/* Popup « Action » : PALO.ENAME (liste d’éléments) ; PALO.DATAC (LIKE/COPY → /cell/replace, formule inchangée). */
(function () {
	var KEYS = {
		url: "palo_connection_url",
		username: "palo_connection_username",
		password: "palo_connection_password",
	};

	var sessionCache = {
		apiBase: "",
		sid: "",
		fp: "",
		at: 0,
	};
	var SESSION_TTL_MS = 4 * 60 * 1000;

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

	function splitPaloCsvLine(line) {
		return line.indexOf(";") >= 0 ? line.split(";") : line.split(",");
	}

	function looksLikePaloErrorDetail(s) {
		var t = String(s || "")
			.toLowerCase()
			.replace(/\s+/g, " ")
			.trim();
		if (!t) {
			return false;
		}
		return (
			t.indexOf("erreur") !== -1 ||
			t.indexOf("error") !== -1 ||
			t.indexOf("invalid") !== -1 ||
			t.indexOf("wrong") !== -1 ||
			t.indexOf("missing") !== -1 ||
			t.indexOf("failed") !== -1 ||
			t.indexOf("internal") !== -1 ||
			t.indexOf("permission") !== -1 ||
			t.indexOf("denied") !== -1
		);
	}

	function shortResponseExcerpt(text) {
		var t = stripBom(String(text || "")).replace(/\s+/g, " ").trim();
		if (!t) {
			return "(vide)";
		}
		if (t.length > 220) {
			return t.slice(0, 220) + "...";
		}
		return t;
	}

	function excerptCellApiBody(text) {
		var t = stripBom(String(text || "")).replace(/\s+/g, " ").trim();
		if (!t) {
			return "(vide)";
		}
		var max = 1600;
		if (t.length > max) {
			return t.slice(0, max) + "...";
		}
		return t;
	}

	function redactPaloSidInUrl(url) {
		return String(url || "").replace(/sid=[^&]*/gi, "sid=***");
	}

	function parsePaloStatus(text) {
		var raw = stripBom(text);
		var lines = raw
			.split(/\r?\n/)
			.map(function (line) {
				return line.replace(/\s+$/, "");
			})
			.filter(function (line) {
				return line.length;
			});
		if (!lines.length) {
			return;
		}
		if (lines[0].charAt(0) === "<" || lines[0].toLowerCase().indexOf("<!doctype") !== -1) {
			throw new Error(shortResponseExcerpt(lines[0]));
		}
		for (var li = 0; li < lines.length; li++) {
			var line = lines[li];
			var cells = splitPaloCsvLine(line);
			var c0 = cells.length ? stripPaloCsvField(cells[0]).trim() : "";
			if (!/^[0-9]{1,10}$/.test(c0)) {
				continue;
			}
			var code = parseInt(c0, 10);
			if (code === 0) {
				continue;
			}
			var c1 = cells.length > 1 ? stripPaloCsvField(String(cells[1] || "").trim()) : "";
			if (code < 100) {
				if (c1 === "ok" || c1 === "1" || c1 === "true" || c1 === "0") {
					continue;
				}
				if (!looksLikePaloErrorDetail(c1) && !looksLikePaloErrorDetail(line)) {
					continue;
				}
			}
			var serverLine = line.length > 800 ? line.slice(0, 800) + "..." : line;
			throw new Error(serverLine);
		}
		var low = raw.toLowerCase();
		if (low.indexOf("erreur interne") !== -1 || low.indexOf("internal error") !== -1) {
			throw new Error(shortResponseExcerpt(raw));
		}
	}

	function fetchCellReplaceSplash(apiBase, sid, database, cube, pathArr, valueText, splashMode) {
		var namePath = (pathArr || []).join(",");
		var valueStr = String(valueText == null ? "" : valueText);
		var sm = Math.round(Number(splashMode));
		if (isNaN(sm) || sm < 0 || sm > 5) {
			sm = 1;
		}
		var q = new URLSearchParams({
			sid: sid,
			name_database: database,
			name_cube: cube,
			name_path: namePath,
			value: valueStr,
			splash: String(sm),
		});
		var url = String(apiBase).replace(/\/$/, "") + "/cell/replace?" + q.toString();
		return fetch(url, {
			method: "GET",
			mode: "cors",
			cache: "no-store",
			credentials: "omit",
		}).then(function (res) {
			return res.text().then(function (text) {
				if (!res.ok) {
					throw new Error(
						"HTTP " + res.status + " — " + (excerptCellApiBody(text) || redactPaloSidInUrl(url)),
					);
				}
				parsePaloStatus(text);
			});
		});
	}

	/** Base, cube et chemin en littéraux "…" uniquement (requis pour appeler /cell/replace depuis le popup). */
	function parseDatacLiteralPathForReplace(formula) {
		if (typeof parsePaloDatacAllArgExpressions !== "function" || typeof tryExcelFormulaStringLiteral !== "function") {
			return { error: "Analyseur de formule indisponible." };
		}
		var args = parsePaloDatacAllArgExpressions(formula);
		if (!args || args.length < 3) {
			return { error: "Formule PALO.DATAC : au moins base, cube et un élément sont requis." };
		}
		var db = tryExcelFormulaStringLiteral(args[0]);
		var cube = tryExcelFormulaStringLiteral(args[1]);
		if (db == null || cube == null) {
			return {
				error:
					"Pour le splash depuis ce popup, la base et le cube doivent être des chaînes littérales entre guillemets dans la formule.",
			};
		}
		var path = [];
		for (var i = 2; i < args.length; i++) {
			var el = tryExcelFormulaStringLiteral(args[i]);
			if (el == null) {
				return {
					error:
						"Chaque élément du chemin doit être un littéral \"…\" (argument " +
						(i - 1) +
						"). Les références de cellules ne sont pas résolues ici pour l’écriture.",
				};
			}
			path.push(el);
		}
		return { database: db, cube: cube, path: path };
	}

	function readSettingsFromLocalStorage() {
		try {
			if (typeof localStorage === "undefined") {
				return { url: "", username: "", password: "" };
			}
			return {
				url: (localStorage.getItem(KEYS.url) || "").trim(),
				username: (localStorage.getItem(KEYS.username) || "").trim(),
				password: localStorage.getItem(KEYS.password) || "",
			};
		} catch (e) {
			return { url: "", username: "", password: "" };
		}
	}

	function mergePaloConnectionFromLocal(cfg) {
		var loc = readSettingsFromLocalStorage();
		var docUrl = cfg.url != null ? String(cfg.url).trim() : "";
		var docUser = cfg.username != null ? String(cfg.username).trim() : "";
		var docPass = cfg.password != null ? String(cfg.password) : "";
		return {
			url: docUrl || loc.url || "",
			username: docUser || loc.username || "",
			password: docPass || loc.password || "",
		};
	}

	function loadSettingsAsync() {
		return new Promise(function (resolve) {
			var s = Office.context && Office.context.document && Office.context.document.settings;
			function readFromDocument() {
				resolve({
					url: (s.get(KEYS.url) || "").trim(),
					username: (s.get(KEYS.username) || "").trim(),
					password: s.get(KEYS.password) || "",
				});
			}
			if (!s) {
				resolve(readSettingsFromLocalStorage());
				return;
			}
			if (typeof s.refreshAsync === "function") {
				s.refreshAsync(function (ar) {
					if (ar.status !== Office.AsyncResultStatus.Succeeded) {
						resolve(readSettingsFromLocalStorage());
						return;
					}
					readFromDocument();
				});
			} else {
				try {
					readFromDocument();
				} catch (e) {
					resolve(readSettingsFromLocalStorage());
				}
			}
		}).then(function (cfg) {
			return mergePaloConnectionFromLocal(cfg);
		});
	}

	function apiBaseCandidates(connectionUrl) {
		var u = new URL(String(connectionUrl).trim());
		return [u.origin];
	}

	function parseLoginSidFromText(text) {
		var lines = stripBom(text)
			.split(/\r?\n/)
			.map(function (line) {
				return line.replace(/\s+$/, "");
			})
			.filter(function (line) {
				return line.length;
			});
		if (!lines.length) {
			throw new Error("Réponse login vide.");
		}
		var cells = lines[0].split(";");
		var sid = cells[0] ? cells[0].trim() : "";
		if (!sid) {
			throw new Error("Identifiant de session manquant.");
		}
		if (/^[0-9]{1,5}$/.test(sid) && cells.length > 1 && cells[1]) {
			var code = parseInt(sid, 10);
			if (code > 0) {
				throw new Error(cells.slice(1).join("; "));
			}
		}
		return sid;
	}

	function loginAtBase(apiBase, user, password) {
		if (typeof md5 !== "function") {
			return Promise.reject(new Error("Bibliothèque MD5 indisponible (md5.js)."));
		}
		var q = new URLSearchParams({
			user: user,
			password: md5(String(password)),
		});
		var url = apiBase + "/server/login?" + q.toString();
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
				return parseLoginSidFromText(text);
			});
		});
	}

	function discoverAndLogin(connectionUrl, user, password) {
		var bases = apiBaseCandidates(connectionUrl);
		function tryAt(i) {
			if (i >= bases.length) {
				return Promise.reject(new Error("Impossible de joindre l’API Palo (URL ou CORS)."));
			}
			var apiBase = bases[i];
			return loginAtBase(apiBase, user, password)
				.then(function (sid) {
					return { apiBase: apiBase, sid: sid };
				})
				.catch(function (err) {
					var msg = err && err.message ? err.message : String(err);
					var retry =
						msg.indexOf("HTTP 404") !== -1 ||
						msg.indexOf("HTTP 405") !== -1 ||
						msg.indexOf("Failed to fetch") !== -1;
					if (retry && i + 1 < bases.length) {
						return tryAt(i + 1);
					}
					throw err;
				});
		}
		return tryAt(0);
	}

	function sessionFingerprint(cfg) {
		return md5(String(cfg.url) + "\t" + String(cfg.username) + "\t" + String(cfg.password));
	}

	function getCachedSession(cfg) {
		if (typeof md5 !== "function") {
			return discoverAndLogin(cfg.url, cfg.username, cfg.password);
		}
		var fp = sessionFingerprint(cfg);
		if (
			sessionCache.sid &&
			sessionCache.apiBase &&
			Date.now() - sessionCache.at < SESSION_TTL_MS &&
			sessionCache.fp === fp
		) {
			return Promise.resolve({ apiBase: sessionCache.apiBase, sid: sessionCache.sid });
		}
		return discoverAndLogin(cfg.url, cfg.username, cfg.password).then(function (sess) {
			sessionCache = {
				apiBase: sess.apiBase,
				sid: sess.sid,
				fp: fp,
				at: Date.now(),
			};
			return sess;
		});
	}

	function buildDimensionElementsRequestUrl(apiBase, sid, nameDatabase, nameDimension) {
		var q = new URLSearchParams({
			sid: sid,
			name_database: nameDatabase,
			name_dimension: nameDimension,
			show_permission: "0",
		});
		return String(apiBase).replace(/\/$/, "") + "/dimension/elements?" + q.toString();
	}

	/** Littéraux uniquement (sans résolution Excel côté commands). */
	function parseEnameLiteralArgs(formula) {
		if (typeof parsePaloEnameFirstThreeArgExpressions !== "function") {
			return null;
		}
		if (typeof tryExcelFormulaStringLiteral !== "function") {
			return null;
		}
		var exprs = parsePaloEnameFirstThreeArgExpressions(formula);
		if (!exprs) {
			return null;
		}
		var db = tryExcelFormulaStringLiteral(exprs[0]);
		var dim = tryExcelFormulaStringLiteral(exprs[1]);
		var el = tryExcelFormulaStringLiteral(exprs[2]);
		if (db == null || dim == null || el == null) {
			return null;
		}
		return { database: db, dimension: dim, element: el };
	}

	function isEnameContext(funcParam, formula) {
		var u = String(funcParam || "").toUpperCase();
		if (u === "PALO.ENAME" || u === "ENAME" || /\.ENAME$/i.test(String(funcParam || ""))) {
			return true;
		}
		var g = String(formula || "").replace(/^\s*=\s*/, "");
		return /^_xlfn\./i.test(g)
			? /^_xlfn\.PALO\.ENAME\s*\(/i.test(g)
			: /^PALO\.ENAME\s*\(/i.test(g);
	}

	function isDatacContext(funcParam, formula) {
		var u = String(funcParam || "").toUpperCase();
		if (u === "PALO.DATAC" || u === "DATAC" || /\.DATAC$/i.test(String(funcParam || ""))) {
			return true;
		}
		var g = String(formula || "").replace(/^\s*=\s*/, "");
		g = g.replace(/^_xlfn\./i, "");
		return /^PALO\.DATAC\s*\(/i.test(g);
	}

	function runDatacLikeCopyPanel(params) {
		var panel = document.getElementById("datacLikeCopyPanel");
		var ta = document.getElementById("datacCommandInput");
		var err = document.getElementById("datacErr");
		var dh = document.getElementById("defaultHint");
		if (dh) {
			dh.style.display = "none";
		}
		if (panel) {
			panel.style.display = "block";
		}
		if (err) {
			err.textContent = "";
			err.style.color = "";
		}
		var btn = document.getElementById("btnDatacSplash");
		if (btn) {
			btn.addEventListener("click", function () {
				if (err) {
					err.textContent = "";
					err.style.color = "";
				}
				var text = ta ? String(ta.value).trim() : "";
				if (!text) {
					if (err) {
						err.style.color = "#a4262c";
						err.textContent = "Saisissez une commande (ex. 100 LIKE dim:élém;année:2025 ou COPY …).";
					}
					return;
				}
				var parsedPath = parseDatacLiteralPathForReplace(params.formula);
				if (parsedPath.error) {
					if (err) {
						err.style.color = "#a4262c";
						err.textContent = parsedPath.error;
					}
					return;
				}
				btn.disabled = true;
				loadSettingsAsync()
					.then(function (cfg) {
						if (!cfg.url || !cfg.username) {
							throw new Error("Configurez l’URL et l’utilisateur dans le volet Connexion (Palo).");
						}
						return getCachedSession(cfg);
					})
					.then(function (sess) {
						if (!sess || !sess.apiBase || !sess.sid) {
							throw new Error("Connexion serveur impossible.");
						}
						return fetchCellReplaceSplash(
							sess.apiBase,
							sess.sid,
							parsedPath.database,
							parsedPath.cube,
							parsedPath.path,
							text,
							1,
						);
					})
					.then(function () {
						if (err) {
							err.style.color = "#107c10";
							err.textContent =
								"Splash appliqué sur le cube. La formule PALO.DATAC de la cellule n’a pas été modifiée. Recalculez la feuille si la valeur affichée ne se met pas à jour.";
						}
						btn.disabled = false;
					})
					.catch(function (e) {
						if (err) {
							err.style.color = "#a4262c";
							err.textContent = e && e.message ? e.message : String(e);
						}
						btn.disabled = false;
					});
			});
		}
	}

	function excelFormulaStringLiteral(s) {
		return '"' + String(s == null ? "" : s).replace(/"/g, '""') + '"';
	}

	function buildEnameFormula(nameDatabase, nameDimension, elementName) {
		return (
			"=PALO.ENAME(" +
			excelFormulaStringLiteral(nameDatabase) +
			"," +
			excelFormulaStringLiteral(nameDimension) +
			"," +
			excelFormulaStringLiteral(elementName) +
			")"
		);
	}

	/**
	 * Ne remplace que le 3ᵉ argument (élément) ; conserve les expressions d’origine
	 * pour la base et la dimension (littéraux ou références de cellules).
	 */
	function buildUpdatedEnameFormulaPreservingArgs12(originalFormula, newElementName) {
		var raw = String(originalFormula || "").trim();
		if (!raw || typeof parsePaloEnameFirstThreeArgExpressions !== "function") {
			return null;
		}
		var exprs = parsePaloEnameFirstThreeArgExpressions(raw);
		if (!exprs || exprs.length < 3) {
			return null;
		}
		var mLead = raw.match(/^(\s*=\s*)(_xlfn\.)?/i);
		if (!mLead) {
			return null;
		}
		var eqPart = mLead[1];
		var xlfnPart = mLead[2] || "";
		var newThird = excelFormulaStringLiteral(newElementName);
		var body = exprs[0] + "," + exprs[1] + "," + newThird;
		return eqPart + xlfnPart + "PALO.ENAME(" + body + ")";
	}

	function fetchDimensionElementNamesSorted(apiBase, sid, nameDatabase, nameDimension) {
		var url = buildDimensionElementsRequestUrl(apiBase, sid, nameDatabase, nameDimension);
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
				var lines = stripBom(text)
					.split(/\r?\n/)
					.map(function (line) {
						return line.replace(/\s+$/, "");
					})
					.filter(function (line) {
						return line.length;
					});
				if (!lines.length) {
					throw new Error("Réponse /dimension/elements vide.");
				}
				if (
					lines[0].charAt(0) === "<" ||
					lines[0].toLowerCase().indexOf("<!doctype") !== -1
				) {
					throw new Error("Réponse inattendue (HTML).");
				}
				var seen = Object.create(null);
				var names = [];
				for (var i = 0; i < lines.length; i++) {
					var cells = splitPaloCsvLine(lines[i]);
					if (!cells || cells.length < 2) {
						continue;
					}
					var name = stripPaloCsvField(cells[1]).trim();
					if (!name || seen[name]) {
						continue;
					}
					seen[name] = true;
					names.push(name);
				}
				names.sort(function (a, b) {
					return a.localeCompare(b, "fr", { sensitivity: "base" });
				});
				return names;
			});
		});
	}

	function q(name) {
		var p = new URLSearchParams(window.location.search);
		return p.get(name) || "";
	}

	function setText(id, txt) {
		var el = document.getElementById(id);
		if (el) {
			el.textContent = txt;
		}
	}

	function sendClose() {
		try {
			Office.context.ui.messageParent("close");
		} catch (e) {
			window.close();
		}
	}

	function sendUpdateFormula(address, formula) {
		try {
			Office.context.ui.messageParent(
				JSON.stringify({
					action: "updateFormula",
					address: address,
					formula: formula,
				}),
			);
		} catch (e) {
			setText("elistErr", "Envoi à Excel impossible : " + (e.message || String(e)));
		}
	}

	function runEnamePicker(params) {
		var formulaRaw = q("formula");
		var dbQ = q("ename_db");
		var dimQ = q("ename_dim");
		var elQ = q("ename_el");
		var parsed = null;
		if (dbQ && dimQ && elQ) {
			parsed = { database: dbQ, dimension: dimQ, element: elQ };
		} else {
			parsed = parseEnameLiteralArgs(formulaRaw);
		}
		var elHost = document.getElementById("enamePicker");
		var elList = document.getElementById("enameList");
		var elHint = document.getElementById("enameHint");
		if (!parsed) {
			if (elHost) {
				elHost.style.display = "block";
			}
			setText(
				"elistErr",
				"Impossible d’obtenir la base, la dimension et l’élément. Utilisez des chaînes " +
					'(ex. "dwh","Dim","Paris"), ou des références vers ce classeur : A1, Feuille!B2, \'Ma feuille\'!C3 (valeurs lues par le complément). ' +
					"Références vers un autre fichier du type [Autre.xlsx]Feuille!A1 : Excel les calcule, mais le complément ne peut pas les lire — mettez la valeur dans une cellule de ce classeur (ex. =[Autre.xlsx]Feuille!A1) et référencez cette cellule. " +
					"Formules complexes : cellules intermédiaires.",
			);
			return;
		}
		if (elHost) {
			elHost.style.display = "block";
		}
		if (elHint) {
			elHint.textContent =
				"Base « " + parsed.database + " », dimension « " + parsed.dimension + " » — élément actuel : « " + parsed.element + " »";
		}
		setText("elistErr", "Chargement des éléments…");
		loadSettingsAsync()
			.then(function (cfg) {
				if (!cfg.url || !cfg.username) {
					throw new Error("Configurez l’URL et l’utilisateur dans le volet Connexion (Palo).");
				}
				return getCachedSession(cfg).then(function (sess) {
					return fetchDimensionElementNamesSorted(
						sess.apiBase,
						sess.sid,
						parsed.database,
						parsed.dimension,
					);
				});
			})
			.then(function (names) {
				setText("elistErr", "");
				if (!elList) {
					return;
				}
				elList.innerHTML = "";
				for (var i = 0; i < names.length; i++) {
					(function (name) {
						var row = document.createElement("button");
						row.type = "button";
						row.className = "ename-item";
						row.textContent = name;
						row.addEventListener("click", function () {
							var newFormula =
								buildUpdatedEnameFormulaPreservingArgs12(formulaRaw, name) ||
								buildEnameFormula(parsed.database, parsed.dimension, name);
							sendUpdateFormula(params.address, newFormula);
						});
						elList.appendChild(row);
					})(names[i]);
				}
			})
			.catch(function (err) {
				setText("elistErr", err && err.message ? err.message : String(err));
			});
	}

	Office.onReady(function () {
		var funcName = q("func");
		var formula = q("formula");
		var address = q("address");

		setText("address", address || "(inconnue)");
		setText("func", funcName || "(aucune)");
		setText("formula", formula || q("value") || "(vide)");

		document.getElementById("btnClose").addEventListener("click", sendClose);

		if (isEnameContext(funcName, formula)) {
			var dh = document.getElementById("defaultHint");
			if (dh) {
				dh.style.display = "none";
			}
			runEnamePicker({ address: address, formula: formula });
		} else if (isDatacContext(funcName, formula)) {
			var dh2 = document.getElementById("defaultHint");
			if (dh2) {
				dh2.style.display = "none";
			}
			runDatacLikeCopyPanel({ address: address, formula: formula });
		}
	});
})();
