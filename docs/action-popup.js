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

	/**
	 * Première ligne de GET /cell/value (même format que parseCellValueFirstLine dans functions.js) :
	 * type ; exists ; valeur [, …] — la valeur utile est cells[2], pas cells[0] (sinon 1+10=11 si type=1).
	 */
	function parsePaloCellValueFirstLine(line) {
		var cells = splitPaloCsvLine(line);
		if (cells.length < 3) {
			throw new Error(
				"Réponse /cell/value inattendue (moins de 3 champs) : " + line.slice(0, 200),
			);
		}
		var type = parseInt(String(stripPaloCsvField(cells[0])).trim(), 10);
		var existsRaw = stripPaloCsvField(cells[1]);
		var exists = existsRaw === "1" || String(existsRaw).toLowerCase() === "true";
		var rawVal = cells.length > 2 ? stripPaloCsvField(cells[2]) : "";
		return { type: type, exists: exists, rawVal: rawVal };
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

	/**
	 * Paramètre datac_r : tableau JSON [base, cube, el1, …] rempli par commands.js
	 * après lecture des cellules référencées (même logique que PALO.ENAME).
	 */
	function parseDatacResolvedFromQuery() {
		var raw = q("datac_r");
		if (!raw) {
			return null;
		}
		try {
			var arr = JSON.parse(raw);
			if (!Array.isArray(arr) || arr.length < 3) {
				return null;
			}
			var db = String(arr[0] == null ? "" : arr[0]).trim();
			var cube = String(arr[1] == null ? "" : arr[1]).trim();
			var path = [];
			for (var i = 2; i < arr.length; i++) {
				path.push(String(arr[i] == null ? "" : arr[i]).trim());
			}
			if (!db || !cube || !path.length) {
				return null;
			}
			return { database: db, cube: cube, path: path };
		} catch (eParse) {
			return null;
		}
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

	/** Jedox : type élément « consolidé » (colonne type /dimension/elements, ex. C++ PaloSpreadsheetFuncs). */
	var DATAC_ELEMENT_TYPE_CONSOLIDATED = 4;

	var datacGuidedWired = false;
	var datacRuntimeCtx = null;

	function fetchTextNoStore(url) {
		return fetch(url, {
			method: "GET",
			mode: "cors",
			cache: "no-store",
			credentials: "omit",
		}).then(function (res) {
			return res.text().then(function (text) {
				if (!res.ok) {
					throw new Error("HTTP " + res.status + " — " + excerptCellApiBody(text));
				}
				return text;
			});
		});
	}

	function parseIdNameListFromCsv(text) {
		var lines = stripBom(text)
			.split(/\r?\n/)
			.map(function (line) {
				return line.replace(/\s+$/, "");
			})
			.filter(function (line) {
				return line.length;
			});
		var list = [];
		for (var i = 0; i < lines.length; i++) {
			if (lines[i].charAt(0) === "<") {
				continue;
			}
			var cells = splitPaloCsvLine(lines[i]);
			if (!cells || cells.length < 2) {
				continue;
			}
			var id = stripPaloCsvField(cells[0]).trim();
			var name = stripPaloCsvField(cells[1]).trim();
			if (!/^[0-9]+$/.test(id) || !name) {
				continue;
			}
			list.push({ id: id, name: name });
		}
		return list;
	}

	function parseCubeInfoDimensionIdsFromCsv(text) {
		var lines = stripBom(text)
			.split(/\r?\n/)
			.map(function (line) {
				return line.replace(/\s+$/, "");
			})
			.filter(function (line) {
				return line.length;
			});
		if (!lines.length) {
			return [];
		}
		var cells = splitPaloCsvLine(lines[0]);
		if (!cells || cells.length < 4) {
			cells = lines[0].split(";").map(function (c) {
				return c.trim();
			});
		}
		if (!cells || cells.length < 4) {
			return [];
		}
		var raw = stripPaloCsvField(cells[3]);
		if (!raw || !String(raw).trim()) {
			return [];
		}
		return String(raw)
			.split(",")
			.map(function (x) {
				return x.trim();
			})
			.filter(function (x) {
				return /^[0-9]+$/.test(x);
			});
	}

	function cubeIdFromName(cubes, cubeName) {
		var want = String(cubeName || "").trim();
		for (var i = 0; i < cubes.length; i++) {
			if (cubes[i].name === want) {
				return cubes[i].id;
			}
		}
		return null;
	}

	function dimensionNameById(dimList, dimId) {
		for (var i = 0; i < dimList.length; i++) {
			if (dimList[i].id === String(dimId)) {
				return dimList[i].name;
			}
		}
		return null;
	}

	function parseDimensionElementRow(line) {
		var cells = splitPaloCsvLine(line);
		if (!cells || cells.length < 7) {
			return null;
		}
		var id = stripPaloCsvField(cells[0]).trim();
		var name = stripPaloCsvField(cells[1]).trim();
		var typeStr = stripPaloCsvField(cells[6]).trim();
		var typeNum = parseInt(typeStr, 10);
		if (!/^[0-9]+$/.test(id) || !name) {
			return null;
		}
		if (isNaN(typeNum)) {
			typeNum = null;
		}
		return { id: id, name: name, type: typeNum };
	}

	function findElementRowInDimensionCsv(text, elementName) {
		var want = String(elementName || "").trim();
		var lines = stripBom(text)
			.split(/\r?\n/)
			.map(function (line) {
				return line.replace(/\s+$/, "");
			})
			.filter(function (line) {
				return line.length;
			});
		for (var i = 0; i < lines.length; i++) {
			if (lines[i].charAt(0) === "<") {
				continue;
			}
			var row = parseDimensionElementRow(lines[i]);
			if (row && row.name === want) {
				return row;
			}
		}
		return null;
	}

	function loadCubesList(apiBase, sid, nameDatabase) {
		var q = new URLSearchParams({
			sid: sid,
			name_database: nameDatabase,
			show_system: "1",
			show_attribute: "1",
			show_info: "1",
		});
		var url = String(apiBase).replace(/\/$/, "") + "/database/cubes?" + q.toString();
		return fetchTextNoStore(url).then(parseIdNameListFromCsv);
	}

	function loadCubeDimensionIds(apiBase, sid, nameDatabase, cubeId) {
		var q = new URLSearchParams({
			sid: sid,
			name_database: nameDatabase,
			cube: String(cubeId),
		});
		var url = String(apiBase).replace(/\/$/, "") + "/cube/info?" + q.toString();
		return fetchTextNoStore(url).then(parseCubeInfoDimensionIdsFromCsv);
	}

	function loadDatabaseDimensionsList(apiBase, sid, nameDatabase) {
		var q = new URLSearchParams({
			sid: sid,
			name_database: nameDatabase,
			show_system: "1",
			show_normal: "1",
			show_attribute: "0",
			show_virtual_attribute: "0",
			show_info: "1",
			show_permission: "0",
		});
		var url = String(apiBase).replace(/\/$/, "") + "/database/dimensions?" + q.toString();
		return fetchTextNoStore(url).then(parseIdNameListFromCsv);
	}

	function loadDimensionElementsText(apiBase, sid, nameDatabase, nameDimension) {
		var url = buildDimensionElementsRequestUrl(apiBase, sid, nameDatabase, nameDimension);
		return fetchTextNoStore(url);
	}

	function datacAnalyzePathForConsolidation(sess, pathInfo) {
		var apiBase = sess.apiBase;
		var sid = sess.sid;
		var db = pathInfo.database;
		var cubeName = pathInfo.cube;
		var path = pathInfo.path || [];
		return loadCubesList(apiBase, sid, db)
			.then(function (cubes) {
				var cid = cubeIdFromName(cubes, cubeName);
				if (!cid) {
					throw new Error('Cube « ' + cubeName + ' » introuvable dans la base « ' + db + ' ».');
				}
				return loadCubeDimensionIds(apiBase, sid, db, cid).then(function (dimIds) {
					return loadDatabaseDimensionsList(apiBase, sid, db).then(function (dimRows) {
						return { dimIds: dimIds, dimRows: dimRows, cubeId: cid };
					});
				});
			})
			.then(function (pack) {
				var dimIds = pack.dimIds;
				if (dimIds.length !== path.length) {
					throw new Error(
						"Nombre de dimensions du cube (" +
							String(dimIds.length) +
							") ≠ nombre d’arguments chemin dans PALO.DATAC (" +
							String(path.length) +
							").",
					);
				}
				var dimOrderNames = [];
				for (var d = 0; d < dimIds.length; d++) {
					var dn = dimensionNameById(pack.dimRows, dimIds[d]);
					if (!dn) {
						throw new Error("Dimension id " + dimIds[d] + " : nom introuvable.");
					}
					dimOrderNames.push(dn);
				}
				var chain = Promise.resolve({
					hasConsolidation: false,
					labels: [],
					dimOrderNames: dimOrderNames,
				});
				for (var j = 0; j < path.length; j++) {
					(function (idx, dimNm, elNm) {
						chain = chain.then(function (acc) {
							return loadDimensionElementsText(apiBase, sid, db, dimNm).then(function (csv) {
								var row = findElementRowInDimensionCsv(csv, elNm);
								if (!row) {
									throw new Error(
										'Élément « ' + elNm + ' » introuvable dans la dimension « ' + dimNm + ' ».',
									);
								}
								if (row.type === DATAC_ELEMENT_TYPE_CONSOLIDATED) {
									acc.hasConsolidation = true;
									acc.labels.push(dimNm + " : « " + elNm + " » (consolidé)");
								}
								return acc;
							});
						});
					})(j, dimOrderNames[j], path[j]);
				}
				return chain;
			});
	}

	function fillDatacSplashSelect(sel) {
		if (!sel) {
			return;
		}
		/** Aligné sur normalizeSplashMode (functions.js) et doc Jedox Splashing / # ! !! !# !!#. */
		var opts = [
			{
				v: "0",
				t: "0 — Aucun splash (pas de décomposition ; souvent impossible sur consolidé pur)",
			},
			{
				v: "1",
				t: "1 — Default : splash serveur par défaut (répartition sur les bases ; poids / logique Jedox — pas « parts égales » # seules)",
			},
			{
				v: "2",
				t: "2 — Add base (!!) : la même valeur est ajoutée à chaque base (pas une répartition du total ; cf. # / mode 1)",
			},
			{
				v: "3",
				t: "3 — Set base (!) : la même valeur sur chaque base — écrase toutes les bases liées avec cette valeur",
			},
			{
				v: "4",
				t: "4 — Set populated (!#) : uniquement bases déjà peuplées — erreur si toutes vides (pas de path référence LIKE)",
			},
			{
				v: "5",
				t: "5 — Add populated (!!#) : idem — addition seulement où une base a déjà une valeur",
			},
		];
		sel.innerHTML = "";
		for (var i = 0; i < opts.length; i++) {
			var o = document.createElement("option");
			o.value = opts[i].v;
			o.textContent = opts[i].t;
			sel.appendChild(o);
		}
		sel.value = "1";
	}

	function fetchCellValueByNames(apiBase, sid, nameDatabase, nameCube, pathArr) {
		var q = new URLSearchParams({
			sid: sid,
			name_database: nameDatabase,
			name_cube: nameCube,
			name_path: (pathArr || []).join(","),
		});
		var url = String(apiBase).replace(/\/$/, "") + "/cell/value?" + q.toString();
		return fetchTextNoStore(url).then(function (text) {
			var line = stripBom(text)
				.split(/\r?\n/)
				.map(function (l) {
					return l.replace(/\s+$/, "");
				})
				.filter(function (l) {
					return l.length;
				})[0];
			if (!line) {
				throw new Error("Réponse /cell/value vide.");
			}
			var o = parsePaloCellValueFirstLine(line);
			if (o.type === 99 || isNaN(o.type)) {
				throw new Error(
					"Cellule non numérique (type " + o.type + ") : addition impossible.",
				);
			}
			if (!o.exists) {
				return 0;
			}
			if (o.type === 1) {
				var n = parseFloat(String(o.rawVal).replace(",", "."));
				if (isNaN(n)) {
					throw new Error(
						"Valeur actuelle non numérique (« " + o.rawVal + " ») : addition impossible.",
					);
				}
				return n;
			}
			throw new Error(
				"Addition réservée aux cellules numériques (type 1) ; type actuel : " + o.type + ".",
			);
		});
	}

	function parseCellCopyOk(text) {
		var t = stripBom(String(text || "")).trim();
		if (t === "1" || t.charAt(0) === "1") {
			return;
		}
		parsePaloStatus(text);
	}

	/**
	 * Mode avancé : « 300 like 2025 » ou « COPY année:2025 » (doc Jedox LIKE — chemins partiels).
	 * LIKE avec valeur → GET /cell/copy + paramètre value (cell_copy.api), pas /cell/replace.
	 */
	function tryParseDatacLikeOrCopy(text) {
		var raw = String(text || "").trim();
		if (!raw) {
			return null;
		}
		var mCopy = raw.match(/^\s*copy\s+(.+)$/i);
		if (mCopy) {
			var ps = String(mCopy[1] || "").trim();
			if (!ps) {
				return null;
			}
			return { kind: "copy", pathSpec: ps };
		}
		var mLikeApos = raw.match(/^\s*'(-?[0-9]+(?:[.,][0-9]+)?)\s+like\s+(.+)$/i);
		if (mLikeApos) {
			var n1 = parseFloat(String(mLikeApos[1]).replace(",", "."));
			if (!isNaN(n1)) {
				return { kind: "like", value: n1, pathSpec: String(mLikeApos[2] || "").trim() };
			}
		}
		var mLike = raw.match(/^\s*([-+]?[0-9]+(?:[.,][0-9]+)?)\s+like\s+(.+)$/i);
		if (mLike) {
			var n2 = parseFloat(String(mLike[1]).replace(",", "."));
			if (!isNaN(n2)) {
				return { kind: "like", value: n2, pathSpec: String(mLike[2] || "").trim() };
			}
		}
		return null;
	}

	function parseLikeCopyPathSpec(spec) {
		var out = [];
		var parts = String(spec || "").split(";");
		for (var i = 0; i < parts.length; i++) {
			var seg = String(parts[i] || "").trim();
			if (!seg) {
				continue;
			}
			var ci = seg.indexOf(":");
			if (ci < 0) {
				out.push({ bare: stripPaloCsvField(seg) });
			} else {
				out.push({
					dim: seg.slice(0, ci).trim(),
					element: stripPaloCsvField(seg.slice(ci + 1).trim()),
				});
			}
		}
		return out;
	}

	function locateBareElementDimensionIndex(sess, nameDatabase, dimOrderNames, elName) {
		var want = String(elName || "").trim();
		if (!want) {
			return Promise.reject(new Error("Segment de chemin LIKE/COPY vide."));
		}
		return Promise.all(
			dimOrderNames.map(function (dimNm) {
				return loadDimensionElementsText(sess.apiBase, sess.sid, nameDatabase, dimNm).then(function (csv) {
					return findElementRowInDimensionCsv(csv, want) ? dimNm : null;
				});
			}),
		).then(function (hits) {
			var diList = [];
			for (var i = 0; i < hits.length; i++) {
				if (hits[i]) {
					diList.push(i);
				}
			}
			if (!diList.length) {
				throw new Error('Élément « ' + want + ' » introuvable dans aucune dimension du cube.');
			}
			if (diList.length > 1) {
				throw new Error(
					'Élément « ' +
						want +
						' » ambigu (plusieurs dimensions). Utilisez NomDimension:élément (doc Jedox LIKE).',
				);
			}
			return diList[0];
		});
	}

	function buildSourcePathForLikeCopy(sess, nameDatabase, targetPath, dimOrderNames, segments) {
		var source = targetPath.slice();
		var chain = Promise.resolve();
		for (var s = 0; s < segments.length; s++) {
			(function (seg) {
				chain = chain.then(function () {
					if (seg.dim) {
						var di = -1;
						for (var j = 0; j < dimOrderNames.length; j++) {
							if (String(dimOrderNames[j]).toLowerCase() === String(seg.dim).toLowerCase()) {
								di = j;
								break;
							}
						}
						if (di < 0) {
							throw new Error(
								'Dimension « ' +
									seg.dim +
									' » introuvable. Ordre du cube : ' +
									dimOrderNames.join(" → ") +
									".",
							);
						}
						return loadDimensionElementsText(
							sess.apiBase,
							sess.sid,
							nameDatabase,
							dimOrderNames[di],
						).then(function (csv) {
							if (!findElementRowInDimensionCsv(csv, seg.element)) {
								throw new Error(
									'Élément « ' +
										seg.element +
										' » introuvable dans « ' +
										dimOrderNames[di] +
										' ».',
								);
							}
							source[di] = seg.element;
						});
					}
					return locateBareElementDimensionIndex(
						sess,
						nameDatabase,
						dimOrderNames,
						seg.bare,
					).then(function (diBare) {
						source[diBare] = seg.bare;
					});
				});
			})(segments[s]);
		}
		return chain.then(function () {
			return source;
		});
	}

	function fetchCellCopyByNames(apiBase, sid, nameDatabase, nameCube, namePathFrom, namePathTo, useRules, copyValue) {
		var q = new URLSearchParams({
			sid: sid,
			name_database: nameDatabase,
			name_cube: nameCube,
			function: "0",
			name_path: namePathFrom,
			name_path_to: namePathTo,
		});
		if (useRules) {
			q.set("use_rules", "1");
		}
		if (copyValue !== undefined && copyValue !== null && copyValue !== "") {
			var vnum = Number(copyValue);
			if (!isNaN(vnum) && isFinite(vnum)) {
				q.set("value", String(vnum));
			}
		}
		var url = String(apiBase).replace(/\/$/, "") + "/cell/copy?" + q.toString();
		return fetchTextNoStore(url).then(parseCellCopyOk);
	}

	function wireDatacGuidedOnce(params, pathInfo) {
		if (datacGuidedWired) {
			return;
		}
		datacGuidedWired = true;
		var err = document.getElementById("datacErr");
		var ta = document.getElementById("datacCommandInput");
		var btnLegacy = document.getElementById("btnDatacSplash");

		function clearErr() {
			if (err) {
				err.textContent = "";
				err.style.color = "";
			}
		}

		document.getElementById("btnDatacApplySet").addEventListener("click", function () {
			clearErr();
			var ctx = datacRuntimeCtx;
			if (!ctx || !ctx.sess) {
				return;
			}
			var v = document.getElementById("datacSetValueInput");
			var val = v ? String(v.value) : "";
			var splash = 1;
			if (ctx.hasConsolidation) {
				var sEl = document.getElementById("datacSplashSelect");
				splash = sEl ? parseInt(sEl.value, 10) : 1;
				if (isNaN(splash) || splash < 0 || splash > 5) {
					splash = 1;
				}
			}
			var btn = document.getElementById("btnDatacApplySet");
			btn.disabled = true;
			fetchCellReplaceSplash(
				ctx.sess.apiBase,
				ctx.sess.sid,
				pathInfo.database,
				pathInfo.cube,
				pathInfo.path,
				val,
				splash,
			)
				.then(function () {
					if (err) {
						err.style.color = "#107c10";
						err.textContent =
							"Valeur envoyée. Recalculez la feuille si l’affichage ne se met pas à jour. PALO.DATAC inchangé.";
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

		document.getElementById("btnDatacApplyAdd").addEventListener("click", function () {
			clearErr();
			var ctx = datacRuntimeCtx;
			if (!ctx || !ctx.sess) {
				return;
			}
			var inp = document.getElementById("datacAddValueInput");
			var addPart = inp ? parseFloat(String(inp.value).replace(",", ".")) : NaN;
			if (isNaN(addPart)) {
				if (err) {
					err.style.color = "#a4262c";
					err.textContent = "Saisissez un nombre à additionner.";
				}
				return;
			}
			var splash = 1;
			if (ctx.hasConsolidation) {
				var sEl2 = document.getElementById("datacSplashSelect");
				splash = sEl2 ? parseInt(sEl2.value, 10) : 1;
				if (isNaN(splash) || splash < 0 || splash > 5) {
					splash = 1;
				}
			}
			var btnA = document.getElementById("btnDatacApplyAdd");
			btnA.disabled = true;
			fetchCellValueByNames(
				ctx.sess.apiBase,
				ctx.sess.sid,
				pathInfo.database,
				pathInfo.cube,
				pathInfo.path,
			)
				.then(function (cur) {
					return fetchCellReplaceSplash(
						ctx.sess.apiBase,
						ctx.sess.sid,
						pathInfo.database,
						pathInfo.cube,
						pathInfo.path,
						String(cur + addPart),
						splash,
					);
				})
				.then(function () {
					if (err) {
						err.style.color = "#107c10";
						err.textContent =
							"Somme écrite. Recalculez la feuille si besoin. PALO.DATAC inchangé.";
					}
					btnA.disabled = false;
				})
				.catch(function (e) {
					if (err) {
						err.style.color = "#a4262c";
						err.textContent = e && e.message ? e.message : String(e);
					}
					btnA.disabled = false;
				});
		});

		document.getElementById("btnDatacApplyCopy").addEventListener("click", function () {
			clearErr();
			var ctx = datacRuntimeCtx;
			if (!ctx || !ctx.sess) {
				return;
			}
			var inpC = document.getElementById("datacCopyFromInput");
			var raw = inpC ? String(inpC.value).trim() : "";
			if (!raw) {
				if (err) {
					err.style.color = "#a4262c";
					err.textContent = "Saisissez le chemin source (éléments séparés par des virgules).";
				}
				return;
			}
			var parts = raw.split(",").map(function (x) {
				return x.trim();
			});
			if (parts.length !== pathInfo.path.length) {
				if (err) {
					err.style.color = "#a4262c";
					err.textContent =
						"Le chemin source doit avoir " +
						String(pathInfo.path.length) +
						" éléments (comme le cube), séparés par des virgules.";
				}
				return;
			}
			var useRules = document.getElementById("datacCopyUseRules");
			var ur = useRules && useRules.checked;
			var fromComma = parts.join(",");
			var toComma = pathInfo.path.join(",");
			var btnC = document.getElementById("btnDatacApplyCopy");
			btnC.disabled = true;
			fetchCellCopyByNames(
				ctx.sess.apiBase,
				ctx.sess.sid,
				pathInfo.database,
				pathInfo.cube,
				fromComma,
				toComma,
				ur,
			)
				.then(function () {
					if (err) {
						err.style.color = "#107c10";
						err.textContent =
							"Copie /cell/copy exécutée. Recalculez la feuille si besoin. PALO.DATAC inchangé.";
					}
					btnC.disabled = false;
				})
				.catch(function (e) {
					if (err) {
						err.style.color = "#a4262c";
						err.textContent = e && e.message ? e.message : String(e);
					}
					btnC.disabled = false;
				});
		});

		if (btnLegacy) {
			btnLegacy.addEventListener("click", function () {
				clearErr();
				var text = ta ? String(ta.value).trim() : "";
				if (!text) {
					if (err) {
						err.style.color = "#a4262c";
						err.textContent = "Saisissez une commande (ex. 300 like 2025 ou COPY Année:2025).";
					}
					return;
				}
				var parsedPath = parseDatacResolvedFromQuery();
				if (!parsedPath) {
					parsedPath = parseDatacLiteralPathForReplace(params.formula);
				}
				if (parsedPath.error) {
					if (err) {
						err.style.color = "#a4262c";
						err.textContent = parsedPath.error;
					}
					return;
				}
				var parsedCmd = tryParseDatacLikeOrCopy(text);
				if (!parsedCmd) {
					if (err) {
						err.style.color = "#a4262c";
						err.textContent =
							"Commande non reconnue. Exemples : « 300 like 2025 » (répartition type LIKE : " +
							"seules les dimensions listées changent entre source et cible), « COPY Année:2025 ». " +
							"Ambiguïté : « NomDimension:élément ». Voir docs/palo-like-copy-datac-action.md § 5.2.";
					}
					return;
				}
				var segs = parseLikeCopyPathSpec(parsedCmd.pathSpec);
				if (!segs.length) {
					if (err) {
						err.style.color = "#a4262c";
						err.textContent = "Chemin après LIKE/COPY vide (séparez par des ;).";
					}
					return;
				}
				btnLegacy.disabled = true;
				var useRulesAdv =
					document.getElementById("datacCopyUseRules") &&
					document.getElementById("datacCopyUseRules").checked;
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
						return datacAnalyzePathForConsolidation(sess, parsedPath).then(function (analysis) {
							return buildSourcePathForLikeCopy(
								sess,
								parsedPath.database,
								parsedPath.path,
								analysis.dimOrderNames,
								segs,
							).then(function (sourceArr) {
								var fromC = sourceArr.join(",");
								var toC = parsedPath.path.join(",");
								var valOpt =
									parsedCmd.kind === "like" ? parsedCmd.value : undefined;
								return fetchCellCopyByNames(
									sess.apiBase,
									sess.sid,
									parsedPath.database,
									parsedPath.cube,
									fromC,
									toC,
									useRulesAdv,
									valOpt,
								);
							});
						});
					})
					.then(function () {
						if (err) {
							err.style.color = "#107c10";
							err.textContent =
								parsedCmd.kind === "like"
									? "LIKE exécuté via /cell/copy (valeur cible). PALO.DATAC inchangé. Recalculez la feuille si besoin."
									: "COPY exécuté via /cell/copy. PALO.DATAC inchangé. Recalculez la feuille si besoin.";
						}
						btnLegacy.disabled = false;
					})
					.catch(function (e) {
						if (err) {
							err.style.color = "#a4262c";
							err.textContent = e && e.message ? e.message : String(e);
						}
						btnLegacy.disabled = false;
					});
			});
		}
	}

	function runDatacLikeCopyPanel(params) {
		var panel = document.getElementById("datacLikeCopyPanel");
		var err = document.getElementById("datacErr");
		var dh = document.getElementById("defaultHint");
		var statusEl = document.getElementById("datacAnalysisStatus");
		var guided = document.getElementById("datacGuidedActions");
		var splashRow = document.getElementById("datacSplashRow");
		var dimHint = document.getElementById("datacDimOrderHint");
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
		if (statusEl) {
			statusEl.textContent = "Analyse du chemin (dimensions / consolidations)…";
		}
		if (guided) {
			guided.style.display = "none";
		}

		var pathInfo = parseDatacResolvedFromQuery();
		if (!pathInfo) {
			pathInfo = parseDatacLiteralPathForReplace(params.formula);
		}
		if (pathInfo.error) {
			if (statusEl) {
				statusEl.textContent = "";
			}
			if (err) {
				err.style.color = "#a4262c";
				err.textContent = pathInfo.error;
			}
			return;
		}

		fillDatacSplashSelect(document.getElementById("datacSplashSelect"));

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
				return datacAnalyzePathForConsolidation(sess, pathInfo).then(function (analysis) {
					return { sess: sess, analysis: analysis };
				});
			})
			.then(function (pack) {
				var analysis = pack.analysis;
				datacRuntimeCtx = { sess: pack.sess, hasConsolidation: analysis.hasConsolidation };
				if (statusEl) {
					statusEl.textContent = analysis.hasConsolidation
						? "Consolidation détectée : " + (analysis.labels.length ? analysis.labels.join(" ; ") : "oui") + " — choisissez un mode splash pour les écritures (§ G.5 B)."
						: "Aucun élément consolidé sur ce chemin — parcours simple (§ G.5 A). Splash masqué sauf besoin avancé.";
				}
				if (splashRow) {
					splashRow.style.display = analysis.hasConsolidation ? "block" : "none";
				}
				if (dimHint && analysis.dimOrderNames) {
					dimHint.textContent =
						"Ordre des dimensions du cube : " + analysis.dimOrderNames.join(" → ");
				}
				if (guided) {
					guided.style.display = "block";
				}
				wireDatacGuidedOnce(params, pathInfo);
			})
			.catch(function (e) {
				datacRuntimeCtx = null;
				if (statusEl) {
					statusEl.textContent =
						"Analyse impossible : " + (e && e.message ? e.message : String(e)) + " — utilisez le mode avancé LIKE/COPY si le serveur est injoignable.";
				}
				if (guided) {
					guided.style.display = "none";
				}
			});
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
