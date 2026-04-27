/** Doit rester aligné avec <Version> dans docs/manifest.xml */
var ADDIN_VERSION = "1.0.35.0";

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

function hello() {
	return "hello world";
}

function version() {
	return ADDIN_VERSION;
}

/**
 * GET sur l’URL en mode CORS — pour tester HTTP/HTTPS et les en-têtes côté serveur Palo.
 * Si la réponse est bloquée (CORS, réseau, certificat), le message d’erreur l’indique souvent.
 */
function info(url) {
	if (url === undefined || url === null) {
		return "Indiquez une URL complète, ex. https://127.0.0.1:7777/server/info";
	}
	var s = String(url).trim();
	if (!s) {
		return "URL vide";
	}
	return fetch(s, {
		method: "GET",
		mode: "cors",
		cache: "no-store",
	})
		.then(function (res) {
			return (
				"OK — HTTP " +
				res.status +
				" " +
				res.statusText +
				" — Content-Type: " +
				(res.headers.get("content-type") || "(none)")
			);
		})
		.catch(function (err) {
			var msg = err && err.message ? err.message : String(err);
			return (
				"Échec: " +
				msg +
				" — (souvent CORS, réseau ou certificat ; serveur Palo en HTTPS si besoin)"
			);
		});
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

	function officeReady() {
		return new Promise(function (resolve, reject) {
			if (typeof Office === "undefined") {
				reject(new Error("Office.js indisponible."));
				return;
			}
			Office.onReady(function () {
				if (!Office.context || !Office.context.document || !Office.context.document.settings) {
					reject(new Error("Paramètres du classeur indisponibles (Office.context.document.settings)."));
					return;
				}
				resolve();
			});
		});
	}

	function loadSettingsAsync() {
		return officeReady().then(function () {
			return new Promise(function (resolve, reject) {
				var s = Office.context.document.settings;
				function read() {
					resolve({
						url: (s.get(KEYS.url) || "").trim(),
						username: (s.get(KEYS.username) || "").trim(),
						password: s.get(KEYS.password) || "",
					});
				}
				if (typeof s.refreshAsync === "function") {
					s.refreshAsync(function (ar) {
						if (ar.status !== Office.AsyncResultStatus.Succeeded) {
							var em = ar.error && ar.error.message;
							reject(new Error(em || "refreshAsync des paramètres a échoué."));
							return;
						}
						read();
					});
				} else {
					read();
				}
			});
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

	function normalizePathElements(pathElements) {
		if (pathElements == null) {
			return [];
		}
		if (Array.isArray(pathElements)) {
			var out = [];
			for (var i = 0; i < pathElements.length; i++) {
				var v = unwrapExcelScalar(pathElements[i]);
				if (v != null && String(v).trim() !== "") {
					out.push(String(v).trim());
				}
			}
			return out;
		}
		var one = unwrapExcelScalar(pathElements);
		return [String(one).trim()].filter(Boolean);
	}

	function parseCellValueFirstLine(text) {
		var lines = stripBom(text)
			.split(/\r?\n/)
			.map(function (line) {
				return line.replace(/\s+$/, "");
			})
			.filter(function (line) {
				return line.length;
			});
		if (!lines.length) {
			throw new Error("Réponse cell/value vide.");
		}
		var line = lines[0];
		if (line.charAt(0) === "<" || line.toLowerCase().indexOf("<!doctype") !== -1) {
			throw new Error("Réponse cell/value : HTML au lieu du CSV Palo.");
		}
		var cells = line.indexOf(";") >= 0 ? line.split(";") : line.split(",");
		if (cells.length < 3) {
			throw new Error("Réponse cell/value illisible : " + line.slice(0, 160));
		}
		var type = parseInt(cells[0].trim(), 10);
		var existsRaw = stripPaloCsvField(cells[1]);
		var exists = existsRaw === "1" || String(existsRaw).toLowerCase() === "true";
		var rawVal = cells.length > 2 ? stripPaloCsvField(cells[2]) : "";
		return { type: type, exists: exists, rawVal: rawVal };
	}

	function fetchCellValue(apiBase, sid, nameDatabase, nameCube, elementNames) {
		var namePath = elementNames.join(",");
		var q = new URLSearchParams({
			sid: sid,
			name_database: nameDatabase,
			name_cube: nameCube,
			name_path: namePath,
		});
		var url = apiBase + "/cell/value?" + q.toString();
		return fetch(url, {
			method: "GET",
			mode: "cors",
			cache: "no-store",
			credentials: "omit",
		}).then(function (res) {
			return res.text().then(function (text) {
				if (!res.ok) {
					throw new Error("cell/value HTTP " + res.status + " — " + text.slice(0, 500));
				}
				var o = parseCellValueFirstLine(text);
				if (o.type === 99 || isNaN(o.type)) {
					return o.rawVal ? String(o.rawVal) : "#ERROR";
				}
				if (!o.exists) {
					return "";
				}
				if (o.type === 1) {
					var n = parseFloat(String(o.rawVal).replace(",", "."));
					return isNaN(n) ? String(o.rawVal) : String(n);
				}
				if (o.type === 2) {
					return String(o.rawVal);
				}
				return String(o.rawVal);
			});
		});
	}

/**
 * Excel peut passer une référence de cellule comme scalaire imbriqué [[3]] ou matrice Office.
 * On extrait la valeur utile avant splash / valeur / noms.
 */
function unwrapExcelScalar(value) {
	if (value === undefined || value === null) {
		return value;
	}
	var cur = value;
	for (var guard = 0; guard < 10; guard++) {
		if (Array.isArray(cur)) {
			if (!cur.length) {
				return undefined;
			}
			cur = cur[0];
			continue;
		}
		if (cur && typeof cur === "object") {
			if (Array.isArray(cur.values) && cur.values.length) {
				var row0 = cur.values[0];
				cur = Array.isArray(row0) && row0.length ? row0[0] : row0;
				continue;
			}
			if (typeof cur.valueOf === "function") {
				var v = cur.valueOf();
				if (v !== cur) {
					cur = v;
					continue;
				}
			}
		}
		break;
	}
	return cur;
}

function parsePaloBooleanLike(value) {
	if (typeof value === "boolean") {
		return value;
	}
	if (value === undefined || value === null) {
		return null;
	}
	var s = String(value).trim().toLowerCase();
	if (!s) {
		return null;
	}
	if (s === "1" || s === "true" || s === "vrai" || s === "yes" || s === "y" || s === "on") {
		return true;
	}
	if (s === "0" || s === "false" || s === "faux" || s === "no" || s === "n" || s === "off") {
		return false;
	}
	return null;
}

function normalizeSplashMode(splash) {
	if (splash === undefined || splash === null || splash === "") {
		return 0;
	}
	if (typeof splash === "number") {
		if (!isFinite(splash)) {
			throw new Error("Paramètre splash invalide.");
		}
		var n = Math.round(splash);
		if (n < 0 || n > 5) {
			throw new Error("Paramètre splash hors plage (0..5).");
		}
		return n;
	}
	var boolLike = parsePaloBooleanLike(splash);
	if (boolLike !== null) {
		return boolLike ? 1 : 0;
	}
	var s = String(splash).trim().toLowerCase();
	if (s === "default") {
		return 1;
	}
	if (s === "add" || s === "add_base") {
		return 2;
	}
	if (s === "set" || s === "set_base") {
		return 3;
	}
	if (s === "set_populated") {
		return 4;
	}
	if (s === "add_populated") {
		return 5;
	}
	var parsed = parseInt(s, 10);
	if (!isNaN(parsed) && String(parsed) === s) {
		if (parsed < 0 || parsed > 5) {
			throw new Error("Paramètre splash hors plage (0..5).");
		}
		return parsed;
	}
	throw new Error("Paramètre splash invalide (attendu: booléen, 0..5, default/add/set).");
}

function parsePaloStatus(text, operationLabel) {
	var lines = stripBom(text)
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
	var first = lines[0];
	if (first.charAt(0) === "<" || first.toLowerCase().indexOf("<!doctype") !== -1) {
		throw new Error(operationLabel + " : HTML renvoyé au lieu du CSV Palo.");
	}
	var cells = first.indexOf(";") >= 0 ? first.split(";") : first.split(",");
	var c0 = cells.length ? stripPaloCsvField(cells[0]).trim() : "";
	if (/^[0-9]{1,6}$/.test(c0)) {
		var code = parseInt(c0, 10);
		if (code > 0) {
			var details = cells
				.slice(1)
				.map(function (c) {
					return stripPaloCsvField(String(c || "").trim());
				})
				.filter(Boolean)
				.join(" — ");
			if (!details) {
				details = "(pas de détail)";
			}
			throw new Error(
				operationLabel + " code " + code + ": " + details + " [ligne1=" + first.slice(0, 240) + "]",
			);
		}
	}
}

function stringifySetDataValue(value) {
	if (value === undefined || value === null) {
		return "";
	}
	if (typeof value === "number") {
		if (!isFinite(value)) {
			throw new Error("Valeur numérique invalide.");
		}
		return String(value);
	}
	if (typeof value === "boolean") {
		return value ? "1" : "0";
	}
	return String(value);
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

function replaceCellValue(apiBase, sid, nameDatabase, nameCube, elementNames, value, splashMode) {
	var namePath = elementNames.join(",");
	var valueAsString = stringifySetDataValue(value);
	var sm =
		splashMode === undefined || splashMode === null || splashMode === ""
			? 0
			: Math.round(Number(splashMode));
	if (isNaN(sm) || sm < 0 || sm > 5) {
		sm = 0;
	}
	var q = new URLSearchParams({
		sid: sid,
		name_database: nameDatabase,
		name_cube: nameCube,
		name_path: namePath,
		value: valueAsString,
		splash: String(sm),
		/** Doc Jedox : attendre la fin des push rules / partie async (défaut souvent 0). */
		wait: "1",
	});
	/** L’UI « API » du serveur est sous /api/... (HTML) ; l’endpoint CSV est /cell/replace (sans /api/). */
	var url = apiBase + "/cell/replace?" + q.toString();
	return fetch(url, {
		method: "GET",
		mode: "cors",
		cache: "no-store",
		credentials: "omit",
	}).then(function (res) {
		return res.text().then(function (text) {
			if (!res.ok) {
				throw new Error(
					"cell/replace HTTP " +
						res.status +
						" — " +
						shortResponseExcerpt(text) +
						" — url=" +
						url.replace(/sid=[^&]*/i, "sid=***"),
				);
			}
			parsePaloStatus(text, "cell/replace");
			return valueAsString;
		});
	});
}

function formatSetdataError(err, ctx) {
	var msg = err && err.message ? String(err.message) : String(err);
	if (!msg || msg === "[object Object]") {
		msg = "Erreur inconnue lors de l'ecriture dans le cube.";
	}
	msg = msg.replace(/\s+/g, " ").trim();
	var details = [];
	if (ctx) {
		details.push("db=" + (ctx.database || "(vide)"));
		details.push("cube=" + (ctx.cube || "(vide)"));
		details.push("path=[" + (ctx.path || []).join(" | ") + "]");
		details.push("splashArg=" + (ctx.splashArg === undefined ? "(absent)" : String(ctx.splashArg)));
		details.push("splashMode=" + String(ctx.splashMode));
		details.push("value=" + ctx.valuePreview);
	}
	var out = "PALO.SETDATA: " + msg;
	if (details.length) {
		out += " — " + details.join(" ; ");
	}
	if (ctx && typeof ctx.splashMode === "number" && ctx.splashMode >= 2) {
		out +=
			" | Splash 2–5 = modes Jedox sur chemins consolidés (répartition vers la base). Pour une intersection feuille, essayez splash 0 ou 1.";
	}
	out += " | API CSV: …/cell/replace (la page /api/cell/replace est la doc HTML, pas l’URL d’appel).";
	if (out.length > 900) {
		out = out.slice(0, 900) + "...";
	}
	return out;
}

	/**
	 * Lecture d’une cellule cube (équivalent Jedox PALO.DATAC : coordonnées par noms d’éléments).
	 * Utilise les identifiants du volet Connexion. Session réutilisée quelques minutes (évite un login par cellule).
	 */
	function datac(database, cube, element) {
		var db = database != null ? String(database).trim() : "";
		var cubeName = cube != null ? String(cube).trim() : "";
		var parts = normalizePathElements(element);
		return loadSettingsAsync()
			.then(function (cfg) {
				if (!cfg.url) {
					throw new Error("Configurez l’URL dans le volet Connexion (Palo).");
				}
				if (!cfg.username) {
					throw new Error("Utilisateur de connexion manquant (volet Connexion).");
				}
				if (!db) {
					throw new Error("Nom de base manquant.");
				}
				if (!cubeName) {
					throw new Error("Nom de cube manquant.");
				}
				if (!parts.length) {
					throw new Error("Indiquez au moins un nom d’élément (un par dimension du cube).");
				}
				return getCachedSession(cfg).then(function (sess) {
					return fetchCellValue(sess.apiBase, sess.sid, db, cubeName, parts);
				});
			});
	}

/**
 * Écriture d’une cellule cube (équivalent Jedox PALO.SETDATA).
 * Signature: value, splash, database, cube, element1, [element2], ...
 */
function setdata(value, splash, database, cube, element) {
	value = unwrapExcelScalar(value);
	splash = unwrapExcelScalar(splash);
	database = unwrapExcelScalar(database);
	cube = unwrapExcelScalar(cube);

	var db = database != null ? String(database).trim() : "";
	var cubeName = cube != null ? String(cube).trim() : "";
	var parts = normalizePathElements(element);
	var splashArg = splash;
	var splashMode;
	try {
		splashMode = normalizeSplashMode(splash);
	} catch (e) {
		return formatSetdataError(e, {
			database: db,
			cube: cubeName,
			path: parts,
			splashArg: splashArg,
			splashMode: "(erreur)",
			valuePreview: "(voir splash)",
		});
	}
	var valuePreview;
	try {
		valuePreview = stringifySetDataValue(value);
	} catch (e2) {
		return formatSetdataError(e2, {
			database: db,
			cube: cubeName,
			path: parts,
			splashArg: splashArg,
			splashMode: splashMode,
			valuePreview: "(valeur invalide)",
		});
	}
	if (valuePreview.length > 80) {
		valuePreview = valuePreview.slice(0, 80) + "...";
	}
	return loadSettingsAsync()
		.then(function (cfg) {
			if (!cfg.url) {
				throw new Error("Configurez l’URL dans le volet Connexion (Palo).");
			}
			if (!cfg.username) {
				throw new Error("Utilisateur de connexion manquant (volet Connexion).");
			}
			if (!db) {
				throw new Error("Nom de base manquant.");
			}
			if (!cubeName) {
				throw new Error("Nom de cube manquant.");
			}
			if (!parts.length) {
				throw new Error("Indiquez au moins un nom d’élément (un par dimension du cube).");
			}
			return getCachedSession(cfg).then(function (sess) {
				return replaceCellValue(sess.apiBase, sess.sid, db, cubeName, parts, value, splashMode);
			});
		})
		.catch(function (err) {
			return formatSetdataError(err, {
				database: db,
				cube: cubeName,
				path: parts,
				splashArg: splashArg,
				splashMode: splashMode,
				valuePreview: valuePreview,
			});
		});
}

CustomFunctions.associate("HELLO", hello);
CustomFunctions.associate("VERSION", version);
CustomFunctions.associate("INFO", info);
CustomFunctions.associate("DATAC", datac);
CustomFunctions.associate("SETDATA", setdata);
