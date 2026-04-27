/** Doit rester aligné avec <Version> dans docs/manifest.xml */
var ADDIN_VERSION = "1.0.41.0";

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
			throw new Error(shortResponseExcerpt(text) || "(vide)");
		}
		var line = lines[0];
		if (line.charAt(0) === "<" || line.toLowerCase().indexOf("<!doctype") !== -1) {
			throw new Error(shortResponseExcerpt(line));
		}
		var cells = line.indexOf(";") >= 0 ? line.split(";") : line.split(",");
		if (cells.length < 3) {
			throw new Error(shortResponseExcerpt(line));
		}
		var type = parseInt(cells[0].trim(), 10);
		var existsRaw = stripPaloCsvField(cells[1]);
		var exists = existsRaw === "1" || String(existsRaw).toLowerCase() === "true";
		var rawVal = cells.length > 2 ? stripPaloCsvField(cells[2]) : "";
		return { type: type, exists: exists, rawVal: rawVal };
	}

	function fetchCellValue(apiBase, sid, nameDatabase, nameCube, elementNames) {
		var url = buildCellValueRequestUrl(apiBase, sid, nameDatabase, nameCube, elementNames);
		var urlRed = redactPaloSidInUrl(url);
		/** Toujours résoudre : Excel remplace souvent tout reject par « Une erreur interne s’est produite. » */
		return fetch(url, {
			method: "GET",
			mode: "cors",
			cache: "no-store",
			credentials: "omit",
		})
			.then(function (res) {
				return res.text().then(function (text) {
					try {
						if (!res.ok) {
							return formatUrlAndServerResponse(
								urlRed,
								excerptCellApiBody(text) || "HTTP " + res.status,
							);
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
					} catch (inner) {
						return formatUrlAndServerResponse(urlRed, inner.message || String(inner));
					}
				});
			})
			.catch(function (err) {
				return formatUrlAndServerResponse(urlRed, err.message || String(err));
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

/**
 * Entier splash pour l’URL HTTP /cell/replace (doc Jedox OLAP : 0..5).
 * Réf. C++ Excel : GenericCell::getSplashMode() (PaloSpreadsheetFuncs) pour un nombre
 * n’accepte que 0..3 : 0 none, 1 default, 2 = MODE_SPLASH_SET, 3 = MODE_SPLASH_ADD.
 * Sur le fil HTTP, « set base / add base » sont 3 et 2 (ordre inverse des libellés C++).
 * Ici les littéraux numériques suivent la doc HTTP ; pour les libellés texte, utiliser
 * add_base / set_base (voir chaînes ci-dessous).
 */
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
			throw new Error("Paramètre splash hors plage (0..5, doc HTTP /cell/replace).");
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
	throw new Error(
		"Paramètre splash invalide (booléen, 0..5 selon doc HTTP, ou mots default/add_base/set_base…).",
	);
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
		var cells = line.indexOf(";") >= 0 ? line.split(";") : line.split(",");
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

/** Corps de réponse HTTP pour affichage d’erreur cellule (plus long que shortResponseExcerpt). */
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

function buildCellValueRequestUrl(apiBase, sid, nameDatabase, nameCube, elementNames) {
	var namePath = (elementNames || []).join(",");
	var q = new URLSearchParams({
		sid: sid,
		name_database: nameDatabase,
		name_cube: nameCube,
		name_path: namePath,
	});
	return String(apiBase).replace(/\/$/, "") + "/cell/value?" + q.toString();
}

function buildCellReplaceRequestUrl(apiBase, sid, nameDatabase, nameCube, elementNames, value, splashMode) {
	var namePath = (elementNames || []).join(",");
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
	});
	return String(apiBase).replace(/\/$/, "") + "/cell/replace?" + q.toString();
}

function replaceCellValue(apiBase, sid, nameDatabase, nameCube, elementNames, value, splashMode) {
	/** L’UI « API » du serveur est sous /api/... (HTML) ; l’endpoint CSV est /cell/replace (sans /api/). */
	var valueAsString;
	try {
		valueAsString = stringifySetDataValue(value);
	} catch (se) {
		return Promise.resolve(formatUrlAndServerResponse("", se.message || String(se)));
	}
	var url = buildCellReplaceRequestUrl(apiBase, sid, nameDatabase, nameCube, elementNames, value, splashMode);
	console.log(url);
	var urlRed = redactPaloSidInUrl(url);
	/** Toujours résoudre : Excel masque souvent les reject par « Une erreur interne s’est produite. » */
	return fetch(url, {
		method: "GET",
		mode: "cors",
		cache: "no-store",
		credentials: "omit",
	})
		.then(function (res) {
			return res.text().then(function (text) {
				console.log(text);
				try {
					if (!res.ok) {
						return formatUrlAndServerResponse(
							urlRed,
							excerptCellApiBody(text) || "HTTP " + res.status,
						);
					}
					parsePaloStatus(text);
					return valueAsString;
				} catch (inner) {
					return formatUrlAndServerResponse(urlRed, inner.message || String(inner));
				}
			});
		})
		.catch(function (err) {
			return formatUrlAndServerResponse(urlRed, err.message || String(err));
		});
}

function normalizeOneLineText(s) {
	return String(s == null ? "" : s).replace(/\s+/g, " ").trim();
}

function stripLegacyUrlSuffix(msg) {
	var m = String(msg == null ? "" : msg);
	var idx = m.indexOf(" — url=");
	if (idx !== -1) {
		return m.slice(0, idx).trim();
	}
	return m.trim();
}

/** Affichage cellule : uniquement l’URL (sid masqué) et le corps / ligne renvoyé par le serveur. */
function formatUrlAndServerResponse(urlRedacted, serverText) {
	var u = normalizeOneLineText(urlRedacted);
	var t = normalizeOneLineText(stripLegacyUrlSuffix(serverText));
	if (!u) {
		return t || "(vide)";
	}
	if (!t) {
		return "url=" + u + " | (vide)";
	}
	var out = "url=" + u + " | " + t;
	if (out.length > 4000) {
		out = out.slice(0, 4000) + "...";
	}
	return out;
}

function formatSetdataError(err, requestUrlRedacted) {
	var urlRedacted = requestUrlRedacted != null ? String(requestUrlRedacted).trim() : "";
	var msg = err && err.message ? String(err.message) : String(err);
	if (!msg || msg === "[object Object]") {
		msg = "(vide)";
	}
	msg = normalizeOneLineText(stripLegacyUrlSuffix(msg));
	return formatUrlAndServerResponse(urlRedacted, msg);
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
			})
			.catch(function (err) {
				var msg = err && err.message ? err.message : String(err);
				return formatUrlAndServerResponse("", msg);
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
		return formatSetdataError(e, "");
	}
	var valuePreview;
	try {
		valuePreview = stringifySetDataValue(value);
	} catch (e2) {
		return formatSetdataError(e2, "");
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
			return formatSetdataError(err, "");
		});
}

CustomFunctions.associate("HELLO", hello);
CustomFunctions.associate("VERSION", version);
CustomFunctions.associate("INFO", info);
CustomFunctions.associate("DATAC", datac);
CustomFunctions.associate("SETDATA", setdata);
