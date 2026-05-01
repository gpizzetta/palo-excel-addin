/* global Office */
/* Popup « Action » : pour PALO.ENAME, charge les éléments de la dimension (tri alphabétique) et met à jour la formule au clic. */
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

	function parseStringLiteral(arg) {
		var t = String(arg == null ? "" : arg).trim();
		if (t.length < 2 || t.charAt(0) !== '"') {
			return null;
		}
		var out = "";
		for (var i = 1; i < t.length; i++) {
			var c = t.charAt(i);
			if (c === '"') {
				if (i + 1 < t.length && t.charAt(i + 1) === '"') {
					out += '"';
					i++;
					continue;
				}
				if (i === t.length - 1) {
					return out;
				}
				return null;
			}
			out += c;
		}
		return null;
	}

	function parseTopLevelArgs(inner) {
		var args = [];
		var cur = "";
		var depth = 0;
		var inStr = false;
		var i = 0;
		while (i < inner.length) {
			var c = inner.charAt(i);
			if (inStr) {
				if (c === '"') {
					if (i + 1 < inner.length && inner.charAt(i + 1) === '"') {
						cur += '""';
						i += 2;
						continue;
					}
					inStr = false;
					cur += c;
					i++;
					continue;
				}
				cur += c;
				i++;
				continue;
			}
			if (c === '"') {
				inStr = true;
				cur += c;
				i++;
				continue;
			}
			if (c === "(") {
				depth++;
				cur += c;
				i++;
				continue;
			}
			if (c === ")") {
				depth--;
				cur += c;
				i++;
				continue;
			}
			if (c === "," && depth === 0) {
				args.push(cur.trim());
				cur = "";
				i++;
				continue;
			}
			cur += c;
			i++;
		}
		if (cur.trim()) {
			args.push(cur.trim());
		}
		return args;
	}

	function parseEnameLiteralArgs(formula) {
		var f = String(formula == null ? "" : formula).trim();
		if (f.charAt(0) === "=") {
			f = f.slice(1).trim();
		}
		f = f.replace(/^_xlfn\./i, "");
		var m = f.match(/^PALO\.ENAME\s*\(/i);
		if (!m) {
			return null;
		}
		var start = f.indexOf("(");
		var depth = 0;
		var inStr = false;
		var end = -1;
		for (var i = start; i < f.length; i++) {
			var c = f.charAt(i);
			if (inStr) {
				if (c === '"') {
					if (i + 1 < f.length && f.charAt(i + 1) === '"') {
						i++;
						continue;
					}
					inStr = false;
				}
				continue;
			}
			if (c === '"') {
				inStr = true;
				continue;
			}
			if (c === "(") {
				depth++;
				continue;
			}
			if (c === ")") {
				depth--;
				if (depth === 0) {
					end = i;
					break;
				}
			}
		}
		if (end < 0) {
			return null;
		}
		var inner = f.slice(start + 1, end);
		var args = parseTopLevelArgs(inner);
		if (args.length < 3) {
			return null;
		}
		var db = parseStringLiteral(args[0]);
		var dim = parseStringLiteral(args[1]);
		var el = parseStringLiteral(args[2]);
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
		var parsed = parseEnameLiteralArgs(formulaRaw);
		var elHost = document.getElementById("enamePicker");
		var elList = document.getElementById("enameList");
		var elHint = document.getElementById("enameHint");
		if (!parsed) {
			if (elHost) {
				elHost.style.display = "block";
			}
			setText(
				"elistErr",
				"PALO.ENAME : les trois premiers arguments doivent être des chaînes littérales " +
					'(ex. ="Base","Dim","Élément") pour afficher la liste.',
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
							var newFormula = buildEnameFormula(parsed.database, parsed.dimension, name);
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
		}
	});
})();
