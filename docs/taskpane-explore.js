(function () {
	var KEYS = {
		url: "palo_connection_url",
		username: "palo_connection_username",
		password: "palo_connection_password",
	};

	var state = {
		apiBase: "",
		sid: "",
		databases: [],
		selectedDb: null,
	};

	function setStatus(msg, kind) {
		var el = document.getElementById("status");
		el.textContent = msg || "";
		el.className = kind || "";
	}

	function getSettings() {
		var s = Office.context.document.settings;
		return {
			url: (s.get(KEYS.url) || "").trim(),
			username: (s.get(KEYS.username) || "").trim(),
			password: s.get(KEYS.password) || "",
		};
	}

	/** Base HTTP Palo : chemins /server/, /database/, etc. à la racine du host (pas de préfixe /api). */
	function apiBaseCandidates(connectionUrl) {
		var u = new URL(connectionUrl.trim());
		return [u.origin];
	}

	function stripBom(text) {
		return String(text).replace(/^\uFEFF/, "");
	}

	/** URL affichée dans les messages de debug (masque le mot de passe). */
	function redactUrlForDebug(url) {
		try {
			var u = new URL(url);
			if (u.searchParams.has("password")) {
				u.searchParams.set("password", "***");
			}
			return u.toString();
		} catch (e) {
			return String(url).replace(/([?&]password=)[^&]*/gi, "$1***");
		}
	}

	function truncateForDebug(text, maxLen) {
		maxLen = maxLen || 900;
		var t = stripBom(text).replace(/\s+/g, " ").trim();
		if (t.length <= maxLen) {
			return t;
		}
		return t.slice(0, maxLen) + "…";
	}

	function rejectIfHtml(text, requestUrl) {
		var t = stripBom(text).trim();
		if (!t.length) {
			return;
		}
		var lower = t.slice(0, 200).toLowerCase();
		if (
			t.charAt(0) === "<" ||
			lower.indexOf("<!doctype") !== -1 ||
			lower.indexOf("<html") !== -1 ||
			lower.indexOf("<head") !== -1 ||
			lower.indexOf("<body") !== -1
		) {
			var safeUrl = requestUrl ? redactUrlForDebug(requestUrl) : "(URL inconnue)";
			var excerpt = truncateForDebug(text, 900);
			throw new Error(
				"Le serveur a renvoyé du HTML au lieu du CSV Palo. Indiquez l’URL racine du serveur OLAP (ex. https://hôte:port), pas une page web.\n\n" +
					"URL interrogée : " +
					safeUrl +
					"\n\n" +
					"Réponse du serveur (extrait) :\n" +
					excerpt,
			);
		}
	}

	/** Nom d’objet Palo (base, dimension, cube) — refuse les fragments HTML issus d’une mauvaise réponse. */
	function isPlausibleObjectName(s) {
		if (s === undefined || s === null) {
			return false;
		}
		var t = String(s).trim();
		if (!t.length || t.length > 512) {
			return false;
		}
		if (/[<>]/.test(t) || /^\W+$/.test(t)) {
			return false;
		}
		return true;
	}

	function isNumericId(s) {
		return s !== undefined && s !== null && /^\d+$/.test(String(s).trim());
	}

	/** Découpe une ligne CSV Palo : point-virgule (Jedox), puis virgule ou tabulation. */
	function splitDataLine(line) {
		var delims = [";", ",", "\t"];
		for (var d = 0; d < delims.length; d++) {
			var parts = line.split(delims[d]);
			if (parts.length >= 2 && isNumericId(parts[0])) {
				var out = [];
				for (var i = 0; i < parts.length; i++) {
					out.push(parts[i].trim());
				}
				return out;
			}
		}
		return null;
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

	/** Login : première ligne CSV, sid en colonne 0 (souvent alphanumérique, pas un id numérique). */
	function parseLoginSidFromLines(lines) {
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

	function fetchCsv(url) {
		var safeUrl = redactUrlForDebug(url);
		return fetch(url, {
			method: "GET",
			mode: "cors",
			cache: "no-store",
			credentials: "omit",
		}).then(function (res) {
			return res.text().then(function (text) {
				if (!res.ok) {
					throw new Error(
						"HTTP " +
							res.status +
							" — URL : " +
							safeUrl +
							"\n\n" +
							"Réponse (extrait) :\n" +
							truncateForDebug(text, 600),
					);
				}
				return text;
			});
		});
	}

	function loginAtBase(apiBase, user, password) {
		if (typeof md5 !== "function") {
			return Promise.reject(
				new Error("Bibliothèque MD5 indisponible : le script md5.js doit être chargé avant taskpane-explore.js."),
			);
		}
		var q = new URLSearchParams({
			user: user,
			password: md5(String(password)),
		});
		var url = apiBase + "/server/login?" + q.toString();
		return fetchCsv(url).then(function (text) {
			rejectIfHtml(text, url);
			var lines = parseCsvLines(text);
			return parseLoginSidFromLines(lines);
		});
	}

	function discoverAndLogin(connectionUrl, user, password) {
		var bases = apiBaseCandidates(connectionUrl);
		function tryAt(i) {
			if (i >= bases.length) {
				return Promise.reject(
					new Error("Impossible de joindre l’API Palo (URL ou CORS ; essayez https://hôte:port)."),
				);
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

	function parseIdNameList(text, kindLabel, requestUrl) {
		rejectIfHtml(text, requestUrl);
		var lines = parseCsvLines(text);
		var list = [];
		for (var i = 0; i < lines.length; i++) {
			var cells = splitDataLine(lines[i]);
			if (!cells || cells.length < 2) {
				continue;
			}
			var id = cells[0];
			var name = cells[1];
			if (!isNumericId(id) || !isPlausibleObjectName(name)) {
				continue;
			}
			list.push({ id: id, name: name });
		}
		if (lines.length && !list.length) {
			var safeUrl = requestUrl ? redactUrlForDebug(requestUrl) : "(URL inconnue)";
			throw new Error(
				"Réponse " +
					kindLabel +
					" illisible (CSV Palo attendu : id numérique puis nom).\n\n" +
					"URL interrogée : " +
					safeUrl +
					"\n\n" +
					"Réponse du serveur (extrait) :\n" +
					truncateForDebug(text, 900),
			);
		}
		return list;
	}

	function loadDatabases() {
		var q = new URLSearchParams({ sid: state.sid });
		var url = state.apiBase + "/server/databases?" + q.toString();
		return fetchCsv(url).then(function (text) {
			return parseIdNameList(text, "bases", url);
		});
	}

	function loadDimensionsForDb(nameDatabase) {
		var q = new URLSearchParams({
			sid: state.sid,
			name_database: nameDatabase,
			show_system: "1",
			show_attribute: "1",
			show_info: "1",
		});
		var url = state.apiBase + "/database/dimensions?" + q.toString();
		return fetchCsv(url).then(function (text) {
			return parseIdNameList(text, "dimensions", url);
		});
	}

	function loadCubesForDb(nameDatabase) {
		var q = new URLSearchParams({
			sid: state.sid,
			name_database: nameDatabase,
			show_system: "1",
			show_attribute: "1",
			show_info: "1",
		});
		var url = state.apiBase + "/database/cubes?" + q.toString();
		return fetchCsv(url).then(function (text) {
			return parseIdNameList(text, "cubes", url);
		});
	}

	function renderDatabaseList() {
		var ul = document.getElementById("listDatabases");
		var empty = document.getElementById("emptyDatabases");
		while (ul.firstChild) {
			ul.removeChild(ul.firstChild);
		}
		var dbs = state.databases;
		if (!dbs.length) {
			empty.style.display = "block";
			return;
		}
		empty.style.display = "none";
		for (var i = 0; i < dbs.length; i++) {
			(function (db) {
				var li = document.createElement("li");
				li.textContent = db.name;
				li.title = "id " + db.id;
				li.addEventListener("click", function () {
					selectDatabase(db);
				});
				ul.appendChild(li);
			})(dbs[i]);
		}
	}

	function renderLists(dimensions, cubes) {
		function fill(ulId, emptyId, items, labelField) {
			var ul = document.getElementById(ulId);
			var empty = document.getElementById(emptyId);
			while (ul.firstChild) {
				ul.removeChild(ul.firstChild);
			}
			if (!items.length) {
				empty.style.display = "block";
				return;
			}
			empty.style.display = "none";
			for (var i = 0; i < items.length; i++) {
				var li = document.createElement("li");
				li.textContent = items[i][labelField];
				li.title = "id " + items[i].id;
				ul.appendChild(li);
			}
		}
		fill("listDimensions", "emptyDimensions", dimensions, "name");
		fill("listCubes", "emptyCubes", cubes, "name");
	}

	function showView(which) {
		document.getElementById("viewDatabases").className = which === "databases" ? "active" : "";
		document.getElementById("viewDetail").className = which === "detail" ? "active" : "";
		document.getElementById("btnBack").style.display = which === "detail" ? "inline-block" : "none";
	}

	function selectDatabase(db) {
		state.selectedDb = db;
		setStatus("Chargement…", "");
		document.getElementById("detailTitle").textContent = "Base : " + db.name;
		Promise.all([loadDimensionsForDb(db.name), loadCubesForDb(db.name)])
			.then(function (results) {
				renderLists(results[0], results[1]);
				setStatus("", "");
				showView("detail");
			})
			.catch(function (err) {
				var msg = err && err.message ? err.message : String(err);
				setStatus(msg, "err");
			});
	}

	function refreshAll() {
		var cfg = getSettings();
		if (!cfg.url) {
			setStatus("Configurez d’abord l’URL dans Connexion.", "err");
			return;
		}
		var btn = document.getElementById("btnRefresh");
		btn.disabled = true;
		setStatus("Connexion au serveur…", "");
		discoverAndLogin(cfg.url, cfg.username, cfg.password)
			.then(function (session) {
				state.apiBase = session.apiBase;
				state.sid = session.sid;
				return loadDatabases();
			})
			.then(function (dbs) {
				state.databases = dbs;
				state.selectedDb = null;
				renderDatabaseList();
				showView("databases");
				setStatus(dbs.length + " base(s) chargée(s).", "ok");
			})
			.catch(function (err) {
				var msg = err && err.message ? err.message : String(err);
				setStatus(msg, "err");
			})
			.then(function () {
				btn.disabled = false;
			});
	}

	function onBack() {
		showView("databases");
		setStatus("", "");
	}

	Office.onReady(function () {
		document.getElementById("btnRefresh").addEventListener("click", refreshAll);
		document.getElementById("btnBack").addEventListener("click", onBack);
		refreshAll();
	});
})();
