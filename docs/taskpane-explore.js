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

	function apiBaseCandidates(connectionUrl) {
		var u = new URL(connectionUrl.trim());
		var path = u.pathname.replace(/\/$/, "");
		if (path === "/api") {
			return [u.origin + "/api"];
		}
		return [u.origin + "/api", u.origin];
	}

	function parseCsvLines(text) {
		return text
			.split(/\r?\n/)
			.map(function (line) {
				return line.replace(/\s+$/, "");
			})
			.filter(function (line) {
				return line.length;
			})
			.map(function (line) {
				return line.split(";");
			});
	}

	function parseLoginSid(rows) {
		if (!rows.length) {
			throw new Error("Réponse login vide.");
		}
		var r = rows[0];
		var sid = r[0];
		if (sid === undefined || sid === "") {
			throw new Error("Identifiant de session manquant.");
		}
		if (/^[0-9]{1,5}$/.test(sid) && r.length > 1 && r[1]) {
			var code = parseInt(sid, 10);
			if (code > 0) {
				throw new Error(r.slice(1).join("; "));
			}
		}
		return sid;
	}

	function fetchCsv(url) {
		return fetch(url, {
			method: "GET",
			mode: "cors",
			cache: "no-store",
			credentials: "omit",
		}).then(function (res) {
			return res.text().then(function (text) {
				if (!res.ok) {
					throw new Error("HTTP " + res.status + " — " + text.slice(0, 300));
				}
				return text;
			});
		});
	}

	function loginAtBase(apiBase, user, password) {
		var q = new URLSearchParams({
			user: user,
			extern_password: password,
		});
		var url = apiBase + "/server/login?" + q.toString();
		return fetchCsv(url).then(function (text) {
			return parseLoginSid(parseCsvLines(text));
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

	function loadDatabases() {
		var q = new URLSearchParams({ sid: state.sid });
		return fetchCsv(state.apiBase + "/server/databases?" + q.toString()).then(function (text) {
			var rows = parseCsvLines(text);
			var list = [];
			for (var i = 0; i < rows.length; i++) {
				var row = rows[i];
				if (row.length < 2) {
					continue;
				}
				list.push({
					id: row[0],
					name: row[1],
				});
			}
			return list;
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
		return fetchCsv(state.apiBase + "/database/dimensions?" + q.toString()).then(function (text) {
			var rows = parseCsvLines(text);
			var list = [];
			for (var i = 0; i < rows.length; i++) {
				var row = rows[i];
				if (row.length < 2) {
					continue;
				}
				list.push({
					id: row[0],
					name: row[1],
				});
			}
			return list;
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
		return fetchCsv(state.apiBase + "/database/cubes?" + q.toString()).then(function (text) {
			var rows = parseCsvLines(text);
			var list = [];
			for (var i = 0; i < rows.length; i++) {
				var row = rows[i];
				if (row.length < 2) {
					continue;
				}
				list.push({
					id: row[0],
					name: row[1],
				});
			}
			return list;
		});
	}

	function renderDatabaseList() {
		var ul = document.getElementById("listDatabases");
		var empty = document.getElementById("emptyDatabases");
		ul.innerHTML = "";
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
			ul.innerHTML = "";
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
