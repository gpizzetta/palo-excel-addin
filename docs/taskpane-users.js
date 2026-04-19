(function () {
	var KEYS = {
		url: "palo_connection_url",
		username: "palo_connection_username",
		password: "palo_connection_password",
	};

	var state = {
		apiBase: "",
		sid: "",
		currentUserId: null,
		currentUserName: null,
		usersMeta: null,
	};

	function setStatus(msg, kind) {
		var el = document.getElementById("status");
		if (el) {
			el.textContent = msg || "";
			el.className = kind || "";
		}
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
		return [u.origin];
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
				"Le serveur a renvoyé du HTML au lieu du CSV Palo.\n\nURL : " +
					safeUrl +
					"\n\nExtrait :\n" +
					excerpt,
			);
		}
	}

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
							truncateForDebug(text, 600),
					);
				}
				return text;
			});
		});
	}

	function loginAtBase(apiBase, user, password) {
		if (typeof md5 !== "function") {
			return Promise.reject(new Error("md5.js doit être chargé avant taskpane-users.js."));
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
			var name = stripPaloCsvField(cells[1]);
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
					" illisible.\n\nURL : " +
					safeUrl +
					"\n\n" +
					truncateForDebug(text, 900),
			);
		}
		return list;
	}

	function parseDatabaseDimensionsList(text, requestUrl) {
		rejectIfHtml(text, requestUrl);
		var lines = parseCsvLines(text);
		var list = [];
		var anyIdNameRow = false;
		for (var i = 0; i < lines.length; i++) {
			var cells = splitDataLine(lines[i]);
			if (!cells || cells.length < 2) {
				continue;
			}
			var id = cells[0];
			var name = stripPaloCsvField(cells[1]);
			if (isNumericId(id) && isPlausibleObjectName(name)) {
				anyIdNameRow = true;
			}
			if (!isNumericId(id) || !isPlausibleObjectName(name)) {
				continue;
			}
			var typeStr = cells.length > 6 ? stripPaloCsvField(cells[6]) : "";
			var typeNum = typeStr === "" ? null : parseInt(typeStr, 10);
			list.push({ id: id, name: name, type: typeNum });
		}
		if (lines.length && !list.length && !anyIdNameRow) {
			throw new Error("Réponse dimensions illisible.\n\n" + truncateForDebug(text, 900));
		}
		return list;
	}

	function parseChildrenIdList(s) {
		if (!s || !String(s).trim()) {
			return [];
		}
		return String(s)
			.split(",")
			.map(function (x) {
				return x.trim();
			})
			.filter(function (x) {
				return isNumericId(x);
			});
	}

	function parseWeightsForChildren(raw, childCount) {
		if (!childCount) {
			return [];
		}
		var out = [];
		var parts = String(raw || "").split(",");
		for (var i = 0; i < childCount; i++) {
			var p = parts[i] ? parts[i].trim() : "";
			var v = parseFloat(String(p).replace(",", "."));
			out.push(isNaN(v) ? 1 : v);
		}
		return out;
	}

	function parseDimensionElementsList(text, requestUrl) {
		rejectIfHtml(text, requestUrl);
		var lines = parseCsvLines(text);
		var list = [];
		var anyRow = false;
		for (var i = 0; i < lines.length; i++) {
			var cells = splitDataLine(lines[i]);
			if (!cells || cells.length < 7) {
				continue;
			}
			var id = cells[0];
			var name = stripPaloCsvField(cells[1]);
			if (isNumericId(id) && isPlausibleObjectName(name)) {
				anyRow = true;
			}
			if (!isNumericId(id) || !isPlausibleObjectName(name)) {
				continue;
			}
			var typeStr = stripPaloCsvField(cells[6]);
			var typeNum = parseInt(typeStr, 10);
			if (isNaN(typeNum)) {
				typeNum = null;
			}
			var childrenRaw = cells.length > 10 ? cells[10] : "";
			var weightsRaw = cells.length > 11 ? cells[11] : "";
			var childrenIds = parseChildrenIdList(childrenRaw);
			var weights = parseWeightsForChildren(weightsRaw, childrenIds.length);
			list.push({
				id: id,
				name: name,
				type: typeNum,
				childrenIds: childrenIds,
				weights: weights,
			});
		}
		if (lines.length && !list.length && !anyRow) {
			throw new Error("Réponse éléments illisible.\n\n" + truncateForDebug(text, 900));
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
			show_normal: "1",
			show_attribute: "1",
			show_virtual_attribute: "1",
			show_info: "1",
			show_permission: "1",
		});
		var url = state.apiBase + "/database/dimensions?" + q.toString();
		return fetchCsv(url).then(function (text) {
			return parseDatabaseDimensionsList(text, url);
		});
	}

	function dimensionKindLabel(typeNum) {
		if (typeNum === null || typeNum === undefined || isNaN(typeNum)) {
			return "";
		}
		if (typeNum === 2) {
			return "propriété";
		}
		if (typeNum === 5) {
			return "propriété virtuelle";
		}
		if (typeNum === 0) {
			return "normale";
		}
		if (typeNum === 3) {
			return "système";
		}
		return "type " + typeNum;
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

	function parseUserInfoFirstLine(text, requestUrl) {
		rejectIfHtml(text, requestUrl);
		var lines = parseCsvLines(text);
		if (!lines.length) {
			throw new Error("Réponse user_info vide.");
		}
		var cells = splitDataLine(lines[0]);
		if (!cells || cells.length < 2) {
			cells = lines[0].split(";").map(function (c) {
				return c.trim();
			});
		}
		if (!cells || cells.length < 2) {
			throw new Error("user_info illisible.");
		}
		var id = cells[0].trim();
		var name = stripPaloCsvField(cells[1]);
		if (!isNumericId(id) || !String(name).trim()) {
			throw new Error("Identité utilisateur illisible dans user_info.");
		}
		return { id: id, name: name };
	}

	function loadCurrentUserFromServer() {
		var q = new URLSearchParams({ sid: state.sid, show_permission: "0", show_info: "0" });
		var url = state.apiBase + "/server/user_info?" + q.toString();
		return fetchCsv(url).then(function (text) {
			var u = parseUserInfoFirstLine(text, url);
			state.currentUserId = u.id;
			state.currentUserName = u.name;
		});
	}

	function pickUsersDimensionName(dims) {
		var preferExact = ["#_USERS_", "#_USERS", "#_USER_", "#_USER"];
		var i, j, nm;
		for (j = 0; j < preferExact.length; j++) {
			for (i = 0; i < dims.length; i++) {
				if (dims[i].name === preferExact[j]) {
					return dims[i].name;
				}
			}
		}
		for (i = 0; i < dims.length; i++) {
			nm = dims[i].name;
			if (nm && /^#_USERS/i.test(nm)) {
				return nm;
			}
		}
		for (i = 0; i < dims.length; i++) {
			nm = dims[i].name;
			if (nm && /^#_USER/i.test(nm)) {
				return nm;
			}
		}
		return null;
	}

	function findDatabaseWithUsersDimension(dbs, idx) {
		if (idx >= dbs.length) {
			return Promise.reject(
				new Error(
					'Dimension utilisateurs « #_USERS_ » introuvable (ni variante #_USERS / #_USER_) dans les bases accessibles.',
				),
			);
		}
		var dbName = dbs[idx].name;
		return loadDimensionsForDb(dbName).then(function (dims) {
			var userDimName = pickUsersDimensionName(dims);
			if (userDimName) {
				return { nameDatabase: dbName, userDimensionName: userDimName };
			}
			return findDatabaseWithUsersDimension(dbs, idx + 1);
		});
	}

	function fetchDimensionElementsByName(nameDatabase, nameDimension) {
		var q = new URLSearchParams({
			sid: state.sid,
			name_database: nameDatabase,
			name_dimension: nameDimension,
			show_permission: "0",
		});
		var url = state.apiBase + "/dimension/elements?" + q.toString();
		return fetchCsv(url).then(function (text) {
			return parseDimensionElementsList(text, url);
		});
	}

	function loadUsersFromUsersDimension() {
		state.usersMeta = null;
		return loadDatabases()
			.then(function (dbs) {
				return findDatabaseWithUsersDimension(dbs, 0);
			})
			.then(function (ctx) {
				state.usersMeta = {
					nameDatabase: ctx.nameDatabase,
					userDimensionName: ctx.userDimensionName,
				};
				return fetchDimensionElementsByName(ctx.nameDatabase, ctx.userDimensionName).then(function (elements) {
					var out = [];
					for (var j = 0; j < elements.length; j++) {
						var el = elements[j];
						if (el.type === 4) {
							continue;
						}
						out.push({ id: el.id, name: el.name });
					}
					return out;
				});
			});
	}

	function renderServerUsersList(users) {
		var ul = document.getElementById("listServerUsers");
		var empty = document.getElementById("emptyServerUsers");
		var hint = document.getElementById("userHint");
		if (!ul || !empty) {
			return;
		}
		while (ul.firstChild) {
			ul.removeChild(ul.firstChild);
		}
		if (hint) {
			if (state.usersMeta) {
				hint.className = "show";
				hint.textContent =
					"Liste issue de la dimension « " +
					state.usersMeta.userDimensionName +
					" » — base « " +
					state.usersMeta.nameDatabase +
					" ».";
			} else {
				hint.className = "";
				hint.textContent = "";
			}
		}
		if (!users.length) {
			empty.style.display = "block";
			empty.textContent = "Aucun utilisateur.";
			return;
		}
		empty.style.display = "none";
		for (var i = 0; i < users.length; i++) {
			(function (u) {
				var li = document.createElement("li");
				var span = document.createElement("span");
				span.className = "user-name";
				span.textContent = u.name;
				span.title = "id " + u.id;
				li.appendChild(span);
				var canDel = state.currentUserId && u.id !== state.currentUserId;
				if (canDel) {
					var b = document.createElement("button");
					b.type = "button";
					b.className = "user-del";
					b.setAttribute("aria-label", "Supprimer l’utilisateur " + u.name);
					b.textContent = "−";
					b.addEventListener("click", function () {
						onDeleteServerUser(u);
					});
					li.appendChild(b);
				}
				ul.appendChild(li);
			})(users[i]);
		}
	}

	function refreshServerUsersPanel() {
		return loadUsersFromUsersDimension().then(function (list) {
			renderServerUsersList(list);
		});
	}

	function loadAllHashCubes() {
		return loadDatabases().then(function (dbs) {
			return Promise.all(
				dbs.map(function (db) {
					return loadCubesForDb(db.name).then(function (cubes) {
						var rows = [];
						for (var i = 0; i < cubes.length; i++) {
							var c = cubes[i];
							if (c.name && c.name.indexOf("#_") === 0) {
								rows.push({
									nameDatabase: db.name,
									cubeId: c.id,
									cubeName: c.name,
								});
							}
						}
						return rows;
					});
				}),
			).then(function (arrays) {
				var out = [];
				for (var a = 0; a < arrays.length; a++) {
					out = out.concat(arrays[a]);
				}
				out.sort(function (x, y) {
					var kx = x.nameDatabase + "\0" + x.cubeName;
					var ky = y.nameDatabase + "\0" + y.cubeName;
					return kx < ky ? -1 : kx > ky ? 1 : 0;
				});
				return out;
			});
		});
	}

	function renderHashCubesList(rows) {
		var ul = document.getElementById("listHashCubes");
		var empty = document.getElementById("emptyHashCubes");
		if (!ul || !empty) {
			return;
		}
		while (ul.firstChild) {
			ul.removeChild(ul.firstChild);
		}
		if (!rows.length) {
			empty.style.display = "block";
			return;
		}
		empty.style.display = "none";
		for (var i = 0; i < rows.length; i++) {
			var r = rows[i];
			var li = document.createElement("li");
			var span = document.createElement("span");
			span.className = "user-name";
			span.textContent = r.nameDatabase + " — " + r.cubeName;
			span.title = "id cube " + r.cubeId;
			li.appendChild(span);
			ul.appendChild(li);
		}
	}

	function refreshHashCubesPanel() {
		return loadAllHashCubes().then(function (rows) {
			renderHashCubesList(rows);
		});
	}

	function loadAllDimensionsAllDatabases() {
		return loadDatabases().then(function (dbs) {
			return Promise.all(
				dbs.map(function (db) {
					return loadDimensionsForDb(db.name).then(function (dims) {
						return dims.map(function (d) {
							return {
								nameDatabase: db.name,
								dimId: d.id,
								dimName: d.name,
								type: d.type,
							};
						});
					});
				}),
			).then(function (arrays) {
				var out = [];
				for (var a = 0; a < arrays.length; a++) {
					out = out.concat(arrays[a]);
				}
				out.sort(function (x, y) {
					var kx = x.nameDatabase + "\0" + x.dimName;
					var ky = y.nameDatabase + "\0" + y.dimName;
					return kx < ky ? -1 : kx > ky ? 1 : 0;
				});
				return out;
			});
		});
	}

	function renderAllDimensionsList(rows) {
		var ul = document.getElementById("listAllDimensions");
		var empty = document.getElementById("emptyAllDimensions");
		if (!ul || !empty) {
			return;
		}
		while (ul.firstChild) {
			ul.removeChild(ul.firstChild);
		}
		if (!rows.length) {
			empty.style.display = "block";
			return;
		}
		empty.style.display = "none";
		for (var i = 0; i < rows.length; i++) {
			var r = rows[i];
			var li = document.createElement("li");
			var span = document.createElement("span");
			span.className = "user-name";
			var kind = dimensionKindLabel(r.type);
			span.textContent =
				r.nameDatabase + " — " + r.dimName + (kind ? " · " + kind : "");
			span.title = "id dimension " + r.dimId + (kind ? " — " + kind : "");
			li.appendChild(span);
			ul.appendChild(li);
		}
	}

	function refreshAllDimensionsPanel() {
		return loadAllDimensionsAllDatabases().then(function (rows) {
			renderAllDimensionsList(rows);
		});
	}

	function onDeleteServerUser(u) {
		if (state.currentUserId && u.id === state.currentUserId) {
			setStatus("Vous ne pouvez pas supprimer votre propre compte depuis cette session.", "err");
			return;
		}
		if (!confirm('Supprimer l’utilisateur « ' + u.name + ' » (id ' + u.id + ') ?')) {
			return;
		}
		if (!state.usersMeta) {
			setStatus("Métadonnées de la dimension utilisateurs indisponibles. Actualisez la liste.", "err");
			return;
		}
		var q = new URLSearchParams({
			sid: state.sid,
			name_database: state.usersMeta.nameDatabase,
			name_dimension: state.usersMeta.userDimensionName,
			element: String(u.id),
		});
		var url = state.apiBase + "/element/destroy?" + q.toString();
		setStatus("Suppression…", "");
		fetchCsv(url)
			.then(function () {
				return refreshServerUsersPanel();
			})
			.then(function () {
				setStatus("Utilisateur supprimé.", "ok");
			})
			.catch(function (err) {
				setStatus(err && err.message ? err.message : String(err), "err");
			});
	}

	function getAddinPageBaseUrl() {
		var href = window.location.href.split("#")[0];
		var i = href.lastIndexOf("/");
		if (i < 0) {
			return href.indexOf("?") >= 0 ? href.substring(0, href.indexOf("?")) + "/" : href + "/";
		}
		return href.substring(0, i + 1);
	}

	function openOfficeDialogPage(htmlFile, queryParams, onRefreshDone) {
		if (
			typeof Office === "undefined" ||
			!Office.context ||
			!Office.context.ui ||
			typeof Office.context.ui.displayDialogAsync !== "function"
		) {
			setStatus("Boîte de dialogue Office indisponible (displayDialogAsync).", "err");
			return;
		}
		var qs = new URLSearchParams(queryParams);
		var url = getAddinPageBaseUrl() + htmlFile + "?v=1.0.24.0&" + qs.toString();
		var dialogOpts = {
			height: 90,
			width: 90,
			displayInIframe: true,
		};
		Office.context.ui.displayDialogAsync(url, dialogOpts, function (asyncResult) {
			if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
				setStatus("Impossible d’ouvrir la fenêtre Office.", "err");
				return;
			}
			var dialog = asyncResult.value;
			dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
				var msg = arg.message;
				try {
					dialog.close();
				} catch (e) {}
				if (msg === "refresh" && typeof onRefreshDone === "function") {
					onRefreshDone();
				}
			});
		});
	}

	function openModalCreateUser() {
		if (!state.sid) {
			setStatus("Connectez-vous d’abord (Actualiser).", "err");
			return;
		}
		if (!state.usersMeta) {
			setStatus("Chargez d’abord la liste utilisateurs (Actualiser).", "err");
			return;
		}
		openOfficeDialogPage(
			"dialog-create-user.html",
			{
				apiBase: state.apiBase,
				sid: state.sid,
				name_database: state.usersMeta.nameDatabase,
				name_dimension: state.usersMeta.userDimensionName,
			},
			function () {
				refreshServerUsersPanel()
					.then(function () {
						setStatus("Utilisateur créé.", "ok");
					})
					.catch(function (err) {
						setStatus(err && err.message ? err.message : String(err), "err");
					});
			},
		);
	}

	function refreshAll() {
		var cfg = getSettings();
		if (!cfg.url) {
			setStatus("Configurez d’abord l’URL dans l’onglet Connexion.", "err");
			return;
		}
		var btn = document.getElementById("btnRefresh");
		if (btn) {
			btn.disabled = true;
		}
		setStatus("Connexion…", "");
		discoverAndLogin(cfg.url, cfg.username, cfg.password)
			.then(function (session) {
				state.apiBase = session.apiBase;
				state.sid = session.sid;
				return loadCurrentUserFromServer().catch(function () {
					state.currentUserId = null;
					state.currentUserName = null;
				});
			})
			.then(function () {
				setStatus("Chargement des utilisateurs…", "");
				return Promise.all([
					refreshServerUsersPanel(),
					refreshHashCubesPanel(),
					refreshAllDimensionsPanel(),
				]);
			})
			.then(function () {
				setStatus("Liste à jour.", "ok");
			})
			.catch(function (err) {
				setStatus(err && err.message ? err.message : String(err), "err");
				renderServerUsersList([]);
				renderHashCubesList([]);
				renderAllDimensionsList([]);
			})
			.then(function () {
				if (btn) {
					btn.disabled = false;
				}
			});
	}

	Office.onReady(function () {
		document.getElementById("btnRefresh").addEventListener("click", refreshAll);
		document.getElementById("btnUserAdd").addEventListener("click", openModalCreateUser);
		refreshAll();
	});
})();
