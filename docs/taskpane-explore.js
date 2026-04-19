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
		selectedDimension: null,
		currentView: "databases",
		/** Droits sur la base courante : "N" | "R" | "W" | "D" (voir `/database/info` + show_permission). */
		databasePermission: null,
		/** Dernière liste d’éléments parsée pour la dimension affichée (voir `parseDimensionElementsList`). */
		dimensionElements: [],
		/** Id Jedox de la dimension d’attributs (#…) liée à la dimension courante (`/dimension/info`, colonne attributes_dimension). */
		attributesDimensionId: null,
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

	/** Champ CSV Palo / Jedox : guillemets optionnels autour du nom ("dwh" → dwh). */
	function stripPaloCsvField(s) {
		var t = String(s).trim();
		if (t.length >= 2 && t.charAt(0) === '"' && t.charAt(t.length - 1) === '"') {
			return t.slice(1, -1).replace(/""/g, '"');
		}
		return t;
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

	function permissionAllowsWrite(perm) {
		return perm === "W" || perm === "D";
	}

	/** Dimensions système (non supprimables via l’UI). */
	function dimensionTypeDeletable(typeNum) {
		if (typeNum === null || typeNum === undefined || isNaN(typeNum)) {
			return true;
		}
		return typeNum === 0 || typeNum === 3;
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
			if (typeNum !== null && !isNaN(typeNum) && (typeNum === 2 || typeNum === 5)) {
				continue;
			}
			var perm = cells.length > 15 ? stripPaloCsvField(cells[15]) : "";
			list.push({ id: id, name: name, type: typeNum, permission: perm });
		}
		if (lines.length && !list.length && !anyIdNameRow) {
			var safeUrl = requestUrl ? redactUrlForDebug(requestUrl) : "(URL inconnue)";
			throw new Error(
				"Réponse dimensions illisible (CSV Palo attendu).\n\n" +
					"URL interrogée : " +
					safeUrl +
					"\n\n" +
					"Réponse du serveur (extrait) :\n" +
					truncateForDebug(text, 900),
			);
		}
		return list;
	}

	function loadDatabasePermissionInfo(nameDatabase) {
		var q = new URLSearchParams({
			sid: state.sid,
			name_database: nameDatabase,
			show_permission: "1",
		});
		var url = state.apiBase + "/database/info?" + q.toString();
		return fetchCsv(url).then(function (text) {
			rejectIfHtml(text, url);
			var lines = parseCsvLines(text);
			if (!lines.length) {
				return null;
			}
			var cells = splitDataLine(lines[0]);
			if (!cells || cells.length < 8) {
				return null;
			}
			return stripPaloCsvField(cells[7]);
		});
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

	function elementTypeLabel(typeNum) {
		if (typeNum === 1) {
			return "N";
		}
		if (typeNum === 2) {
			return "S";
		}
		if (typeNum === 4) {
			return "C";
		}
		return String(typeNum);
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
			var safeUrl = requestUrl ? redactUrlForDebug(requestUrl) : "(URL inconnue)";
			throw new Error(
				"Réponse éléments illisible (CSV Palo attendu).\n\n" +
					"URL interrogée : " +
					safeUrl +
					"\n\n" +
					"Réponse du serveur (extrait) :\n" +
					truncateForDebug(text, 900),
			);
		}
		return list;
	}

	function findElementById(id) {
		var list = state.dimensionElements;
		for (var i = 0; i < list.length; i++) {
			if (list[i].id === id) {
				return list[i];
			}
		}
		return null;
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
			show_attribute: "0",
			show_virtual_attribute: "0",
			show_info: "1",
			show_permission: "1",
		});
		var url = state.apiBase + "/database/dimensions?" + q.toString();
		return fetchCsv(url).then(function (text) {
			return parseDatabaseDimensionsList(text, url);
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

	/** Colonne 7 du CSV `/dimension/info` : id de la dimension d’attributs (#…) contenant les noms de propriétés. */
	function parseAttributesDimensionIdFromInfo(text, requestUrl) {
		rejectIfHtml(text, requestUrl);
		var lines = parseCsvLines(text);
		if (!lines.length) {
			return null;
		}
		var cells = splitDataLine(lines[0]);
		if (!cells || cells.length < 8) {
			cells = lines[0].split(";").map(function (c) {
				return c.trim();
			});
		}
		if (cells.length < 8) {
			return null;
		}
		var raw = stripPaloCsvField(cells[7]);
		if (!raw || raw === "0") {
			return null;
		}
		if (isNumericId(raw)) {
			return raw;
		}
		return null;
	}

	function loadDimensionAttributesDimensionId(nameDatabase, nameDimension) {
		var q = new URLSearchParams({
			sid: state.sid,
			name_database: nameDatabase,
			name_dimension: nameDimension,
			show_permission: "1",
			show_counters: "0",
			show_default_elements: "0",
			show_count_by_type: "0",
			show_virtual: "1",
		});
		var url = state.apiBase + "/dimension/info?" + q.toString();
		return fetchCsv(url).then(function (text) {
			return parseAttributesDimensionIdFromInfo(text, url);
		});
	}

	/** Éléments de la dimension d’attributs (propriétés affichables pour la dimension normale). */
	function loadPropertyElements(nameDatabase, attributesDimId) {
		if (!attributesDimId) {
			return Promise.resolve([]);
		}
		var q = new URLSearchParams({
			sid: state.sid,
			name_database: nameDatabase,
			dimension: attributesDimId,
			show_permission: "0",
		});
		var url = state.apiBase + "/dimension/elements?" + q.toString();
		return fetchCsv(url).then(function (text) {
			return parseDimensionElementsList(text, url);
		});
	}

	function loadDimensionElements(nameDatabase, nameDimension) {
		var q = new URLSearchParams({
			sid: state.sid,
			name_database: nameDatabase,
			name_dimension: nameDimension,
			show_permission: "1",
		});
		var url = state.apiBase + "/dimension/elements?" + q.toString();
		return fetchCsv(url).then(function (text) {
			var list = parseDimensionElementsList(text, url);
			state.dimensionElements = list;
			return list;
		});
	}

	function renderPropertyElementsList(items, attributesDimId) {
		var ul = document.getElementById("listPropertyElements");
		var empty = document.getElementById("emptyPropertyElements");
		var hint = document.getElementById("propertyHint");
		if (!ul || !empty) {
			return;
		}
		while (ul.firstChild) {
			ul.removeChild(ul.firstChild);
		}
		if (hint) {
			hint.style.display = "block";
			if (!attributesDimId) {
				hint.textContent =
					"Aucune dimension d’attributs (#…) n’est liée à cette dimension (champ attributes_dimension dans /dimension/info). Les propriétés sont les éléments de cette dimension « # ».";
				empty.style.display = "none";
				empty.textContent = "";
				return;
			}
			hint.textContent =
				"Liste des éléments de la dimension d’attributs (id " +
				attributesDimId +
				") — ce sont les noms de propriétés pour cette dimension.";
		}
		if (!items.length) {
			empty.style.display = "block";
			empty.textContent = "Aucun élément dans la dimension d’attributs (aucune propriété).";
			return;
		}
		empty.style.display = "none";
		empty.textContent = "—";
		for (var i = 0; i < items.length; i++) {
			var li = document.createElement("li");
			li.textContent = items[i].name;
			li.title = "id " + items[i].id + " — " + elementTypeLabel(items[i].type);
			ul.appendChild(li);
		}
	}

	function canManageElements() {
		return permissionAllowsWrite(state.databasePermission || "");
	}

	/** Dossier du complément (même origine que cette page), pour les pages ouvertes en dialogue Office. */
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
			setStatus("Boîte de dialogue Office indisponible (Office.js / displayDialogAsync).", "err");
			return;
		}
		var qs = new URLSearchParams(queryParams);
		var url = getAddinPageBaseUrl() + htmlFile + "?v=1.0.17.0&" + qs.toString();
		/** Nouvel objet à chaque appel : Excel sur le web peut enrichir l’objet options (ex. callback) ; le réutiliser provoque « le rappel ne peut pas être spécifié à la fois… » au 2ᵉ affichage. */
		var dialogOpts = {
			height: 90,
			width: 90,
			displayInIframe: true,
		};
		Office.context.ui.displayDialogAsync(url, dialogOpts, function (asyncResult) {
			if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
				setStatus("Impossible d’ouvrir la fenêtre Office (displayDialogAsync).", "err");
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

	function openModalAddElement() {
		if (!canManageElements() || !state.selectedDb || !state.selectedDimension) {
			return;
		}
		openOfficeDialogPage(
			"dialog-add-element.html",
			{
				sid: state.sid,
				apiBase: state.apiBase,
				name_database: state.selectedDb.name,
				name_dimension: state.selectedDimension.name,
			},
			function () {
				reloadDimensionView()
					.then(function () {
						setStatus("Élément créé.", "ok");
					})
					.catch(function (err) {
						var m = err && err.message ? err.message : String(err);
						setStatus(m, "err");
					});
			},
		);
	}

	function openModalConsolidation(el) {
		if (!canManageElements() || !state.selectedDb || !state.selectedDimension) {
			return;
		}
		openOfficeDialogPage(
			"dialog-consolidation.html",
			{
				sid: state.sid,
				apiBase: state.apiBase,
				name_database: state.selectedDb.name,
				name_dimension: state.selectedDimension.name,
				element_id: el.id,
				element_name: el.name,
				initial_children: (el.childrenIds || []).join(","),
			},
			function () {
				reloadDimensionView()
					.then(function () {
						setStatus("Consolidation enregistrée.", "ok");
					})
					.catch(function (err) {
						var m = err && err.message ? err.message : String(err);
						setStatus(m, "err");
					});
			},
		);
	}

	function onDeleteElement(el) {
		if (!state.selectedDb || !state.selectedDimension) {
			return;
		}
		if (!confirm('Supprimer l’élément « ' + el.name + ' » ?')) {
			return;
		}
		var q = new URLSearchParams({
			sid: state.sid,
			name_database: state.selectedDb.name,
			name_dimension: state.selectedDimension.name,
			element: el.id,
		});
		var url = state.apiBase + "/element/destroy?" + q.toString();
		setStatus("Suppression de l’élément…", "");
		fetchCsv(url)
			.then(function () {
				return reloadDimensionView();
			})
			.then(function () {
				setStatus("Élément supprimé.", "ok");
			})
			.catch(function (err) {
				var msg = err && err.message ? err.message : String(err);
				setStatus(msg, "err");
			});
	}

	function reloadDimensionView() {
		var db = state.selectedDb;
		var dim = state.selectedDimension;
		if (!db || !dim) {
			return Promise.reject(new Error("Pas de dimension sélectionnée."));
		}
		setStatus("Actualisation…", "");
		return Promise.all([
			loadDimensionElements(db.name, dim.name),
			loadDimensionAttributesDimensionId(db.name, dim.name),
		]).then(function (results) {
			var mainElements = results[0];
			var attrId = results[1];
			state.attributesDimensionId = attrId;
			return loadPropertyElements(db.name, attrId).then(function (propElements) {
				renderElementList(mainElements);
				renderPropertyElementsList(propElements, attrId);
				setStatus("", "");
			});
		});
	}

	function renderElementList(items) {
		var ul = document.getElementById("listElements");
		var empty = document.getElementById("emptyElements");
		var btnAdd = document.getElementById("btnAddElement");
		var can = canManageElements();
		if (btnAdd) {
			btnAdd.style.display = can ? "inline-block" : "none";
		}
		while (ul.firstChild) {
			ul.removeChild(ul.firstChild);
		}
		if (!items.length) {
			empty.style.display = "block";
			return;
		}
		empty.style.display = "none";
		for (var i = 0; i < items.length; i++) {
			(function (el) {
				var li = document.createElement("li");
				li.className = "el-row";
				var main = document.createElement("span");
				main.className = "el-main";
				main.textContent = el.name;
				var ty = document.createElement("span");
				ty.className = "el-type";
				ty.textContent = "(" + elementTypeLabel(el.type) + ")";
				main.appendChild(ty);
				li.appendChild(main);
				if (can) {
					var actions = document.createElement("span");
					actions.className = "el-actions";
					var bCons = document.createElement("button");
					bCons.type = "button";
					bCons.className = "el-consol secondary";
					bCons.textContent = "Consolidation";
					bCons.title = "Définir les enfants (consolidation)";
					bCons.addEventListener("click", function (e) {
						e.stopPropagation();
						openModalConsolidation(el);
					});
					var bDel = document.createElement("button");
					bDel.type = "button";
					bDel.className = "el-del";
					bDel.textContent = "Supprimer";
					bDel.addEventListener("click", function (e) {
						e.stopPropagation();
						onDeleteElement(el);
					});
					actions.appendChild(bCons);
					actions.appendChild(bDel);
					li.appendChild(actions);
				}
				ul.appendChild(li);
			})(items[i]);
		}
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

	function effectiveDimPermission(dim) {
		if (dim.permission && String(dim.permission).trim()) {
			return String(dim.permission).trim();
		}
		return state.databasePermission ? String(state.databasePermission).trim() : "";
	}

	function canDeleteDimensionRow(dim) {
		if (!dimensionTypeDeletable(dim.type)) {
			return false;
		}
		return permissionAllowsWrite(effectiveDimPermission(dim));
	}

	function updateDimensionManageUi() {
		var row = document.getElementById("dimAddRow");
		var hint = document.getElementById("dimRightsHint");
		var p = state.databasePermission;
		var canCreate = permissionAllowsWrite(p || "");
		row.style.display = canCreate ? "flex" : "none";
		if (!hint) {
			return;
		}
		if (p) {
			hint.style.display = "block";
			hint.textContent =
				"Droits sur cette base (API Jedox) : " +
				p +
				" — " +
				(canCreate
					? "création / suppression de dimensions autorisée si le serveur l’accepte (dimensions système non supprimables)."
					: "pas de création ni suppression (lecture seule ou droits insuffisants).");
		} else {
			hint.style.display = "none";
			hint.textContent = "";
		}
	}

	function reloadDatabaseDetail() {
		var db = state.selectedDb;
		if (!db) {
			return Promise.resolve();
		}
		setStatus("Actualisation…", "");
		return Promise.all([
			loadDatabasePermissionInfo(db.name),
			loadDimensionsForDb(db.name),
			loadCubesForDb(db.name),
		]).then(function (results) {
			state.databasePermission = results[0];
			renderLists(results[1], results[2]);
			updateDimensionManageUi();
			setStatus("", "");
		});
	}

	function createDimensionOnServer(newName) {
		var db = state.selectedDb;
		if (!db || !permissionAllowsWrite(state.databasePermission || "")) {
			return Promise.reject(new Error("Action non autorisée."));
		}
		var q = new URLSearchParams({
			sid: state.sid,
			name_database: db.name,
			new_name: newName,
			type: "0",
		});
		var url = state.apiBase + "/dimension/create?" + q.toString();
		return fetchCsv(url);
	}

	function destroyDimensionOnServer(dim) {
		var db = state.selectedDb;
		if (!db) {
			return Promise.reject(new Error("Aucune base sélectionnée."));
		}
		if (!canDeleteDimensionRow(dim)) {
			return Promise.reject(new Error("Suppression non autorisée pour cette dimension."));
		}
		var q = new URLSearchParams({
			sid: state.sid,
			name_database: db.name,
			name_dimension: dim.name,
		});
		var url = state.apiBase + "/dimension/destroy?" + q.toString();
		return fetchCsv(url);
	}

	function onCreateDimension() {
		var input = document.getElementById("inputNewDimension");
		var btn = document.getElementById("btnCreateDimension");
		var name = input ? input.value.trim() : "";
		if (!name) {
			setStatus("Indiquez un nom de dimension.", "err");
			return;
		}
		if (!permissionAllowsWrite(state.databasePermission || "")) {
			setStatus("Création impossible : droits insuffisants sur la base.", "err");
			return;
		}
		btn.disabled = true;
		setStatus("Création de la dimension…", "");
		createDimensionOnServer(name)
			.then(function () {
				if (input) {
					input.value = "";
				}
				return reloadDatabaseDetail();
			})
			.then(function () {
				setStatus("Dimension créée.", "ok");
			})
			.catch(function (err) {
				var msg = err && err.message ? err.message : String(err);
				setStatus(msg, "err");
			})
			.then(function () {
				btn.disabled = false;
			});
	}

	function onDeleteDimension(dim, ev) {
		if (ev) {
			ev.preventDefault();
			ev.stopPropagation();
		}
		if (!canDeleteDimensionRow(dim)) {
			return;
		}
		if (!confirm('Supprimer la dimension « ' + dim.name + ' » ?')) {
			return;
		}
		setStatus("Suppression…", "");
		destroyDimensionOnServer(dim)
			.then(function () {
				if (state.selectedDimension && state.selectedDimension.id === dim.id) {
					state.selectedDimension = null;
					showView("detail");
				}
				return reloadDatabaseDetail();
			})
			.then(function () {
				setStatus("Dimension supprimée.", "ok");
			})
			.catch(function (err) {
				var msg = err && err.message ? err.message : String(err);
				setStatus(msg, "err");
			});
	}

	function renderLists(dimensions, cubes) {
		function fillCubes(ulId, emptyId, items, labelField) {
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

		var ulDim = document.getElementById("listDimensions");
		var emptyDim = document.getElementById("emptyDimensions");
		while (ulDim.firstChild) {
			ulDim.removeChild(ulDim.firstChild);
		}
		if (!dimensions.length) {
			emptyDim.style.display = "block";
		} else {
			emptyDim.style.display = "none";
			for (var j = 0; j < dimensions.length; j++) {
				(function (dim) {
					var li = document.createElement("li");
					li.className = "dim-row";
					var nameSpan = document.createElement("span");
					nameSpan.className = "dim-name";
					nameSpan.textContent = dim.name;
					nameSpan.title = "id " + dim.id + " — ouvrir";
					nameSpan.addEventListener("click", function () {
						selectDimension(dim);
					});
					li.appendChild(nameSpan);
					if (canDeleteDimensionRow(dim)) {
						var delBtn = document.createElement("button");
						delBtn.type = "button";
						delBtn.className = "dim-del";
						delBtn.setAttribute("aria-label", "Supprimer la dimension " + dim.name);
						delBtn.textContent = "Supprimer";
						delBtn.addEventListener("click", function (e) {
							onDeleteDimension(dim, e);
						});
						li.appendChild(delBtn);
					}
					ulDim.appendChild(li);
				})(dimensions[j]);
			}
		}
		fillCubes("listCubes", "emptyCubes", cubes, "name");
	}

	function showView(which) {
		state.currentView = which;
		document.getElementById("viewDatabases").className = which === "databases" ? "active" : "";
		document.getElementById("viewDetail").className = which === "detail" ? "active" : "";
		document.getElementById("viewDimension").className = which === "dimension" ? "active" : "";
		var btn = document.getElementById("btnBack");
		if (which === "databases") {
			btn.style.display = "none";
		} else if (which === "detail") {
			btn.style.display = "inline-block";
			btn.textContent = "← Bases";
		} else if (which === "dimension") {
			btn.style.display = "inline-block";
			btn.textContent = "← Dimensions";
		}
	}

	function selectDatabase(db) {
		state.selectedDb = db;
		state.selectedDimension = null;
		state.databasePermission = null;
		setStatus("Chargement…", "");
		document.getElementById("detailTitle").textContent = "Base : " + db.name;
		Promise.all([
			loadDatabasePermissionInfo(db.name),
			loadDimensionsForDb(db.name),
			loadCubesForDb(db.name),
		])
			.then(function (results) {
				state.databasePermission = results[0];
				renderLists(results[1], results[2]);
				updateDimensionManageUi();
				setStatus("", "");
				showView("detail");
			})
			.catch(function (err) {
				var msg = err && err.message ? err.message : String(err);
				setStatus(msg, "err");
			});
	}

	function selectDimension(dim) {
		var db = state.selectedDb;
		if (!db) {
			return;
		}
		state.selectedDimension = dim;
		state.attributesDimensionId = null;
		setStatus("Chargement de la dimension…", "");
		document.getElementById("dimensionTitle").textContent =
			"Dimension : " + dim.name + " — base : " + db.name;
		Promise.all([
			loadDimensionElements(db.name, dim.name),
			loadDimensionAttributesDimensionId(db.name, dim.name),
		])
			.then(function (results) {
				var mainElements = results[0];
				var attrId = results[1];
				state.attributesDimensionId = attrId;
				return loadPropertyElements(db.name, attrId).then(function (propElements) {
					renderElementList(mainElements);
					renderPropertyElementsList(propElements, attrId);
					setStatus("", "");
					showView("dimension");
				});
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
				state.selectedDimension = null;
				state.databasePermission = null;
				state.dimensionElements = [];
				state.attributesDimensionId = null;
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
		if (state.currentView === "dimension") {
			state.selectedDimension = null;
			state.dimensionElements = [];
			state.attributesDimensionId = null;
			showView("detail");
			setStatus("", "");
			return;
		}
		showView("databases");
		setStatus("", "");
	}

	Office.onReady(function () {
		document.getElementById("btnRefresh").addEventListener("click", refreshAll);
		document.getElementById("btnBack").addEventListener("click", onBack);
		var btnCreate = document.getElementById("btnCreateDimension");
		var inputNew = document.getElementById("inputNewDimension");
		if (btnCreate) {
			btnCreate.addEventListener("click", onCreateDimension);
		}
		if (inputNew) {
			inputNew.addEventListener("keydown", function (ev) {
				if (ev.key === "Enter") {
					ev.preventDefault();
					onCreateDimension();
				}
			});
		}
		var btnAddEl = document.getElementById("btnAddElement");
		if (btnAddEl) {
			btnAddEl.addEventListener("click", openModalAddElement);
		}
		refreshAll();
	});
})();
