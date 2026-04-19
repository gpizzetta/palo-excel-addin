/* Dialog « plein écran » (fenêtre Office) — consolidation d’éléments. */
(function () {
	var ctx = {
		apiBase: "",
		sid: "",
		name_database: "",
		name_dimension: "",
		element_id: "",
		element_name: "",
		selectedChildIds: [],
		dimensionElements: [],
		targetEl: null,
	};

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

	function rejectIfHtml(text) {
		var t = stripBom(text).trim();
		if (!t.length) {
			return;
		}
		var lower = t.slice(0, 200).toLowerCase();
		if (
			t.charAt(0) === "<" ||
			lower.indexOf("<!doctype") !== -1 ||
			lower.indexOf("<html") !== -1
		) {
			throw new Error("Réponse HTML au lieu du CSV Palo.");
		}
	}

	function parseDimensionElementsList(text) {
		rejectIfHtml(text);
		var lines = parseCsvLines(text);
		var list = [];
		for (var i = 0; i < lines.length; i++) {
			var cells = splitDataLine(lines[i]);
			if (!cells || cells.length < 7) {
				continue;
			}
			var id = cells[0];
			var name = stripPaloCsvField(cells[1]);
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
			parseWeightsForChildren(weightsRaw, childrenIds.length);
			list.push({
				id: id,
				name: name,
				type: typeNum,
				childrenIds: childrenIds,
			});
		}
		return list;
	}

	function findElementById(id) {
		for (var i = 0; i < ctx.dimensionElements.length; i++) {
			if (ctx.dimensionElements[i].id === id) {
				return ctx.dimensionElements[i];
			}
		}
		return null;
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
					throw new Error("HTTP " + res.status + " — " + text.slice(0, 400));
				}
				return text;
			});
		});
	}

	function showErr(msg) {
		var el = document.getElementById("err");
		if (el) {
			el.textContent = msg || "Erreur";
			el.className = "show";
		}
	}

	function parseQuery() {
		var p = new URLSearchParams(window.location.search);
		ctx.apiBase = (p.get("apiBase") || "").trim();
		ctx.sid = p.get("sid") || "";
		ctx.name_database = p.get("name_database") || "";
		ctx.name_dimension = p.get("name_dimension") || "";
		ctx.element_id = p.get("element_id") || "";
		ctx.element_name = p.get("element_name") || "";
		var kids = p.get("initial_children") || "";
		ctx.selectedChildIds = kids
			? kids.split(",").map(function (x) {
					return x.trim();
			  })
			: [];
		ctx.selectedChildIds = ctx.selectedChildIds.filter(isNumericId);
	}

	function loadElements() {
		var q = new URLSearchParams({
			sid: ctx.sid,
			name_database: ctx.name_database,
			name_dimension: ctx.name_dimension,
			show_permission: "1",
		});
		var url = ctx.apiBase + "/dimension/elements?" + q.toString();
		return fetchCsv(url).then(function (text) {
			ctx.dimensionElements = parseDimensionElementsList(text);
			ctx.targetEl = findElementById(ctx.element_id);
			if (!ctx.targetEl) {
				throw new Error("Élément introuvable (id " + ctx.element_id + ").");
			}
			if (!ctx.selectedChildIds.length && ctx.targetEl.childrenIds && ctx.targetEl.childrenIds.length) {
				ctx.selectedChildIds = ctx.targetEl.childrenIds.slice();
			}
		});
	}

	function refreshUi() {
		var target = ctx.targetEl;
		if (!target) {
			return;
		}
		var selectedSet = {};
		for (var s = 0; s < ctx.selectedChildIds.length; s++) {
			selectedSet[ctx.selectedChildIds[s]] = true;
		}
		var avail = document.getElementById("consolAvail");
		if (avail) {
			avail.innerHTML = "";
			for (var j = 0; j < ctx.dimensionElements.length; j++) {
				var e = ctx.dimensionElements[j];
				if (e.id === target.id) {
					continue;
				}
				if (selectedSet[e.id]) {
					continue;
				}
				var opt = document.createElement("option");
				opt.value = e.id;
				opt.textContent = e.name + " (" + elementTypeLabel(e.type) + ", id " + e.id + ")";
				avail.appendChild(opt);
			}
		}
		var ul = document.getElementById("consolSelectedList");
		if (!ul) {
			return;
		}
		ul.innerHTML = "";
		for (var k = 0; k < ctx.selectedChildIds.length; k++) {
			var cid = ctx.selectedChildIds[k];
			var child = findElementById(cid);
			var li = document.createElement("li");
			var span = document.createElement("span");
			span.textContent = child ? child.name + " (id " + cid + ")" : "id " + cid;
			var btn = document.createElement("button");
			btn.type = "button";
			btn.className = "secondary";
			btn.textContent = "Retirer";
			(function (id) {
				btn.addEventListener("click", function () {
					ctx.selectedChildIds = ctx.selectedChildIds.filter(function (x) {
						return x !== id;
					});
					refreshUi();
				});
			})(cid);
			li.appendChild(span);
			li.appendChild(btn);
			ul.appendChild(li);
		}
	}

	function consolAddSelected() {
		var avail = document.getElementById("consolAvail");
		if (!avail) {
			return;
		}
		var opts = avail.selectedOptions;
		if (!opts || !opts.length) {
			return;
		}
		for (var i = 0; i < opts.length; i++) {
			var id = opts[i].value;
			if (ctx.selectedChildIds.indexOf(id) === -1) {
				ctx.selectedChildIds.push(id);
			}
		}
		refreshUi();
	}

	function closeDialog() {
		if (Office && Office.context && Office.context.ui && Office.context.ui.messageParent) {
			Office.context.ui.messageParent("cancel", { targetOrigin: "*" });
		}
	}

	function saveAndClose() {
		if (!ctx.targetEl) {
			return;
		}
		if (ctx.selectedChildIds.length === 0) {
			showErr("Ajoutez au moins un enfant pour une consolidation.");
			return;
		}
		var weights = ctx.selectedChildIds.map(function () {
			return 1;
		});
		var q = new URLSearchParams({
			sid: ctx.sid,
			name_database: ctx.name_database,
			name_dimension: ctx.name_dimension,
			element: ctx.targetEl.id,
			type: "4",
			children: ctx.selectedChildIds.join(","),
			weights: weights.join(","),
		});
		var url = ctx.apiBase + "/element/replace?" + q.toString();
		document.getElementById("btnSave").disabled = true;
		fetchCsv(url)
			.then(function () {
				if (Office && Office.context && Office.context.ui && Office.context.ui.messageParent) {
					Office.context.ui.messageParent("refresh", { targetOrigin: "*" });
				}
			})
			.catch(function (err) {
				showErr(err && err.message ? err.message : String(err));
				document.getElementById("btnSave").disabled = false;
			});
	}

	function init() {
		parseQuery();
		loadElements()
			.then(function () {
				var sub = document.getElementById("subtitle");
				if (sub && ctx.targetEl) {
					sub.textContent =
						"Élément : " +
						(ctx.element_name || ctx.targetEl.name) +
						" (id " +
						ctx.element_id +
						") — type actuel : " +
						elementTypeLabel(ctx.targetEl.type);
				}
				refreshUi();
			})
			.catch(function (err) {
				showErr(err && err.message ? err.message : String(err));
			});
	}

	Office.onReady(function () {
		document.getElementById("consolBtnAdd").addEventListener("click", consolAddSelected);
		document.getElementById("consolAvail").addEventListener("dblclick", consolAddSelected);
		document.getElementById("btnCancel").addEventListener("click", closeDialog);
		document.getElementById("btnSave").addEventListener("click", saveAndClose);
		init();
	});
})();
