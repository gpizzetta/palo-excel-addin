(function () {
	var ctx = {
		apiBase: "",
		sid: "",
		name_database: "",
		allDimensions: [],
		selectedNames: [],
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

	function parseDatabaseDimensionsList(text) {
		rejectIfHtml(text);
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
			var typeStr = cells.length > 6 ? stripPaloCsvField(cells[6]) : "";
			var typeNum = typeStr === "" ? null : parseInt(typeStr, 10);
			if (typeNum !== null && !isNaN(typeNum) && (typeNum === 2 || typeNum === 5)) {
				continue;
			}
			list.push({ id: id, name: name, type: typeNum });
		}
		return list;
	}

	function parseQuery() {
		var p = new URLSearchParams(window.location.search);
		ctx.apiBase = (p.get("apiBase") || "").trim();
		ctx.sid = p.get("sid") || "";
		ctx.name_database = p.get("name_database") || "";
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

	function selectedSet() {
		var o = {};
		for (var i = 0; i < ctx.selectedNames.length; i++) {
			o[ctx.selectedNames[i]] = true;
		}
		return o;
	}

	function refreshUi() {
		var sel = selectedSet();
		var avail = document.getElementById("cubeAvail");
		if (avail) {
			avail.innerHTML = "";
			for (var j = 0; j < ctx.allDimensions.length; j++) {
				var d = ctx.allDimensions[j];
				if (sel[d.name]) {
					continue;
				}
				var opt = document.createElement("option");
				opt.value = d.name;
				opt.textContent = d.name + " (id " + d.id + ")";
				avail.appendChild(opt);
			}
		}
		var ul = document.getElementById("cubeSelectedList");
		if (!ul) {
			return;
		}
		ul.innerHTML = "";
		for (var k = 0; k < ctx.selectedNames.length; k++) {
			var nm = ctx.selectedNames[k];
			var li = document.createElement("li");
			var span = document.createElement("span");
			span.textContent = nm;
			var btn = document.createElement("button");
			btn.type = "button";
			btn.className = "secondary";
			btn.textContent = "Retirer";
			(function (name) {
				btn.addEventListener("click", function () {
					ctx.selectedNames = ctx.selectedNames.filter(function (x) {
						return x !== name;
					});
					refreshUi();
				});
			})(nm);
			li.appendChild(span);
			li.appendChild(btn);
			ul.appendChild(li);
		}
	}

	function addSelected() {
		var avail = document.getElementById("cubeAvail");
		if (!avail) {
			return;
		}
		var opts = avail.selectedOptions;
		if (!opts || !opts.length) {
			return;
		}
		for (var i = 0; i < opts.length; i++) {
			var nm = opts[i].value;
			if (ctx.selectedNames.indexOf(nm) === -1) {
				ctx.selectedNames.push(nm);
			}
		}
		refreshUi();
	}

	function loadDimensions() {
		var q = new URLSearchParams({
			sid: ctx.sid,
			name_database: ctx.name_database,
			show_system: "1",
			show_normal: "1",
			show_attribute: "0",
			show_virtual_attribute: "0",
			show_info: "1",
			show_permission: "1",
		});
		var url = ctx.apiBase + "/database/dimensions?" + q.toString();
		return fetchCsv(url).then(function (text) {
			ctx.allDimensions = parseDatabaseDimensionsList(text);
		});
	}

	function closeDialog() {
		if (Office && Office.context && Office.context.ui && Office.context.ui.messageParent) {
			Office.context.ui.messageParent("cancel", { targetOrigin: "*" });
		}
	}

	function saveAndClose() {
		var cubeName = document.getElementById("inputCubeName").value.trim();
		if (!cubeName) {
			showErr("Indiquez un nom de cube.");
			return;
		}
		if (!ctx.selectedNames.length) {
			showErr("Ajoutez au moins une dimension au cube.");
			return;
		}
		var q = new URLSearchParams({
			sid: ctx.sid,
			name_database: ctx.name_database,
			new_name: cubeName,
			name_dimensions: ctx.selectedNames.join(","),
		});
		var url = ctx.apiBase + "/cube/create?" + q.toString();
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
		var sub = document.getElementById("subtitle");
		if (sub) {
			sub.textContent = ctx.name_database ? "Base : " + ctx.name_database : "";
		}
		loadDimensions()
			.then(function () {
				refreshUi();
			})
			.catch(function (err) {
				showErr(err && err.message ? err.message : String(err));
			});
	}

	Office.onReady(function () {
		document.getElementById("cubeBtnAdd").addEventListener("click", addSelected);
		document.getElementById("cubeAvail").addEventListener("dblclick", addSelected);
		document.getElementById("btnCancel").addEventListener("click", closeDialog);
		document.getElementById("btnSave").addEventListener("click", saveAndClose);
		init();
	});
})();
