/* Réponses CSV des opérations HTTP Palo (create / replace / …) : erreurs souvent renvoyées avec HTTP 200. */
(function () {
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

	/**
	 * Lève une Error si le corps de réponse indique un échec côté serveur.
	 * Aligné sur parsePaloStatus dans functions.js.
	 */
	function assertPaloCsvMutationOk(text) {
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

	if (typeof window !== "undefined") {
		window.assertPaloCsvMutationOk = assertPaloCsvMutationOk;
	}
})();
