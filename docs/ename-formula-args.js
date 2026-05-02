/* Parse des arguments de =PALO.ENAME(...) / =PALO.DATAC(...) — partagé par commands.js et action-popup.js */
(function (g) {
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

	function findPaloEnameArgumentListInner(formula) {
		var f = String(formula == null ? "" : formula).trim();
		if (f.charAt(0) === "=") {
			f = f.slice(1).trim();
		}
		f = f.replace(/^_xlfn\./i, "");
		if (!/^PALO\.ENAME\s*\(/i.test(f)) {
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
		return f.slice(start + 1, end);
	}

	function findPaloDatacArgumentListInner(formula) {
		var f = String(formula == null ? "" : formula).trim();
		if (f.charAt(0) === "=") {
			f = f.slice(1).trim();
		}
		f = f.replace(/^_xlfn\./i, "");
		if (!/^PALO\.DATAC\s*\(/i.test(f)) {
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
		return f.slice(start + 1, end);
	}

	/** Tous les arguments expression de =PALO.DATAC(db, cube, el1, …) — au moins 3. */
	function parsePaloDatacAllArgExpressions(formula) {
		var inner = findPaloDatacArgumentListInner(formula);
		if (!inner) {
			return null;
		}
		var args = parseTopLevelArgs(inner);
		if (args.length < 3) {
			return null;
		}
		return args.map(function (a) {
			return a.trim();
		});
	}

	function parsePaloEnameFirstThreeArgExpressions(formula) {
		var inner = findPaloEnameArgumentListInner(formula);
		if (!inner) {
			return null;
		}
		var args = parseTopLevelArgs(inner);
		if (args.length < 3) {
			return null;
		}
		return [args[0].trim(), args[1].trim(), args[2].trim()];
	}

	function tryExcelFormulaStringLiteral(expr) {
		var t = String(expr == null ? "" : expr).trim();
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

	g.parsePaloEnameFirstThreeArgExpressions = parsePaloEnameFirstThreeArgExpressions;
	g.parsePaloDatacAllArgExpressions = parsePaloDatacAllArgExpressions;
	g.tryExcelFormulaStringLiteral = tryExcelFormulaStringLiteral;
})(typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : this);
