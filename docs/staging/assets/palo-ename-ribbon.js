/**
 * Parse PALO.ENAME dans une formule locale et resolution des arguments (litteral ou reference cellule).
 * Utilise par commands.js (ruban) pour le choix d'element dimension.
 */
(function paloEnameRibbonBootstrap() {
  var FUNC = "PALO.ENAME";

  /**
   * Uniformise la formule renvoyee par Excel (ruban) avant parse : trim, @ d'intersection implicite, = manquant.
   */
  function normalizeRibbonFormula(formulaLocal) {
    var s = String(formulaLocal || "").trim();
    if (!s) {
      return "";
    }
    if (s.charCodeAt(0) === 0xfeff) {
      s = s.slice(1).trim();
    }
    if (s.indexOf("=@") === 0) {
      s = ("=" + s.slice(2)).trim();
    }
    if (s.charAt(0) !== "=" && /(?:^|[^A-Z0-9_])PALO\s*\.\s*ENAME\s*\(/i.test(s)) {
      s = "=" + s;
    }
    // Guillemets typographiques / espaces insécables : sans ca le parseur de liste d'arguments casse (popup ruban silencieux).
    s = s
      .replace(/\u201c/g, "\"")
      .replace(/\u201d/g, "\"")
      .replace(/\u00ab/g, "\"")
      .replace(/\u00bb/g, "\"")
      .replace(/\u2018/g, "'")
      .replace(/\u2019/g, "'")
      .replace(/\u00a0/g, " ");
    return s;
  }

  function mentionsPaloEnameInText(s) {
    return /PALO\s*\.\s*ENAME\s*\(/i.test(s) || /_XLFN\.PALO\.ENAME/i.test(s);
  }

  /**
   * Premiere cellule de la matrice formulas* contenant PALO.ENAME (indices relatifs a la plage chargee).
   */
  function scanRangeForPaloEnameFirstHit(range) {
    var maxCells = 500;
    function scanMatrix(m) {
      if (!m || !Array.isArray(m)) {
        return null;
      }
      var r;
      var seen = 0;
      for (r = 0; r < m.length; r += 1) {
        var row = m[r];
        if (!Array.isArray(row)) {
          continue;
        }
        var c;
        for (c = 0; c < row.length; c += 1) {
          seen += 1;
          if (seen > maxCells) {
            return null;
          }
          var cellF = String(row[c] || "").trim();
          if (!cellF) {
            continue;
          }
          if (mentionsPaloEnameInText(cellF)) {
            return { formula: cellF, row: r, col: c };
          }
        }
      }
      return null;
    }
    return scanMatrix(range.formulasLocal) || scanMatrix(range.formulas);
  }

  /**
   * Parcourt formulasLocal / formulas (matrice) pour trouver une cellule contenant PALO.ENAME (selection multi-cellules).
   */
  function scanRangeForPaloEnameFormula(range) {
    var h = scanRangeForPaloEnameFirstHit(range);
    return h ? h.formula : "";
  }

  /**
   * Formule + position relative (pour getCell sur la meme plage que celle chargee).
   */
  function readFormulaBestMetaForPalo(range) {
    var one = readFormulaForPaloFromLoadedRange(range);
    if (mentionsPaloEnameInText(one)) {
      return { formula: one, row: 0, col: 0 };
    }
    var hit = scanRangeForPaloEnameFirstHit(range);
    if (hit) {
      return { formula: hit.formula, row: hit.row, col: hit.col };
    }
    return { formula: one || "", row: 0, col: 0 };
  }

  /**
   * Formule a parser : priorite au contenu qui contient vraiment PALO.ENAME (pas seulement [0][0] sur une plage).
   */
  function readFormulaBestForPalo(range) {
    return readFormulaBestMetaForPalo(range).formula;
  }

  /**
   * Choisit formulasLocal ou formulas selon celui qui contient vraiment PALO.ENAME (Excel peut ne remplir qu'un des deux).
   */
  function readFormulaForPaloFromLoadedRange(range) {
    var local = range.formulasLocal && range.formulasLocal[0] ? range.formulasLocal[0][0] : "";
    var en = range.formulas && range.formulas[0] ? range.formulas[0][0] : "";
    local = String(local || "").trim();
    en = String(en || "").trim();
    function mentionsPaloEname(s) {
      return /_XLFN\.PALO\.ENAME/i.test(s) || /PALO\s*\.\s*ENAME\s*\(/i.test(s);
    }
    if (mentionsPaloEname(local)) {
      return local;
    }
    if (mentionsPaloEname(en)) {
      return en;
    }
    return local || en;
  }

  /**
   * Position de l'appel PALO.ENAME( (supporte espaces autour du point et _xlfn.).
   */
  function findPaloEnameCallStart(formula) {
    var s = String(formula || "");
    var upper = s.toUpperCase();
    var xlfn = "_XLFN.PALO.ENAME(";
    var ix = upper.indexOf(xlfn);
    if (ix !== -1) {
      return { idx: ix, needleLen: xlfn.length };
    }
    var pos = 0;
    while (pos < s.length) {
      var hit = upper.indexOf("PALO", pos);
      if (hit === -1) {
        return null;
      }
      var rest = s.slice(hit);
      var m = /^PALO\s*\.\s*ENAME\s*\(/i.exec(rest);
      if (m) {
        return { idx: hit, needleLen: m[0].length };
      }
      pos = hit + 4;
    }
    return null;
  }

  function tryUnquoteStringLiteral(trimmed) {
    var s = String(trimmed || "");
    if (s.length >= 2 && s[0] === "\"" && s[s.length - 1] === "\"") {
      return s.slice(1, -1).replace(/""/g, "\"");
    }
    if (s.length >= 2 && s[0] === "'" && s[s.length - 1] === "'") {
      return s.slice(1, -1).replace(/''/g, "'");
    }
    return null;
  }

  function detectListSeparator(formula, innerStart, closingIdx) {
    var depth = 0;
    var inStr = false;
    var strQ = "";
    var sawSemi = false;
    var sawComma = false;
    var i;
    for (i = innerStart; i < closingIdx; i++) {
      var ch = formula[i];
      if (inStr) {
        if (ch === strQ) {
          if (ch === "\"" && formula[i + 1] === "\"") {
            i++;
            continue;
          }
          if (ch === "'" && formula[i + 1] === "'") {
            i++;
            continue;
          }
          inStr = false;
        }
        continue;
      }
      if (ch === "\"" || ch === "'") {
        inStr = true;
        strQ = ch;
        continue;
      }
      if (ch === "(") {
        depth++;
      } else if (ch === ")") {
        depth--;
      } else if (depth === 0 && ch === ";") {
        sawSemi = true;
      } else if (depth === 0 && ch === ",") {
        sawComma = true;
      }
    }
    if (sawSemi) {
      return ";";
    }
    if (sawComma) {
      return ",";
    }
    return ";";
  }

  /**
   * Coupe les arguments du premier appel FUNC(...) dans formula (indices globaux).
   * innerStart = index du premier caractere a l'interieur des parentheses ouvrantes.
   */
  function sliceTopLevelArgs(formula, innerStart, listSep) {
    var args = [];
    var segStart = innerStart;
    var depth = 0;
    var inStr = false;
    var strQ = "";
    var i;
    for (i = innerStart; i < formula.length; i++) {
      var ch = formula[i];
      if (inStr) {
        if (ch === strQ) {
          if (ch === "\"" && formula[i + 1] === "\"") {
            i++;
            continue;
          }
          if (ch === "'" && formula[i + 1] === "'") {
            i++;
            continue;
          }
          inStr = false;
        }
        continue;
      }
      if (ch === "\"" || ch === "'") {
        inStr = true;
        strQ = ch;
        continue;
      }
      if (ch === "(") {
        depth++;
      } else if (ch === ")") {
        if (depth === 0) {
          pushArg(formula, segStart, i, args);
          return { args: args, closeParenIndex: i };
        }
        depth--;
      } else if (ch === listSep && depth === 0) {
        pushArg(formula, segStart, i, args);
        segStart = i + 1;
      }
    }
    return null;
  }

  function pushArg(formula, segStart, segEnd, args) {
    var seg = formula.slice(segStart, segEnd);
    var t = seg.trim();
    var lead = t ? seg.indexOf(t) : 0;
    if (lead < 0) {
      lead = 0;
    }
    var start = segStart + lead;
    var end = t.length ? start + t.length : start;
    args.push({ raw: t, start: start, end: end });
  }

  /**
   * @returns {null | { formula: string, callStart: number, callEnd: number, args: Array<{raw:string,start:number,end:number}> }}
   */
  function parsePaloEnameCall(formulaLocal) {
    var formula = normalizeRibbonFormula(formulaLocal);
    if (!formula) {
      return null;
    }
    var found = findPaloEnameCallStart(formula);
    if (!found) {
      return null;
    }
    var idx = found.idx;
    var innerStart = idx + found.needleLen;
    var depth = 0;
    var inStr = false;
    var strQ = "";
    var j;
    for (j = innerStart; j < formula.length; j++) {
      var cj = formula[j];
      if (inStr) {
        if (cj === strQ) {
          if (cj === "\"" && formula[j + 1] === "\"") {
            j++;
            continue;
          }
          if (cj === "'" && formula[j + 1] === "'") {
            j++;
            continue;
          }
          inStr = false;
        }
        continue;
      }
      if (cj === "\"" || cj === "'") {
        inStr = true;
        strQ = cj;
        continue;
      }
      if (cj === "(") {
        depth++;
      } else if (cj === ")") {
        if (depth === 0) {
          break;
        }
        depth--;
      }
    }
    if (j >= formula.length || formula[j] !== ")") {
      return null;
    }
    var closingIdx = j;
    var listSep = detectListSeparator(formula, innerStart, closingIdx);
    var sliced = sliceTopLevelArgs(formula, innerStart, listSep);
    if (!sliced || sliced.args.length < 3) {
      return null;
    }
    return {
      formula: formula,
      callStart: idx,
      callEnd: closingIdx + 1,
      args: sliced.args
    };
  }

  function buildExcelStringLiteral(value) {
    var s = String(value == null ? "" : value);
    return "\"" + s.replace(/"/g, "\"\"") + "\"";
  }

  /**
   * Prepare la lecture d'une valeur scalaire depuis un argument de formule (litteral ou adresse).
   */
  function prepareResolveArgument(context, workbook, worksheet, rawArgText) {
    var trimmed = String(rawArgText || "").trim();
    var lit = tryUnquoteStringLiteral(trimmed);
    if (lit !== null) {
      return { kind: "literal", value: lit };
    }
    if (!trimmed) {
      return { kind: "literal", value: "" };
    }
    var addr = trimmed;
    var range;
    if (addr.indexOf("!") === -1) {
      range = worksheet.getRange(addr);
    } else {
      range = workbook.getRange(addr);
    }
    range.load(["values"]);
    return { kind: "range", range: range };
  }

  function readPrepared(prepared) {
    if (prepared.kind === "literal") {
      return prepared.value;
    }
    var vals = prepared.range.values;
    if (!vals || !vals[0] || vals[0].length === 0) {
      return "";
    }
    var v = vals[0][0];
    if (v == null || v === "") {
      return "";
    }
    return String(v);
  }

  window.paloEnameRibbon = {
    normalizeRibbonFormula: normalizeRibbonFormula,
    scanRangeForPaloEnameFormula: scanRangeForPaloEnameFormula,
    readFormulaBestMetaForPalo: readFormulaBestMetaForPalo,
    readFormulaBestForPalo: readFormulaBestForPalo,
    readFormulaForPaloFromLoadedRange: readFormulaForPaloFromLoadedRange,
    parsePaloEnameCall: parsePaloEnameCall,
    prepareResolveArgument: prepareResolveArgument,
    readPrepared: readPrepared,
    buildExcelStringLiteral: buildExcelStringLiteral,
    tryUnquoteStringLiteral: tryUnquoteStringLiteral
  };
})();
