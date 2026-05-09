(function paloApiBootstrap() {
  var PALO_TRACE_STORAGE_KEY = "palo.office365.trace.v1";
  var PALO_TRACE_MAX_ENTRIES = 300;

  function paloTraceEnabled() {
    return Boolean(window.PALO_TRACE || window.PALO_DEBUG);
  }

  function paloTraceConsole() {
    try {
      if (typeof window !== "undefined" && window.top && window.top.console) {
        return window.top.console;
      }
    } catch (_e) {
      // Ignorer les erreurs cross-origin.
    }
    return console;
  }

  function paloTrace(eventName, payload) {
    if (!paloTraceEnabled()) {
      return;
    }
    var entry = {
      ts: new Date().toISOString(),
      event: String(eventName || "trace"),
      payload: payload || {}
    };
    try {
      var targetConsole = paloTraceConsole();
      if (targetConsole && typeof targetConsole.log === "function") {
        targetConsole.log("[PaloTrace]", entry);
      }
    } catch (_consoleError) {
      // Ne jamais casser le flux applicatif pour un log.
    }
    try {
      var raw = window.localStorage.getItem(PALO_TRACE_STORAGE_KEY);
      var history = [];
      if (raw) {
        history = JSON.parse(raw);
      }
      if (!Array.isArray(history)) {
        history = [];
      }
      history.push(entry);
      if (history.length > PALO_TRACE_MAX_ENTRIES) {
        history = history.slice(history.length - PALO_TRACE_MAX_ENTRIES);
      }
      window.localStorage.setItem(PALO_TRACE_STORAGE_KEY, JSON.stringify(history));
    } catch (_storageError) {
      // Stockage best effort.
    }
  }

  function paloGetTraceHistory() {
    try {
      var raw = window.localStorage.getItem(PALO_TRACE_STORAGE_KEY);
      if (!raw) {
        return [];
      }
      var parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (_e) {
      return [];
    }
  }

  /**
   * Persistance localStorage avec retry si quota (souvent du aux traces Palo).
   */
  function paloLocalStorageSetItem(key, value) {
    try {
      window.localStorage.setItem(key, value);
      return;
    } catch (e) {
      var isQuota =
        e &&
        (e.name === "QuotaExceededError" ||
          e.name === "NS_ERROR_DOM_QUOTA_REACHED" ||
          e.code === 22 ||
          e.number === -2147024882);
      if (isQuota) {
        try {
          window.localStorage.removeItem(PALO_TRACE_STORAGE_KEY);
        } catch (_e2) {
        }
        try {
          window.localStorage.setItem(key, value);
          return;
        } catch (_e3) {
        }
      }
      throw e;
    }
  }

  function paloLocalStorageRemoveItem(key) {
    try {
      window.localStorage.removeItem(key);
    } catch (_e) {
    }
  }

  function paloSetLastApiUrl(url) {
    window.PaloOffice = window.PaloOffice || {};
    window.PaloOffice._lastApiUrl = String(url || "");
  }

  function paloGetLastApiUrl() {
    if (!window.PaloOffice || !window.PaloOffice._lastApiUrl) {
      return "";
    }
    return String(window.PaloOffice._lastApiUrl);
  }

  function leftRotate(x, c) {
    return (x << c) | (x >>> (32 - c));
  }

  function toUtf8Bytes(input) {
    return Array.from(new TextEncoder().encode(input));
  }

  function toWordArrayLittleEndian(bytes) {
    var out = [];
    var i;
    for (i = 0; i < bytes.length; i += 4) {
      out.push(
        (bytes[i] || 0)
        | ((bytes[i + 1] || 0) << 8)
        | ((bytes[i + 2] || 0) << 16)
        | ((bytes[i + 3] || 0) << 24)
      );
    }
    return out;
  }

  function md5(input) {
    var s = [
      7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22,
      5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20,
      4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23,
      6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21
    ];
    var k = Array.from({ length: 64 }, function (_, i) {
      return Math.floor(Math.abs(Math.sin(i + 1)) * Math.pow(2, 32)) >>> 0;
    });

    var bytes = toUtf8Bytes(input);
    var originalByteLength = bytes.length;
    var bitLenLo = (originalByteLength * 8) >>> 0;
    var bitLenHi = Math.floor((originalByteLength * 8) / 0x100000000) >>> 0;
    bytes.push(0x80);
    while ((bytes.length % 64) !== 56) {
      bytes.push(0);
    }
    bytes.push(bitLenLo & 0xff, (bitLenLo >>> 8) & 0xff, (bitLenLo >>> 16) & 0xff, (bitLenLo >>> 24) & 0xff);
    bytes.push(bitLenHi & 0xff, (bitLenHi >>> 8) & 0xff, (bitLenHi >>> 16) & 0xff, (bitLenHi >>> 24) & 0xff);

    var a0 = 0x67452301;
    var b0 = 0xefcdab89;
    var c0 = 0x98badcfe;
    var d0 = 0x10325476;
    var offset;
    for (offset = 0; offset < bytes.length; offset += 64) {
      var chunk = bytes.slice(offset, offset + 64);
      var m = toWordArrayLittleEndian(chunk);
      var a = a0;
      var b = b0;
      var c = c0;
      var d = d0;
      var i;
      for (i = 0; i < 64; i += 1) {
        var f;
        var g;
        if (i < 16) {
          f = (b & c) | (~b & d);
          g = i;
        } else if (i < 32) {
          f = (d & b) | (~d & c);
          g = (5 * i + 1) % 16;
        } else if (i < 48) {
          f = b ^ c ^ d;
          g = (3 * i + 5) % 16;
        } else {
          f = c ^ (b | ~d);
          g = (7 * i) % 16;
        }
        var temp = d;
        d = c;
        c = b;
        b = (b + leftRotate((a + f + k[i] + m[g]) >>> 0, s[i])) >>> 0;
        a = temp;
      }
      a0 = (a0 + a) >>> 0;
      b0 = (b0 + b) >>> 0;
      c0 = (c0 + c) >>> 0;
      d0 = (d0 + d) >>> 0;
    }

    var words = [a0, b0, c0, d0];
    return words
      .flatMap(function (w) { return [w & 0xff, (w >>> 8) & 0xff, (w >>> 16) & 0xff, (w >>> 24) & 0xff]; })
      .map(function (b) { return b.toString(16).padStart(2, "0"); })
      .join("");
  }

  function normalizePasswordForPalo(password) {
    var candidate = String(password || "").trim();
    if (/^[a-fA-F0-9]{32}$/.test(candidate)) {
      return candidate.toLowerCase();
    }
    return md5(String(password || ""));
  }

  function splitSemicolonLine(line) {
    var out = [];
    var cur = "";
    var inQuotes = false;
    var i;
    for (i = 0; i < line.length; i += 1) {
      var ch = line[i];
      if (ch === "\"") {
        inQuotes = !inQuotes;
        continue;
      }
      if (ch === ";" && !inQuotes) {
        out.push(cur);
        cur = "";
        continue;
      }
      cur += ch;
    }
    out.push(cur);
    return out;
  }

  function parseCsvIds(value) {
    if (!value) {
      return [];
    }
    return value.split(",").map(function (v) { return v.trim(); }).filter(Boolean);
  }

  function parseCsvNumbers(value) {
    if (!value) {
      return [];
    }
    return value.split(",").map(function (v) { return Number(v); }).filter(function (v) { return !Number.isNaN(v); });
  }

  function parseServDb(servdb) {
    var idx = String(servdb).indexOf("/");
    if (idx <= 0 || idx >= String(servdb).length - 1) {
      throw new Error("servdb invalide (" + servdb + "), format attendu: Connection/Database");
    }
    return {
      connectionName: String(servdb).slice(0, idx),
      database: String(servdb).slice(idx + 1)
    };
  }

  /**
   * PALO_DEBUG = true : chemins, etapes resolve, apercu des reponses HTTP.
   * PALO_LOG_HTTP = true : traces getValidSid (cache / login).
   * Chaque requete Palo : console.log de l'URL complete (query incluse ; password masque par ***).
   */
  function paloDebugEnabled() {
    return Boolean(window.PALO_DEBUG);
  }

  function paloHttpLogEnabled() {
    return Boolean(window.PALO_DEBUG || window.PALO_LOG_HTTP);
  }

  function paloBulkTraceEnabled() {
    return Boolean(window.PALO_DEBUG || window.PALO_BULK_TRACE);
  }

  function paloRedactUrlForLog(urlString) {
    try {
      var u = new URL(urlString);
      if (u.searchParams.has("password")) {
        u.searchParams.set("password", "***");
      }
      return u.toString();
    } catch (_e) {
      return urlString;
    }
  }

  function paloSnapshotArgForLog(value) {
    var snap = {
      typeof: typeof value,
      isArray: Array.isArray(value),
      ctor: value && value.constructor && value.constructor.name ? value.constructor.name : null
    };
    if (value !== null && value !== undefined && typeof value === "object" && !Array.isArray(value)) {
      try {
        snap.keys = Object.keys(value);
      } catch (_k) {
        snap.keys = "(keys inaccessibles)";
      }
    }
    try {
      snap.json = JSON.stringify(value);
    } catch (_e) {
      snap.json = "(non serialisable en JSON)";
    }
    return snap;
  }

  function paloLogCoordinateDebug(message, value, extra) {
    var payload = { message: message, arg: paloSnapshotArgForLog(value) };
    if (extra) {
      var k;
      for (k in extra) {
        if (Object.prototype.hasOwnProperty.call(extra, k)) {
          payload[k] = extra[k];
        }
      }
    }
    console.warn("[PaloOffice]", payload);
  }

  function paloPageHrefForError() {
    try {
      if (typeof window !== "undefined" && window.location && window.location.href) {
        return window.location.href;
      }
    } catch (_e) {
      // ignorer
    }
    return "(page URL indisponible)";
  }

  /**
   * Excel envoie souvent une reference de cellule comme matrice [["texte"]] ou un wrapper objet.
   * Les nombres (ex. annee 2025) peuvent arriver en number, en objet riche (basicValue) ou en Number boite.
   * Palo attend des segments de chemin en string ; on normalise toujours vers une string en fin de traitement.
   */
  function normalizePaloPathSegment(value, debugFrom) {
    if (value === null || value === undefined) {
      return "";
    }
    var t = typeof value;
    if (t === "string") {
      return value;
    }
    if (t === "number" || t === "boolean") {
      return String(value);
    }
    if (t === "bigint") {
      return String(value);
    }
    if (Array.isArray(value)) {
      if (value.length === 0) {
        return "";
      }
      var first = value[0];
      if (Array.isArray(first)) {
        return normalizePaloPathSegment(first.length ? first[0] : "", debugFrom);
      }
      return normalizePaloPathSegment(first, debugFrom);
    }
    if (t === "object") {
      if (value.text !== undefined && value.text !== null) {
        return String(value.text);
      }
      if (value.value !== undefined) {
        return normalizePaloPathSegment(value.value, debugFrom);
      }
      if (value.basicValue !== undefined && value.basicValue !== null) {
        return normalizePaloPathSegment(value.basicValue, debugFrom);
      }
      if (typeof value.valueOf === "function") {
        var prim = value.valueOf();
        if (prim !== value && (typeof prim === "string" || typeof prim === "number" || typeof prim === "boolean" || typeof prim === "bigint")) {
          return String(prim);
        }
      }
      paloLogCoordinateDebug("Coordonnee Palo: objet sans .text ni .value utilisables", value, debugFrom);
      var keysStr = "";
      try {
        keysStr = Object.keys(value).join(",");
      } catch (_keys) {
        keysStr = "?";
      }
      var segPart = debugFrom && debugFrom.segmentIndex !== undefined
        ? " segmentIndex=" + debugFrom.segmentIndex
        : "";
      throw new Error(
        "Coordonnee Palo inutilisable "+JSON.stringify(value)+". Reference une cellule a une seule valeur ou saisis du texte."
        + " url=" + paloPageHrefForError()
        + segPart
        + " keys=" + keysStr
      );
    }
    return String(value);
  }

  function normalizePaloCellPath(path) {
    var out = [];
    var i;
    for (i = 0; i < path.length; i += 1) {
      if (paloDebugEnabled()) {
        paloLogCoordinateDebug("normalizePaloCellPath segment entrant", path[i], { segmentIndex: i, pathLength: path.length });
      }
      try {
        out.push(String(normalizePaloPathSegment(path[i], { segmentIndex: i, pathLength: path.length })).trim());
      } catch (err) {
        paloLogCoordinateDebug("normalizePaloCellPath echec sur segment", path[i], {
          segmentIndex: i,
          pathLength: path.length,
          errorMessage: err && err.message ? err.message : String(err)
        });
        throw err;
      }
    }
    return out;
  }

  function normalizePaloPathSegmentsInput(pathSegments) {
    var input = Array.isArray(pathSegments) ? pathSegments.slice() : [pathSegments];
    var i;
    var allSingleCellArrays = input.length > 0;
    if (input.length === 1 && Array.isArray(input[0])) {
      return input[0].slice();
    }
    for (i = 0; i < input.length; i += 1) {
      if (!Array.isArray(input[i]) || input[i].length !== 1) {
        allSingleCellArrays = false;
        break;
      }
    }
    if (allSingleCellArrays) {
      return input.map(function (cell) { return cell[0]; });
    }
    return input;
  }

  var PALO_DEFAULT_DIRECT_BASE = "https://palo.berdoz.local";

  function resolvePaloDirectBaseUrl(profile) {
    if (window.PALO_DIRECT_BASE_URL) {
      return String(window.PALO_DIRECT_BASE_URL).replace(/\/+$/, "");
    }
    var raw = String(profile && profile.baseUrl ? profile.baseUrl : "").trim();
    var normalized = raw.replace(/\/+$/, "");
    if (!normalized) {
      normalized = PALO_DEFAULT_DIRECT_BASE;
    }
    return normalized;
  }

  function paloRequestTimeoutMs() {
    if (typeof window !== "undefined" && window.PALO_REQUEST_TIMEOUT_MS != null) {
      var n = Number(window.PALO_REQUEST_TIMEOUT_MS);
      if (!Number.isNaN(n) && n >= 1000) {
        return Math.floor(n);
      }
    }
    return 45000;
  }

  function paloHttpMaxConcurrent() {
    if (typeof window !== "undefined" && window.PALO_HTTP_MAX_CONCURRENT != null) {
      var n = Number(window.PALO_HTTP_MAX_CONCURRENT);
      if (!Number.isNaN(n) && n >= 1) {
        return Math.min(32, Math.floor(n));
      }
    }
    return 8;
  }

  function paloHttpRetryCount() {
    if (typeof window !== "undefined" && window.PALO_HTTP_RETRY_COUNT != null) {
      var n = Number(window.PALO_HTTP_RETRY_COUNT);
      if (!Number.isNaN(n) && n >= 0) {
        return Math.min(5, Math.floor(n));
      }
    }
    return 2;
  }

  function paloHttpRetryDelayMs() {
    if (typeof window !== "undefined" && window.PALO_HTTP_RETRY_DELAY_MS != null) {
      var n = Number(window.PALO_HTTP_RETRY_DELAY_MS);
      if (!Number.isNaN(n) && n >= 0) {
        return Math.floor(n);
      }
    }
    return 280;
  }

  function paloSleepMs(ms) {
    return new Promise(function (resolve) {
      setTimeout(resolve, ms);
    });
  }

  var paloHttpGate = (function createPaloHttpGate() {
    var active = 0;
    var waiters = [];
    function pump() {
      var max = paloHttpMaxConcurrent();
      while (active < max && waiters.length > 0) {
        var w = waiters.shift();
        if (!w) {
          break;
        }
        active += 1;
        w.fn()
          .then(function (v) {
            active -= 1;
            w.resolve(v);
            pump();
          })
          .catch(function (e) {
            active -= 1;
            w.reject(e);
            pump();
          });
      }
    }
    return function runGated(fn) {
      return new Promise(function (resolve, reject) {
        waiters.push({ fn: fn, resolve: resolve, reject: reject });
        pump();
      });
    };
  })();

  function paloErrorIsRetriable(message) {
    var m = String(message || "");
    return (
      m.indexOf("Timeout HTTP") !== -1
      || m.indexOf("Impossible de joindre") !== -1
      || m.indexOf("Failed to fetch") !== -1
      || m.indexOf("NetworkError") !== -1
      || m.indexOf("Load failed") !== -1
    );
  }

  /**
   * Appels Palo depuis le navigateur vers profile.baseUrl (defaut https://palo.berdoz.local). CORS a regler cote serveur Palo.
   */
  function PaloApiClient(profile) {
    this.profile = profile;
  }

  PaloApiClient.prototype.buildUrl = function buildUrl(path, params) {
    var base = resolvePaloDirectBaseUrl(this.profile);
    var url = new URL(base + path);
    Object.keys(params).forEach(function (key) {
      var value = params[key];
      if (value !== undefined && value !== null) {
        url.searchParams.set(key, String(value));
      }
    });
    return url.toString();
  };

  PaloApiClient.prototype.call = async function call(path, params) {
    var self = this;
    var retries = paloHttpRetryCount();
    var delayMs = paloHttpRetryDelayMs();
    var attempt;
    var lastErr;
    for (attempt = 0; attempt <= retries; attempt += 1) {
      try {
        return await paloHttpGate(function () {
          return self.callOnce(path, params);
        });
      } catch (err) {
        lastErr = err;
        var msg = err && err.message ? err.message : String(err);
        if (attempt < retries && paloErrorIsRetriable(msg)) {
          paloTrace("api-retry", { path: path, attempt: attempt + 1, delayMs: delayMs, detail: msg });
          await paloSleepMs(delayMs * (attempt + 1));
          continue;
        }
        throw err;
      }
    }
    throw lastErr;
  };

  PaloApiClient.prototype.callOnce = async function callOnce(path, params) {
    var url = this.buildUrl(path, params);
    var logUrl = paloRedactUrlForLog(url);
    paloSetLastApiUrl(logUrl);
    paloTrace("api-call", { path: path, url: logUrl });
    var response;
    var timeoutMs = paloRequestTimeoutMs();
    var controller = typeof AbortController !== "undefined" ? new AbortController() : null;
    var timeoutId = null;
    try {
      if (controller) {
        timeoutId = setTimeout(function () {
          try {
            controller.abort();
          } catch (_abortErr) {
            // ignore
          }
        }, timeoutMs);
      }
      response = await fetch(url, {
        method: "GET",
        signal: controller ? controller.signal : undefined
      });
    } catch (error) {
      if (controller && controller.signal && controller.signal.aborted) {
        throw new Error("Timeout HTTP (" + timeoutMs + " ms) sur " + logUrl);
      }
      throw new Error(
        "Impossible de joindre " + logUrl + ". CORS, URL Palo, certificat HTTPS ou reseau. Detail: "
        + (error && error.message ? error.message : String(error))
      );
    } finally {
      if (timeoutId) {
        clearTimeout(timeoutId);
      }
    }
    if (!response.ok) {
      throw new Error("HTTP " + response.status + " sur " + logUrl);
    }
    var text = await response.text();
    if (paloDebugEnabled()) {
      console.info("[PaloOffice HTTP] corps", {
        path: path,
        octets: text.length,
        debut: text.slice(0, 240)
      });
    }
    return text;
  };

  PaloApiClient.prototype.login = async function login() {
    var text = await this.call("/server/login", {
      user: this.profile.user,
      password: normalizePasswordForPalo(this.profile.password)
    });
    var line = String(text).trim().split("\n")[0];
    var parts = line.split(";");
    if (!parts[0]) {
      throw new Error("Echec login: sid manquant.");
    }
    return {
      sid: parts[0],
      ttl: Number(parts[1] || "3600") || 3600
    };
  };

  PaloApiClient.prototype.logout = async function logout(sid) {
    await this.call("/server/logout", { sid: sid });
  };

  PaloApiClient.prototype.serverDatabases = async function serverDatabases(sid) {
    var text = await this.call("/server/databases", { sid: sid, show_normal: 1, show_system: 0, show_user_info: 0 });
    return text.split("\n").map(function (line) { return line.trim(); }).filter(Boolean).map(function (line) {
      var p = splitSemicolonLine(line);
      return { id: p[0], name: p[1], type: Number(p[5] || "0") };
    });
  };

  PaloApiClient.prototype.databaseDimensions = async function databaseDimensions(sid, database) {
    var text = await this.call("/database/dimensions", {
      sid: sid,
      name_database: database,
      show_normal: 1,
      show_system: 0,
      show_attribute: 0,
      show_info: 0
    });
    return text.split("\n").map(function (line) { return line.trim(); }).filter(Boolean).map(function (line) {
      var p = splitSemicolonLine(line);
      return { id: p[0], name: p[1], numberElements: Number(p[2] || "0"), type: Number(p[6] || "0") };
    });
  };

  PaloApiClient.prototype.databaseCubes = async function databaseCubes(sid, database) {
    var text = await this.call("/database/cubes", {
      sid: sid,
      name_database: database,
      show_normal: 1,
      show_system: 0,
      show_attribute: 0,
      show_info: 0
    });
    return text.split("\n").map(function (line) { return line.trim(); }).filter(Boolean).map(function (line) {
      var p = splitSemicolonLine(line);
      return { id: p[0], name: p[1], dimensionIds: parseCsvIds(p[3] || ""), type: Number(p[7] || "0") };
    });
  };

  PaloApiClient.prototype.cubeInfo = async function cubeInfo(sid, database, cube) {
    var text = await this.call("/cube/info", { sid: sid, name_database: database, name_cube: cube });
    var p = splitSemicolonLine(String(text).trim().split("\n")[0]);
    return { id: p[0], name: p[1], dimensionIds: parseCsvIds(p[3] || ""), type: Number(p[7] || "0") };
  };

  PaloApiClient.prototype.dimensionCubes = async function dimensionCubes(sid, database, dimension) {
    var text = await this.call("/dimension/cubes", {
      sid: sid,
      name_database: database,
      name_dimension: dimension,
      show_normal: 1,
      show_system: 0,
      show_attribute: 0,
      show_info: 0
    });
    return text.split("\n").map(function (line) { return line.trim(); }).filter(Boolean).map(function (line) {
      var p = splitSemicolonLine(line);
      return { id: p[0], name: p[1], dimensionIds: parseCsvIds(p[3] || ""), type: Number(p[7] || "0") };
    });
  };

  PaloApiClient.prototype.dimensionElements = async function dimensionElements(sid, database, dimension) {
    var text = await this.call("/dimension/elements", { sid: sid, name_database: database, name_dimension: dimension });
    return text.split("\n").map(function (line) { return line.trim(); }).filter(Boolean).map(function (line) {
      var p = splitSemicolonLine(line);
      return {
        id: p[0],
        name: p[1],
        position: Number(p[2] || "0"),
        level: Number(p[3] || "0"),
        indent: Number(p[4] || "0"),
        depth: Number(p[5] || "0"),
        type: Number(p[6] || "1"),
        parentIds: parseCsvIds(p[8] || ""),
        childIds: parseCsvIds(p[10] || ""),
        weights: parseCsvNumbers(p[11] || "")
      };
    });
  };

  PaloApiClient.prototype.dimensionElementByPosition = async function dimensionElementByPosition(sid, database, dimension, position1Based) {
    var text = await this.call("/dimension/element", {
      sid: sid,
      name_database: database,
      name_dimension: dimension,
      position: Math.floor(position1Based) - 1
    });
    var p = splitSemicolonLine(String(text).trim().split("\n")[0]);
    return {
      id: p[0],
      name: p[1],
      position: Number(p[2] || "0"),
      level: Number(p[3] || "0"),
      indent: Number(p[4] || "0"),
      depth: Number(p[5] || "0"),
      type: Number(p[6] || "1"),
      parentIds: parseCsvIds(p[8] || ""),
      childIds: parseCsvIds(p[10] || ""),
      weights: parseCsvNumbers(p[11] || "")
    };
  };

  PaloApiClient.prototype.elementInfo = async function elementInfo(sid, database, dimension, element) {
    var text = await this.call("/element/info", {
      sid: sid,
      name_database: database,
      name_dimension: dimension,
      name_element: element
    });
    var p = splitSemicolonLine(String(text).trim().split("\n")[0]);
    return {
      id: p[0],
      name: p[1],
      position: Number(p[2] || "0"),
      level: Number(p[3] || "0"),
      indent: Number(p[4] || "0"),
      depth: Number(p[5] || "0"),
      type: Number(p[6] || "1"),
      parentIds: parseCsvIds(p[8] || ""),
      childIds: parseCsvIds(p[10] || ""),
      weights: parseCsvNumbers(p[11] || "")
    };
  };

  function parseCellValueResponseLine(line) {
    var p = splitSemicolonLine(String(line || "").trim());
    var type = Number(p[0] || "0");
    var value = p[2];
    if (value === undefined || value === "") {
      value = p[1];
    }
    if (value === undefined || value === null) {
      value = "";
    }
    paloTrace("cell-value-line", {
      raw: line,
      type: type,
      valueCandidate: String(value)
    });
    if (type === 1) {
      var n = Number(value);
      return Number.isNaN(n) ? null : n;
    }
    if (type === 2) {
      return value;
    }
    return null;
  }

  PaloApiClient.prototype.cellValue = async function cellValue(sid, database, cube, path) {
    var pathStr;
    if (typeof path === "string") {
      pathStr = path;
    } else if (Array.isArray(path)) {
      pathStr = path.map(function (seg) { return String(seg).trim(); }).join(",");
    } else {
      pathStr = normalizePaloCellPath(path).join(",");
    }
    var text = await this.call("/cell/value", {
      sid: sid,
      name_database: database,
      name_cube: cube,
      name_path: pathStr
    });
    var line = String(text).trim().split(/\r?\n/)[0];
    return parseCellValueResponseLine(line);
  };

  PaloApiClient.prototype.cellValueByIds = async function cellValueByIds(sid, name_database, name_cube, idPathStr) {
    var text = await this.call("/cell/value", {
      sid: sid,
      name_database: name_database,
      name_cube: name_cube,
      path: idPathStr
    });
    var line = String(text).trim().split(/\r?\n/)[0];
    return parseCellValueResponseLine(line);
  };

  /**
   * Plusieurs cellules en un seul appel HTTP : un seul name_database et un seul name_cube pour toute la requete.
   * Parametre API `paths` : chemins separes par ":", chaque chemin = liste d'identifiants d'elements separes par des virgules
   * (voir doc Jedox /cell/values ; plus compact en GET que name_paths).
   * @returns {Array} une valeur par path (meme ordre).
   */
  PaloApiClient.prototype.cellValues = async function cellValues(sid, name_database, name_cube, pathsJoined) {
    var text = await this.call("/cell/values", {
      sid: sid,
      name_database: name_database,
      name_cube: name_cube,
      paths: pathsJoined
    });
    var raw = String(text || "");
    paloTrace("cell-values-raw", {
      name_database: name_database,
      name_cube: name_cube,
      rawLength: raw.length,
      rawPreview: raw.slice(0, 800)
    });
    var lines = String(text).trim().split(/\r?\n/).map(function (l) { return l.trim(); }).filter(Boolean);
    paloTrace("cell-values-lines", {
      name_database: name_database,
      name_cube: name_cube,
      lineCount: lines.length,
      firstLines: lines.slice(0, 5)
    });
    return lines.map(function (line) {
      return parseCellValueResponseLine(line);
    });
  };

  PaloApiClient.prototype.cellReplace = async function cellReplace(sid, database, cube, path, value, splash) {
    var pathStr;
    if (typeof path === "string") {
      pathStr = path;
    } else if (Array.isArray(path)) {
      pathStr = path.map(function (seg) { return String(seg).trim(); }).join(",");
    } else {
      pathStr = normalizePaloCellPath(path).join(",");
    }
    var text = await this.call("/cell/replace", {
      sid: sid,
      name_database: database,
      name_cube: cube,
      name_path: pathStr,
      value: normalizePaloPathSegment(value),
      splash: splash || 0
    });
    return String(text).trim().indexOf("1") === 0;
  };

  PaloApiClient.prototype.cellReplaceByIds = async function cellReplaceByIds(
    sid,
    name_database,
    name_cube,
    idPathStr,
    value,
    splash
  ) {
    var text = await this.call("/cell/replace", {
      sid: sid,
      name_database: name_database,
      name_cube: name_cube,
      path: idPathStr,
      value: normalizePaloPathSegment(value),
      splash: splash || 0
    });
    return String(text).trim().indexOf("1") === 0;
  };

  function dimElementCacheTtlMs() {
    if (typeof window !== "undefined" && window.PALO_DIM_ELEMENT_CACHE_TTL_MS != null) {
      var n = Number(window.PALO_DIM_ELEMENT_CACHE_TTL_MS);
      if (!Number.isNaN(n) && n >= 1000) {
        return Math.floor(n);
      }
    }
    return 600000;
  }

  function buildNameToIdMapsFromElements(elements) {
    var byExact = new Map();
    var byLower = new Map();
    var j;
    for (j = 0; j < elements.length; j += 1) {
      var el = elements[j];
      byExact.set(el.name, el.id);
      byLower.set(String(el.name).toLowerCase(), el.id);
    }
    return { byExact: byExact, byLower: byLower };
  }

  function PaloConnectionManager() {
    this.storageKey = "palo.office365.connections.v1";
    this.activeKey = "palo.office365.active.v1";
    this.sessions = new Map();
    this._dbListCache = new Map();
    this._cubeListCache = new Map();
    this._dimensionListCache = new Map();
    this._dimensionListInflight = new Map();
    this._dimOrderCache = new Map();
    this._dimOrderInflight = new Map();
    this._dimElementMapCache = new Map();
    this._dimElementMapInflight = new Map();
    this._cellBatchQueues = new Map();
    this._cellBatchTimers = new Map();
  }

  /** Cle de file : une file = une requete /cell/values homogene (meme connexion, sid, base, cube). */
  function cellBatchKey(connectionName, sid, name_database, name_cube) {
    return String(connectionName || "")
      + "\0" + String(sid || "")
      + "\0" + String(name_database || "")
      + "\0" + String(name_cube || "");
  }

  function cellBatchDelayMs() {
    // Bulk /cell/values actif par defaut (24 ms), avec fallback unitaire et decoupage automatique.
    if (typeof window !== "undefined" && window.PALO_DISABLE_BATCH !== undefined) {
      return window.PALO_DISABLE_BATCH ? 0 : 24;
    }
    if (typeof window !== "undefined" && window.PALO_CELL_BATCH_MS != null) {
      var n = Number(window.PALO_CELL_BATCH_MS);
      if (!Number.isNaN(n) && n >= 0) {
        return n;
      }
    }
    return 24;
  }

  function cellValuesMaxUrlChars() {
    if (typeof window !== "undefined" && window.PALO_CELL_VALUES_MAX_URL_CHARS != null) {
      var n = Number(window.PALO_CELL_VALUES_MAX_URL_CHARS);
      if (!Number.isNaN(n) && n >= 512) {
        return Math.floor(n);
      }
    }
    return 5000;
  }

  /**
   * Regroupe plusieurs PALO.DATAC en un appel /cell/values uniquement lorsqu'ils partagent la meme base et le meme cube
   * (cles de file distinctes par name_database + name_cube ; pas de melange inter-base / inter-cube).
   * Pour ce seul cas : resolution des noms d'elements en IDs a la volee (pas de cache id), puis parametre API `paths`.
   * Desactiver le batch : PALO_CELL_BATCH_MS = 0 (un appel /cell/value par cellule).
   */
  PaloConnectionManager.prototype.requestCellValueBatched = function requestCellValueBatched(
    connectionName,
    sid,
    client,
    name_database,
    name_cube,
    namePath,
    pathSegments,
    debugMeta
  ) {
    var manager = this;
    if (cellBatchDelayMs() === 0) {
      if (Array.isArray(pathSegments) && pathSegments.length > 0) {
        return this.buildCellIdPathsListFromSegments(
          connectionName,
          sid,
          client,
          name_database,
          name_cube,
          [pathSegments]
        ).then(function (list) {
          return client.cellValueByIds(sid, name_database, name_cube, list[0]);
        });
      }
      return client.cellValue(sid, name_database, name_cube, namePath);
    }
    return new Promise(function (resolve, reject) {
      var key = cellBatchKey(connectionName, sid, name_database, name_cube);
      var q = manager._cellBatchQueues.get(key);
      if (!q) {
        q = {
          connectionName: connectionName,
          sid: sid,
          client: client,
          name_database: name_database,
          name_cube: name_cube,
          items: []
        };
        manager._cellBatchQueues.set(key, q);
      }
      q.items.push({
        namePath: namePath,
        pathSegments: pathSegments,
        debugMeta: debugMeta || null,
        resolve: resolve,
        reject: reject
      });
      if (paloBulkTraceEnabled()) {
        paloTrace("cell-values-enqueue", {
          key: key,
          connectionName: connectionName,
          name_database: name_database,
          name_cube: name_cube,
          queueSize: q.items.length,
          requestId: debugMeta && debugMeta.requestId ? debugMeta.requestId : null
        });
      }
      var prev = manager._cellBatchTimers.get(key);
      if (prev) {
        clearTimeout(prev);
      }
      manager._cellBatchTimers.set(
        key,
        setTimeout(function () {
          manager._cellBatchTimers.delete(key);
          manager._flushCellValueBatch(key);
        }, cellBatchDelayMs())
      );
    });
  };

  PaloConnectionManager.prototype._flushCellValueBatch = async function _flushCellValueBatch(key) {
    var q = this._cellBatchQueues.get(key);
    if (!q) {
      return;
    }
    this._cellBatchQueues.delete(key);
    var items = q.items;
    if (!items || !items.length) {
      return;
    }
    var client = q.client;
    var sid = q.sid;
    var name_database = q.name_database;
    var name_cube = q.name_cube;
    try {
      if (items.length === 1) {
        var single;
        if (Array.isArray(items[0].pathSegments) && items[0].pathSegments.length > 0) {
          var singleIdPathList = await this.buildCellIdPathsListFromSegments(
            q.connectionName,
            sid,
            client,
            name_database,
            name_cube,
            [items[0].pathSegments]
          );
          single = await client.cellValueByIds(sid, name_database, name_cube, singleIdPathList[0]);
        } else {
          single = await client.cellValue(sid, name_database, name_cube, items[0].namePath);
        }
        items[0].resolve(single);
        if (paloBulkTraceEnabled()) {
          paloTrace("cell-values-single-resolve", {
            key: key,
            requestId: items[0].debugMeta && items[0].debugMeta.requestId ? items[0].debugMeta.requestId : null,
            value: single
          });
        }
        return;
      }
      var allHaveSegments = items.every(function (it) {
        return Array.isArray(it.pathSegments) && it.pathSegments.length > 0;
      });
      var idPaths;
      var namesJoinedLen = 0;
      if (allHaveSegments) {
        idPaths = await this.buildCellIdPathsListFromSegments(
          q.connectionName,
          sid,
          client,
          name_database,
          name_cube,
          items.map(function (it) { return it.pathSegments; })
        );
      } else {
        var namePathStrings = items.map(function (it) {
          return it.namePath;
        });
        namesJoinedLen = namePathStrings.join(":").length;
        idPaths = await this.buildCellIdPathsList(
          q.connectionName,
          sid,
          client,
          name_database,
          name_cube,
          namePathStrings
        );
      }
      paloTrace("cell-values-batch", {
        connectionName: q.connectionName,
        name_database: name_database,
        name_cube: name_cube,
        count: items.length,
        pathsQueryLen: idPaths.join(":").length,
        namePathsHypotheticalLen: namesJoinedLen
      });
      var maxUrlChars = cellValuesMaxUrlChars();
      var start = 0;
      while (start < idPaths.length) {
        var end = start;
        var joined = "";
        while (end < idPaths.length) {
          var candidate = joined ? (joined + ":" + idPaths[end]) : idPaths[end];
          var candidateUrl = client.buildUrl("/cell/values", {
            sid: sid,
            name_database: name_database,
            name_cube: name_cube,
            paths: candidate
          });
          if (end > start && candidateUrl.length > maxUrlChars) {
            break;
          }
          joined = candidate;
          end += 1;
        }
        if (!joined) {
          joined = idPaths[start];
          end = start + 1;
        }
        paloTrace("cell-values-chunk", {
          connectionName: q.connectionName,
          name_database: name_database,
          name_cube: name_cube,
          start: start,
          end: end,
          total: idPaths.length,
          urlLength: client.buildUrl("/cell/values", {
            sid: sid,
            name_database: name_database,
            name_cube: name_cube,
            paths: joined
          }).length,
          maxUrlChars: maxUrlChars
        });
        var arr;
        try {
          arr = await client.cellValues(sid, name_database, name_cube, joined);
          if (arr.length !== (end - start)) {
            throw new Error(
              "cell/values: " + (end - start) + " chemin(s) envoyes, " + arr.length + " ligne(s) recues."
            );
          }
          var allEmpty = arr.length > 0 && arr.every(function (v) {
            return v === null || v === "";
          });
          if (allEmpty) {
            paloTrace("cell-values-chunk-all-empty-fallback-single", {
              connectionName: q.connectionName,
              name_database: name_database,
              name_cube: name_cube,
              start: start,
              end: end
            });
            arr = [];
            var s;
            for (s = start; s < end; s += 1) {
              arr.push(await client.cellValueByIds(sid, name_database, name_cube, idPaths[s]));
            }
          }
        } catch (chunkErr) {
          // Fiabilite prioritaire: si un chunk batch echoue, fallback cellule par cellule
          // pour eviter de renvoyer des cellules vides sur recalcul massif.
          paloTrace("cell-values-chunk-fallback-single", {
            connectionName: q.connectionName,
            name_database: name_database,
            name_cube: name_cube,
            start: start,
            end: end,
            reason: chunkErr && chunkErr.message ? chunkErr.message : String(chunkErr)
          });
          arr = [];
          var f;
          for (f = start; f < end; f += 1) {
            arr.push(await client.cellValueByIds(sid, name_database, name_cube, idPaths[f]));
          }
        }
        var i;
        for (i = start; i < end; i += 1) {
          if (paloBulkTraceEnabled() || i - start < 5) {
            paloTrace("cell-values-resolve", {
              index: i,
              chunkIndex: i - start,
              idPath: idPaths[i],
              value: arr[i - start],
              requestId: items[i].debugMeta && items[i].debugMeta.requestId ? items[i].debugMeta.requestId : null,
              coordinates: items[i].debugMeta && Array.isArray(items[i].debugMeta.coordinates)
                ? items[i].debugMeta.coordinates
                : null
            });
          }
          items[i].resolve(arr[i - start]);
        }
        start = end;
      }
    } catch (err) {
      paloTrace("cell-values-batch-error", {
        key: key,
        count: items.length,
        reason: err && err.message ? err.message : String(err),
        requestIds: items.map(function (it) {
          return it.debugMeta && it.debugMeta.requestId ? it.debugMeta.requestId : null;
        })
      });
      var j;
      for (j = 0; j < items.length; j += 1) {
        items[j].reject(err);
      }
    }
  };

  PaloConnectionManager.prototype.clearCachesForConnection = function clearCachesForConnection(connectionName) {
    var prefix = String(connectionName || "") + "|";
    function wipe(map) {
      if (!map || !map.keys) {
        return;
      }
      Array.from(map.keys()).forEach(function (k) {
        if (String(k).indexOf(prefix) === 0) {
          map.delete(k);
        }
      });
    }
    wipe(this._dbListCache);
    wipe(this._cubeListCache);
    wipe(this._dimensionListCache);
    wipe(this._dimOrderCache);
    wipe(this._dimensionListInflight);
    wipe(this._dimOrderInflight);
    wipe(this._dimElementMapCache);
    wipe(this._dimElementMapInflight);
  };

  PaloConnectionManager.prototype.getServerDatabasesCached = async function getServerDatabasesCached(
    connectionName,
    sid,
    client
  ) {
    var key = connectionName + "|db-list";
    if (this._dbListCache.has(key)) {
      paloTrace("db-list-cache-hit", { connectionName: connectionName });
      return this._dbListCache.get(key);
    }
    var dbs = await client.serverDatabases(sid);
    this._dbListCache.set(key, dbs);
    paloTrace("db-list-cache-fill", { connectionName: connectionName, count: dbs.length });
    return dbs;
  };

  PaloConnectionManager.prototype.getDatabaseCubesCached = async function getDatabaseCubesCached(
    connectionName,
    sid,
    client,
    database
  ) {
    var key = connectionName + "|cube-list|" + database;
    if (this._cubeListCache.has(key)) {
      paloTrace("cube-list-cache-hit", { connectionName: connectionName, database: database });
      return this._cubeListCache.get(key);
    }
    var cubes = await client.databaseCubes(sid, database);
    this._cubeListCache.set(key, cubes);
    paloTrace("cube-list-cache-fill", { connectionName: connectionName, database: database, count: cubes.length });
    return cubes;
  };

  PaloConnectionManager.prototype.getDatabaseDimensionsCached = async function getDatabaseDimensionsCached(
    connectionName,
    sid,
    client,
    database
  ) {
    var key = connectionName + "|dim-list|" + database;
    if (this._dimensionListCache.has(key)) {
      paloTrace("dimension-list-cache-hit", { connectionName: connectionName, database: database });
      return this._dimensionListCache.get(key);
    }
    if (this._dimensionListInflight.has(key)) {
      return this._dimensionListInflight.get(key);
    }
    var self = this;
    var p = (async function () {
      var dims = await client.databaseDimensions(sid, database);
      self._dimensionListCache.set(key, dims);
      paloTrace("dimension-list-cache-fill", { connectionName: connectionName, database: database, count: dims.length });
      return dims;
    })();
    p.finally(function () {
      self._dimensionListInflight.delete(key);
    });
    this._dimensionListInflight.set(key, p);
    return p;
  };

  PaloConnectionManager.prototype.getCubeDimensionNamesOrdered = async function getCubeDimensionNamesOrdered(
    connectionName,
    sid,
    client,
    database,
    cube
  ) {
    var key = connectionName + "|dims|" + database + "|" + cube;
    if (this._dimOrderCache.has(key)) {
      return this._dimOrderCache.get(key);
    }
    if (this._dimOrderInflight.has(key)) {
      return this._dimOrderInflight.get(key);
    }
    var self = this;
    var p = (async function () {
      var info = await client.cubeInfo(sid, database, cube);
      var allDims = await self.getDatabaseDimensionsCached(connectionName, sid, client, database);
      var idToName = {};
      var i;
      for (i = 0; i < allDims.length; i += 1) {
        idToName[allDims[i].id] = allDims[i].name;
      }
      var names = [];
      for (i = 0; i < info.dimensionIds.length; i += 1) {
        var dn = idToName[info.dimensionIds[i]];
        if (!dn) {
          throw new Error("Dimension id " + info.dimensionIds[i] + " introuvable pour le cube " + cube);
        }
        names.push(dn);
      }
      self._dimOrderCache.set(key, names);
      return names;
    })();
    p.finally(function () {
      self._dimOrderInflight.delete(key);
    });
    this._dimOrderInflight.set(key, p);
    return p;
  };
  PaloConnectionManager.prototype._getDimElementMapCacheKey = function _getDimElementMapCacheKey(
    connectionName,
    database,
    dimName
  ) {
    return connectionName + "|dim-elements-map|" + database + "|" + dimName;
  };

  PaloConnectionManager.prototype.getDimensionElementNameIdMapCached = async function getDimensionElementNameIdMapCached(
    connectionName,
    sid,
    client,
    database,
    dimName,
    forceRefresh
  ) {
    var key = this._getDimElementMapCacheKey(connectionName, database, dimName);
    var ttl = dimElementCacheTtlMs();
    var now = Date.now();
    var cached = this._dimElementMapCache.get(key);
    if (!forceRefresh && cached && (now - cached.loadedAt) <= ttl) {
      return cached;
    }
    if (!forceRefresh && this._dimElementMapInflight.has(key)) {
      return this._dimElementMapInflight.get(key);
    }
    var self = this;
    var p = (async function () {
      var elems = await client.dimensionElements(sid, database, dimName);
      var maps = buildNameToIdMapsFromElements(elems);
      var entry = {
        loadedAt: Date.now(),
        byExact: maps.byExact,
        byLower: maps.byLower
      };
      self._dimElementMapCache.set(key, entry);
      paloTrace("dim-element-map-cache-fill", {
        connectionName: connectionName,
        database: database,
        dimension: dimName,
        count: elems.length,
        forceRefresh: Boolean(forceRefresh)
      });
      return entry;
    })();
    this._dimElementMapInflight.set(key, p);
    p.finally(function () {
      self._dimElementMapInflight.delete(key);
    });
    return p;
  };

  PaloConnectionManager.prototype.resolveNameSetToIdMapCached = async function resolveNameSetToIdMapCached(
    connectionName,
    sid,
    client,
    database,
    dimName,
    nameSet
  ) {
    var names = Array.from(nameSet || []);
    var out = new Map();
    if (!names.length) {
      return out;
    }

    var entry = await this.getDimensionElementNameIdMapCached(connectionName, sid, client, database, dimName, false);
    var missing = [];
    var i;
    for (i = 0; i < names.length; i += 1) {
      var nm = names[i];
      var id = entry.byExact.get(nm);
      if (id === undefined) {
        id = entry.byLower.get(String(nm).toLowerCase());
      }
      if (id === undefined) {
        missing.push(nm);
      } else {
        out.set(nm, id);
      }
    }

    if (missing.length > 0) {
      entry = await this.getDimensionElementNameIdMapCached(connectionName, sid, client, database, dimName, true);
      for (i = 0; i < missing.length; i += 1) {
        var name = missing[i];
        var refreshed = entry.byExact.get(name);
        if (refreshed === undefined) {
          refreshed = entry.byLower.get(String(name).toLowerCase());
        }
        if (refreshed === undefined) {
          throw new Error('Element "' + name + '" introuvable dans la dimension ' + dimName);
        }
        out.set(name, refreshed);
      }
    }

    return out;
  };


  /**
   * Construit name_path (segments separes par des virgules, ordre des dimensions du cube) pour /cell/value et /cell/replace.
   * Les API sont appelees avec name_database, name_cube, name_path — pas de resolution id.
   */
  PaloConnectionManager.prototype.buildCellNamePath = async function buildCellNamePath(
    connectionName,
    sid,
    client,
    database,
    cubeName,
    pathSegments
  ) {
    var dimNames = await this.getCubeDimensionNamesOrdered(connectionName, sid, client, database, cubeName);
    var input = normalizePaloPathSegmentsInput(pathSegments);
    var normalized = [];
    var i;
    for (i = 0; i < input.length; i += 1) {
      normalized.push(String(normalizePaloPathSegment(input[i], { segmentIndex: i, pathLength: input.length })).trim());
    }
    if (normalized.length !== dimNames.length) {
      throw new Error(
        "Nombre de coordonnees (" + normalized.length + ") different du nombre de dimensions du cube (" + dimNames.length + ")."
      );
    }
    for (i = 0; i < normalized.length; i += 1) {
      var seg = String(normalized[i]).trim();
      if (!seg) {
        throw new Error("Coordonnee vide pour la dimension " + dimNames[i]);
      }
      paloTrace("cell-name-path-segment", {
        connectionName: connectionName,
        database: database,
        cube: cubeName,
        dimension: dimNames[i],
        element: seg
      });
    }
    return normalized.join(",");
  };

  /**
   * Pour /cell/values uniquement : name_path (noms, virgules) -> parametre API paths (ids, virgules, chemins separes par ":").
   *
   * Enchainement :
   * 1) Parcourir tous les name_path une fois : parser les segments, remplir uniquePerDim[d] (noms distincts par dimension).
   * 2) Par dimension d : remplir dimMaps[d] = Map nom -> id via cache memoire (TTL + refresh force sur nom manquant).
   * 3) Boucler sur parsed : uniquement dimMaps[d].get(segment) en memoire, construire idPaths puis pathsJoined.
   */
  PaloConnectionManager.prototype.buildCellIdPathsList = async function buildCellIdPathsList(
    connectionName,
    sid,
    client,
    name_database,
    name_cube,
    namePathStrings
  ) {
    var dimNames = await this.getCubeDimensionNamesOrdered(connectionName, sid, client, name_database, name_cube);
    var dimCount = dimNames.length;

    // Etape 1 : collecte des noms uniques par dimension + tableau parsed (segments par chemin). Aucun appel Palo ici.
    var uniquePerDim = [];
    var d;
    for (d = 0; d < dimCount; d += 1) {
      uniquePerDim[d] = new Set();
    }
    var parsed = [];
    var p;
    for (p = 0; p < namePathStrings.length; p += 1) {
      var segs = String(namePathStrings[p]).split(",").map(function (s) {
        return s.trim();
      });
      if (segs.length !== dimCount) {
        throw new Error(
          "cell/values: chemin " + (p + 1) + " a " + segs.length + " segments, " + dimCount + " attendus (dimensions du cube)."
        );
      }
      for (d = 0; d < dimCount; d += 1) {
        if (!segs[d]) {
          throw new Error(
            "Segment vide (chemin " + (p + 1) + ", dimension " + dimNames[d] + ")."
          );
        }
        uniquePerDim[d].add(segs[d]);
      }
      parsed.push(segs);
    }

    // Etape 2 : une Map nom->id par dimension via cache memoire (TTL + refresh sur manquants).
    var dimMaps = [];
    for (d = 0; d < dimCount; d += 1) {
      dimMaps[d] = await this.resolveNameSetToIdMapCached(
        connectionName,
        sid,
        client,
        name_database,
        dimNames[d],
        uniquePerDim[d]
      );
    }

    // Etape 3 : assemblage de la chaine paths — lookups Map seulement, pas de resolution reseau.
    var idPaths = [];
    for (p = 0; p < parsed.length; p += 1) {
      var ids = [];
      for (d = 0; d < dimCount; d += 1) {
        var id = dimMaps[d].get(parsed[p][d]);
        if (id === undefined) {
          throw new Error("ID introuvable pour \"" + parsed[p][d] + "\" (" + dimNames[d] + ").");
        }
        ids.push(String(id));
      }
      idPaths.push(ids.join(","));
    }
    return idPaths;
  };

  PaloConnectionManager.prototype.buildCellIdPathsListFromSegments = async function buildCellIdPathsListFromSegments(
    connectionName,
    sid,
    client,
    name_database,
    name_cube,
    pathSegmentsList
  ) {
    var dimNames = await this.getCubeDimensionNamesOrdered(connectionName, sid, client, name_database, name_cube);
    var dimCount = dimNames.length;
    var uniquePerDim = [];
    var d;
    for (d = 0; d < dimCount; d += 1) {
      uniquePerDim[d] = new Set();
    }
    var normalizedPaths = [];
    var p;
    for (p = 0; p < pathSegmentsList.length; p += 1) {
      var rawPath = Array.isArray(pathSegmentsList[p]) ? pathSegmentsList[p] : [pathSegmentsList[p]];
      // Excel peut fournir les coordonnees comme une plage unique (ex. [[a],[b],...]
      // ou [[a,b,...]]). On normalise avant de verifier la cardinalite.
      rawPath = normalizePaloPathSegmentsInput(rawPath);
      if (rawPath.length !== dimCount) {
        throw new Error(
          "cell/value: " + rawPath.length + " coordonnees recues, " + dimCount + " attendues pour le cube " + name_cube + "."
        );
      }
      var norm = [];
      for (d = 0; d < dimCount; d += 1) {
        var seg = String(normalizePaloPathSegment(rawPath[d], { segmentIndex: d, pathLength: rawPath.length })).trim();
        if (!seg) {
          throw new Error("Coordonnee vide pour la dimension " + dimNames[d]);
        }
        norm.push(seg);
        uniquePerDim[d].add(seg);
      }
      normalizedPaths.push(norm);
    }

    var dimMaps = [];
    for (d = 0; d < dimCount; d += 1) {
      dimMaps[d] = await this.resolveNameSetToIdMapCached(
        connectionName,
        sid,
        client,
        name_database,
        dimNames[d],
        uniquePerDim[d]
      );
    }

    var idPaths = [];
    for (p = 0; p < normalizedPaths.length; p += 1) {
      var ids = [];
      for (d = 0; d < dimCount; d += 1) {
        var id = dimMaps[d].get(normalizedPaths[p][d]);
        if (id === undefined) {
          throw new Error("ID introuvable pour \"" + normalizedPaths[p][d] + "\" (" + dimNames[d] + ").");
        }
        ids.push(String(id));
      }
      idPaths.push(ids.join(","));
    }
    return idPaths;
  };

  PaloConnectionManager.prototype.buildCellIdPathFromSegments = async function buildCellIdPathFromSegments(
    connectionName,
    sid,
    client,
    name_database,
    name_cube,
    pathSegments
  ) {
    var list = await this.buildCellIdPathsListFromSegments(
      connectionName,
      sid,
      client,
      name_database,
      name_cube,
      [pathSegments]
    );
    return list[0];
  };

  PaloConnectionManager.prototype.buildCellIdPathsColonJoined = async function buildCellIdPathsColonJoined(
    connectionName,
    sid,
    client,
    name_database,
    name_cube,
    namePathStrings
  ) {
    var list = await this.buildCellIdPathsList(
      connectionName,
      sid,
      client,
      name_database,
      name_cube,
      namePathStrings
    );
    return list.join(":");
  };

  PaloConnectionManager.prototype.listConnections = function listConnections() {
    var raw = null;
    try {
      raw = window.localStorage.getItem(this.storageKey);
    } catch (_e) {
      return [];
    }
    if (!raw) {
      return [];
    }
    try {
      var parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) {
        return [];
      }
      return parsed.filter(function (p) {
        return p && typeof p === "object" && typeof p.name === "string" && p.name.length > 0;
      });
    } catch (_error) {
      return [];
    }
  };

  PaloConnectionManager.prototype.saveConnection = function saveConnection(profile) {
    if (!profile || typeof profile.name !== "string" || !profile.name) {
      throw new Error("Profil connexion invalide (nom requis).");
    }
    var all = this.listConnections().filter(function (p) { return p.name !== profile.name; });
    all.push(profile);
    paloLocalStorageSetItem(this.storageKey, JSON.stringify(all));
    if (!this.getActiveConnectionName()) {
      this.setActiveConnectionName(profile.name);
    }
  };

  PaloConnectionManager.prototype.deleteConnection = function deleteConnection(name) {
    var all = this.listConnections().filter(function (p) { return p.name !== name; });
    paloLocalStorageSetItem(this.storageKey, JSON.stringify(all));
    if (this.getActiveConnectionName() === name) {
      paloLocalStorageRemoveItem(this.activeKey);
    }
    this.sessions.delete(name);
    this.clearCachesForConnection(name);
  };

  PaloConnectionManager.prototype.getConnection = function getConnection(name) {
    var profile = this.listConnections().find(function (p) { return p.name === name; });
    if (!profile) {
      throw new Error("Connexion introuvable: " + name);
    }
    return profile;
  };

  PaloConnectionManager.prototype.setActiveConnectionName = function setActiveConnectionName(name) {
    paloLocalStorageSetItem(this.activeKey, name);
  };

  PaloConnectionManager.prototype.getActiveConnectionName = function getActiveConnectionName() {
    try {
      return window.localStorage.getItem(this.activeKey);
    } catch (_e) {
      return null;
    }
  };

  PaloConnectionManager.prototype.parseServDb = function parseServDbPublic(servdb) {
    return parseServDb(servdb);
  };

  PaloConnectionManager.prototype.getClientAndContext = async function getClientAndContext(servdb) {
    var parsed = parseServDb(servdb);
    var profile = this.getConnection(parsed.connectionName);
    var client = new PaloApiClient(profile);
    var sid = await this.getValidSid(profile.name, client);
    return {
      client: client,
      sid: sid,
      database: parsed.database,
      connectionName: parsed.connectionName
    };
  };

  PaloConnectionManager.prototype.testConnection = async function testConnection(name) {
    var profile = this.getConnection(name);
    var client = new PaloApiClient(profile);
    try {
      var session = await client.login();
      var dbs = await client.serverDatabases(session.sid);
      try {
        await client.logout(session.sid);
      } catch (eLogout) {
        // ignorer
      }
      return {
        ok: true,
        details: "Connexion OK — " + dbs.length + " base(s) Palo accessible(s)."
      };
    } catch (error) {
      return {
        ok: false,
        details: error && error.message ? error.message : String(error)
      };
    }
  };

  PaloConnectionManager.prototype.getValidSid = async function getValidSid(connectionName, client) {
    var cached = this.sessions.get(connectionName);
    if (cached) {
      var ageMs = Date.now() - cached.createdAt;
      var ttlMs = cached.ttlSeconds * 1000;
      if (ageMs < ttlMs - 30000) {
        if (paloHttpLogEnabled()) {
          console.info("[PaloOffice] getValidSid cache OK", {
            connectionName: connectionName,
            ageMs: ageMs,
            sidPrefix: cached.sid ? cached.sid.slice(0, 8) : ""
          });
        }
        return cached.sid;
      }
    }
    if (paloHttpLogEnabled()) {
      console.info("[PaloOffice] getValidSid login Palo", { connectionName: connectionName });
    }
    var login = await client.login();
    this.sessions.set(connectionName, {
      sid: login.sid,
      ttlSeconds: login.ttl,
      createdAt: Date.now()
    });
    if (paloHttpLogEnabled()) {
      console.info("[PaloOffice] getValidSid login OK", {
        connectionName: connectionName,
        sidPrefix: login.sid ? login.sid.slice(0, 8) : "",
        ttlSeconds: login.ttl
      });
    }
    return login.sid;
  };

  window.PaloOffice = window.PaloOffice || {};
  window.PaloOffice.ApiClient = PaloApiClient;
  window.PaloOffice.ConnectionManager = PaloConnectionManager;
  window.PaloOffice.trace = paloTrace;
  window.PaloOffice.getTraceHistory = paloGetTraceHistory;
  window.PaloOffice.getLastApiUrl = paloGetLastApiUrl;
  window.PaloOffice.createConnectionManager = function createConnectionManager() {
    return new PaloConnectionManager();
  };
})();

