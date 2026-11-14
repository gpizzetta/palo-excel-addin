/* global CustomFunctions, OfficeRuntime */
/* Source des fonctions Excel : editer ce fichier puis ./build-bundle.sh (genere functions.js). */
var PALO_CDN_BASE = "https://gpizzetta.github.io/palo-excel-addin";
var PALO_ASSET_VERSION = "1.0.3.38";

(function paloFunctionsBootstrap() {
  var connectionManager = null;
  var datacRequestSeq = 0;
  var paloBootstrapPromise = null;

  function paloGlobalRef() {
    if (typeof globalThis !== "undefined") {
      return globalThis;
    }
    if (typeof self !== "undefined") {
      return self;
    }
    if (typeof window !== "undefined") {
      return window;
    }
    try {
      var viaFunction = Function("return this")();
      if (viaFunction) {
        return viaFunction;
      }
    } catch (_e) {
    }
    if (typeof Function !== "undefined" && Function.__PALO_RUNTIME_BAG__) {
      return Function.__PALO_RUNTIME_BAG__;
    }
    var g = {};
    if (typeof Function !== "undefined") {
      Function.__PALO_RUNTIME_BAG__ = g;
    }
    return g;
  }

  var gBoot = paloGlobalRef();
  if (gBoot.__PALO_FUNCTIONS_CORE_LOADED__) {
    if (typeof gBoot.__paloScheduleRegisterCustomFunctions === "function") {
      gBoot.__paloScheduleRegisterCustomFunctions();
    }
    return;
  }
  gBoot.__PALO_FUNCTIONS_CORE_LOADED__ = true;

  function paloFnTrace() {
    var g = paloGlobalRef();
    return Boolean(g.PALO_DEBUG);
  }

  function hasPaloOffice() {
    var g = paloGlobalRef();
    return Boolean(g.PaloOffice && typeof g.PaloOffice.createConnectionManager === "function");
  }

  async function paloCfAwaitStorageReady() {
    if (!hasPaloOffice()) {
      return;
    }
    var po = paloGlobalRef().PaloOffice;
    if (po && typeof po.paloEnsureStorageReady === "function") {
      await po.paloEnsureStorageReady();
    }
  }

  async function paloCfEnsureRuntimeReady() {
    await ensurePaloOfficeReady();
    await paloCfAwaitStorageReady();
  }

  function paloWarmActiveSession() {
    getConnectionManager()
      .then(function (manager) {
        if (!manager) {
          return;
        }
        var name = manager.getActiveConnectionName();
        if (!name) {
          return;
        }
        var profile = manager.getConnection(name);
        var ApiClient = paloGlobalRef().PaloOffice.ApiClient;
        var client = new ApiClient(profile);
        return manager.getValidSid(name, client);
      })
      .catch(function () {
        // ignore
      });
  }

  function paloWarmStorageAfterRegister() {
    if (!hasPaloOffice()) {
      return;
    }
    var po = paloGlobalRef().PaloOffice;
    // Desktop : charger ORT apres register (onReady peut timeout — gere dans palo-api).
    function warmAfterReady() {
      if (po && typeof po.paloEnsureStorageReady === "function") {
        po.paloEnsureStorageReady(true).catch(function () {});
      }
      paloWarmActiveSession();
    }
    if (typeof Office !== "undefined" && Office && typeof Office.onReady === "function") {
      Office.onReady(warmAfterReady);
    } else {
      warmAfterReady();
    }
  }

  function resolveAfterStorageReady(resolve) {
    var po = paloGlobalRef().PaloOffice;
    var storageReady = po.paloEnsureStorageReady && po.paloEnsureStorageReady();
    if (storageReady && typeof storageReady.then === "function") {
      storageReady.then(resolve).catch(resolve);
    } else {
      resolve();
    }
  }

  function ensurePaloOfficeReady() {
    if (!paloBootstrapPromise) {
      paloBootstrapPromise = new Promise(function (resolve, reject) {
        var deadline = Date.now() + 20000;

        function failNotReady() {
          var diag = "globalThis=" + (typeof globalThis !== "undefined" ? "oui" : "non");
          diag += " self=" + (typeof self !== "undefined" ? "oui" : "non");
          diag += " document=" + (typeof document !== "undefined" ? "oui" : "non");
          diag += " importScripts=" + (typeof importScripts === "function" ? "oui" : "non");
          reject(new Error(
            "PaloOffice indisponible dans le runtime des formules (" + diag + "). "
            + "Testez =PALO.RUNTIME_DIAG() puis ouvrez le volet Connexion et recalculez (Ctrl+Alt+F9)."
          ));
        }

        function tryFinish() {
          if (hasPaloOffice()) {
            resolveAfterStorageReady(resolve);
            return true;
          }
          return false;
        }

        function tryImportScripts() {
          if (typeof importScripts !== "function") {
            return false;
          }
          try {
            importScripts(PALO_CDN_BASE + "/functions-bundle.js?v=" + PALO_ASSET_VERSION);
            return tryFinish();
          } catch (e1) {
            try {
              importScripts(PALO_CDN_BASE + "/assets/palo-api.js?v=" + PALO_ASSET_VERSION);
              return tryFinish();
            } catch (e2) {
              reject(e1);
              return true;
            }
          }
        }

        function schedulePoll() {
          if (tryFinish()) {
            return;
          }
          if (Date.now() > deadline) {
            failNotReady();
            return;
          }
          setTimeout(schedulePoll, 50);
        }

        if (tryFinish()) {
          return;
        }
        if (tryImportScripts()) {
          return;
        }
        if (typeof document !== "undefined" && document.head) {
          var script = document.createElement("script");
          script.src = PALO_CDN_BASE + "/functions-bundle.js?v=" + PALO_ASSET_VERSION;
          script.onload = function () {
            if (!tryFinish()) {
              schedulePoll();
            }
          };
          script.onerror = function () {
            reject(new Error("Echec chargement palo-api.js depuis " + script.src));
          };
          document.head.appendChild(script);
          schedulePoll();
          return;
        }
        failNotReady();
      });
    }
    return paloBootstrapPromise;
  }

  async function getConnectionManager() {
    try {
      await paloCfEnsureRuntimeReady();
    } catch (_boot) {
      return null;
    }
    if (!hasPaloOffice()) {
      return null;
    }
    if (!connectionManager) {
      connectionManager = paloGlobalRef().PaloOffice.createConnectionManager();
    }
    var po = paloGlobalRef().PaloOffice;
    // Desktop CF : si LS vide, recharger depuis OfficeRuntime.storage.
    if (connectionManager.listConnections().length === 0 && po
      && typeof po.paloReloadOfficeRuntimeStorage === "function") {
      try {
        await po.paloReloadOfficeRuntimeStorage();
      } catch (_reload) {
        // ignore
      }
    }
    return connectionManager;
  }

  function toError(error) {
    var message = error && error.message ? error.message : String(error);
    return "#PALO! " + message;
  }

  async function isDebugEnabledForServdb(servdb) {
    try {
      var manager = await getConnectionManager();
      var parsed = manager.parseServDb(servdb);
      var profile = manager.getConnection(parsed.connectionName);
      return Boolean(profile && profile.debug);
    } catch (_e) {
      return false;
    }
  }

  function asStringMatrix(values) {
    return values.map(function (value) { return [value]; });
  }

  function paloDatacPerf(step, fields) {
    var po = paloGlobalRef().PaloOffice;
    if (po && typeof po.datacPerf === "function") {
      po.datacPerf(step, fields || null);
      return;
    }
    try {
      console.log("[Palo DATAC JSON]\n" + JSON.stringify([{ step: step, fields: fields || null }], null, 2));
    } catch (_log) {
      // ignore
    }
  }

  function nextDatacRequestId() {
    datacRequestSeq += 1;
    return "datac-" + String(datacRequestSeq);
  }

  function traceDatac(eventName, payload) {
    var po = paloGlobalRef().PaloOffice;
    if (po && typeof po.trace === "function") {
      po.trace(eventName, payload || {});
    }
  }

  /** Diagnostic runtime Excel (ne depend pas de PaloOffice). */
  function RUNTIME_DIAG() {
    try {
      var g = paloGlobalRef();
      var parts = [
        "v=" + PALO_ASSET_VERSION,
        "cfReg=" + (g.__PALO_CF_REGISTERED__ ? "oui" : "non"),
        "globalThis=" + (typeof globalThis !== "undefined" ? "oui" : "non"),
        "self=" + (typeof self !== "undefined" ? "oui" : "non"),
        "window=" + (typeof window !== "undefined" ? "oui" : "non"),
        "document=" + (typeof document !== "undefined" ? "oui" : "non"),
        "importScripts=" + (typeof importScripts === "function" ? "oui" : "non"),
        "CustomFunctions=" + (typeof CustomFunctions !== "undefined" ? "oui" : "non"),
        "Office=" + (typeof Office !== "undefined" ? "oui" : "non"),
        "paloApi=" + (g.__PALO_API_LOADED__ ? "oui" : "non"),
        "PaloOffice=" + (g.PaloOffice ? "oui" : "non"),
        "createCM=" + (g.PaloOffice && typeof g.PaloOffice.createConnectionManager === "function" ? "oui" : "non"),
        "bundleFile=functions.js"
      ];
      if (g.PaloOffice && typeof g.PaloOffice.paloStorageDiagSync === "function") {
        parts.push(g.PaloOffice.paloStorageDiagSync());
      }
      return parts.join(" ");
    } catch (err) {
      return "RUNTIME_DIAG_ERROR " + (err && err.message ? err.message : String(err));
    }
  }

  /**
   * Diagnostic stockage.
   * Desktop attendu : channel=desktop/ort ort=oui ortPull=oui conn>=1
   * Web attendu : channel=web/ls ls=oui conn>=1
   */
  async function STORAGE_DIAG() {
    var g = paloGlobalRef();
    var po = g.PaloOffice;
    var parts = [RUNTIME_DIAG()];
    if (po && typeof po.paloAwaitOfficeReadyAsync === "function") {
      try {
        parts.push("onReady=" + String(await po.paloAwaitOfficeReadyAsync(1500)));
      } catch (_ready) {
        parts.push("onReady=err");
      }
    }
    if (po && typeof po.paloPullOfficeRuntimeStorageIntoMem === "function") {
      var ortPull = false;
      try {
        ortPull = await po.paloPullOfficeRuntimeStorageIntoMem();
      } catch (_ort) {
        ortPull = false;
      }
      parts.push("ortPull=" + (ortPull ? "oui" : "non"));
    }
    if (po && typeof po.paloStorageDiagSync === "function") {
      parts.push(po.paloStorageDiagSync());
    }
    return parts.join(" ");
  }

  /**
   * Contexte interne Excel (custom functions) : preview, invocation, etc.
   * Ne doit jamais etre traite comme une coordonnee Palo.
   * En value preview, setResult/setError ne sont pas toujours typeof "function".
   */
  function isOfficeCustomFunctionMeta(value) {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      return false;
    }
    if (Object.prototype.hasOwnProperty.call(value, "_functionName")) {
      return true;
    }
    if (Object.prototype.hasOwnProperty.call(value, "_isInValuePreview")) {
      return true;
    }
    if (Object.prototype.hasOwnProperty.call(value, "onCanceled")
      && Object.prototype.hasOwnProperty.call(value, "setResult")) {
      return true;
    }
    if (Object.prototype.hasOwnProperty.call(value, "parameterAddresses")) {
      return true;
    }
    if (Object.prototype.hasOwnProperty.call(value, "setResult")
      && Object.prototype.hasOwnProperty.call(value, "setError")) {
      return true;
    }
    return false;
  }

  function sanitizePaloCoordinates(values) {
    if (!values || !values.length) {
      return [];
    }
    return values.filter(function (value) {
      return !isOfficeCustomFunctionMeta(value);
    });
  }

  /** Deplie plage 1x1 / objet cellule Excel vers une valeur scalaire. */
  /**
   * Si false : les cellules en erreur Excel (#REF!, #BUSY!, …) declenchent encore Palo (comportement ancien).
   * Defaut : true — pas d'appel Palo si amont vide ou en erreur (limite la cascade de recalcul / HTTP).
   */
  function skipExcelUpstreamBlockingEnabled() {
    if (typeof window === "undefined") {
      return true;
    }
    return window.PALO_SKIP_ON_EXCEL_ERROR !== false;
  }

  function isBlankPaloArg(raw) {
    if (isOfficeCustomFunctionMeta(raw)) {
      return false;
    }
    var v = coerceExcelScalarArg(raw);
    if (v === null || v === undefined) {
      return true;
    }
    return String(v).trim() === "";
  }

  function isExcelOrPaloErrorText(scalar) {
    var s = String(scalar == null ? "" : scalar).trim();
    if (!s || s.charAt(0) !== "#") {
      return false;
    }
    if (s.indexOf("#PALO!") === 0) {
      return true;
    }
    var u = s.toUpperCase();
    var known = [
      "#N/A", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#DIV/0!", "#NULL!",
      "#GETTING_DATA", "#SPILL!", "#CALC!", "#BUSY!", "#CONNECT!", "#BLOCKED!",
      "#FIELD!", "#UNKNOWN!"
    ];
    var i;
    for (i = 0; i < known.length; i++) {
      if (u === known[i].toUpperCase()) {
        return true;
      }
    }
    if (/^#N\/A$/i.test(s)) {
      return true;
    }
    if (/^#[A-Z][A-Z0-9]{1,15}[!?]$/i.test(s)) {
      return true;
    }
    return false;
  }

  /** Vide, ou erreur Excel / #PALO! si court-circuit active (pour DATAC, coords, cube, servdb). */
  function shouldBlockPaloDatacArg(raw) {
    if (isBlankPaloArg(raw)) {
      return true;
    }
    if (!skipExcelUpstreamBlockingEnabled()) {
      return false;
    }
    if (isOfficeCustomFunctionMeta(raw)) {
      return false;
    }
    return isExcelOrPaloErrorText(coerceExcelScalarArg(raw));
  }

  /** Erreur Excel / #PALO! seulement (pas le vide — laisse les messages #PALO! ENAME explicites). */
  function shouldBlockPaloErrorOnly(raw) {
    if (!skipExcelUpstreamBlockingEnabled()) {
      return false;
    }
    if (isOfficeCustomFunctionMeta(raw)) {
      return false;
    }
    return isExcelOrPaloErrorText(coerceExcelScalarArg(raw));
  }

  function emptyPaloMatrixCell() {
    return [[""]];
  }

  function coerceExcelScalarArg(value) {
    if (value === null || value === undefined) {
      return value;
    }
    if (isOfficeCustomFunctionMeta(value)) {
      return value;
    }
    if (typeof value === "object" && !Array.isArray(value)) {
      if (Object.prototype.hasOwnProperty.call(value, "text") && value.text !== undefined && value.text !== null) {
        return value.text;
      }
      if (Object.prototype.hasOwnProperty.call(value, "value")) {
        return coerceExcelScalarArg(value.value);
      }
      if (Object.prototype.hasOwnProperty.call(value, "basicValue")) {
        return coerceExcelScalarArg(value.basicValue);
      }
      if (typeof value.valueOf === "function") {
        var prim = value.valueOf();
        if (prim !== value && (typeof prim === "string" || typeof prim === "number" || typeof prim === "boolean")) {
          return prim;
        }
      }
      return value;
    }
    if (Array.isArray(value)) {
      if (value.length === 0) {
        return "";
      }
      return coerceExcelScalarArg(value[0]);
    }
    return value;
  }

  function paloCfIsRuntimeError(error) {
    return typeof CustomFunctions !== "undefined"
      && CustomFunctions.Error
      && error instanceof CustomFunctions.Error;
  }

  /** Erreur Excel #VALEUR! avec message visible au survol (Custom Functions runtime). */
  function paloCfThrowInvalidValue(detail) {
    var msg = "PALO — " + detail;
    if (typeof CustomFunctions !== "undefined" && CustomFunctions.Error && CustomFunctions.ErrorCode) {
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, msg);
    }
    throw new Error(detail);
  }

  function paloCfHandleCatch(error) {
    if (paloCfIsRuntimeError(error)) {
      throw error;
    }
    var message = error && error.message ? error.message : String(error);
    if (typeof CustomFunctions !== "undefined" && CustomFunctions.Error && CustomFunctions.ErrorCode) {
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "PALO — " + message);
    }
    return toError(error);
  }

  function paloCfDatacReturn(value) {
    if (value === null || value === undefined) {
      return "";
    }
    if (typeof value === "number") {
      return Number.isFinite(value) ? value : "";
    }
    if (typeof value === "boolean") {
      return value;
    }
    return String(value);
  }

  function paloBuildNamePathFromCoords(coordinates, cubeName) {
    var po = paloGlobalRef().PaloOffice;
    if (po && typeof po.paloBuildNamePathSafe === "function") {
      return po.paloBuildNamePathSafe(coordinates, cubeName);
    }
    return { ok: false, path: "", error: "paloBuildNamePathSafe indisponible" };
  }

  /** Evite slice(arguments) et recursion profonde — lecture CF argument par argument. */
  function paloCoerceCfArgSafe(raw, depth) {
    if (depth === undefined) {
      depth = 0;
    }
    if (depth > 10) {
      return "";
    }
    try {
      if (isOfficeCustomFunctionMeta(raw)) {
        return "";
      }
      if (raw === null || raw === undefined) {
        return "";
      }
      var t = typeof raw;
      if (t === "string" || t === "number" || t === "boolean") {
        return raw;
      }
      if (Array.isArray(raw)) {
        if (!raw.length) {
          return "";
        }
        return paloCoerceCfArgSafe(raw[0], depth + 1);
      }
      if (t === "object") {
        if (raw.text !== undefined && raw.text !== null) {
          return String(raw.text);
        }
        if (raw.basicValue !== undefined && raw.basicValue !== null) {
          return paloCoerceCfArgSafe(raw.basicValue, depth + 1);
        }
        if (raw.value !== undefined) {
          return paloCoerceCfArgSafe(raw.value, depth + 1);
        }
        if (typeof raw.valueOf === "function") {
          var prim = raw.valueOf();
          if (prim !== raw) {
            return paloCoerceCfArgSafe(prim, depth + 1);
          }
        }
        return "";
      }
      return String(raw);
    } catch (_coerce) {
      return "";
    }
  }

  function paloCollectCoordinateArgs(cfArgs, startIndex) {
    var out = [];
    var i;
    for (i = startIndex; i < cfArgs.length; i += 1) {
      try {
        if (isOfficeCustomFunctionMeta(cfArgs[i])) {
          continue;
        }
        out.push(cfArgs[i]);
      } catch (_skip) {
      }
    }
    // Excel Online passe coord1..coord20 optionnels : les slots non fournis arrivent vides.
    // Les retirer en queue, sinon shouldBlockPaloDatacArg / join renvoient "" ou une erreur.
    while (out.length && isBlankPaloArg(out[out.length - 1])) {
      out.pop();
    }
    return out;
  }

  async function getClientContextForServdb(servdb) {
    var manager = await getConnectionManager();
    if (!manager) {
      return null;
    }
    return manager.getClientAndContext(servdb);
  }

  async function getElementByName(servdb, dimension, elementName) {
    if (shouldBlockPaloDatacArg(servdb) || shouldBlockPaloDatacArg(dimension) || shouldBlockPaloDatacArg(elementName)) {
      return null;
    }
    var context = await getClientContextForServdb(servdb);
    if (!context) {
      return null;
    }
    return context.client.elementInfo(context.sid, context.database, dimension, elementName);
  }

  async function DATAC(servdb, cubeName) {
    var cfArgs = arguments;
    var coordinates = paloCollectCoordinateArgs(cfArgs, 2);
    servdb = paloCoerceCfArgSafe(servdb);
    cubeName = paloCoerceCfArgSafe(cubeName);
    var requestId = nextDatacRequestId();
    var blockedEarly = shouldBlockPaloDatacArg(servdb)
      || shouldBlockPaloDatacArg(cubeName)
      || coordinates.some(function (coord) {
        return shouldBlockPaloDatacArg(coord);
      });
    if (blockedEarly) {
      traceDatac("datac-skip-upstream-blocked", {
        requestId: requestId,
        servdb: String(servdb || ""),
        cubeName: String(cubeName || ""),
        coordinatesCount: coordinates.length
      });
      return "";
    }

    try {
      var started = Date.now();
      var mode = "ids";
      paloDatacPerf("debut", { id: requestId, cube: String(cubeName || ""), mode: mode });
      var stepAt = Date.now();
      var manager = await getConnectionManager();
      var managerMs = Date.now() - stepAt;
      paloDatacPerf("manager", { id: requestId, ms: managerMs });
      if (!manager) {
        paloDatacPerf("arret", { id: requestId, ms: Date.now() - started, error: "runtime pas pret" });
        return "";
      }
      stepAt = Date.now();
      var context = await manager.getClientAndContext(servdb);
      var contextMs = Date.now() - stepAt;
      paloDatacPerf("contexte", {
        id: requestId,
        ms: contextMs,
        sid: context.sidFromCache ? "cache" : "login"
      });
      if (paloFnTrace()) {
        console.info("[PaloOffice DATAC] cell/value params", {
          requestId: requestId,
          name_database: context.database,
          name_cube: cubeName,
          coordinates: coordinates.map(function (coord) {
            return String(coerceExcelScalarArg(coord));
          })
        });
      }
      traceDatac("datac-start", {
        requestId: requestId,
        servdb: String(servdb || ""),
        connectionName: context.connectionName,
        database: context.database,
        cubeName: String(cubeName || ""),
        coordinates: coordinates.map(function (coord) {
          return String(coerceExcelScalarArg(coord));
        }),
        mode: "id_path"
      });
      var coordsScalar = coordinates.map(function (coord) {
        return String(coerceExcelScalarArg(coord));
      });
      stepAt = Date.now();
      var value = await manager.requestCellValueBatched(
        context.connectionName,
        context.sid,
        context.client,
        context.database,
        cubeName,
        "",
        coordsScalar,
        {
          requestId: requestId,
          coordinates: coordsScalar
        }
      );
      var readMs = Date.now() - stepAt;
      paloDatacPerf("lecture", { id: requestId, ms: readMs, mode: mode });
      paloDatacPerf("fin", {
        id: requestId,
        ms: Date.now() - started,
        managerMs: managerMs,
        contexteMs: contextMs,
        sid: context.sidFromCache ? "cache" : "login",
        lectureMs: readMs,
        mode: mode,
        cube: String(cubeName || "")
      });
      if (paloFnTrace()) {
        console.info("[PaloOffice DATAC] fin OK", { requestId: requestId, value: value });
      }
      traceDatac("datac-end", {
        requestId: requestId,
        value: value
      });
      return paloCfDatacReturn(value);
    } catch (error) {
      var msg = error && error.message ? String(error.message) : String(error);
      paloDatacPerf("erreur", { id: requestId, error: msg });
      if (paloFnTrace() || (typeof window !== "undefined" && window.PALO_LOG_HTTP)) {
        console.warn("[PaloOffice DATAC] erreur", { requestId: requestId, message: msg });
      }
      traceDatac("datac-error", {
        requestId: requestId,
        message: msg
      });
      if (
        await isDebugEnabledForServdb(servdb)
        || msg.indexOf("Timeout HTTP") !== -1
        || msg.indexOf("HTTP ") !== -1
        || msg.indexOf("Impossible de joindre") !== -1
      ) {
        return toError(error);
      }
      return "";
    }
  }

  async function PALO_SETDATA(value, splash, servdb, cubeName) {
    var coordinates = paloCollectCoordinateArgs(arguments, 4);
    if (paloFnTrace()) {
      console.info("[PaloOffice PALO_SETDATA] start", {
        servdb: servdb,
        cubeName: cubeName,
        nbCoords: coordinates.length,
        splash: splash
      });
    }
    if (
      shouldBlockPaloDatacArg(value)
      || shouldBlockPaloDatacArg(servdb)
      || shouldBlockPaloDatacArg(cubeName)
      || coordinates.some(function (c) {
        return shouldBlockPaloDatacArg(c);
      })
    ) {
      return 0;
    }
    try {
      var manager = await getConnectionManager();
      if (!manager) {
        return 0;
      }
      var context = await manager.getClientAndContext(servdb);
      var idPath = await manager.buildCellIdPathFromSegments(
        context.connectionName,
        context.sid,
        context.client,
        context.database,
        cubeName,
        coordinates
      );
      if (paloFnTrace()) {
        console.info("[PaloOffice PALO_SETDATA] cell/replace params", {
          name_database: context.database,
          name_cube: cubeName,
          path: idPath
        });
      }
      var ok = await context.client.cellReplaceByIds(
        context.sid,
        context.database,
        cubeName,
        idPath,
        value,
        splash || 0
      );
      if (paloFnTrace()) {
        console.info("[PaloOffice PALO_SETDATA] fin OK", { ok: ok });
      }
      return ok ? 1 : 0;
    } catch (error) {
      if (paloFnTrace() || (typeof window !== "undefined" && window.PALO_LOG_HTTP)) {
        console.warn("[PaloOffice PALO_SETDATA] erreur", error && error.message ? error.message : String(error));
      }
      return toError(error);
    }
  }

  async function ENAME() {
    try {
      var args = Array.prototype.slice.call(arguments);
      if (args.length < 3) {
        return "#PALO! ENAME: il faut au moins 3 arguments (servdb; dimension; element). Recu: " + args.length + ".";
      }
      var servdb = args[0];
      var dimensionName = args[1];
      var index = args[2];

      var manager = await getConnectionManager();
      if (!manager) {
        return "#PALO! ENAME: runtime pas pret (F9 pour recalculer).";
      }
      var servRaw = coerceExcelScalarArg(servdb);
      var dimRaw = coerceExcelScalarArg(dimensionName);
      var elementRaw = coerceExcelScalarArg(index);
      if (isOfficeCustomFunctionMeta(servdb) || isOfficeCustomFunctionMeta(dimensionName) || isOfficeCustomFunctionMeta(index)) {
        return "#PALO! ENAME: argument Office interne detecte. Reference une cellule a valeur unique.";
      }
      var serv = String(servRaw == null ? "" : servRaw).trim();
      var dim = String(dimRaw == null ? "" : dimRaw).trim();
      if (!serv) {
        return "#PALO! ENAME: servdb vide (attendu ex. CONNEXION/BASE).";
      }
      if (!dim) {
        return "#PALO! ENAME: nom de dimension vide.";
      }
      var query = String(elementRaw == null ? "" : elementRaw).trim();
      if (!query) {
        return "#PALO! ENAME: element vide.";
      }
      if (
        shouldBlockPaloErrorOnly(servdb)
        || shouldBlockPaloErrorOnly(dimensionName)
        || shouldBlockPaloErrorOnly(index)
        || (args.length > 3 && shouldBlockPaloErrorOnly(args[3]))
      ) {
        return "";
      }
      var attribute = "";
      if (args.length > 3 && !isOfficeCustomFunctionMeta(args[3]) && !isBlankPaloArg(args[3])) {
        attribute = String(coerceExcelScalarArg(args[3]) == null ? "" : coerceExcelScalarArg(args[3])).trim();
      }
      var context = await manager.getClientAndContext(serv);
      var elementName = "";
      try {
        var elem = await context.client.elementInfo(context.sid, context.database, dim, query);
        elementName = String(elem.name);
      } catch (_exact) {
        var items = await context.client.dimensionElements(context.sid, context.database, dim);
        var qLower = query.toLowerCase();
        var found = items.find(function (item) {
          return String(item.name) === query;
        });
        if (!found) {
          found = items.find(function (item) {
            return String(item.name).toLowerCase() === qLower;
          });
        }
        if (!found) {
          return "#PALO! ENAME: element \"" + query + "\" introuvable dans \"" + dim + "\".";
        }
        elementName = String(found.name);
      }
      if (!attribute) {
        return elementName;
      }
      var cubeName = await context.client.attributeCubeName(context.sid, context.database, dim);
      if (!cubeName) {
        return "#PALO! ENAME: pas de cube d'attributs pour \"" + dim + "\".";
      }
      var value = await context.client.cellValue(context.sid, context.database, cubeName, [attribute, elementName]);
      if (value == null || value === "") {
        return elementName;
      }
      return value;
    } catch (error) {
      return toError(error);
    }
  }

  async function PALO_ECOUNT(servdb, dimensionName) {
    try {
      if (shouldBlockPaloDatacArg(servdb) || shouldBlockPaloDatacArg(dimensionName)) {
        return 0;
      }
      var context = await getClientContextForServdb(servdb);
      if (!context) {
        return 0;
      }
      var items = await context.client.dimensionElements(context.sid, context.database, dimensionName);
      return items.length;
    } catch (error) {
      return toError(error);
    }
  }

  async function PALO_ECHILDCOUNT(servdb, dimensionName, elementName) {
    try {
      var elem = await getElementByName(servdb, dimensionName, elementName);
      if (!elem) {
        return 0;
      }
      return elem.childIds.length;
    } catch (error) {
      return toError(error);
    }
  }

  async function PALO_ECHILD(servdb, dimensionName, elementName, childIndex) {
    try {
      if (
        shouldBlockPaloDatacArg(servdb)
        || shouldBlockPaloDatacArg(dimensionName)
        || shouldBlockPaloDatacArg(elementName)
        || shouldBlockPaloDatacArg(childIndex)
      ) {
        return "";
      }
      var context = await getClientContextForServdb(servdb);
      if (!context) {
        return "";
      }
      var parent = await context.client.elementInfo(context.sid, context.database, dimensionName, elementName);
      var childId = parent.childIds[Math.floor(childIndex) - 1];
      if (!childId) {
        return "";
      }
      var all = await context.client.dimensionElements(context.sid, context.database, dimensionName);
      var found = all.find(function (item) { return item.id === childId; });
      return found ? found.name : "";
    } catch (error) {
      return toError(error);
    }
  }

  async function PALO_EPARENTCOUNT(servdb, dimensionName, elementName) {
    try {
      var elem = await getElementByName(servdb, dimensionName, elementName);
      if (!elem) {
        return 0;
      }
      return elem.parentIds.length;
    } catch (error) {
      return toError(error);
    }
  }

  async function PALO_EPARENT(servdb, dimensionName, elementName, parentIndex) {
    try {
      if (
        shouldBlockPaloDatacArg(servdb)
        || shouldBlockPaloDatacArg(dimensionName)
        || shouldBlockPaloDatacArg(elementName)
        || shouldBlockPaloDatacArg(parentIndex)
      ) {
        return "";
      }
      var context = await getClientContextForServdb(servdb);
      if (!context) {
        return "";
      }
      var child = await context.client.elementInfo(context.sid, context.database, dimensionName, elementName);
      var parentId = child.parentIds[Math.floor(parentIndex) - 1];
      if (!parentId) {
        return "";
      }
      var all = await context.client.dimensionElements(context.sid, context.database, dimensionName);
      var found = all.find(function (item) { return item.id === parentId; });
      return found ? found.name : "";
    } catch (error) {
      return toError(error);
    }
  }

  async function PALO_ELEVEL(servdb, dimensionName, elementName) {
    try {
      var elem = await getElementByName(servdb, dimensionName, elementName);
      if (!elem) {
        return 0;
      }
      return elem.level;
    } catch (error) {
      return toError(error);
    }
  }

  async function PALO_EINDENT(servdb, dimensionName, elementName) {
    try {
      var elem = await getElementByName(servdb, dimensionName, elementName);
      if (!elem) {
        return 0;
      }
      return elem.indent + 1;
    } catch (error) {
      return toError(error);
    }
  }

  async function PALO_ETYPE(servdb, dimensionName, elementName) {
    try {
      var elem = await getElementByName(servdb, dimensionName, elementName);
      if (!elem) {
        return "";
      }
      if (elem.type === 1) {
        return "numeric";
      }
      if (elem.type === 2) {
        return "string";
      }
      if (elem.type === 4) {
        return "consolidated";
      }
      return "numeric";
    } catch (error) {
      return toError(error);
    }
  }

  async function PALO_EWEIGHT(servdb, dimensionName, parentName, childName) {
    try {
      if (
        shouldBlockPaloDatacArg(servdb)
        || shouldBlockPaloDatacArg(dimensionName)
        || shouldBlockPaloDatacArg(parentName)
        || shouldBlockPaloDatacArg(childName)
      ) {
        return 0;
      }
      var context = await getClientContextForServdb(servdb);
      if (!context) {
        return 0;
      }
      var parent = await context.client.elementInfo(context.sid, context.database, dimensionName, parentName);
      var child = await context.client.elementInfo(context.sid, context.database, dimensionName, childName);
      var idx = parent.childIds.findIndex(function (id) { return id === child.id; });
      if (idx < 0) {
        return 0;
      }
      return parent.weights[idx] || 0;
    } catch (error) {
      return toError(error);
    }
  }

  async function PALO_DATABASE_LIST_DIMENSIONS(servdb) {
    try {
      if (shouldBlockPaloDatacArg(servdb)) {
        return emptyPaloMatrixCell();
      }
      var context = await getClientContextForServdb(servdb);
      if (!context) {
        return emptyPaloMatrixCell();
      }
      var dimensions = await context.client.databaseDimensions(context.sid, context.database);
      return asStringMatrix(dimensions.map(function (item) { return item.name; }));
    } catch (error) {
      return toError(error);
    }
  }

  async function PALO_CUBE_LIST_DIMENSIONS(servdb, cubeName) {
    try {
      if (shouldBlockPaloDatacArg(servdb) || shouldBlockPaloDatacArg(cubeName)) {
        return emptyPaloMatrixCell();
      }
      var context = await getClientContextForServdb(servdb);
      if (!context) {
        return emptyPaloMatrixCell();
      }
      var info = await context.client.cubeInfo(context.sid, context.database, cubeName);
      var allDimensions = await context.client.databaseDimensions(context.sid, context.database);
      var names = info.dimensionIds
        .map(function (id) {
          var found = allDimensions.find(function (d) { return d.id === id; });
          return found ? found.name : null;
        })
        .filter(Boolean);
      return asStringMatrix(names);
    } catch (error) {
      return toError(error);
    }
  }

  async function PALO_DIMENSION_LIST_CUBES(servdb, dimensionName) {
    try {
      if (shouldBlockPaloDatacArg(servdb) || shouldBlockPaloDatacArg(dimensionName)) {
        return emptyPaloMatrixCell();
      }
      var context = await getClientContextForServdb(servdb);
      if (!context) {
        return emptyPaloMatrixCell();
      }
      var cubes = await context.client.dimensionCubes(context.sid, context.database, dimensionName);
      return asStringMatrix(cubes.map(function (item) { return item.name; }));
    } catch (error) {
      return toError(error);
    }
  }

  async function PALO_DIMENSION_LIST_ELEMENTS(servdb, dimensionName) {
    try {
      if (shouldBlockPaloDatacArg(servdb) || shouldBlockPaloDatacArg(dimensionName)) {
        return emptyPaloMatrixCell();
      }
      var context = await getClientContextForServdb(servdb);
      if (!context) {
        return emptyPaloMatrixCell();
      }
      var elements = await context.client.dimensionElements(context.sid, context.database, dimensionName);
      return asStringMatrix(elements.map(function (item) { return item.name; }));
    } catch (error) {
      return toError(error);
    }
  }

  var PALO_TODO_FUNCS = [
    ["AFILTER", "AFILTER", "Le filtre attribut supprime des éléments pour lesquels les lignes attribut définies ne correspondent pas."],
    ["AGR_CHILDREN", "AGR.CHILDREN", "Indique à une agrégation de prendre les enfants directs."],
    ["AGR_LEAVES", "AGR.LEAVES", "Indique à une agrégation de prendre les éléments de base situés sous l'élément."],
    ["AGR_SELF", "AGR.SELF", "Indique à une agrégation de prendre l'élément lui-même."],
    ["CONNECTIONUSER", "CONNECTIONUSER", "Retourne l'utilisateur de la connexion"],
    ["COORD", "COORD", "Regroupe des coordonnées dans une matrice pour DATA.SUM et les autres agrégats."],
    ["CUBE", "CUBE", "Désigne la zone de cube lue par une vue Palo."],
    ["DATA", "DATA", "Lit la valeur de la cellule de données indiquée à partir du cube."],
    ["DATA_AVG", "DATA.AVG", "Renvoie la valeur moyenne d'une plage de cellules du cube."],
    ["DATA_CNT", "DATA.CNT", "Renvoie le nombre de cellules d'une plage de cellules du cube."],
    ["DATA_MAX", "DATA.MAX", "Renvoie la valeur maximale d'une plage de cellules du cube."],
    ["DATA_MIN", "DATA.MIN", "Renvoie la valeur minimale d'une plage de cellules du cube."],
    ["DATA_SUM", "DATA.SUM", "Renvoie la somme d'une plage de cellules du cube."],
    ["DATAA", "DATAA", "Lit une valeur numérique en acceptant un alias à la place du nom d'élément."],
    ["DATAAC", "DATAAC", "Comme DATAA, et signale une erreur si la cellule ne peut pas être lue."],
    ["DATAAT", "DATAAT", "Lit une valeur texte en acceptant un alias à la place du nom d'élément."],
    ["DATAATC", "DATAATC", "Comme DATAAT, et signale une erreur si la cellule texte ne peut pas être lue."],
    ["DATABASE_LIST_CUBES", "DATABASE_LIST_CUBES", "La fonction établit, dans une plage que vous aurez prédéfinie, une liste des cubes d’une base de données."],
    ["DATABASE_LIST_DIMENSIONS_EXT", "DATABASE_LIST_DIMENSIONS_EXT", "Liste les dimensions d'une base avec leurs attributs : identifiant, type et nombre d'éléments."],
    ["DATAT", "DATAT", "Lit la valeur texte de la cellule indiquée. Renvoie une chaîne vide si la lecture échoue."],
    ["DATATC", "DATATC", "Lit la valeur texte de la cellule indiquée et renvoie une erreur si la lecture échoue. Le texte est limité à 255 caractères."],
    ["DATAV", "DATAV", "Lit une cellule du cube et se place dans une formule matricielle Excel."],
    ["DATAX", "DATAX", "Lit une cellule du cube, et ne se calcule que lorsque sa feuille est active."],
    ["DFILTER", "DFILTER", "Calcule pour chaque élément d’un sous-ensemble une valeur de données. La valeur peut être numérique ou un String."],
    ["DIMENSION", "DIMENSION", "Désigne une dimension dans la définition d'une vue Palo."],
    ["EADD", "EADD", "Ajoute un élément, ou un lien de consolidation, dans une dimension."],
    ["EALIAS", "EALIAS", "Renvoie l'alias d'un élément, ou les éléments qui portent l'alias indiqué."],
    ["EDELETE", "EDELETE", "Supprime l'élément spécifié."],
    ["EFIRST", "EFIRST", "Renvoie le premier élément de la dimension, ou le premier enfant du parent indiqué."],
    ["EINDEX", "EINDEX", "Renvoie la position de l’élément indiqué dans la dimension."],
    ["EISCHILD", "EISCHILD", "Vérifie si un élément consolidé contient l’élément indiqué."],
    ["ELEMENT_LIST_ANCESTORS", "ELEMENT_LIST_ANCESTORS", "Listes de tous les ancêtres de l'élément spécifié"],
    ["ELEMENT_LIST_CHILDREN", "ELEMENT_LIST_CHILDREN", "Liste les enfants directs d'un élément."],
    ["ELEMENT_LIST_CONSOLIDATION_ELEMENTS", "ELEMENT_LIST_CONSOLIDATION_ELEMENTS", "Retourne une liste avec tous les éléments enfants d'un élément consolidé."],
    ["ELEMENT_LIST_DESCENDANTS", "ELEMENT_LIST_DESCENDANTS", "Liste tous les descendants."],
    ["ELEMENT_LIST_PARENTS", "ELEMENT_LIST_PARENTS", "Retourne une liste de tous les parents d'un élément"],
    ["ELEMENT_LIST_SIBLINGS", "ELEMENT_LIST_SIBLINGS", "Retourne une liste de tous les éléments qui sont sur le même niveau hiérarchique que l’élément spécifié"],
    ["ENEXT", "ENEXT", "Renvoie le nom de l'élément qui suit l'élément indiqué."],
    ["EPREV", "EPREV", "Renvoie le nom de l'élément qui précède l'élément indiqué."],
    ["ERENAME", "ERENAME", "Renomme l’élément indiqué."],
    ["ERROR_LOG", "ERROR_LOG", "Journalise une ligne d'import refusée par le serveur."],
    ["ESIBLING", "ESIBLING", "Renvoie le nom de l’élément frère défini."],
    ["ETOPLEVEL", "ETOPLEVEL", "Renvoie le niveau le plus élevé que possède un élément dans la hiérarchie de consolidation de la dimension indiquée."],
    ["EUPDATE", "EUPDATE", "Met à jour le type, le poids ou la position d'un élément existant."],
    ["EXPANDTYPE", "EXPANDTYPE", "Regroupe les modes d'expansion (SELF, CHILDREN, LEAVES) d'une agrégation."],
    ["HFILTER", "HFILTER", "Filtre un sous-ensemble selon la place de l'élément dans la hiérarchie : niveau, parents, enfants ou feuilles."],
    ["PICKLIST", "PICKLIST", "Peut ajouter des éléments à un sous-ensemble qui ne peuvent en aucun cas être à nouveau supprimés. En outre, il est possible de définir par ces derniers une multitude d’éléments qui feront office de filtre."],
    ["PING", "PING", "Vérifie que le serveur Jedox OLAP de la connexion répond."],
    ["REGISTER", "REGISTER", "Enregistre un serveur PALO pour qu'une formule puisse s'y connecter."],
    ["REMOVE_CONNECTION", "REMOVE_CONNECTION", "Ferme la connexion nommée."],
    ["ROOT_LIST_DATABASES", "ROOT_LIST_DATABASES", "Donne une liste avec toutes les bases de données du serveur spécifié."],
    ["ROOT_LIST_DATABASES_EXT", "ROOT_LIST_DATABASES_EXT", "Liste les bases du serveur avec leurs attributs."],
    ["SERVERSUBSET", "SERVERSUBSET", "Ancien nom de SERVER_SUBSET : calcule un sous-ensemble d'éléments sur le serveur."],
    ["SERVER_INFO", "SERVER_INFO", "Retourne des informations sur le serveur Jedox OLAP d'une connexion donnée."],
    ["SERVER_SUBSET", "SERVER_SUBSET", "Renvoie un sous-ensemble d’éléments de dimension."],
    ["SETDATAA", "SETDATAA", "Écrit une valeur en acceptant un alias à la place du nom d'élément."],
    ["SETDATAAIF", "SETDATAAIF", "Écrit une valeur, en acceptant un alias, lorsque la condition est vraie."],
    ["SETDATAA_WRAP", "SETDATAA_WRAP", "Variante interne de SETDATAA : enveloppe l'écriture et journalise l'erreur."],
    ["SETDATAIF", "SETDATAIF", "Écrit une valeur dans une cellule du cube lorsque la condition est vraie."],
    ["SETDATA_BULK", "SETDATA_BULK", "Écrit plusieurs cellules du cube dans un seul appel, sans interpréter de préfixe."],
    ["SETDATA_ERROR", "SETDATA_ERROR", "Journalise une ligne d'import refusée par le serveur."],
    ["SETDATA_WRAP", "SETDATA_WRAP", "Variante interne de SETDATA : enveloppe l'écriture et journalise l'erreur au lieu de la remonter."],
    ["SORT", "SORT", "Définit l'affichage des éléments du sous-ensemble, par exemple en ordre de définition ou en ordre hiérarchique."],
    ["SUBCUBE", "SUBCUBE", "Décrit la tranche de cube lue par DFILTER. La dimension filtrée reste vide, les autres reçoivent un ou plusieurs éléments."],
    ["SUBSET", "SUBSET", "Renvoie un sous-ensemble d’éléments de dimension."],
    ["SUBSETSIZE", "SUBSETSIZE", "Renvoie la longueur d’un extrait d’ensemble d’un élément de dimension."],
    ["SUBSETX", "SUBSETX", "Renvoie un sous-ensemble d’éléments de dimension si est appelée depuis une feuille active."],
    ["TFILTER", "TFILTER", "Éléments dont les noms correspondent à des modèles définis, sont conservés. Le reste est supprimé de l’affichage. Plusieurs critères de filtre sont indiqués dans la matrice."],
    ["USERGROUPS", "USERGROUPS", "Retourne le groupe d'utilisateurs de la connexion"],
    ["VIEW_AREAGET", "VIEW.AREAGET", "Renvoie les valeurs de la zone d'une vue Palo."],
    ["VIEW_AXISDEF", "VIEW.AXISDEF", "Définit un axe (lignes ou colonnes) d'une vue Palo."],
    ["VIEW_AXISGET", "VIEW.AXISGET", "Renvoie les éléments placés sur un axe de vue Palo."],
    ["VIEW_AXISGETINDEX", "VIEW.AXISGETINDEX", "Renvoie la position d'un élément sur un axe de vue Palo."],
    ["VIEW_SUBSETDEF", "VIEW.SUBSETDEF", "Définit le sous-ensemble utilisé par un axe de vue Palo."],
  ];

  function paloTodoMessage(excelName, description) {
    return "#PALO! TODO " + excelName + " : " + description;
  }

  function registerPaloTodoFunctions() {
    var i;
    for (i = 0; i < PALO_TODO_FUNCS.length; i += 1) {
      (function (item) {
        CustomFunctions.associate(item[0], function () {
          return paloTodoMessage("PALO." + item[1], item[2]);
        });
      })(PALO_TODO_FUNCS[i]);
    }
  }

  function registerCustomFunctions() {
    if (typeof CustomFunctions === "undefined") {
      return false;
    }
    if (gBoot.__PALO_CF_REGISTERED__) {
      return true;
    }
    CustomFunctions.associate("RUNTIME_DIAG", RUNTIME_DIAG);
    CustomFunctions.associate("STORAGE_DIAG", STORAGE_DIAG);
    CustomFunctions.associate("DATAC", DATAC);
    CustomFunctions.associate("SETDATA", PALO_SETDATA);
    CustomFunctions.associate("ENAME", ENAME);
    CustomFunctions.associate("ECOUNT", PALO_ECOUNT);
    CustomFunctions.associate("ECHILDCOUNT", PALO_ECHILDCOUNT);
    CustomFunctions.associate("ECHILD", PALO_ECHILD);
    CustomFunctions.associate("EPARENTCOUNT", PALO_EPARENTCOUNT);
    CustomFunctions.associate("EPARENT", PALO_EPARENT);
    CustomFunctions.associate("ELEVEL", PALO_ELEVEL);
    CustomFunctions.associate("EINDENT", PALO_EINDENT);
    CustomFunctions.associate("ETYPE", PALO_ETYPE);
    CustomFunctions.associate("EWEIGHT", PALO_EWEIGHT);
    CustomFunctions.associate("DATABASE_LIST_DIMENSIONS", PALO_DATABASE_LIST_DIMENSIONS);
    CustomFunctions.associate("CUBE_LIST_DIMENSIONS", PALO_CUBE_LIST_DIMENSIONS);
    CustomFunctions.associate("DIMENSION_LIST_CUBES", PALO_DIMENSION_LIST_CUBES);
    CustomFunctions.associate("DIMENSION_LIST_ELEMENTS", PALO_DIMENSION_LIST_ELEMENTS);
    registerPaloTodoFunctions();
    gBoot.__PALO_CF_REGISTERED__ = true;
    gBoot.__PALO_CF_REGISTERED_AT__ = Date.now();
    paloWarmStorageAfterRegister();
    return true;
  }

  /** CustomFunctions peut arriver apres le premier parse du script (worker / shared runtime). */
  function scheduleRegisterCustomFunctions() {
    if (registerCustomFunctions()) {
      return;
    }
    var deadline = Date.now() + 30000;
    function tick() {
      if (registerCustomFunctions()) {
        return;
      }
      if (Date.now() > deadline) {
        if (typeof console !== "undefined" && console.error) {
          console.error(
            "[PaloOffice] CustomFunctions indisponible apres 30s — aucune fonction PALO.* enregistree. "
            + "Verifiez le manifeste et rechargez le complement."
          );
        }
        return;
      }
      setTimeout(tick, 50);
    }
    tick();
  }

  gBoot.__paloScheduleRegisterCustomFunctions = scheduleRegisterCustomFunctions;

  // Associer immediatement : attendre Office.onReady provoque #VALEUR! sur Excel Online
  // (formules evaluees avant la fin de onReady / associate trop tardif).
  scheduleRegisterCustomFunctions();
  if (typeof Office !== "undefined" && Office && typeof Office.onReady === "function") {
    Office.onReady(function () {
      scheduleRegisterCustomFunctions();
    });
  }
})();

