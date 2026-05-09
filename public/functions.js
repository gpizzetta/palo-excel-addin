(function paloFunctionsBootstrap() {
  var connectionManager = null;
  var datacRequestSeq = 0;

  function paloFnTrace() {
    return typeof window !== "undefined" && Boolean(window.PALO_DEBUG);
  }

  function getConnectionManager() {
    if (!connectionManager) {
      connectionManager = window.PaloOffice.createConnectionManager();
    }
    return connectionManager;
  }

  function toError(error) {
    var message = error && error.message ? error.message : String(error);
    return "#PALO! " + message;
  }

  function isDebugEnabledForServdb(servdb) {
    try {
      var manager = getConnectionManager();
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

  function getLastApiUrl() {
    if (window.PaloOffice && typeof window.PaloOffice.getLastApiUrl === "function") {
      return window.PaloOffice.getLastApiUrl();
    }
    return "";
  }

  function nextDatacRequestId() {
    datacRequestSeq += 1;
    return "datac-" + String(datacRequestSeq);
  }

  function traceDatac(eventName, payload) {
    if (window.PaloOffice && typeof window.PaloOffice.trace === "function") {
      window.PaloOffice.trace(eventName, payload || {});
    }
  }

  function addOperandToNumber(value) {
    if (value === null || value === undefined || value === "") {
      return 0;
    }
    var n = Number(value);
    return Number.isNaN(n) ? 0 : n;
  }

  /** Pas de connexion Palo : ne pas utiliser getConnectionManager. */
  function ADD(cellule1, cellule2) {
    return addOperandToNumber(cellule1) + addOperandToNumber(cellule2);
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

  async function getElementByName(servdb, dimension, elementName) {
    if (shouldBlockPaloDatacArg(servdb) || shouldBlockPaloDatacArg(dimension) || shouldBlockPaloDatacArg(elementName)) {
      return null;
    }
    var context = await getConnectionManager().getClientAndContext(servdb);
    return context.client.elementInfo(context.sid, context.database, dimension, elementName);
  }

  async function DATAC(servdb, cubeName) {
    var coordinates = sanitizePaloCoordinates(Array.prototype.slice.call(arguments, 2));
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
      var manager = getConnectionManager();
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
        console.info("[PaloOffice DATAC] cell/value params", {
          requestId: requestId,
          name_database: context.database,
          name_cube: cubeName,
          path: idPath
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
        idPath: idPath
      });
      var value = await manager.requestCellValueBatched(
        context.connectionName,
        context.sid,
        context.client,
        context.database,
        cubeName,
        "",
        coordinates,
        {
          requestId: requestId,
          coordinates: coordinates.map(function (coord) {
            return String(coerceExcelScalarArg(coord));
          })
        }
      );
      if (paloFnTrace()) {
        console.info("[PaloOffice DATAC] fin OK", { requestId: requestId, value: value });
      }
      traceDatac("datac-end", {
        requestId: requestId,
        value: value
      });
      return value === null ? "" : value;
    } catch (error) {
      var msg = error && error.message ? String(error.message) : String(error);
      if (paloFnTrace() || (typeof window !== "undefined" && window.PALO_LOG_HTTP)) {
        console.warn("[PaloOffice DATAC] erreur", { requestId: requestId, message: msg });
      }
      traceDatac("datac-error", {
        requestId: requestId,
        message: msg
      });
      if (
        isDebugEnabledForServdb(servdb)
        || msg.indexOf("Timeout HTTP") !== -1
        || msg.indexOf("HTTP ") !== -1
        || msg.indexOf("Impossible de joindre") !== -1
      ) {
        return toError(error);
      }
      return "";
    }
  }

  async function DATAC_TEST() {
    var manager = getConnectionManager();
    var active = manager.getActiveConnectionName();
    if (!active) {
      return "#PALO! Aucune connexion active selectionnee. | url=(url indisponible)";
    }
    var testServdb = String(active) + "/DWH";
    var testCube = "PP_BUDGET";
    var testCoordinates = ["16VS", "ACTIVISTA SA", "Chiffres d'affaires HT", "2025", "Décembre", "Réalisé"];
    var value = await DATAC(
      testServdb,
      testCube,
      testCoordinates[0],
      testCoordinates[1],
      testCoordinates[2],
      testCoordinates[3],
      testCoordinates[4],
      testCoordinates[5]
    );
    var url = getLastApiUrl();
    if (typeof value === "string" && value.indexOf("#PALO!") === 0) {
      return value + " | url=" + (url || "(url indisponible)");
    }
    return "url=" + (url || "(url indisponible)") + " | value=" + String(value);
  }

  async function PALO_SETDATA(value, splash, servdb, cubeName) {
    var coordinates = sanitizePaloCoordinates(Array.prototype.slice.call(arguments, 4));
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
      var manager = getConnectionManager();
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

      var manager = getConnectionManager();
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
      if (shouldBlockPaloErrorOnly(servdb) || shouldBlockPaloErrorOnly(dimensionName) || shouldBlockPaloErrorOnly(index)) {
        return "";
      }
      var context = await manager.getClientAndContext(serv);
      try {
        var elem = await context.client.elementInfo(context.sid, context.database, dim, query);
        return String(elem.name);
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
        return String(found.name);
      }
    } catch (error) {
      return toError(error);
    }
  }

  async function PALO_ECOUNT(servdb, dimensionName) {
    try {
      if (shouldBlockPaloDatacArg(servdb) || shouldBlockPaloDatacArg(dimensionName)) {
        return 0;
      }
      var context = await getConnectionManager().getClientAndContext(servdb);
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
      var context = await getConnectionManager().getClientAndContext(servdb);
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
      var context = await getConnectionManager().getClientAndContext(servdb);
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
      var context = await getConnectionManager().getClientAndContext(servdb);
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
      var context = await getConnectionManager().getClientAndContext(servdb);
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
      var context = await getConnectionManager().getClientAndContext(servdb);
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
      var context = await getConnectionManager().getClientAndContext(servdb);
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
      var context = await getConnectionManager().getClientAndContext(servdb);
      var elements = await context.client.dimensionElements(context.sid, context.database, dimensionName);
      return asStringMatrix(elements.map(function (item) { return item.name; }));
    } catch (error) {
      return toError(error);
    }
  }

  if (typeof CustomFunctions !== "undefined") {
    CustomFunctions.associate("ADD", ADD);
    CustomFunctions.associate("DATAC", DATAC);
    CustomFunctions.associate("DATAC_TEST", DATAC_TEST);
    CustomFunctions.associate("PALO_SETDATA", PALO_SETDATA);
    CustomFunctions.associate("ENAME", ENAME);
    CustomFunctions.associate("PALO_ECOUNT", PALO_ECOUNT);
    CustomFunctions.associate("PALO_ECHILDCOUNT", PALO_ECHILDCOUNT);
    CustomFunctions.associate("PALO_ECHILD", PALO_ECHILD);
    CustomFunctions.associate("PALO_EPARENTCOUNT", PALO_EPARENTCOUNT);
    CustomFunctions.associate("PALO_EPARENT", PALO_EPARENT);
    CustomFunctions.associate("PALO_ELEVEL", PALO_ELEVEL);
    CustomFunctions.associate("PALO_EINDENT", PALO_EINDENT);
    CustomFunctions.associate("PALO_ETYPE", PALO_ETYPE);
    CustomFunctions.associate("PALO_EWEIGHT", PALO_EWEIGHT);
    CustomFunctions.associate("PALO_DATABASE_LIST_DIMENSIONS", PALO_DATABASE_LIST_DIMENSIONS);
    CustomFunctions.associate("PALO_CUBE_LIST_DIMENSIONS", PALO_CUBE_LIST_DIMENSIONS);
    CustomFunctions.associate("PALO_DIMENSION_LIST_CUBES", PALO_DIMENSION_LIST_CUBES);
    CustomFunctions.associate("PALO_DIMENSION_LIST_ELEMENTS", PALO_DIMENSION_LIST_ELEMENTS);
  }
})();

