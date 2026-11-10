(function taskpaneBootstrap() {
  var PLUGIN_VERSION = "1.0.3.32";
  var PALO_CDN_BASE = "https://gpizzetta.github.io/palo-excel-addin";
  var manager = null;

  function setText(id, message) {
    var node = document.getElementById(id);
    if (node) {
      node.textContent = message;
    }
  }

  /**
   * @param {"neutral"|"ok"|"error"} kind
   */
  function setStatusState(kind, message) {
    var logEl = document.getElementById("status-log");
    var iconEl = document.getElementById("status-icon");
    var bannerEl = document.getElementById("status-banner");
    if (logEl) {
      logEl.textContent = String(message || "");
    }
    if (bannerEl) {
      bannerEl.classList.remove("status-banner--ok", "status-banner--error");
      if (kind === "ok") {
        bannerEl.classList.add("status-banner--ok");
      } else if (kind === "error") {
        bannerEl.classList.add("status-banner--error");
      }
    }
    if (!iconEl) {
      return;
    }
    iconEl.className = "status-icon";
    iconEl.removeAttribute("aria-label");
    if (kind === "ok") {
      iconEl.classList.add("status-icon--ok");
      iconEl.textContent = "\u2714";
      iconEl.setAttribute("aria-label", "Connexion reussie");
      iconEl.removeAttribute("aria-hidden");
    } else if (kind === "error") {
      iconEl.classList.add("status-icon--fail");
      iconEl.textContent = "\u2716";
      iconEl.setAttribute("aria-label", "Echec de connexion");
      iconEl.removeAttribute("aria-hidden");
    } else {
      iconEl.textContent = "";
      iconEl.setAttribute("aria-hidden", "true");
    }
  }

  function status(message, kind) {
    setStatusState(kind || "neutral", message);
  }

  function getManager() {
    if (manager) {
      return manager;
    }
    if (!window.PaloOffice || typeof window.PaloOffice.createConnectionManager !== "function") {
      throw new Error("PaloOffice indisponible (palo-api.js non charge ?).");
    }
    manager = window.PaloOffice.createConnectionManager();
    return manager;
  }

  function getValue(id) {
    var node = document.getElementById(id);
    return node ? String(node.value || "") : "";
  }

  function refreshConnectionList() {
    var select = document.getElementById("connection-list");
    if (!select) {
      return;
    }

    var active = null;
    var all = [];
    try {
      active = getManager().getActiveConnectionName();
      all = getManager().listConnections();
    } catch (e) {
      status(e && e.message ? e.message : String(e));
      active = null;
      all = [];
    }
    select.innerHTML = "";

    var empty = document.createElement("option");
    empty.value = "";
    empty.textContent = "-- Selectionner --";
    select.appendChild(empty);

    all.forEach(function (conn) {
      var opt = document.createElement("option");
      opt.value = conn.name;
      opt.textContent = conn.name + " (" + (conn.baseUrl || "https://palo.example.com") + ")";
      if (active && active === conn.name) {
        opt.selected = true;
      }
      select.appendChild(opt);
    });
  }

  function getSelectedConnection() {
    var select = document.getElementById("connection-list");
    return select ? String(select.value || "") : "";
  }

  function ensureActiveConnectionSelected() {
    var active = null;
    try {
      active = getManager().getActiveConnectionName();
    } catch (_e) {
      active = null;
    }
    if (active) {
      return active;
    }
    var all = [];
    try {
      all = getManager().listConnections();
    } catch (_e2) {
      all = [];
    }
    if (all.length > 0 && all[0].name) {
      try {
        getManager().setActiveConnectionName(all[0].name);
      } catch (_e3) {
      }
      refreshConnectionList();
      return String(all[0].name);
    }
    return "";
  }

  function flushStorageToOfficeRuntime() {
    try {
      var po = window.PaloOffice;
      if (po && typeof po.paloFlushStorageToOfficeRuntime === "function") {
        return po.paloFlushStorageToOfficeRuntime();
      }
    } catch (_flush) {
    }
    return Promise.resolve();
  }

  function saveConnection() {
    status("Enregistrement connexion...");
    try {
      var name = getValue("conn-name").trim();
      var baseUrl = getValue("conn-url").trim();
      var user = getValue("conn-user").trim();
      var password = getValue("conn-password");
      var debug = false;
      try {
        var current = getManager().getConnection(name);
        debug = Boolean(current && current.debug);
      } catch (_keepDebug) {
      }
      if (!name) {
        status("Nom de connexion requis.");
        return;
      }
      if (!baseUrl) {
        status("URL Palo requise.");
        return;
      }
      if (!user) {
        status("Utilisateur requis.");
        return;
      }
      getManager().saveConnection({
        name: name,
        baseUrl: baseUrl,
        user: user,
        password: password,
        debug: debug
      });
      getManager().setActiveConnectionName(name);
      refreshConnectionList();
      flushStorageToOfficeRuntime().then(function () {
        status("Connexion " + name + " enregistree.", "ok");
      }).catch(function () {
        status("Connexion " + name + " enregistree.", "ok");
      });
    } catch (error) {
      status(error && error.message ? error.message : String(error));
    }
  }

  async function testConnection() {
    var name = getSelectedConnection();
    if (!name) {
      status("Aucune connexion selectionnee.");
      return;
    }
    status("Test connexion en cours...");
    var result = await getManager().testConnection(name);
    if (result.ok) {
      status(result.details, "ok");
    } else {
      status(result.details, "error");
    }
  }

  function deleteConnection() {
    var name = getSelectedConnection();
    if (!name) {
      status("Aucune connexion selectionnee.");
      return;
    }
    getManager().deleteConnection(name);
    refreshConnectionList();
    status("Connexion " + name + " supprimee.");
  }

  function refreshVersionFromServer() {
    var assetsEl = document.getElementById("plugin-version");
    if (assetsEl) {
      assetsEl.textContent = PLUGIN_VERSION;
    }

    var versionUrl = PALO_CDN_BASE + "/version.json?_=" + String(Date.now());
    fetch(versionUrl, { cache: "no-store" })
      .then(function (res) {
        if (!res.ok) {
          throw new Error("HTTP " + res.status);
        }
        return res.json();
      })
      .then(function (data) {
        var live = data && data.version ? String(data.version) : "";
        if (!live || !assetsEl) {
          return;
        }
        if (live === PLUGIN_VERSION) {
          assetsEl.textContent = live;
          return;
        }
        assetsEl.textContent = live + " (charges " + PLUGIN_VERSION + " — recharger)";
      })
      .catch(function () {
        if (assetsEl) {
          assetsEl.textContent = PLUGIN_VERSION;
        }
      });
  }

  function bindUi() {
    refreshVersionFromServer();
    try {
      var po = window.PaloOffice;
      if (po && typeof po.paloEnsureStorageReady === "function") {
        po.paloEnsureStorageReady().catch(function () {});
      }
    } catch (_storageInit) {
    }
    refreshConnectionList();

    var saveBtn = document.getElementById("save-connection");
    var testBtn = document.getElementById("test-connection");
    var delBtn = document.getElementById("delete-connection");
    var list = document.getElementById("connection-list");

    if (saveBtn) {
      saveBtn.addEventListener("click", saveConnection);
    }
    if (testBtn) {
      testBtn.addEventListener("click", function () {
        testConnection().catch(function (error) {
          status(error && error.message ? error.message : String(error), "error");
        });
      });
    }
    if (delBtn) {
      delBtn.addEventListener("click", deleteConnection);
    }
    if (list) {
      list.addEventListener("change", function () {
        var value = getSelectedConnection();
        if (value) {
          try {
            getManager().setActiveConnectionName(value);
          } catch (e) {
            status(e && e.message ? e.message : String(e));
          }
          flushStorageToOfficeRuntime().then(function () {
            status("Connexion active : " + value + ".", "ok");
          }).catch(function () {
            status("Connexion active : " + value + ".", "ok");
          });
        }
      });
    }

    var autoSelected = ensureActiveConnectionSelected();
    if (!autoSelected) {
      status("Aucune connexion. Renseignez le serveur, puis Enregistrer.");
    } else {
      flushStorageToOfficeRuntime().catch(function () {});
    }
  }

  // Office.onReady peut ne pas se declencher selon le host/webview: on bind l'UI au DOM aussi.
  try {
    if (typeof Office !== "undefined" && Office && typeof Office.onReady === "function") {
      Office.onReady(function () {
        bindUi();
      });
    }
  } catch (_e) {
  }
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bindUi);
  } else {
    bindUi();
  }
})();
