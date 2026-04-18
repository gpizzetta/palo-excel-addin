(function () {
	var KEYS = {
		url: "palo_connection_url",
		username: "palo_connection_username",
		password: "palo_connection_password",
	};

	function setStatus(msg, kind) {
		var el = document.getElementById("status");
		el.textContent = msg || "";
		el.className = kind || "";
	}

	function loadSettings() {
		if (!Office || !Office.context || !Office.context.document || !Office.context.document.settings) {
			setStatus("Office.js indisponible.", "err");
			return;
		}
		var s = Office.context.document.settings;
		var urlEl = document.getElementById("url");
		var userEl = document.getElementById("username");
		var passEl = document.getElementById("password");
		try {
			if (s.get(KEYS.url)) urlEl.value = s.get(KEYS.url);
			if (s.get(KEYS.username)) userEl.value = s.get(KEYS.username);
			if (s.get(KEYS.password)) passEl.value = s.get(KEYS.password);
		} catch (e) {
			setStatus("Lecture des paramètres : " + (e.message || e), "err");
		}
	}

	function saveSettings(ev) {
		ev.preventDefault();
		setStatus("");
		var s = Office.context.document.settings;
		var url = document.getElementById("url").value.trim();
		var username = document.getElementById("username").value.trim();
		var password = document.getElementById("password").value;

		var btn = document.getElementById("btnSave");
		btn.disabled = true;

		try {
			s.set(KEYS.url, url);
			s.set(KEYS.username, username);
			s.set(KEYS.password, password);
		} catch (e) {
			setStatus("Erreur : " + (e.message || e), "err");
			btn.disabled = false;
			return;
		}

		s.saveAsync(function (asyncResult) {
			btn.disabled = false;
			if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
				setStatus("Connexion enregistrée pour ce classeur.", "ok");
			} else {
				setStatus(
					"Échec de l’enregistrement : " + (asyncResult.error && asyncResult.error.message),
					"err",
				);
			}
		});
	}

	Office.onReady(function () {
		document.getElementById("form").addEventListener("submit", saveSettings);
		loadSettings();
	});
})();
