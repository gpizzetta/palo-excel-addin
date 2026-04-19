(function () {
	var ctx = { apiBase: "", sid: "" };

	function parseQuery() {
		var p = new URLSearchParams(window.location.search);
		ctx.apiBase = (p.get("apiBase") || "").trim();
		ctx.sid = p.get("sid") || "";
	}

	function showErr(msg) {
		var el = document.getElementById("err");
		if (el) {
			el.textContent = msg || "Erreur";
			el.className = "show";
		}
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

	function tryCreateUser(name, password) {
		var q1 = new URLSearchParams({
			sid: ctx.sid,
			new_name: name,
			password: password,
		});
		var urls = [
			ctx.apiBase + "/user/create?" + q1.toString(),
			ctx.apiBase + "/server/user_create?" + q1.toString(),
		];
		function attempt(i) {
			if (i >= urls.length) {
				return Promise.reject(
					new Error(
						"Aucun endpoint de création d’utilisateur n’a répondu (/user/create, /server/user_create).",
					),
				);
			}
			return fetchCsv(urls[i]).catch(function () {
				return attempt(i + 1);
			});
		}
		return attempt(0);
	}

	function save() {
		parseQuery();
		var name = document.getElementById("inputName").value.trim();
		var p1 = document.getElementById("inputPass").value;
		var p2 = document.getElementById("inputPass2").value;
		if (!name) {
			showErr("Indiquez un nom d’utilisateur.");
			return;
		}
		if (!p1) {
			showErr("Indiquez un mot de passe.");
			return;
		}
		if (p1 !== p2) {
			showErr("Les mots de passe ne correspondent pas.");
			return;
		}
		var btn = document.getElementById("btnOk");
		btn.disabled = true;
		tryCreateUser(name, p1)
			.then(function () {
				if (Office && Office.context && Office.context.ui && Office.context.ui.messageParent) {
					Office.context.ui.messageParent("refresh", { targetOrigin: "*" });
				}
			})
			.catch(function (err) {
				showErr(err && err.message ? err.message : String(err));
				btn.disabled = false;
			});
	}

	function cancel() {
		if (Office && Office.context && Office.context.ui && Office.context.ui.messageParent) {
			Office.context.ui.messageParent("cancel", { targetOrigin: "*" });
		}
	}

	Office.onReady(function () {
		document.getElementById("btnOk").addEventListener("click", save);
		document.getElementById("btnCancel").addEventListener("click", cancel);
	});
})();
