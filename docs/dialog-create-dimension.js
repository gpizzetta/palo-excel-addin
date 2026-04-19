(function () {
	var ctx = { apiBase: "", sid: "", name_database: "" };

	function parseQuery() {
		var p = new URLSearchParams(window.location.search);
		ctx.apiBase = (p.get("apiBase") || "").trim();
		ctx.sid = p.get("sid") || "";
		ctx.name_database = p.get("name_database") || "";
	}

	function showErr(msg) {
		var el = document.getElementById("err");
		if (el) {
			el.textContent = msg || "Erreur";
			el.className = "show";
		}
	}

	function save() {
		parseQuery();
		var name = document.getElementById("inputName").value.trim();
		if (!name) {
			showErr("Indiquez un nom.");
			return;
		}
		var q = new URLSearchParams({
			sid: ctx.sid,
			name_database: ctx.name_database,
			new_name: name,
			type: "0",
		});
		var url = ctx.apiBase + "/dimension/create?" + q.toString();
		var btn = document.getElementById("btnOk");
		btn.disabled = true;
		fetch(url, { method: "GET", mode: "cors", cache: "no-store", credentials: "omit" })
			.then(function (res) {
				return res.text().then(function (text) {
					if (!res.ok) {
						throw new Error("HTTP " + res.status + " — " + text.slice(0, 500));
					}
					if (Office && Office.context && Office.context.ui && Office.context.ui.messageParent) {
						Office.context.ui.messageParent("refresh", { targetOrigin: "*" });
					}
				});
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
		parseQuery();
		var sub = document.getElementById("subtitle");
		if (sub && ctx.name_database) {
			sub.textContent = "Base : " + ctx.name_database;
		}
		document.getElementById("btnOk").addEventListener("click", save);
		document.getElementById("btnCancel").addEventListener("click", cancel);
		document.getElementById("inputName").addEventListener("keydown", function (ev) {
			if (ev.key === "Enter") {
				ev.preventDefault();
				save();
			}
		});
	});
})();
