/** Doit rester aligné avec <Version> dans docs/manifest.xml */
var ADDIN_VERSION = "1.0.21.0";

function hello() {
	return "hello world";
}

function version() {
	return ADDIN_VERSION;
}

/**
 * GET sur l’URL en mode CORS — pour tester HTTP/HTTPS et les en-têtes côté serveur Palo.
 * Si la réponse est bloquée (CORS, réseau, certificat), le message d’erreur l’indique souvent.
 */
function info(url) {
	if (url === undefined || url === null) {
		return "Indiquez une URL complète, ex. https://127.0.0.1:7777/server/info";
	}
	var s = String(url).trim();
	if (!s) {
		return "URL vide";
	}
	return fetch(s, {
		method: "GET",
		mode: "cors",
		cache: "no-store",
	})
		.then(function (res) {
			return (
				"OK — HTTP " +
				res.status +
				" " +
				res.statusText +
				" — Content-Type: " +
				(res.headers.get("content-type") || "(none)")
			);
		})
		.catch(function (err) {
			var msg = err && err.message ? err.message : String(err);
			return (
				"Échec: " +
				msg +
				" — (souvent CORS, réseau ou certificat ; serveur Palo en HTTPS si besoin)"
			);
		});
}

CustomFunctions.associate("HELLO", hello);
CustomFunctions.associate("VERSION", version);
CustomFunctions.associate("INFO", info);
