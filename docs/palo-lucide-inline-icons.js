/**
 * Icônes SVG inline dérivées de Lucide (https://lucide.dev) — licence ISC
 * Voir https://github.com/lucide-icons/lucide/blob/main/LICENSE
 *
 * Aucune police externe ni CDN : compatible compléments Office (CSP).
 */
(function (g) {
	var head =
		'<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" focusable="false" aria-hidden="true">';
	var foot = "</svg>";

	/** Corbeille — équivalent Lucide « trash-2 » */
	g.paloIconSvgTrash2 = function () {
		return (
			head +
			'<path d="M3 6h18"/>' +
			'<path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/>' +
			'<path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>' +
			'<line x1="10" x2="10" y1="11" y2="17"/>' +
			'<line x1="14" x2="14" y1="11" y2="17"/>' +
			foot
		);
	};

	/** Calques — équivalent Lucide « layers » (hiérarchie / consolidation) */
	g.paloIconSvgLayers = function () {
		return (
			head +
			'<path d="m12.83 2.18a2 2 0 0 0-1.66 0L2.6 6.08a1 1 0 0 0 0 1.83l8.58 3.91a2 2 0 0 0 1.66 0l8.58-3.9a1 1 0 0 0 0-1.83Z"/>' +
			'<path d="m22 17.65-9.17 4.16a2 2 0 0 1-1.66 0L2 17.65"/>' +
			'<path d="m22 12.65-9.17 4.16a2 2 0 0 1-1.66 0L2 12.65"/>' +
			foot
		);
	};
})(typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : this);
