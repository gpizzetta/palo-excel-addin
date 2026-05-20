/**
 * Polyfills pour le runtime Custom Functions Excel Desktop (worker sans DOM).
 * Charge en premier dans functions.js (voir build-bundle.sh).
 */
(function paloCfPolyfillsBootstrap() {
  var g = typeof globalThis !== "undefined" ? globalThis
    : typeof self !== "undefined" ? self
    : typeof window !== "undefined" ? window
    : {};

  if (typeof g.TextEncoder === "undefined") {
    function TextEncoderPoly() {}
    TextEncoderPoly.prototype.encode = function encodePoly(str) {
      var s = String(str);
      var encoded = unescape(encodeURIComponent(s));
      var out = new Uint8Array(encoded.length);
      var i;
      for (i = 0; i < encoded.length; i += 1) {
        out[i] = encoded.charCodeAt(i) & 0xff;
      }
      return out;
    };
    g.TextEncoder = TextEncoderPoly;
  }

  function urlSearchParamsBroken() {
    if (typeof g.URLSearchParams === "undefined") {
      return true;
    }
    try {
      var t = new g.URLSearchParams();
      t.set("palo", "1");
      return String(t.get("palo")) !== "1";
    } catch (_e) {
      return true;
    }
  }

  if (urlSearchParamsBroken()) {
    function URLSearchParamsPoly(init) {
      this._pairs = [];
      if (typeof init === "string") {
        var s = String(init).replace(/^\?/, "");
        if (s) {
          s.split("&").forEach(function (part) {
            if (!part) {
              return;
            }
            var eq = part.indexOf("=");
            var k = eq >= 0 ? part.slice(0, eq) : part;
            var v = eq >= 0 ? part.slice(eq + 1) : "";
            this._pairs.push([decodeURIComponent(k.replace(/\+/g, " ")), decodeURIComponent(v.replace(/\+/g, " "))]);
          }, this);
        }
      }
    }
    URLSearchParamsPoly.prototype.set = function set(key, value) {
      this.delete(key);
      this._pairs.push([String(key), String(value)]);
    };
    URLSearchParamsPoly.prototype.has = function has(key) {
      var i;
      var k = String(key);
      for (i = 0; i < this._pairs.length; i += 1) {
        if (this._pairs[i][0] === k) {
          return true;
        }
      }
      return false;
    };
    URLSearchParamsPoly.prototype.get = function get(key) {
      var i;
      var k = String(key);
      for (i = 0; i < this._pairs.length; i += 1) {
        if (this._pairs[i][0] === k) {
          return this._pairs[i][1];
        }
      }
      return null;
    };
    URLSearchParamsPoly.prototype.delete = function del(key) {
      var i;
      var k = String(key);
      var next = [];
      for (i = 0; i < this._pairs.length; i += 1) {
        if (this._pairs[i][0] !== k) {
          next.push(this._pairs[i]);
        }
      }
      this._pairs = next;
    };
    URLSearchParamsPoly.prototype.toString = function toString() {
      return this._pairs.map(function (pair) {
        return encodeURIComponent(pair[0]) + "=" + encodeURIComponent(pair[1]);
      }).join("&");
    };
    g.URLSearchParams = URLSearchParamsPoly;
  }
})();
