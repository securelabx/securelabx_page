// Simple client-side password gate.
// NOTE: This only deters casual visitors. It is NOT real security -
// the hash and page content are downloadable by anyone who inspects
// the site's source. Do not put confidential data behind this alone.

(function () {
  const SESSION_KEY = "slx_authed";

  async function sha256Hex(text) {
    const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
    return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
  }

  function showGate() {
    document.getElementById("slx-gate").style.display = "flex";
    document.getElementById("slx-content").style.display = "none";
  }

  function showContent() {
    document.getElementById("slx-gate").style.display = "none";
    document.getElementById("slx-content").style.display = "block";
  }

  async function tryUnlock(password) {
    const hash = await sha256Hex(password);
    if (hash === SITE_CONFIG.SITE_PASSWORD_HASH) {
      sessionStorage.setItem(SESSION_KEY, "1");
      showContent();
      return true;
    }
    return false;
  }

  document.addEventListener("DOMContentLoaded", () => {
    if (sessionStorage.getItem(SESSION_KEY) === "1") {
      showContent();
    } else {
      showGate();
    }

    const form = document.getElementById("slx-gate-form");
    const input = document.getElementById("slx-gate-password");
    const error = document.getElementById("slx-gate-error");

    form.addEventListener("submit", async (e) => {
      e.preventDefault();
      const ok = await tryUnlock(input.value);
      if (!ok) {
        error.style.display = "block";
        input.value = "";
        input.focus();
      }
    });
  });
})();
