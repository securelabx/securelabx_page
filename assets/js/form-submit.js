// Submits a form to the Google Apps Script Web App via a hidden iframe.
// This avoids CORS issues that plain fetch() has with Apps Script.
(function () {
  function submitViaHiddenIframe(form, onDone) {
    const iframeName = "slx-submit-frame-" + Date.now();
    const iframe = document.createElement("iframe");
    iframe.name = iframeName;
    iframe.style.display = "none";
    document.body.appendChild(iframe);

    form.target = iframeName;
    form.action = SITE_CONFIG.APPS_SCRIPT_URL;
    form.method = "POST";

    const cleanup = () => {
      iframe.removeEventListener("load", cleanup);
      onDone();
      setTimeout(() => iframe.remove(), 2000);
    };
    iframe.addEventListener("load", cleanup);

    form.submit();
  }

  document.addEventListener("DOMContentLoaded", () => {
    const form = document.getElementById("slx-data-form");
    if (!form) return;

    const status = document.getElementById("slx-form-status");
    const submitBtn = form.querySelector('button[type="submit"]');

    form.addEventListener("submit", (e) => {
      e.preventDefault();

      if (!SITE_CONFIG.APPS_SCRIPT_URL || SITE_CONFIG.APPS_SCRIPT_URL.includes("PASTE_YOUR")) {
        status.textContent =
          "Form backend not configured yet. Set APPS_SCRIPT_URL in assets/js/config.js.";
        status.className = "form-status error";
        status.style.display = "block";
        return;
      }

      submitBtn.disabled = true;
      submitBtn.textContent = "Submitting...";
      status.style.display = "none";

      submitViaHiddenIframe(form, () => {
        submitBtn.disabled = false;
        submitBtn.textContent = "Submit";
        status.textContent = "Submitted successfully. The team has been notified.";
        status.className = "form-status success";
        status.style.display = "block";
        form.reset();
      });
    });
  });
})();
