(() => {
  const bindActions = () => {
    const printButton = document.getElementById('cgpt-popup-print-btn');
    const closeButton = document.getElementById('cgpt-popup-close-btn');

    if (printButton && !printButton.hasAttribute('data-cgpt-export-bound')) {
      printButton.setAttribute('data-cgpt-export-bound', '1');
      printButton.addEventListener('click', () => {
        try {
          window.focus();
          setTimeout(() => window.print(), 60);
        } catch (_error) {
          // ignored
        }
      });
    }

    if (closeButton && !closeButton.hasAttribute('data-cgpt-export-bound')) {
      closeButton.setAttribute('data-cgpt-export-bound', '1');
      closeButton.addEventListener('click', () => {
        try {
          window.close();
        } catch (_error) {
          // ignored
        }
      });
    }

    return Boolean(printButton && closeButton);
  };

  const runAutoPrint = () => {
    if (window.__cgptExportAutoPrintStarted) {
      return;
    }
    window.__cgptExportAutoPrintStarted = true;

    const run = () => {
      try {
        window.focus();
        setTimeout(() => window.print(), 220);
      } catch (_error) {
        // ignored
      }
    };

    try {
      const fontsReady = document.fonts?.ready;
      if (fontsReady && typeof fontsReady.then === 'function') {
        fontsReady.then(() => setTimeout(run, 220)).catch(() => setTimeout(run, 220));
      } else {
        setTimeout(run, 320);
      }
    } catch (_error) {
      setTimeout(run, 320);
    }
  };

  const init = () => {
    if (!bindActions()) {
      return;
    }
    runAutoPrint();
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
  } else {
    init();
  }
})();
