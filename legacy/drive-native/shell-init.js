if ('serviceWorker' in navigator && location.protocol === 'https:') {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js').then((registration) => {
      window.__pwaRegistration = registration;
      window.dispatchEvent(new Event('pwa-registration-ready'));
    }).catch(() => {});
  });
}
