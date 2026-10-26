// Microsoft's window only opens from a click: the browser blocks it otherwise.
document.getElementById('continue').addEventListener(
  'click',
  () => window.msalPopupRelay.runPopupRelay({ allowedAuthorityOrigins: ['https://login.microsoftonline.com'] }),
  { once: true }
);
