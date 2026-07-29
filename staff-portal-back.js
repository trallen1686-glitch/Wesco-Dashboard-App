(function(){
  'use strict';

  var scriptUrl = document.currentScript && document.currentScript.src
    ? document.currentScript.src
    : window.location.href;
  var portalUrl = new URL('Staff-Portal-Starter-Template.html', scriptUrl);
  var person = new URL(window.location.href).searchParams.get('person');
  if(person) portalUrl.searchParams.set('person', person);

  var style = document.createElement('style');
  style.textContent =
    '.wesco-portal-back{' +
      'position:fixed;left:18px;bottom:18px;z-index:2147483647;' +
      'display:inline-flex;align-items:center;gap:7px;min-height:42px;' +
      'padding:10px 15px;border:1px solid #f36a00;border-radius:9px;' +
      'background:#f36a00;color:#fff!important;font:800 14px/1.2 Arial,Helvetica,sans-serif;' +
      'text-decoration:none!important;box-shadow:0 10px 28px rgba(0,0,0,.42);' +
    '}' +
    '.wesco-portal-back:hover,.wesco-portal-back:focus{' +
      'background:#ff7a14;outline:3px solid rgba(243,106,0,.28);outline-offset:2px;' +
    '}' +
    '@media(max-width:520px){.wesco-portal-back{left:10px;bottom:10px;min-height:40px;padding:9px 12px;font-size:13px}}' +
    '@media print{.wesco-portal-back{display:none!important}}';
  document.head.appendChild(style);

  var link = document.createElement('a');
  link.className = 'wesco-portal-back';
  link.href = portalUrl.toString();
  link.setAttribute('aria-label', 'Back to Staff Portal');
  link.textContent = '\u2190 Back to Staff Portal';
  link.addEventListener('click', function(event){
    if(window.opener && !window.opener.closed){
      event.preventDefault();
      window.opener.focus();
      window.close();
    }
  });
  document.body.appendChild(link);
})();
