/* Settings page. Two modes:
 *  - task pane (default): reads/writes Office.context.roamingSettings directly.
 *  - dialog (?mode=dialog): opened by the function command when no link is saved;
 *    sends the link back with messageParent (dialogs can't use roamingSettings). */
/* global Office, InteractioLink */
(function () {
  'use strict';
  var L = InteractioLink;
  var isDialog = /[?&]mode=dialog\b/.test(location.search);
  var $ = function (id) { return document.getElementById(id); };
  var input = $('link'), msg = $('msg');

  function say(text, kind) { msg.textContent = text || ''; msg.className = 'msg' + (kind ? ' ' + kind : ''); }

  function showSaved(link) {
    $('current').hidden = !link;
    $('open').hidden = !link;
    $('remove').hidden = !link;
    if (link) {
      $('currentLink').textContent = link; $('currentLink').href = link; $('open').href = link;
    }
  }

  function settings() { return Office.context && Office.context.roamingSettings; }

  function init() {
    $('example').textContent = L.LINK_EXAMPLE;
    input.placeholder = L.LINK_EXAMPLE;

    if (isDialog) {
      $('subtitle').textContent = 'Set up your Interactio link to add it to this meeting';
      $('save').textContent = 'Save and add to meeting';
      $('cancel').hidden = false;
      $('cancel').onclick = function () { Office.context.ui.messageParent(JSON.stringify({ cancel: true })); };
    } else {
      var rs = settings();
      var saved = rs && rs.get(L.SETTINGS_KEY);
      var r = saved ? L.normalizeInteractioLink(saved) : { ok: false };
      if (r.ok) { input.value = r.link; showSaved(r.link); }
      $('remove').onclick = function () {
        rs.remove(L.SETTINGS_KEY);
        rs.saveAsync(function (res) {
          if (res.status === Office.AsyncResultStatus.Succeeded) { input.value = ''; showSaved(null); say('Link removed.', 'ok'); }
          else say('Couldn’t remove the link. Please try again.', 'err');
        });
      };
    }

    $('form').addEventListener('submit', function (e) {
      e.preventDefault();
      var r = L.normalizeInteractioLink(input.value);
      if (!r.ok) { say(r.error, 'err'); input.focus(); return; }
      input.value = r.link;
      if (isDialog) {
        Office.context.ui.messageParent(JSON.stringify({ link: r.link }));
        return;
      }
      var rs = settings();
      $('save').disabled = true;
      rs.set(L.SETTINGS_KEY, r.link);
      rs.saveAsync(function (res) {
        $('save').disabled = false;
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          showSaved(r.link);
          say('Saved. Choose Interactio when you create a meeting to add this link.', 'ok');
        } else say('Couldn’t save the link. Please try again.', 'err');
      });
    });
    input.addEventListener('input', function () { say(''); });
  }

  Office.onReady(init);
})();
