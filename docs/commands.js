/* Function command for the Outlook online-meeting surface.
 * Inserts the user's saved Interactio link into the meeting body.
 * If no valid link is saved, opens the settings dialog first. */
/* global Office, InteractioLink */
(function () {
  'use strict';
  var L = InteractioLink;
  var NOTICE_KEY = 'interactioNotice';

  function notify(type, message) {
    try {
      var item = Office.context.mailbox.item;
      item.notificationMessages.replaceAsync(NOTICE_KEY, {
        type: type === 'error'
          ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
          : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: String(message).slice(0, 150),
        icon: type === 'error' ? undefined : 'Icon.16x16',
        persistent: false
      }, function () {});
    } catch (e) { /* notifications are best-effort */ }
  }

  function savedLink() {
    var raw = Office.context.roamingSettings.get(L.SETTINGS_KEY);
    if (!raw) return null;
    var r = L.normalizeInteractioLink(raw);
    return r.ok ? r.link : null; // re-validate: never trust stored values blindly
  }

  function saveLink(link, cb) {
    Office.context.roamingSettings.set(L.SETTINGS_KEY, link);
    Office.context.roamingSettings.saveAsync(function (res) {
      cb(res.status === Office.AsyncResultStatus.Succeeded);
    });
  }

  function settingsUrl() {
    var here = String((typeof window !== 'undefined' && window.location && window.location.href) || '');
    return here.replace(/commands\.html.*$/i, 'settings.html') + '?mode=dialog';
  }

  /** Opens the settings dialog; calls back with a normalized link or null (cancelled/invalid). */
  function askForLink(cb) {
    var done = false;
    function finish(dialog, link) {
      if (done) return; done = true;
      try { dialog && dialog.close(); } catch (e) { /* ignore */ }
      cb(link);
    }
    Office.context.ui.displayDialogAsync(settingsUrl(), { height: 50, width: 35, displayInIframe: true }, function (res) {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return finish(null, null);
      var dialog = res.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        var msg = {};
        try { msg = JSON.parse(arg.message); } catch (e) { /* ignore */ }
        var r = L.normalizeInteractioLink(msg && msg.link);
        finish(dialog, r.ok ? r.link : null);
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function () { finish(dialog, null); });
    });
  }

  function insertLink(event, link) {
    var body = Office.context.mailbox.item.body;
    body.getAsync(Office.CoercionType.Html, function (res) {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        notify('error', 'Couldn’t read the meeting details. Please try again.');
        return event.completed({ allowEvent: false });
      }
      var merged = L.mergeMeetingBlock(res.value, link);
      if (merged === null) return event.completed({ allowEvent: true }); // already there
      body.setAsync(merged, { coercionType: Office.CoercionType.Html }, function (r2) {
        if (r2.status !== Office.AsyncResultStatus.Succeeded) {
          notify('error', 'Couldn’t add the Interactio link. Please try again.');
          return event.completed({ allowEvent: false });
        }
        event.completed({ allowEvent: true });
      });
    });
  }

  function insertInteractioMeeting(event) {
    var link = savedLink();
    if (link) return insertLink(event, link);
    askForLink(function (newLink) {
      if (!newLink) {
        notify('info', 'Set up your Interactio link to add it to meetings.');
        return event.completed({ allowEvent: false });
      }
      saveLink(newLink, function (saved) {
        if (!saved) notify('error', 'Your link was added but couldn’t be saved for next time.');
        insertLink(event, newLink);
      });
    });
  }

  Office.onReady(function () {});
  Office.actions.associate('insertInteractioMeeting', insertInteractioMeeting);
})();
