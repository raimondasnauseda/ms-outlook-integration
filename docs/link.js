/*
 * Interactio link validation + meeting-block helpers.
 * Shared by the function command, the settings page and the offline tests.
 * Validation is a 1:1 port of src/Link.gs from the Google Calendar add-on.
 */
(function (root) {
  'use strict';

  var ALLOWED_HOST_SUFFIXES = ['interactio.com', 'interactio.io'];
  var LINK_EXAMPLE = 'https://meeting.interactio.com/link/your-name';
  var SETTINGS_KEY = 'interactioLink';
  var BLOCK_ID = 'interactio-meeting';

  function normalizeInteractioLink(raw) {
    var s = raw == null ? '' : String(raw).trim();
    if (!s) return { ok: false, error: 'Please enter your Interactio meeting link.' };
    if (s.length > 2048) return { ok: false, error: 'That link is too long.' };
    if (!/^[a-z][a-z0-9+.-]*:\/\//i.test(s)) s = 'https://' + s; // allow "meeting.interactio.com/..."
    var m = s.match(/^([a-z][a-z0-9+.-]*):\/\/([^\/?#:@\s]+)(:\d{1,5})?([\/?#][^\s]*)?$/i); // '@' rejected
    if (!m) return { ok: false, error: 'That doesn’t look like a valid link.' };
    if (m[1].toLowerCase() !== 'https') return { ok: false, error: 'The link must start with https://' };
    var host = m[2].toLowerCase().replace(/\.$/, '');
    var allowed = ALLOWED_HOST_SUFFIXES.some(function (sfx) {
      return host === sfx || host.endsWith('.' + sfx);
    });
    if (!allowed) return { ok: false, error: 'Please use an Interactio link (interactio.com, interactio.io).' };
    return { ok: true, link: 'https://' + host + (m[3] || '') + (m[4] || '') };
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  /**
   * HTML block appended to the meeting body. `link` must already be normalized.
   * The link text IS the URL, so plain-text renderings of the invite (e.g. Google Calendar,
   * text-only mail clients) show the link once instead of "text<url>" plus the URL again.
   * Divider lines are kept short so they don't wrap in narrow views.
   */
  function buildMeetingBlock(link) {
    var safe = escapeHtml(link);
    var rule = '<div>' + new Array(41).join('_') + '</div>';
    return (
      '<div id="' + BLOCK_ID + '">' +
      '<br>' +
      rule +
      '<div style="font-size:14pt;font-weight:bold;">Interactio meeting</div>' +
      '<div>Join Interactio meeting: <a href="' + safe + '">' + safe + '</a></div>' +
      rule +
      '</div>'
    );
  }

  /**
   * Returns the new body HTML, or null if the body already contains this exact link
   * (nothing to do). A previous Interactio block (e.g. an older link) is replaced.
   * Outlook may rewrite ids (e.g. "x_interactio-meeting"), so the match is tolerant.
   */
  function mergeMeetingBlock(bodyHtml, link) {
    var body = bodyHtml == null ? '' : String(bodyHtml);
    var safe = escapeHtml(link);
    if (body.indexOf(safe) !== -1 || body.indexOf(link) !== -1) return null;
    var block = buildMeetingBlock(link);
    // From the block's opening tag, past the second underscore rule, plus its closing tags.
    var re = /<div[^>]*\bid=["']?[a-z_]*interactio-meeting["']?[^>]*>(?:(?!<div[^>]*interactio-meeting)[\s\S])*?_{20,}(?:(?!<div[^>]*interactio-meeting)[\s\S])*?Join Interactio meeting(?:(?!<div[^>]*interactio-meeting)[\s\S])*?_{20,}(?:\s*<\/(?:div|span|p|font|o:p)>){0,4}/i;
    if (re.test(body)) return body.replace(re, function () { return block; });
    return body + block;
  }

  var api = {
    ALLOWED_HOST_SUFFIXES: ALLOWED_HOST_SUFFIXES,
    LINK_EXAMPLE: LINK_EXAMPLE,
    SETTINGS_KEY: SETTINGS_KEY,
    BLOCK_ID: BLOCK_ID,
    normalizeInteractioLink: normalizeInteractioLink,
    escapeHtml: escapeHtml,
    buildMeetingBlock: buildMeetingBlock,
    mergeMeetingBlock: mergeMeetingBlock
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.InteractioLink = api;
})(typeof self !== 'undefined' ? self : this);
