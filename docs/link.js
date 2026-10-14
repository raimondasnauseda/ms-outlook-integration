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
   * HTML block added to the meeting body. `link` must already be normalized.
   * Renders as two lines, in rich and plain-text views alike:
   *   🟣 Interactio meeting
   *   Join: https://…
   * The link text IS the URL, so plain-text renderings (e.g. Google Calendar) show it once.
   */
  function buildMeetingBlock(link) {
    var safe = escapeHtml(link);
    return (
      '<div id="' + BLOCK_ID + '">' +
      '<div>&#x1F7E3; <b>Interactio meeting</b></div>' +
      '<div>Join: <a href="' + safe + '">' + safe + '</a></div>' +
      '</div>'
    );
  }

  /** True if the HTML has no visible content (Outlook's "empty" body is still a full document). */
  function isBlankHtml(html) {
    var s = String(html || '');
    if (/<(img|table|hr|iframe|object|video)\b/i.test(s)) return false;
    s = s.replace(/<head[\s\S]*?<\/head>/gi, '')
         .replace(/<(style|script)[\s\S]*?<\/\1>/gi, '')
         .replace(/<!--[\s\S]*?-->/g, '')
         .replace(/<[^>]+>/g, '')
         .replace(/&nbsp;|&#160;|&#xa0;|\u00a0|\u200b/gi, '');
    return !/\S/.test(s);
  }

  /**
   * Finds an existing Interactio block (any version; Outlook may rename the id to e.g.
   * "x_interactio-meeting") and returns [start, end) of its outer <div>, or null.
   * Walks <div>/</div> tags to find the matching close, so nested markup is handled.
   */
  function findExistingBlock(html) {
    var open = /<div\b[^>]*\bid=["']?[a-z0-9_]*interactio-meeting["']?[^>]*>/i.exec(html);
    if (!open) return null;
    var tag = /<(\/?)div\b[^>]*>/gi;
    tag.lastIndex = open.index + open[0].length;
    var depth = 1, m;
    while ((m = tag.exec(html))) {
      depth += m[1] ? -1 : 1;
      if (depth === 0) return [open.index, tag.lastIndex];
    }
    return null;
  }

  /**
   * Returns the new body HTML, or null if the body already contains this exact link
   * (nothing to do). An older Interactio block (e.g. a previous link or layout) is replaced.
   * An empty body is replaced by the block, so the text starts at the very top.
   */
  function mergeMeetingBlock(bodyHtml, link) {
    var body = bodyHtml == null ? '' : String(bodyHtml);
    var safe = escapeHtml(link);
    if (body.indexOf(safe) !== -1 || body.indexOf(link) !== -1) return null;
    var block = buildMeetingBlock(link);
    var found = findExistingBlock(body);
    if (found) return body.slice(0, found[0]) + block + body.slice(found[1]);
    if (isBlankHtml(body)) return block;
    var close = body.search(/<\/body>/i);
    return close === -1 ? body + block : body.slice(0, close) + block + body.slice(close);
  }

  var api = {
    ALLOWED_HOST_SUFFIXES: ALLOWED_HOST_SUFFIXES,
    LINK_EXAMPLE: LINK_EXAMPLE,
    SETTINGS_KEY: SETTINGS_KEY,
    BLOCK_ID: BLOCK_ID,
    normalizeInteractioLink: normalizeInteractioLink,
    escapeHtml: escapeHtml,
    buildMeetingBlock: buildMeetingBlock,
    mergeMeetingBlock: mergeMeetingBlock,
    isBlankHtml: isBlankHtml
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.InteractioLink = api;
})(typeof self !== 'undefined' ? self : this);
