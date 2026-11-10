var SHEET_NAME   = "";
var COL_ANCHOR   = 2;
var COL_TARGET   = 3;
var COL_URL      = 5;
var COL_STATUS   = 6;
var PAINT_URL_CELL = true;
var IMAGE_LINK_COUNTS_AS_ANCHOR = true;
var STRICT_URL_MATCH = true;
var BATCH_SIZE   = 10;
var MAX_RUNTIME_MS = 4.5 * 60 * 1000;

var COLORS = {
  success: "#6ed18e",
  warn:    "#ff0000",
  error:   "#ff0000",
  bot:     "#ff0000",
  gray:    "#ff0000"
};

var HEADERS = {
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
  "Accept": "text/html,application/xhtml+xml",
  "Accept-Language": "en-US,en;q=0.9"
};

var NOT_FOUND_PATTERNS = [
  "page not found", "page does not exist", "doesn't exist", "doesn’t exist",
  "not found", "página no encontrada", "pagina no encontrada",
  "seite nicht gefunden", "page introuvable", "pagina non trovata",
  "сторінку не знайдено", "сторінка не знайдена", "страница не найдена"
];

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('SEO Checker')
      .addItem('▶️ Check new rows (empty status)', 'checkNew')
      .addItem('🔁 Check all except success', 'checkRetry')
      .addItem('♻️ Check absolutely all', 'checkAll')
      .addSeparator()
      .addItem('⏹ Stop auto-resume', 'stopAuto')
      .addToUi();
  } catch (e) {
  }
}

function checkNew()   { runCheck_('new'); }
function checkRetry() { runCheck_('retry'); }
function checkAll()   { runCheck_('all'); }

function stopAuto() {
  deleteAutoTriggers_();
  notify_("Auto-resume stopped.", false);
}

function runCheck_(mode, silent) {
  var startTime = Date.now();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    notify_("Error: sheet not found '" + SHEET_NAME + "'.", silent);
    return;
  }

  deleteAutoTriggers_();

  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var url = (data[i][COL_URL - 1] || "").toString().trim();
    var old = (data[i][COL_STATUS - 1] || "").toString().trim();
    if (!/^https?:\/\//i.test(url)) continue;
    if (mode === 'new' && old !== "") continue;
    if (mode === 'retry' && old === "success") continue;
    rows.push({
      row: i + 1,
      url: url,
      anchor: (data[i][COL_ANCHOR - 1] || "").toString().trim(),
      target: (data[i][COL_TARGET - 1] || "").toString().trim()
    });
  }

  if (mode !== 'new' && rows.length) {
    var statusCol = sheet.getRange(1, COL_STATUS, data.length, 1).getValues();
    rows.forEach(function (r) { statusCol[r.row - 1][0] = ""; });
    sheet.getRange(1, COL_STATUS, data.length, 1).setValues(statusCol);
  }

  var done = 0, stoppedEarly = false;

  for (var s = 0; s < rows.length; s += BATCH_SIZE) {
    if (Date.now() - startTime > MAX_RUNTIME_MS) { stoppedEarly = true; break; }

    var chunk = rows.slice(s, s + BATCH_SIZE);

    var responses = null;
    try {
      responses = UrlFetchApp.fetchAll(chunk.map(function (r) {
        return { url: r.url, muteHttpExceptions: true, followRedirects: false, headers: HEADERS };
      }));
    } catch (e) {
      responses = null;
    }

    for (var j = 0; j < chunk.length; j++) {
      var r = chunk[j];
      var res;
      try {
        res = checkPage_(r.url, r.anchor, r.target, responses ? responses[j] : null);
      } catch (err) {
        res = { status: "script error: " + String(err.message || err).substring(0, 60), color: COLORS.error };
      }

      sheet.getRange(r.row, COL_STATUS)
        .setNumberFormat("@")
        .setValue(res.status)
        .setBackground(null)
        .setFontColor("#000000")
        .setFontLine("none")
        .setHorizontalAlignment("left");
      if (PAINT_URL_CELL) sheet.getRange(r.row, COL_URL).setBackground(res.color);
      done++;
    }
    SpreadsheetApp.flush();
  }

  if (stoppedEarly) {
    ScriptApp.newTrigger("autoContinue_").timeBased().after(60 * 1000).create();
    notify_("Checked " + done + " out of " + rows.length + ". The rest will continue automatically in one minute.", silent);
  } else {
    notify_("Done! Rows checked: " + done, silent);
  }
}

function autoContinue_() {
  runCheck_('new', true);
}

function deleteAutoTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "autoContinue_") ScriptApp.deleteTrigger(t);
  });
}

function notify_(msg, silent) {
  if (silent) return;
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

function checkPage_(url, anchor, targetRaw, firstResp) {
  var resp = fetchWithRedirects_(url, firstResp);

  if (resp.error) return { status: "error / DNS (" + resp.error + ")", color: COLORS.error };

  var code = resp.code;
  var html = resp.html || "";

  if (code === 404 || code === 410) return { status: "404", color: COLORS.error };
  if ((code === 403 || code === 503 || code === 429) && isBotPage_(html)) return { status: "bot check", color: COLORS.bot };
  if (code === 403 || code === 401) return { status: code, color: COLORS.bot };
  if (code === 429) return { status: "429", color: COLORS.bot };
  if (code >= 500) return { status: code, color: COLORS.error };
  if (code >= 400) return { status: code, color: COLORS.error };
  if (code >= 300) return { status: "redirect loop / HTTP " + code, color: COLORS.error };

  if (resp.redirected && !sameUrl_(url, resp.finalUrl)) {
    return { status: "redirect → " + resp.finalUrl, color: COLORS.bot };
  }

  var plainText = visibleText_(html);
  if (plainText.length < 50 && html.length < 500) {
    return isSpaShell_(html, plainText)
      ? { status: "202", color: COLORS.error }
      : { status: "white page", color: COLORS.gray };
  }

  if (hasNoindex_(html)) return { status: "no index", color: COLORS.error };

  var result = analyzeLinks_(html, url, anchor, targetRaw);
  if (result.status === "success") return result;

  if (isSoft404_(html, plainText)) return { status: "404", color: COLORS.error };

  if (result.status === "no link") {
    if (rawHtmlLooksNotFound_(html)) return { status: "404", color: COLORS.error };
    if (isSpaShell_(html, plainText)) return { status: "202", color: COLORS.error };
  }
  return result;
}

function analyzeLinks_(html, pageUrl, anchor, targetRaw) {
  var target = normalizeUrl_(targetRaw);
  if (!target) return { status: "no target url", color: COLORS.error };

  var cleanHtml = html.replace(/<!--[\s\S]*?-->/g, "");
  var anchorNorm = normalizeText_(anchor);

  var aTagRegex = /<a\s[^>]*?href\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))[^>]*>([\s\S]*?)<\/a>/gi;
  var m, linkFound = false, linkWithAnchor = false, anchorWithOtherLink = false;

  while ((m = aTagRegex.exec(cleanHtml)) !== null) {
    var href = decodeEntities_(m[1] || m[2] || m[3] || "").trim();
    if (!href || /^(#|mailto:|tel:|javascript:)/i.test(href)) continue;
    if (!/^(https?:)?\/\//i.test(href)) href = resolveUrl_(pageUrl, href);
    var inner = m[4] || "";

    var hrefMatches = urlMatches_(normalizeUrl_(href), target);
    var anchorMatches = anchorTextMatches_(inner, anchorNorm);

    if (hrefMatches) {
      linkFound = true;
      if (anchorMatches) { linkWithAnchor = true; break; }
    } else if (anchorNorm && anchorMatches && !isImageOnly_(inner)) {
      anchorWithOtherLink = true;
    }
  }

  if (linkWithAnchor)      return { status: "success",          color: COLORS.success };
  if (linkFound)           return { status: "incorrect anchor", color: COLORS.warn };
  if (anchorWithOtherLink) return { status: "incorrect link",   color: COLORS.warn };
  return { status: "no link", color: COLORS.error };
}

function fetchWithRedirects_(startUrl, firstResp) {
  var url = startUrl, redirected = false;
  var options = { muteHttpExceptions: true, followRedirects: false, headers: HEADERS };

  for (var hop = 0; hop < 7; hop++) {
    var r = (hop === 0 && firstResp) ? firstResp : null, lastErr = "";
    for (var attempt = 0; attempt < 2 && !r; attempt++) {
      try { r = UrlFetchApp.fetch(url, options); }
      catch (e) { lastErr = String(e.message || e).substring(0, 80); Utilities.sleep(700); }
    }
    if (!r) return { error: lastErr };

    var code = r.getResponseCode();
    if (code >= 300 && code < 400) {
      var h = r.getHeaders();
      var loc = h["Location"] || h["location"];
      if (!loc) return { code: code, html: r.getContentText(), finalUrl: url, redirected: redirected };
      url = resolveUrl_(url, loc);
      redirected = true;
      continue;
    }
    return { code: code, html: r.getContentText(), finalUrl: url, redirected: redirected };
  }
  return { code: 310, html: "", finalUrl: url, redirected: true };
}

function visibleText_(html) {
  return html.replace(/<script[\s\S]*?<\/script>/gi, "")
             .replace(/<style[\s\S]*?<\/style>/gi, "")
             .replace(/<noscript[\s\S]*?<\/noscript>/gi, "")
             .replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
}

function isSpaShell_(html, plainText) {
  var hasRoot = /id\s*=\s*["'](root|app|__next|__nuxt|svelte)["']/i.test(html) || /<app-root/i.test(html);
  var scripts = (html.match(/<script/gi) || []).length;
  return plainText.length < 400 && (hasRoot || scripts >= 3);
}

function rawHtmlLooksNotFound_(html) {
  var low = html.toLowerCase();
  var strong = ["page not found", "page does not exist", "doesn't exist", "doesn\u2019t exist",
                "página no encontrada", "seite nicht gefunden", "page introuvable",
                "сторінку не знайдено", "страница не найдена", "\"statuscode\":404", "\"status\":404"];
  for (var i = 0; i < strong.length; i++) if (low.indexOf(strong[i]) !== -1) return true;
  return false;
}

function isSoft404_(html, plainText) {
  var title = ((html.match(/<title[^>]*>([\s\S]*?)<\/title>/i) || [])[1] || "");
  var h1 = ((html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i) || [])[1] || "").replace(/<[^>]*>/g, "");
  var probe = normalizeText_(title + " " + h1);
  if (plainText && plainText.length < 800) probe += " " + normalizeText_(plainText);
  if (/(^|\D)404(\D|$)/.test(probe)) return true;
  for (var i = 0; i < NOT_FOUND_PATTERNS.length; i++) {
    if (probe.indexOf(NOT_FOUND_PATTERNS[i]) !== -1) return true;
  }
  return false;
}

function resolveUrl_(base, loc) {
  if (/^https?:\/\//i.test(loc)) return loc;
  var origin = (base.match(/^https?:\/\/[^\/]+/i) || [""])[0];
  if (loc.indexOf("//") === 0) return base.split(":")[0] + ":" + loc;
  if (loc.charAt(0) === "/") return origin + loc;
  return base.replace(/[^\/]*$/, "") + loc;
}

function normalizeUrl_(u) {
  return (u || "").toString().toLowerCase().trim()
    .replace(/^https?:\/\//, "").replace(/^\/\//, "").replace(/^www\./, "")
    .replace(/#.*$/, "").replace(/\/+$/, "");
}

function urlMatches_(href, target) {
  if (!href || !target) return false;
  var h = href;
  if (target.indexOf("?") === -1) h = h.replace(/\?.*$/, "");
  h = h.replace(/\/+$/, "");
  if (h === target) return true;
  if (STRICT_URL_MATCH) return false;
  var c = h.charAt(target.length);
  return h.indexOf(target) === 0 && (c === "/" || c === "?");
}

function sameUrl_(a, b) { return normalizeUrl_(a) === normalizeUrl_(b); }

function decodeEntities_(s) {
  return s.replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&").replace(/&quot;/gi, '"')
          .replace(/&#39;|&apos;/gi, "'").replace(/&lt;/gi, "<").replace(/&gt;/gi, ">");
}

function normalizeText_(s) {
  return decodeEntities_((s || "").toString()).replace(/\s+/g, " ").trim().toLowerCase();
}

function isImageOnly_(inner) {
  return /<img/i.test(inner) && normalizeText_(inner.replace(/<[^>]*>/g, "")) === "";
}

function anchorTextMatches_(inner, anchorNorm) {
  if (!anchorNorm) return true;
  var text = normalizeText_(inner.replace(/<[^>]*>/g, " "));
  if (text.indexOf(anchorNorm) !== -1) return true;
  if (/<img/i.test(inner)) {
    if (IMAGE_LINK_COUNTS_AS_ANCHOR) return true;
    var alt = (inner.match(/alt\s*=\s*["']([^"']*)["']/i) || [])[1] || "";
    return normalizeText_(alt).indexOf(anchorNorm) !== -1;
  }
  return false;
}

function isBotPage_(html) {
  return /cloudflare|just a moment|challenge-running|enable javascript|ddos-guard|captcha/i.test(html);
}

function hasNoindex_(html) {
  var metas = html.match(/<meta[^>]+>/gi) || [];
  for (var i = 0; i < metas.length; i++) {
    if (/name\s*=\s*["']robots["']/i.test(metas[i]) && /noindex/i.test(metas[i])) return true;
  }
  return false;
}