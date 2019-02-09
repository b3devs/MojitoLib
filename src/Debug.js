'use strict';
/*
 * Copyright (c) 2013-2019 b3devs@gmail.com
 * MIT License: https://spdx.org/licenses/MIT.html
 */


export const Debug = {
  enabled: false,
  traceEnabled: false,
  CACHE_DEBUG_LOG: 'mojito.debug.log',
  CACHE_DEBUG_LOG_EXPIRATION: 21600, // Set max allowed expiration (6 hours)

  txnDeleteEnabled: false, // Didn't know where else to put this

  assert: function(assertion, msgIfError) {
    if (!assertion) {
      if (msgIfError == undefined || msgIfError == null)
        msgIfError = 'assert() failed';

      Browser.msgBox('Assertion Error', msgIfError, Browser.Buttons.OK);
      throw 'Assertion error: ' + msgIfError;
    }
  },

  log: function(msg) {
    if (!Debug.enabled)
      return;

    if (arguments.length > 1) {
      msg = Utilities.formatString(arguments[0], arguments[1], arguments[2], arguments[3], arguments[4], arguments[5], arguments[6], arguments[7], arguments[8], arguments[9]);
    }

    this.writeLog(msg);
  },

  trace: function(msg) {
    if (!Debug.traceEnabled)
      return;

    if (arguments.length > 1) {
      msg = Utilities.formatString(arguments[0], arguments[1], arguments[2], arguments[3], arguments[4], arguments[5], arguments[6], arguments[7], arguments[8], arguments[9]);
    }

    this.writeLog(msg);
  },

  /**
   * Debug log function only intended for temporary debugging.
   * Same as log() bug with different name so we can quickly search for it and delete it.
   */
  tempLog: function(msg) {
    if (!Debug.enabled)
      return;

    if (arguments.length > 1) {
      msg = Utilities.formatString(arguments[0], arguments[1], arguments[2], arguments[3], arguments[4], arguments[5], arguments[6], arguments[7], arguments[8], arguments[9]);
    }

    this.writeLog(msg);
  },

  writeLog: function(msg) {
    try {
      if (SpreadsheetApp.getActiveSpreadsheet()) {
        var time = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd HH:mm:ss.zzz');
        var caller = '';//arguments.callee.caller.name;
        var logEntry = `[${time}]  ${msg}\n`;

        const cache = Utils.getPrivateCache();
        let debugLog = cache.get(this.CACHE_DEBUG_LOG);
        debugLog = (!debugLog ? logEntry : debugLog + logEntry);
        try {
          cache.put(this.CACHE_DEBUG_LOG, debugLog, this.CACHE_DEBUG_LOG_EXPIRATION);
        }
        catch (e) {
          // Log probably exceeded cache limit. Cut the log in half and try again.
          var purgeMsg = '--- Debug log exceeded max size. First half was purged. ---';
          var len = debugLog.length;
          debugLog = `${purgeMsg}\n\n${debugLog.substring(Math.floor(len / 2))}`;
          cache.put(this.CACHE_DEBUG_LOG, debugLog, this.CACHE_DEBUG_LOG_EXPIRATION);
        }
      }
      else {
        Logger.log(msg);
      }
    }
    catch (e) {
      Browser.msgBox('Error: ' + e.toString());
      // Catch any exceptions. Don't let a 'debug' error interrupt execution...
    }
  },

  getLog: function() {
    var cache = Utils.getPrivateCache();
    const debugLog = cache.get(this.CACHE_DEBUG_LOG);
    return debugLog;
  },

  resetLog: function() {
    var cache = Utils.getPrivateCache();
    cache.remove(this.CACHE_DEBUG_LOG);
  },

  getExceptionInfo: function(e) {
    var exInfo = 'Exception:\n';
    for (var prop in e) {
      exInfo += `    ${prop}:  '${e[prop]}'\n`;
    }
    exInfo += '    e.toString():  ' + e.toString();
    return exInfo;
  },

  displayLogWindow: function() {
    try {
      const htmlOutput = HtmlService.createTemplateFromFile('debug_log.html').evaluate();
      htmlOutput.setTitle("Debug Log Output").setHeight(400).setWidth(800).setSandboxMode(HtmlService.SandboxMode.IFRAME);
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) ss.show(htmlOutput);
    }
    catch (e) {
      Debug.log(Debug.getExceptionInfo(e));
      Browser.msgBox(e);
    }
  }
}
