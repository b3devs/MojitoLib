'use strict';
/*
 * Copyright (c) 2013-2018 b3devs@gmail.com
 * MIT License: https://spdx.org/licenses/MIT.html
 */

import {Const} from './Constants.js';
import {Debug} from './Debug.js';

let MojitoScript = null;

/**
 * Displays popup toast message in the bottom right of the window.
 * 
 * @param msg - Message to display
 * @param title - Optional
 * @param timeoutSec - Optional, default is 5 sec
 */
export function toast(msg, title, timeoutSec)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss)
    return;

  if (title === undefined)
    ss.toast(msg);
  else if (timeoutSec === undefined)
    ss.toast(msg, title);
  else
    ss.toast(msg, title, timeoutSec);
}

///////////////////////////////////////////////////////////////////////////////
// Settings class

const SettingsImpl = {
  getSetting : function(settingIndex) {
    return this.getSettingValue('SettingsRange', settingIndex);
  },
  
  setSetting : function(settingIndex, value) {
    this.setSettingValue('SettingsRange', settingIndex, value);
  },
  
  getInternalSetting : function(settingIndex) {
    return this.getSettingValue('InternalSettingsRange', settingIndex);
  },

  setInternalSetting : function(settingIndex, value) {
    this.setSettingValue('InternalSettingsRange', settingIndex, value);
  },
  
  // Implementations
  getSettingValue : function(rangeName, settingIndex) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsRange = ss.getRangeByName(rangeName);
    var settingCell = settingsRange.getCell(settingIndex, 2);
    return (!settingCell ? null : settingCell.getValue());
  },
  setSettingValue : function(rangeName, settingIndex, value) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsRange = ss.getRangeByName(rangeName);
    var settingCell = settingsRange.getCell(settingIndex, 2);
    if (!settingCell) { throw 'Invalid setting index (' + rangeName + ', ' + settingIndex + ')'; }
    settingCell.setValue(value);
  },
};

export const Settings = SettingsImpl;

///////////////////////////////////////////////////////////////////////////////
// Utils class

export const Utils = {
  isDemoMode : false,

  checkDemoMode : function() {
    if (this.isDemoMode === true) {
      Browser.msgBox('Feature disabled', 'This feature is disabled so the demo data isn\'t overwritten. Feel free to download the non-demo version of Mojito and try things out as you like.', Browser.Buttons.OK);
    }
    return this.isDemoMode;
  },
  
  getPrivateCache : function() {
    if (MojitoScript === null) {
      return CacheService.getPrivateCache();
    }

    return MojitoScript.getPrivateCache();
  },

  getDocumentLock : function() {
    if (MojitoScript === null) {
      if (Debug.traceEnabled) Debug.trace('MojitoScript not set, returning ScriptProperties instead of DocumentProperties');
      return LockService.getScriptLock();
    }

    var docLock = MojitoScript.getDocumentLock();
    return docLock;
  },

  getDocumentProperties : function() {
    if (MojitoScript === null) {
      if (Debug.traceEnabled) Debug.trace('MojitoScript not set, returning ScriptProperties instead of DocumentProperties');
      return PropertiesService.getScriptProperties();
    }

    var docProps = MojitoScript.getDocumentProperties();
    return docProps;
  },

  getMintLoginAccount : function() {
    var mintAcct = Utils.getPrivateCache().get(Const.CACHE_LOGIN_ACCOUNT);
    if (mintAcct === null) {
      // Try getting mint account from settings sheet
      mintAcct = Settings.getSetting(Const.IDX_SETTING_MINT_LOGIN);
      if (mintAcct === '') {
        mintAcct = null;
      }
    }
    return mintAcct;
  },

  getSavedPassword : function() {
    return Utils.getAccountStats();
  },

  getAccountStats : function() {
    var encodedValue = Db.DataStore.getRecord(Const.DSKEY_ACCT_STATS);
    if (!encodedValue) {
      return null;
    }

    var decodedBytes = Utilities.base64Decode(encodedValue);
    var value = Utilities.newBlob(decodedBytes).getDataAsString();
    value = value.substr(Const.ACCT_STATS_EXTRA1.length);
    value = value.substr(0, value.length - Const.ACCT_STATS_EXTRA2.length);
    return value;
  },
  
  saveAccountStats : function(stats) {
    if (!stats) {
      Db.DataStore.removeRecord(Const.DSKEY_ACCT_STATS);
    }
    else {
      var value = Utilities.base64Encode(Const.ACCT_STATS_EXTRA1 + stats + Const.ACCT_STATS_EXTRA2);
      Db.DataStore.saveRecord(Const.DSKEY_ACCT_STATS, value);
    }
  },
  
  clearCacheEntries : function() {
    var cache = this.getPrivateCache();
    cache.remove(Const.CACHE_ACCOUNT_INFO_MAP);
    cache.remove(Const.CACHE_CATEGORY_MAP);
    cache.remove(Const.CACHE_TAG_MAP);
    cache.remove(Const.CACHE_SETTING_CLEARED_TAG);
    cache.remove(Const.CACHE_SETTING_RECONCILED_TAG);
    cache.remove(Const.CACHE_TXNDATA_AMOUNT_COL);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      var sheets = ss.getSheets();
      var sheetCount = ss.getNumSheets();
      for (var i = 0; i < sheetCount; ++i) {
        var sheet = sheets[i];
        var typeCell = sheet.getRange(Const.SHEET_TYPE_CELL);

        if (typeCell !== null) {
          switch (typeCell.getValue())
          {
          case Const.SHEET_TYPE_BUDGET:
            Sheets.Budget.clearNamedRanges(sheet);
            break;

          case Const.SHEET_TYPE_INOUT:
            Sheets.InOut.clearNamedRanges(sheet);
            break;

          case Const.SHEET_TYPE_SAVINGS_GOALS:
            Sheets.SavingsGoal.clearNamedRanges(sheet);
            break;
          }
        }
      }
    }
  },

  getTimezoneOffset : function()
  {
    try
    {
      var now = new Date();
      // The Date.getTimezoneOffset() method returns minutes, and positive values indicate
      // times *behind* UTC, which is the opposite of what we want (or expect!?)
      var tzOffsetInHours = -(now.getTimezoneOffset() / 60);
      Debug.log('timezone offset: %d', tzOffsetInHours);
      return tzOffsetInHours;
    }
    catch (e)
    { Debug.log(Debug.getExceptionInfo(e)); }
  },

  getHumanFriendlyElapsedTime : function(elapsedSeconds) {
    var elapsedTimeStr = '';

    var sec = elapsedSeconds;
    var min = 0;
    var hr = 0;
    if (sec >= 60 ) {
      min = Math.floor(sec / 60);
      sec = sec % 60;
    }
    if (min >= 60) {
      hr = Math.floor(min / 60);
      min = min % 60;
    }

    if (hr > 0) {
      elapsedTimeStr = String(hr) + ' hr';
      if (min > 0) {
        elapsedTimeStr += ' ' + String(min) + ' min';
      }
    }
    else if (min > 0) {
      elapsedTimeStr = String(min) + ' min';
      if (sec > 0) {
        elapsedTimeStr += ' ' + String(sec) + ' sec';
      }
    }
    else {
      elapsedTimeStr = String(sec) + ' seconds';
    }

    return elapsedTimeStr;
  },
  
  /**
   * Supported values for 'units'
   *   Y     years
   *   M     months
   *   W     weeks
   *   D     days
   *   *     human readable
   * @param startDate
   * @param endDate
   * @param units
   * @returns Object, { diff : 0, unit : '' };
   */
  getHumanFriendlyDateDiff : function(startDate, endDate, units) {
    var dateDiff = { diff : 0, unit : '' };

    var daysLeft = Math.round((endDate - startDate)/Const.ONE_DAY_IN_MILLIS);

    if (!units || units === 'D') {
      dateDiff.diff = daysLeft;
      dateDiff.unit = 'days';

    } else if (units === 'W') {
      dateDiff.diff = Math.round(daysLeft / 7);
      dateDiff.unit = 'weeks';

    } else if (units === 'M') {
      dateDiff.diff = Math.round(daysLeft / 30 * 10) / 10; // One decimal point
      dateDiff.unit = 'months';

    } else if (units === 'Y') {
      dateDiff.diff = Math.round(daysLeft / 365.25 * 10) / 10; // One decimal point
      dateDiff.unit = 'years';

    } else if (units === '*') {

      if (daysLeft > 365 + 90) { // After 15 months, use 'years'
        dateDiff.diff = Math.round(daysLeft / 365.25 * 10) / 10; // One decimal point
        dateDiff.unit = 'years';

      } else if (daysLeft > 56) { // After 8 weeks, use 'months'
        dateDiff.diff = Math.round(daysLeft / 30 * 10) / 10; // One decimal point
        dateDiff.unit = 'months';

      } else if (daysLeft > 13) { // After 13 days, use 'weeks'
        dateDiff.diff = Math.round(daysLeft / 7);
        dateDiff.unit = 'weeks';
      }
      else {
        dateDiff.diff = daysLeft;
        dateDiff.unit = 'days';
      }
    }
    
    return dateDiff;
  },
  
  getStartEndDates : function(dateText) {
    var today = new Date();
    var startDate = null;
    var endDate = null;
    
    switch (dateText.toLowerCase())
    {
      case Const.IDX_DATERANGE_THIS_MONTH:
        startDate = new Date(today.getFullYear(), today.getMonth(), 1);
        endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
        break;
        
      case Const.IDX_DATERANGE_LAST_MONTH:
        startDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
        endDate = new Date(today.getFullYear(), today.getMonth(), 0);
        break;
        
      case Const.IDX_DATERANGE_LAST_3_MONTHS:
        startDate = new Date(today.getFullYear(), today.getMonth() - 3, 1);
        endDate = new Date(today.getFullYear(), today.getMonth(), 0);
        break;
        
      case Const.IDX_DATERANGE_LAST_6_MONTHS:
        startDate = new Date(today.getFullYear(), today.getMonth() - 6, 1);
        endDate = new Date(today.getFullYear(), today.getMonth(), 0);
        break;
        
      case Const.IDX_DATERANGE_YEAR_TO_DATE:
        startDate = new Date(today.getFullYear(), 0, 1);
        endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
        break;

      case Const.IDX_DATERANGE_THIS_QUARTER:
        var monthOfQuarter = today.getMonth() % 3;
        startDate = new Date(today.getFullYear(), today.getMonth() - monthOfQuarter, 1);
        endDate = new Date(startDate.getFullYear(), startDate.getMonth() + 3, 0);
        break;

      case Const.IDX_DATERANGE_LAST_QUARTER:
        var monthOfQuarter = today.getMonth() % 3;
        startDate = new Date(today.getFullYear(), today.getMonth() - monthOfQuarter - 3, 1);
        endDate = new Date(startDate.getFullYear(), startDate.getMonth() + 3, 0);
        break;

      case Const.IDX_DATERANGE_THIS_WEEK:
        startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - today.getDay());
        endDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + 6);
        break;

      case Const.IDX_DATERANGE_LAST_WEEK:
        startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - today.getDay() - 7);
        endDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + 6);
        break;

      default:
        // Custom
    }
    
    var result = {startDate : startDate, endDate : endDate};
    return result;
  },

  getDataRange : function(sheetName, lastCol, returnNullIfEmpty)
  {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    var firstRow = sheet.getFrozenRows() + 1;
    var numRows = Math.max(0, sheet.getLastRow() - firstRow + 1);
    if (numRows === 0) {
      if (returnNullIfEmpty === true) {
        return null;
      } else {
        numRows = 1; // Should have at least one row
      }
    }

    var range = sheet.getRange(firstRow, 1, numRows, lastCol);
    if (Debug.enabled) Debug.log(Utilities.formatString('getDataRange(%s), (%d,%d) - (%d, %d)', sheetName, firstRow, 1, numRows, lastCol));

    return range;
  },
  
  getTxnDataRange : function(returnNullIfEmpty)
  {
    var lastCol = Math.max(Const.IDX_TXN_LAST_COL + 1, Utils.getTxnAmountColumn());
    return this.getDataRange(Const.SHEET_NAME_TXNDATA, lastCol, returnNullIfEmpty);
  },
  
  getAccountDataRanges : function(returnNullRangesIfEmpty)
  {
    var acctRanges = {
      hdrRange : null,
      dateRange : null,
      balanceRange : null,
      isEmpty : true,
    };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(Const.SHEET_NAME_ACCTDATA);
    var firstHeaderRow = 2;
    var firstRow = sheet.getFrozenRows() + 1;
    var numRows = Math.max(0, sheet.getLastRow() - firstRow + 1);
    var firstCol = sheet.getFrozenColumns() + 1;
    var numCols = Math.max(0, sheet.getLastColumn() - firstCol + 1);
    var isEmpty = (numCols === 0 || numRows === 0);

    acctRanges.isEmpty = isEmpty;

    if (numCols === 0 && returnNullRangesIfEmpty === true) {
      acctRanges.hdrRange = null;
    } else {
      numCols = Math.max(1, numCols); // Should have at least one column
      acctRanges.hdrRange = sheet.getRange(firstHeaderRow, firstCol, firstRow - 1, numCols);
    }

    if (isEmpty === true && returnNullRangesIfEmpty === true) {
      acctRanges.balanceRange = null;
      acctRanges.dateRange = null;
    } else {
      numRows = Math.max(1, numRows); // Should have at least one row
      // Use the date column to determine the actual number of rows. User could have columns
      // with computed values that extend beyond the actual account data.
      acctRanges.dateRange = sheet.getRange(firstRow, 1, numRows, 1);
      var dateValues = acctRanges.dateRange.getValues();
      // Find the row of the last date (search from the end)
      for (var i = numRows - 1; i >= 0; --i) {
        if (dateValues[i][0] instanceof Date) {
          break;
        }
      }

      numRows = Math.min(Math.max(i + 1, 1), numRows);

      acctRanges.dateRange = sheet.getRange(firstRow, 1, numRows, 1);
      acctRanges.balanceRange = sheet.getRange(firstRow, firstCol, numRows, numCols);
    }

    // Sanity check. Make sure balanceRange is null if hdrRange is null
    if (acctRanges.hdrRange === null && acctRanges.balanceRange !== null) {
      throw 'AccountData sheet is in a bad state. There are account balances, but no accounts are listed at the top. Please delete all of the account balances and try again.';
    }

    return acctRanges;
  },

  getCategoryDataRange : function(returnNullIfEmpty)
  {
    return this.getDataRange(Const.SHEET_NAME_CATEGORYDATA, Const.IDX_CAT_LAST_COL + 1, returnNullIfEmpty);
  },
  
  getTagDataRange : function(returnNullIfEmpty)
  {
    return this.getDataRange(Const.SHEET_NAME_TAGDATA, Const.IDX_TAG_LAST_COL + 1, returnNullIfEmpty);
  },

  getTxnAmountColumn : function() {
    var cache = Utils.getPrivateCache();
    var amountCol = Number(cache.get(Const.CACHE_TXNDATA_AMOUNT_COL));
    if (amountCol > 0) {
      if (Debug.enabled) Debug.log('Found "txn amount column" in cache: ' + amountCol);
      return amountCol;
    }

    amountCol = Const.IDX_TXN_AMOUNT + 1;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(Const.SHEET_NAME_TXNDATA);
    var amountColSetting = Settings.getSetting(Const.IDX_SETTING_TXN_AMOUNT_COL);
    if (amountColSetting) {
      try
      {
        var temp = sheet.getRange(amountColSetting).getColumn();
        amountCol = temp;
        if (Debug.enabled) Debug.log('Using txn amount column ' + amountCol);
      }
      catch (e)
      {
        var errMsg = 'The transaction \'amount\' column specified on the Settings sheet is invalid: ' + amountColSetting;
        Browser.msgBox(errMsg);
        Debug.log(errMsg);
      }
    }

    // Save the amount column in the cache so we don't have to keep looking it up.
    cache.put(Const.CACHE_TXNDATA_AMOUNT_COL, String(amountCol));

    return amountCol;
  },

  find2dArrayMaxValue : function (values, colIndex, rowMatchFunc)
  {
    if (rowMatchFunc === undefined)
      rowMatchFunc = null;

    var maxVal = null;
    var len = values.length;
    for (var i = 0; i < len; ++i) {
      var nextVal = values[i][colIndex];

      if (rowMatchFunc) {
        if (!rowMatchFunc(values[i])) {
          continue;
        }
      }

      if (maxVal === null) {
        maxVal = nextVal;
      } else if (nextVal > maxVal) {
        maxVal = nextVal;
      }
    }
    return maxVal;
  },
  
  find2dArrayMinValue : function(values, colIndex)
  {
    var minVal = null;
    var len = values.length;
    for (var i = 0; i < len; ++i) {
      var nextVal = values[i][colIndex];
      
      if (minVal === null) {
        minVal = nextVal;
      } else if (nextVal < minVal) {
        minVal = nextVal;
      }
    }
    return minVal;
  },

  sort2dArray : function(values, colIndexArray, sortOrderArray) {
    // Values for sortOrderArray are either 1 for ascending, or -1 for descending
    var colIndexCount = colIndexArray.length;

    values.sort(function(a, b) {
      try {
        for (var i = 0; i < colIndexCount; ++i) {
          if (a[i] > b[i]) { return sortOrderArray[i]; }
          else if (a[i] < b[i]) { return -sortOrderArray[i]; }
        };
      } catch (e) {
        Debug.log(Debug.getExceptionInfo(e));
        Debug.log('a: ' + JSON.stringify(a));
        Debug.log('b: ' + JSON.stringify(b));
        throw e;
      }
      return 0;
    });
    
  },

  convertDelimitedStringToArray : function(stringList, delimiter)
  {
    var array = stringList.split(delimiter);
    var map = [];
    for (var i = 0; i < array.length; ++i) {
      var item = array[i].trim();
      if (item === '')
        continue;
      
      map[item] = i;
    }
    return map;
  },

  parseDateString : function(dateString, today) {
    if (today === undefined || today ===  null) {
      today = new Date();
    }

    var parsedDate = null;

    if (dateString.indexOf('/') > 0) {
      // dateString is formatted as month/day/year
      const dateParts = dateString.split('/');

      Debug.assert(dateParts.length >= 3, 'parseDateString: Invalid date, less than 3 date parts');

      let year = parseInt(dateParts[2], 10);
      const month = parseInt(dateParts[0], 10) - 1;
      const day = parseInt(dateParts[1], 10);

      if (year < 100) {
        var century = Math.round(today.getFullYear() / 100) * 100;
        year = century + year;
      }

      parsedDate = new Date(year, month, day);

    } else if (dateString.indexOf(' ') > 0) {
      // dateString is formatted as 'Month Day'
      const dateParts = dateString.split(' ');
      let year = today.getFullYear();
      const month = MONTH_LOOKUP_1[dateParts[0]];
      const day = parseInt(dateParts[1]);

      if (today.getMonth() < month)
        --year;

      Debug.assert(day > 0, 'parseDateString: day is zero');

      parsedDate = new Date(year, month, day);
    }

    return parsedDate;
  },

  invokeFunction : function(rootObj, funcName, args) {
    let ret = null;

    try
    {
      if (Debug.traceEnabled) Debug.trace('invokeFunction(\'%s\', %s)', funcName, (args ? 'args' : 'null'));
      
      // For the function name, parse through nested objects, if any:  MyLib.SomeObj.NestedObj.callbackFunc
      var callbackObjects = funcName.split('.');
      var func = callbackObjects[0];
      var obj = rootObj;
      for (var i = 1; i < callbackObjects.length; ++i) {
        var nestedObj = func;
        obj = obj[nestedObj];
        func = callbackObjects[i];
      }
      
      // Call the function
      if (args instanceof Array) {
        ret = obj[func].apply(obj, args);
      }
      else {
        ret = obj[func](args);
      }
    }
    catch(e)
    {
      Debug.log(Debug.getExceptionInfo(e));
    }

    return ret;
  },
};

///////////////////////////////////////////////////////////////////////////////
// CommandQueue class

export const CommandQueue = {
  pushCommand(cmdFunctionName, args) {

  }
};

///////////////////////////////////////////////////////////////////////////////
// EventServiceX class

export const EventServiceX = {
  registerForEvent(eventName, callbackFunctionName, contextData) {
    var cacheValue = callbackFunctionName + (contextData ? ',' + contextData : '');
    var cache = Utils.getPrivateCache();
    cache.put('mojito.registered_' + eventName, cacheValue);
  },

  triggerEvent (eventName, eventData) {
    var cache = Utils.getPrivateCache();
    cache.put(eventName, (new Date()).toString(), 10); // trigger the event for 10 seconds
    var callbackInfo = cache.get('mojito.registered_' + eventName);

    if (callbackInfo) {
      this.invokeCallback(callbackInfo, eventData);
    }
  },

  // Returns true if event was triggered, or false if timed out
  waitForEvent (eventName, timeoutSeconds) {
    return (!!this.waitForEvents([eventName], timeoutSeconds));
  },

  // Returns name of event that was triggered, or null if timed out
  waitForEvents (eventNames, timeoutSeconds) {
    var cache = Utils.getPrivateCache();

    while(true) {
      for (var i = 0; i < eventNames.length; ++i) {
        if (cache.get(eventNames[i])) {
          cache.remove(eventNames[i]);
          return eventNames[i];
        }
      }

      Utilities.sleep(1000);

      if (timeoutSeconds > 0) {
        if (--timeoutSeconds <= 0) {
          break;
        }
      }
    }
    
    return null;
  },

  clearEvent (eventName) {
    Utils.getPrivateCache().remove(eventName);
  },

  waitForUiClose (closeEvent, stillOpenEvent, timeoutSeconds) {
    var cache = Utils.getPrivateCache();

    while(true) {
      if (null !== cache.get(closeEvent)) {
        cache.remove(closeEvent);
        toast('UI closed');
        return true;

      } else if (null === cache.get(stillOpenEvent)) {
        toast('UI canceled');
        return true;
      }

      Utilities.sleep(1000);

      if (timeoutSeconds > 0) {
        if (--timeoutSeconds <= 0) {
          break;
        }
      }
    }

    toast('UI timeout');
    return false;
  },

  invokeCallback(callbackInfo, data) {
    if (Debug.enabled) Debug.log('triggerEvent: Calling callback: ' + callbackInfo);

    // Parse out the callback function name and the context data
    var contextData = null;
    var callbackFunctionName = callbackInfo;
    
    var contextStart = callbackInfo.indexOf(',');
    if (contextStart > 0) {
      contextData = callbackInfo.substr(contextStart + 1);
      callbackFunctionName = callbackInfo.substr(0, contextStart);
    }
    
    // For the function name, parse through nested objects, if any:  MyLib.SomeObj.NestedObj.callbackFunc
    var callbackObjects = callbackFunctionName.split('.');
    var func = callbackObjects[0];
    var obj = MojitoLib;
    for (var i = 1; i < callbackObjects.length; ++i) {
      var nestedObj = func;
      obj = obj[nestedObj];
      func = callbackObjects[i];
    }
    
    // Call the function
    obj[func](contextData, data);
  },
};
