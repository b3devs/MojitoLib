'use strict';
/*
 * Copyright (c) 2013-2019 b3devs@gmail.com
 * MIT License: https://spdx.org/licenses/MIT.html
 */

import {Const} from './Constants.js';
import {Utils, Settings, toast} from './Utils.js';
import {Sheets} from './Sheets.js';
import {Debug} from './Debug.js';

///////////////////////////////////////////////////////////////////////////////
// Mint class

export const Mint = {

  ///////////////////////////////////////////////////////////////////////////////
  // Mint.TxnData class

  TxnData: {

    // Mint.TxnData
    importTransactions(options, interactive, showImportCount)
    {
      try
      {
        if (interactive == undefined) {
          interactive = false;
        }
        if (showImportCount == undefined) {
          showImportCount = true;
        }

        if (interactive && showImportCount) toast("Starting ...", "Mint transaction import");

        var range = Utils.getTxnDataRange();

        // If replacing existing txns, prompt user to make sure
        if (options.replaceExistingData && range.getNumRows() > 1) {
          var button = Browser.msgBox("Replace existing transactions?", "Are  you sure you want to REPLACE the " + range.getNumRows() + " existing transaction(s)?", Browser.Buttons.OK_CANCEL);
          if (button === "cancel")
            return;
        }

        // Clear all txn formatting (row colors, text colors, bold, italics, etc.)
        // We will re-apply it after the txns have been imported
        range.clear({formatOnly: true});

        var cookies = Mint.Session.getCookies();
        if (!cookies) {
          throw new Error('Mint session has expired');
        }

        if (interactive) {
          if (showImportCount) {
            toast("Retrieving transaction data", "Mint transaction import", 60);
          } else {
            toast("Retrieving the latest transaction data", "Mint transaction import", 60);
          }
        }

        var allTxns = [];
        var offset = 0;
        var progress = 0;
        var mintAccount = Settings.getSetting(Const.IDX_SETTING_MINT_LOGIN);
        var importDate = new Date();

        do
        {
          var startDate = new Date(options.startDate);
          var endDate = new Date(options.endDate);

          var txns = this.downloadTransactions(cookies, offset, startDate, endDate);
          if (!txns || txns.length <= 0)
            break;
          
          var txnValues = this.copyDataIntoValueArray(txns, importDate, mintAccount);
          // Use push.apply() for fast, in-place concatenation
          allTxns.push.apply(allTxns, txnValues);

          offset += txns.length;
          
          // Display toast with status every 500 txns
          progress += txns.length;
          if (progress >= 500)
          {
            if (interactive && showImportCount) toast("Downloaded " + offset + " transactions. More coming ...", "Mint transaction import", 10);
            progress -= 500;
            
//            if (offset % 1000 === 0 && "no" === Browser.msgBox("Mint transaction import", String(offset) + " transactions have been downloaded. Continue?", Browser.Buttons.YES_NO)) {
//              break;
//            }
          }
          
        } while (true);

        if (interactive && showImportCount) toast("Download complete. Importing " + offset + " transactions into Mojito", "Mint transaction import", 60);

        Sheets.TxnData.insertData(allTxns, options.replaceExistingData);

        if (interactive) {
          if (showImportCount) {
            toast("A total of " + offset + " transactions were imported.", "Mint transaction import", 5);
          } else {
            toast("Transactions imported.", "Mint transaction import", 5);
          }
        }
      }
      catch(e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }
      
    },

    // Mint.TxnData
    downloadTransactions(cookies, offset, startDate, endDate, secondAttempt)
    {
      if (!cookies)
        return null;
      
      var startDateString = Utilities.formatString("%d/%d/%d", startDate.getMonth() + 1, startDate.getDate(), startDate.getFullYear());
      var endDateString = Utilities.formatString("%d/%d/%d", endDate.getMonth() + 1, endDate.getDate(), endDate.getFullYear());
      
      var headers = {
        "Cookie": cookies,
        "Accept": "application/json"
      };
      var options = {
        "method": "GET",
        "headers": headers,
      };
      
      var url = "https://mint.intuit.com/getJsonData.xevent";
      var queryParams = "?" +
        "startDate=" + startDateString + 
        "&endDate=" + endDateString + 
        "&task=transactions" +
        "&filterType=cash" + 
        "&rnd=" + String(Date.now()) + 
        "&queryNew=" +
        "&offset=" + offset + 
        "&comparableType=0";
      var response = null;

      Debug.log("downloadTransactions: Starting fetch");
      
      var txnData = null;
      
      try
      {
        response = UrlFetchApp.fetch(url + queryParams, options);
        if (response && (response.getResponseCode() === 200))
        {
          //var respHeaders = response.getAllHeaders();
          //Debug.log(respHeaders.toSource());
          
          var respBody = response.getContentText();
          var respCheck = Mint.checkJsonResponse(respBody);
          if (respCheck.success)
          {
            var results = JSON.parse(respBody);
            //Debug.log("Results: " + results.toSource());

            Debug.assert(results.set[0].id === "transactions", "results.set[0].id === \"transactions\"");
            txnData = results.set[0].data;
            Debug.assert(!!txnData, "txnData !== null");
          }
          else if (respCheck.sessionExpired && !secondAttempt)
          {
            // Re-login, then call this function again to retry
            cookies = Mint.Session.getCookies();
            txnData = this.downloadTransactions(cookies, offset, startDate, endDate, true);
          }
          else
          {
            throw new Error("Data retrieval failed. " + respBody);
          }
        }
        else
        {
          toast("Data retrieval failed. HTTP status code: " + response.getResponseCode(), "Update error");
        }
      }
      catch(e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }
      
      return txnData;
    },

    // Mint.TxnData
    updateTransaction (cookies, updateFormData, secondAttempt)
    {
      var result = { success: false, responseJson: null };

      if (!cookies)
      {
        cookies = Mint.Session.getCookies();
      }

      try
      {
        var response = null;
 
        Debug.log("updateTransaction: updateFormData: " + updateFormData.toSource());

        var headers = {
          "Cookie": cookies,
          "Accept": "application/json",
          "Origin": "https://mint.intuit.com",
          "Referer": "https://mint.intuit.com/transaction.event",
        };

        var options = {
          "method": "POST",
          "headers": headers,
          "payload": updateFormData,
          "followRedirects": false
        };
        
        response = UrlFetchApp.fetch("https://mint.intuit.com/updateTransaction.xevent", options);
        var respBody = "";

        if (response && ((response.getResponseCode() === 200) || response.getResponseCode() === 302))
        {
          //var respHeaders = response.getAllHeaders();
          //Debug.log("Response headers: " + respHeaders.toSource());

          respBody = response.getContentText();
          if (Debug.enabled) Debug.log("updateTransaction: Response Body: " + respBody);
          var respCheck = Mint.checkJsonResponse(respBody);
          if (respCheck.success)
          {
            result.responseJson = JSON.parse(respBody);

            result.success = (result.responseJson.task === updateFormData["task"]);
          }
          else if (respCheck.sessionExpired && !secondAttempt)
          {
            // Re-login, then call this function again to retry
            Debug.log("updateTransaction: Session expired. Refetching session token");
            var token = Mint.Session.getSessionToken(cookies, true);
            if (token) {
              updateFormData["token"] = token;
              result = this.updateTransaction(cookies, updateFormData, true);
            }
          }
        }

        if (!result.success)
        {
            throw new Error("Update failed. Status code: " + response.getResponseCode() + ", body:  " + respBody);
        }
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        toast("Update failed. Exception encountered: " + e.toString());
//        Browser.msgBox("Error: " + e.toString());
      }

      return result;
    },

    // Mint.TxnData
    copyDataIntoValueArray(txnData, importDate, mintAccount)
    {
      var numRows = txnData.length;
      var numCols = Const.IDX_TXN_LAST_COL + 1;
      var txnValues = [];

      var tagCleared = Mint.getClearedTag();
      var tagReconciled = Mint.getReconciledTag();
      
      for (var i = 0; i < numRows; ++i)
      {
        var currRow = txnData[i];

        if (currRow.date === "NOT_FOUND" || isNaN(currRow.id) || currRow.amount === "NOT_FOUND") {
          Browser.msgBox("Skipping invalid transaction: " + JSON.stringify(currRow));
          continue;
        }
        
        if (currRow.isDuplicate)
        {
          // Skip txns flagged as duplicates
          if (Debug.enabled) Debug.log("Skipping duplicate: %s", JSON.stringify(currRow));
          continue;
        }

        var date = Mint.convertMintDateToDate(currRow.odate || currRow.date, importDate);
        var amount = Number(currRow.amount.replace(/[^\d.-]/g, '')); // strip '$' and ','
        if (currRow.isDebit === true)
          amount = -amount; // change debits to negative amounts

        var cleared = false;
        var reconciled = false;

        var tagArray = currRow.labels;
        var tags = "";
        var tagIds = "";
        for (var j = 0; j < tagArray.length; ++j)
        {
          // Show the "cleared" and "reconciled" tags in c/R column
          var tagName = tagArray[j].name;
          if (tagName === tagCleared) {
            cleared = true;
          }
          else if (tagName === tagReconciled) {
            reconciled = true;
          }
          else {
            tags += tagName + Const.DELIM;
          }
          // Include all tag IDs
          tagIds += tagArray[j].id + Const.DELIM;
        }
        
        // See if the memo field contains any Mojito properties
        // If so, move the props to a separate column to avoid confusing the user
        var memo = currRow.note;
        var props = null;
        var propsJson = null;

        if (memo) {
          var extractedParts = this.extractPropsFromString(memo);
          if (extractedParts) {
            memo = extractedParts.text;
            props = extractedParts.props;
            propsJson = extractedParts.propsJson;
          }
        }
        
        // Determine "state" column
        var state = null;
        if (currRow.isPending === true && (!props || props.pending !== "ignore")) {
          state = Const.TXN_STATUS_PENDING;
        } else if (currRow.isChild === true) {
          state = Const.TXN_STATUS_SPLIT;
        }

        var txnRow = new Array(numCols);

        txnRow[Const.IDX_TXN_DATE] = date;
        txnRow[Const.IDX_TXN_EDIT_STATUS] = null;
        txnRow[Const.IDX_TXN_ACCOUNT] = currRow.account;
        txnRow[Const.IDX_TXN_MERCHANT] = currRow.merchant;
        txnRow[Const.IDX_TXN_AMOUNT] = amount;
        txnRow[Const.IDX_TXN_CATEGORY] = currRow.category;
        txnRow[Const.IDX_TXN_TAGS] = tags;
        txnRow[Const.IDX_TXN_CLEAR_RECON] = (reconciled ? "R": (cleared ? "c": null));
        txnRow[Const.IDX_TXN_MEMO] = memo;
        txnRow[Const.IDX_TXN_MATCHES] = "";
        txnRow[Const.IDX_TXN_STATE] = state;
        // Internal values
        txnRow[Const.IDX_TXN_MINT_ACCOUNT] = mintAccount;
        txnRow[Const.IDX_TXN_ORIG_MERCHANT_INFO] = currRow.omerchant;
        txnRow[Const.IDX_TXN_ID] = currRow.id;
        txnRow[Const.IDX_TXN_PARENT_ID] = (currRow.isChild === true ? currRow.pid: null);
        txnRow[Const.IDX_TXN_CAT_ID] = currRow.categoryId;
        txnRow[Const.IDX_TXN_TAG_IDS] = tagIds;
        txnRow[Const.IDX_TXN_MOJITO_PROPS] = propsJson;
        txnRow[Const.IDX_TXN_YEAR_MONTH] = date.getFullYear() * 100 + (date.getMonth() + 1);
        txnRow[Const.IDX_TXN_ORIG_AMOUNT] = amount; // We keep a second copy of txn amount so, if user changes amount, we can compare the before and after value
        txnRow[Const.IDX_TXN_IMPORT_DATE] = importDate;
        
        txnValues.push(txnRow);
      }

      return txnValues;
    },
    
    // Mint.TxnData
    getUpdateFormData(txnRow, isSimpleEdit, editType, fillToken)
    {
      var formData = null;

      var token = null;
      if (fillToken === true) {
        token = Mint.Session.getSessionToken();
        if (!token) {
          toast("Unable to update transaction. Session token could not be obtained.", "Update transaction");
          return null;
        }
        Debug.log("getUpdateFormData: Using token: " + token);
      }

      if (editType === Const.EDITTYPE_DELETE) {
        formData = {
            "task": "delete",
            "token": token,
            "txnId": String(txnRow[Const.IDX_TXN_ID]) + ":0"
          };
      }
      else if (isSimpleEdit && editType !== Const.EDITTYPE_NEW) {

        formData = {
          "task": "simpleEdit",
          "txnId": String(txnRow[Const.IDX_TXN_ID]) + ":0",
          "date": Utilities.formatDate(txnRow[Const.IDX_TXN_DATE], "GMT", "MM/dd/yyyy"),
          "merchant": txnRow[Const.IDX_TXN_MERCHANT],
          "category": txnRow[Const.IDX_TXN_CATEGORY],
          "catId": String(txnRow[Const.IDX_TXN_CAT_ID]),
          "amount": "",
          "token": token,
        };
      }
      else
      {
        var amount = txnRow[Const.IDX_TXN_AMOUNT];
        var memo = txnRow[Const.IDX_TXN_MEMO];
        var propsJson = txnRow[Const.IDX_TXN_MOJITO_PROPS];
        if (propsJson) {
          // If Mojito properties exist, append them to the end of the memo field with a delimeter
          memo = this.appendPropsToString(memo, propsJson);
          if (Debug.enabled) Debug.log("Properties appended to memo: " + propsJson);
        }
        
        formData = {
          "cashTxnType": "on",
          "mtCheckNo": "",
          "price": "",
          "symbol": "",
          "note": memo,
          "isInvestment": "false",
          "catId": String(txnRow[Const.IDX_TXN_CAT_ID]),
          "category": txnRow[Const.IDX_TXN_CATEGORY],
          "merchant": txnRow[Const.IDX_TXN_MERCHANT],
          "date": Utilities.formatDate(txnRow[Const.IDX_TXN_DATE], "GMT", "MM/dd/yyyy"),
          "amount": String(Math.abs(amount)), // change debit/credit to positive value
          "token": token,
        };

        if (editType === Const.EDITTYPE_NEW) {
          var acctInfoMap = Sheets.AccountData.getAccountInfoMap();
          var acctInfo = (!acctInfoMap ? null: acctInfoMap[ String(txnRow[Const.IDX_TXN_ACCOUNT]) ]);
          if (!acctInfo) {
            throw Utilities.formatString("Account \"%s\" not found. Unable to determine account id.", txnRow[Const.IDX_TXN_ACCOUNT]);
          }

          formData["task"] = "txnAdd";
          formData["txnId"] = ":0";
          formData["mtCashSplitPref"] = "1";
          formData["mtAccount"] = String(acctInfo.id);
          formData["mtType"] = "pending-other";
          formData["mtIsExpense"] = (amount < 0 ? "true": "false");
        }
        else {
          var txnId = String(txnRow[Const.IDX_TXN_ID]);

          formData["task"] = "txnEdit";
          formData["txnId"] = txnId + ":0";
          formData["mtCashSplit"] = "on";
          formData["mtAccount"] = "";
          formData["mtType"] = "cash";
        }

        // Add all tags with value "0" first
        var tagMap = Mint.getTagMap();
        for (var prop in tagMap) {
          var tagId = tagMap[prop].tagId;
          if (tagId === "")
            continue;

          formData["tag" + tagId] = "0";
        }

        // Overwrite actual tags for this txn with value "2"        
        var tagIds = txnRow[Const.IDX_TXN_TAG_IDS];
        var tagIdArray = tagIds.split(Const.DELIM);
        for (var i = 0; i < tagIdArray.length; ++i) {
          var tagId = tagIdArray[i];
          if (tagId === "")
            continue;

          formData["tag" + tagIdArray[i]] = "2";
        }
      }

      return formData;
    },

    // Mint.TxnData
    getSplitUpdateFormData(splitRows, txnValues, fillToken)
    {
      var token = null;
      if (fillToken === true) {
        token = Mint.Session.getSessionToken();
        if (!token) {
          toast("Unable to update transaction. Session token could not be obtained.", "Update transaction");
          return null;
        }
        if (Debug.enabled) Debug.log("getSplitUpdateFormData: Using token: " + token);
      }

      // Is this a new split or existing split?
      var firstSplitRow = splitRows[0];
      var parentId = txnValues[firstSplitRow][Const.IDX_TXN_PARENT_ID];

      var formData = {
        "task": "split",
        "data": "",
        "txnId": String(parentId) + ":0",
        "token": token,
      };

      var splitCount = splitRows.length;
      // If there is only one split txn in this group, then we are effectively deleting the split group
      // and reverting back to a 'normal' transaction. The above formData will accomplish this.

      if (splitCount > 1) {
        for (var i = 0; i < splitCount ; ++i) {
          var rowNum = splitRows[i];
          var iStr = String(i);
          formData["amount" + iStr] = String(-txnValues[rowNum][Const.IDX_TXN_AMOUNT]); // Debits are changed to positive
          formData["category" + iStr] = txnValues[rowNum][Const.IDX_TXN_CATEGORY];
          formData["merchant" + iStr] = txnValues[rowNum][Const.IDX_TXN_MERCHANT];
          formData["txnId" + iStr] = String(txnValues[rowNum][Const.IDX_TXN_ID]) + ":0";
          formData["percentAmount" + iStr] = "0";
          formData["categoryId" + iStr] = String(txnValues[rowNum][Const.IDX_TXN_CAT_ID]);
        }
      }

      return formData;
    },

    // Mint.TxnData
    appendPropsToString(strValue, propsJson) {
      return Utilities.formatString("%s\n\n\n%s%s", (strValue || ''), Const.DELIM_2, propsJson);
    },

    // Mint.TxnData
    extractPropsFromString(strValue) {
      var extractedParts = {
        text: null,
        props: null,
        propsJson: null,
      };

      var propDelim = strValue.indexOf(Const.DELIM_2);
      if (propDelim >= 0) {
        extractedParts.propsJson = strValue.substr(propDelim + Const.DELIM_2.length);
        if (Debug.enabled) Debug.log("Mojito props: " + extractedParts.propsJson);
        try {
          extractedParts.props = JSON.parse(extractedParts.propsJson);
        } catch (e) {
          if (Debug.enabled) Debug.log("Unable to parse mojito props. " + e.toString());
          extractedParts.props = null;
        }
        extractedParts.text = strValue.substr(0, propDelim).trim();

      } else {
        extractedParts.text = strValue;
      }

      return extractedParts;
    },
    
  },

  ///////////////////////////////////////////////////////////////////////////////
  // Mint.AccountData

  AccountData: {

    mintTimestampForToday: 0,

    // Mint.AccountData
    downloadAccountInfo()
    {
      var cookies = Mint.Session.getCookies();
      var jsonData = Mint.downloadJsonData(cookies, "accounts", null);
      if (Debug.enabled) Debug.log("%s account(s) found", (!jsonData ? "0": String(jsonData.length)));
      
      return jsonData;
    },

    // Mint.AccountData
    downloadBalanceHistory(cookies, account, startDate, endDate, importTodaysBalance, interactive)
    {
      if (!cookies)
        return null;

      var balanceHistory = [];
      var mintTimestamp = this.getTodayTimestamp();

      try
      {
        // Determine the type of trend data to fetch, "asset" or "debt"
        var reportType = "AT"; // assume asset
        if (account.klass === "credit" || account.klass === "loan") {
          reportType = "DT"; // debt
        }

        var secondAttempt = false;
        var daysToFetch = Math.max(0, (endDate - startDate)/Const.ONE_DAY_IN_MILLIS);
        var iterStartDate = startDate;
        if (Debug.enabled) Debug.log("Retrieving %d days of balance history for account %s, from '%s' to '%s'", daysToFetch, account.id, startDate.toString(), endDate.toString());

        while (daysToFetch > 0) {
          // If we want to get daily account balances (which we do) then
          // we can only fetch about 40 days at a time

          var iterEndDate = new Date(iterStartDate.getFullYear(), iterStartDate.getMonth(), iterStartDate.getDate() + 40);
          if (iterEndDate > endDate) {
            iterEndDate = endDate;
          }

          var strStart = Utilities.formatDate(iterStartDate, "GMT", "MM/dd/yyyy");
          var strEnd = Utilities.formatDate(iterEndDate, "GMT", "MM/dd/yyyy");
          if (Debug.enabled) Debug.log("Fetching account trend data for date range %s - %s", strStart, strEnd);

          // After a lot of trial-and-error, it seems we must send the form data already html-encoded.
          var formData = Utilities.formatString("searchQuery=%%7B%%22reportType%%22%%3A%%22%s%%22%%2C%%22chartType%%22%%3A%%22P%%22%%2C%%22comparison%%22%%3A%%22%%22%%2C%%22matchAny%%22%%3Atrue%%2C%%22terms%%22%%3A%%5B%%5D%%2C%%22accounts%%22%%3A%%7B%%22groupIds%%22%%3A%%5B%%5D%%2C%%22accountIds%%22%%3A%%5B%d%%5D%%2C%%22count%%22%%3A1%%7D%%2C%%22dateRange%%22%%3A%%7B%%22period%%22%%3A%%7B%%22label%%22%%3A%%22Custom%%22%%2C%%22value%%22%%3A%%22CS%%22%%7D%%2C%%22start%%22%%3A%%22%s%%22%%2C%%22end%%22%%3A%%22%s%%22%%7D%%2C%%22drilldown%%22%%3Anull%%2C%%22categoryTypeFilter%%22%%3A%%22all%%22%%7D&token=%s",
                  reportType, account.id, strStart, strEnd, Mint.Session.getSessionToken(cookies, false));
          /* The encoded formData above is equivalent to this: 

            var formData = {
              "searchQuery": {
                "reportType": "AT",  (or "DT" for debt accounts)
                "chartType": "P",
                "comparison": "",
                "matchAny": true,
                "terms": [],
                "accounts": {
                  "groupIds": [],
                  "accountIds": [ 1633319 ],
                  "count": 1
                },
                "dateRange": {
                  "period": {
                    "label": "Custom",
                    "value": "CS"
                  },
                  "start": "12/1/2013",
                  "end": "12/31/2013"
                },
                "drilldown": null,
                "categoryTypeFilter": "all"
              },
              "token": Mint.Session.getSessionToken(cookies, true),
            };
          */
          //Debug.log("formData:  %s", formData);

          var headers = {
            "Cookie": cookies,
            "Accept": "application/json",
            "Origin": "https://mint.intuit.com",
            "Referer": "https://mint.intuit.com/trend.event",
          };
          var options = {
              "method": "POST",
              "headers": headers,
              "payload": formData,
              "followRedirects": false,
              "escaping": false,
          };
    
          if (Debug.enabled) Debug.log(Utilities.formatString("downloadBalanceHistory(%d): Starting fetch", account.id));

          var url = "https://mint.intuit.com/trendData.xevent";
          var response = UrlFetchApp.fetch(url, options);

          if (Debug.enabled) Debug.log(Utilities.formatString("downloadBalanceHistory(%d): Fetch completed. Response code: %d", account.id, (response ? response.getResponseCode(): "<unknown>")));
  
          if (response && (response.getResponseCode() === 200 || response.getResponseCode() === 302))
          {
            //var respHeaders = response.getAllHeaders();
            //Debug.log(respHeaders.toSource());

            var respBody = response.getContentText();
            //if (Debug.enabled) Debug.log("trendData.xevent:  %s", respBody);
            var respCheck = Mint.checkJsonResponse(respBody);
            if (respCheck.success)
            {
              var results = JSON.parse(respBody);
              //if (Debug.enabled) Debug.log("Results: " + results.toSource());

              // Success. We got the account balances.
              // The results.trendList array contains an entry for each day
              Debug.assert(results.granularity === "DAY", "results.granularity === \"DAY\"");
              var data = results.trendList;
              Debug.assert(!!data, "data !== null && data !== undefined");

              if (data.length === 0) {
                // The trendList was empty. Should we use today's balance in account.bal?
                if (Debug.enabled) Debug.log("trendList data was empty for account '%s'.", account.name);
                var today = new Date();
                var todayTime = new Date(today.getFullYear(), today.getMonth(), today.getDate());
                var startTime = startDate.getTime();
                var endTime = endDate.getTime();

                if (importTodaysBalance === undefined) {
                  var includeToday = (startTime <= todayTime && todayTime <= endTime);
                  if (includeToday) {
                    account.balanceHistoryNotAvailable = true;
                    importTodaysBalance = true;
                  } else {
                    var msg = `Account balance history does not exist for account '${account.name}'. Would you like import today's balance even though it falls outside your specified date range?`;
                    var choice = Browser.msgBox("Account without history detected", msg, Browser.Buttons.YES_NO);
                    importTodaysBalance = (choice === "yes");
                  }
                  Settings.setInternalSetting(Const.IDX_INT_SETTING_CURR_DAY_ACCT_IMPORT, importTodaysBalance);

                }

                if (importTodaysBalance === true) {
                  // We will "fail gracefully" and use today's balance from the account info.
                  if (Debug.enabled) Debug.log("Substituting 'account.bal' for today's balance: %s", account.name, account.bal);

                  data = [{ date: mintTimestamp, value: account.bal }];
                }
              } else {
                // trendList data is available. Use the date of the first entry to calculate the time
                // offset from UTC. We may need to use this for accounts that do not have any trendList data.
                this.setTodayTimestamp(data[0].date);
              }
              // Append this data to the end of the balanceHistory array
              balanceHistory.push.apply(balanceHistory, data);
              
            }
            else if (respCheck.sessionExpired && !secondAttempt)
            {
              // Re-login, then try again
              cookies = Mint.Session.getCookies();
              secondAttempt = true;
              continue;
            }
            else
            {
              throw new Error("Data retrieval failed. " + respBody);
            }
          }
          else
          {
            throw new Error("Data retrieval failed. HTTP status code: " + response.getResponseCode());
          }

          iterStartDate = new Date(iterStartDate.getTime() + (41 * Const.ONE_DAY_IN_MILLIS));
          if (Debug.enabled) Debug.log("iterStartDate=%s", iterStartDate.toString());
          daysToFetch = Math.max(0, (endDate - iterStartDate)/Const.ONE_DAY_IN_MILLIS);
          if (Debug.enabled) Debug.log("daysToFetch=%d", daysToFetch);

        } // while

        // Add a 'balanceHistory' member to the account
        account.balanceHistory = balanceHistory;
      }
      catch(e)
      {
        if (Debug.enabled) Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }

      return account;
    },

    setTodayTimestamp(mintTimestamp) {
      // mint timestamp could be for any date. We will calculate the mint time offset
      // from UTC, then use that to calculate the timestamp for today.

      var mintDate = new Date(mintTimestamp);
      var utcTimestamp = Date.UTC(mintDate.getFullYear(), mintDate.getMonth(), mintDate.getDate());
      var timeOffset = mintDate.getTime() - utcTimestamp;
      if (Debug.enabled) Debug.log("Mint import timestamp offset from UTC: %d hours", timeOffset/(60*60*1000));

      var today = new Date();
      this.mintTimestampForToday = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()) + timeOffset;
      if (Debug.enabled) Debug.log("Calculated mint timestamp for today: %s", this.mintTimestampForToday);
    },

    getTodayTimestamp() {
      // If no balance history is available for the specified account, then we use the
      // current balance in account.bal and associate it with today's date (adjusted to match Mint).
      // This function calculates the timestamp for today;
      var timestamp = this.mintTimestampForToday;

      if (timestamp === 0) {
        var today = new Date();
        // Parse timezone info from date string to determine if this it is daylight savings time or not
        var todayStr = today.toString();
        var tzStr = todayStr.match(/\([A-Z]+\)/g); // example: "(MDT)" = Moutain Daylight Time
        var dstChar = (tzStr ? String(tzStr).charAt(2): null); // Parse second char. Should be 'D' or 'S'
        if (Debug.enabled) Debug.log("Determining if it is daylight savings time: %s, '%s'", tzStr, dstChar);
        var isDST = (dstChar ? (dstChar === 'D' ? true: false): false);
        // Determine time offset for Pacific Time (that's what Mint uses)
        var mintTimeOffset = (isDST === true ? 7*60*60*1000 /*7 hours*/: 8*60*60*1000 /*8 hours*/);
        if (Debug.enabled) Debug.log("timezone offset: %d hours", mintTimeOffset / (60*60*1000));
        timestamp = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()) + mintTimeOffset;
        if (Debug.log) Debug.log("Estimated mint timestamp for today: %s", timestamp);
      }
      return timestamp;
    },

  },

  ///////////////////////////////////////////////////////////////////////////////
  // Mint.Categories class

  Categories: {

    // Mint.Categories
    import(interactive)
    {
      try
      {
        if (interactive) toast("Starting update", "Category update", 30);

        var cookies = Mint.Session.getCookies();

        if (interactive) toast("Retrieving categories", "Category update");

        var jsonData = Mint.downloadJsonData(cookies, "categories", null);

        // Clear cached category map
        Utils.getPrivateCache().remove(Const.CACHE_CATEGORY_MAP);

        if (interactive) toast("Retrieved " + jsonData.length + " categories", "Category update");

        this.insertDataIntoSheet(jsonData, interactive);
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        toast("Import failed. Exception encountered: " + e.toString(), "Error", 20);
      }
    },

    // Mint.Categories
    insertDataIntoSheet(categoryData, interactive)
    {
      if (!categoryData || categoryData.length === 0)
        return;
      
      var numTopLevelCategories = categoryData.length;
      var numCols = Const.IDX_CAT_LAST_COL + 1;
      Debug.log("Looping through " + numTopLevelCategories + " top-level categories");
      
//      Browser.msgBox(categoryData.toSource());
      var catValues = new Array();

      for (var i = 0; i < numTopLevelCategories; ++i) {
        var currRow = categoryData[i];
        var targetRow = new Array(numCols);

        if (isNaN(currRow.id))
        {
          if (interactive) Browser.msgBox("Skipping invalid category, " + currRow.value);
          Debug.log(Utilities.formatString("Skipping import of invalid category. Name: %s", currRow.value));
          continue;
        }
        
        targetRow[Const.IDX_CAT_NAME] = currRow.value;
        targetRow[Const.IDX_CAT_ID] = currRow.id;
        targetRow[Const.IDX_CAT_STANDARD] = true;
        targetRow[Const.IDX_CAT_PARENT_ID] = 0;
        // Add row to catValues array
        catValues.push(targetRow);
        
        if (currRow.children && currRow.children.length > 0) {
          var childCount = currRow.children.length;

          for (var j = 0; j < childCount; ++j) {
            var childRow = currRow.children[j];
            var targetRow = new Array(numCols);

            targetRow[Const.IDX_CAT_NAME] = childRow.value;
            targetRow[Const.IDX_CAT_ID] = childRow.id;
            targetRow[Const.IDX_CAT_STANDARD] = childRow.isStandard;
            targetRow[Const.IDX_CAT_PARENT_ID] = currRow.id;
            Debug.assert(childRow.children == undefined, "childRow.children == undefined");
            // Add child row to catValues array
            catValues.push(targetRow);
          } // for j
        }
      } // for i

      Debug.log("Inserting " + catValues.length + " categories");

      var range = Utils.getCategoryDataRange();
      range.clear(); // Replace existing categories
      var catRange = range.offset(0, 0, catValues.length, numCols);
      catRange.setValues(catValues);

      if (interactive) toast("Inserted " + catValues.length + " categories", "Category update");
    },
  },
  
  ///////////////////////////////////////////////////////////////////////////////
  // Mint.Tags class

  Tags: {

    // Mint.Tags
    import(interactive)
    {
      try
      {
        if (interactive) toast("Starting update", "Tag update", 30);

        var cookies = Mint.Session.getCookies();

        if (interactive) toast("Retrieving categories", "Tag update");

        var jsonData = Mint.downloadJsonData(cookies, "tags", null);

        // Clear cached tag map
        Utils.getPrivateCache().remove(Const.CACHE_TAG_MAP);

        //Browser.msgBox(jsonData.toSource());
        if (interactive) toast("Retrieved " + jsonData.length + " tags", "Tag update");

        this.insertDataIntoSheet(jsonData, interactive);
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        toast("Import failed. Exception encountered: " + e.toString(), "Error", 20);
      }
    },

    // Mint.Tags
    insertDataIntoSheet(tagData, interactive)
    {
      if (!tagData || tagData.length === 0)
        return;
      
      var numRows = tagData.length;
      var numCols = Const.IDX_TAG_LAST_COL + 1;
      Debug.log("Looping through " + numRows + " tags");
      
      var tagValues = new Array();

      for (var i = 0; i < numRows; ++i) {
        var currRow = tagData[i];
        var targetRow = new Array(numCols);

        if (isNaN(currRow.id))
        {
          if (interactive) Browser.msgBox("Skipping invalid category, " + currRow.value);
          Debug.log(Utilities.formatString("Skipping import of invalid category. Name: %s", currRow.value));
          continue;
        }
        
        targetRow[Const.IDX_TAG_NAME] = currRow.value;
        targetRow[Const.IDX_TAG_ID] = currRow.id;
        // Add row to tagValues array
        tagValues.push(targetRow);
       
      } // for i

      Debug.log("Inserting " + tagValues.length + " tags");

      var range = Utils.getTagDataRange();
      range.clear(); // Replace existing tags
      var tagRange = range.offset(0, 0, tagValues.length, numCols);
      tagRange.setValues(tagValues);

      if (interactive) toast("Inserted " + tagValues.length + " tags", "Tag update");
    },

  },

  ///////////////////////////////////////////////////////////////////////////////
  // Mint.Session class

  Session: {

    resetSession() {
      var cache = Utils.getPrivateCache();
      cache.remove(Const.CACHE_SESSION_TOKEN);
      
      Mint.Session.clearCookies();
    },

    getCookies(throwIfNone = true)
    {
      var cache = Utils.getPrivateCache();
      var cookies = cache.get(Const.CACHE_LOGIN_COOKIES);
      var mintAccount = Utils.getMintLoginAccount();

      if (cookies) {
        // Put the same cookies and mint account back in the cache to reset the expiration
        cache.put(Const.CACHE_LOGIN_COOKIES, cookies, Const.CACHE_SESSION_EXPIRE_SEC);
        cache.put(Const.CACHE_LOGIN_ACCOUNT, mintAccount, Const.CACHE_SESSION_EXPIRE_SEC);
      }
      else
      {
        if (Debug.enabled) Debug.log("Mint.Session.getCookies: No login cookies in cache");

        // If cookies are not in the cache, prompt the user to provide mint auth data.
        if (throwIfNone) {
          throw new Error("Mint authentication has expired.");
        }
      }

      return cookies;
    },

    clearCookies()
    {
      Debug.log("Clearing login cookies and token from cache.");
      var cache = Utils.getPrivateCache();
      cache.remove(Const.CACHE_LOGIN_COOKIES);
      cache.remove(Const.CACHE_LOGIN_ACCOUNT);
    },

    getCookiesFromResponse(response) {
      var respHeaders = response.getAllHeaders();
      var setCookieArray = respHeaders["Set-Cookie"];
      if (!setCookieArray) {
        setCookieArray = [];
      }
      // Make sure the setCookieArray is actually an array, and not just a single string
      if (typeof setCookieArray === 'string') {
        setCookieArray = [setCookieArray];
      }
      if (Debug.traceEnabled) Debug.trace("Cookies in response: " + setCookieArray.toSource());

      var cookies = {};
      for (var i = 0; i < setCookieArray.length; ++i)
      {
        var cookie = setCookieArray[i];
        var cookieParts = cookie.split('; ');
        //Debug.log('cookieParts: ' + cookieParts.toSource());
        cookie = cookieParts[0].split('=');
        //Debug.log('cookie: ' + cookie.toSource());
        cookies[ cookie[0] ] = cookie[1];
      }

      //Debug.log('******** cookies: ' + cookies.toSource());      
      return cookies;
    },

    showManualMintAuth() {
      try {
        var htmlOutput = HtmlService.createTemplateFromFile('manual_mint_auth.html').evaluate();
        htmlOutput.setHeight(350).setWidth(600).setSandboxMode(HtmlService.SandboxMode.IFRAME);
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Mint Authentication - HTTP headers');
      }
      catch (e) {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }
    },

    /**
     * Called from manual_mint_auth.html
     * @param args
     * @returns {boolean}
     */
    verifyManualAuth(args) {
      var success = false;

      try {
        var cache = Utils.getPrivateCache();
        var cookies = args.cookies;
        Debug.log("Cookies from HTTP headers: " + cookies);

        var token = args.token;
        Debug.log("Token from HTTP headers: " + token);
        var result = Mint.Session.fetchTokenAndUsername(cookies);
        let retrievedToken = result && result.token;

        success = (retrievedToken && token === retrievedToken);
        if (success) {
          // Cache cookies and token
          cache.put(Const.CACHE_LOGIN_COOKIES, cookies, Const.CACHE_SESSION_EXPIRE_SEC);
          cache.put(Const.CACHE_SESSION_TOKEN, token, Const.CACHE_SESSION_EXPIRE_SEC);

          // Save username (email)
          if (result.username) {
            Settings.setSetting(Const.IDX_SETTING_MINT_LOGIN, result.username);
          }
        }
        toast('Mint authentication ' + (success ? 'succeeded': 'FAILED'));
      }
      catch (e) {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }
      
      return success;
    },

    /**
     * @obsolete This function no longer works.
     */
    loginMintUser(username, password)
    {
      if (!username) {
        return null;
      }

      password = password || '';

      username = username.toLowerCase();
      var isDemoUser = (username == Const.DEMO_MINT_LOGIN);
      
      var msg = Utilities.formatString("Logging in user %s. %s", username, (isDemoUser ? " Note: The demo user can take a while. Be patient.": ""));
      toast(msg, "Mint Login", 60);
      Debug.log("Logging in user " + username);

      var cookies = null;
      
      try
      {
        var podCookie = Mint.Session.getUserPodCookie(username);
        var response = null;
        
        if (isDemoUser)
        {
          // login to mint demo account
          
          var options = {
            "method": "GET",
            "followRedirects": false
          };
          
          // call the demo login page
          response = UrlFetchApp.fetch("https://mint.intuit.com/demoUser.event", options);
        }
        else // normal login
        {
          var headers = {
            "Cookie": podCookie,
            "Accept": "application/json",
            "X-Request-With": "XMLHttpRequest",
            "X-NewRelic-ID": "UA4OVVFWGwEGV1VaBwc=",
            "Referrer": "https://mint.intuit.com/login.event?task=L&messageId=1&country=US&nextPage=overview.event"
          };

          var formData = {
            "username": username,
            "password": password,
            "task": "L",
            "timezone": Utils.getTimezoneOffset(),
            "browser": "Chrome",
            "browserVersion": 39,
            "os": "win"
          };
          
          var options = {
            "method": "POST",
            "headers": headers,
            "payload": formData,
            "followRedirects": false
            //    "muteHttpExceptions": true
          };
          
          // call the login page
          response = UrlFetchApp.fetch("https://mint.intuit.com/loginUserSubmit.xevent", options);
        }

        Debug.log("Response code: " + response.getResponseCode());
        if (response && (response.getResponseCode() == 200 || response.getResponseCode() == 302))
        {
          var respBody = response.getContentText();
          Debug.log("Response: " + respBody);
          var respJson = (respBody ? JSON.parse(respBody): {});
          
          if (respJson.action) {// && respJson.action === 'CHALLENGE') {
            Debug.log('login action: ' + respJson.action);
          }
    
          // get the cookies (including auth info) so we can use them in subsequent json requests
          var respHeaders = response.getAllHeaders();
          if (Debug.traceEnabled) Debug.trace("loginUser: all headers: " + respHeaders.toSource());
          
          var setCookieArray = respHeaders["Set-Cookie"];
          if (!setCookieArray) {
            setCookieArray = [];
          }
          if (Debug.traceEnabled) Debug.trace("loginUser: cookies: " + setCookieArray.toSource());
          var success = (isDemoUser ? true: false);

          // Save all of the cookies returned in the login response
          cookies = "";
          for (var i = 0; i < setCookieArray.length; i++)
          {
            var thisCookie = setCookieArray[i];
            
            // Make sure the login was successful by looking for a specific cookie
            if (!success && thisCookie.indexOf(username) > 0)
            {
              success = true;
            }
            
            cookies += thisCookie + "; ";
          }

          // Add the pod cookie to the login response cookies so we have the full set
          cookies += podCookie;
          
          var cache = Utils.getPrivateCache();
          
          if (success)
          {
            var token = respJson.sUser.token || respJson.CSRFToken;
            if (!token) {
              if (Debug.enabled) Debug.log("Token was not returned in login response. Response text: " + respBody);
            }

            // Login succeeded. Save the cookies and mint account in the cache.
            cache.put(Const.CACHE_LOGIN_COOKIES, cookies, Const.CACHE_SESSION_EXPIRE_SEC);
            cache.put(Const.CACHE_LOGIN_ACCOUNT, username.toLowerCase(), Const.CACHE_SESSION_EXPIRE_SEC);

            Debug.log("Login succeeded");
            toast("Login succeeded");

            if (token) {
              if (Debug.enabled) Debug.log("Saving token in cache: " + token);
              cache.put(Const.CACHE_SESSION_TOKEN, token, Const.CACHE_SESSION_EXPIRE_SEC);
            }
            else {
              // Remove the session token, if any, because this login just changed it.
              if (Debug.enabled) Debug.log("No token was return in login response. Removing existing token (if any) from cache.");
              cache.remove(Const.CACHE_SESSION_TOKEN);
            }

            if (Debug.traceEnabled) Debug.trace("Cookies: %s", cookies);
          }
          else
          {
            // The login wasn't successful. Clear the cookies.
            cache.remove(Const.CACHE_LOGIN_COOKIES);
            cache.remove(Const.CACHE_LOGIN_ACCOUNT);
            cookies = null;
            toast("Login failed.");
            if (Debug.enabled) Debug.log("Login failed: Response: %s", response.getContentText());
          }
        }
        else
        {
          var msg = "Login failed. HTTP status code: " + response.getResponseCode();
          toast(msg);
          if (Debug.enabled) Debug.log(msg);
        }
        
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        toast("Login failed.");
        Browser.msgBox("Error: " + e.toString());
      }
      
      return cookies;
    },

    getIusSession() {
      var headers = {
      };
      var options = {
          "method": "GET",
          "followRedirects": false
      };

      Debug.log("getIusSession: Starting request");
      var cookies = "";

      var response = UrlFetchApp.fetch("https://accounts.intuit.com/xdr.html", options);
      if (response && response.getResponseCode() === 200)
      {
        cookies = Mint.Session.getCookiesFromResponse(response);
        Debug.log('Cookies from iussession request: ' + cookies.toSource());
      }
      else
      {
        throw new Error("getIusSession failed. " + response.getResponseCode());
      }

      return cookies['ius_session'];
    },
    
    clientSignIn(username, password, iusSession) {

      var headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json; charset=utf-8',
        'Cookie': 'ius_session=' + iusSession + ';',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.84 Safari/537.36'
      };
      var options = {
        "method": "POST",
        "headers": headers,
        "payload": {'username': username, 'password': password },
        "followRedirects": false,
        "muteHttpExceptions": true
      };

      Debug.log("clientSignIn: Starting request");
      var cookies = "";

      var response = null;

      // From Google Apps Script, the below http requests always fails with a 503...      
      response = UrlFetchApp.fetch("https://accounts.intuit.com/access_client/sign_in", options);
      if (response)
      {
        Debug.log("Response code: " + response.getResponseCode());
        var respBody = response.getContentText();
        Debug.log("Response:  " + respBody);
        if (response.getResponseCode() !== 200) {
          throw new Error("clientSignIn POST failed. code: " + response.getResponseCode());
        }
      }
      else
      {
        throw new Error("clientSignIn POST failed. " + respBody);
      }

      return respBody;
    },
    
    getUserPodCookie(mintAccount)
    {
      var headers = {
          "Cookie": "mintUserName=\"" + mintAccount + "\"; "
      };
      var options = {
          "method": "POST",
          "headers": headers,
          "payload": { "username": mintAccount },
          "followRedirects": false
      };

      Debug.log("getUserPodCookie: Starting request");
      var podCookie = "";

      try
      {
        var response = null;
        
        response = UrlFetchApp.fetch("https://mint.intuit.com/getUserPod.xevent", options);
        if (response && (response.getResponseCode() === 200))
        {
          var respBody = response.getContentText();
          //Debug.log("Response:  " + respBody);
          
          var respCheck = Mint.checkJsonResponse(respBody);
          if (respCheck.success)
          {
            var results = JSON.parse(respBody);
            //Debug.log("Results: " + results.toSource());
            
            Debug.assert(results.mintPN, "results.mintPN != undefined");

            var respHeaders = response.getAllHeaders();
            var setCookieArray = respHeaders["Set-Cookie"] || [];
            if (Debug.traceEnabled) Debug.trace("getUserPodCookie: cookies: " + setCookieArray.toSource());

            for (var i = 0; i < setCookieArray.length; i++)
            {
              var thisCookie = setCookieArray[i];
              if (thisCookie.indexOf("mintPN") == 0) {
                podCookie += thisCookie + "; ";
                break;
              }
            }
            podCookie += "mintUserName=\"" + mintAccount + "\"; ";
            Debug.log("podCookie: %s", podCookie);
          }
          else
          {
            throw new Error("getUserPodCookie failed. " + respBody);
          }
        }
      }
      catch(e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }

      return podCookie;
    },

    // Mint.Session
    getSessionToken(cookies, force)
    {
      if (force == undefined)
        force = false;

      let token = '';
      const cache = Utils.getPrivateCache();

      if (!force)
      {
        // Is the token still in the cache?
        Debug.log("Looking up token in cache");
        token = cache.get(Const.CACHE_SESSION_TOKEN);
        if (token)
        {
          Debug.log("Token found in cache: " + token);
          return token;
        }
      }

      if (!cookies)
      {
        cookies = Mint.Session.getCookies();

        if (cookies) {
          token = cache.get(Const.CACHE_SESSION_TOKEN);
          if (token) {
            Debug.log("Token found in cache after re-login: " + token);
            return token;
          }
        }
      }

      Debug.log("getSessionToken: Fetching token from Mint overview page...");
      const result = Mint.Session.fetchTokenAndUsername(cookies);

      return (result ? result.token : null);
    },

    fetchTokenAndUsername(cookies) {
      let result = {};
      const cache = Utils.getPrivateCache();

      var headers = {
        "Cookie": cookies
      };
      var options = {
        "method": "GET",
        "headers": headers
      };
      var url = "https://mint.intuit.com/overview.event";

      try
      {
        let response = UrlFetchApp.fetch(url, options);
        if (response && (response.getResponseCode() === 200))
        {
          var respBody = response.getContentText();
          //Debug.log("Session token fetch response:  " + respBody);

          result = this.parseTokenAndUsernameFromOverviewHtml(respBody);
          if (!result) {
            throw new Error('Invalid headers provided.')
          }
          if (result.token) {
            cache.put(Const.CACHE_SESSION_TOKEN, result.token, Const.CACHE_SESSION_EXPIRE_SEC);
          }
          else {
            Debug.log("Unable to parse token from overview.event html");
          }

          if (result.username) {
            cache.put(Const.CACHE_LOGIN_ACCOUNT, result.username, Const.CACHE_SESSION_EXPIRE_SEC);
          }
          else {
            Debug.log("Unable to parse username from overview.event html");
          }
        }
        else
        {
          toast("Data retrieval failed. HTTP status code: " + response.getResponseCode(), "Error");
          // Failure to get session token is likely due to a problem with the login cookies.
          // We'll reset the session so the user will be forced to login again if another attempt is made.
          Mint.Session.resetSession();
        }
      }
      catch(e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e.toString());
      }

      return result;
    },

    parseTokenAndUsernameFromOverviewHtml(html)
    {
      Debug.log('Parsing overview.html for token');
      var parse1 = html.match(/<input\s+type="hidden"\s+id="javascript-user"\s[^>]*value="[^"]*"\s*\/>/gi);
      if (!parse1) {
        Debug.log('Unable to find hidden input with id "javascript-user"');
        return null;
      }
      Logger.log(parse1[0]);

      var parse2 = parse1[0].match(/\{.*\}/gi);
      if (!parse2) {
        Debug.log('Unable to find "javascript-user" JSON value');
        return null;
      }
      Logger.log(parse2[0]);

      var json = parse2[0].replace(/&quot;/g, "\"");
      Debug.log("token json: %s", json);
      var userValues = JSON.parse(json);
      if (!userValues) {
        Debug.log('Unable to parse "javascript-user" value as JSON');
        return null;
      }

      if (!userValues.token) {
        Debug.log('Unable to find "token" property in "javascript-user" JSON');
        return null;
      }

      var token = userValues.token;
      if (Debug.enabled) Debug.log('Retrieved token: %s', token);

      return userValues;
    },
  },

  ///////////////////////////////////////////////////////////////////////////////
  // Helpers

  // Mint
  waitForMintFiRefresh(interactive)
  {
    if (interactive == undefined)
      interactive = false;

    var isMintDataReady = false;
    
    try
    {
      var maxWaitSeconds = Settings.getSetting(Const.IDX_SETTING_MINT_FI_SYNC_TIMEOUT);
      if (Debug.enabled) Debug.log("Waiting a maxiumum of %s seconds for Mint to refresh data from financial institutions.", maxWaitSeconds);
      var elapsedSeconds = 0;
      while (elapsedSeconds < maxWaitSeconds) {
        if (!this.isMintRefreshingData()) {
          Debug.log("Mint is done refreshing account data from financial institutions");
          isMintDataReady = true;
          break;
        }

        // Sleep 10 seconds
        var waitSec = 10;
        if (interactive) {
          var currentWait = (elapsedSeconds > 0 ? Utils.getHumanFriendlyElapsedTime(elapsedSeconds): "");
          toast(Utilities.formatString("Waiting for Mint to sync data from your financial institutions %s", currentWait), "Mint data update", 10);
        }
        
        Utilities.sleep(waitSec * 1000);
        elapsedSeconds += waitSec;
      }

      if (!isMintDataReady) {
        if (Debug.enabled) Debug.log("Mint took too long to refresh FI data.");
        if (interactive) toast("Mint is taking too long to refresh your account data. Try logging in to the mint.com website to make sure your account information is up to date.", "", 10);
      }

    }
    catch (e)
    {
      if (interactive) Browser.msgBox("Unable to determine if Mint is ready. Error: " + e.toString());
    }

    return isMintDataReady;
  },
  
  // Mint
  isMintRefreshingData()
  {
    // This function calls Mint's userStatus API to see if Mint is currently sync'ing account data
    // from the financial institutions.
    
    var isRefreshing = true;
    
    var cookies = Mint.Session.getCookies();
    var headers = {
      "Cookie": cookies,
      "Accept": "application/json"
    };
    var options = {
      "method": "GET",
      "headers": headers,
    };
    
    var url = "https://mint.intuit.com/userStatus.xevent";
    var queryParams = Utilities.formatString("?rnd=%d", Date.now());
    
    //Debug.log("userStatus.xevent: Starting fetch");
    //Debug.log("  -- using cookies: %s", cookies);

    try
    {
      var response = UrlFetchApp.fetch(url + queryParams, options);
      if (response && (response.getResponseCode() === 200))
      {
        //Debug.log(response.getResponseCode());
        
        var respBody = response.getContentText();
        var respCheck = Mint.checkJsonResponse(respBody);
        if (respCheck.success)
        {
          var results = JSON.parse(respBody);
          if (Debug.enabled) Debug.log("isMintRefreshingData: Results: " + results.toSource());

          isRefreshing = results.isRefreshing;
        }
        else
        {
          if (Debug.enabled) Debug.log("isMintRefreshingData: Retrieval failed: " + respBody);
          throw respBody;
        }
      }
      else
      {
        Debug.log("isMintRefreshingData: Retrieval failed. HTTP status code: " + response.getResponseCode());
      }
    }
    catch(e)
    {
      Debug.log(Debug.getExceptionInfo(e));
      throw e;
    }
    
    return isRefreshing;
  },

  // Mint
  downloadJsonData(cookies, task, queryParamMap, secondAttempt)
  {
    if (!cookies)
      return null;
    
    if (queryParamMap == undefined)
      queryParamMap = null;
    
    var headers = {
      "Cookie": cookies,
      "Accept": "application/json"
    };
    var options = {
      "method": "GET",
      "headers": headers,
    };
    
    var url = "https://mint.intuit.com/getJsonData.xevent";
    var queryParams = Utilities.formatString("?task=%s&rnd=%d", task, Date.now());
    if (queryParamMap) {
      for (var param in queryParamMap) {
        queryParams += Utilities.formatString("&%s=%s", param, queryParamsMap[param]);
      }
    }
    
    var response = null;
    
    if (Debug.enabled) Debug.log(Utilities.formatString("downloadJsonData(%s): Starting fetch", task));
    
    var jsonData = null;
    
    try
    {
      response = UrlFetchApp.fetch(url + queryParams, options);
      if (Debug.enabled) Debug.log(Utilities.formatString("downloadJsonData(%s): Fetch completed. Response code: %d", task, (response ? response.getResponseCode(): "<unknown>")));

      if (response && (response.getResponseCode() === 200))
      {
        //var respHeaders = response.getAllHeaders();
        //Debug.log(respHeaders.toSource());

        var respBody = response.getContentText();
        var respCheck = Mint.checkJsonResponse(respBody);
        if (respCheck.success)
        {
          var results = JSON.parse(respBody);
          //Debug.log("Results: " + results.toSource());
          
          Debug.assert(results.set[0].id === task, "results.set[0].id === \"" + task + "\"");
          jsonData = results.set[0].data;
          Debug.assert(!!jsonData, "jsonData !== null");
        }
        else if (respCheck.sessionExpired && !secondAttempt)
        {
          // Re-login, then call this function again to retry
          cookies = Mint.Session.getCookies();
          jsonData = this.downloadJsonData(cookies, task, queryParamMap, true);
        }
        else
        {
          throw new Error("Data retrieval failed. " + respBody);
        }
      }
      else
      {
        toast("Data retrieval failed. HTTP status code: " + response.getResponseCode(), "Download error");
      }
    }
    catch(e)
    {
      if (Debug.enabled) Debug.log(Debug.getExceptionInfo(e));
      Browser.msgBox("Error: " + e.toString());
    }
    
    return jsonData;
  },

  // Mint
  checkJsonResponse(json)
  {
    var result = {
      success: (!!json && (json.indexOf('<error>') < 0)),
      sessionExpired: false
    };

    if (!result.success) {
      Debug.log("Request failed");

      if (json.indexOf("Session has expired") > 0)
      {
        result.sessionExpired = true;
        Debug.log("Session expired");
        Mint.Session.resetSession();
      }
    }
    
    return result;
  },

  // Mint
  convertMintDateToDate(mintDate, today)
  {
    var convertedDate = null;

    if (mintDate.indexOf("/") > 0) {
      // mintDate is formatted as month/day/year
      var dateParts = mintDate.split("/");

      Debug.assert(dateParts.length >= 3, "convertMintDateToDate: Invalid date, less than 3 date parts");

      var year = parseInt(dateParts[2], 10);
      var month = parseInt(dateParts[0], 10) - 1;
      var day = parseInt(dateParts[1], 10);

      if (year < 100) {
        var century = Math.round(today.getFullYear() / 100) * 100;
        year = century + year;
      }

      convertedDate = new Date(year, month, day);

    } else if (mintDate.indexOf(" ") > 0) {
      // mintDate is formatted as "Month Day"
      var dateParts = mintDate.split(" ");
      var year = today.getFullYear();
      var month = Const.MONTH_LOOKUP_1[dateParts[0]];
      var day = parseInt(dateParts[1]);

      if (today.getMonth() < month)
        --year;

      Debug.assert(day > 0, "convertMintDateToDate: day is zero");

      convertedDate = new Date(year, month, day);
    }
    else
    {
      convertedDate = new Date(2000, 0, 1);
    }

    return convertedDate;
  },

  // Mint
  getMintAccounts(fiAccount) {
    var txnRange = Utils.getTxnDataRange(true);
    if (!txnRange) {
      return null;
    }

    var mintAccountMap = [];
    
    var mintAcctRange = txnRange.offset(0, Const.IDX_TXN_MINT_ACCOUNT, txnRange.getNumRows(), 1);
    var mintAcctValues = mintAcctRange.getValues();
    var mintAcctValuesLen = mintAcctValues.length;

    var fiAcctValues = null;
    if (fiAccount) {
      var fiAcctRange = txnRange.offset(0, Const.IDX_TXN_ACCOUNT, txnRange.getNumRows(), 1);
      fiAcctValues = fiAcctRange.getValues();
    }

    for (var i = 0; i < mintAcctValuesLen; ++i) {
      if (fiAcctValues && fiAccount !== fiAcctValues[i][0]) {
        continue; // Financial institution account doesn't match specified account. Skip it.
      }
      var mintAcct = mintAcctValues[i][0];
      if (mintAcct && !mintAccountMap[mintAcct]) {
        mintAccountMap[mintAcct] = true;
        if (Debug.enabled) Debug.log("Found mint account: " + mintAcct);
      }
    }

    var mintAccounts = [];
    for (var mintAcct in mintAccountMap) {
      mintAccounts.push(mintAcct);
    }

    if (Debug.enabled) Debug.log("mintAccounts array: " + mintAccounts);
    return mintAccounts;
  },

  _categoryMap: null,  // This variable is only valid while server-side code is executing. It resets each time.
  
  // Mint
  getCategoryMap() {
    if (!this._categoryMap) {
      var cache = Utils.getPrivateCache();
      var catMap = {};
      //cache.remove(CACHE_CATEGORY_MAP);
      var catMapJson = cache.get(Const.CACHE_CATEGORY_MAP);
      if (catMapJson && catMapJson !== "{}")
      {
        Debug.log("Category map found in cache. Parsing JSON.");
        catMap = JSON.parse(catMapJson);
      }
      else
      {
        Debug.log("Rebuilding category map");
        var range = Utils.getCategoryDataRange();
        var catRange = range.offset(0, 0, range.getNumRows(), Const.IDX_CAT_ID + 1);
        var catValues = catRange.getValues();
        var catCount = catValues.length;
        for (var i = 0; i < catCount; ++i) {
          var catName = catValues[i][Const.IDX_CAT_NAME];
          catMap[ catName.toLowerCase() ] = { catId: catValues[i][Const.IDX_CAT_ID], displayName: catName };
        }

        Debug.log("Saving category map in cache");
        catMapJson = JSON.stringify(catMap);
        cache.put(Const.CACHE_CATEGORY_MAP, catMapJson, Const.CACHE_MAP_EXPIRE_SEC);
      }

      this._categoryMap = catMap;
    }

    return this._categoryMap;
  },

  // Mint
  validateCategory(category, interactive)
  {
    if (interactive == undefined)
      interactive = false;

    var validationInfo = { isValid: false, displayName: "", catId: 0 };
    var catInfo = this.lookupCategoryId(category);
    if (catInfo) {
      validationInfo.isValid = true;
      validationInfo.displayName = catInfo.displayName;
      validationInfo.catId = catInfo.catId;
    }

    if (!validationInfo.isValid && interactive) toast(Utilities.formatString("The category \"%s\" is not valid. Please \"undo\" your change.", category));
    return validationInfo;
  },

  // Mint
  lookupCategoryId(category)
  {
    if (!category)
      return null;

    var catInfo = null;

    var catMap = this.getCategoryMap();

    var catLower = category.toLowerCase();
    var lookupVal = catMap[catLower];
    if (lookupVal) {
      catInfo = lookupVal;
  //    if (Debug.enabled) Debug.log(Utilities.formatString("lookupCategoryId: Found catId %d for category \"%s\"", catInfo.catId, category));
    }
    else {
      if (Debug.enabled) Debug.log(Utilities.formatString("lookupCategoryId: Category \"%s\" not found", category));
    }

    return catInfo;
  },

  _tagMap: null,  // This variable is only valid while server-side code is executing. It resets each time.
  
  // Mint
  getTagMap() {
    if (!this._tagMap) {
      var tagMap = {};

      var cache = Utils.getPrivateCache();
      var tagMapJson = cache.get(Const.CACHE_TAG_MAP);
      if (tagMapJson && tagMapJson !== "{}")
      {
//        Debug.log("Tag map found in cache. Parsing JSON.");
        tagMap = JSON.parse(tagMapJson);
      }
      else
      {
        Debug.log("Rebuilding tag map");
        var range = Utils.getTagDataRange();
        var tagValues = range.getValues();
        var tagCount = tagValues.length;
        for (var i = 0; i < tagCount; ++i) {
          var tagName = tagValues[i][Const.IDX_TAG_NAME];
          tagMap[ tagName.toLowerCase() ] = { displayName: tagName, tagId: tagValues[i][Const.IDX_TAG_ID] };
        }

        Debug.log("Saving tag map in cache");
        tagMapJson = JSON.stringify(tagMap);
//        Debug.log(tagMapJson);
        cache.put(Const.CACHE_TAG_MAP, tagMapJson, Const.CACHE_MAP_EXPIRE_SEC);
        //Debug.log(tagMap.toSource());
      }

      this._tagMap = tagMap;
    }
    
    return this._tagMap;
  },

  composeTxnTagArray(tagsVal, clearReconVal) {
    let tagArray = [];

    const tagCleared = Mint.getClearedTag();
    const tagReconciled = Mint.getReconciledTag();

    const reconciled = (clearReconVal.toUpperCase() === "R");
    const cleared = (clearReconVal !== null && clearReconVal !== ""); // "Cleared" is anything other than "R" or empty
    if (reconciled) {
      // Reconciled transactions are also 'cleared', so we'll include both tags
      tagArray.push(tagReconciled);
      tagArray.push(tagCleared);
    }
    else if (cleared) {
      tagArray.push(tagCleared);
    }

    if (tagsVal) {
      tagArray.push.apply(tagArray, tagsVal.split(Const.DELIM));
    }

    return tagArray;
  },

  // Mint
  /**
   * Validate the tags of a transaction row, separating out 'cleared' and 'reconcied' status.
   * @param tags {string[]}
   * @param [interactive] {boolean}
   * @returns {{isValid: boolean, tagNames: string, tagIds: string, cleared: boolean, reconciled: boolean}}
   */
  validateTxnTags(tagArray, interactive)
  {
    let validationInfo = {
      isValid: true,
      tagNames: '',
      tagIds: '',
      cleared: false,
      reconciled: false
    };

    if (!tagArray) {
      return validationInfo;
    }

    if (interactive == undefined)
      interactive = false;

    var tagNames = '';
    var tagIds = '';
    var tagCleared = Mint.getClearedTag();
    var tagReconciled = Mint.getReconciledTag();

    var tagMap = this.getTagMap();
    for (var i = 0; i < tagArray.length; ++i) {
      var tag = tagArray[i].trim();
      if (tag === '')
        continue;

      var lookupVal = tagMap[tag.toLowerCase()];
      if (lookupVal) {
//        Debug.log(lookupVal.toSource());
//        if (Debug.enabled) Debug.log(Utilities.formatString("Found tagId %d for tag \"%s\"", lookupVal.tagId, tag));

        // Build list of tags (using exact case from Mint).
        // Cleared and reconciled tags are handled separately.
        if (lookupVal.displayName === tagCleared) {
          validationInfo.cleared = true;
        }
        else if (lookupVal.displayName === tagReconciled) {
          validationInfo.reconciled = true;
        }
        else {
          tagNames += lookupVal.displayName + Const.DELIM;
        }

        // Build list of tag IDs
        tagIds += lookupVal.tagId + Const.DELIM;
      }
      else {
        if (interactive) { toast(Utilities.formatString("Tag \"%s\" is not valid. You must add this tag using the Mint website or mobile app.", tag)); }
        if (Debug.enabled) Debug.log(Utilities.formatString("No tagId found for tag \"%s\"", tag));
        validationInfo.isValid = false;
        break;
      }
    }

    if (validationInfo.isValid) {
      validationInfo.tagNames = tagNames;
      validationInfo.tagIds = tagIds;
    }

    return validationInfo;
  },

  _clearedTag: undefined, // This variable is only valid while server-side code is executing. It resets each time.

  // Mint
  getClearedTag() {
    if (this._clearedTag === undefined) {
      var tag = Utils.getPrivateCache().get(Const.CACHE_SETTING_CLEARED_TAG);
      if (tag === null) {
        tag = Settings.getSetting(Const.IDX_SETTING_CLEARED_TAG);
        // If setting is empty, store empty string in cache
        Utils.getPrivateCache().put(Const.CACHE_SETTING_CLEARED_TAG, tag || '', 300); // Save the tag in the cache for 5 minutes
  
        if (Debug.enabled) Debug.log("getClearedTag: %s", tag);
      }

      this._clearedTag = tag || null;
    }

    return this._clearedTag;
  },

  _reconciledTag: undefined, // This variable is only valid while server-side code is executing. It resets each time.

  // Mint
  getReconciledTag() {
    if (this._reconciledTag === undefined) {
      var tag = Utils.getPrivateCache().get(Const.CACHE_SETTING_RECONCILED_TAG);
      if (tag === null) {
        tag = Settings.getSetting(Const.IDX_SETTING_RECONCILED_TAG);
        // If setting is empty, store empty string in cache
        Utils.getPrivateCache().put(Const.CACHE_SETTING_RECONCILED_TAG, tag || '', 300); // Save the tag in the cache for 5 minutes
  
        if (Debug.enabled) Debug.log("getReconciledTag: %s", tag);
      }

      this._reconciledTag = tag || null;
    }
    
    return this._reconciledTag;
  }
};
