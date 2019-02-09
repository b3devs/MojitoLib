'use strict';
/*
 * Copyright (c) 2013-2019 b3devs@gmail.com
 * MIT License: https://spdx.org/licenses/MIT.html
 */

import {Const} from './Constants.js'
import {Mint} from './MintApi.js';
import {Sheets} from './Sheets.js';
import {Utils, Settings, EventServiceX, toast} from './Utils.js';
import {Upgrade} from './Upgrade.js';
import {Debug} from './Debug.js';


export const Ui = {

  Menu: {

    setupMojitoMenu: function(isOnOpen) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      // Add a custom "Mojito" menu to the active spreadsheet.
      let entries = [];

//      entries.push({ name: "Show / hide sidebar",           functionName: "onMenu_" + Const.ID_TOGGLE_SIDEBAR });
      entries.push({ name: "Connect Mojito to your active Mint session",  functionName: "onMenu_" + Const.ID_SET_MINT_AUTH });
      entries.push(null); // separator

      entries.push({ name: "Sync all with Mint",            functionName: "onMenu_" + Const.ID_SYNC_ALL_WITH_MINT });
      entries.push({ name: "Sync: Import txn data",         functionName: "onMenu_" + Const.ID_IMPORT_TXNS });
      entries.push({ name: "Sync: Import account balances", functionName: "onMenu_" + Const.ID_IMPORT_ACCOUNT_DATA});
      entries.push({ name: "Sync: Save txn changes",        functionName: "onMenu_" + Const.ID_UPLOAD_CHANGES});
      entries.push(null); // separator
      entries.push({ name: "Reconcile an account",          functionName: "onMenu_" + Const.ID_RECONCILE_ACCOUNT});
      // Don't show following menu item until it's implemented
      //entries.push(null); // separator
      //entries.push({ name: "Check for Mojito updates",      functionName: "onMenu_" + Const.ID_CHECK_FOR_UPDATES});

      if (Debug.enabled) {
        entries.push(null); // separator
        entries.push({ name: "Display log window", functionName: "onMenu_displayLogWindow" });
      }

      if (isOnOpen) {    
        ss.addMenu("Mojito", entries);
      } else {
        ss.updateMenu("Mojito", entries);
      }
    },

    onMenu: function(id) {
      switch (id) {
        case Const.ID_SYNC_ALL_WITH_MINT:
          this.onSyncAllWithMint();
          break;

        case Const.ID_IMPORT_TXNS:
          this.onImportTxnData();
          break;

        case Const.ID_IMPORT_ACCOUNT_DATA:
          this.onImportAccountData();
          break;

        case Const.ID_UPLOAD_CHANGES:
          this.onUploadChanges();
          break;

        case Const.ID_RECONCILE_ACCOUNT:
          this.onReconcileAccount();
          break;

        case Const.ID_CANCEL_RECONCILE:
          this.onCancelReconcile();
          break;

        case Const.ID_CHECK_FOR_UPDATES:
          this.onCheckForUpdates();
          break;

        case Const.ID_TOGGLE_SIDEBAR:
          //this.onToggleSidebar();
          break;

        case Const.ID_SET_MINT_AUTH:
          Mint.Session.showManualMintAuth();
          break;

        default:
          if (Debug.enabled) Debug.log("Unknown menu id: " + id);
          break;
      }
    },
    
    //-----------------------------------------------------------------------------
    onSyncAllWithMint: function()
    {
      this.syncWithMint(true, true, false, true, true, true);
    },
    
    //-----------------------------------------------------------------------------
    onImportTxnData: function()
    {
      this.syncWithMint(false, false, true, true, false, false);
    },

    //-----------------------------------------------------------------------------
    onImportAccountData: function()
    {
      this.syncWithMint(false, false, false, false, true, false);
    },

    //-----------------------------------------------------------------------------
    onUploadChanges: function()
    {
      this.syncWithMint(false, true, false, false, false, false);
    },

    //-----------------------------------------------------------------------------
    onReconcileAccount: function()
    {
      Reconcile.startReconcile();
      Sheets.About.turnOffAuthMsg();
    },

    //-----------------------------------------------------------------------------
    onCancelReconcile: function()
    {
      Reconcile.cancelReconcile();
      // Refresh the Mojito menu
      Ui.Menu.setupMojitoMenu(false);
      Sheets.About.turnOffAuthMsg();
    },

    //-----------------------------------------------------------------------------
    onCheckForUpdates: function()
    {
      // TODO: Open web page showing release history
    },

    //-----------------------------------------------------------------------------
    onToggleSidebar: function() {
      // Experimenting with this...
      const htmloutput = HtmlService.createTemplateFromFile('sidebar.html').evaluate()
                        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                        .setTitle('Mojito sidebar')
                        .setWidth(300);
      SpreadsheetApp.getUi().showSidebar(htmloutput);
    },
    
    //-----------------------------------------------------------------------------
    onDisplayLogWindow: function()
    {
      Debug.displayLogWindow();
    },
    
    /////////////////////////////////////////////////////////////////////////////
    // Helpers

    syncWithMint: function(syncAll, saveEdits, promptSaveEdits, importTxns, importAccountBal, importCatsTags)
    {
      // If we're running int "demo mode" then don't do anything
      if (Utils.checkDemoMode()) {
        return;
      }

      // Apply Mojito updates, if any. If the spreadsheet version cannot be updated
      // to this MojitoLib version, then abort.
      const version = Upgrade.autoUpgradeMojitoIfApplicable();
      if (version !== Const.CURRENT_MOJITO_VERSION) {
        // Upgrade function already displayed a message. Just return.
        return;
      }

      // Before doing anything, make sure the mint session is still valid.
      // If expired (no cookies), then show message box and abort.
      const cookies = Mint.Session.getCookies(false);
      if (!cookies) {
        Browser.msgBox('Mint authentication expired', 'The Mint authentication token has not been provided or it has expired. Please re-enter the Mint authentication headers then retry the operation.', Browser.Buttons.OK);
        Debug.log('No auth cookies found. Showing mint auth ui.');
        Mint.Session.showManualMintAuth();
        return;
      }

      let saveFailed = false;

      try
      {
        let txnDateRange = null;
        let acctDateRange = null;

        if (syncAll) {
          importAccountBal = !Sheets.AccountData.isUpToDate();
          // If account balances aren't up to date, get the date range for the latest balances
          acctDateRange = (importAccountBal ? Sheets.AccountData.determineImportDateRange() : null);

          // Get date range for latest transactions
          txnDateRange = Sheets.TxnData.determineImportDateRange(Utils.getMintLoginAccount());
        }

        // Wait for mint to get data from financial institutions
        //Disabled: Mint doesn't seem to support this any more
        const isMintDataReady = Mint.waitForMintFiRefresh(true);

        if (!isMintDataReady) {
          return;
        }

        if (importAccountBal) {
          // Don't activate the AccountData sheet during import. It just slows it down.
          //Sheets.AccountData.getSheet().activate();

          if (acctDateRange != null)
          {
            // Download latest account balances
            if (Debug.enabled) Debug.log("Sync account date range: " + Utilities.formatDate(acctDateRange.startDate, "GMT", "MM/dd/yyyy") + " - " + Utilities.formatDate(acctDateRange.endDate, "GMT", "MM/dd/yyyy"));

            const args = {
                startDate: acctDateRange.startDate.getTime(),
                endDate: acctDateRange.endDate.getTime(),
                replaceExistingData: false
              };

            Ui.AccountBalanceImportWindow.onImport(args);
          } else {
            // No date range specified. Show the account import window
            const args = Sheets.AccountData.determineImportDateRange();
            Ui.AccountBalanceImportWindow.show(args);
         }
        }
        
        if (saveEdits === true || promptSaveEdits === true) {
          const mintAccount = Utils.getMintLoginAccount();

          if (promptSaveEdits === true) {
            const pendingUpdates = Sheets.TxnData.getModifiedTransactionRows(mintAccount);
            if (pendingUpdates != null && pendingUpdates.length > 0) {
              if ("yes" === Browser.msgBox("", "Would you like to save your modified transactions first? (If you click \"No\", your changes may be overwritten.)", Browser.Buttons.YES_NO)) {
                saveEdits = true;
              }
            }
          }

          if (saveEdits === true) {
            Sheets.TxnData.getSheet().activate();

            // Upload any edited txns
            const success = Sheets.TxnData.saveModifiedTransactions(mintAccount, true);
            saveFailed = !success;
          }
        }

        if (importTxns) {
          if (saveFailed) {
            toast("Not all changes were saved. Skipping transaction import.");
            Utilities.sleep(3000);

          }
          else {
            Sheets.TxnData.getSheet().activate();
  
            // Import the "latest" txns, or show import window?
            if (txnDateRange != null) {
              // Download the latest txns
              if (Debug.enabled) Debug.log("Sync txn date range: " + Utilities.formatDate(txnDateRange.startDate, "GMT", "MM/dd/yyyy") + " - " + Utilities.formatDate(txnDateRange.endDate, "GMT", "MM/dd/yyyy"));
  
              const args = {
                  startDate: txnDateRange.startDate.getTime(),
                  endDate: txnDateRange.endDate.getTime(),
                  replaceExistingData: false,
                };
  
              Mint.TxnData.importTransactions(args, true, !syncAll);
  
            }
            else {
              // No date range specified. Show the txn import window
              
              // Default date range will be year-to-date
              const today = new Date;
              const startDate = new Date(today.getYear(), 0, 1);

              const args = { startDate: startDate, endDate: today };
              Ui.TxnImportWindow.show(args);
            }
          }
        }

        if (importCatsTags) {
          // Download latest categories and tags
          Debug.log("Syncing categories");
          Mint.Categories.import(false);
          Debug.log("Syncing tags");
          Mint.Tags.import(false);
        }

        Sheets.About.turnOffAuthMsg();
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e);
      }
    },


  }, // Menu

  ///////////////////////////////////////////////////////////////////////////////
  // LoginWindow class

  /**
   * @obsolete The login window no longer works because Mint now requires two-factor
   * authentication (via emailed code) if it doesn't recognize the computer where
   * the login request is coming from (which is a Google server in this case).
   */
  LoginWindow: {

    login: function()
    {
      // Only allow one login call at a time
      let loginMutex = Utils.getDocumentLock();
      if (!loginMutex.tryLock(1000))
      {
        toast("Multiple logins are occurring at once. Please wait a moment ...", "", 10);

        if (!loginMutex.tryLock(Const.MINT_LOGIN_START_TIMEOUT_SEC * 1000))
        {
          if (Debug.enabled) Debug.log("Unable to aquire login lock");
          return null;
        }
      }

      try
      {
        Debug.log("Login required");
        Ui.LoginWindow.show();
        
        // The code that follows only exists because we want to open the login window and wait for it
        // to close, either because the user successfully logged in, or the login was cancelled. This sort of
        // "modal dialog" behavior is not supported by Google Apps Script for windows created by container-bound
        // scripts. To complicate matters, the user could click the little "X" in the upper right corner to
        // close the window. There is no way (that I have found) to intercept this action and notify this login
        // script that the login has been canceled. So to handle this case, we have the login window send a 
        // "window ping" event every few seconds so we know it is still open. If we don't see the ping 
        // for 10 seconds, then we assume the user closed the window with the "X". The fact that any of this
        //  code needs to exist is pretty lame. Google Apps Script should support this simple use case.
        let loginFinished = false;
        let loginSucceeded = false;
        let loginWaitEvents = [Const.EVT_MINT_LOGIN_SUCCEEDED, Const.EVT_MINT_LOGIN_FAILED, Const.EVT_MINT_LOGIN_CANCELED, Const.EVT_MINT_LOGIN_WINDOW_PING];
        let timeoutSec = Const.MINT_LOGIN_TIMEOUT_SEC;
        let windowOpened = false;
        let timeoutCount = 0;
        
        while (true) {
          let loginEvent = EventServiceX.waitForEvents(loginWaitEvents, timeoutSec);
          switch (loginEvent) {
              
            case Const.EVT_MINT_LOGIN_SUCCEEDED:
              if (Debug.enabled) Debug.log("Wait event: Login succeeded");
              loginFinished = true;
              loginSucceeded = true;
              break;
              
            case Const.EVT_MINT_LOGIN_CANCELED:
              if (Debug.enabled) Debug.log("Wait event: Login was canceled");
              loginFinished = true;
              loginSucceeded = false;
              break;
              
            case Const.EVT_MINT_LOGIN_FAILED:
              if (Debug.enabled) Debug.log("Wait event: Login failed.");
              break;
              
            case Const.EVT_MINT_LOGIN_WINDOW_PING:
              // Login window is still open, waiting for user to click OK or Cancel. Keep waiting ...
              windowOpened = true;
              if (Debug.traceEnabled) Debug.trace("Wait event: Login window ping");
              break;
              
            default:
              if (windowOpened || timeoutCount > 2)
              {
                if (Debug.enabled) Debug.log("Login timeout: Assuming login window has been closed.");
                toast("Login timed out.", "Mint login", 5);
                loginFinished = true;
                loginSucceeded = false;
              } else {
                ++timeoutCount;
                // don't give up until the window has opened. Could just be a really slow network connection.
                if (Debug.enabled) Debug.log("Login timeout: Ignoring. Window is not open yet.");
              }
              break;
          }
          
          if (loginFinished)
            break;
        }
      }
      finally
      {
        loginMutex.releaseLock();
      }

      return loginSucceeded;
    },

    show: function()
    {
      EventServiceX.clearEvent(Const.EVT_MINT_LOGIN_STARTED);
      EventServiceX.clearEvent(Const.EVT_MINT_LOGIN_CANCELED);

      const htmlOutput = HtmlService.createTemplateFromFile('mint_login.html').evaluate();
      htmlOutput.setHeight(150).setWidth(310).setSandboxMode(HtmlService.SandboxMode.IFRAME);
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Log in to Mint');
    },

    onDoLogin: function(args) {
      let success = false;

      const loginCookies = Mint.Session.loginMintUser(args.email, args.password);
      
      if (loginCookies)
      {
        // Save successful email to settings
        Settings.setSetting(Const.IDX_SETTING_MINT_LOGIN, args.email);

        EventServiceX.triggerEvent(Const.EVT_MINT_LOGIN_SUCCEEDED, {result:"success", cookies: loginCookies});
        success = true;
      }
      else
      {
        EventServiceX.triggerEvent(Const.EVT_MINT_LOGIN_FAILED, null);
      }

      return success;
    },

    onCancel: function() {
      EventServiceX.triggerEvent(Const.EVT_MINT_LOGIN_CANCELED, null);
    },
    
    onWindowPing: function() {
      EventServiceX.triggerEvent(Const.EVT_MINT_LOGIN_WINDOW_PING, null);
    },
  },
  
  ///////////////////////////////////////////////////////////////////////////////
  TxnImportWindow: {
    /**
     *
     * @param dates {{ startDate: Date, endDate: Date }}
     */
    show: function(dates) {
      try
      {
        const args = { startDate: dates.startDate.getTime(), endDate: dates.endDate.getTime() };
        Utils.getPrivateCache().put(Const.CACHE_TXN_IMPORT_WINDOW_ARGS, JSON.stringify(args), 60);

        const htmlOutput = HtmlService.createTemplateFromFile('txn_import.html').evaluate();
        htmlOutput.setTitle("Import Transactions from Mint").setHeight(190).setWidth(250).setSandboxMode(HtmlService.SandboxMode.IFRAME);
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss != null) ss.show(htmlOutput);
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e);
      }
    },

    /**
     * onImport - Called from txn_import.html
     * args = { startDate: <date>, endDate: <date>, replaceExistingData: <true/false> }
     */
    onImport: function(args) {
      try
      {
        Mint.TxnData.importTransactions(args, true, true);
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e);
      }

      toast("Done");
    }

  },

  ///////////////////////////////////////////////////////////////////////////////
  AccountBalanceImportWindow: {
    show: function(dates) {
      // dates = { startDate: <date>, endDate: <date> }
      try
      {
        const args = { startDate: dates.startDate.getTime(), endDate: dates.endDate.getTime() };
        Utils.getPrivateCache().put(Const.CACHE_ACCOUNT_IMPORT_WINDOW_ARGS, JSON.stringify(args), 60);

        const htmlOutput = HtmlService.createTemplateFromFile('account_balance_import.html').evaluate();
        htmlOutput.setTitle("Import account balances").setHeight(220).setWidth(250).setSandboxMode(HtmlService.SandboxMode.IFRAME);
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss) ss.show(htmlOutput);
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e);
      }
    },

    /**
     * onImport - Called from account_balance_import.html
     * @param args {{
     *   startDate: Date,
     *   endDate: Date,
     *   replaceExistingData: boolean,
     *   saveReplaceExisting: boolean,
     *   importTodaysBalance: undefined|boolean,
     *   }}
     */
    onImport: function(args) {
      try
      {
        if (Debug.enabled) Debug.log("AccountBalanceImportWindow.onImport(): %s", args.toSource());

        const cookies = Mint.Session.getCookies();
        const acctInfoArray = Mint.AccountData.downloadAccountInfo();

        if (args.saveReplaceExisting) {
          Settings.setSetting(Const.IDX_SETTING_REPLACE_ALL_ON_ACCT_IMPORT, args.replaceExistingData);
        }

        // Clear the existing account data, if requested
        if (args.replaceExistingData === true) {
          const accountRanges = Sheets.AccountData.getAccountDataRanges();
          const balRange = accountRanges.balanceRange;
          const balCount = (balRange != null ? balRange.getNumRows() : 0);

          if (balCount > 0) {
            const button = Browser.msgBox("Replace existing balances?", "Are  you sure you want to REPLACE the " + balCount + " existing account balance(s)?", Browser.Buttons.OK_CANCEL);
            if (button === "cancel")
              return;
          }

          if (accountRanges.hdrRange != null) {
            accountRanges.hdrRange.clear();
            accountRanges.hdrRange.setWrap(true);
          }
          const range = (balRange != null ? balRange.offset(0, -1, balRange.getNumRows(), balRange.getNumColumns() + 1) : null);
          if (range) {
            range.clear();
          }
        }

        // Activate the last date cell so user can see the balances as they are imported.
        let balRange = Utils.getAccountDataRanges(false).balanceRange;
        let lastDateCell = (balRange != null ? balRange.offset(balRange.getNumRows() - 1, -1, 1, 1) : null);
        if (lastDateCell) {
          // Don't activate the AccountData sheet during import. It just slows it down.
          //lastDateCell.activate();
        }

        const startDate = new Date(args.startDate);
        const endDate = new Date(args.endDate);
        const acctCount = acctInfoArray.length;
        let accountsWithNoHistory = [];

        toast(Utilities.formatString("Retrieving balances for %d account(s)", acctCount), "Account balance import", 120);
        let showToastForEachAcct = false;

        for (let i = 0; i < acctCount; ++i) {
            const timeStart = Date.now();

            const currAcct = acctInfoArray[i];
            if (currAcct.isHidden)
            {
              if (Debug.enabled) Debug.log("Not retrieving balances for hidden account '%s'", currAcct.name);
              continue;
            }
            if (currAcct.isClosed)
            {
              if (Debug.enabled) Debug.log("Not retrieving balances for closed account '%s'", currAcct.name);
              continue;
            }

            if (showToastForEachAcct) {
              toast(Utilities.formatString("Retrieving balances for account %d of %d:  %s", i + 1, acctCount, currAcct.name, 60), "Account balance import", 60);
            }

            const acctWithBalances = Mint.AccountData.downloadBalanceHistory(cookies, currAcct, startDate, endDate, args.importTodaysBalance, true);
            if (acctWithBalances.balanceHistoryNotAvailable) {
              accountsWithNoHistory.push(acctWithBalances.name);
            }

            Sheets.AccountData.insertAccountBalanceHistory(acctWithBalances);

            const timeElapsed = Date.now() - timeStart;
            if (timeElapsed > 3000) {
              showToastForEachAcct = true;
            }
        }

        // Sort the the account balances by date, ascending
        balRange = Sheets.AccountData.getAccountDataRanges(true).balanceRange;
        const range = (balRange != null ? balRange.offset(0, -1, balRange.getNumRows(), balRange.getNumColumns() + 1) : null);
        if (range != null) {
          range.sort(1);

          // Activate the last date cell so user can quickly see the latest balances
          lastDateCell = range.offset(balRange.getNumRows() - 1, 0, 1, 1);
          lastDateCell.activate();
        }

        if (accountsWithNoHistory.length > 0) {
          let msg = `No balance history was found for the following ${accountsWithNoHistory.length} account(s). Only today\'s balance was imported. -- ${accountsWithNoHistory.join(' --\r\n ')}`;
          Browser.msgBox("Accounts with no balance history", msg, Browser.Buttons.OK);
        }
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e);
      }

      toast("Done", "Account balance import");
    }

  }

};
