'use strict';
/*
 * Author: b3devs@gmail.com
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
      var ss = SpreadsheetApp.getActiveSpreadsheet();

      // Add a custom "Mojito" menu to the active spreadsheet.
      var entries = [];

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
      var htmloutput = HtmlService.createTemplateFromFile('sidebar.html').evaluate()
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
      if (Utils.checkDemoMode())
        return;

      var saveFailed = false;
      var mojitoVersionCheck = false;

      try
      {
        // Prompt for login, if necessary
        var cookies = Mint.Session.getCookies();
        if (cookies == null)
          return;

        var txnDateRange = null;
        var acctDateRange = null;

        if (syncAll) {
          importAccountBal = !Sheets.AccountData.isUpToDate();
          // If account balances aren't up to date, get the date range for the latest balances
          acctDateRange = (importAccountBal ? Sheets.AccountData.determineImportDateRange() : null);

          // Get date range for latest transactions
          txnDateRange = Sheets.TxnData.determineImportDateRange(Utils.getMintLoginAccount());
        }

        mojitoVersionCheck = true; // At this point, we'll go ahead and check for Mojito updates at the end.

        // Wait for mint to get data from financial institutions
        //Disabled: Mint doesn't seem to support this any more
        var isMintDataReady = Mint.waitForMintFiRefresh(true);

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

            var args = {
                startDate: acctDateRange.startDate.getTime(),
                endDate: acctDateRange.endDate.getTime(),
                replaceExistingData: false,
              };

            Ui.AccountBalanceImportWindow.onImport(args);
          } else {
            // No date range specified. Show the account import window
            var args = Sheets.AccountData.determineImportDateRange();
            Ui.AccountBalanceImportWindow.show(args);

            mojitoVersionCheck = false; // Don't check Mojito version when we are displaying a window
         }
        }
        
        if (saveEdits === true || promptSaveEdits === true) {
          var mintAccount = Utils.getMintLoginAccount();

          if (promptSaveEdits === true) {
            var pendingUpdates = Sheets.TxnData.getModifiedTransactionRows(mintAccount);
            if (pendingUpdates != null && pendingUpdates.length > 0) {
              if ("yes" === Browser.msgBox("", "Would you like to save your modified transactions first? (If you click \"No\", your changes may be overwritten.)", Browser.Buttons.YES_NO)) {
                saveEdits = true;
              }
            }
          }

          if (saveEdits === true) {
            Sheets.TxnData.getSheet().activate();

            // Upload any edited txns
            var success = Sheets.TxnData.saveModifiedTransactions(mintAccount, true);
            saveFailed = !success;
          }
        }

        if (importTxns) {
          if (saveFailed) {
            toast("Not all changes were saved. Skipping transaction import.");
            Utilities.sleep(3000);

          } else {
            Sheets.TxnData.getSheet().activate();
  
            // Import the "latest" txns, or show import window?
            if (txnDateRange != null) {
              // Download the latest txns
              if (Debug.enabled) Debug.log("Sync txn date range: " + Utilities.formatDate(txnDateRange.startDate, "GMT", "MM/dd/yyyy") + " - " + Utilities.formatDate(txnDateRange.endDate, "GMT", "MM/dd/yyyy"));
  
              var args = {
                  startDate: txnDateRange.startDate.getTime(),
                  endDate: txnDateRange.endDate.getTime(),
                  replaceExistingData: false,
                };
  
              Mint.TxnData.importTransactions(args, true, !syncAll);
  
            } else {
              // No date range specified. Show the txn import window
              
              // Default date range will be year-to-date
              var today = new Date;
              var startDate = new Date(today.getYear(), 0, 1);
              var endDate = today;
              
              var args = { startDate: startDate, endDate: endDate };
              Ui.TxnImportWindow.show(args);

              mojitoVersionCheck = false; // Don't check Mojito version when we are displaying a window
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
      finally
      {
        // Check for Mojito updates
        if (mojitoVersionCheck === true) {
          Upgrade.autoUpgradeMojitoIfApplicable();
        }
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
      var loginMutex = Utils.getDocumentLock();
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
        var loginFinished = false;
        var loginSucceeded = false;
        var loginWaitEvents = [Const.EVT_MINT_LOGIN_SUCCEEDED, Const.EVT_MINT_LOGIN_FAILED, Const.EVT_MINT_LOGIN_CANCELED, Const.EVT_MINT_LOGIN_WINDOW_PING];
        var timeoutSec = Const.MINT_LOGIN_TIMEOUT_SEC;
        var windowOpened = false;
        var timeoutCount = 0;
        
        while (true) {
          var loginEvent = EventServiceX.waitForEvents(loginWaitEvents, timeoutSec);
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

      var htmlOutput = HtmlService.createTemplateFromFile('mint_login.html').evaluate();
      htmlOutput.setHeight(150).setWidth(310).setSandboxMode(HtmlService.SandboxMode.IFRAME);
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Log in to Mint');
    },

    onDoLogin: function(args) {
      var success = false;

      var loginCookies = Mint.Session.loginMintUser(args.email, args.password);
      
      if (loginCookies != null && loginCookies != "")
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

    show: function(dates)
    {
      // dates = { startDate: <date>, endDate: <date> }
      var uiApp = UiApp.createApplication().setWidth(250).setHeight(230).setTitle("Import transactions from Mint");
      
      var startDateField = uiApp.createDateBox().setName('startDate').setId("start_date");
      var endDateField = uiApp.createDateBox().setName('endDate').setId("end_date");
      
      var grid = uiApp.createGrid(2, 2);
      grid.setWidget(0, 0, uiApp.createLabel('Start date:'));
      grid.setWidget(0, 1, startDateField);
      grid.setWidget(1, 0, uiApp.createLabel('End date:'));
      grid.setWidget(1, 1, endDateField);
      
      var checkboxReplaceData = uiApp.createCheckBox("Replace existing transactions").setName('replaceData').setValue(false).setHeight(30);
      
      var btnOk = uiApp.createButton('OK').setId("ok_button").setHeight(30).setWidth(75);
      btnOk.addClickHandler(uiApp.createClientHandler().forTargets(btnOk).setEnabled(false));
      btnOk.addClickHandler(uiApp.createServerHandler('TxnImportWindow_onOkClicked').addCallbackElement(grid).addCallbackElement(checkboxReplaceData));
      var btnCancel = uiApp.createButton('Cancel').setHeight(30).setWidth(75);
      btnCancel.addClickHandler(uiApp.createServerHandler('TxnImportWindow_onCancelClicked'));
      
      var spacerPanel = uiApp.createVerticalPanel().setHeight(30).add(uiApp.createLabel());
      var vPanel = uiApp.createVerticalPanel();
      vPanel.add(grid).add(spacerPanel).add(checkboxReplaceData);
      var buttonPanel = uiApp.createHorizontalPanel().setHeight(100).setWidth("100%");
      buttonPanel.setVerticalAlignment(UiApp.VerticalAlignment.BOTTOM).setHorizontalAlignment(UiApp.HorizontalAlignment.CENTER);
      buttonPanel.add(btnOk).add(btnCancel);
      
      uiApp.add(vPanel).add(buttonPanel);
      
      // Pre-populate some fields
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      startDateField.setValue(dates.startDate);
      startDateField.setFocus(true);
      endDateField.setValue(dates.endDate);
      checkboxReplaceData.setValue(Settings.getSetting(Const.IDX_SETTING_REPLACE_ALL_ON_TXN_IMPORT) === true);

      // Show the window
      spreadsheet.show(uiApp);
    },

    // Html version not being used. Too slow to load jquery.
    show_html: function(dates) {
      // dates = { startDate: <date>, endDate: <date> }
      try
      {
        var args = { startDate: dates.startDate.getTime(), endDate: dates.endDate.getTime() };
        Utils.getPrivateCache().put(Const.CACHE_TXN_IMPORT_WINDOW_ARGS, JSON.stringify(args), 60);

        var htmlOutput = HtmlService.createTemplateFromFile('txn_import.html').evaluate();
        htmlOutput.setTitle("Import Transactions from Mint").setHeight(235).setWidth(250).setSandboxMode(HtmlService.SandboxMode.IFRAME);
        var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    },

    //-------------------------------------------------------------------------
    // Event handlers

    onOkClicked: function(e)
    {
      var uiApp = UiApp.getActiveApplication();
      var startDate = e.parameter.startDate;
      var endDate = e.parameter.endDate;
      var validInputs = true;

      if (!(startDate instanceof Date)) {
        toast("Invalid start date");
        validInputs = false;
      }
      if (!(endDate instanceof Date)) {
        toast("Invalid end date");
        validInputs = false;
      }

      if (!validInputs) {
        uiApp.getElementById("ok_button").setEnabled(true);
        return uiApp;
      }

      var replaceData = (e.parameter.replaceData == "true");
      if (replaceData !== Settings.getSetting(Const.IDX_SETTING_REPLACE_ALL_ON_TXN_IMPORT)) {
        if ("yes" === Browser.msgBox("Do you want this \"Replace existing transactions\" choice to be the default for future imports?", Browser.Buttons.YES_NO)) {
          Settings.setSetting(Const.IDX_SETTING_REPLACE_ALL_ON_TXN_IMPORT, replaceData);
        }
      }

      var args = {
        startDate: startDate.getTime(),
        endDate: endDate.getTime(),
        replaceExistingData: replaceData,
      };

      uiApp.close();

      Ui.TxnImportWindow.onImport(args);

      return uiApp;
    },

    onCancelClicked: function(e)
    {
      var uiApp = UiApp.getActiveApplication();
      uiApp.close();
      return uiApp;
    }
  },

  ///////////////////////////////////////////////////////////////////////////////
  AccountBalanceImportWindow: {

    show: function(dates)
    {
      // dates = { startDate: <date>, endDate: <date> }
      var uiApp = UiApp.createApplication().setWidth(250).setHeight(230).setTitle("Import account balances");
      
      var startDateField = uiApp.createDateBox().setName('startDate').setId("start_date");
      var endDateField = uiApp.createDateBox().setName('endDate').setId("end_date");
      
      var grid = uiApp.createGrid(2, 2);
      grid.setWidget(0, 0, uiApp.createLabel('Start date:'));
      grid.setWidget(0, 1, startDateField);
      grid.setWidget(1, 0, uiApp.createLabel('End date:'));
      grid.setWidget(1, 1, endDateField);
      
      var checkboxReplaceData = uiApp.createCheckBox("Replace existing account data").setName('replaceData').setValue(false).setHeight(30);
      var checkboxImportCurrentDay = uiApp.createCheckBox("Include today's balance for accounts with no available history").setName('includeToday').setHeight(30);
      var vPanel = uiApp.createVerticalPanel();
      vPanel.add(grid);

      let okServerClickHandler = uiApp.createServerHandler('AccountBalanceImportWindow_onOkClicked')
                                      .addCallbackElement(grid)
                                      .addCallbackElement(checkboxReplaceData);


      var includeTodaySetting = Settings.getInternalSetting(Const.IDX_INT_SETTING_CURR_DAY_ACCT_IMPORT);
      var showIncludeTodayCheckbox = (includeTodaySetting !== "");

      if (showIncludeTodayCheckbox) {
        // Add "include today's balance" checkbox to UI
        vPanel.add(uiApp.createLabel().setHeight(10)); // spacer
        checkboxImportCurrentDay.setValue(includeTodaySetting);
        vPanel.add(checkboxImportCurrentDay);
        vPanel.add(uiApp.createLabel().setHeight(10)); // spacer
        vPanel.add(checkboxReplaceData);
        vPanel.add(uiApp.createLabel().setHeight(50)); // spacer
        
        // Add callback element to click handler
        okServerClickHandler.addCallbackElement(checkboxImportCurrentDay);
      } else {
        vPanel.add(uiApp.createLabel().setHeight(30)); // spacer
        vPanel.add(checkboxReplaceData);
        vPanel.add(uiApp.createLabel().setHeight(70)); // spacer
      }

      var btnOk = uiApp.createButton('OK').setId("ok_button").setHeight(30).setWidth(75);
      var btnCancel = uiApp.createButton('Cancel').setHeight(30).setWidth(75);
      var buttonPanel = uiApp.createHorizontalPanel().setHeight(30).setWidth("100%");
      buttonPanel.setVerticalAlignment(UiApp.VerticalAlignment.BOTTOM).setHorizontalAlignment(UiApp.HorizontalAlignment.CENTER);
      buttonPanel.add(btnOk).add(btnCancel);

      uiApp.add(vPanel).add(buttonPanel);

      btnOk.addClickHandler(okServerClickHandler);
      btnOk.addClickHandler(uiApp.createClientHandler().forTargets(btnOk).setEnabled(false));
      btnCancel.addClickHandler(uiApp.createServerHandler('AccountBalanceImportWindow_onCancelClicked'));

      // Pre-populate some fields
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      startDateField.setValue(dates.startDate);
      startDateField.setFocus(true);
      endDateField.setValue(dates.endDate);
      checkboxReplaceData.setValue(Settings.getSetting(Const.IDX_SETTING_REPLACE_ALL_ON_ACCT_IMPORT) === true);

      // Show the window
      spreadsheet.show(uiApp);
    },

    // Html version not being used. Too slow to load jquery.
    show_html: function(dates) {
      // dates = { startDate: <date>, endDate: <date> }
      try
      {
        var args = { startDate: dates.startDate.getTime(), endDate: dates.endDate.getTime() };
        Utils.getPrivateCache().put(Const.CACHE_ACCOUNT_IMPORT_WINDOW_ARGS, JSON.stringify(args), 60);

        var htmlOutput = HtmlService.createTemplateFromFile('account_balance_import.html').evaluate();
        htmlOutput.setTitle("Import account balances").setHeight(235).setWidth(250).setSandboxMode(HtmlService.SandboxMode.IFRAME);
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss != null) ss.show(htmlOutput);
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e);
      }
    },

    /**
     * onImport - Called from account_balance_import.html
     * args = {
     *   startDate: <date>,
     *   endDate: <date>,
     *   replaceExistingData: true/false,
     *   importTodaysBalance: undefined/true/false,
     * }
     */
    onImport: function(args) {
      try
      {
        if (Debug.enabled) Debug.log("AccountBalanceImportWindow.onImport(): %s", args.toSource());
        var cookies = Mint.Session.getCookies();
        var acctInfoArray = Mint.AccountData.downloadAccountInfo();

        // Clear the existing account data, if requested
        if (args.replaceExistingData === true) {
          var accountRanges = Sheets.AccountData.getAccountDataRanges();
          var balRange = accountRanges.balanceRange;
          var balCount = (balRange != null ? balRange.getNumRows() : 0);

          if (balCount > 0) {
            var button = Browser.msgBox("Replace existing balances?", "Are  you sure you want to REPLACE the " + balCount + " existing account balance(s)?", Browser.Buttons.OK_CANCEL);
            if (button === "cancel")
              return;
          }

          if (accountRanges.hdrRange != null) {
            accountRanges.hdrRange.clear();
            accountRanges.hdrRange.setWrap(true);
          }
          var range = (balRange != null ? balRange.offset(0, -1, balRange.getNumRows(), balRange.getNumColumns() + 1) : null);
          if (range != null) {
            range.clear();
          }
        }

        // Activate the last date cell so user can see the balances as they are imported.
        var balRange = Utils.getAccountDataRanges(false).balanceRange;
        var lastDateCell = (balRange != null ? balRange.offset(balRange.getNumRows() - 1, -1, 1, 1) : null);
        if (lastDateCell != null) {
          // Don't activate the AccountData sheet during import. It just slows it down.
          //lastDateCell.activate();
        }

        var startDate = new Date(args.startDate);
        var endDate = new Date(args.endDate);
        var acctCount = acctInfoArray.length;

        toast(Utilities.formatString("Retrieving balances for %d account(s)", acctCount), "Account balance import", 120);
        var showToastForEachAcct = false;

        for (var i = 0; i < acctCount; ++i) {
            var timeStart = Date.now();

            var currAcct = acctInfoArray[i];
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

            var acctWithBalances = Mint.AccountData.downloadBalanceHistory(cookies, currAcct, startDate, endDate, args.importTodaysBalance, true);

            Sheets.AccountData.insertAccountBalanceHistory(acctWithBalances);

            var timeElapsed = Date.now() - timeStart;
            if (timeElapsed > 3000) {
              showToastForEachAcct = true;
            }
        }

        // Sort the the account balances by date, ascending
        balRange = Sheets.AccountData.getAccountDataRanges(true).balanceRange;
        var range = (balRange != null ? balRange.offset(0, -1, balRange.getNumRows(), balRange.getNumColumns() + 1) : null);
        if (range != null) {
          range.sort(1);

          // Activate the last date cell so user can quickly see the latest balances
          lastDateCell = range.offset(balRange.getNumRows() - 1, 0, 1, 1);
          lastDateCell.activate();
        }
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(e);
      }

      toast("Done", "Account balance import");
    },

    //-----------------------------------------------------------------------------
    // Event handlers

    onOkClicked: function(e)
    {
      var uiApp = UiApp.getActiveApplication();
      var startDate = e.parameter.startDate;
      var endDate = e.parameter.endDate;
      var validInputs = true;

      if (!(startDate instanceof Date)) {
        toast("Invalid start date");
        validInputs = false;
      }
      if (!(endDate instanceof Date)) {
        toast("Invalid end date");
        validInputs = false;
      }

      if (!validInputs) {
        uiApp.getElementById("ok_button").setEnabled(true);
        return uiApp;
      }

      var replaceData = (e.parameter.replaceData == "true");
      if (replaceData !== Settings.getSetting(Const.IDX_SETTING_REPLACE_ALL_ON_ACCT_IMPORT)) {
        if ("yes" === Browser.msgBox("Do you want this \"Replace existing account balances\" choice to be the default for future imports?", Browser.Buttons.YES_NO)) {
          Settings.setSetting(Const.IDX_SETTING_REPLACE_ALL_ON_ACCT_IMPORT, replaceData);
        }
      }

      // Convert start date to midnight and end date to midnight + 1 (so the difference isn't 0 if they are the same date)
      startDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
      endDate = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate(), 0, 0, 1);

      var args = {
        startDate: startDate.getTime(),
        endDate: endDate.getTime(),
        replaceExistingData: replaceData,
      };

      //Debug.log("e.parameter.replaceData = " + e.parameter.replaceData);
      //Debug.log("e.parameter.includeToday = " + e.parameter.includeToday);
      if (e.parameter.includeToday != undefined) {
        args.importTodaysBalance = (e.parameter.includeToday == "true");
        Debug.log("Import today's balance for accounts with no balance history: %s", args.importTodaysBalance);
      } else {
        Debug.log("includeToday is undefined");
      }

      uiApp.close();

      Ui.AccountBalanceImportWindow.onImport(args);

      return uiApp;
    },

    onCancelClicked: function(e)
    {
      var uiApp = UiApp.getActiveApplication();
      uiApp.close();
      return uiApp;
    }
  }

};
