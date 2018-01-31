'use strict';
/*
 * Author: b3devs@gmail.com
 */

import {Const} from './Constants.js';
import {Mint} from './MintApi.js';
import {Ui} from './Ui.js';
import {Reconcile} from './Reconcile.js';
import {Utils, Settings, toast} from './Utils.js';
import {SpreadsheetUtils} from './SpreadsheetUtils.js';
import {Debug} from './Debug.js';


export const Sheets = {

  getSheetType: function(sheetName) {
    var sheet = sheetName 
                ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
                : SpreadsheetApp.getActiveSheet();

    var type = null;
    if (sheet) {
      var sheetTypeCell = sheet.getRange(Const.SHEET_TYPE_CELL);
      type = sheetTypeCell.getValue();
    }
    return type;
  },

  ////////////////////////////////////////////////////////////////////////////////////////////////
  Triggers: {

    ///////////////////////////////////////////////////////////////////////////////
    onOpen: function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();

      try
      {
        var showAuthMsg = Settings.getInternalSetting(Const.IDX_INT_SETTING_SHOW_AUTH_MSG);
        Sheets.About.showAuthorizationMsg(showAuthMsg);
        
        Ui.Menu.setupMojitoMenu(true);
        
        // Set default values
        var txnActionCell = ss.getRangeByName("TxnSheetAction");
        txnActionCell.setValue(Const.TXN_ACTION_DEFAULT);
      }
      catch (e) {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox(Debug.getExceptionInfo(e));
      }
    },

    ///////////////////////////////////////////////////////////////////////////////
    onEdit: function(e) {
      try
      {
        var ss = e.source;
        var sheet = e.range.getSheet();
        var sheetTypeCell = sheet.getRange(Const.SHEET_TYPE_CELL);
        var sheetType = (sheetTypeCell !== null ? sheetTypeCell.getValue() : null);
        var sheetName = sheet.getName();
        var editRowFirst = e.range.getRow();
        var editRowLast = e.range.getLastRow();
        var editColFirst = e.range.getColumn();
        var editColLast = e.range.getLastColumn();
        var numRows = e.range.getNumRows();
        var numColumns = e.range.getNumColumns();

        // Figure out what sheet was edited. Several sheets have a hidden "sheet type" value. We'll look for that first.
        if (sheetType === Const.SHEET_TYPE_BUDGET)
        {
          if (numRows === 1 && numColumns === 1)
          {
            // Was a new date range selected from the drop-down?
            var dateRangeCell = sheet.getRange(3, 1);
            if (editRowFirst === dateRangeCell.getRow() && editColFirst === dateRangeCell.getColumn())
            {
              this.onBudgetDateRangeEdit(e);
            }
            else
            {
              var startDateCell = Sheets.Budget.getRangeByName(sheet, "BudgetStartDate");
              var endDateCell = Sheets.Budget.getRangeByName(sheet, "BudgetEndDate");
              var budgetRange = Sheets.Budget.getRangeByName(sheet, "BudgetItemsRange");

              // Was start date edited directly?
              if (editRowFirst === startDateCell.getRow() && editColFirst === startDateCell.getColumn())
              {
                dateRangeCell.setValue(Const.IDX_DATERANGE_CUSTOM);
                Sheets.Budget.updateCalculations(sheet);
              }
              // Was end date edited directly?
              else if (editRowFirst === endDateCell.getRow() && editColFirst === endDateCell.getColumn())
              {
                dateRangeCell.setValue(Const.IDX_DATERANGE_CUSTOM);
                Sheets.Budget.updateCalculations(sheet);
              }
              // Was a budget value changed?
              else if (budgetRange.getRow() <= editRowFirst && editRowLast <= budgetRange.getLastRow()
                && (editColFirst >= Const.IDX_BUDGET_NAME + 1
                    && editColFirst <= Const.IDX_BUDGET_INCLUDE_ANDOR + 1))
                {
                  Sheets.Budget.updateCalculations(sheet);
                }
            }
          }
        }
        else if (sheetType === Const.SHEET_TYPE_SAVINGS_GOALS)
        {
          if (numRows === 1 && numColumns === 1)
          {
            var goalRange = Sheets.SavingsGoal.getRangeByName(sheet, "GoalsRange");
            
            // Was a budget value changed?
            if (goalRange.getRow() <= editRowFirst && editRowLast <= goalRange.getLastRow()
              && ((editColFirst >= Const.IDX_GOAL_NAME + 1 && editColFirst <= Const.IDX_GOAL_INCLUDE_ANDOR + 1)
                  || editColFirst === Const.IDX_GOAL_CARRY_FWD + 1
                  || editColFirst === Const.IDX_GOAL_CREATE_DATE + 1))
              {
                Sheets.SavingsGoal.updateCalculations(sheet);
              }
          }
        }
        else if (sheetType === Const.SHEET_TYPE_INOUT)
        {
          if (numRows == 1 && numColumns == 1)
          {
            // Was a new date range selected from the drop-down?
            var dateRangeCell = sheet.getRange(3, 1);
            if (editRowFirst === dateRangeCell.getRow() && editColFirst === dateRangeCell.getColumn())
            {
              this.onInOutDateRangeEdit(e);
            }
            else
            {
              var startDateCell = Sheets.InOut.getRangeByName(sheet, "InOutStartDate");
              var endDateCell = Sheets.InOut.getRangeByName(sheet, "InOutEndDate");
              
              // Was start date edited directly?
              if (editRowFirst == startDateCell.getRow() && editColFirst == startDateCell.getColumn())
              {
                dateRangeCell.setValue(Const.IDX_DATERANGE_CUSTOM);
                Sheets.InOut.updateCalculations(sheet);
              }
              // Was end date edited directly?
              else if (editRowFirst == endDateCell.getRow() && editColFirst == endDateCell.getColumn())
              {
                dateRangeCell.setValue(Const.IDX_DATERANGE_CUSTOM);
                Sheets.InOut.updateCalculations(sheet);
              }
            }
          }
        }
        else if (sheetType === Const.SHEET_TYPE_TXNDATA)
        {
          var txnRange = Utils.getTxnDataRange();
          // Was a transaction edited?
          if (editRowFirst >= txnRange.getRow()
            && (editColFirst >= Const.IDX_TXN_DATE + 1 && editColFirst <= Const.IDX_TXN_LAST_COL + 1))
            {
              this.onTxnUpdate(e);
            }
          else if (numRows == 1 && numColumns == 1)
          {
            var txnActionCell = ss.getRangeByName("TxnSheetAction");
            
            if (editRowFirst == txnActionCell.getRow() && editColFirst == txnActionCell.getColumn())
            {
              this.onTxnAction(e);
            }
          }
        }
        else if (sheetType === Const.SHEET_TYPE_RECONCILE)
        {
          Reconcile.onReconcileSheetEdit(e);
        }
        else if (sheetName === Const.SHEET_NAME_SETTINGS)
        {
          if (numRows == 1 && numColumns == 1 && editColFirst == 2)
          {
            var settingsRange = ss.getRangeByName("SettingsRange");

            if (settingsRange.getRow() <= editRowFirst && editRowLast <= settingsRange.getLastRow()) {
              var rowIndex = editRowFirst - settingsRange.getRow() + 1;

              if (rowIndex === Const.IDX_SETTING_MINT_PWD) {
                var value = e.value;
                Utils.saveAccountStats(String(value));
                if (value) {
                  e.range.setValue("** Saved **");
                }
                else {
                  if (Debug.enabled) Debug.log("Saved password has been deleted");
                }
              } else if (rowIndex === Const.IDX_SETTING_CLEARED_TAG) {
                Utils.getPrivateCache().remove(Const.CACHE_SETTING_CLEARED_TAG);

              } else if (rowIndex === Const.IDX_SETTING_RECONCILED_TAG) {
                Utils.getPrivateCache().remove(Const.CACHE_SETTING_RECONCILED_TAG);

              } else if (rowIndex === Const.IDX_SETTING_TXN_AMOUNT_COL) {
                Utils.getPrivateCache().remove(Const.CACHE_TXNDATA_AMOUNT_COL);
              }
            }
          }
        }
        else if (sheetName === Const.SHEET_NAME_CATEGORYDATA)
        {
          toast("Changes to categories are not supported in Mojito and will be overwritten when you sync with Mint.", "Not supported", 8);
        }
        else if (sheetName === Const.SHEET_NAME_TAGDATA)
        {
          toast("Changes to tags are not supported in Mojito and will be overwritten when you sync with Mint.", "Not supported", 8);
        }
        
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }
      
    },
    
    //-----------------------------------------------------------------------------
    onBudgetDateRangeEdit: function(e) {
      
      var result = Utils.getStartEndDates(e.value);
      
      if (result != null && result.startDate != null && result.endDate != null) {
        var sheet = e.range.getSheet();
        Sheets.Budget.getRangeByName(sheet, "BudgetStartDate").setValue(result.startDate);
        Sheets.Budget.getRangeByName(sheet, "BudgetEndDate").setValue(result.endDate);

        Sheets.Budget.updateCalculations(sheet);
      }
    },

    //-----------------------------------------------------------------------------
    onInOutDateRangeEdit: function(e) {

      var result = Utils.getStartEndDates(e.value);

      if (result != null && result.startDate != null && result.endDate != null) {
        var sheet = e.range.getSheet();
        Sheets.InOut.getRangeByName(sheet, "InOutStartDate").setValue(result.startDate);
        Sheets.InOut.getRangeByName(sheet, "InOutEndDate").setValue(result.endDate);

        Sheets.InOut.updateCalculations(sheet);
      }
    },
    
    //-----------------------------------------------------------------------------
    onTxnAction: function(e) {

      try
      {

        switch(e.value.toLowerCase())
        {
          case Const.TXN_ACTION_SORT_BY_DATE_DESC:
            toast("Sorting transactions by date (descending)", "Action", 10);
            var txnDataRange = Utils.getTxnDataRange();
            txnDataRange.sort([{column: Const.IDX_TXN_DATE + 1, ascending: false}, {column: Const.IDX_TXN_PARENT_ID + 1, ascending: true}, Const.IDX_TXN_ACCOUNT + 1]);
            break;
            
          case Const.TXN_ACTION_SORT_BY_DATE_ASC:
            toast("Sorting transactions by date (ascending)", "Action", 10);
            var txnDataRange = Utils.getTxnDataRange();
            txnDataRange.sort([{column: Const.IDX_TXN_DATE + 1, ascending: true}, {column: Const.IDX_TXN_PARENT_ID + 1, ascending: true}, Const.IDX_TXN_ACCOUNT + 1]);
            break;
            
          case Const.TXN_ACTION_SORT_BY_MONTH_AMOUNT:
            var txnAmountCol = Utils.getTxnAmountColumn();
            var toastMsg = "Sorting txns by month / amount";
            if (txnAmountCol !== Const.IDX_TXN_AMOUNT + 1) {
              toastMsg += "  (Using Amount column " + String(txnAmountCol) + ")";
            }
            toast(toastMsg, "Action", 10);
            var txnDataRange = Utils.getTxnDataRange();
            txnDataRange.sort([{column: Const.IDX_TXN_YEAR_MONTH + 1, ascending: false}, txnAmountCol]);
            break;

          case Const.TXN_ACTION_CLEAR_TXN_MATCHES:
            toast("Clearing transaction highlights", "Action", 20);
            Sheets.TxnData.clearRowHighlights();
            break;
        }
        toast("Complete", "Action");
        
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }
      
      e.range.setValue(Const.TXN_ACTION_DEFAULT);
    },
    
    //-----------------------------------------------------------------------------
    onTxnUpdate: function(e) {
      
      var ss = e.source;
      
      var txnSheet = e.range.getSheet();
      var editRow = e.range.getRow();
      var editCol = e.range.getColumn();
      var editColLast = e.range.getLastColumn();
      
      if (editCol === Const.IDX_TXN_AMOUNT + 1) {
        Sheets.TxnData.handleAmountEdit(editRow, editCol);
        return;
      }
      
      var startRow = Math.min(e.range.getRow(), e.range.getLastRow());
      var endRow = Math.max(e.range.getRow(), e.range.getLastRow());
      for (var i = startRow; i <= endRow; ++i) {
        if (!Sheets.TxnData.isTransactionEditAllowed(txnSheet, editRow, editCol, editColLast, Const.EDITTYPE_EDIT, true)) {
          break;
        }
        
        Sheets.TxnData.validateTransactionEdit(txnSheet, i, editCol, editColLast, Const.EDITTYPE_EDIT);
      }
    },

    ///////////////////////////////////////////////////////////////////////////////
    onChange: function(e) {
      try
      {
        if (e.changeType === "INSERT_ROW" || e.changeType === "REMOVE_ROW") {
          var ss = e.source;
          var sheet = ss.getActiveSheet();
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
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }

    },
    
  }, // Triggers

  ////////////////////////////////////////////////////////////////////////////////////////////////
  About: {

    getSheet: function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(Const.SHEET_NAME_ABOUT);
      return sheet;
    },

    showAuthorizationMsg: function(show) {
      var sheet = this.getSheet();
      if (show == true) {
        sheet.showRows(6, 3);
      } else {
        sheet.hideRows(6, 3);
      }
    },
    
    turnOffAuthMsg: function() {
      var showAuthMsg = Settings.getInternalSetting(Const.IDX_INT_SETTING_SHOW_AUTH_MSG);
      if (showAuthMsg != false) {
        Settings.setInternalSetting(Const.IDX_INT_SETTING_SHOW_AUTH_MSG, false);
        this.showAuthorizationMsg(false);
      }
    },

  }, // About

  ////////////////////////////////////////////////////////////////////////////////////////////////
  TxnData: {

    getSheet: function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(Const.SHEET_NAME_TXNDATA);
      return sheet;
    },

    getColumnName: function(txnSheet, colNum) {
      return txnSheet.getRange(txnSheet.getFrozenRows(), colNum).getValue();
    },

    _txnRangeAndValues : null,  // This variable is only valid while server-side code is executing. It resets each time.
    
    getTxnRangeAndValues: function()
    {
      if (this._txnRangeAndValues == null) {
        if (Debug.enabled) Debug.log("** Fetching all txn values **");
        var txnRange = Utils.getTxnDataRange();
        this._txnRangeAndValues = {
          txnRange : txnRange,
          txnValues : txnRange.getValues(),
        };
      }
      
      return this._txnRangeAndValues;
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    determineImportDateRange: function(mintAccount)
    {
      var today = new Date();
      var startDate = null;
      // We'll use the last day of the month as the end date
      var endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);

      // For calculating the start date, we'll use the date of the first pending txn minus
      // a "fudge factor". There's no harm in importing more txns than we actually need to.
      var FUDGE_FACTOR = 14;  // We'll use a fudge factor of 14 days

      var trv = this.getTxnRangeAndValues();
      var existingValues = trv.txnValues;
      var existingLen = existingValues.length;

      if (existingLen === 0) {
        // If no transactions have been imported yet, then just use January 1st
        // as the start date.
        startDate = new Date(today.getFullYear(), 0, 1);
        
      } else {
        
        // Find the date of the first pending txn, if any
        // NOTE: There is a bug in Mint where editing the category of a pending txn will
        // cause the txn to forever be stuck as pending. This bug creates a problem for us
        // because we are looking for the first pending txn, expecting it to be only a few
        // days old, not several months old.
        // Our work-around is to ignore any pending txn whose date is 30 days before the
        // most recent txn.
        var lastTxnDate = Utils.find2dArrayMaxValue(existingValues, Const.IDX_TXN_DATE, function(row) { return (row[Const.IDX_TXN_MINT_ACCOUNT] === mintAccount); } );
        if (lastTxnDate == null) {
          lastTxnDate = new Date();
        }
        var invalidPendingTxnDate = new Date(lastTxnDate.getFullYear(), lastTxnDate.getMonth(), lastTxnDate.getDate() - 30);
        var firstPendingTxnDate = null;

        var existingLen = existingValues.length;
        for (var i = 0; i < existingLen; ++i) {
          var txnDate = existingValues[i][Const.IDX_TXN_DATE];

          if (existingValues[i][Const.IDX_TXN_MINT_ACCOUNT] === mintAccount &&
                (firstPendingTxnDate == null ||
                (existingValues[i][Const.IDX_TXN_STATE] === Const.TXN_STATUS_PENDING && txnDate < firstPendingTxnDate && txnDate > invalidPendingTxnDate)))
          {
            firstPendingTxnDate = txnDate;
          }
        }

        if (firstPendingTxnDate == null) {
          firstPendingTxnDate = lastTxnDate;
        }
        if (Debug.enabled) Debug.log("firstPendingTxnDate: " + Utilities.formatDate(firstPendingTxnDate, "GMT", "MM/dd/yyyy"));

        startDate = new Date(firstPendingTxnDate.getFullYear(), firstPendingTxnDate.getMonth(), firstPendingTxnDate.getDate() - FUDGE_FACTOR);
      }

      return { startDate : startDate, endDate : endDate };
    },
    
    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    insertData: function(txnValues, replaceExistingData)
    {
      if (txnValues == null || txnValues.length === 0)
        return;

      var numCols = Const.IDX_TXN_LAST_COL + 1;
      Debug.log("inserting " + txnValues.length + " txns");

      var txnRange = Utils.getTxnDataRange();

      // Merge new txns with existing txns
      var rowCountBefore = txnRange.getNumRows();
      var existingValues = (replaceExistingData ? [] : txnRange.getValues());

      Sheets.TxnData.mergeTxnValues(txnValues, existingValues);
      
      var range = txnRange.offset(0, 0, existingValues.length, numCols);
      range.setValues(existingValues);
      
      // If there are fewer txns after the merge, then delete the extras from the spreadsheet
      var rowCountAfter = existingValues.length;
      var rowDiff = rowCountBefore - rowCountAfter;
      if (rowDiff > 0) {
        Debug.log("Deleting " + rowDiff + " extra txn rows from TxnData sheet.");
        var delRange = txnRange.offset(rowCountAfter, 0, rowDiff, numCols);
        delRange.clear();
      }

      // Pretty-up the appearance
      this.formatSheet(range, existingValues, rowCountAfter);
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    isTransactionEditAllowed: function(txnSheet, editRowNum, updateColFirst, updateColLast, editType, interactive)
    {
      // If transaction has already been flagged as edited, make sure this new edit
      // is allowed before saving the previous one.
      var editStatusCell = txnSheet.getRange(editRowNum, Const.IDX_TXN_EDIT_STATUS + 1);
      var prevEditType = editStatusCell.getValue();
      if (prevEditType === Const.EDITTYPE_EDIT && editType === Const.EDITTYPE_NEW) {
        // This should never happen
        throw "The transaction edit status cannot be changed from \"E\" to \"N\".";
      }

      // Do not allow edits to "pending" transactions
      var pendingCell = txnSheet.getRange(editRowNum, Const.IDX_TXN_STATE + 1);
      if (pendingCell.getValue() === Const.TXN_STATUS_PENDING) {
        toast("Do NOT edit pending transactions! Mint has a bug related to editing pending transactions.", "Edit not allowed", 8);
        return false;
      }

      if (editType === Const.EDITTYPE_EDIT) {
        for (var i = updateColFirst - 1; i < updateColLast; ++i) {
          if (Const.TXN_EDITABLE_FIELDS.indexOf(i) < 0) {
            toast("You cannot update column \"" + Sheets.TxnData.getColumnName(txnSheet, i+1) + "\". Please \"undo\" your change.", "Edit not allowed", 5);
            return false;
          }
        }
      }

      return true;
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    validateTransactionEdit: function(txnSheet, editRowNum, updateColFirst, updateColLast, editType)
    {
      var isValid = false;

      try
      {
        if (Debug.traceEnabled) Debug.trace("Validating row %d edit '%s', first col %d, last col %d", editRowNum, editType, updateColFirst, updateColLast);

        // Validate the transaction and mark it as "edited" with editType
        var allTags = "";
        var tagsValidated = false;

        for (var colIndex = updateColFirst - 1; colIndex < updateColLast; ++colIndex) {

          //if (Debug.enabled) Debug.log("Validating column " + String(colIndex + 1));
          switch (colIndex)
          {
            case Const.IDX_TXN_CATEGORY:
              var categoryCell = txnSheet.getRange(editRowNum, Const.IDX_TXN_CATEGORY + 1);
              var category = categoryCell.getValue();

              var validationInfo = Mint.validateCategory(category, true);
              
              isValid = validationInfo.isValid;
              if (isValid) {
                // Set Category field to exact value in Mint (matching upper/lower case)
                if (validationInfo.displayName != category) { categoryCell.setValue(validationInfo.displayName); }
                // Set Category ID field
                var categoryIdCell = txnSheet.getRange(editRowNum, Const.IDX_TXN_CAT_ID + 1);
                categoryIdCell.setValue(validationInfo.catId);
              }
              break;

            case Const.IDX_TXN_TAGS:
            case Const.IDX_TXN_CLEAR_RECON:
              if (tagsValidated)
                break; // No point in validating tags twice

              var tagCleared = Mint.getClearedTag();
              var tagReconciled = Mint.getReconciledTag();
              
              if (colIndex == Const.IDX_TXN_CLEAR_RECON && (tagCleared == null || tagReconciled == null))
              {
                Browser.msgBox("Clear / Reconcile not enabled", "You cannot mark a transaction as cleared (c) or reconciled (R) until you enable this feature by specifying the corresponding tags on the Settings sheet. Refer to the Help sheet for instructions on how to do this.", Browser.Buttons.OK);
                isValid = false;
                break;
              }
              
              // Get the reconciled or cleared tags
              var crCell = txnSheet.getRange(editRowNum, Const.IDX_TXN_CLEAR_RECON + 1);
              var crVal = crCell.getValue();
              var reconciled = (crVal.toUpperCase() === "R");
              var cleared = (crVal != null && crVal != ""); // "Cleared" is anything other than "R" or empty
              crVal = null;
              if (reconciled) {
                // Reconciled transactions are also 'cleared', so we'll include both tags
                crVal = tagReconciled + Const.DELIM + tagCleared;
              }
              else if (cleared) {
                crVal = tagCleared;
              }
              
              if (crVal != null) {
                allTags += crVal + Const.DELIM;
              }
              
              var tagsCell = txnSheet.getRange(editRowNum, Const.IDX_TXN_TAGS + 1);
              var tags = tagsCell.getValue();
              if (tags != null) {
                allTags += tags;
              }

              var validationInfo = Mint.validateTags(allTags, true);
              
              isValid = validationInfo.isValid;
              if (isValid) {
                // Set Tags field to exact values in Mint (matching upper/lower case)
                tagsCell.setValue(validationInfo.tagNames);

                // Set cleared/reconciled field
                crCell.setValue(validationInfo.reconciled ? "R" : (validationInfo.cleared ? "c" : null));

                // Set Tag IDs field
                var tagIdsCell = txnSheet.getRange(editRowNum, Const.IDX_TXN_TAG_IDS + 1);
                tagIdsCell.setValue(validationInfo.tagIds);
              }
              
              tagsValidated = true;
              break;
              
            default:
              // For all other columns ...
              if (editType === Const.EDITTYPE_EDIT) {
                // Existing txn edited: Any change to an *editable* column is considered valid. Edits to other columns are invalid.
                isValid = (Const.TXN_EDITABLE_FIELDS.indexOf(colIndex) >= 0);
              } else {
                isValid = true;  // We assume values for a new txn are valid
              }
              break;
          }
          
          if (!isValid) {
            if (Debug.enabled) Debug.log(Utilities.formatString("Txn row %d, column %d is not valid", editRowNum, colIndex + 1));
            break;  // if a column value is invalid, there's no point in validating other columns
          }

        } // for

        if (Debug.traceEnabled) Debug.trace("Row edit is %s", (isValid ? "valid" : "INVALID"));

        if (isValid) {
          this.setTransactionEditStatus(txnSheet, editRowNum, editType);
        }
      }
      catch(e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }
      
      return isValid;
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    setTransactionEditStatus: function(txnSheet, row, editType)
    {
      var editStatusCell = txnSheet.getRange(row, Const.IDX_TXN_EDIT_STATUS + 1);
      var prevEditStatus = editStatusCell.getValue();
      var newEditStatus = "";

      if (prevEditStatus === null || prevEditStatus === "") {
        newEditStatus = editType;
      }
      else if (prevEditStatus === Const.EDITTYPE_NEW) {
        // If the current edit status is 'new', then leave it as-is
        newEditStatus = prevEditStatus;
      }
      else if (prevEditStatus.indexOf(editType) >= 0)
      {
        if (Debug.enabled) Debug.log("Edit status already has edit type '%s'. Leaving it intact: '%s'", editType, prevEditStatus);
        newEditStatus = prevEditStatus;
      }
      else {
        // Make sure edit type 'split' comes before 'edit' (S, E, or SE, but not ES).
        if (editType === Const.EDITTYPE_SPLIT) {
          newEditStatus = Const.EDITTYPE_SPLIT + prevEditStatus;
        }
        else if (editType === Const.EDITTYPE_EDIT) {
          newEditStatus = prevEditStatus + Const.EDITTYPE_EDIT;
        }
        else {
          newEditStatus = editType;
        }
      }

      editStatusCell.setValue(newEditStatus);
      editStatusCell.setFontWeight("bold");
      if (Debug.traceEnabled) Debug.trace("Row %d edit status set to '%s'", row, newEditStatus);
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    getModifiedTransactionRows: function(mintAccount, editValues, mintAcctValues)
    {
      var pendingUpdates = [];

      if (editValues === undefined)
        editValues = null;
      if (mintAcctValues === undefined)
        mintAcctValues = null;

      if (editValues == null || mintAcctValues == null) {

        var trv = Sheets.TxnData.getTxnRangeAndValues();
        var txnRange = trv.txnRange;

        if (editValues == null) {
          // Get the "Edit" column
          var editRange = txnRange.offset(0, Const.IDX_TXN_EDIT_STATUS, txnRange.getNumRows(), 1);
          editValues = editRange.getValues();
        }
        
        if (mintAcctValues == null) {
          //  Get the "Mint Account" column
          var mintAcctRange = txnRange.offset(0, Const.IDX_TXN_MINT_ACCOUNT, txnRange.getNumRows(), 1);
          mintAcctValues = mintAcctRange.getValues();
        }
      }

      var numRows = editValues.length;

      // Get the list of edited transactions for the specified mint account
      for (var i = 0; i < numRows; ++i) {
        var editVal = editValues[i][0];
        var txnMintAcct = mintAcctValues[i][0];
        if (editVal != null && editVal != ""
            && txnMintAcct != null && txnMintAcct.toLowerCase() === mintAccount)
        {
          pendingUpdates.push(i);
        }
      }
      
      return pendingUpdates;
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    saveModifiedTransactions: function(mintAccount, interactive)
    {
      var allSucceeded = true;

      try
      {
        mintAccount = mintAccount.toLowerCase();
        if (Debug.enabled) Debug.log(Utilities.formatString("Saving modified txns for mint account '%s'", mintAccount));

        if (interactive === undefined) {
          interactive = false;
        }

        var trv = Sheets.TxnData.getTxnRangeAndValues();
        var txnRange = trv.txnRange;
        var numRows = txnRange.getNumRows();
        var startRow = txnRange.getRow();

        // Get the "Edit" column and the "Mint Account" column
        var editRange = txnRange.offset(0, Const.IDX_TXN_EDIT_STATUS, numRows, 1);
        var editValues = editRange.getValues();
        var editBgColors = editRange.getBackgrounds();
        var mintAcctRange = txnRange.offset(0, Const.IDX_TXN_MINT_ACCOUNT, numRows, 1);
        var mintAcctValues = mintAcctRange.getValues();

        var pendingUpdates = this.getModifiedTransactionRows(mintAccount, editValues, mintAcctValues);

        var updateCount = pendingUpdates.length;
        if (updateCount === 0)
        {
          if (interactive) toast("There are no transaction changes to save.", "Transaction update");
          return true; // nothing to do
        }

        var failedUpdates = [];
        var successCount = 0;

        var cookies = Mint.Session.getCookies();

        if (interactive) toast(Utilities.formatString("Uploading %d transaction(s) to Mint", updateCount), "Transaction update", 60);

        for (var i = 0; i < updateCount; ++i) {
          var rowIndex = pendingUpdates[i];

          var editStatus = editValues[rowIndex][0];
          if (editStatus === null || editStatus === "") {
            // This was probably a split txn that we already handled.
            continue;
          }
          editStatus = editStatus.toUpperCase();

          var rowRange = txnRange.offset(rowIndex, 0, 1, Const.IDX_TXN_LAST_COL + 1);
          var rowValues = rowRange.getValues();
          var isSplitTxn = (editStatus.indexOf(Const.EDITTYPE_SPLIT) >= 0);
          var splitRows = [];
          var updateFormData = null;
          var trv = null;

          if (isSplitTxn) {
            // Handle all transactions in the same split group as one update
            trv = Sheets.TxnData.getTxnRangeAndValues();
            splitRows = this.getTransactionRowsForSplitGroup(trv.txnValues, rowValues[0][Const.IDX_TXN_PARENT_ID], -1);
            updateFormData = Mint.TxnData.getSplitUpdateFormData(splitRows, trv.txnValues, true);
          }
          else if (editStatus === Const.EDITTYPE_DELETE) {
            // Deleting a txn is only supported in debug mode, and user must manually delete
            // the spreadsheet rows after the mint update is complete.
            if (Debug.txnDeleteEnabled) {
              updateFormData = Mint.TxnData.getUpdateFormData(rowValues[0], false, Const.EDITTYPE_DELETE, true);
            }
          }
          else {
            updateFormData = Mint.TxnData.getUpdateFormData(rowValues[0], false, editStatus, true);
          }

          if (updateFormData == null) {
            failedUpdates.push(i);
            if (Debug.enabled) Debug.log(Utilities.formatString("Unable to get form data for row %d. Skipping.", startRow + rowIndex));
            continue;
          }

          //
          // Upload the transaction to Mint
          //
          var result = Mint.TxnData.updateTransaction(cookies, updateFormData, false);

          if (result.success) {

            // Handle split changes first
            if (isSplitTxn) {
              var splitTxnIds = result.responseJson.txnId;
              Debug.assert(rowValues[0][Const.IDX_TXN_PARENT_ID] === splitTxnIds[0], "rowValues[0][IDX_TXN_PARENT_ID] === splitTxnIds[0]");
              var splitCount = splitRows.length;
              
              if (splitCount === 1) {
                // Split group only has one txn, so revert it to a non-split txn
                Debug.assert(splitTxnIds.length === 1, "result.responseJson.txnId.length === 1");
                var splitRowRange = trv.txnRange.offset(splitRows[0], 0, 1, Const.IDX_TXN_LAST_COL);
                var splitRowValues = splitRowRange.getValues();
                splitRowValues[0][Const.IDX_TXN_ID] = splitTxnIds[0];
                splitRowValues[0][Const.IDX_TXN_PARENT_ID] = null;
                splitRowValues[0][Const.IDX_TXN_STATE] = null;
                splitRowRange.setValues(splitRowValues);
                Debug.log("Reverting single split txn to a regular non-split txn");

                // Clear "S" from edit status
                var newEditStatus = editStatus.replace(Const.EDITTYPE_SPLIT, "");
                editValues[ splitRows[0] ][0] = newEditStatus;
                editBgColors[rowIndex][0] = Const.NO_COLOR;
                if (Debug.enabled) Debug.log("Split removed: Changing txn row index %s to edit status '%s'", splitRows[0], newEditStatus);
                if (newEditStatus.length === 0) {
                  // This txn has no more changes to save. Increment the 'success' counter
                  ++successCount;
                }
              } else {
                // The response contains a txn id for each of the split items (sorted in numerical order).
                // We need to sort our array of split item rows so that the txn ids are also in numerical
                // order to line up with the response. Note that new split items will have a txn id of 0, so
                // we should put those at the end so they get the newest txn ids in the response.
                if (Debug.traceEnabled) {
                  Debug.trace("result.responseJson.txnId: " + JSON.stringify(result.responseJson.txnId));
                  Debug.trace("splitRows BEFORE: " + JSON.stringify(splitRows));
                }
                Debug.assert(splitCount === splitTxnIds.length - 1, "splitCount === result.responseJson.txnId.length - 1");
                splitRows = splitRows.sort(function(a, b) {
                  var aTxnId = trv.txnValues[a][Const.IDX_TXN_ID];
                  var bTxnId = trv.txnValues[b][Const.IDX_TXN_ID];
                  var result = 0;
                  if (aTxnId === 0) {
                    result = 1;
                  }
                  else if (bTxnId === 0) {
                      result = -1;
                  }
                  else {
                    result = (aTxnId < bTxnId ? -1 : (aTxnId > bTxnId ? 1 : 0));
                  }
                  //Debug.log("(a, b): %s, %s, return %s", aTxnId, bTxnId, result);
                  return result;
                });
                if (Debug.traceEnabled) Debug.trace("splitRows AFTER: " + JSON.stringify(splitRows));

                // Set the txn id of each new split txn in the group, and clear the edit status column for all split txns
                for (var j = 0; j < splitCount; ++j) {
                  var splitRowNum = splitRows[j];
                  var responseTxnId = splitTxnIds[j + 1]; // + 1 because the first entry in the array is the parent txn id
                  var splitTxnId = trv.txnValues[splitRowNum][Const.IDX_TXN_ID];

                  if (splitTxnId === 0) {
                    var splitTxnIdCell = trv.txnRange.offset(splitRowNum, Const.IDX_TXN_ID, 1, 1);
                    splitTxnIdCell.setValue(responseTxnId);
                  }
                  else {
                    Debug.assert(responseTxnId === splitTxnId, "responseTxnId (" + responseTxnId + ") === splitTxnId (" + splitTxnId + ")");
                  }
                  
                  // Clear "S" from the edit status
                  var newEditStatus = editValues[splitRowNum][0]; // Get current edit type for this split item
                  newEditStatus = newEditStatus.replace(Const.EDITTYPE_SPLIT, "");
                  editValues[splitRowNum][0] = newEditStatus;
                  editBgColors[rowIndex][0] = Const.NO_COLOR;
                  if (Debug.enabled) Debug.log("Split: Changing txn row index %s to edit status '%s'", splitRowNum, newEditStatus);
                  if (newEditStatus.length === 0) {
                    // This txn has no more changes to save. Increment the 'success' counter
                    ++successCount;
                  }
                } // for j

              }

              --i; // decrement 'i' so we look at this row again. It my also have the 'E' edit status.
            }
            else {
              // For a all non-split changes (edit, new, delete), just clear the edit status column
              var editStatusRow = editValues[rowIndex];
              // The txn should have only one edit status: E, N, or D
              Debug.assert(editStatusRow[0].length === 1, Utilities.formatString("Unexpected: Txn at index %s has multiple edit statuses: %s", rowIndex, editStatusRow[0]));
              editStatusRow[0] = "";
              editBgColors[rowIndex][0] = Const.NO_COLOR;
              if (Debug.enabled) Debug.log("Clearing edit status of txn row index %s", rowIndex);
              // This txn has no more changes to save. Increment the 'success' counter.
              ++successCount;
            }

          } else { // result.success == false
            Debug.log("Update failed -- One possible reason for this is that the Account ID for '%s' has changed. Check the AccountData sheet.", rowValues[0][Const.IDX_TXN_ACCOUNT]);
            failedUpdates.push(rowIndex);
            // Change background color of edit status column to light red
            editBgColors[rowIndex][0] = Const.COLOR_ERROR;
          }
        } // for i

        var failMsg = "";
        var failCount = failedUpdates.length;
        if (failCount > 0) {
          allSucceeded = false;
          failMsg = Utilities.formatString(", but %d failed. If the problem persists, try re-importing the failing transaction(s) and updating them again.", failCount);
        }

        editRange.setValues(editValues);
        editRange.setBackgrounds(editBgColors);
        if (interactive) toast(Utilities.formatString("%d transaction(s) successfully uploaded%s", successCount, failMsg), "Transaction update", (failCount > 0 ? 15 : 5));
      }
      catch (e)
      {
        allSucceeded = false;
        Debug.log(Debug.getExceptionInfo(e));
        
        toast("Update failed.  Exception: " + Debug.getExceptionInfo(e), "Transaction update", 30);
      }

      return allSucceeded;
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    getTransactionRowsForSplitGroup: function(txnValues, parentId, rowToExclude)
    {
      return this.findAllTxnRowsUnsorted(txnValues, txnValues.length, [Const.IDX_TXN_PARENT_ID], [parentId], rowToExclude);
    },
    
    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    handleAmountEdit: function(editRow, editCol)
    {
      try
      {
        toast("Processing split transaction change.");

        var trv = Sheets.TxnData.getTxnRangeAndValues();
        var sheet = trv.txnRange.getSheet();

        if (!this.isTransactionEditAllowed(sheet, editRow, editCol, editCol, Const.EDITTYPE_SPLIT, true))
            return;

        var rowRange = sheet.getRange(editRow, 1, 1, Const.IDX_TXN_LAST_COL + 1);
        var rowValues = rowRange.getValues();

        var isSplitTxn = (rowValues[0][Const.IDX_TXN_STATE] === Const.TXN_STATUS_SPLIT);
        var newAmount = rowValues[0][Const.IDX_TXN_AMOUNT];
        var origAmount = rowValues[0][Const.IDX_TXN_ORIG_AMOUNT];
        var splitAmount = Math.round((origAmount - newAmount)*100) / 100; // Round off to cents
        // If this is not already a split txn, then the txn id will become the parent id
        var parentId = (isSplitTxn ? rowValues[0][Const.IDX_TXN_PARENT_ID] : rowValues[0][Const.IDX_TXN_ID]);

        var offsetEditRow = editRow - trv.txnRange.getRow();

        var insertNewSplit = false;
        var rebalanceSplitAmount = false;
        var deleteRow = false;
        var revertChange = false;

        if (Debug.enabled) Debug.log(Utilities.formatString("Processing split. newAmount: %f, origAmount: %f", newAmount, origAmount));

        if (newAmount == null || Math.round(newAmount * 100) === 0) {
          // Amount was set to null or 0

          if (isSplitTxn === true) {
            if ("yes" === Browser.msgBox("Remove split transaction?", Browser.Buttons.YES_NO)) {
              // User wants to remove this split txn
              deleteRow = true;

              // Transfer split amount of edited row to another split txn
              rebalanceSplitAmount = true;

            } else {
              // If user clicked "No", then we can't just leave the amount as 0.
              // Revert the amount to the original value.
              revertChange = true;
            }

          } else {
            // This is NOT a split txn. Setting amount to 0 is invalid.
            toast("You can only set the amount to 0.00 to remove a split transaction. Please \"undo\" your change.");
            return;
          }
        } else { // Amount is > 0.00

          if (isSplitTxn === true && "no" === Browser.msgBox("Split transaction", "Would you like to add a new split transaction?", Browser.Buttons.YES_NO)) {
            // Transfer remaining split amount to another split txn in the group
            rebalanceSplitAmount = true;
          } else {
            insertNewSplit = true;
          }
        }

        if (insertNewSplit === true)
        {
          // insert new row with copy of fields, amount set to original_amount - new_amount
          toast("Inserting new split transaction", "", 60);
          var insertRowIndex = offsetEditRow + 2; // insert new row right after the edited row
          var newRowRange = Sheets.TxnData.insertNewTransaction(Const.EDITTYPE_SPLIT, insertRowIndex, rowValues[0][Const.IDX_TXN_DATE],
                                                                       rowValues[0][Const.IDX_TXN_ACCOUNT], rowValues[0][Const.IDX_TXN_MERCHANT],
                                                                       splitAmount, rowValues[0][Const.IDX_TXN_CATEGORY], "", "", "", parentId,
                                                                       rowValues[0][Const.IDX_TXN_MOJITO_PROPS], rowValues[0][Const.IDX_TXN_MINT_ACCOUNT]);
          var txnIdCell = newRowRange.offset(0, Const.IDX_TXN_ID, 1, 1);
          txnIdCell.setValue(0);

          toast("Split transaction inserted", "", 2);

          // Set focus to amount column of new row
          var newAmountCell = newRowRange.offset(0, Const.IDX_TXN_AMOUNT, 1, 1);
          newAmountCell.activate();
          
        }

        if (rebalanceSplitAmount === true) {
          
          toast("Rebalancing split amounts", "", 60);
          
          // Find another txn in split group
          var otherRowNum = this.findTxnRowInsideOut(trv.txnValues, trv.txnValues.length, [Const.IDX_TXN_PARENT_ID, Const.IDX_TXN_ID], [rowValues[0][Const.IDX_TXN_PARENT_ID], rowValues[0][Const.IDX_TXN_ID]], offsetEditRow + 1, offsetEditRow);
          if (otherRowNum < 0) {
            throw "No other split transaction found to transfer the balance to.";
          }
          
          var otherRowRange = trv.txnRange.offset(otherRowNum, 0, 1, Const.IDX_TXN_LAST_COL + 1);
          var otherRowValues = otherRowRange.getValues();
          
          // Transfer split amount of edited row to the other split txn
          var otherSplitAmount = Math.round((otherRowValues[0][Const.IDX_TXN_AMOUNT] + splitAmount) * 100) / 100; // Round to cents
          otherRowValues[0][Const.IDX_TXN_AMOUNT] = otherSplitAmount;
          otherRowValues[0][Const.IDX_TXN_ORIG_AMOUNT] = otherSplitAmount;
          otherRowRange.setValues(otherRowValues);
          
          // Mark other row with "S"
          this.setTransactionEditStatus(sheet, otherRowNum + trv.txnRange.getRow(), Const.EDITTYPE_SPLIT);

          toast("Split rebalance complete", "", 2);

        }

        if (deleteRow === true) {
          sheet.deleteRow(editRow);

        } else if (revertChange === true) {
          toast("No change made", "", 2);
          var amountCell = rowRange.offset(0, Const.IDX_TXN_AMOUNT, 1, 1);
          amountCell.setValue(origAmount);

        } else {
          // Update the edited row's original amount column
          rowValues[0][Const.IDX_TXN_ORIG_AMOUNT] = newAmount;
          // If this is a new split txn, then set the txn ID, parent ID, and txn status
          if (!isSplitTxn) {
            rowValues[0][Const.IDX_TXN_ID] = 0;
            rowValues[0][Const.IDX_TXN_PARENT_ID] = parentId;
            rowValues[0][Const.IDX_TXN_STATE] = Const.TXN_STATUS_SPLIT;
          }
          rowRange.setValues(rowValues);
          // Mark row with "S"
          this.setTransactionEditStatus(sheet, editRow, Const.EDITTYPE_SPLIT);
        }
      }
      catch(e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    insertNewTransaction: function(editType, rowIndex, date, account, merchant, amount, category, tags, cR, memo, parentId, propsJson, mintAccount)
    {
      var sheet = this.getSheet();
      var insertPos = sheet.getFrozenRows() + rowIndex;
      var numCols = Const.IDX_TXN_ORIG_AMOUNT + 1;

      Debug.log("Inserting new row at position %d", insertPos);
      sheet.insertRows(insertPos);

      var txnValues = [];
      var txnRow = new Array(numCols);
      txnRow[Const.IDX_TXN_DATE] = date;
      txnRow[Const.IDX_TXN_EDIT_STATUS] = null;
      txnRow[Const.IDX_TXN_ACCOUNT] = account;
      txnRow[Const.IDX_TXN_MERCHANT] = merchant;
      txnRow[Const.IDX_TXN_AMOUNT] = amount;
      txnRow[Const.IDX_TXN_CATEGORY] = category;
      txnRow[Const.IDX_TXN_TAGS] = tags;
      txnRow[Const.IDX_TXN_CLEAR_RECON] = cR;
      txnRow[Const.IDX_TXN_MEMO] = memo;
      txnRow[Const.IDX_TXN_MATCHES] = null;
      txnRow[Const.IDX_TXN_STATE] = (parentId != null ? Const.TXN_STATUS_SPLIT : null);
      // Internal values
      txnRow[Const.IDX_TXN_MINT_ACCOUNT] = mintAccount;
      txnRow[Const.IDX_TXN_ORIG_MERCHANT_INFO] = null;
      txnRow[Const.IDX_TXN_ID] = null;
      txnRow[Const.IDX_TXN_PARENT_ID] = parentId;
      txnRow[Const.IDX_TXN_CAT_ID] =  null;
      txnRow[Const.IDX_TXN_TAG_IDS] =  null;
      txnRow[Const.IDX_TXN_MOJITO_PROPS] = propsJson;
      txnRow[Const.IDX_TXN_YEAR_MONTH] = date.getFullYear() * 100 + (date.getMonth() + 1);
      txnRow[Const.IDX_TXN_ORIG_AMOUNT] = amount;
      txnValues.push(txnRow);

      // Format the cells and set the row values
      var rowRange = sheet.getRange(insertPos, 1, 1, numCols);
      var dateCell = rowRange.offset(0, Const.IDX_TXN_DATE, 1, 1);
      dateCell.setNumberFormat("M/d/yyyy");

      rowRange.setValues(txnValues);

      // Flag the txn as new. This will validate some fields.
      Sheets.TxnData.validateTransactionEdit(sheet, insertPos, Const.IDX_TXN_MERCHANT + 1, Const.IDX_TXN_CLEAR_RECON + 1, editType);

      return rowRange;
    },

    // Mint.TxnData
    formatSheet: function (range, txnValues, rowCount)
    {
      Debug.log("formatTxnSheet start");
      var numCols = Const.IDX_TXN_LAST_VIEWABLE_COL + 1;

      // Set text of pending txn rows to dark-gray and italics
      for (var i = 0; i < rowCount; ++i) {
        if (txnValues[i][Const.IDX_TXN_STATE] === Const.TXN_STATUS_PENDING) {
          var rowRange = range.offset(i, 0, 1, numCols);
          var fontValues = rowRange.getFontColors();
          for (var j = 0; j < numCols; ++j) {
            fontValues[0][j] = Const.COLOR_TXN_PENDING;
          }
          rowRange.setFontColors(fontValues);

          for (var j = 0; j < numCols; ++j) {
            fontValues[0][j] = "italic";
          }
          rowRange.setFontStyles(fontValues);
        }
      }

// Formatting the appearance is taking too long.
// Commenting out parts that aren't critical
//
//      // Set "E" and "c/R" columns to bold
//      var eColRange = range.offset(0, Const.IDX_TXN_EDIT_STATUS, rowCount, 1);
//      var crColRange = range.offset(0, Const.IDX_TXN_CLEAR_RECON, rowCount, 1);
//      var fontWeights = eColRange.getFontWeights()
//      for (var i = 0; i < rowCount; ++i) {
//            fontWeights[i][0] = "bold";
//      }
//      eColRange.setFontWeights(fontWeights);
//      crColRange.setFontWeights(fontWeights);

//      Debug.log("formatTxnSheet internal cols");
//
//      // Set the text color of the "internal only" columns to light gray
//      var numInternalCols = Const.IDX_TXN_LAST_COL - Const.IDX_TXN_LAST_VIEWABLE_COL;
//      var internalRange = range.offset(0,Const.IDX_TXN_LAST_VIEWABLE_COL + 1, rowCount, numInternalCols);
//      SpreadsheetUtils.setRowColors(internalRange, null, false, COLOR_TXN_INTERNAL_FIELD, true);
//
//      // Set some internal columns to "no wrapping"
//      var wrapValues = internalRange.getWraps();
//      for (var i = 0; i < rowCount; ++i) {
//          for (var j = 0; j < numInternalCols; ++j) {
//            wrapValues[i][j] = false;
//          }
//      }
//      internalRange.setWraps(wrapValues);

      if (Debug.enabled) Debug.log("Sheets.TxnData.formatSheet exit");
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    clearRowHighlights: function() {
      var txnDataRange = Utils.getTxnDataRange();
      var highlightStartCol = txnDataRange.getSheet().getFrozenColumns();
      var txnHighlightRange = txnDataRange.offset(0, highlightStartCol, txnDataRange.getNumRows(), Const.IDX_TXN_MATCHES - highlightStartCol);
      SpreadsheetUtils.setRowColors(txnHighlightRange, null, false, Const.NO_COLOR);
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    showTxnMatches: function(matchName, updateObj, sortCriteria) {

      var currDisplayedMatches = this.getTxnMatchesHeader();
      if (currDisplayedMatches !== matchName) {
        updateObj.updateCalculations();
      }

      toast(Utilities.formatString("Sorting transactions by '%s' and highlighting rows.", matchName), "Action", 30);
      var txnDataRange = Utils.getTxnDataRange();
      
      var txnDataLen = txnDataRange.getNumRows();
      var txnMatchesRange = txnDataRange.offset(0, Const.IDX_TXN_MATCHES, txnDataLen, 1);
      var highlightColors = txnMatchesRange.getBackgrounds();
      var highlightStartCol = txnDataRange.getSheet().getFrozenColumns();
      
      var txnHighlightRange = txnDataRange.offset(0, highlightStartCol, txnDataLen, Const.IDX_TXN_MATCHES - highlightStartCol);
      SpreadsheetUtils.setRowColors(txnHighlightRange, highlightColors, true, Const.NO_COLOR, false);

      if (sortCriteria === undefined) {
        sortCriteria = updateObj.getTxnSortCriteria();
      }
      txnDataRange.sort(sortCriteria);

      var firstCell = txnDataRange.offset(0, Const.IDX_TXN_ACCOUNT, 1, 1);
      firstCell.activate();

      toast("Done");
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    getTxnMatchesHeader: function() {
      var sheet = this.getSheet();
      var txnMatchesCell = sheet.getRange(sheet.getFrozenRows(), Const.IDX_TXN_MATCHES + 1);
      return txnMatchesCell.getValue();
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    setTxnMatchesHeader: function(hdrName) {
      var sheet = this.getSheet();
      var txnMatchesCell = sheet.getRange(sheet.getFrozenRows(), Const.IDX_TXN_MATCHES + 1);
      return txnMatchesCell.setValue(hdrName);
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    getRecentReconcileBalances: function (fiAccount) {
      var accountMap = {};
      
      var txnRange = Utils.getTxnDataRange();
      var txnDataLen = txnRange.getNumRows();
      var propsRange = txnRange.offset(0, Const.IDX_TXN_MOJITO_PROPS, txnDataLen, 1);
      var propsValues = propsRange.getValues();
      var accountRange = txnRange.offset(0, Const.IDX_TXN_ACCOUNT, txnDataLen, 1);
      var accountValues = accountRange.getValues();
      var dateRange = txnRange.offset(0, Const.IDX_TXN_DATE, txnDataLen, 1);
      var dateValues = dateRange.getValues();

      for (var i = 0; i < txnDataLen; ++i) {
        if (fiAccount != null && fiAccount !== accountValues[i][0]) {
          continue; // Financial institution account doesn't match specified account. Skip it.
        }

        var propsJson = propsValues[i][0];
        if ( propsJson === "" || propsJson == null) {
          continue; // no props
        }

        var acct = accountValues[i][0];
        if (acct == null) {
          continue; // no account!?
        }

        var reconInfo = accountMap[acct];

        if (reconInfo != null) {
          if (reconInfo.date >= dateValues[i][0]) {
            //Debug.log("reconInfo date is not the most recent");
            continue;  // We already have reconcile info that is more recent than this one
          }
        }

        var props = JSON.parse(propsJson);
        if (props == null || props.type !== "reconcile") {
          continue; // props are invalid, or this is not a reconcile record
        }
        
        if (reconInfo == null) {
          reconInfo = { date : dateValues[i][0], balance : props.balance };
          accountMap[acct] = reconInfo;
        } else {
          reconInfo.date = dateValues[i][0];
          reconInfo.balance = props.balance;
        }
      }

      if (Debug.enabled) Debug.log("Reconciled account balances map: " + accountMap.toSource());
      return accountMap;
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    findAllTxnRowsUnsorted: function(unsortedTxnValues, arrayLen, colIndexes, values, excludeRow)
    {
      // Return the all matching rows, but skip excludeRow, if specified
      var matches = [];
      var colCount = colIndexes.length;
      
      if (excludeRow === undefined)
        excludeRow = -1;

      for (var i = 0; i < arrayLen; ++i) {
        if (i === excludeRow)
          continue;

        for (var j = 0; j < colCount; ++j) {
          var colIndex = colIndexes[j];
          if (unsortedTxnValues[i][colIndex] === values[j]) {
            matches.push(i);
            break; // Row matches. Move on to next one
          }
        }
      }
      return matches;
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    findTxnRowInsideOut: function(unsortedTxnValues, arrayLen, colIndexes, values, startRow, excludeRow)
    {
      // Return the first matching row, but skip excludeRow, if specified
      var idx = -1;
      var colCount = colIndexes.length;

      if (excludeRow === undefined)
        excludeRow = -1;

      // Search from startRow to the end
      for (var i = startRow; i < arrayLen && idx < 0; ++i) {
        if (i === excludeRow) {
          continue;
        }

        for (var j = 0; j < colCount; ++j) {
          var colIndex = colIndexes[j];
          if (unsortedTxnValues[i][colIndex] === values[j]) {
            idx = i;
            break;
          }
        }
      }

      // Search from startRow to the beginning
      for (var i = startRow - 1; i >= 0 && idx < 0; --i) {
        if (i === excludeRow) {
          continue;
        }

        for (var j = 0; j < colCount; ++j) {
          var colIndex = colIndexes[j];
          if (unsortedTxnValues[i][colIndex] === values[j]) {
            idx = i;
            break;
          }
        }
      }

      return idx;
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    findTxnSorted: function(sortedTxnValues, arrayLen, txnId)
    {
      if (Debug.enabled) Debug.log("Finding txn id " + txnId);
      var found = false;
      var idx = 0;

      try {
      // Perform binary search
      var start = 0;
      var end = arrayLen - 1;
      while(end - start < 1)
      {
        idx = parseInt(start + (((end - start) / 2)));// | 0); // bit-OR with zero to remove decimal, if any
//        Debug.log(Utilities.formatString("start: %d, end: %d, idx: %d", start, end, idx));
        var thisTxnId = sortedTxnValues[idx][Const.IDX_TXN_ID];
        if (thisTxnId === txnId) {
          found = true;
          break;
        }
        
        if (txnId > thisTxnId) {
          start = idx;
        }
        else {
          end = idx;
        }
      }

      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
      }

      return (found ? idx : -1);
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    mergeTxnValues: function(newValues, existingValues)
    {
      // Overall algorithnm
      // 1. Delete pending txns and all txns with dates on or after the import start date
      // 2. Add new txns
      // 3. Sort existing txns in memory, by date descending


      // 1. Delete pending txns and all txns with dates on or after the import start date
      var firstDateNewTxns = Utils.find2dArrayMinValue(newValues, Const.IDX_TXN_DATE);
      var existingLen = existingValues.length;

      // Loop through backwards so the index 'i' doesn't get messed up if we delete an entry
      for (var i = existingLen - 1; i >= 0; --i) {
        if (existingValues[i][Const.IDX_TXN_STATE] === Const.TXN_STATUS_PENDING) {
          // Delete the pending transaction
          existingValues.splice(i, 1);
          
        } else if (existingValues[i][Const.IDX_TXN_DATE] >= firstDateNewTxns) {
          // Delete existing txn with date on or after the import start date
          existingValues.splice(i, 1);
        }
      }

      // 2. Add new txns
      // NOTE: It's possible that a new txn could already exist (same txn id, but txn date changed), so caller
      // should play it safe by retrieving txns that occurred a few days before the desired date.
      var newTxnsLen = newValues.length;
      for (var i = 0; i < newTxnsLen; ++i) {
        existingValues.push(newValues[i]);
      }

      // 3. Sort existing txns in memory, by date descending
      Debug.log("Sorting txn array by date, descending");
      Utils.sort2dArray(existingValues, [Const.IDX_TXN_DATE, Const.IDX_TXN_PARENT_ID], [-1, 1]);
      Debug.log("Sort by date finished.");
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    findTxnRowUnsorted: function(unsortedTxnValues, arrayLen, colIndexes, values, excludeRow, matchAll)
    {
      // Return the first matching row, but skip excludeRow, if specified
      var idx = -1;
      var colCount = colIndexes.length;

      if (excludeRow === undefined)
        excludeRow = -1;
      if (matchAll === undefined)
        matchAll = false;

      for (var i = 0; i < arrayLen && idx < 0; ++i) {
        if (i === excludeRow) {
          continue;
        }

        var colMatchCount = 0;
        for (var j = 0; j < colCount; ++j) {
          var colIndex = colIndexes[j];
          if (unsortedTxnValues[i][colIndex] === values[j]) {
            if (matchAll !== true || (matchAll === true && ++colMatchCount === colCount)) {
              idx = i;
              break;
            }
          }
        }
      }
      return idx;
    },

    //-----------------------------------------------------------------------------
    // Sheets.TxnData
    mergeTxnValuesOld: function(txnValues, existingValues)
    {
      Debug.log("Sorting txn array by txn id");
      Utils.sort2dArray(existingValues, [Const.IDX_TXN_ID], [1]);
      Debug.log("Sort by txn id finished.");
      var existingLen = existingValues.length;
      var length = txnValues.length;
      for (var i = 0; i < length; ++i) {
        var matchingTxnRow = this.findTxnSorted(existingValues, existingLen, txnValues[i][Const.IDX_TXN_ID]);
        if (matchingTxnRow >= 0) {
          Debug.log("Found existing txn id " + txnValues[i][Const.IDX_TXN_ID]);
          existingValues.splice(matchingTxnRow, 1, txnValues[i]);
        } else {
          Debug.log("Adding new txn id " + txnValues[i][Const.IDX_TXN_ID]);
          existingValues.push(txnValues[i]);
        }
      }
      Debug.log("Sorting txn array by date");
      Utils.sort2dArray(existingValues, [Const.IDX_TXN_DATE], [1]);
      Debug.log("Sort by date finished.");
    },

  }, // TxnData

  ////////////////////////////////////////////////////////////////////////////////////////////////
  AccountData: {

    getSheet: function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(Const.SHEET_NAME_ACCTDATA);
      return sheet;
    },

    _acctRanges : null,  // This variable is only valid while server-side code is executing. It resets each time.

    getAccountDataRanges: function(reset) {
      if (this._acctRanges === null || reset === true) {
        this._acctRanges = Utils.getAccountDataRanges(true);
      }
      return this._acctRanges;
    },

    //-----------------------------------------------------------------------------
    // Sheets.AccountData
    getAccountInfoMap: function()
    {
      var acctInfoMap = null;

      var cache = Utils.getPrivateCache();
      var acctInfoMapJson = cache.get(Const.CACHE_ACCOUNT_INFO_MAP);
      if (acctInfoMapJson != null && acctInfoMapJson != "{}")
      {
        Debug.log("AccountInfo map found in cache. Parsing JSON.");
        acctInfoMap = JSON.parse(acctInfoMapJson);
      }
      else
      {
        Debug.log("Rebuilding AccountInfo map");
        var acctRanges = this.getAccountDataRanges();
        var acctValues = (acctRanges.hdrRange == null ? null : acctRanges.hdrRange.getValues());
        var acctCount = ( acctValues == null ? 0 : acctValues[0].length);
        if (acctCount > 0) {
          acctInfoMap = {};
        }

        for (var i = 0; i < acctCount; ++i) {
          var name = acctValues[Const.IDX_ACCT_NAME][i];
          var acctEntry = {
            name : name,
            id : acctValues[Const.IDX_ACCT_ID][i],
            fi : acctValues[Const.IDX_ACCT_FINANCIAL_INST][i],
            type : acctValues[Const.IDX_ACCT_TYPE][i],
          };
          
          acctInfoMap[name] = acctEntry;
        }

        if (acctInfoMap != null)
        {
          acctInfoMapJson = JSON.stringify(acctInfoMap);
          if (Debug.enabled) Debug.log("Saving AccountInfo map in cache: %s", acctInfoMapJson);
          cache.put(Const.CACHE_ACCOUNT_INFO_MAP, acctInfoMapJson, 30*60); // Save acctInfoMap in cache for 30 minutes
        }
      }

      return acctInfoMap;
    },

    //-----------------------------------------------------------------------------
    // Sheets.AccountData
    isUpToDate: function()
    {
      var lastUpdate = this.getLastUpdate();
      if (Debug.enabled) Debug.log("AccountData, last update: " + lastUpdate);
      var now = new Date();
      var yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      if (lastUpdate != null && lastUpdate >= yesterday)
      {
        return true;
      }

      return false;
    },

    //-----------------------------------------------------------------------------
    // Sheets.AccountData
    getLastUpdate: function()
    {
      var lastUpdate = null;
      var ranges = this.getAccountDataRanges();
      if (ranges.balanceRange !=  null) {
        var dateColRange = ranges.balanceRange.offset(0, -1, ranges.balanceRange.getNumRows(), 1);
        var dateValues = dateColRange.getValues();
        lastUpdate = Utils.find2dArrayMaxValue(dateValues, 0);
        if (!(lastUpdate instanceof Date)) {
          lastUpdate = null;
        }
      }
      return lastUpdate;
    },

    //-----------------------------------------------------------------------------
    // Sheets.AccountData
    determineImportDateRange: function() {
      var today = new Date();
      var lastUpdate = this.getLastUpdate();
      if (Debug.enabled) Debug.log("AccountData, last update: " + lastUpdate);
      var startDate = null;
      if (lastUpdate == null) {
        // If no transactions have been imported yet, then just use January 1st
        // as the start date.
        startDate = new Date(today.getFullYear(), 0, 1);
      } else {
        if (today.getTime() - lastUpdate.getTime() <= Const.ONE_DAY_IN_MILLIS) {
          startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 15);
        } else {
          startDate = lastUpdate;
        }
      }
      var endDate = today; // Set end date to today

      return { startDate : startDate, endDate : endDate };
    },

    //-----------------------------------------------------------------------------
    // Sheets.AccountData
    insertAccountBalanceHistory: function(account)
    {
      var newBalCount = account.balanceHistory.length;

      var acctRanges = Utils.getAccountDataRanges(); // Call Utils impl to get ranges with 1 row if no data exists.
      var firstInsert = (acctRanges.isEmpty === true);

      var numCols = acctRanges.hdrRange.getNumColumns();
      var hdrRange = acctRanges.hdrRange.offset(0, 0, acctRanges.hdrRange.getNumRows(), numCols);
      var hdrValues = hdrRange.getValues();

      var mintLogin = Utils.getMintLoginAccount();
      mintLogin = (mintLogin != null ? mintLogin.toLowerCase() : null);

      // Find column index of this account
      var j = 0;
      for (; j < numCols; ++j)
      {
        // if the account id at index 'j' is the one we're looking for
        // or we've reached the end of the accounts (meaning this is a new account)
        // then we have our position.
        var acctId = hdrValues[Const.IDX_ACCT_ID][j];
        if (acctId === account.id || acctId === "" || isNaN(acctId) || acctId == null)
        {
          break;
        }
        // If this is the Mint demo data, then the account IDs probably won't match.
        // We'll try matching the account name and financial institution instead.
        if (mintLogin === Const.DEMO_MINT_LOGIN
            && account.name === hdrValues[Const.IDX_ACCT_NAME][j]
            && account.fi === hdrValues[Const.IDX_ACCT_FINANCIAL_INST][j])
        {
          break;
        }
      }

      var acctCol = j;
      var newHdrRange = hdrRange.offset(0, acctCol, hdrRange.getNumRows(), 1);
      var newHdrValues = newHdrRange.getValues();
      newHdrValues[Const.IDX_ACCT_NAME][0] = account.name;
      newHdrValues[Const.IDX_ACCT_FINANCIAL_INST][0] = account.fi;
      newHdrValues[Const.IDX_ACCT_TYPE][0] = account.klass;
      newHdrValues[Const.IDX_ACCT_ID][0] = account.id;
      var acctNameCell = newHdrRange.offset(0, 0, 1, 1);
      acctNameCell.setFontWeight("bold");

      var newDateValues = new Array(newBalCount);
      var newBalValues = new Array(newBalCount);
      var isDebtAcct = (account.klass === "credit" || account.klass === "loan");

      // Copy balances to newBalValues array and corresponding dates to newDateValues array
      for (var i = 0; i < newBalCount; ++i) {
        var balanceEntry = account.balanceHistory[i];
        newDateValues[i] = [balanceEntry.date];
        var balance = (isDebtAcct ? -balanceEntry.value : balanceEntry.value);
        newBalValues[i] = [balance];
      }

      var balCount = acctRanges.balanceRange.getNumRows();
      var dateRange = acctRanges.dateRange;
      var dateValues = (firstInsert ? [] : dateRange.getValues());
      var balRange = acctRanges.balanceRange.offset(0, acctCol, balCount, 1);
      var balValues = (firstInsert ? [] : balRange.getValues());

      // NOTE: Entries in newDataValues[] are epoch timestamps, whereas entries in dateValues are Date's.
      var insertDates = (dateValues.length === 0 || dateValues.length !== newDateValues.length || dateValues[0][0].getTime() !== newDateValues[0][0] || dateValues[balCount - 1][0].getTime() !== newDateValues[newBalCount - 1][0]);

      Sheets.AccountData.mergeBalanceValues(newDateValues, newBalValues, dateValues, balValues);

      if (insertDates) {
        if (Debug.enabled) Debug.log("Inserting dates for account balances");
        dateRange = acctRanges.dateRange.offset(0, 0, dateValues.length, 1);
        dateRange.setValues(dateValues);
      }

      newHdrRange.setValues(newHdrValues);
      balRange = acctRanges.balanceRange.offset(0, acctCol, balValues.length, 1);
      balRange.setValues(balValues);
    },

    //-----------------------------------------------------------------------------
    // Sheets.AccountData
//    insertAccountBalanceHistory_old: function(acctBalanceHistoryArray)
//    {
//      if (acctBalanceHistoryArray == null || acctBalanceHistoryArray.length === 0)
//        return;
//
//      var acctCount = acctBalanceHistoryArray.length;
//      Debug.log("inserting " + acctCount + " accounts");
//      var newBalanceCount = acctBalanceHistoryArray[0].balanceHistory.length;
//
//      var firstInsert = (this.getAccountDataRanges().balanceRange == null ? true : false);
//      var acctRanges = Utils.getAccountDataRanges(); // Call Utils impl to get ranges with 1 row if no data exists.
//
//      var numCols = Math.max(acctRanges.hdrRange.getNumColumns(), acctCount);
//      var hdrRange = acctRanges.hdrRange.offset(0, 0, acctRanges.hdrRange.getNumRows(), numCols);
//      var existingCount = (firstInsert ? 0 : acctRanges.balanceRange.getNumRows());
//      var existingRange = acctRanges.balanceRange.offset(0, -1, acctRanges.balanceRange.getNumRows(), numCols + 1);
//      var range = acctRanges.balanceRange.offset(existingCount, -1, newBalanceCount, numCols + 1);
//
//      var hdrValues = hdrRange.getValues();
//      var existingValues = (existingCount === 0 ? null : existingRange.getValues());
//      var newValues = range.getValues();
//
//      // Copy new dates into column 0
//      for (var i = 0; i < newBalanceCount; ++i)
//      {
//        newValues[i][0] = new Date(acctBalanceHistoryArray[0].balanceHistory[i].date);
//      }
//
//      for (var i = 0; i < acctCount; ++i)
//      {
//        var currAccount = acctBalanceHistoryArray[i];
//        
//        if (isNaN(currAccount.id) || currAccount.bal === "NOT_FOUND")
//        {
//          Browser.msgBox("Skipping invalid account: " + JSON.stringify(currAccount));
//          continue;
//        }
//
//        if (currAccount.isHidden)
//        {
//          if (Debug.enabled) Debug.log("Not displaying hidden account '%s'", currAccount.name);
//          continue;
//        }
//
//        // Find column index of this account
//        var j = 0;
//        for (; j < numCols; ++j)
//        {
//          // if the account id at index 'j' is the one we're looking for
//          // or we've reached the end of the accounts (meaning this is a new account)
//          // then we have our position.
//          var acctId = hdrValues[Const.IDX_ACCT_ID][j];
//          if (acctId === currAccount.id || acctId === "" || isNaN(acctId) || acctId == null)
//          {
//            break;
//          }
//          // If this is the Mint demo data, then the account IDs probably won't match.
//          // We'll try matching the account name and financial institution instead.
//          if (currAccount.name === hdrValues[Const.IDX_ACCT_NAME][j]
//              && currAccount.fi === hdrValues[Const.IDX_ACCT_FINANCIAL_INST][j])
//          {
//            break;
//          }
//        }
//        
//        hdrValues[Const.IDX_ACCT_NAME][j] = currAccount.name;
//        hdrValues[Const.IDX_ACCT_FINANCIAL_INST][j] = currAccount.fi;
//        hdrValues[Const.IDX_ACCT_TYPE][j] = currAccount.klass;
//        hdrValues[Const.IDX_ACCT_ID][j] = currAccount.id;
//
//        var isDebtAcct = (currAccount.klass === "credit" || currAccount.klass === "loan");
//
//        // Copy balances for this account into the newValues array
//        // 
//        var acctIndex = j + 1; // Index 0 of newValues is the date column, so add 1
//        for (var k = 0; k < newBalanceCount; ++k) {
//          var balanceEntry = currAccount.balanceHistory[k];
//          newValues[k][acctIndex] = (isDebtAcct ? -balanceEntry.value : balanceEntry.value);
//        }
//      }
//
//      // Merge newValues with the existingValues.
//      if (existingValues != null) {
//        Sheets.AccountData.mergeBalanceValues(newValues, existingValues);
//      } else {
//        existingValues = newValues;
//      }
//
//      // Get column count and refresh ranges in case new accounts were added.
//      Debug.log(Utilities.formatString("Inserting account balances:  hdrValues.columns: %d, accountCount: %d, numCols: %d", hdrValues[0].length, acctCount, numCols));
//      hdrRange = hdrRange.offset(0, 0, hdrRange.getNumRows(), numCols);
//      existingRange = existingRange.offset(0, 0, existingValues.length, numCols + 1);
//
//      hdrRange.setValues(hdrValues);
//      existingRange.setValues(existingValues);
////      dateColRange.setValues(dateValues);
//    },
//

    //-----------------------------------------------------------------------------
    // Sheets.AccountData
    mergeBalanceValues: function(newDateValues, newBalValues, dateValues, balValues)
    {
      if (newBalValues == null || newBalValues.length === 0) {
        if (Debug.enabled) Debug.log("mergeBalanceValues: No new balances to merge.");
        return; // nothing to merge
      }

      // Overall algorithnm
      // 1. If new dates overlap with existing ones, then update the existing balances.
      // 2. Add new dates/balances to the end
      var mergeStartDate = Utils.find2dArrayMinValue(newDateValues, 0);
      var mergeEndDate = Utils.find2dArrayMaxValue(newDateValues, 0);

      // Loop through backwards so the index 'i' doesn't get messed up if we delete an entry
      var balLen = balValues.length;
      for (var i = balLen - 1; i >= 0; --i) {
        // Sanity check: make sure dateValues[i][0] contains an actual date. If it doesn't,
        // the call to getTime() will throw an exception.
        if (dateValues[i][0].getTime === undefined) {
          var warningMsg = "WARNING: Invalid date encountered when importing account balances. Delete the row with the invalid date.";
          if (Debug.enabled) Debug.log(warningMsg);
          toast(warningMsg, "Invalid date");
          dateValues[i][0] = new Date(0);
        }
        var date = dateValues[i][0].getTime();

        if (date >= mergeStartDate && date <= mergeEndDate) {
          // Check if we have a new account balance for an existing date.
          // If so, overwrite the old balance with the new value and remove
          // the new balance from the array.

          // Loop through backwards so the index 'i' doesn't get messed up if we delete an entry
          var newBalLen = newBalValues.length;
          for (var j = newBalLen - 1; j >= 0; --j) {
            if (date === newDateValues[j][0]) {
              if (Debug.traceEnabled) Debug.trace("mergeBalanceValues: Date match found");
              // We have a new balance for this date, we'll replace the existing
              // balance with the new one.
              balValues[i][0] = newBalValues[j][0];
              // Remove the new date/balance since we assigned it to an existing slot
              newBalValues.splice(j, 1);
              newDateValues.splice(j, 1);
              break;
            }
          }
        }
      }

      // 2. Add new balances
      var newBalLen = newBalValues.length;
      for (var i = 0; i < newBalLen; ++i) {
        balValues.push(newBalValues[i]);
        dateValues.push([new Date(newDateValues[i][0])]);
      }

      // 3. Sort existing txns in memory, by date ascending
      // Skipping sort. Account balances will be sorted all at once by the caller.
      //Utils.sort2dArray(existingValues, [0], [1]);
    },
    
    //-----------------------------------------------------------------------------
    // Sheets.AccountData
    mergeBalanceValues_notUsed: function(newDateValues, newBalValues, dateValues, balValues)
    {
      Debug.log("merge start");
            // 1. Sort existingValues and newValues by date
            var dateColIndex = 0;
            Utils.sort2dArray(balValues, [dateColIndex], [1]);
            Utils.sort2dArray(newBalValues, [dateColIndex], [1]);

            var numCols = balValues[0].length;
            var existingLen = balValues.length;
            var newLen = newBalValues.length;
            var e = 0, n = 0;
            while (e < existingLen && n < newLen) {
              var eDate = balValues[e][dateColIndex];
              var nDate = newBalValues[n][dateColIndex];

      Debug.log(Utilities.formatDate(nDate, "GMT", "MM/dd/yyyy"));
              if (eDate < nDate) {
                ++e;
              } else if (nDate < eDate) {
                balValues.push(newBalValues[n]);
                ++n;
              } else {
      Debug.log("match %d", n);
                // Copy newValues of current row over existingValues, but only if new value is a number
                for (var i = 0; i < numCols; ++i) {
                  var newVal = newBalValues[n][i];
                  if (typeof newVal == "number") {
                    balValues[e][i] = newVal;
                  }
                }
                ++e;
                ++n;
              }
            }

            for (; n < newLen; ++n) {
              balValues.push(newBalValues[n]);
            }

            Utils.sort2dArray(balValues, [dateColIndex], [1]);
      Debug.log("merge end");
    },

  },

  ////////////////////////////////////////////////////////////////////////////////////////////////
  Budget: {

    getSheet: function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getActiveSheet();
      var typeCell = sheet.getRange(Const.SHEET_TYPE_CELL);
      if (typeCell === null || typeCell.getValue() !== Const.SHEET_TYPE_BUDGET) {
        sheet = ss.getSheetByName(Const.SHEET_NAME_BUDGET);
      }

      if (Debug.enabled) Debug.log("Using Budget sheet: " + sheet.getName());
      return sheet;
    },

    //-----------------------------------------------------------------------------
    // Sheets.Budget
    getRangeByName: function(sheet, rangeName) {

      // Check the cache first
      const cacheKeyPrefix = Const.CACHE_NAMED_RANGE_PREFIX + String(sheet.getSheetId()) + ".";
      const cacheKeyName = cacheKeyPrefix + rangeName;
      const cache = Utils.getPrivateCache();
      let rangeA1Notation = cache.get(cacheKeyName);

      if (rangeA1Notation) {
        // Need to verify that the cached range is up to date before using it.
        const everythingElseRow = cache.get(cacheKeyPrefix + "EverythingElseRow") || 1;
        const everythingElseLabel = sheet.getRange(+everythingElseRow, 1).getValue();
        if (everythingElseLabel !== 'Everything else') {
          Debug.log('Cached budget ranges are out of date. Recreating them.');
          rangeA1Notation = null;
        }
      }

      if (!rangeA1Notation) {
        // Get first column
        var firstColRange = sheet.getRange(1, 1, sheet.getLastRow(), 1);
        var firstColValues = firstColRange.getValues();
  
        var budgetItemsStart = 0;
        var budgetItemsEnd = 0;
        var budgetOptionsStart = 0;

        var rowCount = firstColValues.length;
        for (var i = 0; i < rowCount; ++i) {
          var value = firstColValues[i][0];
          if (budgetItemsStart === 0 && value === "Budget Item Name") {
            budgetItemsStart = i + 3; // + 3 to get row num of first budget item
          }
          else if (budgetItemsEnd === 0 && value === "Everything else") {
            budgetItemsEnd = i + 1; // + 1 to get row num of "Everything else" row
          }
          else if (budgetOptionsStart === 0) {
            if (value === "Budget Options:") {
              budgetOptionsStart = i + 1;
            }
          }
          else {
            // We're done
            break;
          }
        }

        if (budgetItemsStart === 0 || budgetItemsEnd === 0 || budgetOptionsStart === 0) {
          throw new Error("Unable to determine Budget cell ranges");
        }

        // Get the various named ranges and cache them
        var tempRange = null;

        // BudgetItemsRange
        tempRange = sheet.getRange(budgetItemsStart, 1, budgetItemsEnd - budgetItemsStart + 1, Const.IDX_BUDGET_TXN_COUNT + 1);
        cache.put(cacheKeyPrefix + "BudgetItemsRange", tempRange.getA1Notation());
        // BudgetStartDate
        cache.put(cacheKeyPrefix + "BudgetStartDate", Const.BUDGET_CELL_START_DATE);
        // BudgetEndDate
        cache.put(cacheKeyPrefix + "BudgetEndDate", Const.BUDGET_CELL_END_DATE);
        // EverythingElseCell
        tempRange = sheet.getRange(budgetItemsEnd, Const.IDX_BUDGET_ACTUAL + 1, 1, 1);
        cache.put(cacheKeyPrefix + "EverythingElseCell", tempRange.getA1Notation());
        // EverythingElseRow
        tempRange = sheet.getRange(budgetItemsEnd, Const.IDX_BUDGET_AMOUNT + 1, 1, 1);
        cache.put(cacheKeyPrefix + "EverythingElseRow", String(budgetItemsEnd));
        // BudgetExcludeAccountsCell
        tempRange = sheet.getRange(budgetOptionsStart + 1, 2, 1, 1);
        cache.put(cacheKeyPrefix + "BudgetExcludeAccountsCell", tempRange.getA1Notation());
        // BudgetExcludeCategoriesCell
        tempRange = sheet.getRange(budgetOptionsStart + 2, 2, 1, 1);
        cache.put(cacheKeyPrefix + "BudgetExcludeCategoriesCell", tempRange.getA1Notation());

        if (Debug.enabled) Debug.log("Named ranges for sheet '%s' saved in cache", sheet.getName());

        // Fetch the requested named range from the cache
        rangeA1Notation = cache.get(cacheKeyName);
        if (!rangeA1Notation) {
          throw new Error("Named range '" + rangeName + "' not found for sheet '" + sheet.getName() + "'");
        }
      }
      else {
        if (Debug.enabled) Debug.log("Named range '%s' found in cache", cacheKeyName);
      }

      var range = sheet.getRange(rangeA1Notation);
      return range;
    },

    //-----------------------------------------------------------------------------
    // Sheets.Budget
    clearNamedRanges: function(sheet) {
      var cacheKeyPrefix = Const.CACHE_NAMED_RANGE_PREFIX + String(sheet.getSheetId()) + ".";
      var cache = Utils.getPrivateCache();
      cache.remove(cacheKeyPrefix + "BudgetItemsRange");
      cache.remove(cacheKeyPrefix + "BudgetStartDate");
      cache.remove(cacheKeyPrefix + "BudgetEndDate");
      cache.remove(cacheKeyPrefix + "EverythingElseCell");
      cache.remove(cacheKeyPrefix + "EverythingElseRow");
      cache.remove(cacheKeyPrefix + "BudgetExcludeAccountsCell");
      cache.remove(cacheKeyPrefix + "BudgetExcludeCategoriesCell");

      if (Debug.enabled) Debug.log("Named ranges for sheet '%s' (%s) removed from cache", sheet.getName(), sheet.getSheetId());
    },

    //-----------------------------------------------------------------------------
    // Sheets.Budget
    getTxnSortCriteria: function() {
      return [Const.IDX_TXN_MATCHES + 1, {column: Const.IDX_TXN_DATE + 1, ascending: false}];
    },

    //-----------------------------------------------------------------------------
    // Sheets.Budget
    updateCalculations: function(sheet) {

      toast("Recalculating Budgets", "Budget update");

      // Get data range for transaction data
      var trv = Sheets.TxnData.getTxnRangeAndValues();
      var txnDataRange = trv.txnRange;
      var txnData = trv.txnValues;
      var txnDataLen = txnData.length;
      var matchingTxnsRange = txnDataRange.offset(0, Const.IDX_TXN_MATCHES, txnDataLen, 1);
      var matchingTxnsArray = matchingTxnsRange.getValues();
      var txnColorArray = new Array(txnDataLen);
      var txnAmountIndex = Utils.getTxnAmountColumn() - 1;
      
      try
      {
        if (!sheet) {
          sheet = Sheets.Budget.getSheet();
        }

        // Get start and end dates to include in budget calculations
        var startDate = Sheets.Budget.getRangeByName(sheet, "BudgetStartDate").getValue();
        var endDate = Sheets.Budget.getRangeByName(sheet, "BudgetEndDate").getValue();
        
        // Clear txn matches
        for (var i = 0; i < txnDataLen; ++i) {
          matchingTxnsArray[i][0] = "";
        }
        
        for (var i = 0; i < txnDataLen; ++i) {
          txnColorArray[i] = Const.NO_COLOR;
        }
        
        // Get budget data
        var budgetRange = Sheets.Budget.getRangeByName(sheet, "BudgetItemsRange");
        var budgetData = budgetRange.getValues();
        var budgetDataLen = budgetData.length;
        var budgetActualSpentRange = budgetRange.offset(0, Const.IDX_BUDGET_ACTUAL, budgetRange.getNumRows(), 1);
        var budgetSums = budgetActualSpentRange.getValues();
        var matchingTxnCountRange = budgetRange.offset(0, Const.IDX_BUDGET_TXN_COUNT, budgetRange.getNumRows(), 1);
        var matchingTxnCounts = matchingTxnCountRange.getValues();
        // Clear the 'actual budget' and 'txn count' columns
        for (var i = 0; i < budgetDataLen; ++i) {
          budgetSums[i][0] = 0;
          matchingTxnCounts[i][0] = 0;
          if (budgetData[i][Const.IDX_BUDGET_NAME] == "") {
            // If there is no budget item for this row, set the budget sum and txn count to 0.
            budgetSums[i][0] = null;
            matchingTxnCounts[i][0] = null;
          }
        }
        budgetActualSpentRange.setValues(budgetSums);
        
        // Get index of the 'Everything else' row. It should always be the last row.
        var everythingElseIdx = budgetData.length - 1;
        if (budgetData[everythingElseIdx][Const.IDX_BUDGET_NAME] != "Everything else") {
          throw new Error("Unexpected error. The 'Everything else' item should be the last row in the 'BudgetItemsRange' named range!");
        }
        
        // Get budget highlight colors
        var budgetColorRange = budgetRange.offset(0, Const.IDX_BUDGET_COLOR, budgetRange.getNumRows(), 1);
        var budgetColorArray = budgetColorRange.getBackgrounds();
        
        // Get accounts to ignore
        var ignoreAccounts = String(Sheets.Budget.getRangeByName(sheet, "BudgetExcludeAccountsCell").getValue());
        var ignoreAccountsMap = (ignoreAccounts == null ? [] : Utils.convertDelimitedStringToArray(ignoreAccounts.toLowerCase(), Const.DELIM));
        
        // Get categories and tags to ignore
        var ignoreCategories = String(Sheets.Budget.getRangeByName(sheet, "BudgetExcludeCategoriesCell").getValue());
        var ignoreCategoriesMap = (ignoreCategories == null ? [] : Utils.convertDelimitedStringToArray(ignoreCategories.toLowerCase(), Const.DELIM));
        
        // Loop through the budgets and create map of categories to budgets
        var budgetCategoryMap = [];
        var budgetIncludeAllArray = [];
        
        for (var i = 0; i < budgetDataLen; ++i) {
          if (budgetData[i][Const.IDX_BUDGET_NAME].trim() === "")
            continue;

          var budgetItem = i;
          
          var categories = budgetData[i][Const.IDX_BUDGET_INCLUDE_CATEGORIES].toLowerCase().split(Const.DELIM);
          var includeAll = ("and" === budgetData[i][Const.IDX_BUDGET_INCLUDE_ANDOR].toLowerCase().trim() ? true : false);
          budgetIncludeAllArray[i] = 0;
          
          for (var j = 0; j < categories.length; ++j) {
            var cat = categories[j].trim();
            if (cat === "")
              continue;
            
            // Save category in map. Value is index of budget item.
            // Since multiple budget items can include the same category,
            // the map value is actually an array of budget item indexes.
            var mapEntry = budgetCategoryMap[cat];
            if (mapEntry == null) {
              mapEntry = new Array();
              budgetCategoryMap[cat] = mapEntry;
            }
            mapEntry.push(budgetItem);
            
            // If this is a "match all" budget item, keep track of the
            // total number of categories that must match. This works
            // because a txn will not have the same category / tag more than once.
            // (if the budget item only includes one category / tag, then "match all"
            // is unnecessary)
            if (includeAll && j === 0 && categories.length > 1) {
              budgetIncludeAllArray[i] = categories.length;
            }
          }
        }
        
        // Loop through the array of transactions
        for (var i = 0; i < txnDataLen; ++i) {
          var txnDate = txnData[i][Const.IDX_TXN_DATE];
          if (txnDate < startDate || txnDate > endDate) {
            continue;  // skip txns that are outside of the desired date range
          }

          var txnAmount = txnData[i][txnAmountIndex];
          if (isNaN(txnAmount)) {
            continue;  // Skip invalid amount
          }
          
          if (ignoreAccountsMap[String(txnData[i][Const.IDX_TXN_ACCOUNT]).toLowerCase()] != null) {
            continue;  // Ignore this account
          }

          var match = false;
          var ignore = false;
          matchingTxnsArray[i][0] = "";
          
          var catsAndTags = String(txnData[i][Const.IDX_TXN_CATEGORY]).toLowerCase() + Const.DELIM + String(txnData[i][Const.IDX_TXN_TAGS]).toLowerCase();
          var txnCatArray = catsAndTags.split(Const.DELIM);
          
          // Loop through the array of categories we are trying to match
          var matchingBudgetsArray = [];
          var budgetIncludeAllArrayCopy = [];
          
          for (var j = 0; j < txnCatArray.length; ++j) {
            
            var cat = txnCatArray[j];
            var ignoreCat = (ignoreCategoriesMap[cat] != null);
            
            // Should we ignore this category, and is this the
            // first "ignored category" we've encountered?
            if (ignoreCat && !ignore) {
              // Clear the array of matching budgets
              // Only budgets that include "ignored categories" will be included now.
              matchingBudgetsArray = [];
              ignore = true;
              match = false;
            }
            
            // Does txn category match a category we're looking for?
            var budgetArray = budgetCategoryMap[cat];
            if (budgetArray != null) {
              for (var k = 0; k < budgetArray.length; ++k) {
                var budget = budgetArray[k];
                
                var includeAll = (budgetIncludeAllArray[budget] > 0);
                if (includeAll) {
                  // Matching all categories for this budget is required.
                  var remainingMatchCount = budgetIncludeAllArrayCopy[budget];
                  if (remainingMatchCount == null) {
                    remainingMatchCount = budgetIncludeAllArray[budget];
                  }
                  --remainingMatchCount;
                  
                  Debug.assert(remainingMatchCount >= 0, "remainingMatchCount < 0");
                  
                  budgetIncludeAllArrayCopy[budget] = remainingMatchCount;
                  
                  // Have we matched all of the categories for this budget?
                  if (remainingMatchCount === 0) {
                    // txn matches budget item
                    matchingBudgetsArray[budget] = true;
                    match = true;
                  }
                } else {
                  if (ignore) {
                    if (ignoreCat) {
                      // This is an "ignored category" but it was explicitly included
                      // in a budget item, so the budget item will override.
                      matchingBudgetsArray[budget] = true;
                      match = true;
                    }
                    
                  } else {
                    
                    if (!includeAll) {
                      // Matching all is not required, so one match is good enough
                      matchingBudgetsArray[budget] = true;
                      match = true;
                      
                    }
                  }
                }
              } // for (k)
            } // if (budget != null)
          } // for (j)
          
          // If there were no matching categories found, then include this txn in
          // in the "everything else" total
          if (!match && !ignore) {
            matchingBudgetsArray[everythingElseIdx] = true;
          }
          
          // Add this txn amount to the budget items it matched
          for (var b = 0; b < matchingBudgetsArray.length; ++b) {
            if (matchingBudgetsArray[b] === true) {
              matchingTxnsArray[i][0] += budgetData[b][Const.IDX_BUDGET_NAME] + Const.DELIM;
              
              budgetSums[b][0] += -txnData[i][txnAmountIndex]; // Change sign of value so expenses are positive
              
              // Store budget color in txnColorArray while we're at it
              if (txnColorArray[i] === Const.NO_COLOR) {
                txnColorArray[i] = budgetColorArray[b][0];
              }
              
              // Increment the number of matching txns for this budget item
              ++matchingTxnCounts[b][0];
            }
          }
          
        } // for (i)
        
        // Set calculated values
        budgetActualSpentRange.setValues(budgetSums);
        matchingTxnCountRange.setValues(matchingTxnCounts);
        
        // Set colors for progress bars
        var budgetProgressRange = budgetRange.offset(0, Const.IDX_BUDGET_PERCENT_PROGRESS, budgetDataLen, 1);
        var budgetProgressColors = budgetProgressRange.getFontColors();
        
        for (var i = 0; i < budgetDataLen; ++i) {
          var actualPercent = 0.0;
          var budgetAmount = budgetData[i][Const.IDX_BUDGET_TOTAL];
          var actualAmount = budgetSums[i][0];
          var actualDiff = actualAmount - budgetAmount;
          if (budgetAmount > 0) {
            actualPercent = actualAmount / budgetAmount;
          }
          
          // Set the progress bar color (color index is from 0 to 10)
          //    var progress = Math.round((Math.max(-250, Math.min(250, actualDiff)) + 250) / 50);
          var percent = Math.min(200, actualPercent*100);
          var progress = 0;
          if (percent < 95) {
            progress = Math.round(Math.abs(percent) / 25);
          } else if (percent <= 101) {
            progress = 5;
          } else {
            progress = 6 + Math.round((percent - 100) / 25);
          }
          
          budgetProgressColors[i][0] = Const.BUDGET_PROGRESS_COLORS[progress];
        }
        budgetProgressRange.setFontColors(budgetProgressColors);

      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }
      finally
      {
        toast("Complete", "Budget update");
      }

      Sheets.TxnData.setTxnMatchesHeader(Const.TXN_MATCHES_BUDGET_HDR);
      matchingTxnsRange.setValues(matchingTxnsArray);
      SpreadsheetUtils.setRowColors(matchingTxnsRange, txnColorArray, false, Const.NO_COLOR, false);
    },
    
    //-----------------------------------------------------------------------------
    // Sheets.Budget
    insertBudgetItem: function() {
      var sheet = Sheets.Budget.getSheet();
      var budgetRange = Sheets.Budget.getRangeByName(sheet, "BudgetItemsRange");
      var sheet = budgetRange.getSheet();

      var lastRow = budgetRange.getLastRow() - 2; // Minus to get last budget item, skipping "Everything Else" and blue separator row
      var lastRowOffset = lastRow - budgetRange.getRow();
      var lastBudgetItemRange = budgetRange.offset(lastRowOffset, 0, 1, Const.IDX_BUDGET_TXN_COUNT + 1);

      // Insert new budget row at the end
      sheet.insertRowAfter(lastRow);
      // Clear the named ranges for this budget sheet so they will be re-built
      Sheets.Budget.clearNamedRanges(sheet);

      // Copy previous budget item to new row (so formatting, formulas, etc. are copied)
      var newRowRange = budgetRange.offset(lastRowOffset + 1, 0, 1, Const.IDX_BUDGET_TXN_COUNT + 1);
      lastBudgetItemRange.copyTo(newRowRange, {contentsOnly:false});

      // Clear editable fields of new budget item
      var budgetCell = newRowRange.offset(0, Const.IDX_BUDGET_NAME, 1, 1);
      // Clear budget item name
      budgetCell.setValue("");
      budgetCell.clearNote();
      budgetCell.activate();
      // Set highlight color to white
      budgetCell = newRowRange.offset(0, Const.IDX_BUDGET_COLOR, 1, 1);
      budgetCell.setBackground(Const.NO_COLOR);
      // Set default budget amount to 1 to avoid divide-by-zero errors in computed cells
      budgetCell = newRowRange.offset(0, Const.IDX_BUDGET_AMOUNT, 1, 1);
      budgetCell.setValue(1);
      // Set default frequency period to 'M' (monthly)
      budgetCell = newRowRange.offset(0, Const.IDX_BUDGET_FREQ, 1, 1);
      budgetCell.setValue("M");
      // Clear budget categories and tags
      budgetCell = newRowRange.offset(0, Const.IDX_BUDGET_INCLUDE_CATEGORIES, 1, 1);
      budgetCell.setValue("");
      // Set AND/OR field to OR
      budgetCell = newRowRange.offset(0, Const.IDX_BUDGET_INCLUDE_ANDOR, 1, 1);
      budgetCell.setValue("OR");
      // Clear actual amount
      budgetCell = newRowRange.offset(0, Const.IDX_BUDGET_ACTUAL, 1, 1);
      budgetCell.setValue(null);
      // Clear txn count
      budgetCell = newRowRange.offset(0, Const.IDX_BUDGET_TXN_COUNT, 1, 1);
      budgetCell.setValue(null);
    },

  }, // Budget

  ////////////////////////////////////////////////////////////////////////////////////////////////
  InOut: {

    getSheet: function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getActiveSheet();
      var typeCell = sheet.getRange(Const.SHEET_TYPE_CELL);
      if (typeCell === null || typeCell.getValue() !== Const.SHEET_TYPE_INOUT) {
        sheet = ss.getSheetByName(Const.SHEET_NAME_INOUT);
      }

      if (Debug.enabled) Debug.log("Using In/Out sheet: " + sheet.getName());
      return sheet;
    },

    //-----------------------------------------------------------------------------
    // Sheets.InOut
    getRangeByName: function(sheet, rangeName) {

      // Check the cache first
      var cacheKeyPrefix = Const.CACHE_NAMED_RANGE_PREFIX + String(sheet.getSheetId()) + ".";
      var cacheKeyName = cacheKeyPrefix + rangeName;
      var cache = Utils.getPrivateCache();
      var rangeA1Notation = cache.get(cacheKeyName);

      if (!rangeA1Notation) {
        // Get first column
        var firstColRange = sheet.getRange(1, 1, sheet.getLastRow(), 1);
        var firstColValues = firstColRange.getValues();
  
        var inoutIncomeRow = 0;
        var inoutExpensesRow = 0;
        var inoutOptionsStart = 0;

        var rowCount = firstColValues.length;
        for (var i = 0; i < rowCount; ++i) {
          var value = firstColValues[i][0];
          if (inoutIncomeRow === 0 && value === "Income") {
            inoutIncomeRow = i + 1; // + 1 to get row num
          }
          else if (inoutExpensesRow === 0 && value === "Expenses") {
            inoutExpensesRow = i + 1; // + 1 to get row num
          }
          else if (inoutOptionsStart === 0) {
            if (value === "Options:") {
              inoutOptionsStart = i + 1;
            }
          }
          else {
            // We're done
            break;
          }
        }

        if (inoutIncomeRow === 0 || inoutExpensesRow === 0 || inoutOptionsStart === 0) {
          throw "Unable to determine In/Out cell ranges";
        }

        // Get the various named ranges and cache them
        var tempRange = null;

        // InOutIncomeCell
        tempRange = sheet.getRange(inoutIncomeRow, 2, 1, 1);
        cache.put(cacheKeyPrefix + "InOutIncomeCell", tempRange.getA1Notation());
        // InOutExpensesCell
        tempRange = sheet.getRange(inoutExpensesRow, 2, 1, 1);
        cache.put(cacheKeyPrefix + "InOutExpensesCell", tempRange.getA1Notation());
        // InOutStartDate
        cache.put(cacheKeyPrefix + "InOutStartDate", Const.INOUT_CELL_START_DATE);
        // InOutEndDate
        cache.put(cacheKeyPrefix + "InOutEndDate", Const.INOUT_CELL_END_DATE);
        // InOutExcludeAccountsCell
        tempRange = sheet.getRange(inoutOptionsStart + 1, 2, 1, 1);
        cache.put(cacheKeyPrefix + "InOutExcludeAccountsCell", tempRange.getA1Notation());
        // InOutExcludeCategoriesCell
        tempRange = sheet.getRange(inoutOptionsStart + 2, 2, 1, 1);
        cache.put(cacheKeyPrefix + "InOutExcludeCategoriesCell", tempRange.getA1Notation());

        if (Debug.enabled) Debug.log("Named ranges for sheet '%s' saved in cache", sheet.getName());

        // Fetch the requested named range from the cache
        rangeA1Notation = cache.get(cacheKeyName);
        if (!rangeA1Notation) {
          throw "Named range '" + rangeName + "' not found for sheet '" + sheet.getName() + "'";
        }
      }
      else {
        if (Debug.enabled) Debug.log("Named range '%s' found in cache", cacheKeyName);
      }

      var range = sheet.getRange(rangeA1Notation);
      return range;
    },

    //-----------------------------------------------------------------------------
    // Sheets.InOut
    clearNamedRanges: function(sheet) {
      var cacheKeyPrefix = Const.CACHE_NAMED_RANGE_PREFIX + String(sheet.getSheetId()) + ".";
      var cache = Utils.getPrivateCache();
      cache.remove(cacheKeyPrefix + "InOutIncomeCell");
      cache.remove(cacheKeyPrefix + "InOutExpensesCell");
      cache.remove(cacheKeyPrefix + "InOutStartDate");
      cache.remove(cacheKeyPrefix + "InOutEndDate");
      cache.remove(cacheKeyPrefix + "InOutExcludeAccountsCell");
      cache.remove(cacheKeyPrefix + "InOutExcludeCategoriesCell");      

      if (Debug.enabled) Debug.log("Named ranges for sheet '%s' (%s) removed from cache", sheet.getName(), sheet.getSheetId());
    },

    //-----------------------------------------------------------------------------
    // Sheets.InOut
    getTxnSortCriteria: function() {
      var txnAmountCol = Utils.getTxnAmountColumn();
      return [Const.IDX_TXN_MATCHES + 1, {column: txnAmountCol, ascending: true}, {column: Const.IDX_TXN_DATE + 1, ascending: false}];
    },

    //-----------------------------------------------------------------------------
    // Sheets.InOut
    updateCalculations: function(sheet) {
      
      toast("Recalculating inflows and outflows", "In / Out update");

      // Get data range for transaction data
      var txnDataRange = Utils.getTxnDataRange();
      var txnData = txnDataRange.getValues();
      var txnDataLen = txnData.length;
      var matchingTxnsRange = txnDataRange.offset(0, Const.IDX_TXN_MATCHES, txnDataLen, 1);
      var matchingTxnsArray = matchingTxnsRange.getValues();
      var txnColorArray = new Array(txnDataLen);
      var txnAmountIndex = Utils.getTxnAmountColumn() - 1;

      try
      {
        if (!sheet) {
          sheet = Sheets.InOut.getSheet();
        }

        // Get start and end dates to include in in/out calculations
        var startDate = Sheets.InOut.getRangeByName(sheet, "InOutStartDate").getValue();
        var endDate = Sheets.InOut.getRangeByName(sheet, "InOutEndDate").getValue();

        // Clear txn matches
        for (var i = 0; i < txnDataLen; ++i) {
          matchingTxnsArray[i][0] = "";
        }

        for (var i = 0; i < txnDataLen; ++i) {
          txnColorArray[i] = Const.NO_COLOR;
        }
        
        var incomeSum = 0;
        var expenseSum = 0;
        
        // Get accounts to ignore
        var ignoreAccounts = String(Sheets.InOut.getRangeByName(sheet, "InOutExcludeAccountsCell").getValue());
        var ignoreAccountsMap = (ignoreAccounts == null ? [] : Utils.convertDelimitedStringToArray(ignoreAccounts.toLowerCase(), Const.DELIM));

        // Get categories and tags to ignore
        var ignoreCategories = String(Sheets.InOut.getRangeByName(sheet, "InOutExcludeCategoriesCell").getValue());
        var ignoreCategoriesMap = (ignoreCategories == null ? [] : Utils.convertDelimitedStringToArray(ignoreCategories.toLowerCase(), Const.DELIM));
        
        // Loop through the array of transactions
        for (var i = 0; i < txnDataLen; ++i) {
          var txnDate = new Date(txnData[i][Const.IDX_TXN_DATE]);
          if (txnDate < startDate || txnDate > endDate)
            continue;  // skip txns that are outside of the desired date range
          
          var txnAmount = txnData[i][txnAmountIndex];
          if (isNaN(txnAmount)) {
            continue;  // Skip invalid amount
          }
          
          if (ignoreAccountsMap[String(txnData[i][Const.IDX_TXN_ACCOUNT]).toLowerCase()] != null) {
            continue;  // Ignore this account
          }
          
          var catsAndTags = String(txnData[i][Const.IDX_TXN_CATEGORY]).toLowerCase() + Const.DELIM + String(txnData[i][Const.IDX_TXN_TAGS]).toLowerCase();
          var txnCatArray = catsAndTags.split(Const.DELIM);
          var ignore = false;
          
          // Loop through the array of categories we are excluding.
          // If this txn's category / tag matches, we'll skip it.
          for (var j = 0; j < txnCatArray.length; ++j) {
            
            var cat = txnCatArray[j];
            if (ignoreCategoriesMap[cat] != null) {
              ignore = true;
              break;
            }
            
          } // for (j)
          
          if (ignore)
            continue;
          
          if (txnAmount > 0) {
            incomeSum += txnAmount;
            matchingTxnsArray[i][0] = "Income";
            txnColorArray[i] = Const.COLOR_POSITIVE;
          } else {
            expenseSum += -txnAmount;
            matchingTxnsArray[i][0] = "Expense";
            txnColorArray[i] = Const.COLOR_NEGATIVE;
          }
        } // for (i)

        var incomeCell = Sheets.InOut.getRangeByName(sheet, "InOutIncomeCell");
        incomeCell.setValue(incomeSum);
        var expenseCell = Sheets.InOut.getRangeByName(sheet, "InOutExpensesCell");
        expenseCell.setValue(expenseSum);
      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }
      finally
      {
        toast("Complete", "In / Out update");
      }

      Sheets.TxnData.setTxnMatchesHeader(Const.TXN_MATCHES_INOUT_HDR);
      matchingTxnsRange.setValues(matchingTxnsArray);
      SpreadsheetUtils.setRowColors(matchingTxnsRange, txnColorArray, false, Const.NO_COLOR, false);
    },

  }, // InOut

  ////////////////////////////////////////////////////////////////////////////////////////////////
  SavingsGoal: {

    getSheet: function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getActiveSheet();
      var typeCell = sheet.getRange(Const.SHEET_TYPE_CELL);
      if (typeCell === null || typeCell.getValue() !== Const.SHEET_TYPE_SAVINGS_GOALS) {
        sheet = ss.getSheetByName(Const.SHEET_NAME_SAVINGS_GOALS);
      }
      return sheet;
    },

    //-----------------------------------------------------------------------------
    // Sheets.SavingsGoal
    getRangeByName: function(sheet, rangeName) {

      // Check the cache first
      var cacheKeyPrefix = Const.CACHE_NAMED_RANGE_PREFIX + String(sheet.getSheetId()) + ".";
      var cacheKeyName = cacheKeyPrefix + rangeName;
      var cache = Utils.getPrivateCache();
      var rangeA1Notation = cache.get(cacheKeyName);

      if (rangeA1Notation) {
        // Need to verify that the cached range is up to date before using it.
        const optionsRow = cache.get(cacheKeyPrefix + "GoalsIncludeAccountsLabelRow") || 1;
        const optionsLabel = sheet.getRange(+optionsRow, 1).getValue();
        if (optionsLabel !== 'Savings Goal Options:') {
          Debug.log('Cached savings goal ranges are out of date. Recreating them.');
          rangeA1Notation = null;
        }
      }

      if (!rangeA1Notation) {
        // Get first column
        var firstColRange = sheet.getRange(1, 1, sheet.getLastRow(), 1);
        var firstColValues = firstColRange.getValues();
  
        var goalsStart = 0;
        var goalsEnd = 0;
        var goalsOptionsStart = 0;
  
        var rowCount = firstColValues.length;
        for (var i = 0; i < rowCount; ++i) {
          var value = firstColValues[i][0];
          if (goalsStart === 0 && value === "Savings Goal Name") {
            goalsStart = i + 3; // + 3 to get row num of first Savings Goal item
          }
          else if (goalsEnd === 0 || goalsOptionsStart === 0) {
            if (value === "Savings Goal Options:") {
              goalsEnd = i - 1; // - 1 to get last Savings Goal row
              goalsOptionsStart = i + 1;
            }
          }
          else {
            // We're done
            break;
          }
        }

        if (goalsStart === 0 || goalsEnd === 0 || goalsOptionsStart === 0) {
          throw "Unable to determine Savings Goal cell ranges";
        }

        // Get the various named ranges and cache them
        var tempRange = null;

        // GoalsRange
        tempRange = sheet.getRange(goalsStart, 1, goalsEnd - goalsStart + 1, Const.IDX_GOAL_CREATE_DATE + 1);
        cache.put(cacheKeyPrefix + "GoalsRange", tempRange.getA1Notation());
        // GoalsIncludeAccountsCell
        tempRange = sheet.getRange(goalsOptionsStart + 1, 2, 1, 1);
        cache.put(cacheKeyPrefix + "GoalsIncludeAccountsCell", tempRange.getA1Notation());
        cache.put(cacheKeyPrefix + "GoalsIncludeAccountsLabelRow", String(goalsOptionsStart));

        if (Debug.enabled) Debug.log("Named ranges for sheet '%s' saved in cache", sheet.getName());

        // Fetch the requested named range from the cache
        rangeA1Notation = cache.get(cacheKeyName);
        if (!rangeA1Notation) {
          throw "Named range '" + rangeName + "' not found for sheet '" + sheet.getName() + "'";
        }
      }
      else {
        if (Debug.enabled) Debug.log("Named range '%s' found in cache", cacheKeyName);
      }

      var range = sheet.getRange(rangeA1Notation);
      return range;
    },

    //-----------------------------------------------------------------------------
    // Sheets.SavingsGoal
    clearNamedRanges: function(sheet) {
      var cacheKeyPrefix = Const.CACHE_NAMED_RANGE_PREFIX + String(sheet.getSheetId()) + ".";
      var cache = Utils.getPrivateCache();
      cache.remove(cacheKeyPrefix + "GoalssRange");
      cache.remove(cacheKeyPrefix + "GoalsIncludeAccountsCell");

      if (Debug.enabled) Debug.log("Named ranges for sheet '%s' (%s) removed from cache", sheet.getName(), sheet.getSheetId());
    },

    //-----------------------------------------------------------------------------
    // Sheets.SavingsGoal
    getTxnSortCriteria: function() {
      return [Const.IDX_TXN_MATCHES + 1, {column: Const.IDX_TXN_DATE + 1, ascending: false}];
    },

    //-----------------------------------------------------------------------------
    // Sheets.SavingsGoal
    updateCalculations: function(sheet) {
      
      toast("Recalculating Savings Goals", "Savings goal update", 60);

      // Get data range for transaction data
      var trv = Sheets.TxnData.getTxnRangeAndValues();
      var txnDataRange = trv.txnRange;
      var txnData = trv.txnValues;
      var txnDataLen = txnData.length;
      var matchingTxnsRange = txnDataRange.offset(0, Const.IDX_TXN_MATCHES, txnDataLen, 1);
      var matchingTxnsArray = matchingTxnsRange.getValues();
      var txnColorArray = new Array(txnDataLen);
      var txnAmountIndex = Utils.getTxnAmountColumn() - 1;
      
      try
      {
        if (!sheet) {
          sheet = Sheets.SavingsGoal.getSheet();
        }

        var now = new Date();

        // Get accounts to include
        var includeAccounts = String(Sheets.SavingsGoal.getRangeByName(sheet, "GoalsIncludeAccountsCell").getValue());
        var includeAccountsMap = (includeAccounts == null ? [] : Utils.convertDelimitedStringToArray(includeAccounts.toLowerCase(), Const.DELIM));

        // Make sure at least one account was specified; otherwise, no txns will be found for savings goals
        if (Object.keys(includeAccountsMap).length == 0) {
          Sheets.SavingsGoal.getRangeByName(sheet, "GoalsIncludeAccountsCell").activate();
          Browser.msgBox("You must specify at least one account in the \"Accounts to include\" field.\n\nThis should be accounts that you are DEDUCTING funds from to save towards the goal, such as a checking account");
          return;
        }
        
        // Clear txn matches
        for (var i = 0; i < txnDataLen; ++i) {
          matchingTxnsArray[i][0] = "";
        }
        
        for (var i = 0; i < txnDataLen; ++i) {
          txnColorArray[i] = Const.NO_COLOR;
        }
        
        // Get goal data
        var goalRange = Sheets.SavingsGoal.getRangeByName(sheet, "GoalsRange");
        var goalData = goalRange.getValues();
        var goalDataLen = goalData.length;
        var goalActualRange = goalRange.offset(0, Const.IDX_GOAL_ACTUAL, goalDataLen, 1);
        var goalSums = goalActualRange.getValues();
        var goalTxnCountRange = goalRange.offset(0, Const.IDX_GOAL_TXN_COUNT, goalDataLen, 1);
        var matchingTxnCounts = goalTxnCountRange.getValues();
        // Clear the 'actual goal' and 'txn count' columns. Also set Created Date if it is empty.
        for (var i = 0; i < goalDataLen; ++i) {
          if (!goalData[i][Const.IDX_GOAL_NAME]) {
            // If there is no goal for this row, set the goal sum and txn count to 0.
            goalSums[i][0] = null;
            matchingTxnCounts[i][0] = null;
          } else {
            var carryFwdAmount = goalData[i][Const.IDX_GOAL_CARRY_FWD];
            goalSums[i][0] = (isNaN(carryFwdAmount) || carryFwdAmount == null || carryFwdAmount == "" ? 0 : Number(carryFwdAmount));
            matchingTxnCounts[i][0] = 0;
            
            // If create date hasn't been set, then set it to today.
            if (goalData[i][Const.IDX_GOAL_CREATE_DATE] == "") {
              goalRange.offset(i, Const.IDX_GOAL_CREATE_DATE, 1, 1).setValue(now);
            }
          }
        }
        goalActualRange.setValues(goalSums);
        
        // Get goal highlight colors
        var goalColorRange = goalRange.offset(0, Const.IDX_GOAL_COLOR, goalDataLen, 1);
        var goalColorArray = goalColorRange.getBackgrounds();
        
        // Get categories and tags to ignore
        // (Currently disabled. It doesn't seem useful.)
        var ignoreCategories = "";//String(Sheets.SavingsGoal.getRangeByName(sheet, "GoalsExcludeCategoriesCell").getValue());
        var ignoreCategoriesMap = (ignoreCategories == null ? [] : Utils.convertDelimitedStringToArray(ignoreCategories.toLowerCase(), Const.DELIM));

        // Loop through the goals and create map of categories to goals
        var goalCategoryMap = [];
        var goalIncludeAllArray = [];
        
        for (var i = 0; i < goalDataLen; ++i) {
          if (goalData[i][Const.IDX_GOAL_NAME].trim() === "")
            continue;
          
          var goalIndex = i;
          
          var categories = goalData[i][Const.IDX_GOAL_INCLUDE_CATEGORIES].toLowerCase().split(Const.DELIM);
          var includeAll = ("and" === goalData[i][Const.IDX_GOAL_INCLUDE_ANDOR].toLowerCase().trim() ? true : false);
          goalIncludeAllArray[i] = 0;
          
          for (var j = 0; j < categories.length; ++j) {
            var cat = categories[j].trim();
            if (cat === "")
              continue;
            
            // Save category in map. Value is index of goal.
            // Since multiple goals can include the same category,
            // the map value is actually an array of goal indexes.
            var mapEntry = goalCategoryMap[cat];
            if (mapEntry == null) {
              mapEntry = new Array();
              goalCategoryMap[cat] = mapEntry;
            }
            mapEntry.push(goalIndex);
            
            // If this is a "match all" goal, keep track of the
            // total number of categories that must match. This works
            // because a txn will not have the same category / tag more than once.
            // (if the goal only includes one category / tag, then "match all"
            // is unnecessary)
            if (includeAll && j === 0 && categories.length > 1) {
              goalIncludeAllArray[i] = categories.length;
            }
          }
        }

        // Loop through the array of transactions
        for (var i = 0; i < txnDataLen; ++i) {

          var txnAmount = txnData[i][txnAmountIndex];
          if (isNaN(txnAmount)) {
            continue;  // Skip invalid amount
          }

          if (includeAccountsMap[String(txnData[i][Const.IDX_TXN_ACCOUNT]).toLowerCase()] == null) {
            continue;  // This account is not in the "include" list. Skip it.
          }
          
          var match = false;
          var ignore = false;
          matchingTxnsArray[i][0] = "";
          
          var catsAndTags = String(txnData[i][Const.IDX_TXN_CATEGORY]).toLowerCase() + Const.DELIM + String(txnData[i][Const.IDX_TXN_TAGS]).toLowerCase();
          var txnCatArray = catsAndTags.split(Const.DELIM);
          
          // Loop through the array of categories we are trying to match
          var matchingGoalsArray = [];
          var goalIncludeAllArrayCopy = [];
          
          for (var j = 0; j < txnCatArray.length; ++j) {
            
            var cat = txnCatArray[j];
            var ignoreCat = (ignoreCategoriesMap[cat] != null);
            
            // Should we ignore this category, and is this the
            // first "ignored category" we've encountered?
            if (ignoreCat && !ignore) {
              // Clear the array of matching goals
              // Only goals that include "ignored categories" will be included now.
              matchingGoalsArray = [];
              ignore = true;
              match = false;
            }
            
            // Does txn category match a category we're looking for?
            var goalArray = goalCategoryMap[cat];
            if (goalArray != null) {
              for (var k = 0; k < goalArray.length; ++k) {
                var goal = goalArray[k];
                
                var includeAll = (goalIncludeAllArray[goal] > 0);
                if (includeAll) {
                  // Matching all categories for this goal is required.
                  var remainingMatchCount = goalIncludeAllArrayCopy[goal];
                  if (remainingMatchCount == null) {
                    remainingMatchCount = goalIncludeAllArray[goal];
                  }
                  --remainingMatchCount;
                  
                  Debug.assert(remainingMatchCount >= 0, "remainingMatchCount < 0");
                  
                  goalIncludeAllArrayCopy[goal] = remainingMatchCount;
                  
                  // Have we matched all of the categories for this goal?
                  if (remainingMatchCount === 0) {
                    // txn matches goal
                    matchingGoalsArray[goal] = true;
                    match = true;
                  }
                } else {
                  if (ignore) {
                    if (ignoreCat) {
                      // This is an "ignored category" but it was explicitly included
                      // in a goal, so the goal will override.
                      matchingGoalsArray[goal] = true;
                      match = true;
                    }
                    
                  } else {
                    
                    if (!includeAll) {
                      // Matching all is not required, so one match is good enough
                      matchingGoalsArray[goal] = true;
                      match = true;
                    }
                  }
                }
              } // for (k)
            } // if (goal != null)
          } // for (j)
          
          // Add this txn amount to the goals it matched
          for (var b = 0; b < matchingGoalsArray.length; ++b) {
            if (matchingGoalsArray[b] === true) {
              matchingTxnsArray[i][0] += goalData[b][Const.IDX_GOAL_NAME] + Const.DELIM;
              
              goalSums[b][0] += -txnData[i][txnAmountIndex]; // Change sign of value so expenses are positive
              
              // Store goal color in txnColorArray while we're at it
              if (txnColorArray[i] === Const.NO_COLOR) {
                txnColorArray[i] = goalColorArray[b][0];
              }
              
              // Increment the number of matching txns for this goal
              ++matchingTxnCounts[b][0];
            }
            
          }
          
        } // for (i)
        
        // Set calculated values
        goalActualRange.setValues(goalSums);
        goalTxnCountRange.setValues(matchingTxnCounts);
        
        // Set Time Remaining values and colors, and also colors for progress bars
        var goalTimeLeftRange = goalRange.offset(0, Const.IDX_GOAL_TIME_LEFT, goalDataLen, 1);
        var goalTimeLeftArray = goalTimeLeftRange.getValues();
        var goalTimeLeftColors = goalTimeLeftRange.getBackgrounds();
        var goalProgressRange = goalRange.offset(0, Const.IDX_GOAL_PROGRESS, goalDataLen, 1);
        var goalProgressColors = goalProgressRange.getFontColors();
        
        for (var i = 0; i < goalDataLen; ++i) {
          var actualPercent = 0.0;
          var endDate = goalData[i][Const.IDX_GOAL_END_DATE];
          var goalAmount = goalData[i][Const.IDX_GOAL_AMOUNT];
          if (goalAmount > 0) {
            var actualAmount = goalSums[i][0];
            actualPercent = actualAmount / goalAmount;
            
            // Calculate the time remaining for this goal (in human readable units)
            var dateDiff = Utils.getHumanFriendlyDateDiff(now, endDate, "*");
            goalTimeLeftArray[i][0] = Utilities.formatString("%s %s", String(dateDiff.diff), dateDiff.unit);
            
          } else {
            goalTimeLeftArray[i][0] = null;
          }
          
          var createDate = goalData[i][Const.IDX_GOAL_CREATE_DATE];
          var daysLeft = Math.round((endDate - now) / Const.ONE_DAY_IN_MILLIS);
          var daysTotal = Math.round((endDate - createDate) / Const.ONE_DAY_IN_MILLIS);
          // If savings goal has not been met, then show some "heat" as the end date draws nearer
          var timeLeftColorIndex = 10;
          if (actualPercent > 0.0 && actualPercent < 1.0) {
            timeLeftColorIndex = Math.round(Math.max(0, Math.min(daysLeft / daysTotal * 10, 10)));
            if (Debug.enabled) Debug.log("Days remaining for goal '%s': %s. Color index: %s", goalData[i][Const.IDX_GOAL_NAME], daysLeft, timeLeftColorIndex);
          }
          goalTimeLeftColors[i][0] = Const.TIME_LEFT_COLORS[timeLeftColorIndex];
          
          // Set the progress bar color
          var progress = Math.round(Math.max(0, Math.min(actualPercent * 20, 20)));
          goalProgressColors[i][0] = Const.GOAL_PROGRESS_COLORS[progress];
        }
        goalTimeLeftRange.setValues(goalTimeLeftArray);
        goalTimeLeftRange.setBackgrounds(goalTimeLeftColors);
        goalProgressRange.setFontColors(goalProgressColors);

      }
      catch (e)
      {
        Debug.log(Debug.getExceptionInfo(e));
        Browser.msgBox("Error: " + e.toString());
      }
      finally
      {
        toast("Complete", "Savings goal update");
      }

      Sheets.TxnData.setTxnMatchesHeader(Const.TXN_MATCHES_GOAL_HDR);
      matchingTxnsRange.setValues(matchingTxnsArray);
      SpreadsheetUtils.setRowColors(matchingTxnsRange, txnColorArray, false, Const.NO_COLOR, false);
    },

    //-----------------------------------------------------------------------------
    insertSavingsGoal: function() {
      var sheet = Sheets.SavingsGoal.getSheet();
      var goalRange = Sheets.SavingsGoal.getRangeByName(sheet, "GoalsRange");

      var lastRow = goalRange.getLastRow();
      var lastRowOffset = lastRow - goalRange.getRow();
      var lastGoalItemRange = goalRange.offset(lastRowOffset, 0, 1, Const.IDX_GOAL_CREATE_DATE + 1);

      // Insert new goal row at the end
      sheet.insertRowAfter(lastRow);
      // Clear the named ranges for this Savings Goal sheet so they will be re-built
      Sheets.SavingsGoal.clearNamedRanges(sheet);

      // Copy previous goal to new row (so formatting, formulas, etc. are copied)
      var newRowRange = goalRange.offset(lastRowOffset + 1, 0, 1, Const.IDX_GOAL_CREATE_DATE + 1);
      lastGoalItemRange.copyTo(newRowRange, {contentsOnly:false});

      // Clear editable fields of new goal
      var goalCell = newRowRange.offset(0, Const.IDX_GOAL_NAME, 1, 1);
      // Clear goal name
      goalCell.setValue("");
      goalCell.clearNote();
      goalCell.activate();
      // Clear goal end date
      goalCell = newRowRange.offset(0, Const.IDX_GOAL_END_DATE, 1, 1);
      goalCell.setValue("");
      // Set highlight color to white
      goalCell = newRowRange.offset(0, Const.IDX_GOAL_COLOR, 1, 1);
      goalCell.setBackground(Const.NO_COLOR);
      // Set initial goal amount to 0
      goalCell = newRowRange.offset(0, Const.IDX_GOAL_AMOUNT, 1, 1);
      goalCell.setValue(0);
      // Clear goal categories and tags
      goalCell = newRowRange.offset(0, Const.IDX_GOAL_INCLUDE_CATEGORIES, 1, 1);
      goalCell.setValue("");
      // Set AND/OR field to AND
      goalCell = newRowRange.offset(0, Const.IDX_GOAL_INCLUDE_ANDOR, 1, 1);
      goalCell.setValue("AND");
      // Clear carry forward amount
      goalCell = newRowRange.offset(0, Const.IDX_GOAL_CARRY_FWD, 1, 1);
      goalCell.setValue(null);
      // Set create date to today
      goalCell = newRowRange.offset(0, Const.IDX_GOAL_CREATE_DATE, 1, 1);
      goalCell.setValue(new Date);
      // Clear current savings
      goalCell = newRowRange.offset(0, Const.IDX_GOAL_ACTUAL, 1, 1);
      goalCell.setValue(null);
      // Clear time remaining
      goalCell = newRowRange.offset(0, Const.IDX_GOAL_TIME_LEFT, 1, 1);
      goalCell.setValue(null);
      // Clear txn count
      goalCell = newRowRange.offset(0, Const.IDX_GOAL_TXN_COUNT, 1, 1);
      goalCell.setValue(null);
    },
  }, // SavingsGoal

};
