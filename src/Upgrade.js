'use strict';
/*
 * Copyright (c) 2013-2018 b3devs@gmail.com
 * MIT License: https://spdx.org/licenses/MIT.html
 */

import {Const} from './Constants.js';
import {Debug} from './Debug.js';
import {Utils, Settings, toast} from './Utils.js';


export const Upgrade = {

  autoUpgradeMojitoIfApplicable : function() {
    // Get spreadsheet version from Settings sheet
    const ssVer = String(Settings.getInternalSetting(Const.IDX_INT_SETTING_MOJITO_VERSION));
    // Get MojitoLib version from constants
    const libVer = String(Const.CURRENT_MOJITO_VERSION);

    if (ssVer === libVer) {
      return libVer;
    }

    const newVer = Upgrade.upgradeMojito(ssVer, libVer);
    return newVer;
  },

  upgradeMojito : function(fromVer, toVer) {

    if (Debug.enabled) Debug.log("Attempting to upgrade Mojito from version %s to %s", fromVer, toVer);
    toast(Utilities.formatString("Upgrading Mojito from version %s to version %s", fromVer, toVer), "Mojito upgrade");

    if (this.compareVersions(fromVer, toVer) > 0) {
      // Spreadsheet version is greater than MojitoLib ver. We don't support downgrading.
      Debug.log("upgradeMojito - fromVer (%s) is greater than toVer (%s). Aborting.", fromVer, toVer);
      Browser.msgBox("Internal version mismatch", "Your Mojito spreadsheet version (Settings sheet) is greater than the MojitoLib version it is using. This kind of version mismatch is not supported. Please download a new copy of Mojito to fix this problem.", Browser.Buttons.OK);
      return toVer; // Return the lesser of the two versions
    }

    let newVer = fromVer;
    let upgradeFailed = false;

    try
    {
  
      //TODO: Finish implementing this function
      while (newVer !== toVer && !upgradeFailed) {

        switch (newVer)
        {
        case "1.0.0.1":
          newVer = this.upgradeFrom_1_0_0_1(newVer);
          break;
  
        case "1.1.0":
        case "1.1.1":
          newVer = this.upgradeFrom_1_1_0(newVer);
          break;
  
        case "1.1.2":
          newVer = this.upgradeFrom_1_1_2(newVer);
          break;
          
        case "1.1.2.1":
        case "1.1.2.2":
        case "1.1.2.3":
        case "1.1.2.4":
          newVer = this.upgradeFrom_1_1_2_1(newVer);
          break;
  
        case "1.1.2.5":
          newVer = this.upgradeFrom_1_1_2_5(newVer);
          break;
  
        case "1.1.2.6":
          newVer = this.upgradeFrom_1_1_2_6(newVer);
          break;
  
        case "1.1.2.7":
          newVer = this.upgradeFrom_1_1_2_7(newVer);
          break;

        //case "1.1.3":
          // No auto-upgrade from version 1.1.3. Too many spreadsheet changes.

        case "1.1.4":
          newVer = this.upgradeFrom_1_1_4(newVer);
          break;

        case "1.1.4.1":
          newVer = this.upgradeFrom_1_1_4_1(newVer);
          break;

        //case "1.1.4.2":
          // No auto-upgrade from version 1.1.4.2. Script change and trigger added to spreadsheet.

        //case "1.1.4.3":
          // No auto-upgrade from version 1.1.4.3. Mint login code moved from spreadsheet to MojitoLib, plus new Setting.

        case "1.1.5":
        case "1.1.5.1":
        case "1.1.5.2":
        case "1.1.6":
        case "1.1.6.1":
        case "1.1.6.2":
          newVer = this.upgradeFrom_1_1_5(newVer);
          break;

        //case "1.1.6.3":
        // No auto-upgrade from version 1.1.6.3 to 1.2.0. Various updates to spreadsheet, bug fixes to MojitoLib, plus MojitoLib is now fully open source.

        default:
          throw Utilities.formatString("Auto-upgrade from version %s to %s is not supported. Please download the latest version of Mojito instead.", fromVer, toVer);
        }

      } // while

      if (Debug.enabled) Debug.log("Upgrade to version %s complete", toVer);
    }
    catch(e)
    {
      upgradeFailed = true;
      Debug.log(Debug.getExceptionInfo(e));
      //TODO: Need better message to user
      Browser.msgBox("Mojito upgrade failed", Utilities.formatString("Mojito upgrade from version %s to %s failed. Error: %s", fromVer, toVer, e.toString()), Browser.Buttons.OK);
    }

    if (newVer !== fromVer) {
      // Update version in spreadsheet
      Settings.setInternalSetting(Const.IDX_INT_SETTING_MOJITO_VERSION, newVer);
    }

    toast(upgradeFailed ? "Failed." : "Done.", "Mojito upgrade");

    return newVer;
  },

  upgradeFrom_1_0_0_1 : function(fromVer) {
    var targetVer = "1.1.0";
    if (Debug.enabled) Debug.log("Performing upgrade from %s to %s", fromVer, targetVer);

    // Update version / copyright cell on About sheet with a formula so we don't have to keep updating it manually.
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aboutSheet = ss.getSheetByName(Const.SHEET_NAME_ABOUT);
    var cell = aboutSheet.getRange(2, 1);

    var verAndCopyrightFormula = "=\"Mojito, version \" & Settings!B23 & CHAR(10) & CHAR(10) & \"Copyright (c) 2013 - \" & YEAR(TODAY()) & \", b3devs@gmail.com\"";
    cell.setFormula(verAndCopyrightFormula);

    return targetVer;
  },

  upgradeFrom_1_1_0 : function(fromVer) {
    // Just a code upgrade. Nothing to do.
    return "1.1.2";
  },

  upgradeFrom_1_1_2 : function(fromVer) {
    // Add new internal setting to Settings sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsRange = ss.getRangeByName("InternalSettingsRange");
    var labelCell = settingsRange.getCell(Const.IDX_INT_SETTING_CURR_DAY_ACCT_IMPORT, 1);
    var valueCell = settingsRange.getCell(Const.IDX_INT_SETTING_CURR_DAY_ACCT_IMPORT, 2);

    labelCell.setValue("Import current day's balance for accounts having no balance history");
    valueCell.setValue("");

    return "1.1.2.1";
  },

  upgradeFrom_1_1_2_1 : function(fromVer) {
    // Just a code upgrade. Nothing to do.
    return "1.1.2.5";
  },

  upgradeFrom_1_1_2_5 : function(fromVer) {
    // Add new setting to Settings sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsRange = ss.getRangeByName("SettingsRange");
    var labelCell = settingsRange.getCell(Const.IDX_SETTING_TXN_AMOUNT_COL, 1);
    var valueCell = settingsRange.getCell(Const.IDX_SETTING_TXN_AMOUNT_COL, 2);
    var valueFormatCell = settingsRange.getCell(Const.IDX_SETTING_TXN_AMOUNT_COL, 3);

    labelCell.setValue("TxnData column to use for txn amounts");
    labelCell.setNote("This is useful if you want to create a custom \"Amount\" column to make txn amount adjustments, such as currency conversions. The Budget, In/Out, and Savings Goals sheets will use this column for their calculations.");
    valueCell.setValue("");
    valueFormatCell.setValue("Cell located in the desired \"Amount\" column, such as E5 or V10");

    return "1.1.2.6";
  },

  upgradeFrom_1_1_2_6 : function(fromVer) {
    // Just a code upgrade. Nothing to do.
    return "1.1.2.7";
  },

  upgradeFrom_1_1_2_7 : function(fromVer) {
    // Verify that user has added the getDocumentProperties() function to MojtoScripts
    if (MojitoScript.getDocumentProperties === undefined || typeof MojitoScript.getDocumentProperties !== "function") {
      throw "Upgrade to version 1.1.3 will not work until you manually add getDocumentProperties() to the Functions.gs script. See http://b3devs.blogspot.com for how to make this change (it's easy); or you can download a new copy of Mojito.";
    }

    MojitoScript.removeUpdateTimestamps();

    return "1.1.3";
  },

  upgradeFrom_1_1_4 : function(fromVer) {
    // Just a code upgrade. Nothing to do.
    return "1.1.4.1";
  },

  upgradeFrom_1_1_4_1 : function(fromVer) {
    // Just a code upgrade. Nothing to do. (User must download new Mojito spreadsheet to see "Mint password" on Settings sheet)
    return "1.1.4.2";
  },

  // Below function was never actually used because a Mint login problem surfaced before 1.1.4.4 was released
  // that required additional spreadsheet changes.
  upgradeFrom_1_1_4_3 : function(fromVer) {
    // Add new setting to Settings sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsRange = ss.getRangeByName("SettingsRange");
    var labelCell = settingsRange.getCell(Const.IDX_SETTING_MINT_FI_SYNC_TIMEOUT, 1);
    var valueCell = settingsRange.getCell(Const.IDX_SETTING_MINT_FI_SYNC_TIMEOUT, 2);
    var valueFormatCell = settingsRange.getCell(Const.IDX_SETTING_MINT_FI_SYNC_TIMEOUT, 3);

    labelCell.setValue("How long to wait for Mint to finish its back end data refresh with financial institutions before aborting");
    labelCell.setNote("Increase this number if Mint seems to take an extraordinarily long time to retrieve data from your financial institutions.");
    valueCell.setValue(300);
    valueFormatCell.setValue("Time in seconds (300 = 5 minutes)");

    return "1.1.4.4";
  },

  upgradeFrom_1_1_5: function(fromVer) {
    // Just a code upgrade. Nothing to do.
    return "1.1.6.3";
  },

  //--------------------------------------------------------------------------
  compareVersions : function(ver1, ver2) {
    if (ver1 === ver2) {
      return 0;
    }

    if (ver1 == null) {
      return -1;
    }
    if (ver2 == null) {
      return 1;
    }

    // Version format should be 1.2.3.4
    var ver1Parts = ver1.split(".");
    var ver2Parts = ver2.split(".");
    var ver1Size = ver1Parts.length;
    var ver2Size = ver2Parts.length;
    for (var i = 0; i < ver1Size && i < ver2Size; ++i) {
      var v1 = parseInt(ver1Parts[i]);
      var v2 = parseInt(ver2Parts[i]);
      if (isNaN(v1) || isNaN(v2))
        break;

      if (v1 === v2)
        continue;

      if (v1 < v2)
        return -1;
      if (v1 > v2)
        return 1;
    }

    // If we finished the loop, then one version must have fewer parts than the other
    if (ver1Size < ver2Size)
      return -1;

    return 1;
  },
};
