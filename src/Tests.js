/*
 * Copyright (c) 2013-2019 b3devs@gmail.com
 * MIT License: https://spdx.org/licenses/MIT.html
 */
import {Const} from './Constants.js'
import {Debug} from './Debug.js';
import {Utils, Settings} from './Utils.js';
import {Db} from './Db.js';
import {Mint} from './MintApi.js';

/////////////////////////////////////////////////////////////////////////////////////////////////
export const Tests = {
  TEST_USER: "test@mojito.com",

  executeTests: function() {
    Debug.enabled = false;
    Debug.traceEnabled = false;

    for (var test in Tests) {
      // Skip any functions that don't start with "test"
      if (test.indexOf("test") !== 0)
        continue;

      // execute the test
      try
      {
        // Run setup() before and teardown() after every test()
        Tests.setup();

        var testFunc = `Tests.${test}()`;
        eval(testFunc);
        Logger.log("PASS : %s", test);

        Tests.teardown();
      }
      catch (e)
      {
        Logger.log("FAIL : %s: %s", test, e.toString());
        Logger.log("    Call stack: %s", Debug.getExceptionInfo(e));
      }
    }
  },

  debugTest: function() {
    Debug.enabled = true;
    Debug.traceEnabled = true;

    Tests.setup();
    try
    {
      Tests.test_messageGetSet();
    }
    catch (e)
    {
      Logger.log(Debug.getExceptionInfo(e));
    }

    Tests.teardown();
  },

  //---------------------------------------------------------------------------

  setup : function() {
    // Set up mocks
    Utils.getPrivateCache().put(Const.CACHE_LOGIN_ACCOUNT, this.TEST_USER);
  },

  teardown: function() {
    // Undo mocks
    //Settings = SettingsImpl;
    // Remove login cookies / account from cache
    Utils.getPrivateCache().remove(Const.CACHE_LOGIN_ACCOUNT);
    Mint.Session.clearCookies();
  },

  test_datatoreGetSet: function() {
    var key = "mojito.test.val";
    var testVal = "abc";
    Db.DataStore.saveRecord(key, testVal);
    assertEqual(testVal, Db.DataStore.getRecord(key));
  },

  test_messageGetSet: function() {
    var msgType = "msgtype.";
    var msgId = "mojito.test.msg";
    var testMsg = { msg_id: msgId };
    Db.Messages.saveMessage(msgType, msgId, testMsg);
    var retMsg = Db.Messages.getMessage(msgType, msgId);
    assertEqual(testMsg.msg_id, retMsg.msg_id);
  },
};

function assert(expr) {
  if (!expr)
    throw "assert failed";
}
function assertEqual(expected, actual) {
  if (expected !== actual)
    throw new Error(`assertEqual failed: ${String(expected)} !== ${String(actual)}`);
}
function assertNull(obj) {
  if (obj !== null)
    throw new Error('assertNull failed: obj !== null');
}
function assertNotNull(obj) {
  if (obj === null)
    throw new Error('assertNotNull failed: obj === null');
}

const Mock = {
  Settings: {
    settingsMap: {},
    internalSettingsMap: {},

    getSetting: function(settingIndex) {
      return this.settingsMap[settingIndex];
    },
    
    setSetting: function(settingIndex, value) {
      this.settingsMap[settingIndex] = value;
    },
    
    getInternalSetting: function(settingIndex) {
      return this.internalSettingsMap[settingIndex];
    },
    
    setInternalSetting: function(settingIndex, value) {
      this.internalSettingsMap[settingIndex] = value;
    },
  },  

};

//Settings = Mock.Settings;
