/*
 * Copyright (c) 2013 - 2014, b3devs@gmail.com
 *
 * Licensed under the Common Development and Distribution License (the "License"); you may not use this
 * source code or associated Google document except in compliance with the License. You may obtain a
 * copy of the License at
 *
 * http://opensource.org/licenses/cddl1.php
 *
 * Any modifications to this Software that You distribute or otherwise make available in Executable
 * form must also be made available in Source Code form and that Source Code form must be distributed
 * only under the terms of this License.
 *
 * Unless required by applicable law or agreed to in writing, software distributed under the License
 * is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
 * or implied. See the License for the specific language governing permissions and limitations under
 * the License.
 */
import {Const} from './Constants.js'
import {Debug} from './Debug.js';
import {Utils} from './Utils.js';
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

        var testFunc = Utilities.formatString("Tests.%s()", test);
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
    throw Utilities.formatString("assertEqual failed: %s !== %s", String(expected), String(actual));
}
function assertNull(obj) {
  if (obj !== null)
    throw "assertNull failed: obj !== null";
}
function assertNotNull(obj) {
  if (obj === null)
    throw "assertNotNull failed: obj === null";
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

Settings = Mock.Settings;
