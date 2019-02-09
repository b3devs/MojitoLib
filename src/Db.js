'use strict';
/*
 * Copyright (c) 2013-2019 b3devs@gmail.com
 * MIT License: https://spdx.org/licenses/MIT.html
 */


export const Db = {

  load(recId) {
    const docProps = Utils.getDocumentProperties();
    return docProps.getProperty(recId);
  },
  
  save(recId, rec) {
    const docProps = Utils.getDocumentProperties();
    docProps.setProperty(recId, rec);
  },

  remove(recId) {
    const docProps = Utils.getDocumentProperties();
    docProps.deleteProperty(recId);
  },

  removeAll(recIdPrefix) {
    // USE WITH CARE!
    var allKeys = Utils.getDocumentProperties().getProperties();
    for (let i = 0; i < allKeys.length; i++) {
      var recId = allKeys[i];
      if (recId.indexOf(recIdPrefix) === 0) {
        if (Debug.enabled) Debug.log("Removing datastore key: %s", recId);
        this.remove(recId);
      }
    }
  },

  DataStore: {

    saveRecord(key, obj) {
      var propKey = Const.DBTYPE_MOJITO_DATA_STORE + key;
      if (Debug.traceEnabled) Debug.trace("Saving datastore key: %s, value: %s", propKey, JSON.stringify(obj));
      Db.save(propKey, JSON.stringify(obj));
    },

    // Save a specific value to a record
    setRecordValue(key, valueName, value) {
      var obj = {};
      obj[valueName] = value;
      return this.saveRecord(key, obj);
    },

    getRecord(key) {
      var propKey = Const.DBTYPE_MOJITO_DATA_STORE + key;
      var recJson = Db.load(propKey);
      if (Debug.traceEnabled) Debug.trace("Getting datastore key: %s, value: %s", propKey, (recJson == null ? "<null>" : recJson));
      return (recJson ? JSON.parse(recJson) : null);
    },

    // Retrieve a specific value from a record
    getRecordValue(key, valueName, defaultVal) {
      var value = undefined;
      var rec = this.getRecord(key);
      if (rec) {
        value = rec[valueName];
      }

      if (value == undefined) {
        if (defaultVal !== undefined) {
          value = defaultVal;
        } else {
          value = null;
        }
      }

      return value;
    },

    getAllRecords() {
      var allData = Utils.getDocumentProperties().getProperties();
      var allRecs = {};
      for (let key in allData) {
        if (key.indexOf(Const.DBTYPE_MOJITO_DATA_STORE) === 0) {
          allRecs[key.substr(Const.DBTYPE_MOJITO_DATA_STORE.length)] = allData[key];
        }
      }

      for (let key in allRecs) {
        var json = allRecs[key];
        Debug.log("datastore rec: " + json);
      }
      return allRecs;
    },

    removeRecord(key) {
      if (Debug.enabled) { Debug.log("Removing datastore key: %s", key); }
      var propKey = Const.DBTYPE_MOJITO_DATA_STORE + key;
      Db.remove(propKey);
    },

    removeAllRecords() {
      // USE WITH CARE!
      Db.removeAll(Const.DBTYPE_MOJITO_DATA_STORE);
    },
  },

  Messages: {
    getMessage(msgType, msgId) {
      var propKey = msgType + msgId;
      var msgJson = Db.load(propKey);
      if (Debug.traceEnabled) Debug.trace("Getting message, msgId: %s, msg: %s", propKey, (msgJson == null ? "<null>" : msgJson));
      return (msgJson ? JSON.parse(msgJson) : null);
    },

    getAllMessages(msgType) {
      var allData = Utils.getDocumentProperties().getProperties();
      var allMsgs = [];
      Object.keys(allData).forEach((key) => {
        if (key.indexOf(msgType) === 0) {
          var json = allData[key];
          if (Debug.traceEnabled) Debug.trace("Db.Messages.getAllMessages: [%s] %s", key, json);
          allMsgs.push(JSON.parse(json));
        }
      });

      return allMsgs;
    },

    saveMessage(msgType, msgId, msg) {
      var propKey = msgType + msgId;
      Db.save(propKey, JSON.stringify(msg));
    },

    removeMessage(msgType, msgId) {
      Db.remove(msgType + msgId);
    },

    removeMessages(msgType) {
      // USE WITH CARE!
      Db.removeAll(msgType);
    },
  },

};

function getAllDataStoreRecords() {
  Debug.enabled = true;
  Db.DataStore.getAllRecords();
}

function removeAllDataStoreRecords() {
  Debug.enabled = true;
  Db.DataStore.removeAllRecords();
}
