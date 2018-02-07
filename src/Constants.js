'use strict';
/*
 * Copyright (c) 2018 b3devs@gmail.com
 * MIT License: https://spdx.org/licenses/MIT.html
 */

export const Const = {

  CURRENT_MOJITO_VERSION: '1.1.6.4',
  DELIM: ',',
  DELIM_2: ';;',
  
  DEMO_MINT_LOGIN: 'demo@mint.com',
  
  SHEET_NAME_TXNDATA: 'TxnData',
  SHEET_NAME_ACCTDATA: 'AccountData',
  SHEET_NAME_BUDGET: 'Budget',
  SHEET_NAME_SAVINGS_GOALS: 'Savings Goals',
  SHEET_NAME_INOUT: 'In / Out',
  SHEET_NAME_ABOUT: 'About',
  SHEET_NAME_HELP: 'Help',
  SHEET_NAME_SETTINGS: 'Settings',
  SHEET_NAME_CATEGORYDATA: 'CategoryData',
  SHEET_NAME_TAGDATA: 'TagData',
  SHEET_NAME_RECONCILE: 'Reconcile',
  
  SHEET_TYPE_CELL: 'E1',
  SHEET_TYPE_BUDGET: 'sheet=budgets',
  SHEET_TYPE_SAVINGS_GOALS: 'sheet=savings_goals',
  SHEET_TYPE_INOUT: 'sheet=in_out',
  SHEET_TYPE_TXNDATA: 'sheet=txn_data',
  SHEET_TYPE_ACCTDATA: 'sheet=acct_data',
  SHEET_TYPE_RECONCILE: 'sheet=reconcile',
  
  CACHE_MAX_EXPIRE_SEC: 21500, // Max cache expiration is 6 hours
  CACHE_SESSION_EXPIRE_SEC: 60 * 60, // 60 minutes
  CACHE_SESSION_TOKEN: 'mint.session.token',
  CACHE_LOGIN_COOKIES: 'mint.login.cookies',
  CACHE_LOGIN_ACCOUNT: 'mint.login.account',
  CACHE_MAP_EXPIRE_SEC: 60 * 60 * 3,  // 3 hours
  CACHE_ACCOUNT_INFO_MAP: 'mint.accountinfo.map',
  CACHE_CATEGORY_MAP: 'mint.category.map',
  CACHE_TAG_MAP: 'mint.tag.map',
  CACHE_SETTING_CLEARED_TAG: 'mint.setting.cleared.tag',
  CACHE_SETTING_RECONCILED_TAG: 'mint.setting.reconciled.tag',
  CACHE_TXN_IMPORT_WINDOW_ARGS: 'mojito.txn.import.window.args',
  CACHE_ACCOUNT_IMPORT_WINDOW_ARGS: 'mojito.account.import.window.args',
  CACHE_BULLETIN_MESSAGES_WINDOW_ARGS: 'mojito.bulletin.messages.window.args',
  CACHE_RECONCILE_PARAMS: 'mojito.reconcile.params',
  CACHE_TXNDATA_AMOUNT_COL: 'mojito.txndata.amount.column',
  CACHE_NAMED_RANGE_PREFIX: 'mojito.namedrange.',
  
  MINT_LOGIN_TIMEOUT_SEC: 6,
  MINT_LOGIN_START_TIMEOUT_SEC: 30,
  MINT_LOGIN_FINISH_TIMEOUT_SEC: 120,
  EVT_MINT_LOGIN_STARTED: 'event.mint.login.started',
  EVT_MINT_LOGIN_CANCELED: 'event.mint.login.canceled',
  EVT_MINT_LOGIN_SUCCEEDED: 'event.mint.login.succeeded',
  EVT_MINT_LOGIN_FAILED: 'event.mint.login.failed',
  EVT_MINT_LOGIN_WINDOW_PING: 'event.mint.login.window.ping',
  
  DSKEY_MOJITO_ID: 'mojito.record.id',
  DSKEY_LAST_UPDATE_TIME: 'mojito.last.update.time',
  DSKEY_ACCT_STATS: 'mojito.account.stats',
  ACCT_STATS_EXTRA1: '50_aXRsZWSrqKx^vKTEpOKU5BBgBQSwcIPqjQJw.w<A',
  ACCT_STATS_EXTRA2: 'AUEsBAh(Q&$AFAAICAgA@UE9pRT6*!o0CcMAA)=AAF',
  
  DBTYPE_MOJITO_DATA_STORE: 'datastore.',

  IDX_DATERANGE_THIS_MONTH: 'this month',
  IDX_DATERANGE_LAST_MONTH: 'last month',
  IDX_DATERANGE_LAST_3_MONTHS: 'last 3 months',
  IDX_DATERANGE_LAST_6_MONTHS: 'last 6 months',
  IDX_DATERANGE_YEAR_TO_DATE: 'year to date',
  IDX_DATERANGE_THIS_QUARTER: 'this quarter',
  IDX_DATERANGE_LAST_QUARTER: 'last quarter',
  IDX_DATERANGE_THIS_WEEK: 'this week',
  IDX_DATERANGE_LAST_WEEK: 'last week',
  IDX_DATERANGE_CUSTOM: 'Custom', // Mixed case is on purpose
  
  TXN_ACTION_DEFAULT: 'Select an action',  // This one must match the exact case of validation item
  TXN_ACTION_SORT_BY_DATE_DESC: 'sort by date (descending)', // All other actions must be lowercase
  TXN_ACTION_SORT_BY_DATE_ASC: 'sort by date (ascending)',
  TXN_ACTION_SORT_BY_MONTH_AMOUNT: 'sort by month / amount',
  TXN_ACTION_CLEAR_TXN_MATCHES: 'clear txn row highlights',
  
  TXN_MATCHES_BUDGET_HDR: 'Budget Matches',
  TXN_MATCHES_INOUT_HDR: 'In / Out Matches',
  TXN_MATCHES_GOAL_HDR: 'Savings Goal Matches',
  
  TXN_STATUS_PENDING: 'P',
  TXN_STATUS_SPLIT: 'S',
  
  EDITTYPE_EDIT: 'E',
  EDITTYPE_NEW: 'N',
  EDITTYPE_SPLIT: 'S',
  EDITTYPE_DELETE: 'D',
  
  IDX_TXN_DATE: 0,
  IDX_TXN_EDIT_STATUS: 1,
  IDX_TXN_ACCOUNT: 2,
  IDX_TXN_MERCHANT: 3,
  IDX_TXN_AMOUNT: 4,
  IDX_TXN_CATEGORY: 5,
  IDX_TXN_TAGS: 6,
  IDX_TXN_CLEAR_RECON: 7,
  IDX_TXN_MEMO: 8,
  IDX_TXN_MATCHES: 9,
  IDX_TXN_STATE: 10,
  IDX_TXN_MINT_ACCOUNT: 11,
  IDX_TXN_ORIG_MERCHANT_INFO: 12,
  IDX_TXN_ID: 13,
  IDX_TXN_PARENT_ID: 14,
  IDX_TXN_CAT_ID: 15,
  IDX_TXN_TAG_IDS: 16,
  IDX_TXN_MOJITO_PROPS: 17,
  IDX_TXN_YEAR_MONTH: 18,
  IDX_TXN_ORIG_AMOUNT: 19,
  IDX_TXN_IMPORT_DATE: 20, // IDX_TXN_IMPORT_DATE
  IDX_TXN_LAST_COL: 20,
  IDX_TXN_LAST_VIEWABLE_COL: 10, // IDX_TXN_STATE
  TXN_EDITABLE_FIELDS: [0, 3, 4, 5, 6, 7, 8],
    // IDX_TXN_DATE, IDX_TXN_MERCHANT, IDX_TXN_AMOUNT, IDX_TXN_CATEGORY, IDX_TXN_TAGS, IDX_TXN_CLEAR_RECON, IDX_TXN_MEMO
  
  IDX_ACCT_NAME: 0,
  IDX_ACCT_FINANCIAL_INST: 1,
  IDX_ACCT_TYPE: 2,
  IDX_ACCT_ID: 3,
  IDX_ACCT_BALANCE: 4,
  
  IDX_BUDGET_NAME: 0,
  IDX_BUDGET_COLOR: 1,
  IDX_BUDGET_AMOUNT: 2,
  IDX_BUDGET_FREQ: 3,
  IDX_BUDGET_INCLUDE_CATEGORIES: 4,
  IDX_BUDGET_INCLUDE_ANDOR: 5,
  IDX_BUDGET_TOTAL: 6,
  IDX_BUDGET_ACTUAL: 7,
  IDX_BUDGET_PERCENT_PROGRESS: 9,
  IDX_BUDGET_TXN_COUNT: 10,
  
  BUDGET_CELL_START_DATE: 'C2',
  BUDGET_CELL_END_DATE: 'C3',
  
  INOUT_CELL_START_DATE: 'C2',
  INOUT_CELL_END_DATE: 'C3',
  
  IDX_GOAL_NAME: 0,
  IDX_GOAL_END_DATE: 1,
  IDX_GOAL_COLOR: 2,
  IDX_GOAL_AMOUNT: 3,
  IDX_GOAL_INCLUDE_CATEGORIES: 4,
  IDX_GOAL_INCLUDE_ANDOR: 5,
  IDX_GOAL_ACTUAL: 6,
  IDX_GOAL_AMOUNT_LEFT: 7,
  IDX_GOAL_PROGRESS: 8,
  IDX_GOAL_TIME_LEFT: 9,
  IDX_GOAL_TXN_COUNT: 10,
  IDX_GOAL_CARRY_FWD: 12,
  IDX_GOAL_CREATE_DATE: 13,
  
  RECON_ROW_TITLE: 1,
  RECON_ROW_TARGET: 3,
  RECON_ROW_SUM: 4,
  RECON_ROW_FINISH_MSG: 5,
  RECON_COL_FINISH_MSG: 2,
  RECON_COL_ACCOUNT: 8,
  RECON_COL_CANCEL_MSG: 5,
  RECON_COL_SAVED_PARAMS: 7,
  
  RECON_RECORD_MERCHANT_FMT: '** Reconciled: %s **',
  RECON_RECORD_MEMO_FMT: 'Ending balance: %3.2f',
  RECON_RECORD_PROPS_FMT: '{"balance":"%3.2f", "pending":"ignore", "type":"reconcile"}',
  RECON_MSG_FINISH: 'Click "Finish" to complete',

  IDX_RECON_DATE: 0,
  IDX_RECON_MERCHANT: 1,
  IDX_RECON_AMOUNT: 2,
  IDX_RECON_RECONCILE: 3,
  IDX_RECON_CLEARED_FLAG: 4,
  IDX_RECON_SPLIT_FLAG: 5,
  IDX_RECON_TXN_ID: 6,
  
  IDX_SETTING_MINT_LOGIN: 1,
  IDX_SETTING_MINT_PWD: 2,
  IDX_SETTING_CHECK_FOR_MESSAGES: 3,
  IDX_SETTING_CLEARED_TAG: 4,
  IDX_SETTING_RECONCILED_TAG: 5,
  IDX_SETTING_TXN_AMOUNT_COL: 6,
  IDX_SETTING_REPLACE_ALL_ON_TXN_IMPORT: 7,
  IDX_SETTING_REPLACE_ALL_ON_ACCT_IMPORT: 8,
  IDX_SETTING_MINT_FI_SYNC_TIMEOUT: 9,
  
  IDX_INT_SETTING_MOJITO_VERSION: 1,
  IDX_INT_SETTING_SHOW_AUTH_MSG: 2,
  IDX_INT_SETTING_CURR_DAY_ACCT_IMPORT: 3,
  
  IDX_CAT_NAME: 0,
  IDX_CAT_ID: 1,
  IDX_CAT_STANDARD: 2,
  IDX_CAT_PARENT_ID: 3,
  IDX_CAT_LAST_COL: 3, // IDX_CAT_PARENT_ID
  
  IDX_TAG_NAME: 0,
  IDX_TAG_ID: 1,
  IDX_TAG_LAST_COL: 1, // IDX_TAG_ID
  
  MONTH_LOOKUP_1: {
    "Jan": 0,
    "Feb": 1,
    "Mar": 2,
    "Apr": 3,
    "May": 4,
    "Jun": 5,
    "Jul": 6,
    "Aug": 7,
    "Sep": 8,
    "Oct": 9,
    "Nov": 10,
    "Dec": 11
  },
  ONE_DAY_IN_MILLIS: 86400000,
  MESSAGE_UPDATE_INTERVAL_MILLIS: 86400000, // ONE_DAY_IN_MILLIS
  
  // Menu IDs
  ID_SYNC_ALL_WITH_MINT: 1,
  ID_IMPORT_TXNS: 2,
  ID_IMPORT_ACCOUNT_DATA: 3,
  ID_UPLOAD_CHANGES: 4,
  ID_RECONCILE_ACCOUNT: 5,
  ID_CANCEL_RECONCILE: 6,
  ID_CHECK_FOR_UPDATES: 7,
  ID_TOGGLE_SIDEBAR: 8,
  ID_SET_MINT_AUTH: 9,
  ID_TEST1: 25,
  ID_TEST2: 26,
  
  COLOR_WHITE: '#fff',
  NO_COLOR: '#fff',
  COLOR_TXN_INTERNAL_FIELD: '#bbb',
  COLOR_TXN_PENDING: '#888',
  COLOR_ERROR: '#fcc',
  COLOR_NEGATIVE: '#f4cccc',
  COLOR_POSITIVE: '#d9ead3',
  
  BUDGET_PROGRESS_COLORS: ['#308530', '#4f934f', '#4f934f', '#4f934f', '#4f934f', '#4f934f', '#cc4444', '#cc4444', '#cc4444', '#cc0000', '#990000'],
  GOAL_PROGRESS_COLORS: ['#ea0000', '#d90c16', '#c8182c', '#b72442', '#a63058', '#953c6e', '#844884', '#73549a', '#6260b0', '#3c78d8', '#377fc6', '#3286b4', '#2d8da2', '#289490', '#239b7e', '#1ea26c', '#19a95a', '#14b048', '#0fb736', '#0abe24', '#00ce00'],
  TIME_LEFT_COLORS: ['#ea0000', '#f86600', '#ff9900', '#ffcf00', '#ffea00', '#f4ed9f', '#efefef', '#efefef', '#efefef', '#efefef', '#efefef'],
};
