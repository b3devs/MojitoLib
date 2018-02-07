/*
 * Copyright (c) 2018 b3devs@gmail.com
 * MIT License: https://spdx.org/licenses/MIT.html
 */
MojitoLib = this;

// Wire up global button handlers (for gapps dialogs)
MojitoLib.Debug_onOkClicked = Debug.onOkClicked;
MojitoLib.Debug_onRefreshClicked = Debug.onRefreshClicked;
MojitoLib.Debug_onClearClicked = Debug.onClearClicked;
MojitoLib.Reconcile_Window_onOkClicked = Reconcile.Window.onOkClicked;
MojitoLib.Reconcile_Window_onCancelClicked = Reconcile.Window.onCancelClicked;
MojitoLib.Reconcile_Window_onAccountChanged = Reconcile.Window.onAccountChanged;
MojitoLib.TxnImportWindow_onOkClicked = Ui.TxnImportWindow.onOkClicked;
MojitoLib.TxnImportWindow_onCancelClicked = Ui.TxnImportWindow.onCancelClicked;
MojitoLib.AccountBalanceImportWindow_onOkClicked = Ui.AccountBalanceImportWindow.onOkClicked;
MojitoLib.AccountBalanceImportWindow_onCancelClicked = Ui.AccountBalanceImportWindow.onCancelClicked;
