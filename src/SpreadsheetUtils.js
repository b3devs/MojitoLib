'use strict';
/*
 * Copyright (c) 2013-2019 b3devs@gmail.com
 * MIT License: https://spdx.org/licenses/MIT.html
 */

export const SpreadsheetUtils = {
  copySheet(srcSS, sheetName, destSS) {
    var sheet = srcSS.getSheetByName(sheetName);
    sheet.copyTo(destSS).setName(sheetName + "- Copy");
  },

  setRowColors(range, rowColors, isRowColors2d, defaultColor, setFontColor) {
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();
    var rowColorsLen = (rowColors === null ? 0 : rowColors.length);

    let array = new Array(numRows);
    for (var i = 0; i < numRows; ++i) {
      array[i] = new Array(numCols);

      for (var j = 0; j < numCols; ++j) {
        var color = (i < rowColorsLen ? (isRowColors2d ? rowColors[i][0] : rowColors[i]) : defaultColor);
        array[i][j] = color;
      }
    }

    if (setFontColor === true) {
      range.setFontColors(array);
    } else {
      range.setBackgrounds(array);
    }
  },

  setRangeStrikeThrough(range)
  {
    var fontLines = range.getFontLines();
    if (fontLines === null || fontLines.length === 0)
      return;

    var numRows = fontLines.length;
    var numCols = fontLines[0].length;

    for (var i = 0; i < numRows; ++i)
    {
      for (var j = 0; j < numCols; ++j)
      {
        fontLines[i][j] = "line-through";
      }
    }

    range.setFontLines(fontLines);
  }
};
