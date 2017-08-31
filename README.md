# InteractExcelScala

- Simple Interaction with XLS file
- get Metadata like Number of Worksheet, Names etc in simple Scala Data Structure
- get elements of specific worksheet in basic Data Structure of Scala

Usage:

1. All Methods to get metadata of XLS file

import com.innovative.interactexcel.XLS

val xls = new XLS(filepath)

xls.getSheetNames  // returns all sheet names as List
xls.getAllSheets   // returns all SheetObj(Sheet Objects) 
xls.getSheetCounts // return sheet counts
xls.getSheet(sheetIndex: Int) // returns sheet with given index
xls.getSheetHead(index: Int) // get Header of the given sheet
xls.getSheetTail(index: Int) // ç


2. All Methods on SheetObj (Sheet Object)

// get sheet object
val sheet = xls.getSheet(0)

sheet.getSheetName // returns name of the sheets
sheet.getRowCounts // returns counts of Rows
sheet.getColumnCounts // returns counts of Columns
sheet.getColumnDataTypes // returns column types as List
sheet.getHeader // get Header of the sheet
sheet.getTail // get Taul of the sheet
getTailByQuery(columnNum: Int, operator: Char, operand: String) //get filtered Rows based on the given Query
