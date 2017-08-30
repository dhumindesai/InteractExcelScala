package com.innovative.interactexcel

import java.util
import java.util.Iterator

import org.apache.poi.ss.usermodel._
import org.apache.poi.ss.util.CellUtil
import org.apache.poi.ss.usermodel.DataFormatter


import scala.collection.JavaConverters
import scala.collection.mutable.ListBuffer
import scala.util.Try

case class SheetObj (var workbook: Workbook, var index: Int){

  val sheet = workbook.getSheetAt(index)

  def getSheetName(): String = workbook.getSheetName(index)

  def getRowCounts(): Int = sheet.getPhysicalNumberOfRows

  def getColumnCounts(): Int = sheet.getRow(0).getPhysicalNumberOfCells

  def getColumnDataTypes(): List[String] = {

    val row = sheet.getRow(1)
    val cellIterator = JavaConverters.asScalaIterator(row.iterator)

    (for(cell <- cellIterator) yield Helper.getCellData(cell,false)).toList
  }

  def getHeader(): RowObj = RowObj(sheet.getRow(0))

  def getTail(): Stream[RowObj] = {
    val rowIterator = JavaConverters.asScalaIterator(sheet.iterator).toStream
    for (row <- rowIterator.drop(1)) yield RowObj(row)
  }

  def getTailByQuery(columnNum: Int, operator: Char, operand: String): ListBuffer[RowObj] = {

    val isInteger = Try(operand.toInt).isSuccess
    val rowIterator = JavaConverters.asScalaIterator(sheet.iterator).toStream
    val result: ListBuffer[RowObj] = new ListBuffer[RowObj]
    //result += this.getHeader()

      for (row <- rowIterator.drop(1)){

      val currentCell = CellUtil.getCell(row, columnNum)

      val rowObj = if (isInteger) {
            val numVal = operand.toFloat
            operator match {
              case '=' => if (numVal == currentCell.getNumericCellValue){ result += RowObj(row)}
              case '>' => if (numVal < currentCell.getNumericCellValue) { result += RowObj(row)}
              case '<' => if (numVal > currentCell.getNumericCellValue) { result += RowObj(row)}
              case _ => throw new IllegalArgumentException
            } // end of operator case
          } // end num case

          else {
            val strVal = operand.asInstanceOf[String]
            operator match {
              case '=' => if (strVal.equals(currentCell.getStringCellValue)) { result += RowObj(row)}
              //case _ => {throw IllegalArgumentException}
            }
          } // end string case

          //case _ => throw new IllegalArgumentException
    }  // end of result
    result
  } // end of method

}

case class RowObj(val row: Row){

  val cellIterator = JavaConverters.asScalaIterator(row.iterator)

  def getCellValues(): Stream[String] = {
    (for (cell <- cellIterator) yield Helper.getCellData(cell,true)).toStream
  }
}
