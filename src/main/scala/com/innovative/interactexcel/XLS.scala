package com.innovative.interactexcel

import java.io.{File, FileInputStream}

import org.apache.poi.ss.usermodel._
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.util.CellUtil

import scala.collection.mutable.ListBuffer
import scala.collection.JavaConverters
import scala.collection.immutable.Stream


case class XLS(var filePath: String) {

  val excelFile: FileInputStream = new FileInputStream(new File(filePath))
  val workbook: Workbook = new HSSFWorkbook(excelFile)

      def getSheetNames() = {
          val sheetNames = ListBuffer[String]()
          for ( i <- 0 until workbook.getNumberOfSheets) {
            sheetNames += workbook.getSheetName(i)
          }
          sheetNames.toList
      }

      def getSheetCounts() = {
          workbook.getNumberOfSheets
      }

      def getAllSheets() = {
        val sheets = ListBuffer[SheetObj]()
        for ( i <- 0 until workbook.getNumberOfSheets) {
          sheets += SheetObj(workbook,i)
        }
        sheets.toList
      }

      def getSheet(index: Int): SheetObj = SheetObj(workbook,index)

      def getSheetTail(index: Int): Stream[RowObj] = SheetObj(workbook,index).getTail
      def getSheetHead(index: Int): RowObj = SheetObj(workbook,index).getHeader

}


