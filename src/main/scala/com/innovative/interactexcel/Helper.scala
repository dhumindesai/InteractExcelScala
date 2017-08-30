package com.innovative.interactexcel

import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel._
import java.text.SimpleDateFormat
import java.util.Date

import org.apache.poi.hssf.usermodel.HSSFDateUtil


object Helper {

  def getCellData(currentCell: Cell, isValue: Boolean): String = {
    val cellTypeEnum = currentCell.getCellTypeEnum

    cellTypeEnum match{
      case CellType.STRING => if (isValue) currentCell.getStringCellValue else classOf [String].getSimpleName

      case CellType.NUMERIC => {
        if (DateUtil.isCellDateFormatted(currentCell)){ // Checking if NUMERIC type is Date or not
          if (isValue) new SimpleDateFormat("MM/dd/yyyy").format(currentCell.getDateCellValue)
          else classOf[Date].getSimpleName
        }
        else if (currentCell.getNumericCellValue == Math.ceil(currentCell.getNumericCellValue)) { // checking if the value is float
          if (isValue) currentCell.getNumericCellValue.toString else classOf[Int].getSimpleName
        }
        else
          if (isValue) currentCell.getNumericCellValue.toString else classOf[Float].getSimpleName;
      } //celltype numeric

      case CellType.FORMULA => {
        if (currentCell.getCachedFormulaResultTypeEnum eq CellType.NUMERIC){
          if (DateUtil.isCellDateFormatted(currentCell)){
            if (isValue) new SimpleDateFormat("MM/dd/yyyy").format(currentCell.getDateCellValue)  else classOf[Date].getSimpleName
          }
          else {
            if(isValue) currentCell.getNumericCellValue.toString else classOf[Float].getSimpleName
          }
        }
        else if (currentCell.getCachedFormulaResultTypeEnum eq CellType.STRING){
          if(isValue) currentCell.getStringCellValue  else  classOf[String].getSimpleName
        }
        else{
          if (isValue) "NULL" else  "Unknown"
        }
      } // celltype formula

      case CellType.BOOLEAN => if (isValue) currentCell.getBooleanCellValue.toString else classOf[Boolean].getSimpleName

      case CellType.BLANK => "NULL"

      case _ => "None"
    } // match

  }
}
