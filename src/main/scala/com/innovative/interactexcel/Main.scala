/*
package com.innovative.interactexcel

object Main extends App{
    println("Started .. .. ..")

    val xls = new XLS("/Users/dhrumindesai/Desktop/JavaProjects/Sample - Superstore Sales (Excel).xls")
   // println(xls.getSheetNames)
    //println(xls.getAllSheets())
    /*
    for (sheet <- xls.getAllSheets){
        println(sheet.getSheetName())
        println(sheet.getColumnDataTypes)
        println(sheet.getRowCounts)
        println(sheet.getColumnCounts)
        println
    }
    */

    /*
    xls.getSheetHead(0).getCellValues.foreach(c=>printf(s" $c |"))
    xls.getSheetTail(0).foreach((row)=>{
        row.getCellValues.foreach(c=>print(c+" |"))
        println
    })
    */

    xls.getSheet(0).getTailByQuery(4,'=',"49").foreach((row)=>{
        row match{
            case r: RowObj => {
                r.getCellValues.foreach(c=>print(c+" |"))
                println
            }
            case _ => println("Not found")
        }

    })
}
*/
