(ns excellence.spreadsheet-test
  (:use [excellence.spreadsheet] :reload-all)
  (:use clojure.test)
  (:require [clojure.java.io :as io]
            [clojure.string :as str])
  (:import 
   (java.io File FileInputStream FileOutputStream)
   (java.text DecimalFormat SimpleDateFormat)
   (java.util Calendar Date GregorianCalendar Locale)
   (org.apache.poi.xssf.usermodel XSSFWorkbook XSSFRichTextString)
   (org.apache.poi.hssf.usermodel HSSFWorkbook HSSFRichTextString)
   (org.apache.poi.ss.usermodel Cell CellStyle DataFormatter DateUtil Font
                                FormulaEvaluator IndexedColors RichTextString
                                Row Sheet Workbook WorkbookFactory)
   (org.apache.poi.ss.util CellReference)))

(deftest test-workbook
  (testing "'workbook'"
    (testing "neue Arbeitsmappe erzeugen"
        (is (or (instance? (workbook)) HSSFWorkbook)
            (or (instance? (workbook)) XSSFWorkbook)))
    (testing "vorhandene Arbeitsmappe laden")))