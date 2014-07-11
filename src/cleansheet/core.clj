(ns cleansheet.core
  (:import (org.apache.poi.xssf.usermodel XSSFWorkbook)
           (org.apache.poi.hssf.usermodel HSSFWorkbook)
           (org.apache.poi.ss.usermodel Row Cell WorkbookFactory CellStyle Font Hyperlink Workbook Sheet))
  (:use [clojure.java.io :only [output-stream input-stream]]))

(defn workbook
  ([] (HSSFWorkbook.))
  ([m] (if (= :xssf m) (XSSFWorkbook.) (workbook))))

(defn sheet
  ([^Sheet w] (.createSheet w))
  ([^Sheet w title] (.createSheet w title)))

(defn row
  [^Row s y] (.createRow s y))

(defn cell
  [^Cell r x] (.createCell r x))

(defn load-workbook
  [^Workbook path] (with-open [i (input-stream path)] (WorkbookFactory/create i)))

(defn save-workbook
  [^Workbook wb path] (with-open [o (output-stream path)] (.write wb o)))
