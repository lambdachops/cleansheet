(ns cleansheet.core-test
  (:require [clojure.test :refer :all]
            [exclj.core :refer :all]))

(deftest save-a-simple-sheet
  (testing "load and traverse a spreadsheet"
    (load-workbook "44010-SingleChart.xls"))
  (testing "save a simple spreadsheet"
    (let [wb (workbook)] (sheet wb "main"))))
