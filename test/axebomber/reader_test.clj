(ns axebomber.reader-test
  (:require [axebomber.reader :as reader]
            [clojure.java.io :as io])
  (:use [axebomber usermodel]
        [axebomber.renderer excel]
        [clojure.pprint]
        [midje.sweet]))

(fact "Read an Excel grid sheet."
  (with-open [file (io/input-stream (io/resource "reader-test.xlsx"))]
    (let [wb (open-workbook file)
          components (reader/read wb "読取テスト用方眼紙")
          out-wb (create-workbook)
          sheet (to-grid (.createSheet out-wb "Excel"))]
      (doseq [component components]
        (apply render sheet component))
      (reader/copy-grid (.getSheet wb "読取テスト用方眼紙") sheet)
      (with-open [out (io/output-stream "target/reader.xlsx")]
        (.write out-wb out)))))

