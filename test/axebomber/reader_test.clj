(ns axebomber.reader-test
  (:require [axebomber.reader :as reader]
            [clojure.java.io :as io])
  (:use [axebomber usermodel]
        [axebomber.renderer excel]
        [clojure.pprint]
        [midje.sweet]))

(fact "Read an Excel grid sheet."
  (with-open [file (io/input-stream (io/resource "apply.xls"))]
    (let [wb (open-workbook file)
          components (reader/read wb "記載事項（第1号様式）")
          out-wb (create-workbook)
          sheet (to-grid (.createSheet out-wb "Excel"))]
      (doseq [component components]
        (apply render sheet component))
      (with-open [out (io/output-stream "target/reader.xlsx")]
        (.write out-wb out)))))


