(ns axebomber.reader-test
  (:require [axebomber.reader :as reader]
            [clojure.java.io :as io])
  (:use [axebomber usermodel]
        [axebomber.renderer excel]
        [clojure.pprint]
        [midje.sweet]))

(fact "Read an Excel grid sheet."
  (with-open [file (io/input-stream (io/resource "非機能テスト-1.4.xlsx"))]
    (let [wb (open-workbook file)
          components (reader/read wb "大規模アプリテスト")
          out-wb (create-workbook)
          sheet (to-grid (.createSheet out-wb "Excel"))]
      (doseq [component components]
        (prn component)
        (apply render sheet component))
      (reader/copy-grid (.getSheet wb "大規模アプリテスト") sheet)
      (with-open [out (io/output-stream "target/reader.xlsx")]
        (.write out-wb out)))))

