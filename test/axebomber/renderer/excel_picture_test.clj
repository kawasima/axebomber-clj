(ns axebomber.renderer.excel-picture-test
  (:require [clojure.java.io :as io])
  (:use [axebomber.renderer.excel-picture]
        [axebomber usermodel]
        [midje.sweet]))

(fact "Generate hogan."
      (let [wb (create-workbook)
            sheet (to-grid (.createSheet wb "スクショ"))
            [w1 h1 _] (draw-image sheet 10 3 (io/file "dev-resources/google.png"))]
        (draw-image sheet 10 (+ 3 h1 1) (io/file "dev-resources/google.png") :data-width 8)
        (with-open [out (io/output-stream "target/images.xlsx")]
          (.write wb out))))


