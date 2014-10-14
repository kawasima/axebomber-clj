(ns axebomber.util-test
  (:require [clojure.java.io :as io])
  (:use [axebomber util usermodel]
        [clojure.pprint]
        [midje.sweet]))

(fact "String width."
      (let [wb (create-workbook)]
        (Math/round (string-width "こんにちはExcel" (.getFontAt wb (short 0)))) => 74))
      


