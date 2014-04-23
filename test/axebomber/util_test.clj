(ns axebomber.util-test
  (:require [axebomber.reader :as reader]
            [clojure.java.io :as io])
  (:use [axebomber util]
        [clojure.pprint]
        [midje.sweet])
  (:import [java.awt Font]))

(fact "String width."
      (string-width "こんにちはExcel" (Font. "ＭＳ Ｐゴシック" Font/PLAIN 12)) => 80)


