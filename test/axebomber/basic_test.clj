(ns axebomber.basic-test
  (:require [clojure.java.io :as io])
  (:use [axebomber usermodel]
    [axebomber.renderer excel]
    [axebomber.style :only [create-style]]
    [midje.sweet])
  (:import [org.apache.poi.ss.usermodel CellStyle IndexedColors]
           [java.util Date]))

(facts "Renders basic features"
       (let [wb (create-workbook)]
         (fact "Simple render"
               (let [sheet (to-grid (.createSheet wb "Simple render"))]
                 (render sheet 1 1 "こんにちは、Excel方眼紙！")
                 (render sheet 1 2 "さようなら、UIとしてのExcel…")))
         (fact "Simple tabel"
               (let [sheet (to-grid (.createSheet wb "Simple table"))]
                 (render sheet 0 0
                         [:table
                          [:tr
                           [:td "ID"]
                           [:td "名前"]]
                          [:tr
                           [:td 1]
                           [:td "りんご"]]
                          [:tr
                           [:td 2]
                           [:td "ばなな"]]])))
         (fact "Styled table"
               (let [sheet (to-grid (.createSheet wb "Styled table"))]
                 (render sheet 0 0
                         [:table
                          [:tr
                           [:td {:data-width 3 :style "background-color: lightblue"} "ID"]
                           [:td {:data-width 8 :style "background-color: lightgreen"} "名前"]]
                          [:tr
                           [:td 1]
                           [:td "りんご"]]
                          [:tr
                           [:td 2]
                           [:td "ばなな"]]])))
         (fact "Class-based style"
               (let [sheet (to-grid (.createSheet wb "Class-based style"))]
                 (create-style ".title1" :background-color "lightblue")
                 (create-style ".title2" :background-color "lightgreen")
                 (render sheet 1 1
                         [:table
                          [:tr
                           [:td.title1 {:data-width 3} "ID"]
                           [:td.title2 {:data-width 8} "名前"]]
                          [:tr
                           [:td 1]
                           [:td "りんご"]]
                          [:tr
                           [:td 2]
                           [:td "ばなな"]]])))

         (fact "Table in table."
               (let [sheet (to-grid (.createSheet wb "Table in table"))]
                 (render sheet 1 1
                         [:table
                          [:tr
                           [:td.title1 {:data-width 3} "ID"]
                           [:td.title2 {:data-width 8} "名前"]]
                          [:tr
                           [:td 1]
                           [:td [:table {:data-margin-top 1 :data-margin-left 1 :data-margin-bottom 1}
                                 [:tr
                                  [:td {:data-width 3} ""]
                                  [:td {:data-width 3} "産地"]]
                                 [:tr
                                  [:td "ふじ"]
                                  [:td "青森"]]
                                 [:tr
                                  [:td "あずさ"]
                                  [:td "長野"]]]]]
                          [:tr
                           [:td 2]
                           [:td "ばなな"]]])))

         (fact "Basic list"
               (let [sheet (to-grid (.createSheet wb "Basic list"))]
                 (render sheet 1 1
                         [:ul
                          [:li "りんご"]
                          [:li "ばなな"]
                          [:li "いちご"]])))

         (fact "Ordered list"
               (let [sheet (to-grid (.createSheet wb "Orderd list"))]
                 (render sheet 1 1
                         [:ol
                          [:li "りんご"]
                          [:li "ばなな"]
                          [:li "いちご"]])))

         (fact "Graphics"
               (let [sheet (to-grid (.createSheet wb "Graphics"))]
                 (render sheet 1 1
                         [:graphics {:data-width 30 :data-height 30}
                          [:box {:x 1 :y 1 :w 3 :h 3} "ハコ1"]
                          [:box {:x 5 :y 2 :w 4 :h 2} "ハコ2"]])))


         (with-open [out (io/output-stream "target/simple.xlsx")]
           (.write wb out))))
