(ns axebomber.complex-test
  (:require [clojure.java.io :as io])
  (:use [axebomber usermodel]
    [axebomber.renderer excel]
    [axebomber.style :only [create-style]]
    [midje.sweet])
  (:import [org.apache.poi.ss.usermodel CellStyle IndexedColors]
           [java.util Date]))

(defn template [sheet model]
  (render sheet {:x 1 :y 1}
          [:table
           [:tr
            [:td.header {:data-width 28} "営業日報"]]
           [:tr
            [:td.title {:data-width 3} "店舗"]
            [:td {:data-width 8} (:店舗 model)]
            [:td.title {:data-width 3} "担当"]
            [:td {:data-width 7} (:担当 model)]
            [:td.title {:data-width 2} "日付"]
            [:td {:data-width 5} (:日付 model)]]
           [:tr
            [:td.title {:data-width 4} ""]
            [:td.title {:data-width 5} "計画"]
            [:td.title {:data-width 5} "実績"]
            [:td.title {:data-width 5} "達成率"]
            [:td.title {:data-width 9} "備考"]]
           [:tr
            [:td.title "来客数"]
            [:td.number (get-in model [:計画 :来客数])]
            [:td.number (get-in model [:実績 :来客数])]
            [:td.number (get-in model [:達成率 :来客数])]
            [:td {:rowspan 4} (get-in model [:備考])]]
           [:tr
            [:td.title "売上高"]
            [:td.number (get-in model [:計画 :売上高])]
            [:td.number (get-in model [:実績 :売上高])]
            [:td.number (get-in model [:達成率 :売上高])]]
           [:tr
            [:td.title "客単価"]
            [:td.number (get-in model [:計画 :客単価])]
            [:td.number (get-in model [:実績 :客単価])]
            [:td.number (get-in model [:達成率 :客単価])]]
           [:tr
            [:td.title "販売点数"]
            [:td.number (get-in model [:計画 :来客数])]
            [:td.number (get-in model [:実績 :来客数])]
            [:td.number (get-in model [:達成率 :来客数])]]
           [:tr
            [:td.title.vertical {:data-width  2 :rowspan (inc (count (:取引記録 model)))} "取引記録"]
            [:td.title.center   {:data-width  7} "取引先名"]
            [:td.title.center   {:data-width  6} "連絡先"]
            [:td.title.center   {:data-width 13} "内容"]]
           (for [rec (:取引記録 model)]
             [:tr
              [:td (:取引先名 rec)]
              [:td (:連絡先 rec)]
              [:td (:内容 rec)]])
           [:tr
            [:td.title.vertical {:data-width 2 :data-height 8} "特記事項"]
            [:td {:data-width 26 :data-height 8 :style "vertical-align: top"} (:特記事項 model)]]])

  (render sheet 31 1
          [:div
           "スケジュール"
           [:table {:data-margin-left 1 :data-margin-bottom 1}
            [:tr
             [:td {:class "header" :data-width 24} "2014年"]]
            [:tr
             (for [mon (range 1 13)]
               [:td {:class "title" :data-width 2} (str mon "月")])]
            [:tr
             [:td {:colspan 12}
              [:graphics {:data-width 24 :data-height 8}
               (for [task (range 1 11)]
                 [:box {:x (+ task 4) :y 1 :w 1 :h 1} task])]]]]]))

(def model
  {:店舗 "西新宿店"
   :担当 "川島義隆"
   :日付 (Date.)
   :計画
   {:来客数 150
    :売上高 150000
    :販売点数 1000}
   :実績
   {:来客数 2
    :売上高 1800
    :販売点数 30}
   :備考 "さっぱりでした…"

   :取引記録
   [{:取引先名 "A社"
     :連絡先   "03-XXXX-XXXX"
     :内容     "担当者の笑顔が素敵です。"}
    {:取引先名 "B社"
     :連絡先   "03-XXXX-XXXX"
     :内容     "担当者怖いです。"}
    {:取引先名 "C社"
     :連絡先   "03-XXXX-XXXX"
     :内容     "電話にでんわ。"}]

   :特記事項 "特になし\n\n\n\nですが、改行の分だけ行数が伸びます。"})

(fact "Generate hogan."
    (let [wb (create-workbook)
        sheet (to-grid (.createSheet wb "営業日報"))]
      (create-style ".title"
                    :border-style "solid"
                    :background-color "lightgreen")
      (create-style ".header"
                    :border-style "solid"
                    :color "white"
                    :background-color "seagreen")
      (create-style ".number"
                    :text-align "right")
      (create-style ".vertical"
                    :writing-mode "vertical-rl"
                    :vertical-align "top"
                    :text-align "center")
      (create-style ".center"
                    :text-align "center")
      (template sheet model)
      (with-open [out (io/output-stream "target/営業日報.xlsx")]
        (.write wb out))))

