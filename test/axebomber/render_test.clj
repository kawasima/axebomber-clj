(ns axebomber.render-test
  (:require [clojure.java.io :as io])
  (:use [axebomber usermodel render]
    [axebomber.style :only [create-style-set]]
    [midje.sweet])
  (:import [org.apache.poi.ss.usermodel CellStyle IndexedColors]
           [java.util Date]))

(defn template [sheet model]
  (render sheet 1 1
          [:table
           [:tr
            [:td {:size 28 :class "header" :text-align "center"} "営業日報"]]
           [:tr
            [:td {:size 3 :class "title"} "店舗"]
            [:td {:size 8} (:店舗 model)]
            [:td {:size 3 :class "title"} "担当"]
            [:td {:size 7} (:担当 model)]
            [:td {:size 2 :class "title"} "日付"]
            [:td {:size 5} (:日付 model)]]
           [:tr
            [:td {:size 4 :class "title"} ""]
            [:td {:size 5 :class "title"} "計画"]
            [:td {:size 5 :class "title"} "実績"]
            [:td {:size 5 :class "title"} "達成率"]
            [:td {:size 9 :class "title"} "備考"]]
           [:tr
            [:td {:class "title"} "来客数"]
            [:td {:text-align "right"} (get-in model [:計画 :来客数])]
            [:td {:text-align "right"} (get-in model [:実績 :来客数])]
            [:td {:text-align "right"} (get-in model [:達成率 :来客数])]
            [:td {:rowspan 4} (get-in model [:備考])]]
           [:tr
            [:td {:class "title"} "売上高"]
            [:td {:text-align "right"} (get-in model [:計画 :売上高])]
            [:td {:text-align "right"} (get-in model [:実績 :売上高])]
            [:td {:text-align "right"} (get-in model [:達成率 :売上高])]]
           [:tr
            [:td {:class "title"} "客単価"]
            [:td {:text-align "right"} (get-in model [:計画 :客単価])]
            [:td {:text-align "right"} (get-in model [:実績 :客単価])]
            [:td {:text-align "right"} (get-in model [:達成率 :客単価])]]
           [:tr
            [:td {:class "title"} "販売点数"]
            [:td {:text-align "right"} (get-in model [:計画 :来客数])]
            [:td {:text-align "right"} (get-in model [:実績 :来客数])]
            [:td {:text-align "right"} (get-in model [:達成率 :来客数])]]
           [:tr
            [:td {:size 2 :rowspan (inc (count (:取引記録 model)))
                  :writing-mode "vertical-rl"
                  :class "title"}
             "取引記録"]
            [:td {:size 7 :class "title" :text-align "center"} "取引先名"]
            [:td {:size 6 :class "title" :text-align "center"} "連絡先"]
            [:td {:size 13 :class "title" :text-align "center"} "内容"]]
           (for [rec (:取引記録 model)]
             [:tr
              [:td (:取引先名 rec)]
              [:td (:連絡先 rec)]
              [:td (:内容 rec)]])
           [:tr
            [:td {:size 2 :class "title" :writing-mode "vertical-rl" :rowspan 8} "特記事項"]
            [:td {:size 26 :rowspan 8} (:特記事項 model)]]])

  (render sheet 31 1
          [:div
           "スケジュール"
           [:table {:margin-left 1 :margin-bottom 1}
            [:tr
             [:td {:class "header" :size 24} "2014年"]]
            [:tr
             (for [mon (range 1 13)]
               [:td {:class "title" :size 2} (str mon "月")])]
            [:tr
             [:td {:colspan 12}
              [:graphics {:size 24 :height 8}
               (for [task (range 1 11)]
                 [:box {:x (+ task 4) :y 1
                        :width 1 :height 1} task])]]]]]))

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
      (create-style-set wb :default :border-type CellStyle/BORDER_THIN)
      (create-style-set wb "title"
                        :border-type CellStyle/BORDER_THIN
                        :background-color (.getIndex IndexedColors/LIGHT_GREEN))
      (create-style-set wb "header"
                        :border-type CellStyle/BORDER_THIN
                        :color (.getIndex IndexedColors/WHITE)
                        :background-color (.getIndex IndexedColors/SEA_GREEN))
      (template sheet model)
      (with-open [out (io/output-stream "target/営業日報.xlsx")]
        (.write wb out))))

