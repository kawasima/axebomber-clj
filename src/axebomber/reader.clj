(ns axebomber.reader
  (:refer-clojure :exclude [read])
  (:require [clojure.java.io :as io]
            [clojure.string :as string]
            [axebomber.util :refer :all]
            [axebomber.style :refer :all])
  (:import [org.apache.poi.ss.usermodel Cell DateUtil FormulaError CellStyle]))

(defmulti cell-value (fn [cell]
                       (let [type (. cell getCellType)]
                         (if (= type Cell/CELL_TYPE_FORMULA)
                           (. cell getCachedFormulaResultType)
                           type))))

(defmethod cell-value Cell/CELL_TYPE_STRING [cell]
  (. cell getStringCellValue))

(defmethod cell-value Cell/CELL_TYPE_BOOLEAN [cell]
  (. cell getBooleanCellValue))

(defmethod cell-value Cell/CELL_TYPE_NUMERIC [cell]
  (if (DateUtil/isCellDateFormatted cell)
    (. cell getDateCellValue)
    (. cell getNumericCellValue)))

(defmethod cell-value Cell/CELL_TYPE_ERROR [cell]
  (-> cell (.getErrorCellValue) (FormulaError/forInt)))

(defmethod cell-value Cell/CELL_TYPE_BLANK [cell] nil)

(defn top-left? [cell]
  (let [style (.getCellStyle cell)]
    (and (or (border? cell :top)
             (border? (neighbor cell :up) :bottom))
         (or (border? cell :left)
             (border? (neighbor cell :left) :right)))))


(defn find-box [cell]
  (loop [x (.getColumnIndex cell) y (.getRowIndex cell) direction :right
         layout {:x x :y y :w 1 :h 1 :style (read-style cell)}]
    (when-let [c (get-cell (.getSheet cell) x y)]
      (case direction
        :right (if (or (border? c :top) (border? (neighbor c :up) :bottom))
                 (if (or (border? c :right) (border? (neighbor c :right) :left))
                   (recur x y :down layout)
                   (recur (inc x) y :right (update-in layout [:w] inc))))
        :down  (if (or (border? c :right) (border? (neighbor c :right) :left))
                 (if (or (border? c :bottom) (border? (neighbor c :down) :top))
                   (recur x y :left layout)
                   (recur x (inc y) :down (update-in layout [:h] inc))))
        :left  (if (or (border? c :bottom) (border? (neighbor c :down) :top))
                 (if (or (border? c :left) (border? (neighbor c :left) :right))
                   (recur x y :up layout)
                   (recur (dec x) y :left layout)))
        :up    (if (or (border? c :left) (border? (neighbor c :left) :right))
                 (if (or (border? c :top) (border? (neighbor c :up) :bottom))
                   layout
                   (recur x (dec y) :up layout)))))))

(defn same-row? [[x1 y1 w1 h1] [x2 y2 w2 h2]]
  (and (= y1 y2)
       (< x1 x2)))
;; ↑これではマルチカラムの場合にだめ。
;; テーブルの先頭行の範囲をx2 が超えてないかをチェックするひつようあり

(defn next-row? [[x1 y1 w1 h1] [x2 y2 w2 h2] row-height]
  (= (+ y1 @row-height) y2))

(defn- make-td [[x y w h style value]]
  (let [style-text (map (fn [[k v]] (str (name k) ":" v ";")) style)]
    [:td {:data-width w :data-height h :style (apply str style-text)} value]))

(defn boxes-to-tables [boxes]
  (let [tables (atom [])
        row-size (atom nil)]
    (loop [cb nil, nb (first boxes), boxes (rest boxes)]
      (if-not cb
        (do
          (swap! tables conj
                  [:table {:x (nth nb 0) :y (nth nb 1)}
                   [:tr (make-td nb)]])
          (reset! row-size (nth nb 3)))
        (if (same-row? cb nb)
          (do
            (swap! tables update-in [(->> @tables count dec) (->> @tables last count dec)]
                   conj (make-td nb))
            (swap! row-size min (nth nb 3)))
          (if (next-row? cb nb row-size)
            (do
              (swap! tables update-in [(->> @tables count dec)]
                     conj
                     [:tr (make-td nb)])
              (reset! row-size (nth nb 3)))
            (swap! tables conj
              [:table {:x (nth nb 0) :y (nth nb 1)}
                [:tr (make-td nb)]]))))
      (if (not-empty boxes)
        (recur nb (first boxes) (rest boxes))
        (map #(vector {:x (-> % second :x) :y (-> % second :y)} %) @tables)))))

(defn read-text [cell value]
  (let [style-text (map (fn [[k v]] (str (name k) ":" v ";")) (read-style cell))]
    (if-let [region (get-merged-region cell)]
      [{:x (.getColumnIndex cell) :y (.getRowIndex cell)}
        [:p value]]
      [{:x (.getColumnIndex cell) :y (.getRowIndex cell)}
        [:p value]])))

(defn read [workbook sheet-name]
  (let [sheet (.getSheet workbook sheet-name)
        parsed-cell (atom #{})
        boxes (atom [])
        exprs (atom [])]
    (loop [y (.getFirstRowNum sheet)]
      (when-let [row (.getRow sheet y)]
        (loop [x (long (.getFirstCellNum row))]
          (when-let [cell (and (not (get @parsed-cell [x y])) (get-cell sheet x y))]
            (if (top-left? cell)
              (when-let [{cx :x cy :y w :w h :h style :style} (find-box cell)]
                (let [text (atom "")]
                  (doseq [ix (range cx (+ cx w)) iy (range cy (+ cy h))]
                    (swap! text str (cell-value (get-cell sheet ix iy)))
                    (swap! parsed-cell conj [ix iy]))
                  (swap! boxes conj [cx cy w h style @text])))
              (when-let [v (cell-value cell)]
                (swap! exprs conj (read-text cell v)))))
          (when (< x (dec (.getLastCellNum row)))
            (recur (inc x)))))
      (when (< y (dec (.getLastRowNum sheet)))
        (recur (inc y))))
    (concat @exprs (boxes-to-tables @boxes))))

(defn copy-grid [src dest]
  (let [column-index (loop [y (.getFirstRowNum src) last-column 0]
                       (let [row (.getRow src y)]
                         (some-> (.getRow dest y)
                                 (.setHeightInPoints (if row (.getHeightInPoints row) (.getDefaultRowHeightInPoints src))))
                         (if (< y (dec (.getLastRowNum src)))
                           (recur (inc y) (max (if row (.getLastCellNum row) 0) last-column))
                           (max (if row (.getLastCellNum row) 0) last-column))))]
    (doseq [x (range (inc column-index))]
      (.setColumnWidth dest x (int (.getColumnWidth src x))))))
