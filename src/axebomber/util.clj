(ns axebomber.util
  (:import [org.apache.poi.ss.util CellUtil]
           [org.apache.poi.ss.usermodel CellStyle]
           [java.awt Font Toolkit Canvas]))

(def ^:dynamic *base-url* nil)
(def canvas (Canvas.))
(def dpi (.getScreenResolution (Toolkit/getDefaultToolkit)))

(defmacro with-base-url
  "Sets a base URL that will be prepended onto relative URIs. Note that for this
  to work correctly, it needs to be placed outside the html macro."
  [base-url & body]
  `(binding [*base-url* ~base-url]
     ~@body))

(defprotocol ToString
  (^String to-str [x] "Convert a value into a string."))

(extend-protocol ToString
  clojure.lang.Keyword
  (to-str [k] (name k))
  clojure.lang.Ratio
  (to-str [r] (str (float r)))
  java.net.URI
  (to-str [u]
    (if (or (.getHost u)
            (nil? (.getPath u))
            (not (-> (.getPath u) (.startsWith "/"))))
      (str u)
      (let [base (str *base-url*)]
        (if (.endsWith base "/")
          (str (subs base 0 (dec (count base))) u)
          (str base u)))))
  Object
  (to-str [x] (str x))
  nil
  (to-str [_] ""))

(defn ^String as-str
  "Converts its arguments into a string using to-str."
  [& xs]
  (apply str (map to-str xs)))

(defn get-cell
  "get a cell located by (x, y)"
  [sheet x y]
  (when (and (>= x 0) (>= y 0))
    (CellUtil/getCell (CellUtil/getRow y sheet) x)))

(defn get-merged-region [cell]
  (let [sheet (.getSheet cell)]
    (some #(when (.isInRange % (.getRowIndex cell) (.getColumnIndex cell)) %)
          (map #(.getMergedRegion sheet %) (range (. sheet getNumMergedRegions))))))

(defn neighbor [cell direction]
  (let [x (.getColumnIndex cell) y (.getRowIndex cell)]
    (case direction
      :up    (when (> y 0)
               (some-> (.getSheet cell) (.getRow (dec y)) (.getCell x)))
      :down  (some-> (.getSheet cell) (.getRow (inc y)) (.getCell x))
      :left  (when (> x 0)
               (some-> (.getSheet cell) (.getRow y) (.getCell (dec x))))
      :right (some-> (.getSheet cell) (.getRow y) (.getCell (inc x))))))

(defn border? [cell pos]
  (when-let [style (some-> cell (.getCellStyle))]
    (case pos
      :top    (not= (.getBorderTop style) CellStyle/BORDER_NONE)
      :bottom (not= (.getBorderBottom style) CellStyle/BORDER_NONE)
      :left   (not= (.getBorderLeft style) CellStyle/BORDER_NONE)
      :right  (not= (.getBorderRight style) CellStyle/BORDER_NONE))))

(defn string-width
  "Calculate the width of a given string."
  [s font]
  (let [metrics (.getFontMetrics canvas font)]
    (float (* (.stringWidth metrics s) (/ dpi 72)))))
