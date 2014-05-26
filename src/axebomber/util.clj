(ns axebomber.util
  (:import [org.apache.poi.ss.util CellUtil]
           [org.apache.poi.ss.usermodel CellStyle Font]
           [java.awt Toolkit]
           [java.awt.font TextAttribute FontRenderContext TextLayout]
           [java.text AttributedString]))

(def ^:dynamic *base-url* nil)

(def dpi (.getScreenResolution (Toolkit/getDefaultToolkit)))
(def font-render-context (FontRenderContext. nil (boolean true) (boolean true)))


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

(defn unmerge-region
  "Remove merged region that contains a given cell."
  [cell]
  (let [sheet (.getSheet cell)
        region-indexes (->> (range (.getNumMergedRegions sheet))
                            (map #(vector % (.getMergedRegion sheet %)))
                            (filter #(.isInRange (second %) (.getRowIndex cell) (.getColumnIndex cell)))
                            (map first))]
    (doseq [idx region-indexes]
      (.removeMergedRegion sheet idx))))

(defn border? [cell pos]
  (when-let [style (some-> cell (.getCellStyle))]
    (case pos
      :top    (not= (.getBorderTop style) CellStyle/BORDER_NONE)
      :bottom (not= (.getBorderBottom style) CellStyle/BORDER_NONE)
      :left   (not= (.getBorderLeft style) CellStyle/BORDER_NONE)
      :right  (not= (.getBorderRight style) CellStyle/BORDER_NONE))))

(defn in-box [cell]
  (let [x (.getColumnIndex cell) y (.getRowIndex cell)
        sheet (.getSheet cell)
        box-pos (for [direction [:top :bottom :left :right]]
                  (loop [cx x, cy y]
                    (if (border? (get-cell sheet cx cy) direction)
                      [cx cy]
                      (when (and
                             (< 0 cx (dec (.. sheet (getRow cy) (getLastCellNum))))
                             (< 0 cy (inc (.. sheet (getLastRowNum)))))
                        (recur (case direction :left (dec cx) :right (inc cx) cx)
                               (case direction :top (dec cy) :bottom (inc cy) cy))))))]
    (when-not (some nil? box-pos)
      box-pos)))


(defn neighbor [cell direction]
  (let [x (.getColumnIndex cell) y (.getRowIndex cell)]
    (case direction
      :up    (when (> y 0)
               (some-> (.getSheet cell) (.getRow (dec y)) (.getCell x)))
      :down  (some-> (.getSheet cell) (.getRow (inc y)) (.getCell x))
      :left  (when (> x 0)
               (some-> (.getSheet cell) (.getRow y) (.getCell (dec x))))
      :right (some-> (.getSheet cell) (.getRow y) (.getCell (inc x))))))

(defn- make-attributed-string [s font]
  (let [txt (AttributedString. s)
        len (.length s)]
    (doto txt
      (.addAttribute TextAttribute/FAMILY (.getFontName font) 0 len)
      (.addAttribute TextAttribute/SIZE (.getFontHeightInPoints font) 0 len))
    (when (= (.getBoldweight font) Font/BOLDWEIGHT_BOLD)
      (.addAtribute TextAttribute/WEIGHT TextAttribute/WEIGHT_BOLD 0 len))
    (when (.getItalic font)
      (.addAtribute TextAttribute/POSTURE TextAttribute/POSTURE_OBLIQUE 0 len))
    (when (= (.getUnderline font) Font/U_SINGLE)
      (.addAtribute TextAttribute/UNDERLINE TextAttribute/UNDERLINE_ON 0 len))
    txt))

(defn string-width [s font]
  (let [layout (TextLayout. (.getIterator (make-attributed-string s font))
                            font-render-context)]
    (float (* (.. layout getBounds getWidth) (/ dpi 72)))))

(defn bsearch-char-count
  ([char-seq width font]
   (bsearch-char-count char-seq 0 (count char-seq) width font))
  ([char-seq l u width font]
   (if (< u 2) 1
     (let [m (quot (+ l u) 2)
           len (string-width (apply str (take m char-seq)) font)]
       (cond
        (>= l u) m
        (> (- width len) 1) (recur char-seq (inc m) u width font)
        (< width len) (recur char-seq l (dec m) width font))))))

(defn width-range [sheet from to]
  (let [wb (.getWorkbook sheet)
        char-width (string-width "0" (.getFontAt wb (short 0)))
        margin (- (string-width "00" (.getFontAt wb (short 0))) (* char-width 2))]
    (map #(+ (* (quot (.getColumnWidth sheet %) 256) (+ char-width margin)) margin) (range from to))))

(defn split-by-width [s width font]
  (if (empty? s) [""]
    (loop [char-seq (seq s) splitted []]
      (let [idx (bsearch-char-count char-seq
                  width
                  font)
             splitted (conj splitted (apply str (take idx char-seq)))]
        (if (= (count char-seq) idx)
          splitted
          (recur (drop idx char-seq) splitted))))))


