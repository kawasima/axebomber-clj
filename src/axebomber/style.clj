(ns axebomber.style
  (:require [clojure.string :as string]
            [clojure.set])
  (:use [axebomber.util])
  (:import [org.apache.poi.xssf.usermodel XSSFWorkbook]
           [org.apache.poi.ss.util CellUtil CellRangeAddress]
           [org.apache.poi.ss.usermodel Font CellStyle IndexedColors]))

(def style-sets (atom {}))
(def cell-style-cache (atom {}))

(def ^:const border-styles
  {"none"   CellStyle/BORDER_NONE
   "solid"  CellStyle/BORDER_THIN
   "double" CellStyle/BORDER_DOUBLE
   "dashed" CellStyle/BORDER_DASHED
   "dotted" CellStyle/BORDER_DOTTED})

(def ^:const text-aligns
  {"center" CellStyle/ALIGN_CENTER
   "left"   CellStyle/ALIGN_LEFT
   "right"  CellStyle/ALIGN_RIGHT})

(def ^:const vertical-aligns
  {"top"    CellStyle/VERTICAL_TOP
   "middle" CellStyle/VERTICAL_CENTER
   "bottom" CellStyle/VERTICAL_BOTTOM})

(def ^:const writing-modes
  {"vertical-rl" 0xFF
   "horizontal-tb" 0})

(def colors
  {"aqua"   (.getIndex IndexedColors/AQUA)
   "black"  (.getIndex IndexedColors/BLACK)
   "blue"   (.getIndex IndexedColors/BLUE)
   "bluegrey" (.getIndex IndexedColors/BLUE_GREY)
   "brightgreen" (.getIndex IndexedColors/BRIGHT_GREEN)
   "brown"  (.getIndex IndexedColors/BROWN)
   "coral"  (.getIndex IndexedColors/CORAL)
   "cornflowerblue" (.getIndex IndexedColors/CORNFLOWER_BLUE)
   "darkblue" (.getIndex IndexedColors/DARK_BLUE)
   "darkgreen" (.getIndex IndexedColors/DARK_GREEN)
   "darkred" (.getIndex IndexedColors/DARK_RED)
   "darkteal" (.getIndex IndexedColors/DARK_TEAL)
   "darkyellow" (.getIndex IndexedColors/DARK_YELLOW)
   "gold"   (.getIndex IndexedColors/GOLD)
   "green"  (.getIndex IndexedColors/GREEN)
   "grey25" (.getIndex IndexedColors/GREY_25_PERCENT)
   "grey40" (.getIndex IndexedColors/GREY_40_PERCENT)
   "grey50" (.getIndex IndexedColors/GREY_50_PERCENT)
   "grey80" (.getIndex IndexedColors/GREY_80_PERCENT)
   "indigo" (.getIndex IndexedColors/INDIGO)
   "lavender" (.getIndex IndexedColors/LAVENDER)
   "lemonchiffon" (.getIndex IndexedColors/LEMON_CHIFFON)
   "lightblue" (.getIndex IndexedColors/LIGHT_BLUE)
   "lightcornflowerblue" (.getIndex IndexedColors/LIGHT_CORNFLOWER_BLUE)
   "lightgreen" (.getIndex IndexedColors/LIGHT_GREEN)
   "lightorange" (.getIndex IndexedColors/LIGHT_ORANGE)
   "lightturquoise" (.getIndex IndexedColors/LIGHT_TURQUOISE)
   "lightyellow" (.getIndex IndexedColors/LIGHT_YELLOW)
   "lime"   (.getIndex IndexedColors/LIME)
   "maroon" (.getIndex IndexedColors/MAROON)
   "olivegreen" (.getIndex IndexedColors/OLIVE_GREEN)
   "orange" (.getIndex IndexedColors/ORANGE)
   "orchid" (.getIndex IndexedColors/ORCHID)
   "paleblue" (.getIndex IndexedColors/PALE_BLUE)
   "pink"   (.getIndex IndexedColors/PINK)
   "red"    (.getIndex IndexedColors/RED)
   "rose"   (.getIndex IndexedColors/ROSE)
   "royalblue" (.getIndex IndexedColors/ROYAL_BLUE)
   "seagreen"  (.getIndex IndexedColors/SEA_GREEN)
   "skyblue"   (.getIndex IndexedColors/SKY_BLUE)
   "tan"    (.getIndex IndexedColors/TAN)
   "teal"   (.getIndex IndexedColors/TEAL)
   "turquoise" (.getIndex IndexedColors/TURQUOISE)
   "violet" (.getIndex IndexedColors/VIOLET)
   "white"  (.getIndex IndexedColors/WHITE)
   "yellow" (.getIndex IndexedColors/YELLOW)})

(def colors-invert (clojure.set/map-invert colors))

(def ^:const font-weights
  {"bold" Font/BOLDWEIGHT_BOLD
   "normal" Font/BOLDWEIGHT_NORMAL})

(defn font-map [font]
  {:weight (.getBoldweight font)
   :height (.getFontHeight font)
   :name   (.getFontName font)
   :color  (.getColor font)
   :italic? (.getItalic font)
   :strikeout? (.getStrikeout font)
   :offset (.getTypeOffset font)
   :underline (.getUnderline font)})

(defn- get-cell-style [sheet style style-index]
  (if-let [cell-style (get-in @cell-style-cache [sheet style style-index])]
    cell-style
    (let [cs (.. sheet getWorkbook createCellStyle)
          bs (get border-styles (get style :border-style "none"))
          default-font (font-map (.. sheet getWorkbook (getFontAt (short 0))))
          font (atom default-font)]
      (when (not= (bit-and style-index 1) 0)
        (.setBorderTop cs bs))
      (when (not= (bit-and style-index 2) 0)
        (.setBorderRight cs bs))
      (when (not= (bit-and style-index 4) 0)
        (.setBorderBottom cs bs))
      (when (not= (bit-and style-index 8) 0)
        (.setBorderLeft cs bs))
      (when-let [background-color (:background-color style)]
        (.setFillForegroundColor cs (get colors background-color (.getIndex IndexedColors/WHITE)))
        (.setFillPattern cs CellStyle/SOLID_FOREGROUND))
      (when-let [align (:text-align style)]
        (.setAlignment cs (get text-aligns align "left")))

      (when-let [mode (:writing-mode style)]
        (.setRotation cs (get writing-modes mode "horizontal-tb")))

      (when-let [vertical-align (:vertical-align style)]
        (.setVerticalAlignment cs (get vertical-aligns vertical-align CellStyle/VERTICAL_TOP)))

      (when-let [color (:color style)]
        (swap! font assoc :color (get colors color (.getIndex IndexedColors/BLACK))))

      (when-let [font-style (:font-style style)]
        (swap! font assoc :italic? (= font-style "italic")))

      (when-let [font-size (:font-size style)]
        (swap! font assoc :height (short (* 20 (Integer/parseInt (str font-size))))))

      (when-let [font-weight (:font-weight style)]
        (swap! font assoc :weight (get font-weights font-weight Font/BOLDWEIGHT_NORMAL)))

      (when (not= default-font @font)
        (if-let [defined-font (.. sheet getWorkbook
                              (findFont (:weight @font) (:color @font) (:height @font) (:name @font)
                                (:italic? @font) (:strikeout? @font) (:offset @font) (:underline @font)))]
        (.setFont cs defined-font)
        (let [new-font (.. sheet getWorkbook createFont)]
          (doto new-font
            (.setBoldweight (:weight @font))
            (.setFontHeight (:height @font))
            (.setColor      (:color  @font))
            (.setFontName   (:name   @font))
            (.setItalic     (:italic? @font))
            (.setStrikeout  (:strikeout? @font))
            (.setTypeOffset (:offset @font))
            (.setUnderline  (:underline @font)))
          (.setFont cs new-font))))

      (swap! cell-style-cache assoc-in [sheet style style-index] cs)
      cs)))

(defn- apply-style-each-cells [sheet x y w h style]
  (doseq [cx (range x (+ x w)) cy (range y (+ y h))]
    (let [style-index (+ (if (= cy y) 1 0)
                         (if (= cx (+ x w -1)) 2 0)
                         (if (= cy (+ y h -1)) 4 0)
                         (if (= cx x) 8 0))
          cell (get-cell sheet cx cy)]
      (when (= (.getCellStyle cell)
               (get-cell-style sheet (:default style-sets) 0))
        (.setCellStyle cell
                       (get-cell-style sheet style style-index))))))

(defn- apply-style-merged-cells [sheet x y w h style]
  (let [cell (get-cell sheet x y)]
    (.addMergedRegion sheet (CellRangeAddress. y (+ y h -1) x (+ x w -1)))
    (apply-style-each-cells sheet x y w h style)))

(defn- merge-style [selector attrs]
  (let [selector-style (or (get @style-sets selector) (:default @style-sets))
        class-styles (some->> (string/split (or (get attrs :class) "") #"\s+")
                             (filter not-empty)
                             (map #(get @style-sets (str "." %))))
        attr-style (->> (string/split (or (get attrs :style) "") #"\s*;\s*")
                        (filter not-empty)
                        (map #(string/split % #"\s*:\s*" 2))
                        (reduce #(assoc %1 (keyword (first %2)) (second %2)) {}))]
    (-> (apply merge selector-style class-styles)
        (merge attr-style))))

(defn apply-style [selector sheet x y w h attrs]
  (let [merged-style (merge-style selector attrs)
        height (max (+ h (get attrs :rowspan 1) -1)
                    (get attrs :data-height 1))]
    (if (or (not= (get merged-style :text-align "left") "left")
            (not= (get attrs :writing-mode "horizontal-tb") "horizontal-tb"))
      (apply-style-merged-cells sheet x y w
                                height
                                merged-style)
      (apply-style-each-cells sheet x y w height merged-style))))

(defn get-location-styles [selector attrs]
  (let [merged-style (merge-style selector attrs)]
    (-> merged-style)))

(defn create-style [selector & {:as style}]
  (swap! style-sets assoc selector style))

(create-style :default :border-style "none")
(create-style "td" :border-style "solid")
(create-style "h1" :font-weight "bold" :font-size 18)
(create-style "h2" :font-weight "bold")
(create-style "h3" :font-weight "bold")
(create-style "box" :border-style "solid" :text-align "center" :vertical-align "middle")

(defn read-style [cell]
  (let [style (.getCellStyle cell)]
    (merge
      (condp = (.getAlignment style)
        CellStyle/ALIGN_RIGHT {:text-align "right"}
        CellStyle/ALIGN_CENTER {:text-align "center"}
        nil)
      (condp = (.getVerticalAlignment style)
        CellStyle/VERTICAL_CENTER {:vertical-align "middle"}
        CellStyle/VERTICAL_BOTTOM {:vertical-align "bottom"}
        nil)
      (case (.getRotation style)
        0xFF {:writing-mode "vertical-rl"}
        nil)
      (when-let [font (some-> cell
                        (.getSheet)
                        (.getWorkbook)
                        (.getFontAt (short (.getFontIndex style))))]
        (when-not (= (.getColor font) (.getIndex IndexedColors/AUTOMATIC))
          {:color (get colors-invert (.getColor font))}))
      (when-let [bg-color (get colors-invert (.getFillForegroundColor style))]
        {:background-color bg-color}))))

