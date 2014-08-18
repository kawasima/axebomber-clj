(ns axebomber.renderer.excel
  (:require [clj-time.format :as time-fmt]
            [clj-time.coerce :as c]
            [clojure.string :as string])
  (:use [clojure.walk :only [prewalk]]
        [axebomber util style])
  (:import [java.awt Font]
           [org.apache.poi.ss.util CellUtil SheetUtil]
           [org.apache.poi.ss.usermodel CellStyle IndexedColors]))

(declare render)

(def ^{:doc "Regular expression that parses a CSS-style id and class from an element name."
       :private true}
  re-tag #"([^\s\.#]+)(?:#([^\s\.#]+))?(?:\.([^\s#]+))?")

(def time-formatter (time-fmt/formatter "yy年MM月dd日"))

(def ^:const default-column-width 5)

(defn- literal? [expr]
  (or (string? expr)
      (keyword? expr)
      (number? expr)
      (instance? java.util.Date expr)))

(defn- merge-attributes [{:keys [id class]} map-attrs]
  (->> map-attrs
       (merge (if id {:id id}))
       (merge-with #(if %1 (str %1 " " %2) %2) (if class {:class class}))))

(defn normalize-element
  "Ensure an element vector is of the form [tag-name attrs content]."
  [[tag & content]]
  (when (not (or (keyword? tag) (symbol? tag) (string? tag)))
    (throw (IllegalArgumentException. (str tag " is not a valid element name."))))
  (let [[_ tag id class] (re-matches re-tag (as-str tag))
        tag-attrs        {:id id
                          :class (if class (.replace ^String class "." " "))}
        map-attrs        (first content)]
    (if (map? map-attrs)
      [tag (merge-attributes tag-attrs map-attrs) (next content)]
      [tag tag-attrs content])))

(defn extend-cell [sheet x y w h]
  (let [cell (get-cell sheet x y)
        style (read-style cell)
        [[_ y] _ [x _] _] (in-box cell)]
    (when-let [merged (get-merged-region cell)]
      (unmerge-region cell))
    (apply-style "td" sheet x y w h style)))

(defn correct-td-position [sheet x y & {:keys [height]}]
  (loop [cx x]
    (if-let [[top bottom _ right] (in-box (get-cell sheet cx y))]
      (do
        (when (and height (> height 1))
          (extend-cell sheet cx y (inc (- (first right) cx)) (- (second bottom) (second top) (- height))))
        (recur (inc (first right))))
      cx)))

(defn- inherit-size [sheet x y & {:keys [colspan] :or {colspan 1}}]
  (if (or (<= y 0)
          (= (.. (get-cell sheet x (dec y)) getCellStyle getBorderLeft) CellStyle/BORDER_NONE))
    default-column-width
    (loop [cx x, cols 1]
      (let [cell (get-cell sheet cx (dec y))]
        (if-let [merged-cell (get-merged-region cell)]
          (if (>= cols colspan)
            (- (.getLastColumn merged-cell) (.getFirstColumn merged-cell) -1)
            (recur (+ (.getLastColumn merged-cell) 1) (inc cols)))
          (if (not= (.. cell getCellStyle getBorderRight) CellStyle/BORDER_NONE)
            (if (>= cols colspan)
              (- cx x -1)
              (recur (inc cx) (inc cols)))
            (if (> cx 255)
              (- cx x -1)
              (recur (inc cx) cols))))))))

(defn filter-children
  ([content tag]
   (filter-children content tag []))
  ([content tag filtered]
   (cond
    (vector? content) (if (= (name tag) (name (first content))) (conj filtered content) filtered)
    (seq? content) (->> content
                        (map #(filter-children % tag filtered))
                        (reduce into))
    :else filtered)))


;; TODO Sould determine the default data-width.
(defn render-literal [sheet {:keys [x y data-width]} lit]
  (let [lit (if (instance? java.util.Date lit)
              (time-fmt/unparse time-formatter (c/from-date lit)) lit)
        font-index (.. (get-cell sheet x y) getCellStyle getFontIndex)
        font (.. (.getWorkbook sheet) (getFontAt font-index))]
    [(or data-width 1)
     (->> (string/split (str lit) #"(\r\n|\r|\n)")
          (map #(split-by-width %
                  (if data-width
                    (reduce + (width-range sheet x (+ x data-width)))
                    65536)
                 font))
          (reduce into)
          (map-indexed #(.setCellValue (get-cell sheet x (+ y %1)) %2))
          (count))
     lit]))

(defn render-horizontal [sheet {x :x y :y :as ctx} [tag attrs content]]
  (let [max-height (atom 0)
        cy (+ y (get attrs :data-margin-top 0))]
    (loop [cx (+ x (get attrs :data-margin-left 0)), content content, children []]
      (let [[w h child] (render sheet (assoc ctx :x cx :y cy) (first content)
                                :direction :horizontal)
             cx (+ cx w)]
        (reset! max-height (max h @max-height))
        (if (not-empty (rest content))
          (recur cx (rest content) (conj children child))
          [(- cx x) @max-height (seq (conj children child))])))))

(defn render-vertical [sheet {x :x y :y :as ctx} [tag attrs content]]
  (let [max-width (atom 0)
        cx (+ x (get attrs :data-margin-left 0))]
    (loop [cy (+ y (get attrs :data-margin-top 0)), content content, children []]
      (let [[w h child] (render sheet (assoc ctx :x cx :y cy) (first content)
                                :direction :vertical)
             cy (+ cy h)]
        (reset! max-width (max w @max-width))
        (if (not-empty (rest content))
          (recur cy (rest content) (conj children child))
          [@max-width
           (- cy y (- (get attrs :data-margin-bottom 0)))
           (seq (conj children child))])))))

(defmulti render-tag (fn [sheet ctx tag & rst] tag))

(defmethod render-tag "table" [sheet ctx tag attrs content]
  (render-vertical sheet ctx [tag attrs content]))

(defmethod render-tag "tr" [sheet {x :x y :y :as ctx} tag attrs content]
  (let [[w h td-tags] (render-horizontal sheet ctx [tag attrs content])
        td-tags (filter-children (seq td-tags) "td")]
    (loop [cx x, idx 0, row-height 65536]
      (let [cx (correct-td-position sheet cx y :height h)
            [td-tag td-attrs _] (nth td-tags idx)
            size (get td-attrs :data-width 3)
            height (get td-attrs :data-height h)]
        (apply-style "td" sheet cx y size height td-attrs)
        (if (< idx (dec (count td-tags)))
          (recur (+ cx size) (inc idx) (min row-height height))
          [w (min row-height height) ["tr" attrs td-tags]])))))


(defmethod render-tag "td" [sheet {x :x y :y :as ctx} tag attrs content]
  (let [cx (correct-td-position sheet x y)
        size (or (and (:colspan attrs) (inherit-size sheet cx y :colspan (:colspan attrs)))
                  (:data-width attrs)
                  (inherit-size sheet cx y))
        attrs (assoc (merge ctx attrs) :data-width size)
        [w h child] (render sheet (assoc attrs :x cx) content)]
    [(+ size (- cx x)) h ["td" attrs child]]))

(defmethod render-tag "box" [sheet {px :x py :y :as ctx} tag {:keys [x y w h] :as attrs} content]
  (let [[cw ch child] (render sheet (assoc ctx :x (+ px x) :y (+ py y)) content)]
    (apply-style "box" sheet (+ px x) (+ py y) w h attrs)
    [cw ch child]))

(defmethod render-tag "graphics" [sheet {x :x y :y :as ctx} tag attrs content]
  (let [w (:data-width attrs)
        h (:data-height attrs)]
    (doseq [cont content]
      (render sheet ctx cont :direction :none))
    [w h nil]))


(defmethod render-tag "ul" [sheet {x :x y :y :as ctx} tag attrs content]
  (let [list-style-type (get attrs :list-style-type "・")
        [w h children] (render-vertical sheet ctx [tag attrs content])]
    (->> children
         (tree-seq sequential? seq)
         (filter #(and (vector? %) (= (first %) "li")))
         (map #(render-literal sheet (second %) list-style-type))
         doall)
    [w h children]))

(defmethod render-tag "ol" [sheet {x :x y :y :as ctx} tag attrs content]
  (let [[w h children] (render-vertical sheet ctx [tag attrs content])]
    (->> children
         (tree-seq sequential? seq)
         (filter #(and (vector? %) (= (first %) "li")))
         (map-indexed #(render-literal sheet (second %2) (inc %1)))
         doall)
    [w h children]))

(defmethod render-tag "li" [sheet {x :x y :y :as ctx} tag attrs content]
  (let [[w h children] (render sheet (assoc ctx :x (inc x)) content)]
    [w h [tag (merge attrs {:x x :y y}) content]]))

(defmethod render-tag "dd" [sheet {:keys [x y dt-width] :or {dt-width 3} :as ctx} tag attrs content]
  (render sheet (assoc ctx :x (+ x dt-width)) content))

(defmethod render-tag "dt" [sheet {:keys [x y] :as ctx} tag attrs content]
  (let [[w h children] (render sheet ctx content)
        cell (get-cell sheet x y)]
    (.setCellValue cell (str (.getStringCellValue cell) ":"))
    [w 0 children]))

(defmethod render-tag "dl" [sheet {x :x y :y :as ctx} tag attrs content]
  (let [font-index (.. (get-cell sheet x y) getCellStyle getFontIndex)
        font (.. (.getWorkbook sheet) (getFontAt font-index))
        title-length (->> (filter-children content "dt")
                          (map last)
                          (map #(string-width % font))
                          (apply max))]
    (render sheet (assoc ctx :dt-width (inc (Math/floor (/ title-length 20)))) content)))

(defmethod render-tag "img" [sheet {:keys [x y src] :as ctx} tag attrs content]
  (let [baos (ByteArrayOutputStream.)
        img  (ImageIO/read src)]
    (ImageIO/write img "png" baos)))

(defmethod render-tag :default
  [sheet ctx tag attrs content]
  (render-vertical sheet ctx [tag attrs content]))

(defn- element-render-strategy
  "Returns the compilation strategy to use for a given element."
  [sheet ctx [tag attrs & content]]
  (cond
    (every? literal? (list tag attrs))
      ::all-literal                    ; e.g. [:span "foo"]
    (and (literal? tag) (map? attrs))
      ::literal-tag-and-attributes     ; e.g. [:span {} x]
    (literal? tag)
      ::literal-tag                    ; e.g. [:span x]
    :else
      ::default))

(defmulti render-element element-render-strategy)

(defmethod render-element ::all-literal
  [sheet ctx [tag & literal]]
  (let [[tag tag-attrs _] (normalize-element [tag])]
    (render-tag sheet ctx tag tag-attrs literal)))

(defmethod render-element ::literal-tag-and-attributes
  [sheet ctx [tag attrs & content]]
  (let [[tag attrs _] (normalize-element [tag attrs])]
    (render-tag sheet ctx tag attrs content)))

(defmethod render-element ::literal-tag
  [sheet ctx [tag & content]]
  (let [[tag tag-attrs _] (normalize-element [tag])]
    (render-tag sheet ctx tag tag-attrs content)))

(defn- merge-attributes [{:keys [id class]} map-attrs]
  (->> map-attrs
       (merge (if id {:id id}))
       (merge-with #(if %1 (str %1 " " %2) %2) (if class {:class class}))))

(defn render-seq [sheet ctx content options]
  (case (get options :direction :vertical)
    :vertical   (render-vertical sheet ctx [:dummy {} content])

    :horizontal (render-horizontal sheet ctx [:dummy {} content])
    (reduce #(identity [(max (nth %1 0) (nth %2 0))
                        (max (nth %1 1) (nth %2 1))
                        (conj (nth %1 2) (nth %2 2))])
      [0 0 []]
      (for [cont content]
        (render sheet ctx cont)))))

(defn render [sheet {x :x y :y :as context} expr & {:as options}]
  (cond
   (vector? expr) (render-element sheet context expr)
   (literal? expr) (render-literal sheet context expr)
   (seq? expr) (render-seq sheet context expr options)
   (nil? expr) [1 1 ""]
   :else (render-literal sheet context (str expr))))

