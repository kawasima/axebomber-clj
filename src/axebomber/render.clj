(ns axebomber.render
  (:require [clj-time.format :as time-fmt]
            [clj-time.coerce :as c])
  (:use [axebomber.util]
        [axebomber.style])
  (:import [org.apache.poi.ss.util CellUtil]
           [org.apache.poi.ss.usermodel CellStyle IndexedColors]))

(declare render)

(def ^{:doc "Regular expression that parses a CSS-style id and class from an element name."
       :private true}
  re-tag #"([^\s\.#]+)(?:#([^\s\.#]+))?(?:\.([^\s#]+))?")

(def time-formatter (time-fmt/formatter "yy年MM月dd日"))

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

(defn correct-td-position [sheet x y]
  (loop [cx x]
    (if-let [merged-region (get-merged-region (get-cell sheet cx y))]
      (do
        (recur (inc (.getLastColumn merged-region))))
      cx)))

(defn- inherit-size [sheet x y & {:keys [colspan] :or {colspan 1}}]
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
            (recur (inc cx) cols)))))))

(defn render-literal [sheet x y lit]
  (let [lit (if (instance? java.util.Date lit)
              (time-fmt/unparse time-formatter (c/from-date lit)) lit)
        lines (-> (map-indexed
                   (fn [idx v]
                     (let [c (get-cell sheet x (+ y idx))]
                       (.setCellValue c v)))
                   (clojure.string/split (str lit) #"\n"))
                count)]
    [1 lines lit]))

(defn render-horizontal [sheet x y [tag attrs content]]
  (let [max-height (atom 0)
        cy (+ y (get attrs :margin-top 0))]
    (loop [cx (+ x (get attrs :margin-left 0)), content content, children []]
      (let [[w h child] (render sheet cx cy (first content)
                                :direction :horizontal)
             cx (+ cx w)]
        (reset! max-height (max h @max-height))
        (if (not-empty (rest content))
          (recur cx (rest content) (conj children child))
          [(- cx x) @max-height (conj children child)])))))

(defn render-vertical [sheet x y [tag attrs content]]
  (let [max-width (atom 0)
        cx (+ x (get attrs :margin-left 0))]
    (loop [cy (+ y (get attrs :margin-top 0)), content content, children []]
      (let [[w h child] (render sheet cx cy (first content)
                                :direction :vertical)
             cy (+ cy h)]
        (reset! max-width (max w @max-width))
        (if (not-empty (rest content))
          (recur cy (rest content) (conj children child))
          [@max-width
           (- cy y (- (get attrs :margin-bottom 0)))
           (conj children child)])))))

(defmulti render-tag (fn [sheet x y tag & rst] tag))

(defmethod render-tag "table" [sheet x y tag attrs content]
  (render-vertical sheet x y [tag attrs content]))

(defmethod render-tag "tr" [sheet x y tag attrs content]
  (let [[w h td-tags] (render-horizontal sheet x y [tag attrs content])]
    (loop [cx x, idx 0]
      (let [cx (correct-td-position sheet cx y)
            [td-tag td-attrs _] (nth td-tags idx)
            size (get td-attrs :size 3)]
        (apply-style sheet cx y size h td-attrs)
        (if (< idx (dec (count td-tags)))
          (recur (+ cx size) (inc idx))
          [w h [:tr attrs td-tags]])))))

(defmethod render-tag "td" [sheet x y tag attrs content]
  (let [cx (correct-td-position sheet x y)
        [w h child] (render sheet cx y content)
        size (or (and (:colspan attrs) (inherit-size sheet cx y :colspan (:colspan attrs)))
                  (:size attrs)
                  (inherit-size sheet cx y))
        attrs (assoc attrs :size size)]
    [(+ size (- cx x)) h [:td attrs child]]))

(defmethod render-tag "box" [sheet px py tag {:keys [x y width height] :as attrs} content]
  (let [[w h child] (render sheet (+ px x) (+ py y) content)]
    (apply-style sheet (+ px x) (+ py y) width height attrs)
    [w h child]))

(defmethod render-tag "graphics" [sheet x y tag attrs content]
  (let [w (:size attrs)
        h (:height attrs)]
    (doseq [cont content]
      (render sheet x y cont))
    [w h nil]))

(defmethod render-tag "ul" [sheet x y tag attrs content]
  (let [list-style-type (get attrs :list-style-type "・")
        [w h children] (render-vertical sheet x y [tag attrs content])]
    (->> children
         (tree-seq sequential? seq)
         (filter #(and (vector? %) (= (first %) "li")))
         (map #(render-literal sheet (:x (second %)) (:y (second %)) list-style-type))
         doall)
    [w h children]))

(defmethod render-tag "ol" [sheet x y tag attrs content]
  (let [list-style-type (get attrs :list-style-type "・")
        [w h children] (render-vertical sheet x y [tag attrs content])]
    (->> children
         (tree-seq sequential? seq)
         (filter #(and (vector? %) (= (first %) "li")))
         (map-indexed #(render-literal sheet (:x (second %2)) (:y (second %2)) (inc %1)))
         doall)
    [w h children]))

(defmethod render-tag "li" [sheet x y tag attrs content]
  (let [[w h children] (render sheet (inc x) y content)]
    [w h [tag (merge attrs {:x x :y y}) content]]))

(defmethod render-tag :default
  [sheet x y tag attrs content]
  (render-vertical sheet x y [tag attrs content]))

(defn- element-render-strategy
  "Returns the compilation strategy to use for a given element."
  [sheet x y [tag attrs & content]]
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
  [sheet x y [tag & literal]]
  (let [[tag tag-attrs _] (normalize-element [tag])]
    (render-tag sheet x y tag tag-attrs literal)))

(defmethod render-element ::literal-tag-and-attributes
  [sheet x y [tag attrs & content]]
  (let [[tag attrs _] (normalize-element [tag attrs])]
    (render-tag sheet x y tag attrs content)))

(defmethod render-element ::literal-tag
  [sheet x y [tag & content]]
  (let [[tag tag-attrs _] (normalize-element [tag])]
    (render-tag sheet x y tag tag-attrs content)))

(defn- merge-attributes [{:keys [id class]} map-attrs]
  (->> map-attrs
       (merge (if id {:id id}))
       (merge-with #(if %1 (str %1 " " %2) %2) (if class {:class class}))))

(defn render-seq [sheet x y content options]
  (case (get options :direction :vertical)
    :vertical   (render-vertical sheet x y [:dummy {} content])
    :horizontal (render-horizontal sheet x y [:dummy {} content])
    (reduce #(identity [(max (nth %1 0) (nth %2 0))
                        (max (nth %1 1) (nth %2 1))
                        (conj (nth %1 2) (nth %2 2))])
      [0 0 []]
      (for [cont content]
        (render sheet x y cont)))))

(defn render [sheet x y expr & options]
  (cond
   (vector? expr) (render-element sheet x y expr)
   (literal? expr) (render-literal sheet x y expr)
   (seq? expr) (render-seq sheet x y expr options)
   (nil? expr) [1 1 ""]
   :else (render-literal sheet x y (str expr))))

