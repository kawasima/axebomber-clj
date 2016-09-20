(ns axebomber.renderer.html
  (:require [hiccup.core :refer :all]
            [hiccup.page :refer :all]))

(declare render)

(defn render-element [sheet {x :x y :y} expr]
  (conj sheet
    (if (map? (second expr))
      (-> expr
        (assoc-in expr [1 :data-x] x)
        (assoc-in expr [1 :data-y] y))
      (-> (drop 2 expr)
        (conj {:x x :y y} (first expr))
        vec))))

(defn render [sheet ctx expr & {:as options}]
  (cond
   (vector? expr) (render-element sheet ctx expr)
;   (literal? expr) (render-literal sheet x y expr)
;   (seq? expr) (render-seq sheet x y expr options)
;   (nil? expr) [1 1 ""]
;   :else (render-literal sheet x y (str expr))
    ))
