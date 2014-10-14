(ns axebomber.renderer.excel-picture
  (:use [axebomber util])
  (:require [clojure.java.io :as io])
  (:import [org.apache.poi.ss.usermodel Workbook Picture Drawing ClientAnchor]
           [javax.imageio ImageIO]
           [java.awt.image BufferedImage]
           [java.io ByteArrayOutputStream]))

(defn- read-image [file baos]
  (let [img (ImageIO/read file)]
    (ImageIO/write img "png" baos)
    img))

(defn- scan-columns [img sheet x]
  (let [image-width (.getWidth img)]
    (loop [cx x, width 0]
      (let [width (+ width (/ (.getColumnWidth sheet cx) 32))]
        (if (> width image-width)
          [(- cx x -1) width]
          (recur (inc cx) width))))))

(defn- scan-rows [img sheet y w-px]
  (let [image-height (* (.getHeight img) (/ w-px (.getWidth img)))]
    (loop [cy y, height 0]
      (let [height (+ height (* (.. sheet (createRow y) getHeightInPoints) (/ 4 3)))]
        (if (> height image-height)
          [(- cy y -1) height]
          (recur (inc cy) height))))))

(defn draw-image [sheet x y image-file & {data-width :data-width}]
  (let [baos (ByteArrayOutputStream.)
        img (read-image (io/file image-file) baos)
        drawing (.createDrawingPatriarch sheet)
        [w w-px] (if data-width
                   [data-width (->> (range x (+ x data-width 1))
                                    (map #(double (/ (.getColumnWidth sheet %) 32)))
                                    (reduce +))]
                   (scan-columns img sheet x))
        [h h-px] (scan-rows img sheet y w-px)
        anchor (.createAnchor drawing 0 0 0 0 x y (+ x w) (+ y h))
        pic-index (.addPicture (.getWorkbook sheet) (.toByteArray baos) Workbook/PICTURE_TYPE_PNG)]
    (.setAnchorType anchor 0)
    (.createPicture drawing anchor pic-index)
    [w h nil]))
