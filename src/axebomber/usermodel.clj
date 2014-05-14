(ns axebomber.usermodel
  (:import [org.apache.poi.xssf.usermodel XSSFWorkbook]
           [org.apache.poi.ss.usermodel WorkbookFactory]))

(defn create-workbook [& {:keys [font-family font-size]
                          :or {font-family "Meiryo" font-size 8}}]
  (let [wb (XSSFWorkbook.)]
    (doto (.getFontAt wb (short 0))
      (.setFontName font-family)
      (.setFontHeightInPoints font-size))
    wb))

;; Font Offset in MS Pゴシック
;;
;;  8 = 200
;;  9 = 200
;; 10 = 200
;; 11 = 160
;; 12 = 160
;; 13 = 200
;; 14 = 180
;; 15 = 180
;; 16 = 160
;; 17 = 160
;; 18 = 160
;; 19 = 180
;; 20 = 160
;; 22 = 160
;; 24 = 140
;; 26 = 160
;; 28 = 150
;; 36 = 140

(defn to-grid [sheet]
  (let [font (.getFontAt (.getWorkbook sheet) (short 0))]
    (doseq [col-index (range 256)]
      (.setColumnWidth sheet col-index (+ 200 (* 256 2)))))
  sheet)

(defn open-workbook [filename]
  (WorkbookFactory/create filename))

