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

(defn to-grid [sheet]
  (let [font (.getFontAt (.getWorkbook sheet) (short 0))]
    (doseq [col-index (range 256)]
      (.setColumnWidth sheet col-index (+ 200 (* 256 2)))))
  sheet)

(defn open-workbook [filename]
  (WorkbookFactory/create filename))

