(ns axebomber.usermodel
  (:import [org.apache.poi.xssf.usermodel XSSFWorkbook]))

(defn create-workbook [& {:keys [font-family font-size]
                          :or {font-family "Meiryo" font-size 8}}]
  (let [wb (XSSFWorkbook.)]
    (doto (.getFontAt wb (short 0))
      (.setFontName font-family)
      (.setFontHeightInPoints font-size))
    wb))

(defn to-grid [sheet]
  (doseq [col-index (range 256)]
    (.setColumnWidth sheet col-index 768))
  sheet)


