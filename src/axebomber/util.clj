(ns axebomber.util
  (:import [org.apache.poi.ss.util CellUtil]))

(def ^:dynamic *base-url* nil)

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

(defn get-cell [sheet x y]
  (CellUtil/getCell (CellUtil/getRow y sheet) x))

