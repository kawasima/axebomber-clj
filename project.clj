(defproject axebomber-clj "0.1.0-SNAPSHOT"
  :description "The generator for MS-Excel grid sheet"
  :license {:name "Eclipse Public License - v 1.0"
            :url "http://www.eclipse.org/legal/epl-v10.html"
            :distribution :repo
            :comments "same as Clojure"}
  :dependencies [[org.clojure/clojure "1.6.0"]
                 [org.apache.poi/poi "3.10-FINAL"]
                 [org.apache.poi/poi-ooxml "3.10-FINAL"]
                 [clj-time "0.6.0"]
                 [hiccup "1.0.5"]
;;                 [net.sf.jacob-project/jacob "1.17"]
;;                 [org.apache.pdfbox/pdfbox "1.8.4"]
;;                 [org.apache.xmlgraphics/batik-svggen "1.7"]
;;                 [org.apache.xmlgraphics/batik-dom "1.7"]
	]
  :jvm-opts ["-Djava.library.path=native"]
  :profiles {:dev
              {:dependencies [[midje "1.6.3"]]
               :plugins [[lein-midje "3.1.1"]]}})

