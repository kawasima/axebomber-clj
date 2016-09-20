(defproject net.unit8.axebomber/axebomber-clj "0.1.1"
  :url "https://github.com/kawasima/axebomber-clj"
  :description "The generator for MS-Excel grid sheet"
  :license {:name "Eclipse Public License - v 1.0"
            :url "http://www.eclipse.org/legal/epl-v10.html"
            :distribution :repo
            :comments "same as Clojure"}
  :dependencies [[org.clojure/clojure "1.8.0"]
                 [org.apache.poi/poi "3.14"]
                 [org.apache.poi/poi-ooxml "3.14"]
                 [clj-time "0.12.0"]
                 [hiccup "1.0.5"]]
  :profiles {:dev
              {:dependencies [[midje "1.8.3"]]
               :plugins [[lein-midje "3.2.1"]]}})
