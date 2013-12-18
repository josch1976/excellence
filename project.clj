(defproject excellence "1.0.4"
  :description "Funktionen für die Datenextraktion aus Excel-Dateien"
  :url "http://github.com/josch1976/excellence"
  :license {:name "MIT License"
            :url "http:///en.wikipedia.org/wiki/MIT_License"}
  :dependencies [[org.clojure/clojure "1.5.1"]
		 [org.apache.poi/poi "3.9"]
		 [org.apache.poi/poi-ooxml "3.9"]]
  :plugins [[lein-difftest "2.0.0"]])
