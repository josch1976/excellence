 ;; Copyright (c) <2013> <Joachim Scheffold>

;; Permission is hereby granted, free of charge, to any person obtaining a copy of
;; this software and associated documentation files (the "Software"), to deal in the
;; Software without restriction, including without limitation the rights to use,
;; copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
;; Software, and to permit persons to whom the Software is furnished to do so,
;; subject to the following conditions:

;; The above copyright notice and this permission notice shall be included in all
;; copies or substantial portions of the Software.

;; THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
;; INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
;; PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
;; HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
;; OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
;; SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


;; Funktionen zur Steuerung von Excel

(ns excellence.spreadsheet
  (:require [clojure.java.io :as io]
            [clojure.string :as str])
  (:import
   (java.io File FileInputStream FileOutputStream)
   (java.text DecimalFormat SimpleDateFormat)
   (java.util Calendar Date GregorianCalendar Locale)
   (org.apache.poi.xssf.usermodel XSSFWorkbook)
   (org.apache.poi.hssf.usermodel HSSFWorkbook)
   (org.apache.poi.ss.usermodel Cell CellStyle DataFormatter DateUtil Font
                                FormulaEvaluator IndexedColors RichTextString
                                Row Sheet Workbook WorkbookFactory)
   (org.apache.poi.ss.util CellReference)))


(declare
 ;; allgemeine Funktionen fuer Arbeitsmappen
 create-named-range! get-formula-evaluator create-workbook! load-workbook
 save-workbook!
 ;; Funktionen fuer Arbeitsblaetter
 add-sheet! delete-sheet! get-sheet get-sheet-name sheet-seq set-sheet-name!
 update-formulas!
 ;; Funktionen fuer Excel-Zeilen (rows)
 add-row-after-last! delete-all-rows! delete-row! get-create-next-row!
 get-create-row! get-last-row-num get-row insert-row-after! insert-row-before! row-seq
;; Funktionen fuer Excel-Zellen (cells)
 add-db-values-seq! add-value-rows! add-values! all->varchar apply-date-format!
 cell-iterator cell-reference cell-seq column-types create-cell!
 data-row-seq db-values-seq
 get-cell-formula-value get-cell-value get-create-next-cell! get-create-cell!
 get-last-column-num get-transformed-cell-value
 indexed-cell-map indexed-value-map insert-db-values! into-seq
 last-day-in-month
 set-cell-value!
 value-columnindex-map varchar->date varchar->double
 write-header!)

;;; ---------------------------------------------------------------------------
;;; Funktionen fuer Arbeitsmappen

(defn get-formula-evaluator
  "Erzeugt den Formel-Auswerter fuer die Arbeitsmappe.
Parameter:
- workbook: Referenz auf eine 'workbook'-Referenz
Beispiel:
(get-formula-evaluator wb)"
  [#^Workbook workbook]
  (.. workbook (getCreationHelper) (createFormulaEvaluator)))


(defn workbook
  "Erzeugt oder oeffne eine neue Arbeitsmappe.
Parameter:
- path:  Pfad zu einer Datei (optional)
Beispiel:
(workbook)
(workbook '/User/Documents/xy/test.xls') "
  ([] (XSSFWorkbook.))
  ([path] (WorkbookFactory/create (io/input-stream path))))


(defn create-workbook!
  "Erzeugt eine leere Arbeitsmmappe.
Parameter:
- sheet-name: Name (String) des zu erzeugenden Tabellenblattes (optional)
(create-workbook!)
(create-workbook! 'Blatt1') "
  ([]
   (create-workbook! "1"))
  ([sheet-name]
   (doto
     (XSSFWorkbook.)
     (add-sheet! sheet-name))))


(defn load-workbook
  "Laedt eine Arbeitsmappe mit dem Namen 'file-name'."
  [file-name]
  (with-open [stream (FileInputStream. file-name)]
    (WorkbookFactory/create stream)))


(defn save-workbook!
  "speichert eine Arbeitsmappe unter dem Namen 'save-name'."
  [^Workbook workbook file-name]
  (with-open [file-out (FileOutputStream. file-name)]
    (.write workbook file-out)))

(defn create-named-range!
  "erzeuge benannten Bereich"
  [^Sheet sheet ref-string ref-name]
  (doto (.createName (.getWorkbook sheet))
    (.setNameName ref-name)
    (.setRefersToFormula (str (get-sheet-name sheet) "!" ref-string))))

(defn alter-named-range!
  "aktualisiere den Zellbereich für einen benannten Bereich"
  [^Workbook workbook ref-name new-ref-string]
  (let [named (.getName workbook ref-name)]
    (.setRefersToFormula named
                         (str (.getSheetName named)  "!" new-ref-string))))

(defn has-named-range?
  [^Workbook workbook ref-name]
  (> (.getNameIndex workbook ref-name) -1))

(defn remove-named-range!
  [^Workbook workbook ref-name]
  (let [i (.getNameIndex workbook ref-name)]
    (when (> i 0)
      (.removeName workbook i))))


;;; ---------------------------------------------------------------------------
;;; Funktionen fuer Arbeitsblaetter

(defn add-sheet!
  "Fuegt ein leeres Blatt mit dem Namen 'sheet-name' hinzu."
  [^Workbook workbook sheet-name]
  (.createSheet workbook sheet-name))


(defmulti delete-sheet!
  "Loescht ein Arbeitsblatt einer Arbeitsmappe per Index oder Name."
  (fn [^Workbook workbook index-or-name]
    (if (integer? index-or-name)
      :indexed
      :named)))
(defmethod delete-sheet! :indexed [^Workbook workbook index-or-name]
  (doto workbook
    (.removeSheetAt index-or-name)))
(defmethod delete-sheet! :named [^Workbook workbook index-or-name]
  (let [index (.getSheetIndex workbook index-or-name)]
    (delete-sheet! workbook index)))


(defmulti get-sheet
  "Zugriff auf das Arbeitsblatt einer Arbeitsmappe per Index oder Name."
  (fn [^Workbook workbook index-or-name]
    (if (integer? index-or-name)
      :indexed
      :named)))
(defmethod get-sheet :indexed [^Workbook workbook index-or-name]
  (. workbook getSheetAt index-or-name))
(defmethod get-sheet :named [^Workbook workbook index-or-name]
  (. workbook getSheet (str index-or-name)))


(defn get-sheet-name
  "Liefert den Namen eines Arbeitsblattes."
  [#^Sheet sheet]
  (.getSheetName sheet))


 (defn set-sheet-name!
  "Benennt ein Arbeitsblatt um."
  [#^Sheet sheet sheet-name]
  (.setSheetName sheet sheet-name)
  sheet)


(defn sheet-seq
  "Liefert eine lazy seq aller Arbeitsblaetter einer Arbeitsmappe."
  [#^Workbook workbook]
  (for [idx (range (.getNumberOfSheets workbook))]
    (.getSheetAt workbook idx)))


(defn update-formulas!
  "aktualisiert den Zellinhalt einer Zelle mit einer Formel"
  [^Sheet sheet]
  (let [ev (get-formula-evaluator (.getWorkbook sheet))]
    (doseq [c (cell-seq sheet)]
      (when (= (.getCellType c) Cell/CELL_TYPE_FORMULA)
        (.evaluateFormulaCell ev c)))))

;;; ---------------------------------------------------------------------------
;;; Funktion fuer Excel-Zeilen (rows)
;;; Vor dem Zugriff auf einzelne Zellen muss die entsprechende Zeile (row)
;;; vorhanden sein.

(def col-index-mapping
  (let [c-names (concat (map (fn [ch] (keyword (str (char ch))))
                             (range 65 91))
                        (for [x (range 65 91) y (range 65 91)]
                          (keyword (str (char x) (char y)))))]
    (zipmap c-names
            (range (count c-names)))))

(defn insert-row-after!
  "Fuegt eine neue Zeile unterhalb der angegebenen ein"
  ([^Sheet sheet row-index]
     (if (= (.getLastRowNum sheet) row-index)
       (add-row-after-last! sheet)
       (insert-row-before! sheet (inc row-index))))
  ([^Row row]
     (insert-row-after! (.getSheet row) (.getRowNum row))))

(defn add-row-after-last!
  "Fuegt eine neue Zeile nach der letzten (vorhandenen) Zeile an."
  [^Sheet sheet]
  (let [row-num (if (= 0 (.getPhysicalNumberOfRows sheet))
                  0
                  (inc (.getLastRowNum sheet)))]
    (.createRow sheet row-num)))


(defn insert-row-before!
  "fuegt vor der Zeile eine neue leere Zeile ein"
  ([^Sheet sheet row-index]
     (.shiftRows sheet row-index (.getLastRowNum sheet) 1 true false)
     (.createRow sheet row-index))
  ([^Row row]
     (insert-row-before! (.getSheet row) (.getRowNum row))))


(defn get-create-row!
  "liefert ein Row-Objekt - wenn diese Zeile noch nicht existiert, wird sie angelegt"
  [^Sheet sheet row-index]
  (let [row (get-row sheet (int row-index))]
    (if (nil? row)
      (.createRow sheet (int row-index))
      row)))

(defn get-create-next-row!
 "Liefert die auf aktuelle Zeile folgende Zeile. Wenn diese noch nicht
existiert, wird diese neu erzeugt (z.B. wenn die aktuelle Zeile
die letzte des Blattes ist."
 [^Row row]
 (get-create-row! (.getSheet row)
                  (+ (.getRowNum row) 1)))

(defn delete-row!
  "Loescht eine Tabellenzeile eines Blattes"
  ([^Sheet sheet row-index]
     (let [last-row-num (if (= 0 (.getPhysicalNumberOfRows sheet))
                          0
                          (inc (.getLastRowNum sheet)))]
       (cond (and (>= row-index 0) (< row-index last-row-num))
             (.shiftRows sheet (inc row-index) last-row-num -1)
             (= row-index last-row-num)
             (.removeRow sheet (.getRow sheet row-index)))
       sheet))
  ([^Row row]
     (delete-row! (.getSheet row) (.getRowNum row))))

(defn delete-rows!
  "Loescht mehrere Zeilen eines Blattes"
  [^Sheet sheet row-indexes]
  (doseq [i (sort > row-indexes)]
    (delete-row! sheet i)))

(defn delete-all-rows!
  "Loescht alle Zeilen des Arbeitsblattes."
  [^Sheet sheet]
  (doall
   (for [row (doall (row-seq sheet))]
     (delete-row! row)))
  sheet)

(defn row-seq
  "Liefert eine lazy sequence aller Zeilen einer Arbeitsmappe."
  [#^Sheet sheet]
  (iterator-seq (.iterator sheet)))


(defn get-row
  "Gibt die Zeile eines Blattes ueber den Index zurueck. Falls diese nicht
existiert wird eine Exception geworfen"
  [^Sheet sheet row-index]
  (try (.getRow sheet row-index)
       (catch Exception e (prn "der Rowindex " row-index " existiert nicht"))))


(defn get-last-row-num
  "Ermittelt die letzte Zeile eines Arbeitsblattes 'sheet'"
  [^Sheet sheet]
  (apply max (map (fn [r] (.getRowNum r)) (row-seq  sheet))))


;;; ---------------------------------------------------------------------------
;;; Funktion für Excel-Zellen (cells)

(defn cell-reference [cell]
  "Liefert den als String formatierten Zellbezug des null-basierten Koordinaten-
   system "
  (.formatAsString (CellReference. (.getRowIndex cell) (.getColumnIndex cell))))


(defmulti cell-seq
  "Liefert eine Sequenz von Tabellenzellen abhaengig vom Argument, das ein
   Tabellenblatt, eine Zeile oder eine Collection sein kann.
   Die Sequenz ist sortiert nach Blatt, Zeile und Spalte."
  (fn [x]
    (cond
      (isa? (class x) Sheet) :sheet
      (isa? (class x) Row)   :row
      (seq? x)               :coll
      ;:else                  :default
      )))
(defmethod cell-seq :row  [row]
  (iterator-seq (.iterator row)))
(defmethod cell-seq :sheet [sheet]
  (for [row  (row-seq sheet)
        cell (cell-seq row)]
    cell))
(defmethod cell-seq :coll [coll]
  (for [x coll,
        cell (cell-seq x)]
    cell))

(defn extract-key
  [^String s]
  (let [b (.indexOf s "<-")
        e (.indexOf s "->")]
    (if (and (> b -1) (> e 2))
      (.toUpperCase (.substring s (+ b 2) e))
      "FEHLER")))

(extract-key "<-sds->")

(defmulti indexed-*-map
  (fn [x index-type fun]
    index-type))
(defmethod indexed-*-map :address-key  [x index-type fun]
  (let [cs (cell-seq x)]
    (zipmap (map (fn [c] (keyword (cell-reference c))) cs)
            (map (fn [c] (fun c)) cs))))
(defmethod indexed-*-map :address-nested-map [x index-type fun]
  (let [cs (cell-seq x)]
    (apply merge-with merge
           (map (fn [c]
                  (let [cf (cell-reference c)
                        rw  (re-find #"\d+" cf)
                        cl  (re-find #"\D+" cf)]
                    {(keyword rw)
                     {(keyword cl)
                      (fun c)}}))
                cs))))
(defmethod indexed-*-map :row-col-vec  [x index-type fun]
  (let [cs (cell-seq x)]
    (zipmap (map (fn [c] [(.getRowIndex c) (.getColumnIndex c)]) cs)
            (map (fn [c] (fun c)) cs))))
(defmethod indexed-*-map :comment-key [x index-type fun]
  (let [cs (cell-seq x)]
    (zipmap (map (fn [c] (if-let [cmt (.getCellComment c)]
                           (extract-key (.getString (.getString cmt)))
                           ("empty")))
                   cs)
            (map (fn [c] (fun c)) cs))))
(defmethod indexed-*-map :row-col-nested-map [x index-type fun]
  (let [cs (cell-seq x)]
    (apply merge-with merge
           (map (fn [c] {(.getRowIndex c)
                         {(.getColumnIndex c)

                          (fun c)}})
                cs))))
(defmethod indexed-*-map :coordinates-vec  [x index-type fun]
  (let [cs (cell-seq x)
        fr (first (row-seq x))]
    (zipmap
     (map (fn [c] [(get-cell-value (get-create-cell! (.getRow c) 0))
                   (get-cell-value (get-create-cell! fr (.getColumnIndex c)))])
          cs)
     (map (fn [c] (fun c)) cs))))
(defmethod indexed-*-map :coordinates-nested-map  [x index-type fun]
  (let [cs (cell-seq x)
        fr (first (row-seq x))]
    (apply merge-with merge
           (map (fn [c] {(keyword (str (get-cell-value
                                        (get-create-cell! (.getRow c) 0))))
                         {(keyword (str (get-cell-value (get-create-cell!
                                                         fr
                                                         (.getColumnIndex c)))))
                          (fun c)}})
                cs))))

(defn indexed-cell-map
  "Speichert Zellreferenzen eines Zellbereiches in einen
assoziativen Speicher (map):
   'index-type' steuert den Aufbau der Datenstruktur:
   :address-key            - {:A1 foo :A2 bar}
   :address-nested-map     - {:1 {:A foo} :2 {:A bar}}
   :row-col-vec            - {[0 0] foo [0 1] bar}
   :row-col-nested-map     - {0 {0 foo} 0 {1 :bar}}
   :coordinates-vec        - {[zeile1 spalte1] foo [zeile1 spalte1] bar}
   :coordinates-nested-map - {:zeile1 {:spalte1 foo} :zeile1 {:spalte2 bar}}"
  [x index-type]
  (indexed-*-map x index-type identity))


(defn indexed-value-map
  "Laedt den Inhalt eines Zellbereiches in einen assoziativen Speicher (map):
'index-type' steuert den Aufbau der Datenstruktur:
:address-key            - {:A1 foo :A2 bar}
:address-nested-map     - {:1 {:A foo} :2 {:A bar}}
:row-col-vec            - {[0 0] foo [0 1] bar}
:row-col-nested-map     - {0 {0 foo} 0 {1 :bar}}
:coordinates-vec        - {[zeile1 spalte1] foo [zeile1 spalte1] bar}
:coordinates-nested-map - {:zeile1 {:spalte1 foo} :zeile1 {:spalte2 bar}}"
  [x index-type]
  (indexed-*-map x index-type get-cell-value))


(defn get-last-column-num
  "Ermittelt die letzte Spalte einer Reihe oder einem Tabellenblatt 'x'"
  [x]
  (apply max (map (fn [c] (.getRowIndex c)) (cell-seq x))))


(defn into-seq
  "Liefert alle Zellen eines Blattes oder einer Zeile als 'sequence'"
  [sheet-or-row]
  (vec (for [item (iterator-seq (.iterator sheet-or-row))] item)))


(defn add-values! [#^Sheet sheet values]
  "Fuegt eine Reihe von Werten (als Auflistung) nacheinander in die Zellen ein.
   Bsp: (add-values sheet [1 3 5 ])"
  (let [row (add-row-after-last! sheet)]
    (doseq [[column-index value] (partition 2 (interleave (iterate inc 0) values))]
      (set-cell-value! (.createCell row column-index) value))
    row))


(defn add-value-rows! [#^Sheet sheet rows]
  "Fuegt einem Arbeitsblatt mehrere Zeilen (Sequenz in Sequence verschachtelt)
   mit Daten hinzu.
   Bsp: (add-values sheet [[1 3 5] [3 nil 4 t]])"
  (doseq [values rows]
    (add-values! sheet values)))


(defn add-db-values-seq!
  "Baut auf der Funktion 'insert-db-values' auf. Nimmt als ersten Paramter eine
   Sequenz, die aus 'maps' mit Spaltennamen und Werten besteht, entgegen.
   z.B.: (insert-db-values! ({:name tom :alter 24 :beruf schreiner}
                             {:name heinz :alter 22 :beruf schlosser})
                            {:name 0 :alter 2 :beruf 4})"
  [sheet data columnnames]
  (doseq [row-data data]
    (insert-db-values!
      (add-row-after-last! sheet)
      row-data
      columnnames)))


(defn insert-db-values!
  "nimmt als ersten Parameter eine 'map' mit den Spaltennamen und Werten
   und als zweiten Parameter eine 'map' mit den Spaltennummern (nullbasiert)
   und den Spaltennamen und schreibt die Daten der ersten Map (vals) in
   eine Zeile eines Tabellenblattes
   z.B.: (insert-db-values! {:name tom :alter 24 :beruf schreiner}
                            {:name 0 :alter 2 :beruf 4})"
  [row row-data columnnames]
  (let [nc (zipmap (vals columnnames) (keys columnnames))]
    (doseq [cell-data row-data]
      (set-cell-value!
       (create-cell! row (get nc (key cell-data)))
        (val cell-data)))))


(defn cell-iterator [^Row row]
  "Liefert einen Iterator fuer die Zellen einer Reihe."
  (for [idx (range (.getFirstCellNum row) (.getLastCellNum row))]
    (if-let [cell (.getCell row idx)]
      cell
      (.createCell row idx Cell/CELL_TYPE_BLANK))))


(defmulti get-cell-value
  "Liest den Inhalt einer Tabellenzelle abhaengig vom Zellentyp"
  (fn [cell]
    (when-let [ct (. cell getCellType)]
      (if (not (= Cell/CELL_TYPE_NUMERIC ct))
        ct
        (if (DateUtil/isCellDateFormatted cell)
          :date
          ct)))))
(defmethod get-cell-value Cell/CELL_TYPE_BLANK   [cell]
  nil)
(defmethod get-cell-value Cell/CELL_TYPE_FORMULA [cell]
  (let [val            (. (get-formula-evaluator
                             (.. cell getSheet getWorkbook))
                            evaluate cell)
        evaluated-type (. val getCellType)]
    (get-cell-formula-value val (if (= Cell/CELL_TYPE_NUMERIC evaluated-type)
                                  (if (DateUtil/isCellInternalDateFormatted cell)
                                    :date
                                    :number)
                                  evaluated-type))))
(defmethod get-cell-value Cell/CELL_TYPE_BOOLEAN [cell]
  (. cell getBooleanCellValue))
(defmethod get-cell-value Cell/CELL_TYPE_STRING  [cell]
  (. cell getStringCellValue))
(defmethod get-cell-value Cell/CELL_TYPE_NUMERIC [cell]
  (. cell getNumericCellValue))
(defmethod get-cell-value :date                  [cell]
  (. cell getDateCellValue))
(defmethod get-cell-value :default [cell]
  (str "Unknown cell type " (. cell getCellType)))


(defn- get-index-type [index]
    (cond
     (string? index)
     :comment-key
     (keyword? index)
      :address-key
      (and (= (type index) clojure.lang.PersistentVector)
           (number? (first index))
           (number? (second index)))
      :row-col-vec
      (and (= (type index) clojure.lang.PersistentVector)
           (keyword? (first index))
           (keyword? (second index)))
      :coordinates-vec))


(defmulti get-*-by-index
  (fn [sheet index fun] (get-index-type index)))
(defmethod get-*-by-index :address-key [sheet index fun]
  (when-let [res (first (filter (fn [c] (= index (keyword (cell-reference c))))
                                (cell-seq sheet)))]
      (fun res)))
(defmethod get-*-by-index :row-col-vec [sheet index fun]
  (when-let [res (first (filter (fn [c] (and (= (first index) (.getRowIndex c))
                                             (= (second index) (.getColumnIndex c))))
                                (cell-seq sheet)))]
      (fun res)))
(defmethod get-*-by-index :comment-key [sheet index fun]
  (when-let [res (first (filter (fn [c] (= index
                                           (extract-key (.getString (.getString (.getCellComment c))))))
                                (cell-seq sheet)))]
      (fun res)))
(defmethod get-*-by-index :coordinates-vec [sheet index fun]
  (let [fr (get-row sheet 0)]
    (when-let [res (first
                    (filter
                     (fn [c] (and (= (name (first index))
                                     (get-cell-value (get-create-cell! (.getRow c) 0)))
                                  (= (name (second index))
                                     (get-cell-value
                                      (get-create-cell! fr  (.getColumnIndex c))))))
                     (cell-seq sheet)))]
      (fun res))))

(defn get-value-by-index
 "Liest den Inhalt einer Tabellenzelle wenn die Adresse der Zelle
  über einen Index angegeben wird"
 [sheet index]
 (get-*-by-index sheet index get-cell-value))

(defn get-cell-by-index
 "Liefert eine Referenz (pointer) auf eine Tabellenzelle wenn die Adresse der Zelle
  über einen Index angegeben wird"
 [sheet index]
 (get-*-by-index sheet index identity))

(defmulti get-cell-formula-value
  "Liefert den Wert der Tabellenzelle nach der Auswertung der hinterlegten
   Formel, siehe http://poi.apache.org/spreadsheet/eval.html#Evaluate"
  (fn [evaled-cell evaled-type]
    evaled-type))
(defmethod get-cell-formula-value Cell/CELL_TYPE_BOOLEAN [evaled-cell evaled-type]
  (. evaled-cell getBooleanValue))
(defmethod get-cell-formula-value Cell/CELL_TYPE_STRING  [evaled-cell evaled-type]
  (. evaled-cell getStringValue))
(defmethod get-cell-formula-value :number  [evaled-cell evaled-type]
  (. evaled-cell getNumberValue))
(defmethod get-cell-formula-value :date    [evaled-cell evaled-type]
  (DateUtil/getJavaDate (. evaled-cell getNumberValue)))
(defmethod get-cell-formula-value :default [evaled-cell evaled-type]
  (str "Unknown cell type " (. evaled-cell getCellType)))


(defn get-transformed-cell-value
  "Liefert den Wert der Tabellenzelle nach der Umwandlung in den
   vorgesehenen Spaltentyp: Argument als String oder Funktion"
  [cell  transformation]
  (if (= String (type transformation)) ; Prüft ob Funktion oder String
    ((get column-types transformation) cell)
    (transformation cell)))


(defmulti set-cell-value!
  "Schreibt den als Argument (String, Date oder Double) uebergebenen
   Wert in die Tabellenzelle"
  (fn [c v] (type v)))
(defmethod set-cell-value! java.lang.String [c v]
  (.setCellValue c v))
(defmethod set-cell-value! java.util.Date [c v]
  (doto c
    (.setCellValue v)
    (apply-date-format! "m/d/yy")))
(defmethod set-cell-value! java.lang.Double [c v]
  (.setCellValue c v))
(defmethod set-cell-value! java.lang.Integer [c v]
  (.setCellValue c (double v)))
(defmethod set-cell-value! java.lang.Long [c v]
  (.setCellValue c (double v)))
(defmethod set-cell-value! nil [c v])


(defmulti create-cell!
  "Erzeugt eine neue Zelle in einer vorhandenen Tabellenzeile. Die Bezeichnung
   der Spalte kann als String (z.b. 'A'), Keyword (z.B. ':A') oder Integer (z.B. '0')
   erfolgen"
  (fn [r c]
    (type c)))
(defmethod create-cell! java.lang.String         [r c]
  (create-cell! r (CellReference/convertColStringToIndex c)))
(defmethod create-cell! clojure.lang.Keyword     [r c]
  (create-cell! r (name c)))
(defmethod create-cell! java.lang.Integer        [r c]
  (get-create-cell! r c))
(defmethod create-cell! java.lang.Long           [r c]
  (get-create-cell! r (int c)))


(defn get-cell
  [r i]
  (try (.getCell r i)
       (catch Exception e (prn "Der Cell-Index " i " existiert nicht"))))


(defn get-create-cell!
 "Liefert eine Referenz auf ein 'Cell'-Objekt, das durch die Angaben
  Reihe 'r' und Zellennummer 'i' bestimmt wird. Wenn die Zelle nicht besteht
  wird eine neue angelegt."
  [^Row r
   i]
  (let [cell (.getCell r (int i))]
    (if (nil? cell)
      (.createCell r (int i))
      cell)))

(defn get-create-next-cell!
 "Liefert die auf aktuelle Zelle folgende Zelle. Wenn diese noch nicht
existiert, wird diese neu erzeugt (z.B. wenn die aktuelle Zelle
die letzte der Zeile ist."
 [^Cell cell]
 (get-create-cell! (.getRow cell)
                   (+ (.getColumnIndex cell) 1)))

;;; ---------------------------------------------------------------------------
;;; Umwandlungsfunktionen

(defn varchar->double
  "Wandelt einen String in einen Double-Wert um. Falls die Umwandlung nicht gelingt,
wir nil zurueckgegeben."
  [c]
  (let [v (get-cell-value c)]
    (condp = (type v)
      java.lang.String   (if (= (.trim v) "")
                           nil
                           (.doubleValue (.parse (DecimalFormat. "#,###.00") v)))
      java.lang.Double   v
      nil)))


(defn varchar->date
  "Wandelt einen String in einen Date-Wert um. Falls die Umwandlung nicht gelingt,
wird nil zurueckgegeben."
  [c]
  (let [v (get-cell-value c)]
    (condp = (type v)
      java.lang.String   (if (= (.trim v) "")
                           nil
                           (.parse (SimpleDateFormat. "d.M.yy") v))
      java.util.Date     v
      nil)))

(defn all->varchar
"Wandelt jeden Datentyp in einer Zelle grundsätzlich in einen String-Wert um."
  [c]
  ;; (condp = (type v)
  ;;   java.lang.String
  ;;   java.lang.Double   (.format (DecimalFormat. "#,###.00") v)
  ;;   java.util.Date     (.format (SimpleDateFormat. "dd.MM.yyyy") v)
  ;;   nil)
  ;;
  (.formatCellValue (DataFormatter.) c))



(def column-types
  {"varchar"  all->varchar
   "double"   varchar->double
   "date"     varchar->date  })


(defn columnindex-value-map
  "Liest eine Tabellenzeile und erzeugt eine Map mit der Spaltennummer und
dem Zellenwert.
   Zeilennummern beginnen bei null.
   z. B. {0 2 , 1 Wert , 2 2 , 3 Hello}"
  ([#^Sheet sheet row-num]
    (let [rows (get-row sheet row-num)]
      (into
        {}
        (map (fn [c]
               {(.getColumnIndex c)
                (get-transformed-cell-value c all->varchar)})
             rows))))
  ([#^Sheet sheet row-num first-col]
    (filter (fn [t] (>= (key t) first-col))
            (columnindex-value-map sheet row-num)))
  ([#^Sheet sheet row-num first-col last-col]
    (filter (fn [t] (and (>= (key t) first-col)
                         (<= (key t) last-col)))
            (columnindex-value-map sheet row-num))))


(defn db-values-seq
  "Gibt eine Sequenz mit 'maps' für jede Datenzeilen zurueck.
Parameter:
- sheet:        Referenz auf das Tabellenblatt
- columnnames:  Map mit den Spaltennummern und einer Spaltenbezeichnung (als String)
                z. B. {0 spalte1 1 spalte 2 ...}
- db-types:     Map mit den Spaltennummern und dem Datentyp (als String, siehe 'column-types'
                in den der Zellenwert umgewandelt werden soll
                z. B. {0 varchar 1 double 2 date}
- begin-row:    in welcher Zeile begonnen wird
Bsp:
(db-values-seq sh {0 spalte1 1 spalte2} {0 varchar 1 varchar} 0)
({:spalte1 00, :spalte2 bb}
  {:spalte1 11, :spalte2 11bb}
  {:spalte1 22, :spalte2 22bb}
  {:spalte1 33, :spalte2 33bb})"
  ([sheet columnnames db-types begin-row]
     (db-values-seq sheet columnnames db-types begin-row (get-last-row-num sheet)))
  ([sheet columnnames db-types begin-row end-row]
     (let [rows (filter
                 (fn [r] (and (>= (.getRowNum r) begin-row)
                              (<= (.getRowNum r) end-row)))
                 (row-seq sheet))]
       (for [r rows]
         (into
          {}
          (map
           (fn [cnx]
             {(keyword (val cnx))
              (get-transformed-cell-value (get-create-cell! r (key cnx))
                                          (db-types (key cnx)))})
           columnnames))))))


(defn data-row-seq
  "Liefert alle Zeilen eines Blattes aus und liefert die Daten als Sequenz.
   Die Startzeile wird über den Parameter 'begin-row-num' angegeben."
  ([#^Sheet sheet begin-row-num]
     (columnindex-value-map sheet  (int begin-row-num)))
  ([#^Row begin-row]
     (columnindex-value-map (.getSheet begin-row) (.getRowNum begin-row))))





;;; ---------------------------------------------------------------------------
;;; Allgemeine Hilfsfunktionen

(defn write-header!
  "Schreibt die Spaltenueberschriften in das Ziel-Tabellenblatt.
   Zeilennummern beginnen bei null."
  ([cloumn-name-map #^Sheet dest-sheet dest-row-num]
    (write-header! cloumn-name-map (nth (row-seq dest-sheet) dest-row-num)))
  ([cloumn-name-map #^Row dest-row]
    (doseq [[c v] cloumn-name-map]
      (->
        dest-row
        (get-create-cell! (.getColumnIndex c))
        (.setCellValue v)))))


(defn apply-date-format! [cell format]
  "Format eine Tabellenzeile mit dem als Argument uebergebenen Datumsformat"
  (let [workbook (.. cell getSheet getWorkbook)
        date-style (.createCellStyle workbook)
        format-helper (.getCreationHelper workbook)]
    (.setDataFormat date-style
		    (.. format-helper createDataFormat (getFormat format)))
    (.setCellStyle cell date-style)))


(defn last-day-in-month [m y]
"Ermittlung des letzten Tages eines Monats."
  (->(doto
       (GregorianCalendar.)
       (.set Calendar/DATE 1)
       (.set Calendar/MONTH m)
       (.set Calendar/YEAR y)
       (.add Calendar/DATE -1))
    (.get Calendar/DATE)))


