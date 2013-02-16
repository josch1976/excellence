;; Funktionen zur Steuerung von Excel

(ns excellence.spreadsheet
  (:import 
   (java.io File FileInputStream FileOutputStream)
   (java.text DecimalFormat SimpleDateFormat)
   (java.util Calendar Date GregorianCalendar Locale)
   (org.apache.poi.hssf.usermodel HSSFWorkbook HSSFRichTextString)
   (org.apache.poi.ss.usermodel Workbook Sheet Cell DataFormatter
                                Row WorkbookFactory DateUtil
                                IndexedColors CellStyle Font FormulaEvaluator)
   (org.apache.poi.ss.util CellReference)))


(declare
  get-formula-evaluator 
  create-workbook
  load-workbook
  save-workbook!
  create-named-range!
  add-sheet!
  delete-sheet!
  get-sheet
  get-sheet-name
  set-sheet-name!
  sheet-seq
  update-formula!
  add-row-after-last!
  insert-row!
  get-create-row!
  delete-row!
  delete-all-rows!
  row-seq
  get-row
  get-last-row-num
  cell-reference
  cell-seq
  get-last-column-num
  into-seq
  add-values!
  add-value-rows!
  add-db-values-seq!
  insert-db-values!
  cell-iterator
  get-cell-value
  get-cell-formula-value
  get-transformed-cell-value
  set-cell-value!
  create-cell!
  get-create-cell!
  column-types
  varchar->double
  varchar->date
  all->varchar
  value-columnindex-map
  db-values-seq
  data-row-seq
  write-header!
  apply-date-format!
  last-day-in-month)

;;; ---------------------------------------------------------------------------
;;; Funktionen fuer Arbeitsmappen
 
(defn get-formula-evaluator 
  "Erzeugt den Formel-Auswerter fuer die Arbeitsmappe"
  [#^Workbook workbook]
  (.. workbook (getCreationHelper) (createFormulaEvaluator)))


(defn create-workbook!
  "erzeugt eine leere Arbeitsmmappe"
  ([]
   (create-workbook! "1"))
  ([sheet-name]
   (doto
     (HSSFWorkbook.)
     (add-sheet! sheet-name))))


(defn load-workbook 
  "Laedt eine Arbeitsmappe mit dem Namen 'file-name'."  
  [file-name]
  (with-open [stream (FileInputStream. file-name)]
    (WorkbookFactory/create stream)))


(defn save-workbook! 
  "speichert eine Arbeitsmappe unter dem Namen 'save-name'."
  [#^Workbook workbook file-name]
  (with-open [file-out (FileOutputStream. file-name)]
    (.write workbook file-out)))


(defn create-named-range!
  "erzeuge benannten Bereich"
  [workbook sheet-name ref-string ref-name]
  (doto (.createName workbook) 
    (.setNameName ref-name) 
    (.setReference (str sheet-name "!" ref-string)))) 


(defn update-named-range!
  "aktualisiere den Zellbereich für einen benannten Bereich"
  [workbook ref-name ref-string]
  (let [named (.getName workbook ref-name)]
        (.setReference 
          named
          (str (apply str (take-while (fn [c] (not= c \!)) (.getReference named)))
               "!" 
               ref-string))))


;;; ---------------------------------------------------------------------------
;;; Funktionen fuer Arbeitsblaetter

(defn add-sheet! 
  "Fuegt ein leeres Blatt mit dem Namen 'sheet-name' hinzu."
  [#^Workbook workbook sheet-name]
  (doto workbook
    (.createSheet sheet-name)))


(defmulti delete-sheet!
  "Loescht ein Arbeitsblatt einer Arbeitsmappe per Index oder Name."
  (fn [workbook index-or-name] 
    (if (integer? index-or-name) 
      :indexed 
      :named)))
(defmethod delete-sheet! :indexed [workbook index-or-name]
  (doto workbook
    (.removeSheetAt index-or-name)))
(defmethod delete-sheet! :named [workbook index-or-name]
  (let [index (.getSheetIndex workbook index-or-name)]
    (delete-sheet! workbook index)))


(defmulti get-sheet
  "Zugriff auf das Arbeitsblatt einer Arbeitsmappe per Index oder Name."
  (fn [workbook index-or-name] 
    (if (integer? index-or-name) 
      :indexed 
      :named)))
(defmethod get-sheet :indexed [workbook index-or-name]
  (. workbook getSheetAt index-or-name))
(defmethod get-sheet :named [workbook index-or-name]
  (. workbook getSheet (str index-or-name)))


(defn get-sheet-name
  "Liefert den Namen eines Arbeitsblattes."
  [#^Sheet sheet]
  (.getSheetName sheet))


(defn set-sheet-name!
  "Benennt ein Arbeitsblatt um."
  [#^Sheet sheet sheet-name]
  (.setSheetName sheet sheet-name))


(defn sheet-seq 
  "Liefert eine lazy seq aller Arbeitsblaetter einer Arbeitsmappe."
  [#^Workbook workbook]
  (for [idx (range (.getNumberOfSheets workbook))]
    (.getSheetAt workbook idx)))


(defn update-formula!
  "aktualisiert den Zellinhalt einer Zelle mit einer Formel"
  [sh]
  (let [ev (get-formula-evaluator (.getWorkbook sh))]
    (doseq [c (cell-seq sh)]
      (when (= (.getCellType c) Cell/CELL_TYPE_FORMULA)
        (.evaluateFormulaCell ev c)))))

;;; ---------------------------------------------------------------------------
;;; Funktion fuer Excel-Zeilen (rows)
;;; Vor dem Zugriff auf einzelne Zellen muss die entsprechende Zeile (row)
;;; vorhanden sein.

(defn add-row-after-last! 
  "Fuegt eine neue Zeile nach der letzten vorhandenen Zeile an."
 [^Sheet sheet]
  (let [row-num (if (= 0 (.getPhysicalNumberOfRows sheet)) 
                  0 
                  (inc (.getLastRowNum sheet)))]
    (.createRow sheet row-num)))


(defn insert-row-before! 
  "fuegt vor der Zeile eine neue leere Zeile ein"
  [sheet row-index]
  (.shiftRows sheet row-index (.getLastRowNum sheet) 1 true false)
  (.createRow sheet row-index))


(defn get-create-row!
  "liefert ein Row-Objekt - wenn diese Zeile noch nicht existiert, wird sie angelegt"
  [sheet row-index]
  (let [row (get-row sheet row-index)]
    (if (nil? row)
      (let [new-row (add-row-after-last! sheet)]
        (loop [new-row new-row]
            (if (= (.getRowNum new-row) row-index)
              new-row
              (recur (add-row-after-last! sheet)))))
        row)))


(defmulti delete-row!
  "Loescht eine Zeile des Arbeitsblattes."
  (fn [sheet integer-or-row] 
    (cond
      (isa? (class integer-or-row) Row) :row
      (isa? (class integer-or-row) Integer) :integer
      :else :default)))
(defmethod delete-row! :row [sheet integer-or-row]
 (delete-row! sheet (.getRowNum sheet integer-or-row)))
(defmethod delete-row! :integer [sheet integer-or-row]
  (let [last-row-num (if (= 0 (.getPhysicalNumberOfRows sheet)) 
                       0 
                       (inc (.getLastRowNum sheet)))]
    (cond (and (>= integer-or-row 0) (< integer-or-row last-row-num))
          (.shiftRows sheet (inc integer-or-row) last-row-num -1)
          (= integer-or-row last-row-num)
          (.removeRow sheet (.getRow sheet integer-or-row)))
    sheet))


(defn delete-all-rows!
  "Loescht alle Zeilen des Arbeitsblattes."
  [sheet]
  (doall
   (for [row (doall (row-seq sheet))]
     (delete-row! sheet row)))
  sheet)


(defn row-seq 
  "Liefert eine lazy sequence aller Zeilen einer Arbeitsmappe."
  [#^Sheet sheet]
  (iterator-seq (.iterator sheet)))


(defn get-row
  [#^Sheet sheet row-index]
  (try (.getRow sheet row-index)
       (catch Exception e (prn "Rowindex does't exist"))))


(defn get-last-row-num
  "Ermittelt die letzte Zeile eines Arbeitsblattes 'sheet'"
  [sheet]
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


(defmulti indexed-cell-map
  (fn [x index-type]
    index-type))
(defmethod indexed-cell-map :address-string  [x index-type]
  (let [cs (cell-seq x)]
    (zipmap (map (fn [c] (cell-reference c)) cs)
            cs)))
(defmethod indexed-cell-map :row-col-vec  [x index-type]
  (let [cs (cell-seq x)]
    (zipmap (map (fn [c] [(.getRowIndex c) (.getColumnIndex c)]) cs)
            cs)))
(defmethod indexed-cell-map :coord-sys-vec  [x index-type]
  (let [cs (cell-seq x)
        fr (first (row-seq x))]
    (zipmap
     (map (fn [c] [(get-cell-value (get-create-cell! (.getRow c) 0))
                   (get-cell-value (get-create-cell! fr (.getColumnIndex c)))])
          cs)
     cs)))

(defmulti indexed-value-map
  (fn [x index-type]
    index-type))
(defmethod indexed-value-map :address-string  [x index-type]
  (let [cs (cell-seq x)]
    (zipmap (map (fn [c] (cell-reference c)) cs)
            (map (fn [c] (get-cell-value c)) cs))))
(defmethod indexed-value-map :row-col-vec  [x index-type]
  (let [cs (cell-seq x)]
    (zipmap (map (fn [c] [(.getRowIndex c) (.getColumnIndex c)]) cs)
            (map (fn [c] (get-cell-value c)) cs))))
(defmethod indexed-value-map :coord-sys-vec  [x index-type]
  (let [cs (cell-seq x)
        fr (first (row-seq x))]
    (zipmap (map (fn [c] [(get-cell-value (get-create-cell! (.getRow c) 0))
                          (get-cell-value (get-create-cell! fr (.getColumnIndex c)))])
                 cs)
            (map (fn [c] (get-cell-value c)) cs))))

(defn get-last-column-num
  "Ermittelt die letzte Spalte einer Reihe oder einem Tabellenblatt 'x'"
  [x]
  (apply max (map (fn [c] (.getRowIndex c)) (cell-seq x))))


(defn into-seq
  [sheet-or-row]
  (vec (for [item (iterator-seq (.iterator sheet-or-row))] item)))


(defn add-values! [#^Sheet sheet values]
  "Fuegt eine Reihe von Werten (als Auflistung) nacheinander in die Zellen ein."
  (let [row (add-row-after-last! sheet)]
    (doseq [[column-index value] (partition 2 (interleave (iterate inc 0) values))]
      (set-cell-value! (.createCell row column-index) value))
    row))


(defn add-value-rows! [#^Sheet sheet rows]
  "Fuegt einem Arbeitsblatt mehrere Zeilen (Sequenz in Sequence verschachtelt)
   mit Daten hinzu."
  (doseq [values rows]
    (add-values! sheet values)))


(defn add-db-values-seq!
  [sheet data columnnames]
  (doseq [row-data data]
    (insert-db-values! 
      (add-row-after-last! sheet) 
      row-data
      columnnames)))


(defn insert-db-values!
  [row row-data columnnames]
  (let [nc (zipmap (vals columnnames) (keys columnnames))]
    (doseq [cell-data row-data]
      (set-cell-value! 
        (create-cell! row (nc (name (key cell-data)))) 
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
    (let [ct (. cell getCellType)]
      (if (not (= Cell/CELL_TYPE_NUMERIC ct))
        ct
        (if (DateUtil/isCellDateFormatted cell)
          :date
          ct)))))
(defmethod get-cell-value Cell/CELL_TYPE_BLANK   [cell]
  nil)
(defmethod get-cell-value Cell/CELL_TYPE_FORMULA [cell]
  (let [val              (. (get-formula-evaluator (.. cell getSheet getWorkbook)) evaluate cell)
        evaluated-type   (. val getCellType)]
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


;; (comment (defmulti get-cell-value-by-index
;;   "Liest den Inhalt einer Tabellenzelle wenn die Adresse der Zelle
;; über einen Index angegeben wird"
;;   (fn [sheet index] (type index)))
;; (defmethod get-cell-value-by-index PersitentVector [sheet index]
;;   (cond (every? #(= Integer (type %)))  () 
;;         (every? #(= String (type %))) (do-something2)
;;         :error (throw (Exception. "Index not Valid"))))
;; (defmethod get-cell-value-by-index String [sheet index]
;;   (CellReference. (str get-sheet-name sheet) "!" index)))

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
  (let [df (DataFormatter.)]
    (if (= String (type transformation)) ; Prüft ob Funktion oder String
      ((get column-types transformation) (get-cell-value cell) (.formatCellValue df cell))
      (transformation (get-cell-value cell) (.formatCellValue df cell)))))
  

(defmulti set-cell-value! 
  "Schreibt den als Argument (String, Date oder Double) uebergebenen
   Wert in die Tabellenzelle"
  (fn [c v] (type v)))
(defmethod set-cell-value! java.lang.String [c v] 
  (.setCellValue c (HSSFRichTextString. v)))
(defmethod set-cell-value! java.util.Date [c v] 
  (doto c
    (.setCellValue v)
    (apply-date-format! "m/d/yy")))
(defmethod set-cell-value! java.lang.Double [c v] 
  (.setCellValue c v))
(defmethod set-cell-value! java.lang.Integer [c v] 
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


(defn get-cell
  [r i]
  (try (.getCell r i)
       (catch Exception e (prn "Index doesn't exist"))))


(defmulti get-cell-by-index
  (fn [& args]
    (condp = (count args)
      1                                 ; ein Argument
      (str "")        
      2                                 ; 2 Argumente
      (str ""))))
  
(defn- get-create-cell!
 "Liefert eine Referenz auf ein 'Cell'-Objekt, das durch die Angaben
  Reihe 'r' und Zellennummer 'i' bestimmt wird. Wenn die Zelle nicht besteht
  wird eine neue angelegt."
  [^Row r 
   i]
  (let [cell (.getCell r (int i))]
    (if (nil? cell)
      (.createCell r (int i))
      cell)))


;;; ---------------------------------------------------------------------------
;;; Umwandlungsfunktionen

(defn varchar->double
  [v fv]
  (condp = (type v)
    java.lang.String   (if (= (.trim v) "")
                         nil
                         (.doubleValue (.parse (DecimalFormat. "#,###.00") v)))
    java.lang.Double   v
    nil))


(defn varchar->date
  [v fv]
  (condp = (type v)
    java.lang.String   (if (= (.trim v) "")
                         nil
                         (.parse (SimpleDateFormat. "d.M.yy") v))
    java.util.Date     v
    nil))

(defn all->varchar
  [v fv]
  ;; (condp = (type v)
  ;;   java.lang.String   
  ;;   java.lang.Double   (.format (DecimalFormat. "#,###.00") v)
  ;;   java.util.Date     (.format (SimpleDateFormat. "dd.MM.yyyy") v)
  ;;   nil)
  ;;
  fv)



(def column-types
  {"varchar"  all->varchar
   "double"   varchar->double
   "date"     varchar->date  })


(defn columnindex-value-map
  "Liest eine Tabellenzeile und erzeugt eine Map mit der Spaltennummer und dem Zellenwert.
   Zeilennummern beginnen bei null."
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
  ([sheet columnnames db-types begin-row]
     (db-values-seq sheet columnnames db-types begin-row (get-last-row-num sheet)))
  ([sheet columnnames db-types begin-row end-row]
     (let [rows (filter 
                 (fn [r] (and (>= (.getRowNum r) begin-row)(<= (.getRowNum r) end-row))) 
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
  "Liefert alle Zeilen mit Daten als Sequenz"
  ([#^Sheet sheet begin-row-num]
    (columnindex-value-map sheet  (nth (row-seq sheet) begin-row-num)))
  ([#^Row begin-row]
    (drop (- (.getRowNum begin-row) 1) (row-seq (.getSheet (.getRowNum begin-row))))))




 
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
        (.setCellValue (HSSFRichTextString. v))))))


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



