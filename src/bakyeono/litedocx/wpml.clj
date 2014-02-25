(ns bakyeono.litedocx.wpml
  "WordProcessingML templates & snippets used in litedocx."
  (:require [clojure.string :as str])
  (:require [clojure.data.xml :as xml])
  (:use [bakyeono.litedocx.util]))

;;; OOXML Namespaces
(defconst- xmlns-a "http://schemas.openxmlformats.org/drawingml/2006/main")
(defconst- xmlns-bibliography "http://schemas.openxmlformats.org/officeDocument/2006/bibliography")
(defconst- xmlns-c "http://schemas.openxmlformats.org/drawingml/2006/chart")
(defconst- xmlns-cdr "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing")
(defconst- xmlns-characteristics "http://schemas.openxmlformats.org/officeDocument/2006/characteristics")
(defconst- xmlns-core-properties "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
(defconst- xmlns-custom-properties "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties")
(defconst- xmlns-customXml "http://schemas.openxmlformats.org/officeDocument/2006/customXml")
(defconst- xmlns-dc "http://purl.org/dc/elements/1.1/")
(defconst- xmlns-dcterms "http://purl.org/dc/terms/")
(defconst- xmlns-dgm "http://schemas.openxmlformats.org/drawingml/2006/diagram")
(defconst- xmlns-docPropsVTypes "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
(defconst- xmlns-draw-chart "http://schemas.openxmlformats.org/drawingml/2006/chart")
(defconst- xmlns-draw-compat "http://schemas.openxmlformats.org/drawingml/2006/compatibility")
(defconst- xmlns-draw-diag "http://schemas.openxmlformats.org/drawingml/2006/diagram")
(defconst- xmlns-draw-lc "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas")
(defconst- xmlns-draw-pic "http://schemas.openxmlformats.org/drawingml/2006/picture")
(defconst- xmlns-draw-ssdraw "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
(defconst- xmlns-dsp "http://schemas.microsoft.com/office/drawing/2008/diagram")
(defconst- xmlns-extended-properties "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")
(defconst- xmlns-m "http://schemas.openxmlformats.org/officeDocument/2006/math")
(defconst- xmlns-ns11 "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
(defconst- xmlns-ns13 "urn:schemas-microsoft-com:office:excel")
(defconst- xmlns-ns17 "urn:schemas-microsoft-com:office:powerpoint")
(defconst- xmlns-ns24 "http://schemas.openxmlformats.org/officeDocument/2006/bibliography")
(defconst- xmlns-ns25 "http://schemas.openxmlformats.org/drawingml/2006/compatibility")
(defconst- xmlns-ns26 "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas")
(defconst- xmlns-ns6 "http://schemas.openxmlformats.org/schemaLibrary/2006/main")
(defconst- xmlns-ns8 "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing")
(defconst- xmlns-o "urn:schemas-microsoft-com:office:office")
(defconst- xmlns-odc "http://opendope.org/conditions")
(defconst- xmlns-odgm "http://opendope.org/SmartArt/DataHierarchy")
(defconst- xmlns-odi "http://opendope.org/components")
(defconst- xmlns-odq "http://opendope.org/questions")
(defconst- xmlns-odx "http://opendope.org/xpaths")
(defconst- xmlns-p "http://schemas.openxmlformats.org/presentationml/2006/main")
(defconst- xmlns-pack-content-types "http://schemas.openxmlformats.org/package/2006/content-types")
(defconst- xmlns-pack-r "http://schemas.openxmlformats.org/package/2006/relationships")
(defconst- xmlns-pic "http://schemas.openxmlformats.org/drawingml/2006/picture")
(defconst- xmlns-ppt "urn:schemas-microsoft-com:office:powerpoint")
(defconst- xmlns-r "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
(defconst- xmlns-sl "http://schemas.openxmlformats.org/schemaLibrary/2006/main")
(defconst- xmlns-ssml "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
(defconst- xmlns-v "urn:schemas-microsoft-com:vml")
(defconst- xmlns-vt "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
(defconst- xmlns-w "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
(defconst- xmlns-w10 "urn:schemas-microsoft-com:office:word")
(defconst- xmlns-wp "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
(defconst- xmlns-wvml "urn:schemas-microsoft-com:office:word")
(defconst- xmlns-x "urn:schemas-microsoft-com:office:excel")

;;; e -> clojure.data.xml/element
;;; Since clojure.data.xml/element is used so much,
;;; use 'node' as an abbreviation of it.
(defconst- node clojure.data.xml/element)

;;; Constants
(defconst- content-types-xml-xmlns
  {:xmlns xmlns-pack-content-types})

(defconst- relationships-xmlns 
  {:xmlns xmlns-pack-r})

(defconst- doc-props-app-xml-xmlns
  {:xmlns:vt xmlns-vt
   :xmlns:properties xmlns-extended-properties})

(defconst- doc-props-core-xml-xmlns
  {:xmlns:cp xmlns-core-properties
   :xmlns:dc xmlns-dc
   :xmlns:dcterms xmlns-dcterms})

(defconst- word-xmlns
  {:xmlns:a xmlns-a
   :xmlns:c xmlns-c
   :xmlns:dgm xmlns-dgm
   :xmlns:dsp xmlns-dsp
   :xmlns:m xmlns-m
   :xmlns:ns11 xmlns-ns11
   :xmlns:ns13 xmlns-ns13
   :xmlns:ns17 xmlns-ns17
   :xmlns:ns24 xmlns-ns24
   :xmlns:ns25 xmlns-ns25
   :xmlns:ns26 xmlns-ns26
   :xmlns:ns6 xmlns-ns6
   :xmlns:ns8 xmlns-ns8
   :xmlns:o xmlns-o
   :xmlns:odc xmlns-odc
   :xmlns:odgm xmlns-odgm
   :xmlns:odi xmlns-odi
   :xmlns:odq xmlns-odq
   :xmlns:odx xmlns-odx
   :xmlns:pic xmlns-pic
   :xmlns:r xmlns-r
   :xmlns:v xmlns-v
   :xmlns:w xmlns-w
   :xmlns:w10 xmlns-w10
   :xmlns:wp xmlns-wp})

(defconst- default-content-types
  [(node :Default {:Extension "rels"
                   :ContentType "application/vnd.openxmlformats-package.relationships+xml"})
   (node :Override {:ContentType "application/vnd.openxmlformats-officedocument.extended-properties+xml"
                    :PartName "/docProps/app.xml"})
   (node :Override {:ContentType "application/vnd.openxmlformats-package.core-properties+xml"
                    :PartName "/docProps/core.xml"})
   (node :Override {:ContentType "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                    :PartName "/word/document.xml"})])

(defconst- default-styles
  [(node :w:style {:w:type "paragraph" :w:styleId "a" :w:default "true"}
         (node :w:name {:w:val "Normal"})
         (node :w:qFormat)
         (node :w:pPr {}
               (node :w:widowControl {:w:val "false"})
               (node :w:wordWrap {:w:val "false"})
               (node :w:autoSpaceDE {:w:val "false"})
               (node :w:autoSpaceDN {:w:val "false"})
               (node :w:jc {:w:val "both"})))
   (node :w:style {:w:type "character" :w:styleId "a0" :w:default "true"}
         (node :w:name {:w:val "Default Paragraph Font"})
         (node :w:uiPriority {:w:val "1"})
         (node :w:semiHidden)
         (node :w:unhideWhenUsed))
   (node :w:style {:w:type "table" :w:styleId "a1" :w:default "true"}
         (node :w:name {:w:val "Normal Table"})
         (node :w:uiPriority {:w:val "99"})
         (node :w:semiHidden)
         (node :w:unhideWhenUsed)
         (node :w:qFormat)
         (node :w:tblPr {}
               (node :w:tblInd {:w:w "0" :w:type "dxa"})
               (node :w:tblCellMar {}
                     (node :w:top {:w:w "0" :w:type "dxa"})
                     (node :w:left {:w:w "108" :w:type "dxa"})
                     (node :w:bottom {:w:w "0" :w:type "dxa"})
                     (node :w:right {:w:w "108" :w:type "dxa"}))))
   (node :w:style {:w:type "numbering" :w:styleId "a2" :w:default "true"}
         (node :w:name {:w:val "No List"})
         (node :w:uiPriority {:w:val "99"})
         (node :w:semiHidden)
         (node :w:unhideWhenUsed))])

(defn- when-v-node
  "Returns an XML tag node when the value exists."
  [v tag]
  (cond (nil? v) nil
        (= "" v) (node tag)
        true (node tag {} (str v))))

(defn- when-v-kv
  "Returns a vector of [k v] when the value exists."
  [v k]
  (if (nil? v) nil
    [k v]))

(defn- when-v-attrs
  "Returns an XML tag attributes without nil values."
  [& kvs]
  (into {}
        (remove-nil
          (for [[k v] (partition 2 kvs)]
            (when-v-kv v k)))))

(defconst- styles-xml-resource
  {:id "rId1"
   :part-name "/word/styles.xml"
   :type (str xmlns-r "/styles")
   :content-type "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
   :target "styles.xml"})

(defn- content-type->relationship-type
  "Called by make-resource"
  [content-type]
  (str xmlns-r "/" (re-find #"[\w]*" content-type)))

(defn- make-resource
  "Creates an external resource information map for DOCX pakcage.
  Called by make-resources"
  [i [content-type filename body]]
  {:id (str "rId" i)
   :part-name (str "/word/media/" filename)
   :type (content-type->relationship-type content-type)
   :content-type content-type
   :target (str "media/" filename)
   :path (str "word/media/" filename)
   :body body})

(defn make-resources
  "Creates a vector of external resource information maps for DOCX pakcage.

  Parameters:
  - & specs: [content-type filename body] of the resource, ...

  Examples:
  - (make-resources [\"image/png\" \"embeded_image1.png\" FILE_BODY] ...)"
  [& specs]
  (let [custom-resources (map make-resource
                              (range 2 (+ 2 (count specs)))
                              specs)] 
    (cons styles-xml-resource custom-resources)))

(defn- resource-type-override
  "Returns Override XML tag node of given resource information map.
  Called by content-types-xml."
  [{:keys [content-type part-name]}]
  (node :Override {:ContentType content-type
                   :PartName part-name}))

(defn content-types-xml
  "Creates XML data for [Content_Types].xml file in DOCX package.

  Parameters:
  - resources: <vector of {:content-type <string>, :part-name <string>}>}

  Note that resources should be created by make-resources."
  [resources]
  (node :Types content-types-xml-xmlns
        default-content-types
        (map resource-type-override resources)))

(defn doc-props-app-xml
  "Creates XML data for docProps/app.xml file in DOCX pakcage.

  Parameters:
  - & properties: option, value, ...

  Examples:
  - (doc-props-app-xml :application \"Microsoft Office Word\" :app-version \"12.0000\")"
  [& {:as options
      :keys [application app-version]}]
  (node :properties:Properties doc-props-app-xml-xmlns
        (map when-v-node
             [application app-version]
             [:properties:Application :properties:AppVersion])))

(defn rels-rels
  "Creates XML data for _rels/.rels file in DOCX package."
  []
  (node :Relationships relationships-xmlns
        (node :Relationship {:Id "rId1"
                             :Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" 
                             :Target "word/document.xml"})
        (node :Relationship {:Id "rId2"
                             :Type "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                             :Target "docProps/core.xml"})
        (node :Relationship {:Id "rId3"
                             :Target "docProps/app.xml"
                             :Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"})))

(defn doc-props-core-xml
  "Creates XML data for docProps/core.xml file in DOCX pakcage.

  Parameters:
  - & properties: option, value, ...

  Examples:
  - (doc-props-core-xml :creator \"Bak Yeon O\" :last-modified-by \"Bak Yeon O\")"
  [& {:as properties
      :keys [title subject creator keywords description last-modified-by]}]
  (node :cp:coreProperties doc-props-core-xml-xmlns
        (map when-v-node
             [title subject creator keywords description last-modified-by]
             [:dc:title :dc:subject :dc:creator :cp:keywords
              :dc:description :cp:lastModifiedBy])))

(defn word-document-xml
  "Creates XML data for word/document.xml file in DOCX pakcage."
  [& body]
  ;; Document
  (node :w:document word-xmlns
        ;; w:body
        (node :w:body {}
              body
              ;; Document Final Section Properties
              (node :w:sectPr {}
                    ;; Page Size
                    (node :w:pgSz {:w:w 11907 ; Page Width
                                   :w:h 16839 ; Page Height
                                   :w:orient "portrait" ; Page Orientation
                                   :w:code 9}) ; Printer Paper Code
                    ;; Page Margins
                    (node :w:pgMar {:w:top 720
                                    :w:right 720
                                    :w:bottom 720
                                    :w:left 720
                                    :w:header 0
                                    :w:footer 0})))))

(defn- run-properties
  "Creates w:rPr tag node. Called by paragaph-style."
  [{:as options
    :keys [font font-size font-color bold? italics? underline? strike?]}]
  ;; Run Properties
  (node :w:rPr {}
        (when font
          ;; Run Fonts
          (node :w:rFonts {:w:hint "default" ; Font Type Hint: default, eastAsia, cs
                           :w:ascii font ; ASCII Font
                           :w:hAnsi font ; High ASCII Font
                           :w:eastAsia font ; East Asian Font
                           :w:cs font})) ; Complex Script Font
        (when font-size
          ;; Font Size
          (node :w:sz {:w:val font-size}))
        (when font-color
          ;; Run Content Color
          (node :w:color {:w:val font-color}))
        (when bold?
          ;; Bold
          (node :w:b {:w:val true}))
        (when italics?
          ;; Italics
          (node :w:i {:w:val true}))
        (when underline?
          ;; Underline
          (node :w:u {:w:val "single"})) ; Underline Style
        (when strike?
          ;; Single Strikethrough
          (node :w:strike {:w:val true}))))

(defn character-style
  "Returns w:style tag for character style.

  Parameters:
  - name: <string> Used as index. Must be unique.
  - & options: option, value, ...

  Examples:
  - (character-style \"font1\")
  - (character-style \"font2\" :font \"Arial\" :font-color \"0000AA\" :align \"both\")

  See More: "
  [name
   & {:as options
      :keys [font font-size font-color
             bold? italics? underline? strike?]}]
  ;; Font Style Definition for above Paragraph Style
  (node :w:style {:w:type "character" ; Style Type
                  :w:styleId name ; Style ID
                  :w:customStyle "true"} ; User-Defined Style
        (node :w:name {:w:val name}) ; Primary Style Name
        (node :w:basedOn {:w:val "a0"}) ; Parent Style ID
        (run-properties options)))

(defn- paragraph-properties
  "Creates w:pPr tag node. Called by paragaph-style."
  [{:as options
    :keys [word-wrap?
           align
           space-before space-after space-line
           indent-left indent-right indent-first-line
           mirror-indents?]}]
  ;; Paragraph Properties
  (node :w:pPr {}
        (when word-wrap?
          ;; Allow Line Breaking at Character Level
          (node :w:wordWrap {:w:val true}))
        (when (or space-before space-after space-line)
          ;; Spacing Between Lines and Above/Below Paragraph
          (node :w:spacing (when-v-attrs :w:before space-before ; Spacing Above Paragraph
                                         :w:after space-after ; Spacing Below Paragraph
                                         :w:line space-line ; Spacing Between Lines in Paragraph
                                         :w:lineRule "auto"))) ; Type of Spacing Between Lines
        (when (or indent-left indent-right indent-first-line)
          ;; Paragraph Indentation
          (node :w:ind (when-v-attrs :w:left indent-left ; Left Indentation
                                     :w:right indent-right ; Right Indentation
                                     :w:firstLine indent-first-line))) ; Additional First Line Indentation
        (when mirror-indents?
          ;; Use Left/Right Indents as Inside/Outside Indents
          (node :w:mirrorIndents))
        (when align
          ;; Paragraph Alignment
          (node :w:jc {:w:val align}))))

(defn paragraph-style
  "Returns w:style tag for paragraph style with w:style tag for the font,
  as a vector containing both of them.

  Parameters:
  - name: <string> Used as index. Must be unique.
  - & options: option, value, ...

  Examples:
  - (paragraph-style \"paragraph1\")
  - (paragraph-style \"paragraph2\" :font \"Arial\" :font-color \"0000AA\" :align \"both\")

  See More: "
  [name
   & {:as options
      :keys [font font-size font-color
             bold? italics? underline? strike?
             word-wrap?
             align
             space-before space-after space-line
             indent-left indent-right indent-first-line mirror-indents?]}]
  ;; Paragraph Style Definition
  (node :w:style {:w:type "paragraph" ; Style Type
                  :w:styleId name ; Style ID
                  :w:customStyle "true"} ; User-Defined Style
        (node :w:name {:w:val name}) ; Primary Style Name
        (node :w:basedOn {:w:val "a"}) ; Parent Style ID
        (node :w:qFormat) ; Primary Style
        (paragraph-properties options)
        (run-properties options)))

(defn- table-border-attrs
  "Returns an attributes map for content tags of w:tblBorders tag.
  Called by table-borders."
  [[style color width space]]
  {:w:val style ; Border Style : single, ...
   :w:color color ; Border Color : auto, ...
   :w:sz width ; Border Width
   :w:space space}) ; Border Spacing Measurement)

(defn- table-borders
  "Returns w:tblBorders tag. Called by table-properties.

  Parameters:
  Takes one or six parameters.
  If only one parameter is given, takes it as every side of borders.
  When six parameters are fiven, takes them as top, left, bottom, right,
  inside-horizontal, inside-vertical sides of borders.
  Each parameter is a sequence of [style color width space] attributes of
  each side of borders."
  ([side]
   (table-borders side side side side side side))
  ([top left bottom right inside-h inside-v]
   (node :w:tblBorders {}
     (node :w:top (table-border-attrs top))
     (node :w:left (table-border-attrs left))
     (node :w:bottom (table-border-attrs bottom))
     (node :w:right (table-border-attrs right))
     (node :w:insideH (table-border-attrs inside-h))
     (node :w:insideV (table-border-attrs inside-v)))))

(defn- table-properties
  "Creates w:tblPr tag node. Called by table-style.

  Options: See 'Options' of table-style"
  [{:as options
    :keys [table-align
           cell-margin-top cell-margin-left cell-margin-bottom cell-margin-right]}]
  ;; Table Properties
  (node :w:tblPr {}
        ;; Number of Rows in Row Band
        (node :w:tblStyleRowBandSize {:w:val 1})
        ;; Table Alignment
        (node :w:jc {:w:val table-align})
        ;; Table Indent from Leading Margin
        (node :w:tblInd {:w:val 0 :w:type "dxa"})
        ;; Table Borders
        (table-borders ["single" "auto" 4 0])
        ;; Table Cell Margin
        (node :w:tblCellMar {}
              (node :w:top {:w:w cell-margin-top :w:type "dxa"})
              (node :w:left {:w:w cell-margin-left :w:type "dxa"})
              (node :w:bottom {:w:w cell-margin-bottom :w:type "dxa"})
              (node :w:right {:w:w cell-margin-right :w:type "dxa"}))))

(defn- table-row-properties
  "Creates w:trPr tag node. Called by table-style."
  [{:as options
    :keys [cell-h-align]}]
  ;; Table Row Properties
  (node :w:trPr {}
        ;; Table Row Alignment
        (node :w:jc {:w:val cell-h-align})))

(defn- table-cell-properties
  "Creates w:tcPr tag node. Called by table-style."
  [{:as options
    :keys [cell-v-align]}]
  ;; Table Cell Properties
  (node :w:tcPr {}
        ;; Table Cell Vertical Alignment
        (node :w:vAlign {:w:val cell-v-align})))

(defn table-style
  "Returns w:style tag for table style with w:style tag for the font,
  as a vector containing both of them.

  Parameters:
  - name: <string> Used as index. Must be unique.
  - & options: option, value, ...

  Options:
  - table-align: Alignment of table itself. Values: left, right, both, ...
  - cell-margin-top: Margin in this side of cell. Value: numbers.
  - cell-margin-left: Same as above.
  - cell-margin-bottom: Same as above.
  - cell-margin-right: Same as above.
  - cell-h-align: Horizontal alignment inside cells. Values: left, right, both, ...
  - cell-v-align: Vertical alignment inside cells. Values: top, bottom, center, ...

  Examples:
  - (table-style \"table1\")
  - (table-style \"table2\" :font \"Arial\" :font-color \"0000AA\" :align \"both\")

  See More: "
  [name
   & {:as options
      :keys [font font-size font-color
             bold? italics? underline? strike?
             word-wrap?
             align
             space-before space-after space-line
             indent-left indent-right indent-first-line mirror-indents?
             table-align
             cell-margin-top cell-margin-bottom cell-margin-left cell-margin-right
             cell-h-align cell-v-align]}]
  ;; Table Style Definition
  (node :w:style {:w:type "table" ; Style Type
                  :w:styleId name ; Style ID
                  :w:customStyle "true"} ; User-Defined Style
        (node :w:name {:w:val name}) ; Primary Style Name
        (node :w:basedOn {:w:val "a1"}) ; Parent Style ID
        (node :w:qFormat) ; Primary Style
        (paragraph-properties options)
        (run-properties options)
        (table-properties options)
        (table-row-properties options)
        (table-cell-properties options)
        ;; First Row Formatting Properties
        (node :w:tblStylePr {:w:type "firstRow"})
        ;; Band 1 Horizontal Formatting Properties
        (node :w:tblStylePr {:w:type "band1Horz"})
        ;; Band 2 Horizontal Formatting Properties
        (node :w:tblStylePr {:w:type "band2Horz"})))

(defn word-styles-xml
  "descripted on: http://officeopenxml.com/WPstyles.php"
  [& styles]
  ;; Style Definitions
  (node :w:styles word-xmlns
        (flatten [default-styles styles])))

(defn word-media-resources
  "Creates a vector of resource data map. Each map has :path and :byte-array
  keys and their values.
  Parameters:
  - resources: <vector of {:path <string>, :body <byte-array>}>}

  Note that resources should be created by make-resources."
  [resources]
  (flatten
    (for [{:keys [path body]} (drop 1 resources)]
      [path body])))

(defn- rels
  "Returns Relationship XML tag node of given resource information map.
  Called by word-rels-document-xml-rels."
  [{:keys [id type target]}]
  (node :Relationship {:Id id
                       :Type type
                       :Target target}))

(defn word-rels-document-xml-rels
  "Creates XML data for word/_rels/document.xml.rels file in DOCX package.

  Parameters:
  - resources: <vector of {:id <number>, :type <string>, :target <string>}>}

  Note that resources should be created by make-resources."
  [resources]
  (node :Relationships relationships-xmlns
        (map rels resources)))

(defn table
  "Returns XML tag node of table.

  Parameters:
  - style: <string> style of table
  - align: <string or keyword> horizontal alignment option in each cell.
  - widths: <vector of numbers> widths of cells
  - & body: content of table (tr is expected)

  Examples:
  (table \"table-style1\" :center [964 964 8072] ...)"
  [style align widths & body]
  ;; Table
  (node :w:tbl {}
        ;; Table Properties
        (node :w:tblPr {}
              ;; Referenced Table Style
              (node :w:tblStyle {:w:val style})
              ;; Preferred Table Width
              (node :w:tblW {:w:w (reduce + widths)
                             :w:type "dxa"}) ; Table Width Units : nil, pct, dxa, auto
              ;; Table Alignment
              (node :w:jc {:w:val (name align)}))
        ;; Table Grid
        (node :w:tblGrid {}
              (for [w widths]
                ;; Grid Column Width
                (node :w:gridCol {:w:gridCol w})))
        body))

(defn tr
  "Returns XML tag node of table row.

  Parameters:
  - & body: content of row (td is expected)"
  [& body]
  ;; Table Row
  (node :w:tr {}
        body))

(defn td
  "Returns XML tag node of table cell.

  Parameters:
  - width: width of cell
  - & body: content of cell"
  [width & body]
  ;; Table Cell
  (node :w:tc {}
        ;; Table Cell Properties
        (node :w:tcPr {}
              ;; Preferred Table Cell Width
              (node :w:tcW {:w:w width
                            :w:type "dxa"}))
        body))

(defn p
  "Returns XML tag node of paragraph.

  Parameters:
  - style: <string or keyword> style of paragraph
  - text: <string> text in paragraph"
  ([text]
   ;; Paragraph
   (node :w:p {}
         ;; Text Run
         (node :w:r {}
               ;; Text
               (node :w:t {}
                     text))))
  ([style text]
   ;; Paragraph
   (node :w:p {}
         ;; Paragraph Properties
         (node :w:pPr {}
               ;; Referenced Paragraph Style
               (node :w:pStyle {:w:val (name style)}))
         ;; Text Run
         (node :w:r {}
               ;; Text
               (node :w:t {}
                     text)))))

(defn graphic
  "Returns a:graphic tag."
  [resource-id
   {:as options
    :keys [name desc width height]}]
  ;; Graphic Object
  (node :a:graphic {}
        ;; Graphic Object Data
        (node :a:graphicData {:uri "http://schemas.openxmlformats.org/drawingml/2006/picture"}
              ;; Picture
              (node :pic:pic {}
                    ;; Non-Visual Picture Properties
                    (node :pic:nvPicPr {}
                          ;; Non-Visual Drawing Properties
                          (node :pic:cNvPr {:id 0 ; Unique ID
                                            :name name ; Name
                                            :descr desc}) ; Description
                          ;; Non-Visual Picture Drawing Properties
                          (node :pic:cNvPicPr))
                    ;; Picture Fill
                    (node :pic:blipFill {}
                          ;; Blip
                          (node :a:blip {:r:embed resource-id}) ; Embedded Picture Reference
                          ;; Stretch
                          (node :a:stretch {}
                                ;; Fill Rectangle
                                (node :a:fillRect)))
                    ;; Shape Properties
                    (node :pic:spPr {}
                          ;; 2D Transform for Individual Objects
                          (node :a:xfrm {}
                                ;; Offset
                                (node :a:off {:x 0 :y 0})
                                ;; Extents
                                (node :ext {:cx width :cy height}))
                          ;; Preset Geometry
                          (node :a:prstGeom {:prst "rect"} ; Preset Shape 
                                ;; List of Shape Adjust Values
                                (node :a:avLst)))))))

(defn img
  "Returns XML tag node of image.

  Parameters:
  ..."
  [resource-id
   & {:as options
      :keys [name desc width height]}]
  (node :w:p {}
        (node :w:r {}
              (node :w:drawing {}
                    ;; Distance from Text on ...
                    (node :wp:inline {:distT 0 ; top edge
                                      :distB 0 ; bottom edge
                                      :distL 0 ; left edge
                                      :distR 0} ; right edge
                          ;; Extent ...
                          (node :wp:extent {:cx width :cy height})
                          ;; Effect Extent ...
                          (node :wp:effectExtent {:t 0 :b 0 :l 0 :r 0})
                          ;; Drawing Object Non-Visual Properties
                          (node :wp:docPr {:id 0 ; Unique ID
                                           :name name ; Name
                                           :descr desc}) ; Alternative text for Object
                          ;; Common DrawingML Non-Visual Properties
                          (node :wp:cNvGraphicFramePr {}
                                ;; Graphic Frame Locks
                                (node :a:graphicFrameLocks {:noSelect true
                                                            :noMove false
                                                            :noResize false
                                                            :noChangeAspect true}))
                          ;; Graphic
                          (graphic resource-id options))))))

