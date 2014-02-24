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
(defconst- xmlns-custom-properties "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties")
(defconst- xmlns-customXml "http://schemas.openxmlformats.org/officeDocument/2006/customXml")
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
(defconst- xmlns-pic "http://schemas.openxmlformats.org/drawingml/2006/picture")
(defconst- xmlns-ppt "urn:schemas-microsoft-com:office:powerpoint")
(defconst- xmlns-r "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
(defconst- xmlns-sl "http://schemas.openxmlformats.org/schemaLibrary/2006/main")
(defconst- xmlns-ssml "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
(defconst- xmlns-v "urn:schemas-microsoft-com:vml")
(defconst- xmlns-w "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
(defconst- xmlns-w10 "urn:schemas-microsoft-com:office:word")
(defconst- xmlns-wp "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
(defconst- xmlns-wvml "urn:schemas-microsoft-com:office:word")
(defconst- xmlns-x "urn:schemas-microsoft-com:office:excel")


;;; e -> clojure.data.xml/element
;;; Since clojure.data.xml/element is used so much,
;;; use 'node' as an abbreviation of it.
;
(defconst- node clojure.data.xml/element)

;;; Constants
(defconst- content-types-xml-xmlns
  {:xmlns "http://schemas.openxmlformats.org/package/2006/content-types"})

(defconst- relationships-xmlns 
  {:xmlns "http://schemas.openxmlformats.org/package/2006/relationships"})
(defconst- doc-props-app-xml-xmlns
  {:xmlns:vt "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
   :xmlns:properties "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"})

(defconst- doc-props-core-xml-xmlns
  {:xmlns:cp "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
   :xmlns:dcterms "http://purl.org/dc/terms/"
   :xmlns:dc "http://purl.org/dc/elements/1.1/"})

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
  [(node :w:style
         {:w:type "paragraph" :w:styleId "a" :w:default "true"}
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

(defn when-v-node
  "Returns an XML tag node when the value exists."
  [v tag]
  (cond (nil? v) nil
        (= "" v) (node tag)
        true (node tag {} (str v))))

(defn when-v-kv
  "Returns a vector of [k v] when the value exists."
  [v k]
  (if (nil? v) nil
    [k v]))

(defn make-attrs
  "Returns an XML tag attributes without nil values."
  [& kvs]
  (into {}
        (remove-nil
          (for [[k v] (partition 2 kvs)]
            (when-v-kv v k)))))

(defconst- styles-xml-resource
  {:id "rId1"
   :part-name "/word/styles.xml"
   :type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
   :content-type "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
   :target "styles.xml"})

(defn- content-type->relationship-type
  [content-type]
  (str "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
       (re-find #"[\w]*" content-type)))

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
  (node :w:document word-xmlns
        (node :w:body {}
              body
              (node :w:sectPr {}
                    (node :w:pgSz {:w:w 11907
                                   :w:h 16839
                                   :w:code 9})
                    (node :w:pgMar {:w:top 720
                                    :w:right 720
                                    :w:bottom 720
                                    :w:left 720})))))

(defn- ppr
  "Creates w:pPr tag node. Called by paragaph-style."
  [{:as options
    :keys [align
           space-before space-after space-line
           indent-left indent-right indent-first-line
           mirror-indents?]}]
  (node :w:pPr {}
        (when (or space-before space-after space-line)
          (node :w:spacing (make-attrs :w:before space-before
                                       :w:after space-after
                                       :w:line space-line
                                       :w:lineRule "auto")))
        (when (or indent-left indent-right indent-first-line)
          (node :w:ind (make-attrs :w:left indent-left
                                   :w:right indent-right
                                   :w:firstLine indent-first-line)))
        (when mirror-indents?
          (node :w:mirrorIndents))
        (when align
          (node :w:jc {:w:val align}))))

(defn- rpr
  "Creates w:rPr tag node. Called by paragaph-style."
  [{:as options
    :keys [font font-size font-color bold? italic? underline? strike?]}]
  (node :w:rPr {}
        (when font
          (node :w:rFonts {:w:ascii font
                           :w:hAnsi font
                           :w:eastAsia font}))
        (when font-size
          (node :w:sz {:w:val font-size}))
        (when font-color
          (node :w:color {:w:val font-color}))
        (when bold?
          (node :w:b {:w:val true}))
        (when italic?
          (node :w:b {:w:val true}))
        (when underline?
          (node :w:u {:w:val "single"}))
        (when strike?
          (node :w:strike {:w:val true}))))

(defn paragraph-style
  "Creates a w:style tag of paragraph. Returns w:style tag for paragraph style with
  w:style tag for the font, as a vector containing both of them.
  Needed for word-styles-xml.

  Parameters:
  - name: <string> What others call this style. Also used as style-id. Must be unique.
  - & options: option, value, ...

  Examples:
  - (paragraph-style \"paragraph1\")
  - (paragraph-style \"paragraph2\" :font \"Arial\" :font-color \"0000AA\" :align \"both\")

  See More: "
  [name
   & {:as options  
      :keys [font font-size font-color
             bold? italic? underline? strike?
             align
             space-before space-after space-line   
             indent-left indent-right indent-first-line mirror-indents?]}]
  (let [font-style-name (str "character-style-of-" name)]
    [;; paragraph style
     (node :w:style {:w:type "paragraph"
                     :w:styleId "normal"
                     :w:customStyle "true"}
           (node :w:name {:w:val name})
           (node :w:basedOn {:w:val "a"})
           (when font
             (node :w:link {:w:val font-style-name}))
           (node :w:qFormat)
           (when (or align
                     space-before space-after space-line
                     indent-left indent-right indent-first-line mirror-indents?)
             (ppr options))
           (when (or font font-size font-color bold? italic? underline?)
             (rpr options)))
     ;; font style of the paragraph style
     (when font
       (node :w:style {:w:type "character"
                       :w:styleId font-style-name
                       :w:customStyle "true"}
             (node :w:name {:w:val font-style-name})
             (node :w:basedOn {:w:val "a0"})
             (rpr options)))]))

(defn word-styles-xml
  "descripted on: http://officeopenxml.com/WPstyles.php"
  [& styles]
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
  - h-align: <string or keyword> horizontal alignment option in each cell.
  - widths: <vector of numbers> widths of cells
  - & body: content of table (tr is expected)

  Examples:
  (table \"table-style1\" :center [964 964 8072] ...)"
  [style h-align widths & body]
  (node :w:tbl {}
        (node :w:tblPr {}
              (node :w:tblStyle {:w:val style})
              (node :w:tblW {:w:w (reduce + widths)
                             :w:type "dxa"})
              (node :w:jc {:w:val (name h-align)}))
        (node :w:tblGrid {}
              (for [w widths]
                (node :w:gridCol {:w:gridCol w})))
        body))

(defn tr
  "Returns XML tag node of table.

  Parameters:
  - & body: content of row (td is expected)"
  [& body]
  (node :w:tr {}
        body))

(defn td
  "Returns XML tag node of table.

  Parameters:
  - width: width of cell
  - & body: content of cell"
  [width & body]
  (node :w:tc {}
        (node :w:tcPr {}
              (node :w:tcW {:w:w width
                            :w:type "dxa"}))
        body))

(defn p
  "Returns XML tag node of paragraph.

  Parameters:
  - style: <string or keyword> style of paragraph
  - text: <string> text in paragraph"
  ([text]
   (node :w:p {}
         (node :w:r {}
               (node :w:t {}
                     text))))
  ([style text]
   (node :w:p {}
         (node :w:pPr {}
               (node :w:pStyle {:w:val (name style)}))
         (node :w:r {}
               (node :w:t {}
                     text)))))

(defn img
  "Returns XML tag node of image.
  
  Parameters:
  ..."
  [id name desc extent-x extent-y]
  (node :w:p {}
        (node :w:r {}
              (node :w:drawing {}
                    ;; Distance from Text on ...
                    (node :wp:inline {:distT 0 ; top edge
                                      :distB 0 ; bottom edge
                                      :distL 0 ; left edge
                                      :distR 0}) ; right edge
                    ;; Extent ...
                    (node :wp:extent {:cx extent-x
                                      :cy extent-y})
                    ;; Effect Extent ...
                    (node :wp:effectExtent {:t 0 ; top
                                            :b 0 ; bottom
                                            :l 0 ; left
                                            :r 0}) ; right 
                    ;; Drawing Object Non-Visual Properties
                    (node :wp:docPr {:id id ; Unique ID
                                     :name name ; Name
                                     :descr desc}) ; Alternative text for Object 
                    ;; Common DrawingML Non-Visual Properties
                    (node :wp:cNvGraphicFramePr {}
                          ;; Graphic Frame Locks
                          (node :a:graphicFrameLocks {:noSelect false
                                                      :noMove false
                                                      :noResize false
                                                      :noChangeAspect true}))
                    ;; Graphic Object
                    (node :a:graphic {}
                          ;; Graphic Object Data
                          (node :a:graphicData {:uri "http://schemas.openxmlformats.org/drawingml/2006/picture"}
                                ;; Picture
                                (node :pic:pic {}
                                      ;; Non-Visual Picture Properties
                                      (node :pic:nvPicPr {}
                                            ;; Non-Visual Drawing Properties
                                            (node :pic:cNvPr {:id id ; Unique ID
                                                              :name name ; Name
                                                              :descr desc}) ; Description
                                            ;; Non-Visual Picture Drawing Properties
                                            (node :pic:cNvPicPr))
                                      ;; Picture Fill
                                      (node :pic:blipFill {}
                                            ;; Blip
                                            (node :a:blip {:r:embed id ; Embedded Picture Reference
                                                           :r:link ""}) ; Linked Picture Reference
                                            ;; Stretch
                                            (node :a:stretch {}
                                                  ;; Fill Rectangle
                                                  (node :a:fillRect)))
                                      ;; Shape Properties
                                      (node :pic:spPr {}
                                            ;; 2D Transform for Individual Objects
                                            (node :a:xfrm {}
                                                  ;; Offset
                                                  (node :a:off {:x 0
                                                                :y 0})
                                                  ;; Extents
                                                  (node :ext {:x extent-x
                                                              :y extent-y}))
                                            ;; Preset Geometry
                                            (node :a:prstGeom {:prst "rect"} ; Preset Shape 
                                                  ;; List of Shape Adjust Values
                                                  (node :a:avLst))))))))))

