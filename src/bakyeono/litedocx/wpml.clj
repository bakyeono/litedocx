(ns bakyeono.litedocx.wpml
  "WordProcessingML templates & snippets used in litedocx."
  (:require [clojure.string :as str])
  (:use [bakyeono.litedocx.util]))

(def ^:const ^{:private true} content-types-xml-xmlns
  {:xmlns "http://schemas.openxmlformats.org/package/2006/content-types"})

(def ^:const ^{:private true} relationships-xmlns 
  {:xmlns "http://schemas.openxmlformats.org/package/2006/relationships"})

(def ^:const ^{:private true} doc-props-app-xml-xmlns
  {:xmlns:vt "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
   :xmlns:properties "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"})

(def ^:const ^{:private true} doc-props-core-xml-xmlns
  {:xmlns:cp "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
   :xmlns:dcterms "http://purl.org/dc/terms/"
   :xmlns:dc "http://purl.org/dc/elements/1.1/"})

(def ^:const ^{:private true} word-xmlns
  {:xmlns:v "urn:schemas-microsoft-com:vml"
   :xmlns:pic "http://schemas.openxmlformats.org/drawingml/2006/picture"
   :xmlns:ns11 "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
   :xmlns:r "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
   :xmlns:o "urn:schemas-microsoft-com:office:office"
   :xmlns:odq "http://opendope.org/questions"
   :xmlns:ns13 "urn:schemas-microsoft-com:office:excel"
   :xmlns:w10 "urn:schemas-microsoft-com:office:word"
   :xmlns:dsp "http://schemas.microsoft.com/office/drawing/2008/diagram"
   :xmlns:ns8 "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"
   :xmlns:m "http://schemas.openxmlformats.org/officeDocument/2006/math"
   :xmlns:ns25 "http://schemas.openxmlformats.org/drawingml/2006/compatibility"
   :xmlns:ns26 "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"
   :xmlns:ns6 "http://schemas.openxmlformats.org/schemaLibrary/2006/main"
   :xmlns:ns24 "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
   :xmlns:odx "http://opendope.org/xpaths"
   :xmlns:odi "http://opendope.org/components"
   :xmlns:dgm "http://schemas.openxmlformats.org/drawingml/2006/diagram"
   :xmlns:a "http://schemas.openxmlformats.org/drawingml/2006/main"
   :xmlns:odc "http://opendope.org/conditions"
   :xmlns:c "http://schemas.openxmlformats.org/drawingml/2006/chart"
   :xmlns:wp "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
   :xmlns:ns17 "urn:schemas-microsoft-com:office:powerpoint"
   :xmlns:odgm "http://opendope.org/SmartArt/DataHierarchy"
   :xmlns:w "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})

(def ^:const ^{:private true} default-content-types
  [{:tag :Default
    :attrs {:Extension "rels"
            :ContentType "application/vnd.openxmlformats-package.relationships+xml"}}
   {:tag :Override
    :attrs {:ContentType "application/vnd.openxmlformats-officedocument.extended-properties+xml"
            :PartName "/docProps/app.xml"}}
   {:tag :Override
    :attrs {:ContentType "application/vnd.openxmlformats-package.core-properties+xml"
            :PartName "/docProps/core.xml"}}
   {:tag :Override
    :attrs {:ContentType "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
            :PartName "/word/document.xml"}}])

(def ^:const ^{:private true} default-styles
  [{:tag :w:style
    :attrs {:w:type "paragraph" :w:styleId "a" :w:default "true"}
    :content
    [{:tag :w:name :attrs {:w:val "Normal"} :content nil}
     {:tag :w:qFormat :attrs nil :content nil}
     {:tag :w:pPr :attrs nil :content
      [{:tag :w:widowControl :attrs {:w:val "false"} :content nil}
       {:tag :w:wordWrap :attrs {:w:val "false"} :content nil}
       {:tag :w:autoSpaceDE :attrs {:w:val "false"} :content nil}
       {:tag :w:autoSpaceDN :attrs {:w:val "false"} :content nil}
       {:tag :w:jc :attrs {:w:val "both"} :content nil}]}]}
   {:tag :w:style
    :attrs {:w:type "character" :w:styleId "a0" :w:default "true"}
    :content
    [{:tag :w:name :attrs {:w:val "Default Paragraph Font"} :content nil}
     {:tag :w:uiPriority :attrs {:w:val "1"} :content nil}
     {:tag :w:semiHidden :attrs nil :content nil}
     {:tag :w:unhideWhenUsed :attrs nil :content nil}]}
   {:tag :w:style
    :attrs {:w:type "table" :w:styleId "a1" :w:default "true"}
    :content
    [{:tag :w:name :attrs {:w:val "Normal Table"} :content nil}
     {:tag :w:uiPriority :attrs {:w:val "99"} :content nil}
     {:tag :w:semiHidden :attrs nil :content nil}
     {:tag :w:unhideWhenUsed :attrs nil :content nil}
     {:tag :w:qFormat :attrs nil :content nil}
     {:tag :w:tblPr
      :attrs nil
      :content
      [{:tag :w:tblInd :attrs {:w:w "0" :w:type "dxa"} :content nil}
       {:tag :w:tblCellMar
        :attrs nil
        :content
        [{:tag :w:top :attrs {:w:w "0" :w:type "dxa"} :content nil}
         {:tag :w:left :attrs {:w:w "108" :w:type "dxa"} :content nil}
         {:tag :w:bottom :attrs {:w:w "0" :w:type "dxa"} :content nil}
         {:tag :w:right :attrs {:w:w "108" :w:type "dxa"} :content nil}]}]}]}
   {:tag :w:style
    :attrs {:w:type "numbering" :w:styleId "a2" :w:default "true"}
    :content
    [{:tag :w:name :attrs {:w:val "No List"} :content nil}
     {:tag :w:uiPriority :attrs {:w:val "99"} :content nil}
     {:tag :w:semiHidden :attrs nil :content nil}
     {:tag :w:unhideWhenUsed :attrs nil :content nil}]}])

(defn when-v-tag
  "Returns an XML tag node when the value exists."
  [v tag]
  (cond (nil? v) nil
        (= "" v) {:tag tag}
        true {:tag tag
              :content [(str v)]}))

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

(def ^:const ^{:private true} styles-xml-resource
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
  {:tag :Override
   :attrs {:ContentType content-type
           :PartName part-name}})

(defn content-types-xml
  "Creates XML data for [Content_Types].xml file in DOCX package.

  Parameters:
  - resources: <vector of {:content-type <string>, :part-name <string>}>}

  Note that resources should be created by make-resources."
  [resources]
  {:tag :Types
   :attrs content-types-xml-xmlns
   :content (into default-content-types
                  (map resource-type-override resources))})

(defn doc-props-app-xml
  "Creates XML data for docProps/app.xml file in DOCX pakcage.

  Parameters:
  - & properties: option, value, ...

  Examples:
  - (doc-props-app-xml :application \"Microsoft Office Word\" :app-version \"12.0000\")"
  [& {:as options
      :keys [application app-version]}]
  {:tag :properties:Properties
   :attrs doc-props-app-xml-xmlns
   :content (remove-nil
              (map when-v-tag
                   [application app-version]
                   [:properties:Application :properties:AppVersion]))})

(defn rels-rels
  "Creates XML data for _rels/.rels file in DOCX package."
  []
  {:tag :Relationships
   :attrs relationships-xmlns
   :content
   [{:tag :Relationship
     :attrs
     {:Id "rId1"
      :Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" 
      :Target "word/document.xml"}}
    {:tag :Relationship
     :attrs
     {:Id "rId2"
      :Type "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
      :Target "docProps/core.xml"}}
    {:tag :Relationship
     :attrs
     {:Id "rId3"
      :Target "docProps/app.xml"
      :Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"}}]})

(defn doc-props-core-xml
  "Creates XML data for docProps/core.xml file in DOCX pakcage.

  Parameters:
  - & properties: option, value, ...

  Examples:
  - (doc-props-core-xml :creator \"Bak Yeon O\" :last-modified-by \"Bak Yeon O\")"
  [& {:as properties
      :keys [title subject creator keywords description last-modified-by]}]
  {:tag :cp:coreProperties
   :attrs doc-props-core-xml-xmlns
   :content (remove-nil
              (map when-v-tag
                   [title subject creator keywords description last-modified-by]
                   [:dc:title :dc:subject :dc:creator :cp:keywords
                    :dc:description :cp:lastModifiedBy]))})

(defn word-document-xml
  "Creates XML data for word/document.xml file in DOCX pakcage."
  [& body]
  {:tag :w:document
   :attrs word-xmlns
   :content
   [{:tag :w:body
     :content 
     (remove-nil
       (conj
         (remove-nil body) 
         {:tag :w:sectPr
          :content
          [{:tag :w:pgSz
            :attrs {:w:w 11907
                    :w:h 16839
                    :w:code 9}}
           {:tag :w:pgMar
            :attrs {:w:top 720
                    :w:right 720
                    :w:bottom 720
                    :w:left 720}}]}))}]})

(defn- ppr
  "Creates w:pPr tag node. Called by paragaph-style."
  [{:as options
    :keys [align
           space-before space-after space-line
           indent-left indent-right indent-first-line
           mirror-indents?]}]
  {:tag :w:pPr
   :content (remove-nil
              [(when (or space-before space-after space-line)
                 {:tag :w:spacing
                  :attrs (make-attrs :w:before space-before
                                     :w:after space-after
                                     :w:line space-line
                                     :w:lineRule "auto")})
               (when (or indent-left indent-right indent-first-line)
                 {:tag :w:ind
                  :attrs (make-attrs :w:left indent-left
                                     :w:right indent-right
                                     :w:firstLine indent-first-line)})
               (when mirror-indents?
                 {:tag :w:mirrorIndents})
               (when align
                 {:tag :w:jc
                  :attrs {:w:val align}})])})

(defn- rpr
  "Creates w:rPr tag node. Called by paragaph-style."
  [{:as options
    :keys [font font-size font-color bold? italic? underline? strike?]}]
  {:tag :w:rPr
   :content (remove-nil
              [(when font
                 {:tag :w:rFonts
                  :attrs {:w:ascii font
                          :w:hAnsi font
                          :w:eastAsia font}})
               (when font-size
                 {:tag :w:sz
                  :attrs {:w:val font-size}})
               (when font-color
                 {:tag :w:color
                  :attrs {:w:val font-color}})
               (when bold?
                 {:tag :w:b
                  :attrs {:w:val true}})
               (when italic?
                 {:tag :w:b
                  :attrs {:w:val true}})
               (when underline?
                 {:tag :w:u
                  :attrs {:w:val "single"}})
               (when strike?
                 {:tag :w:strike
                  :attrs {:w:val true}})])})

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
    (remove-nil
      [;; paragraph style
       {:tag :w:style
        :attrs {:w:type "paragraph"
                :w:styleId "normal"
                :w:customStyle "true"}
        :content
        (remove-nil
          [{:tag :w:name
            :attrs {:w:val name}}
           {:tag :w:basedOn
            :attrs {:w:val "a"}}
           (when font {:tag :w:link
                       :attrs {:w:val font-style-name}})
           {:tag :w:qFormat}
           (when (or align
                     space-before space-after space-line
                     indent-left indent-right indent-first-line mirror-indents?)
             (ppr options))
           (when (or font font-size font-color bold? italic? underline?)
             (rpr options))])}
       ;; font style of the paragraph style
       (when font
         {:tag :w:style
          :attrs {:w:type "character"
                  :w:styleId font-style-name
                  :w:customStyle "true"}
          :content
          [{:tag :w:name
            :attrs {:w:val font-style-name}}
           {:tag :w:basedOn
            :attrs {:w:val "a0"}}
           (rpr options)]})])))

(defn word-styles-xml
  "descripted on: http://officeopenxml.com/WPstyles.php"
  [& styles]
  {:tag :w:styles
   :attrs word-xmlns
   :content (remove-nil (flatten [default-styles styles]))})

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
  {:tag :Relationship
   :attrs {:Id id
           :Type type
           :Target target}})

(defn word-rels-document-xml-rels
  "Creates XML data for word/_rels/document.xml.rels file in DOCX package.

  Parameters:
  - resources: <vector of {:id <number>, :type <string>, :target <string>}>}

  Note that resources should be created by make-resources."
  [resources]
  {:tag :Relationships
   :attrs relationships-xmlns
   :content (map rels resources)})

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
  {:tag :w:tbl
   :content
   (remove-nil
     (into
       [{:tag :w:tblPr
         :content
         [{:tag :w:tblStyle
           :attrs {:w:val style}}
          {:tag :w:tblW
           :attrs {:w:w (reduce + widths)
                   :w:type "dxa"}}
          {:tag :w:jc
           :attrs {:w:val (name h-align)}}]}
        {:tag :w:tblGrid
         :content
         (for [w widths]
           {:tag :w:gridCol
            :attrs {:w:gridCol w}})}]
       body))})

(defn tr
  "Returns XML tag node of table.

  Parameters:
  - & body: content of row (td is expected)"
  [& body]
  {:tag :w:tr
   :content (remove-nil body)})

(defn td
  "Returns XML tag node of table.

  Parameters:
  - width: width of cell
  - & body: content of cell"
  [width & body]
  {:tag :w:tc
   :content
   (remove-nil
     (into
       [{:tag :w:tcPr
         :content
         [{:tag :w:tcW
           :attrs {:w:w width
                   :w:type "dxa"}}]}]
       body))})

(defn p
  "Returns XML tag node of paragraph.

  Parameters:
  - style: <string or keyword> style of paragraph
  - text: <string> text in paragraph"
  ([text]
   {:tag :w:p
    :content
    [{:tag :w:r
      :content
      [{:tag :w:t
        :content [text]}]}]})
  ([style text]
   {:tag :w:p
    :content
    [{:tag :w:pPr
      :content
      [{:tag :w:pStyle
        :attrs {:w:val (name style)}}]}
     {:tag :w:r
      :content
      [{:tag :w:rPr
        :content
        [{:tag :w:rFonts
          :attrs {:w:hint "eastAsia"}}]}
       {:tag :w:t
        :content [text]}]}]})) 

