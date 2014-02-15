(ns bakyeono.litedocx.wpml
  "WordProcessingML templates & snippets used in litedocx."
  (:require [clojure.string :as str])
  (:use [bakyeono.litedocx.util]))

(def ^:const ^{:private true} content-types-xml-xmlns
  {:xmlns "http://schemas.openxmlformats.org/package/2006/content-types"})

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
            :PartName "/word/document.xml"}}
   {:tag :Override
    :attrs {:ContentType "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
            :PartName "/word/styles.xml"}}])

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
  "Returns an xml tag node when the value exists."
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
  "Returns an xml tag attributes without nil values."
  [& kvs]
  (into {}
        (remove-nil
          (for [[k v] (partition 2 kvs)]
            (when-v-kv v k)))))

(defn content-types-xml
  "Creates xml data for [Content_Types].xml file in DOCX package.\n
  Parameters:
  - resources: <vector of {:content-type <string>, :part-name <string>}>}"
  [resources]
  {:tag :Types
   :attrs content-types-xml-xmlns
   :content
   (flatten
     [;; default content types
      default-content-types
      ;; additional content types
      (for [resource resources]
        {:tag :Override
         :attrs {:ContentType (:content-type resource)
                 :PartName (:part-name resource)}
         :content nil})])})

(defn doc-props-app-xml
  "Creates xml data for docProps/app.xml file in DOCX pakcage.\n
  Parameters:
  - & properties: option, value, ...\n
  Examples:
  - (doc-props-app-xml :application \"Microsoft Office Word\" :app-version \"12.0000\")"
  [& {:as options
      :keys [application app-version]}]
  {:tag :properties:Properties
   :attrs doc-props-app-xml-xmlns
   :content
   (remove-nil
     (map when-v-tag
          [application app-version]
          [:properties:Application :properties:AppVersion]))})

(defn doc-props-core-xml
  "Creates xml data for docProps/core.xml file in DOCX pakcage.\n
  Parameters:
  - & properties: option, value, ...\n
  Examples:
  - (doc-props-core-xml :creator \"Bak Yeon O\" :last-modified-by \"Bak Yeon O\")"
  [& {:as properties
      :keys [title subject creator keywords description last-modified-by]}]
  {:tag :cp:coreProperties
   :attrs doc-props-core-xml-xmlns
   :content
   (remove-nil
     (map when-v-tag
          [title subject creator keywords description last-modified-by]
          [:dc:title :dc:subject :dc:creator :cp:keywords :dc:description :cp:lastModifiedBy]))})

(defn word-document-xml
  "Creates xml data for word/document.xml file in DOCX pakcage.\n"
  []
  {:tag :w:document
   :attrs word-xmlns
   :content
   nil})

(defn- ppr
  "Creates w:pPr tag node. Called by paragaph-style."
  [{:as options
    :keys [align
           space-before space-after space-line
           indent-left indent-right indent-first-line
           mirror-indents?]}]
  {:tag :w:pPr
   :content
   (remove-nil
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
   :content
   (remove-nil
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
  Needed for word-styles-xml.\n
  Parameters:
  - name: <string> What others call this style. Also used as style-id. Must be unique.
  - & options: option, value, ...\n
  Examples:
  - (paragraph-style \"paragraph1\")
  - (paragraph-style \"paragraph2\" :font \"Arial\" :font-color \"0000AA\" :align \"both\")\n
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
  ""
  []
  nil)

(defn word-rels-document-xml-rels
  ""
  [resources]
  nil)

