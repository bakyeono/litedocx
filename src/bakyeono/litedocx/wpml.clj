(ns bakyeono.litedocx.wpml
  "WordProcessingML templates & snippets used in litedocx."
  (:require [clojure.string :as str])
  (:use [bakyeono.litedocx.util]))

(defn when-v-tag
  "Returns tag and value when the value exists."
  [v tag]
  (cond (nil? v) nil
        (= "" v) (format "<%s/>" tag)
        true (format "<%s>%s</%s>" tag v tag)))

(defn content-types-xml
  "Creates xml code of [Content_Types].xml file for DOCX package.\n
  Parameters:
  - resources: <vector of Resource>}"
  [resources]
  (str
    ;; header
    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
    ;; default extension
    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
    ;; default files
    "<Override ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\" PartName=\"/docProps/app.xml\"/>"
    "<Override ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\" PartName=\"/docProps/core.xml\"/>"
    "<Override ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\" PartName=\"/word/document.xml\"/>"
    ;; resources
    (mapstr #(format "<Override ContentType=\"%s\" PartName=\"%s\"/>"
                     (:content-type %)
                     (:part-name %))
            resources)
    ;; default files
    "<Override ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\" PartName=\"/word/styles.xml\"/>"
    ;; footer
    "</Types>"))

(defn doc-props-app-xml
  "Creates a xml code of docProps/app.xml file for DOCX pakcage.\n
  Parameters:
  - & properties: option, value, ...\n
  Examples:
  - (doc-props-app-xml :application \"Microsoft Office Word\" :app-version \"12.0000\")"
  [& {:keys [application company app-version]}]
  (str
    ;; header
    "<properties:Properties xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\" xmlns:properties=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">"
    ;; properties
    (mapstr when-v-tag
            [application company app-version]
            ["Application" "Company" "AppVersion"])
    ;; footer
    "</properties:Properties>"))

(defn doc-props-core-xml
  "Creates a xml code of docProps/core.xml file for DOCX pakcage.\n
  Parameters:
  - & properties: option, value, ...\n
  Examples:
  - (doc-props-core-xml :creator \"Bak Yeon O\" :last-modified-by \"Bak Yeon O\")"
  [& {:as properties
      :keys [title subject creator keywords description last-modified-by]}]
  (str
    ;; header
    "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\">"
    ;; properties
    (mapstr when-v-tag
            [title subject creator keywords description last-modified-by]
            ["dc:title" "dc:subject" "dc:creator" "cp:keywords" "dc:description" "cp:lastModifiedBy"])
    ;; footer
    "</cp:coreProperties>"))

(defn word-document-xml
  ""
  []
  nil)

(defn- rpr
  "Creates w:rPr tag. Called by paragaph-style."
  [{:as options :keys [font font-size font-color bold? italic? underline? strike?]}]
  (str
    "<w:rPr>"
    (when font (format "<w:rFonts w:ascii=\"%s\" w:hAnsi=\"%s\" w:eastAsia=\"%s\"/>"
                       font font font))
    (when font-size (format "<w:sz w:val=\"%s\"/>" font-size))
    (when font-color (format "<w:color w:val=\"%s\"/>" font-color))
    (when bold? "<w:b w:val=\"true\"/>")
    (when italic? "<w:i w:val=\"true\"/>")
    (when underline? "<w:u w:val=\"single\"/>")
    (when strike? "<w:strike w:val=\"true\"/>")
    "</w:rPr>"))

(defn- ppr
  "Creates w:pPr tag. Called by paragaph-style."
  [{:as options :keys [align
                       space-before space-after space-line
                       indent-left indent-right indent-first-line
                       mirror-indents?]}]
  (str
    "<w:pPr>"
    (when (or space-before space-after space-line)
      (str "<w:spacing"
           (when space-before (format " w:before=\"%s\"" space-before))
           (when space-after (format " w:after=\"%s\"" space-after))
           (when space-line (format " w:line=\"%s\" w:lineRule=\"auto\"" space-line))
           "/>"))
    (when (or indent-left indent-right indent-first-line)
      (str "<w:ind"
           (when indent-left (format " w:left=\"%s\"" indent-left))
           (when indent-right (format " w:right=\"%s\"" indent-right))
           (when indent-first-line (format " w:firstLine=\"%s\"" indent-first-line))
           "/>"))
    (when mirror-indents? "<w:mirrorIndents/>")
    (when align (format "<w:jc w:val=\"%s\"/>" align))
    "</w:pPr>"))

(defn paragraph-style
  "Creates a w:style tag of paragraph. Needed for word-styles-xml.\n
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
  (str
    (format "<w:style w:type=\"paragraph\" w:styleId=\"%s\" w:customStyle=\"true\">" name)
    (format "<w:name w:val=\"%s\"/>" name)
    "<w:basedOn w:val=\"a\"/>"
    (when font (format "<w:link w:val=\"character-style-for-paragraph-style-%s\"/>" name))
    "<w:qFormat/>"
    ; "<w:rsid w:val=\"%s\"/>"
    (when (or align
              space-before space-after space-line
              indent-left indent-right indent-first-line mirror-indents?)
      (ppr options))
    (when (or font font-size font-color bold? italic? underline?)
      (rpr options))
    "</w:style>"
    (when font
      (str
        (format "<w:style w:type=\"character\" w:styleId=\"character-style-for-paragraph-style-%s\" w:customStyle=\"true\">" name)
        (format "<w:name w:val=\"character-style-for-paragraph-style-%s\"/>" name)
        "<w:basedOn w:val=\"a0\"/>"
        (format "<w:link w:val=\"%s\"/>" name)
        (rpr options)
        "</w:style>"))))

(defn word-styles-xml
  ""
  [& custom-styles]
  (str
    ;; header
    "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:ns6=\"http://schemas.openxmlformats.org/schemaLibrary/2006/main\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:ns8=\"http://schemas.openxmlformats.org/drawingml/2006/chartDrawing\" xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:ns11=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\" xmlns:ns13=\"urn:schemas-microsoft-com:office:excel\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:ns17=\"urn:schemas-microsoft-com:office:powerpoint\" xmlns:odx=\"http://opendope.org/xpaths\" xmlns:odc=\"http://opendope.org/conditions\" xmlns:odq=\"http://opendope.org/questions\" xmlns:odi=\"http://opendope.org/components\" xmlns:odgm=\"http://opendope.org/SmartArt/DataHierarchy\" xmlns:ns24=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\" xmlns:ns25=\"http://schemas.openxmlformats.org/drawingml/2006/compatibility\" xmlns:ns26=\"http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas\">"
    ;; <w:docDefaults>
    ;; Specifies the default set of properties which are inherited by every paragraph
    ;; and run within the document. These will define the formatting unless they are
    ;; overridden by other styles or by direct formatting. Note that if the docDefaults
    ;; is not specified, then the application can define defaults.
    ;; (description from -- http://officeopenxml.com/WPstyles.php)
    
    ;; <w:latentStyles>
    ;; Latent styles refer to style definitions known to an application which have not
    ;; been included in the current document. The latentStyles element provides a mechanism
    ;; for storing information regarding certain behaviors of such styles without storing
    ;; the actual formatting properties of the styles. Such behaviors include such things
    ;; as how many latent styles must be initialized to their defaults when the document is
    ;; opened, whether latent styles should be locked so that instances of the styles cannot
    ;; be created, what the uiPriority should be for latent styles, etc.
    ;; (description from -- http://officeopenxml.com/WPstyles.php)

    ;; base styles
    "<w:style w:type=\"paragraph\" w:styleId=\"a\" w:default=\"true\">"
    "<w:name w:val=\"Normal\"/>"
    "<w:qFormat/>"
    "<w:rsid w:val=\"009914CB\"/>"
    "<w:pPr>"
    "<w:widowControl w:val=\"false\"/>"
    "<w:wordWrap w:val=\"false\"/>"
    "<w:autoSpaceDE w:val=\"false\"/>"
    "<w:autoSpaceDN w:val=\"false\"/>"
    "<w:jc w:val=\"both\"/>"
    "</w:pPr>"
    "</w:style>"
    "<w:style w:type=\"character\" w:styleId=\"a0\" w:default=\"true\">"
    "<w:name w:val=\"Default Paragraph Font\"/>"
    "<w:uiPriority w:val=\"1\"/>"
    "<w:semiHidden/>"
    "<w:unhideWhenUsed/>"
    "</w:style>"
    "<w:style w:type=\"table\" w:styleId=\"a1\" w:default=\"true\">"
    "<w:name w:val=\"Normal Table\"/>"
    "<w:uiPriority w:val=\"99\"/>"
    "<w:semiHidden/>"
    "<w:unhideWhenUsed/>"
    "<w:qFormat/>"
    "<w:tblPr>"
    "<w:tblInd w:w=\"0\" w:type=\"dxa\"/>"
    "<w:tblCellMar>"
    "<w:top w:w=\"0\" w:type=\"dxa\"/>"
    "<w:left w:w=\"108\" w:type=\"dxa\"/>"
    "<w:bottom w:w=\"0\" w:type=\"dxa\"/>"
    "<w:right w:w=\"108\" w:type=\"dxa\"/>"
    "</w:tblCellMar>"
    "</w:tblPr>"
    "</w:style>"
    "<w:style w:type=\"numbering\" w:styleId=\"a2\" w:default=\"true\">"
    "<w:name w:val=\"No List\"/>"
    "<w:uiPriority w:val=\"99\"/>"
    "<w:semiHidden/>"
    "<w:unhideWhenUsed/>"
    "</w:style>"

    ;; custom styles
    (apply str custom-styles)

    ;; footer
    "</w:styles>"))

(defn word-media-resources
  ""
  []
  nil)

(defn word-rels-document-xml-rels
  ""
  [resources]
  nil)


