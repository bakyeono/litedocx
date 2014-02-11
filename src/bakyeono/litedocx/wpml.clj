(ns bakyeono.litedocx.wpml
  "WordProcessingML snippets used in litedocx."
  (:require [clojure.string :as str])
  (:use [bakyeono.litedocx.util]))

(defn tag-when-v
  "Returns tag and value when the value exists."
  [v tag]
  (cond (nil? v) nil
        (= "" v) (format "<%s/>" tag)
        true (format "<%s>%s</%s>" tag v tag)))

(defn snippet-content-types-xml
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

(defn snippet-doc-props-app-xml
  "Creates a xml code of docProps/app.xml file for DOCX pakcage.\n
  Parameters:
  - & properties: option, value, ...\n
  Examples:
  - (snippet-doc-props-app-xml :application \"Microsoft Office Word\" :app-version \"12.0000\")"
  [& {:keys [application company app-version]}]
  (str
    ;; header
    "<properties:Properties xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\" xmlns:properties=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">"
    ;; properties
    (mapstr tag-when-v
            [application company app-version]
            ["Application" "Company" "AppVersion"])
    ;; footer
    "</properties:Properties>"))

(defn snippet-doc-props-core-xml
  "Creates a xml code of docProps/core.xml file for DOCX pakcage.\n
  Parameters:
  - & properties: option, value, ...\n
  Examples:
  - (snippet-doc-props-core-xml :creator \"Bak Yeon O\" :last-modified-by \"Bak Yeon O\")"
  [& {:keys [title subject creator keywords description last-modified-by]
      :as properties}]
  (str
    ;; header
    "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\">"
    ;; properties
    (mapstr tag-when-v
            [title subject creator keywords description last-modified-by]
            ["dc:title" "dc:subject" "dc:creator" "cp:keywords" "dc:description" "cp:lastModifiedBy"])
    ;; footer
    "</cp:coreProperties>"))

(defn snippet-word-document-xml
  ""
  []
  nil)

(defn snippet-word-styles-xml
  ""
  [styles]
  (str
    ;; header
    "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:ns6=\"http://schemas.openxmlformats.org/schemaLibrary/2006/main\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:ns8=\"http://schemas.openxmlformats.org/drawingml/2006/chartDrawing\" xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:ns11=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\" xmlns:ns13=\"urn:schemas-microsoft-com:office:excel\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:ns17=\"urn:schemas-microsoft-com:office:powerpoint\" xmlns:odx=\"http://opendope.org/xpaths\" xmlns:odc=\"http://opendope.org/conditions\" xmlns:odq=\"http://opendope.org/questions\" xmlns:odi=\"http://opendope.org/components\" xmlns:odgm=\"http://opendope.org/SmartArt/DataHierarchy\" xmlns:ns24=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\" xmlns:ns25=\"http://schemas.openxmlformats.org/drawingml/2006/compatibility\" xmlns:ns26=\"http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas\">"
    ;; <w:docDefaults>
    ;; Specifies the default set of properties which are inherited by every paragraph
    ;; and run within the document. These will define the formatting unless they are
    ;; overridden by other styles or by direct formatting. Note that if the docDefaults
    ;; is not specified, then the application can define defaults.
    ;; (description from -- http://officeopenxml.com/WPstyles.php)
    "<w:docDefaults>"
    "<w:rPrDefault>"
    "<w:rPr>"
    "<w:rFonts w:asciiTheme=\"minorHAnsi\" w:hAnsiTheme=\"minorHAnsi\" w:eastAsiaTheme=\"minorEastAsia\" w:cstheme=\"minorBidi\"/>"
    "<w:kern w:val=\"2\"/>"
    "<w:szCs w:val=\"22\"/>"
    "<w:lang w:val=\"en-US\" w:eastAsia=\"ko-KR\" w:bidi=\"ar-SA\"/>"
    "</w:rPr>"
    "</w:rPrDefault>"
    "<w:pPrDefault/>"
    "</w:docDefaults>"
    ;; <w:latentStyles>
    ;; Latent styles refer to style definitions known to an application which have not
    ;; been included in the current document. The latentStyles element provides a mechanism
    ;; for storing information regarding certain behaviors of such styles without storing
    ;; the actual formatting properties of the styles. Such behaviors include such things
    ;; as how many latent styles must be initialized to their defaults when the document is
    ;; opened, whether latent styles should be locked so that instances of the styles cannot
    ;; be created, what the uiPriority should be for latent styles, etc.
    ;; (description from -- http://officeopenxml.com/WPstyles.php)

    ;; styles
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
    ;; footer
    "</w:styles>"))

(defn snippet-word-media-resources
  ""
  []
  nil)

(defn snippet-word-rels-document-xml-rels
  ""
  [resources]
  nil)


