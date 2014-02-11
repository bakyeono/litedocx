(ns bakyeono.litedocx
  "litedocx is a light-weight DOCX format writer library for Clojure."
  (:use [bakyeono.litedocx.util])
  (:use [bakyeono.litedocx.wpml]))

(defrecord Resource [id type content-type part-name])

;;; DOCX package

(defn create-docx-package
  "Creates DOCX package."
  [dst]
  (let [content-types-xml (snippet-content-types-xml [])
        doc-props-app-xml (snippet-doc-props-app-xml)
        doc-props-core-xml (snippet-doc-props-core-xml)
        word-document-xml (snippet-word-document-xml)
        word-styles-xml (snippet-word-styles-xml [])
        word-media-resources nil
        word-rels-document-xml-rels (snippet-word-rels-document-xml-rels [])]
    (write-zip dst
               "[Content_Types].xml" content-types-xml
               "docProps/app.xml" doc-props-app-xml
               "docProps/core.xml" doc-props-core-xml
               "word/document.xml" word-document-xml
               "word/styles.xml" word-styles-xml
               "word/_rels/document.xml.rels" word-rels-document-xml-rels 
               )))

