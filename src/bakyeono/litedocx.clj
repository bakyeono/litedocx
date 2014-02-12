(ns bakyeono.litedocx
  "litedocx is a light-weight DOCX format writer library for Clojure."
  (:use [bakyeono.litedocx.util])
  (:require [bakyeono.litedocx.wpml :as w]))

(defrecord Resource [id type content-type part-name])

;;; DOCX package

(defn create-docx-package
  "Creates DOCX package."
  [dst]
  (let [content-types-xml (w/content-types-xml [])
        doc-props-app-xml (w/doc-props-app-xml)
        doc-props-core-xml (w/doc-props-core-xml)
        word-document-xml (w/word-document-xml)
        word-styles-xml (w/word-styles-xml [])
        word-media-resources nil
        word-rels-document-xml-rels (w/word-rels-document-xml-rels [])]
    (write-zip dst
               "[Content_Types].xml" content-types-xml
               "docProps/app.xml" doc-props-app-xml
               "docProps/core.xml" doc-props-core-xml
               "word/document.xml" word-document-xml
               "word/styles.xml" word-styles-xml
               "word/_rels/document.xml.rels" word-rels-document-xml-rels 
               )))

