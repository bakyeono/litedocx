(ns bakyeono.litedocx
  "litedocx is a light-weight DOCX format writer library for Clojure."
  (:require [clojure.xml :as xml])
  (:require [bakyeono.litedocx.wpml :as w]) 
  (:use [bakyeono.litedocx.util]))

(defrecord Resource [id type content-type part-name])

;;; DOCX package
(defn pack-docx
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
               "[Content_Types].xml" (xml/emit content-types-xml)
               "docProps/app.xml" (xml/emit doc-props-app-xml)
               "docProps/core.xml" (xml/emit doc-props-core-xml)
               "word/document.xml" (xml/emit word-document-xml)
               "word/styles.xml" (xml/emit word-styles-xml)
               "word/_rels/document.xml.rels" (xml/emit word-rels-document-xml-rels))))

