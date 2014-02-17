(ns bakyeono.litedocx
  "litedocx is a light-weight DOCX format writer library for Clojure."
  (:require [clojure.xml :as xml])
  (:require [bakyeono.litedocx.wpml :as w]) 
  (:use [bakyeono.litedocx.util]))

(defrecord Resource [id type content-type part-name])

;;; DOCX package
(defn pack-docx
  "Creates DOCX package.
  Resources are created by w/make-resources."
  [dst styles resources]
  (let [content-types-xml (w/content-types-xml resources)
        doc-props-app-xml (w/doc-props-app-xml)
        doc-props-core-xml (w/doc-props-core-xml)
        word-document-xml (w/word-document-xml)
        word-styles-xml (w/word-styles-xml styles)
        word-rels-document-xml-rels (w/word-rels-document-xml-rels resources)
        word-media-resources (w/word-media-resources resources)
        entries ["[Content_Types].xml" (emit-xml-as-str content-types-xml)
                 "docProps/app.xml" (emit-xml-as-str doc-props-app-xml)
                 "docProps/core.xml" (emit-xml-as-str doc-props-core-xml)
                 "word/document.xml" (emit-xml-as-str word-document-xml)
                 "word/styles.xml" (emit-xml-as-str word-styles-xml)
                 "word/_rels/document.xml.rels" (emit-xml-as-str word-rels-document-xml-rels)]]
    (apply write-zip dst (into entries word-media-resources))))

(defn sample-pack
  []
  (pack-docx
    "sample-pack.docx"
    []
    (w/make-resources
      ["image/png" "foo.png" (load-byte-array "foo.png")])))

