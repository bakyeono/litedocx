(ns bakyeono.litedocx
  "litedocx is a light-weight DOCX format writer library for Clojure."
  (:require [clojure.data.xml :as xml])
  (:require [bakyeono.litedocx.wpml :as w]) 
  (:use [bakyeono.litedocx.util]))

(defrecord Resource [id type content-type part-name])

;;; DOCX package
(defn check-styles
  [styles document]
  )

(defn pack-docx
  "Creates DOCX package.
  Resources are created by w/make-resources."
  [dst styles resources]
  (let [content-types-xml (w/content-types-xml resources)
        rels-rels (w/rels-rels)
        doc-props-app-xml (w/doc-props-app-xml :application "litedocx" :app-version 1.0)
        doc-props-core-xml (w/doc-props-core-xml)
        word-document-xml (w/word-document-xml
                            (w/table "a1" :center [964 964 8072]
                                     (w/tr (w/td 964 (w/p "a" "name"))
                                           (w/td 964 (w/p "a" "age"))
                                           (w/td 8072 (w/p "a" "description"))))
                            (w/p "para1"
                                 "This is a paragraph."))
        word-styles-xml (w/word-styles-xml styles)
        word-rels-document-xml-rels (w/word-rels-document-xml-rels resources)
        word-media-resources (w/word-media-resources resources)
        entries ["[Content_Types].xml" (emit-xml-as-str content-types-xml)
                 "_rels/.rels" (emit-xml-as-str rels-rels)
                 "docProps/app.xml" (emit-xml-as-str doc-props-app-xml)
                 "docProps/core.xml" (emit-xml-as-str doc-props-core-xml)
                 "word/document.xml" (emit-xml-as-str word-document-xml)
                 "word/styles.xml" (emit-xml-as-str word-styles-xml)
                 "word/_rels/document.xml.rels" (emit-xml-as-str word-rels-document-xml-rels)]]
    (apply write-zip dst (into entries word-media-resources))))

(defn sample-pack
  []
  (pack-docx
    "SAMPLE.docx"
    [(w/paragraph-style "para1" :font "NanumGothic" :font-color "0000AA" :align "both")]
    (w/make-resources
      ["image/png" "foo.png" (load-byte-array "foo.png")])))

(defn sample-pack2
  []
  (pack-docx "SAMPLE2.docx" [] []))

(defn sample-pack3
  []
  (pack-docx
    "SAMPLE3.docx"
    [(w/paragraph-style "para1" :font "NanumGothic" :font-color "0000AA" :align "both")]
    (w/make-resources)))
