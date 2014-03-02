(ns bakyeono.litedocx
  "litedocx is a light-weight DOCX format writer library for Clojure."
  (:require [clojure.data.xml :as xml])
  (:require [bakyeono.litedocx.wpml :as w]) 
  (:use [bakyeono.litedocx.util]))

;;; DOCX package
(defn check-styles
  [styles document]
  )

(defn check-resources
  [styles document]
  )

(defn- pack-docx
  "Creates DOCX package.
  Resources are created by w/resources."
  [dst styles resources]
  (let [content-types-xml (w/content-types-xml resources)
        rels-rels (w/rels-rels)
        doc-props-app-xml (w/doc-props-app-xml :application "litedocx" :app-version 1.0)
        doc-props-core-xml (w/doc-props-core-xml)
        word-document-xml (w/word-document-xml)
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

(defn write-docx
  "Write API for users."
  [dst
   & {:as params
      :keys [styles resources body
             application app-version
             page-width page-height page-orientation
             page-margin-top page-margin-bottom page-margin-left page-margin-right
             page-margin-header page-margin-footer
             title subject creator keywords description last-modified-by]
      :or {:application "" :app-version ""
           :page-width "210mm" :page-height "297mm" :page-orientation "portrait"
           :page-margin-top 720 :page-margin-bottom 720
           :page-margin-left 720 :page-margin-right 720
           :page-margin-header 0 :page-margin-footer 0
           :title "" :subject ""
           :creator "" :last-modified-by ""
           :keywords "" :description ""}}]
  (let [content-types-xml (w/content-types-xml resources)
        rels-rels (w/rels-rels)
        doc-props-app-xml (w/doc-props-app-xml params)
        doc-props-core-xml (w/doc-props-core-xml params)
        word-document-xml (w/word-document-xml params)
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

(defn sample1
  []
  (write-docx :styles [(w/paragraph-style "headline"
                                          :font "Sans Serif"
                                          :font-color "000000"
                                          :font-size 20
                                          :bold true
                                          :align "both")
                       (w/paragraph-style "date"
                                          :font "Arial"
                                          :font-color "4444CC"
                                          :font-size 12
                                          :italics true
                                          :algin "right")
                       (w/paragraph-style "author"
                                          :font "Arial"
                                          :font-color "aa0000"
                                          :font-size 13
                                          :algin "right")
                       (w/paragraph-style "monospace"
                                          :font "Monospace"
                                          :font-color "000000"
                                          :algin "left")
                       (w/paragraph-style "body")
                       (w/table-style "table"
                                      :font "Arial"
                                      :font-color "449944"
                                      :align "center")]
              ;; TODO: need a better way to create resource meta datas
              :resources (w/resources ["image/png" "foo.png" (load-byte-array "foo.png")])
              :document [(w/table "a1" :center [964 964 8072]
                                  (w/tr (w/td 964 (w/p "a" "name"))
                                        (w/td 964 (w/p "a" "age"))
                                        (w/td 8072 (w/p "a" "description"))))
                         (w/p "para1"
                              "This is a paragraph.")
                         (w/img "rId2"
                                "800px"
                                "600px"
                                :name "foo.png"
                                :desc "Image Description: this is a testing image.")]))

(defn sample-pack
  []
  (pack-docx
    "SAMPLE.docx"
    [(w/paragraph-style "para1" :font "NanumGothic" :font-color "0000AA" :align "both")]
    (w/resources
      ["image/png" "foo.png" (load-byte-array "foo.png")])))

(defn sample-pack2
  []
  (pack-docx "SAMPLE2.docx" [] []))

(defn sample-pack3
  []
  (pack-docx
    "SAMPLE3.docx"
    [(w/paragraph-style "para1" :font "NanumGothic" :font-color "0000AA" :align "both")]
    (w/resources)))

