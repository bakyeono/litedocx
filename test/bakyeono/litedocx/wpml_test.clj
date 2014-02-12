(ns bakyeono.litedocx.wpml-test
  (:require [clojure.test :refer :all]
            [bakyeono.litedocx.wpml :refer :all]))

(deftest when-v-tag-test
  (is (= (when-v-tag nil "Tag1") nil))
  (is (= (when-v-tag false "Boolean") "<Boolean>false</Boolean>"))
  (is (= (when-v-tag "" "EmptyTag") "<EmptyTag/>"))
  (is (= (when-v-tag 911 "Number") "<Number>911</Number>"))
  (is (= (when-v-tag "100" "Value") "<Value>100</Value>")))

(deftest content-types-xml-test
  nil)

(deftest doc-props-app-xml-test
  (is (= (doc-props-app-xml)
         "<properties:Properties xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\" xmlns:properties=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"></properties:Properties>")) 
  (is (= (doc-props-app-xml
           :application "Microsoft Office Word"
           :company ""
           :app-version "12.0000")
         "<properties:Properties xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\" xmlns:properties=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>Microsoft Office Word</Application><Company/><AppVersion>12.0000</AppVersion></properties:Properties>")))

(deftest doc-props-core-xml-test
  (is (= (doc-props-core-xml)
         "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"></cp:coreProperties>"))
  (is (= (doc-props-core-xml :title "MongoDB in Action"
                                     :subject ""
                                     :creator "Kyle Banker"
                                     :keywords ""
                                     :description ""
                                     :last-modified-by "Kyle Banker"))
      "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>MongoDB in Action</dc:title><dc:subject/><dc:creator>Kyle Banker</dc:creator><cp:keywords/><dc:description/><cp:lastModifiedBy>Kyle Banker</cp:lastModifiedBy></cp:coreProperties>"))

(deftest word-document-xml-test
  nil)

(deftest word-styles-xml-test
  nil)

(deftest word-media-resources-test
  nil)

(deftest word-rels-document-xml-rels-test
  nil)


(deftest paragraph-style-test
  (is (paragraph-style "PARAGRAPH1"
                       :font "Arial"
                       :font-size 12
                       :font-color "0022CC")
      (str "<w:style w:type=\"paragraph\" w:styleId=\"PARAGRAPH1\" w:customStyle=\"true\">"
           "<w:name w:val=\"PARAGRAPH1\"/>"
           "<w:basedOn w:val=\"a\"/>"
           "<w:link w:val=\"character-style-for-paragraph-style-PARAGRAPH1\"/>"
           "<w:qFormat/>"
           "<w:rPr>"
           "<w:rFonts w:ascii=\"NanumGothic\" w:hAnsi=\"NanumGothic\" w:eastAsia=\"NanumGothic\"/>"
           "<w:sz w:val=\"12\"/>"
           "<w:color w:val=\"0022CC\"/>"
           "</w:rPr>"
           "</w:style>")))


