(ns bakyeono.litedocx.wpml-test
  (:require [clojure.test :refer :all]
            [bakyeono.litedocx.wpml :refer :all]))

(deftest tag-when-v-test
  (is (= (tag-when-v nil "Tag1") nil))
  (is (= (tag-when-v false "Boolean") "<Boolean>false</Boolean>"))
  (is (= (tag-when-v "" "EmptyTag") "<EmptyTag/>"))
  (is (= (tag-when-v 911 "Number") "<Number>911</Number>"))
  (is (= (tag-when-v "100" "Value") "<Value>100</Value>")))

(deftest snippet-content-types-xml-test
  nil)

(deftest snippet-doc-props-app-xml-test
  (is (= (snippet-doc-props-app-xml)
         "<properties:Properties xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\" xmlns:properties=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"></properties:Properties>")) 
  (is (= (snippet-doc-props-app-xml
           :application "Microsoft Office Word"
           :company ""
           :app-version "12.0000")
         "<properties:Properties xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\" xmlns:properties=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>Microsoft Office Word</Application><Company/><AppVersion>12.0000</AppVersion></properties:Properties>")))

(deftest snippet-doc-props-core-xml-test
  (is (= (snippet-doc-props-core-xml)
         "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"></cp:coreProperties>"))
  (is (= (snippet-doc-props-core-xml :title "MongoDB in Action"
                                     :subject ""
                                     :creator "Kyle Banker"
                                     :keywords ""
                                     :description ""
                                     :last-modified-by "Kyle Banker"))
      "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>MongoDB in Action</dc:title><dc:subject/><dc:creator>Kyle Banker</dc:creator><cp:keywords/><dc:description/><cp:lastModifiedBy>Kyle Banker</cp:lastModifiedBy></cp:coreProperties>"))

(deftest snippet-word-document-xml-test
  nil)

(deftest snippet-word-styles-xml-test
  nil)

(deftest snippet-word-media-resources-test
  nil)

(deftest snippet-word-rels-document-xml-rels-test
  nil)



