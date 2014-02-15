(ns bakyeono.litedocx.wpml-test
  (:require [clojure.test :refer :all]
            [bakyeono.litedocx.wpml :refer :all]))

(deftest when-v-tag-test
  (is (= (when-v-tag nil :Tag1)
         nil))
  (is (= (when-v-tag false :Boolean)
         {:tag :Boolean
          :content ["false"]}))
  (is (= (when-v-tag "" :emptyTag)
         {:tag :emptyTag}))
  (is (= (when-v-tag 32767 :number)
         {:tag :number
          :content ["32767"]}))
  (is (= (when-v-tag "100" :Value)
         {:tag :Value
          :content ["100"]})))

(deftest when-v-kv-test
  (is (= (when-v-kv nil :name)
         nil))
  (is (= (when-v-kv "Bak Yeon O" :name)
         [:name "Bak Yeon O"]))
  (is (= (when-v-kv 28 :age)
         [:age 28]))
  (is (= (when-v-kv false :dead?)
         [:dead? false])))

(deftest make-attrs-test
  (is (= (make-attrs :w:before nil
                     :w:after "20"
                     :w:line "30"
                     :w:lineRule "auto")
         {:w:after "20"
          :w:line"30"
          :w:lineRule "auto"})))

(deftest content-types-xml-test
  nil)

(deftest doc-props-app-xml-test
  (is (= (doc-props-app-xml)
         {:tag :properties:Properties
          :attrs {:xmlns:vt "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                  :xmlns:properties "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"}
          :content '()}))
  (is (= (doc-props-app-xml
           :application "Microsoft Office Word"
           :app-version "12.0000")
         {:tag :properties:Properties
          :attrs {:xmlns:vt "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                  :xmlns:properties "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"} 
          :content
          [{:tag :properties:Application
            :content ["Microsoft Office Word"]}
           {:tag :properties:AppVersion
            :content ["12.0000"]}]})))

(deftest doc-props-core-xml-test
  (is (= (doc-props-core-xml)
         {:tag :cp:coreProperties
          :attrs 
          {:xmlns:cp "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
           :xmlns:dcterms "http://purl.org/dc/terms/"
           :xmlns:dc "http://purl.org/dc/elements/1.1/"} 
          :content '()}))
  (is (= (doc-props-core-xml :title "MongoDB in Action"
                             :subject ""
                             :creator "Kyle Banker"
                             :keywords ""
                             :description ""
                             :last-modified-by "Kyle Banker")
      {:tag :cp:coreProperties
       :attrs 
       {:xmlns:cp "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
        :xmlns:dcterms "http://purl.org/dc/terms/"
        :xmlns:dc "http://purl.org/dc/elements/1.1/"} 
       :content
       [{:tag :dc:title :content ["MongoDB in Action"]}
        {:tag :dc:subject}
        {:tag :dc:creator :content ["Kyle Banker"]}
        {:tag :cp:keywords}
        {:tag :dc:description}
        {:tag :cp:lastModifiedBy :content ["Kyle Banker"]}]})))

(deftest word-document-xml-test
  nil)

(deftest word-styles-xml-test
  nil)

(deftest word-media-resources-test
  nil)

(deftest word-rels-document-xml-rels-test
  nil)


