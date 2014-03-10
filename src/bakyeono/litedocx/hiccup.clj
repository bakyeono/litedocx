(ns bakyeono.litedocx.hiccup
  "Support hiccup-like style DSL for creating DOCX."
  (:require [clojure.string :as str])
  (:use [bakyeono.litedocx.util])
  (:gen-class))

(def sample1
  '(docx
    [:styles [:headline {:font "Sans Serif"
                         :font-color "000000"
                         :font-size 20
                         :bold true
                         :align "both"}]
             [:date {:font "Arial"
                     :font-color "4444CC"
                     :font-size 11
                     :italics true
                     :align "right"}]
             [:author {:font "Arial"
                       :font-color "aa0000"
                       :font-size 11
                       :algin "right"}]
             [:monospace {:font "Monospace"
                          :font-color "000000"
                          :align "left"}]
             [:h1 {:font "Arial"
                   :font-color "000000"
                   :font-size 16
                   :bold true}]
             [:body {:font "Arial"
                     :font-color "000000"
                     :font-size 12}]
             [:table {:font "Arial"
                      :font-color "449944"
                      :align "center"}]]
    [:resources [:image1 {:type "image/png"
                          :byte-array (load-byte-array "foo.png")}]]
    [:doc [:p {:style :headline}
              "A document created with hiccup-like style DSL"]
          [:p {:style :date}
              (str (java.util.Date.))]
          [:p {:style :author}
              "Bak Yeon O"]
          [:p {:style :body}
              ""]
          [:p {:style :body}
              "An easy way to make DOCX documents with Clojure."]
          [:p {:style :monospace}
              "int main(int argc, char** argv)"]]))
