(ns bakyeono.litedocx.unit
  "Thinchgs for unit conversion."
  (:require [clojure.string :as str])
  (:use [bakyeono.litedocx.util])
  (:gen-class))

;;; Length units
;;; Base conversion
(defconst cm-per-inch 2.54)
(defconst inch-per-emu 1/914400)
(defconst mm-per-cm 10)
(defconst pt-per-px 3/4)
(defconst pt-per-twip 1/20)
(defconst px-per-inch 96)

;;; Reversion
(defconst inch-per-cm (/ 1 cm-per-inch))
(defconst emu-per-inch (/ inch-per-emu))
(defconst cm-per-mm (/ 1 mm-per-cm))
(defconst px-per-pt (/ 1 pt-per-px))
(defconst twip-per-pt (/ pt-per-twip))
(defconst inch-per-px (/ 1 px-per-inch))

;;; Combination
(defconst cm-per-emu (* cm-per-inch inch-per-emu))
(defconst cm-per-px (* cm-per-inch inch-per-px))
(defconst cm-per-pt (* cm-per-px px-per-pt))
(defconst cm-per-twip (* cm-per-pt pt-per-twip))
(defconst inch-per-mm (* inch-per-cm cm-per-mm))
(defconst inch-per-pt (* inch-per-px px-per-pt))
(defconst inch-per-twip (* inch-per-pt pt-per-twip))
(defconst mm-per-inch (* mm-per-cm cm-per-inch))
(defconst mm-per-emu (* mm-per-inch inch-per-emu))
(defconst mm-per-px (* mm-per-inch inch-per-px))
(defconst mm-per-pt (* mm-per-px px-per-pt))
(defconst mm-per-twip (* mm-per-pt pt-per-twip))
(defconst pt-per-inch (* pt-per-px px-per-inch))
(defconst pt-per-emu (* pt-per-inch inch-per-emu))
(defconst pt-per-cm (* pt-per-inch inch-per-cm))
(defconst pt-per-mm (* pt-per-cm cm-per-mm))
(defconst px-per-emu (* px-per-inch inch-per-emu))
(defconst px-per-cm (* px-per-inch inch-per-cm))
(defconst px-per-mm (* px-per-cm cm-per-mm))
(defconst px-per-twip (* px-per-pt pt-per-twip))
(defconst twip-per-px (* twip-per-pt pt-per-px))
(defconst twip-per-inch (* twip-per-px px-per-inch))
(defconst twip-per-emu (* twip-per-inch inch-per-emu))
(defconst twip-per-cm (* twip-per-inch inch-per-cm))
(defconst twip-per-mm (* twip-per-cm cm-per-mm))
(defconst emu-per-cm (* emu-per-inch inch-per-cm))
(defconst emu-per-mm (* emu-per-cm cm-per-mm))
(defconst emu-per-px (* emu-per-inch inch-per-px))
(defconst emu-per-pt (* emu-per-px px-per-pt))
(defconst emu-per-twip (* emu-per-pt pt-per-twip))

;;; Functions
(defn parse-unit
  "Takes a string expression of a number and returns it as a [number unit]
  vector."
  [s]
  (let [s (str/trim s)
        n (Double/parseDouble (re-find #"^-?\d+\.?\d*" s))
        unit (re-find #"[a-z]+$" s)]
    [n unit]))

(defn conversion-rate
  "Returns conversion rate for 'from' unit -> 'to' unit."
  [from to]
  (let [sym (symbol (str "bakyeono.litedocx.unit/" to "-per-" from))]
    (cond (resolve sym) (eval sym)
          (= from to) 1
          true (throw (Exception. (str "Unsupported unit conversion: "
                                       from " -> " to))))))

(defn convert
  "Takes a string expression of a value and its unit and then converts it into
  given to-unit.

  Parameters:
  - value: <string> value with unit.
  - to: <string or keyword> target unit type.

  Examples:
  - (convert \"10px\" \"mm\")
  - (convert \"29.7 cm\" :twip)

  Supported Unit Types:
  - Length Units: inch, cm, mm, px, pt, twip"
  [value to]
  (let [[n unit] (parse-unit value)
        rate (conversion-rate unit (name to))]
    (* n rate)))
