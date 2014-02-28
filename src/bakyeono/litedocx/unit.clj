(ns bakyeono.litedocx.unit
  "Things for unit conversion."
  (:require [clojure.string :as str])
  (:use [bakyeono.litedocx.util])
  (:gen-class))

;;; Conversion rate constants
(defn- reverse-conversion-map
  [m]
  (into {}
    (for [[k v] m]
      [k (/ 1 (k m))])))

(defconst flat-unit-per-m
  {:cm       100
   :km       1/1000
   :m        1
   :mm       1000})

(defconst unit-per-inch
  {:cm       2.54
   :dxa      1440
   :emu      914400
   :feet     12
   :ft       12
   :in.      1
   :inch     1
   :km       0.0000254
   :m        0.0254
   :mm       25.4
   :pt       72
   :px       96
   :twip     1440
   :yard     36
   :yd       36})

(defconst inch-per-unit (reverse-conversion-map unit-per-inch))

(defconst unit-per-m
  (into (into {} (for [k (keys unit-per-inch)]
                   [k (* (inch-per-unit :m)
                         (k unit-per-inch))]))
        flat-unit-per-m))

(defconst m-per-unit (reverse-conversion-map unit-per-m))

;;; Conversion rate map selector
(defconst metric-set (set (keys flat-unit-per-m)))
(defconst inch-set (clojure.set/difference (set (keys unit-per-inch))
                                           metric-set))

;;; Functions
(defn- parse-unit
  "Takes a string expression of a number and returns it as a [number unit]
  vector."
  [s]
  (let [s (str/trim s)
        n (Double/parseDouble (re-find #"^-?\d+\.?\d*" s))
        unit (-> (re-find #"[a-z]+$" s) clojure.string/lower-case keyword)]
    [n unit]))

(defn- conversion-medium
  "Takes a keyword of unit and returns the preffered conversion medium of the
  unit."
  [u]
  (cond (metric-set u) [m-per-unit unit-per-m]
        (inch-set u) [inch-per-unit unit-per-inch]))

(defn- conversion-rate
  "Returns conversion rate for 'from' unit -> 'to' unit. The parameters should
  be keywords."
  [from to]
  (let [medium (conversion-medium from)
        medium-per-from (from (first medium))
        to-per-medium (to (second medium))]
    (if (and medium-per-from to-per-medium)
      (* medium-per-from to-per-medium)
      (throw (Exception. (str "Unsupported unit conversion: " from " -> " to))))))

(defn convert
  "Takes a string expression of a value and its unit and then converts it into
  given to-unit.

  Parameters:
  - value: <string> value with unit.
  - to: <string or keyword> target unit type.

  Supported Unit Types:
  - Length Units: inch, cm, km, m, mm, px, pt, dxa, emu

  Examples:
  - (convert \"10px\" \"mm\")
  - (convert \"29.7 cm\" :dxa)"
  [value to]
  (let [[n unit] (parse-unit value)
        rate (conversion-rate unit (keyword to))]
    (* n rate)))
