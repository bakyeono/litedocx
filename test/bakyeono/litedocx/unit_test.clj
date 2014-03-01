(ns bakyeono.litedocx.unit-test
  (:require [clojure.test :refer :all]
            [bakyeono.litedocx.unit :refer :all]))

(deftest convert-test
  (is (== (convert "1m" "m") 1))
  (is (== (convert "1 cm" "cm") 1))
  (is (== (convert "1.0mm" "mm") 1))
  (is (== (convert "1inch" "inch") 1))
  (is (== (convert "1 in." "inch") 1))
  (is (== (convert "1 inch" "cm") 2.54))
  (is (== (convert "72 pt" "inch") 1))
  (is (== (convert "1 yd" "inch") 1/36)))


