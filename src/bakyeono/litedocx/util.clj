(ns bakyeono.litedocx.util
  "Custom utilities used in litedocx."
  (:import [java.io OutputStream])
  (:import [java.util.zip ZipEntry ZipOutputStream])
  (:require [clojure.string :as str])
  (:require [clojure.zip :as z])
  (:require [clojure.java.io :as io]))

;;; Collection
(defn mapstr
  "Returns the result of applying str to the result of applying map
  to f and colls.  Thus function f should return a collection."
  [f & colls]
  (apply str (apply map f colls)))

;;; String
(defn to-camel-case
  [str]
  (str/replace str #"-(\w)" #(str/upper-case (second %))))

;;; Java
(def array-of-bytes-type (Class/forName "[B")) 
(defn is-byte-array?
  [object]
  (= (class object) array-of-bytes-type))

;;; I/O
(defn- write-zip-entry-body-as-text
  "Writes a text entry on zip-ostream.
  Called by write-zip-entry."
  [zip-ostream entry-body]
  (let [writer (io/writer zip-ostream)]
    (.write writer #^String entry-body)
    (.flush writer)))

(defn- write-zip-entry-body-as-bytes
  "Writes a binary entry on zip-ostream.
  Called by write-zip-entry."
  [zip-ostream entry-body]
  (.write zip-ostream #^array-of-bytes-type entry-body)
  (.flush zip-ostream))

(defn- write-zip-entry
  "Writes an entry on zip-ostream with given filename & body of entry.
  Called by write-zip."
  [zip-ostream [entry-filename entry-body]]
  (.putNextEntry zip-ostream (ZipEntry. #^String entry-filename))
  (cond (string? entry-body)
        (write-zip-entry-body-as-text zip-ostream entry-body)
        (is-byte-array? entry-body)
        (write-zip-entry-body-as-bytes zip-ostream entry-body))
  (.closeEntry zip-ostream))

(defn write-zip
  "Writes a zip file into the 'dst' path, with given 'entries'.
  Parameters:
  - dst: <string> filename of the file to be written
  - & entries: <string> filename of entry, <string or byte array> body of entry, ...
  you can add as many entries as needed as '& entries' argument.
  Examples:
  (write-zip \"foo.zip\" \"bar.txt\" \"Lorem ipsum\" \"baz.bin\" (byte-array 10))"
  [dst & entries]
  (with-open [ostream (io/output-stream dst)
              zip-ostream (ZipOutputStream. ostream)]
    (doseq [entry (partition 2 entries)]
      (write-zip-entry zip-ostream entry))))

(defn compress-files-as-zip
  "Compress files as zip format."
  [target]
  (println "Hello, World!"))


