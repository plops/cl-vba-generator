(eval-when (:compile-toplevel :execute :load-toplevel)
  (ql:quickload "alexandria")
  (ql:quickload "cl-vba-generator"))
(in-package :cl-vba-generator)



(progn
  (defparameter *path* "/home/martin/stage/cl-vba-generator/example/01_word_insert_img")
  (defparameter *day-names*
    '("Monday" "Tuesday" "Wednesday"
      "Thursday" "Friday" "Saturday"
      "Sunday"))

  (write-source 
   (format nil "~a/source/01_word_insert_img" *path*)
   `(
     (do0
      (setf
       _code_git_version
       (string ,(let ((str (with-output-to-string (s)
			     (sb-ext:run-program "/usr/bin/git" (list "rev-parse" "HEAD") :output s))))
		  (subseq str 0 (1- (length str)))))
       _code_repository (string ,(format nil "https://github.com/plops/cl-vba-generator/tree/master/example/01_word_insert_img"))
       _code_generation_time
       (string ,(multiple-value-bind
		      (second minute hour date month year day-of-week dst-p tz)
		    (get-decoded-time)
		  (declare (ignorable dst-p))
		  (format nil "~2,'0d:~2,'0d:~2,'0d of ~a, ~d-~2,'0d-~2,'0d (GMT~@d)"
			  hour
			  minute
			  second
			  (nth day-of-week *day-names*)
			  year
			  month
			  date
			  (- tz)))))))))



