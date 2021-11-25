(in-package :cl-vba-generator)
(setf (readtable-case *readtable*) :invert)

(defparameter *file-hashes* (make-hash-table))

(defun write-source (name code &optional (dir (user-homedir-pathname))
				 ignore-hash)
  (let* ((fn (merge-pathnames (format nil "~a.vba" name)
			      dir))
	(code-str (emit-vba
		   :clear-env t
		   :code code))
	(fn-hash (sxhash fn))
	 (code-hash (sxhash code-str)))
    (multiple-value-bind (old-code-hash exists) (gethash fn-hash *file-hashes*)
     (when (or (not exists) ignore-hash (/= code-hash old-code-hash))
       ;; store the sxhash of the c source in the hash table
       ;; *file-hashes* with the key formed by the sxhash of the full
       ;; pathname
       (setf (gethash fn-hash *file-hashes*) code-hash)
       (with-open-file (s fn
			  :direction :output
			  :if-exists :supersede
			  :if-does-not-exist :create)
	 (write-sequence code-str s))
       ;; FIXME: find a programm to format VBA scripts
       #+nil (sb-ext:run-program "/usr/bin/yapf" (list "-i" (namestring fn)))))))

(defun print-sufficient-digits-f64 (f)
  "print a double floating point number as a string with a given nr. of                                                                                                                                             
  digits. parse it again and increase nr. of digits until the same bit                                                                                                                                              
  pattern."

  (let* ((a f)
         (digits 1)
         (b (- a 1)))
    (unless (= a 0)
      (loop while (< 1d-12
		     (/ (abs (- a b))
		       (abs a))
		    ) do
          (setf b (read-from-string (format nil "~,vG" digits a)))
           (incf digits)
	   ))
    (substitute #\e #\d (format nil "~,vG" digits a))))

(defparameter *env-functions* nil)
(defparameter *env-macros* nil)

(defun emit-vba (&key code (str nil) (clear-env nil) (level 0))
  ;(format t "emit ~a ~a~%" level code)
  (when clear-env
    (setf *env-functions* nil
	  *env-macros* nil))
  (flet ((emit (code &optional (dl 0))
	   (emit-vba :code code :clear-env nil :level (+ dl level))))
    (format nil "emit-py ~a" level)
    (if code
	(if (listp code)
	    (case (car code)
	      (tuple (let ((args (cdr code)))
		       (format nil "(~{~a,~})" (mapcar #'emit args))))
	      (paren (let ((args (cdr code)))
		       (format nil "(~{~a~^, ~})" (mapcar #'emit args))))
	      (ntuple (let ((args (cdr code)))
		       (format nil "~{~a~^, ~}" (mapcar #'emit args))))
	      (list (let ((args (cdr code)))
		      (format nil "[~{~a~^, ~}]" (mapcar #'emit args))))
	      (curly (let ((args (cdr code)))
		      (format nil "{~{~a~^, ~}}" (mapcar #'emit args))))
              (indent (format nil "~{~a~}~a"
			      (loop for i below level collect "    ")
			      (emit (cadr code))))
	      (do (with-output-to-string (s)
		    (format s "~{~&~a~}" (mapcar #'(lambda (x) (emit `(indent ,x) 1)) (cdr code)))))
	      (do0 (with-output-to-string (s)
		     (format s "~&~a~{~&~a~}"
			     (emit (cadr code))
			     (mapcar #'(lambda (x) (emit `(indent ,x) 0)) (cddr code)))))
	      (space (with-output-to-string (s)
		     (format s "~{~a~^ ~}"
			     (mapcar #'(lambda (x) (emit x)) (cdr code)))))
	      (sub (destructuring-bind (name lambda-list &rest body) (cdr code)
		     (multiple-value-bind (req-param opt-param res-param
						     key-param other-key-p aux-param key-exist-p)
			 (parse-ordinary-lambda-list lambda-list)
		       (declare (ignorable req-param opt-param res-param
					   key-param other-key-p aux-param key-exist-p))
		       (with-output-to-string (s)
			 (format s "Sub ~a~a~%"
				 name
				 (emit `(paren
					 ,@(append (mapcar #'emit req-param)
						   (loop for e in key-param collect 
							(destructuring-bind ((keyword-name name) init suppliedp)
							    e
							  (declare (ignorable keyword-name suppliedp))
							  (if init
							      `(= ,name ,init)
							      `(= ,name "None"))))))))
			 (format s "~a" (emit `(do ,@body)))
			 (format s "~&End Sub~%")))))
	      (= (destructuring-bind (a b) (cdr code)
		   (format nil "~a=~a" (emit a) (emit b))))
	      (|:=| (destructuring-bind (a b) (cdr code)
		   (format nil "~a:=~a" (emit a) (emit b))))
	      (setf (let ((args (cdr code)))
		      (format nil "~a"
			      (emit `(do0 
				      ,@(loop for i below (length args) by 2 collect
					     (let ((a (elt args i))
						   (b (elt args (+ 1 i))))
					       `(= ,a ,b))))))))
	      (incf (destructuring-bind (target &optional (val 1)) (cdr code)
		      (format nil "~a += ~a" (emit target) (emit val))))
	      (decf (destructuring-bind (target &optional (val 1)) (cdr code)
		      (format nil "~a -= ~a"
			      (emit target)
			      (emit val))))
	      (aref (destructuring-bind (name &rest indices) (cdr code)
		      (format nil "~a[~{~a~^,~}]" (emit name) (mapcar #'emit indices))))
	      (dot (let ((args (cdr code)))
		   (format nil "~{~a~^.~}" (mapcar #'emit args))))
	      (+ (let ((args (cdr code)))
		   (format nil "(~{(~a)~^+~})" (mapcar #'emit args))))
	      (- (let ((args (cdr code)))
		   (format nil "(~{(~a)~^-~})" (mapcar #'emit args))))
	      (* (let ((args (cdr code)))
		   (format nil "(~{(~a)~^*~})" (mapcar #'emit args))))
	      (== (let ((args (cdr code)))
		    (format nil "(~{(~a)~^==~})" (mapcar #'emit args))))
	      (<< (let ((args (cdr code)))
		    (format nil "(~{(~a)~^<<~})" (mapcar #'emit args))))
	      (!= (let ((args (cdr code)))
		   (format nil "(~{(~a)~^!=~})" (mapcar #'emit args))))
	      (< (let ((args (cdr code)))
		   (format nil "(~{(~a)~^<~})" (mapcar #'emit args))))
	      (<= (let ((args (cdr code)))
		    (format nil "(~{(~a)~^<=~})" (mapcar #'emit args))))
	      (>> (let ((args (cdr code)))
		   (format nil "(~{(~a)~^>>~})" (mapcar #'emit args))))
	      (/ (let ((args (cdr code)))
		   (format nil "((~a)/(~a))"
			   (emit (first args))
			   (emit (second args)))))
	      (** (let ((args (cdr code)))
		   (format nil "((~a)**(~a))"
			   (emit (first args))
			   (emit (second args)))))
	      (// (let ((args (cdr code)))
		   (format nil "((~a)//(~a))"
			   (emit (first args))
			   (emit (second args)))))
	      (% (let ((args (cdr code)))
		   (format nil "((~a)%(~a))"
			   (emit (first args))
			   (emit (second args)))))
	      (and (let ((args (cdr code)))
		     (format nil "(~{(~a)~^ and ~})" (mapcar #'emit args))))
	      (& (let ((args (cdr code)))
		   (format nil "(~{(~a)~^ & ~})" (mapcar #'emit args))))
	      (logand (let ((args (cdr code)))
			(format nil "(~{(~a)~^ & ~})" (mapcar #'emit args))))
	      (logxor (let ((args (cdr code)))
		   (format nil "(~{(~a)~^ ^ ~})" (mapcar #'emit args))))
	      (|\|| (let ((args (cdr code)))
		      (format nil "(~{(~a)~^ | ~})" (mapcar #'emit args))))
	      (^ (let ((args (cdr code)))
		      (format nil "(~{(~a)~^ ^ ~})" (mapcar #'emit args))))
	      (logior (let ((args (cdr code)))
		     (format nil "(~{(~a)~^ | ~})" (mapcar #'emit args))))
	      (or (let ((args (cdr code)))
		    (format nil "(~{(~a)~^ or ~})" (mapcar #'emit args))))
	      (comment (format nil "' ~a~%" (cadr code)))
	      (comments (let ((args (cdr code)))
			  (format nil "~{' ~a~%~}" args)))
	      (string (format nil "\"~a\"" (cadr code)))
	      (for (destructuring-bind ((vs start end &key step) &rest body) (cdr code)
		     (with-output-to-string (s)
		       ;(format s "~a" (emit '(indent)))
		       (format s "For ~a = ~a To ~a"
			       (emit vs)
			       (emit start)
			       (emit end))
		       (when step
			 (format s " Step ~a" (emit step)))
		       (format s "~%")
		       (format s "~a" (emit `(do ,@body)))
		       (format s "~&Next ~a~%"
			       (emit vs)))))

	      (for-each (destructuring-bind ((vs group) &rest body) (cdr code)
		     (with-output-to-string (s)
		       
		       (format s "For Each ~a in~%"
			       (emit vs)
			       (emit group))
		       (format s "~a" (emit `(do ,@body)))
		       (format s "~&Next~%"))))
	      (while (destructuring-bind (vs &rest body) (cdr code)
		     (with-output-to-string (s)
		       (format s "Do While ~a~%"
			       (emit `(paren ,vs)))
		       (format s "~a" (emit `(do ,@body)))
		       (format s "~&Loop~%"))))

	      (if (destructuring-bind (condition true-statement &optional false-statement) (cdr code)
		    (with-output-to-string (s)
		      (format s "If ( ~a ) Then~%~a"
			      (emit condition)
			      (emit `(do ,true-statement)))
		      (when false-statement
			(format s "~&~a:~%~a"
				(emit `(indent "Else"))
				(emit `(do ,false-statement)))))))
	      (cond (destructuring-bind (&rest clauses) (cdr code)
		      (with-output-to-string (s)
			(loop for clause in clauses and i from 0
			      do
				 (destructuring-bind (condition &rest statements) clause
				   (format s "~&~a~%~a"
					   (cond ((and (eq condition 't) (eq i 0))
						  ;; this special case may happen when you comment out all but the last cond clauses
						  (format nil "If ( True ) Then"))
						 ((eq i 0) (format nil "If ( ~a ) Then" (emit condition)))
						 ((eq condition 't) (emit `(indent "Else")))
						 (t (emit `(indent ,(format nil "ElseIf ( ~a ) Then" (emit condition)))))
						 )
					   (emit `(do ,@statements)))))
		      )))
	      (when (destructuring-bind (condition &rest forms) (cdr code)
                      (emit `(if ,condition
                                 (do0
                                  ,@forms)))))
              (unless (destructuring-bind (condition &rest forms) (cdr code)
                        (emit `(if (not ,condition)
                                   (do0
                                    ,@forms)))))
	      #+nil
	      (with (destructuring-bind (form &rest body) (cdr code)
		      (with-output-to-string (s)
		       (format s "~a~a:~%~a"
			       (emit "with ")
			       (emit form)
			       (emit `(do ,@body))))))
	      
	      (t (destructuring-bind (name &rest args) code
		   
		   (if (listp name)
		       ;; lambda call and similar complex constructs
		       (format nil "(~a)(~a)" (emit name) (if args
							      (emit `(paren ,@args))
							      ""))
		       #+nil(if (eq 'lambda (car name))
			   (format nil "(~a)(~a)" (emit name) (emit `(paren ,@args)))
			   (break "error: unknown call"))
		       ;; function call
		       (let* ((positional (loop for i below (length args) until (keywordp (elt args i)) collect
					       (elt args i)))
			      (plist (subseq args (length positional)))
			      (props (loop for e in plist by #'cddr collect e)))
			 (format nil "~a~a" name
				 (emit `(paren ,@(append
						  positional
						  (loop for e in props collect
						       `(|:=| ,(format nil "~a" e) ,(getf plist e))))))))))))
	    (cond
	      ((symbolp code) ;; print variable
	       (format nil "~a" code))
	      ((stringp code)
		(substitute #\: #\- (format nil "~a" code)))
	      ((numberp code) ;; print constants
	       (cond ((integerp code) (format str "~a" code))
		     ((floatp code)
		      (format str "(~a)" (print-sufficient-digits-f64 code)))
		     ((complexp code)
		      (format str "((~a) + 1j * (~a))"
			      (print-sufficient-digits-f64 (realpart code))
			      (print-sufficient-digits-f64 (imagpart code))))))))
	"")))

