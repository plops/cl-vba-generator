(asdf:defsystem cl-vba-generator
    :version "0"
    :description "Emit Visual Basic for Applications code"
    :maintainer " <kielhorn.martin@gmail.com>"
    :author " <kielhorn.martin@gmail.com>"
    :licence "GPL"
    :depends-on ("alexandria")
    :serial t
    :components ((:file "package")
		 (:file "vba")) )
