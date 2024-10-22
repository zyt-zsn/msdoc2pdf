(require 'cl)
(add-to-list 'auto-mode-alist '("\\.\\(docx?\\|DOCX?\\|pptx?\\|PPTX?\\|xlsx?\\|XLSX?\\)\\'" . zyt/doc-mode))
(defvar doc2pdf-convert-process nil)
(defvar doc2pdf-convert-in-process nil)
;; (defvar doc2pdf-arguments '("/bookmarks" "/password" "2023"))
;; (defvar doc2pdf-arguments '("/password" "2023"))
(defvar doc2pdf-arguments '("/password" "202"))
(define-derived-mode zyt/doc-mode special-mode "ZYT Microsoft office document View"
  "View ms office files with pdf"
  (if doc2pdf-convert-in-process
	  (message "前一次转换操作尚未结束，请稍候尝试...")
	(lexical-let* (
				   (cur-buffer (current-buffer))
				   (doc-file (buffer-file-name))
				   (directory (file-name-directory doc-file))
				   (pdf-file (concat (file-name-directory doc-file)
									 (file-name-base doc-file)
									 ".pdf"))
				   (pdf-file (concat directory  (file-name-base doc-file) "_" (secure-hash 'md5 cur-buffer) ".pdf"))
				   )
	  (if (file-exists-p pdf-file)
		  (progn
			(kill-buffer cur-buffer)
			(find-file pdf-file)
			;; (setf buffer-file-name (concat (file-name-directory doc-file)
			;; (file-name-base doc-file)
			;; ".pdf"))
			)
		(setq doc2pdf-convert-in-process t)
		(setq doc2pdf-convert-process
			  (apply
			   #'start-process
			   "doc2pdf"
			   (get-buffer-create "doc2pdf output")
			   "D:/Software/Editor/OfficeToPdf/OfficeToPDF/OfficeToPDF.exe"
			   (append
				doc2pdf-arguments
				(list doc-file)
				(list pdf-file)
				)
			   )
			  )
		(set-process-sentinel doc2pdf-convert-process
							  (lambda(process output)
								(setq doc2pdf-convert-in-process nil)
								(message (format "process %s code: %d\n" (process-status doc2pdf-convert-process) (process-exit-status doc2pdf-convert-process)))
								(if (null (string-prefix-p "finished" output))
									(progn
									  (message (format "转换失败：%s, 使用office打开相应文件" output))
									  ;; (dired-visit-node-in-external-application)
									  )
								  (kill-buffer cur-buffer)
								  (message (format "转换结束:%s" output))
								  (find-file pdf-file)
								  ;; (setf buffer-file-name (concat (file-name-directory doc-file)
								  ;; (file-name-base doc-file)
								  ;; ".pdf"))
								  )
								(with-current-buffer
									(process-buffer doc2pdf-convert-process)
								  (end-of-buffer)
								  (insert output)
								  )
								)
							  )
		(message "文件转换中，请稍候...")
		(kill-buffer cur-buffer)
		(find-file directory)
		)
	  )
	)
  )
(provide 'msdoc2pdf)
