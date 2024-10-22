(require 'cl)
(require 'cl-macs)

(add-to-list 'auto-mode-alist '("\\.\\(docx?\\|DOCX?\\|pptx?\\|PPTX?\\|xlsx?\\|XLSX?\\)\\'" . zyt/doc-mode))
(defvar doc2pdf-convert-process nil)
(defvar doc2pdf-convert-in-process nil)
(defvar doc2pdf-arguments '("/readonly" "/hidden"))

(defcustom msdoc-to-pdf-program
  (let ((executable (if (eq system-type 'windows-nt)
                        "OfficeToPDF.exe" "OfficeToPDF"))
        (default-directory
         (or (and load-file-name
                  (file-name-directory load-file-name))
             default-directory)))
    (cl-labels ((try-directory (directory)
                  (and (file-directory-p directory)
                       (file-executable-p (expand-file-name executable directory))
                       (expand-file-name executable directory))))
      (or (executable-find executable)
          ;; This works if epdfinfo is in the same place as emacs and
          ;; the editor was started with an absolute path, i.e. it is
          ;; meant for Windows/Msys2.
          (and (stringp (car-safe command-line-args))
               (file-name-directory (car command-line-args))
               (try-directory
                (file-name-directory (car command-line-args))))
          ;; If we are running directly from the git repo.
          (try-directory (expand-file-name "../server"))
          ;; Fall back to epdfinfo in the directory of this file.
          (expand-file-name executable))))
  "Filename of the msdoc2pdf executable."
  :type 'file)


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
		(message (format "Convert cmd:\n%s %s"
						 msdoc-to-pdf-program
						 (mapconcat
						  'identity
						  (append
						   doc2pdf-arguments
						   (list doc-file)
						   (list pdf-file)
						   )
						  " "
						  )))
		(defun process-handler(process output)
		  (setq doc2pdf-convert-in-process nil)
		  (let ((exit-code (process-exit-status doc2pdf-convert-process)))
			(if (member exit-code '(1 5))
				(when-let ((password
							(read-from-minibuffer
							 (or (and (= 5 exit-code) "请输入文件密码: ")
								 "密码错误，请重新输入:")
							 nil nil nil  
							 'minibuffer-history)))
				  (setq doc2pdf-convert-process
						(apply
						 #'start-process
						 "doc2pdf"
						 (get-buffer-create "doc2pdf output")
						 msdoc-to-pdf-program
						 (append
						  doc2pdf-arguments
						  (list "/password" password)
						  (list "/pdf_user_pass" password)
						  (list doc-file)
						  (list pdf-file)
						  )
						 )
						)
				  (set-process-sentinel doc2pdf-convert-process #'process-handler)					  
				  (message "文件转换中，请稍候...")
				  )
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
			))
		(setq doc2pdf-convert-process
			  (apply
			   #'start-process
			   "doc2pdf"
			   (get-buffer-create "doc2pdf output")
			   msdoc-to-pdf-program
			   (append
				doc2pdf-arguments
				(list doc-file)
				(list pdf-file)
				)
			   )
			  )
		(setq doc2pdf-convert-in-process t)
		(set-process-sentinel doc2pdf-convert-process #'process-handler)					  
		(message "文件转换中，请稍候...")
		(kill-buffer cur-buffer)
		(find-file directory)
		)
	  )
	)
  )
(provide 'msdoc2pdf)
