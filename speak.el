;;; speak.el --- text to speech extension.

;; Copyright (C) 2014  Daisuke Kobayashi

;; Author: Daisuke Kobayashi <d5884jp@gmail.com>
;; Version: 0.1
;; Keywords: multimedia

;; This program is free software; you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation, either version 3 of the License, or
;; (at your option) any later version.

;; This program is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.

;; You should have received a copy of the GNU General Public License
;; along with this program.  If not, see <http://www.gnu.org/licenses/>.

;;; Commentary:

;; 

;;; Code:

(defcustom speak-method
  (or (executable-find "say")
      (executable-find "espeak")
      (and (eq system-type 'windows-nt)
	   'speak--w32-sapi))
  "Method for speaking.
If simple string, execute as a command.
If function, call as a function.")

;;;###autoload
(defun speak (sentence)
  "Speak SENTENCE."
  (interactive "sSentence: ")
  (speak--internal sentence))

;;;###autoload
(defun speak-region (start end)
  "Speak text on region from START to END."
  (interactive "r")
  (speak (buffer-substring start end)))

;;;###autoload
(defun speak-at-point ()
  "Speak word at point."
  (interactive)
  (let ((word (if (region-active-p)
		  (buffer-substring (region-beginning) (region-end))
		(thing-at-point 'word))))
    (when word
      (speak word))))

(defvar speak--proc nil)

(defun speak--internal (sentence)
  (when (and speak--proc
	     (process-live-p speak--proc))
    (kill-process speak--proc)
    (setq speak-proc nil))
  (cond
   ((stringp speak-method)
    (setq speak--proc (start-process "speak" nil speak-method sentence))
    (set-process-query-on-exit-flag speak--proc nil))
   ((functionp speak-method)
    (funcall speak-method sentence))
   (t
    (message "Please set `speak-method'.")))
  t)

;; internal functions

(defvar speak--w32-proc nil)

(defun speak--w32-start ()
  (when (or (not speak--w32-proc)
	    (not (process-live-p speak--w32-proc)))
    (setq speak--w32-proc (start-process "pws-for-speak" nil "powershell" "-Command" "-"))
    (let ((process-connection-type nil))
      (process-send-string speak--w32-proc "$speak = New-Object -ComObject SAPI.SPVoice;\n"))
    (set-process-query-on-exit-flag speak--w32-proc nil)))

(defun speak--w32-sapi (sentence)
  (speak--w32-start)
  (process-send-string speak--w32-proc
		       (format "$speak.Speak('%s')\n"
			       (replace-regexp-in-string "'" "''" sentence))))


(defun speak--w32-set-rate (rate)
  (speak--w32-start)
  (process-send-string speak--w32-proc
		       (format "$speak.Rate=%d\n" rate)))

(provide 'speak)

;;; speak.el ends here
