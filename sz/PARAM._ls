;-------------------------------------------------------------------------------
; Parametryk dla czesci MTZ-BA-11 - pierscien dzielony.
; Komputerowe wspomaganie projektowania (CAD), prof. dr hab. in¿. Piotr Gendarz
; 
; Wykonane przez:
; Seweryn Zeman
; Grupa IV
; MT, AiR, semestr VI
; 2012
;-------------------------------------------------------------------------------


(vl-load-com)
(load "D:\\parametryk\\sz\\utilities\\getexcel.lsp")
					; narzedzia do pracy z arkuszem Excel

(GetExcel "D:\\parametryk\\sz\\values.xls" "Arkusz1" "J14")

(setq pPoczStr (list 0 0))		; punkt rozpoczecia rysowania ramki
(setq pPoczRys (list 50 50))		; punkt rozpoczecia rysowania rysunku

(setq wariant (getint "Podaj numer szeregu [-4,4]: "))
					; h1, h2, h3, h4, w1, w2, w3, fazaM, fazaD
					; komórki B6-J14

(setq kolumna (+ wariant 70))		; konwertuje nr wariantu na kod ASCII odpowiadajacy kolumnie Excela F=70
(setq pWiersz 5)			; pierwszy wiersz - 1 w Excelu, uzycie np. (+ pWiersz 1) dla pierwszego

(setq dane (list 0))			; utworz liste z zerem dla indeksu 0
(foreach n '("6" "7" "8" "9" "10" "11" "12" "13" "14")
					; dla kazdego elementu
  (setq dane (cons (atof (GetCell (strcat (chr kolumna) n))) dane))
  (print "Odczytano z arkusza:")
  (print (GetCell (strcat (chr kolumna) n)))
					; dodaj dane do listy
)
(setq dane (reverse dane))		; odwroc liste do odpowiedniej kolejnosci

(setq
  pA1 (list (car pPoczRys) (cadr pPoczRys))
  pA2 (list (car pA1) (+ (cadr pA1) (/ (nth 1 dane) 2)))
  pA3 (list () (cadr pA2))
)

(command "_line" pA1 pA2 "")
					;(command "_circle" p14 "_d" d3)
