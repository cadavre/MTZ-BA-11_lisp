;-------------------------------------------------------------------------------
; Parametryk dla czesci MTZ-BA-11 - pierscien dzielony
; Komputerowe wspomaganie projektowania (CAD), prof. dr hab. in¿. Piotr Gendarz
; 
; Wykonane przez:
; Seweryn Zeman
; Grupa IV
; MT, AiR, semestr VI
; 2012
;-------------------------------------------------------------------------------

(vl-load-com)
(command "_osnap" "_none")			; wylacz przyciaganie
(setq mypath "D:\\parametryk\\sz\\")

(load (strcat mypath "getexcel.lsp"))
									; narzedzia do pracy z arkuszem Excel

(GetExcel (strcat mypath "values.xls") "Arkusz1" "J14")

(setq wariant (getint "Podaj numer szeregu [-4,4]: "))
									; h1, h2, h3, h4, w1, w2, w3, fazaM, fazaD
									; komórki B6-J14
									
(setq pPoczStr (list 0 0))			; punkt rozpoczecia rysowania ramki
(setq pPoczRys (list 100 (- 90 (* wariant 10))))		; punkt rozpoczecia rysowania przeciecia
(setq karb 3)

(setq kolumna (+ wariant 70))		; konwertuje nr wariantu na kod ASCII odpowiadajacy kolumnie Excela F=70
(setq pWiersz 5)					; pierwszy wiersz - 1 w Excelu, uzycie np. (+ pWiersz 1) dla pierwszego

(setq dane (list 0))				; utworz liste z zerem dla indeksu 0
(foreach n '("6" "7" "8" "9" "10" "11" "12" "13" "14")
									; dla kazdego elementu
  (setq dane (cons (atof (GetCell (strcat (chr kolumna) n))) dane))
  (print "Odczytano z arkusza:")
  (print (GetCell (strcat (chr kolumna) n)))
									; dodaj dane do listy
)
(setq dane (reverse dane))			; odwroc liste do odpowiedniej kolejnosci

(setq pSrodKol (list (+ (car pPoczRys) (nth 1 dane) (- 190 (nth 1 dane))) (+ (cadr pPoczRys) (/ (nth 1 dane) 1.5)))) ; punkt rozpoczecia rysowania widoku
(setq
  pA1 (list (car pPoczRys) (cadr pPoczRys))
  pA2a (list (car pA1) (- (+ (cadr pA1) (/ (nth 1 dane) 2)) (nth 8 dane)))
  pA2b (list (+ (car pA2a) (nth 8 dane)) (+ (cadr pA2a) (nth 8 dane)))
  pA3 (list (+ (- (car pA2b) (nth 8 dane)) (- (nth 5 dane) (nth 6 dane))) (cadr pA2b))
  pA4 (list (car pA3) (- (cadr pA3) (/ (- (nth 1 dane) (nth 2 dane)) 2)))
  pA5 (list (+ (car pA4) (nth 6 dane)) (cadr pA4))
  pA6 (list (car pA5) (cadr pPoczRys))
  pB1 (list (+ (car pPoczRys) (nth 8 dane)) (cadr pPoczRys))
  pB2 (list (car pB1) (+ (cadr pB1) (/ (nth 3 dane) 2)))
  pB3 (list (- (+ (car pB2) (- (nth 5 dane) (nth 7 dane))) (nth 8 dane)) (cadr pB2))
  pB4 (list (car pB3) (cadr pPoczRys))
  pC1 (list (+ (car pPoczRys) (+ (- (nth 5 dane) (nth 7 dane)) (nth 9 dane))) (cadr pPoczRys))
  pC2 (list (car pC1) (+ (cadr pC1) (/ (nth 4 dane) 2)))
  pC3 (list (+ (car pC2) (- (nth 7 dane) (* 2 (nth 9 dane)))) (cadr pC2))
  pC4 (list (car pC3) (cadr pPoczRys))
  pD1 (list (car pPoczRys) (+ (+ (cadr pB1) (/ (nth 3 dane) 2)) (nth 8 dane)))
  pD2 (list (car pB3) (+ (cadr pC2) (nth 9 dane)))
  pD3 (list (car pA6) (+ (cadr pC3) (nth 9 dane)))
)

(command "_LAYER" "T" "0" "K" "7" "0" "L" "CONTINUOUS" "0" "")
									; warstwa rysunku

(command "_pline" pA1 pA2a pA2b pA3 pA4 pA5 pA6 pA1 "")
(command "_pline" pB1 pB2 pB3 pB4 "")
(command "_pline" pC1 pC2 pC3 pC4 "")
(command "_pline" pD1 pB2 "")
(command "_pline" pD2 pC2 "")
(command "_pline" pD3 pC3 "")

(setq								; ramka strony
  pR1 (list 0 0)
  pR2 (list 420 297)
  pR3 (list (+ (car pR1) 5) (+ (cadr pR1) 5))
  pR4 (list (- (car pR2) 5) (- (cadr pR2) 5))
)
(command "_RECTANGLE" pR1 pR2)
(command "_RECTANGLE" pR3 pR4)

									; tekst bialy
(command "_text" (list (+ (car pSrodKol) 30) (- (cadr pPoczRys) 10)) "5" "" "MTZ-BA-11")
(command "_text" (list (+ (car pSrodKol) 30) (- (cadr pPoczRys) 30)) "5" "" "Pierœcieñ dzielony")


(setq
  pE1 (list (- (car pSrodKol) (/ (nth 1 dane) 4)) (+ (cadr pSrodKol) 0.75))
  pE2 (list (+ (car pE1) (/ (nth 1 dane) 2)) (cadr pE1))
  pE3 (list (car pE1) (- (cadr pE1) 1.5))
  pE4 (list (car pE2) (- (cadr pE2) 1.5))
)

(command "_pline" pE1 pE2 "")
(command "_pline" pE3 pE4 "")

(command "_circle" pSrodKol "_d" (/ (nth 1 dane) 2))
(command "_circle" pSrodKol "_d" (/ (nth 3 dane) 2))
(command "_circle" pSrodKol "_d" (/ (nth 4 dane) 2))
(command "_circle" pSrodKol "_d" (/ (- (nth 1 dane) (* 2 (nth 8 dane))) 2))
(command "_circle" pSrodKol "_d" (/ (- (nth 3 dane) (* 2 (nth 8 dane))) 2))
(command "_circle" pSrodKol "_d" (/ (- (nth 4 dane) (* 2 (nth 9 dane))) 2))
;(command "_arc" "_c" pSrodKol pE2 pE1)
;(command "_arc" "_c" pSrodKol pE3 pE4)
;(command "_arc" "_c" pSrodKol (list (- (car pE2) (nth 8 dane)) (cadr pE2)) (list (+ (car pE1) (nth 8 dane)) (cadr pE1)))
;(command "_arc" "_c" pSrodKol (list (+ (car pE3) (nth 8 dane)) (cadr pE3)) (list (- (car pE4) (nth 8 dane)) (cadr pE4)))

(command "_LAYER" "T" "1" "L" "ACAD_ISO04W100" "1" "")
									; warstwa linii srodkowych

(command "_pline" (list (- (car pSrodKol) (/ (nth 1 dane) 3)) (cadr pSrodKol)) (list (+ (car pSrodKol) (/ (nth 1 dane) 3)) (cadr pSrodKol)) "")
(command "_pline" (list (car pSrodKol) (- (cadr pSrodKol) (/ (nth 1 dane) 3))) (list (car pSrodKol) (+ (cadr pSrodKol) (/ (nth 1 dane) 3))) "")

(command "_pline" (list (- (car pPoczRys) 40) (- (cadr pPoczRys) 1.5)) (list (+ (+ (car pPoczRys) 60) (nth 5 dane)) (- (cadr pPoczRys) 1.5)) "")

(command "_LAYER" "T" "2" "K" "3" "2" "L" "CONTINUOUS" "2" "")
									; warstwa kreskowania

(command "-KRESKUJ" "A" "ANSI31" "1" "0" (list (car pB3) (+ (cadr pB3) 10)) "")

(command "_LAYER" "T" "3" "K" "2" "3" "L" "CONTINUOUS" "3" "")
									; warstwa wymiarow

(command "_DIM" "_VER" pE2 pE4 (list (+ (car pE2) 20) (cadr pE2)) (rtos karb 2 1) "_EXIT")

(command "DIMBLK1" "_NONE")			; brak drugiej strzalki wymiarowej
(command "DIMSAH" "1")				; wlacz tryb bez jednej strzalki
(command "_DIM" "_VER" pB1 pB2 (list (- (car pB1) 30) (cadr pB1)) (strcat "%%c" (rtos (nth 3 dane) 2 1) "H10") "_EXIT")
(command "_DIM" "_VER" pC3 pC4 (list (+ (car pA6) 30) (cadr pA6)) (strcat "%%c" (rtos (nth 4 dane) 2 1) "H8") "_EXIT")
(command "_DIM" "_VER" pA5 pA6 (list (+ (car pA6) 40) (cadr pA6)) (strcat "%%c" (rtos (nth 2 dane) 2 1) "h8") "_EXIT")
(command "_DIM" "_VER" pA3 pA6 (list (+ (car pA6) 50) (cadr pA6)) (strcat "%%c" (rtos (nth 1 dane) 2 1) "h10") "_EXIT")
(command "DIMSAH" "0")				; wlacz tryb z dwiema strzalkami
(command "_DIM" "_HOR" pB4 pA6 (list (car pA6) (- (cadr pA6) 10)) (strcat (rtos (nth 7 dane) 2 1) "h10") "_EXIT")
(command "_DIM" "_HOR" pA2a pA5 (list (car pA3) (+ (cadr pA3) 50)) (rtos (nth 5 dane) 2 1) "_EXIT")
(command "_DIM" "_HOR" pA3 pA5 (list (car pA3) (+ (cadr pA3) 40)) (strcat (rtos (nth 6 dane) 2 1) "H12") "_EXIT")
(command "_DIM" "_HOR" pA2b pA2a (list (car pA3) (+ (cadr pA3) 10)) (strcat (rtos (nth 8 dane) 2 1) "x45%%d") "_EXIT")
(command "_DIM" "_HOR" pC1 pB4 (list (car pB4) (+ (cadr pB4) 5)) (strcat (rtos (nth 9 dane) 2 1) "x45%%d") "_EXIT")
(command "_DIM" "_HOR" pC4 pA6 (list (car pB4) (+ (cadr pB4) 5)) (strcat (rtos (nth 9 dane) 2 1) "x45%%d") "_EXIT")

; bloki
(command "_insert" (strcat mypath "przekr.dwg") (list (car pSrodKol) (+ (cadr pSrodKol) (/ (nth 1 dane) 4) 15)) 0.7 0.7 "")
(command "_insert" (strcat mypath "przekr.dwg") (list (car pSrodKol) (- (cadr pSrodKol) 10)) 0.7 0.7 "")
(command "_insert" (strcat mypath "blok_tol_wym005.dwg") (list (+ (car pA6) 50) (cadr pA3)) 1.0 1.0 "")
(command "_insert" (strcat mypath "bazaA.dwg") (list (+ (car pA6) 17) (cadr pC3)) 1.0 1.0 "")
(command "_insert" (strcat mypath "chrop25.dwg") (list (+ (car pA4) 7) (cadr pA4)) 1.0 1.0 "")
(command "_insert" (strcat mypath "chrop25.dwg") (list (+ (car pA3) 40) (cadr pA3)) 1.0 1.0 "")
(command "_insert" (strcat mypath "chrop25.dwg") "_R" "90" (list (car pB3) (- (cadr pB3) 20)) 1.0 1.0)
(command "_insert" (strcat mypath "chrop25.dwg") "_R" "270" (list (car pA3) (+ (cadr pA3) 20)) 1.0 1.0)
(command "_insert" (strcat mypath "chrop.dwg") (list (- (car pR4) 10) (- (cadr pR4) 10)) 1.0 1.0 "")

; tekst
(command "_text" (list (car pSrodKol) (+ (cadr pPoczRys) 20)) "2.5" "" "Uwagi:" "_text" "" "1. Galwanizowaæ antykorozyjnie, gruboœæ warstwy min. 0.02mm" "_text" "" "2. Wymiary tolerowane sprawdziæ po galwanizacji." "_text" "" "3. Ostre krawêdzie stêpiæ.")
(command "_text" (list (car pSrodKol) (- (cadr pPoczRys) 20)) "5" "" "1:2    1.344    41Cr4")
(command "_text" (list (- (car pPoczRys) 40) (+ (cadr pSrodKol) (/ (nth 1 dane) 4) 10)) "5" "" "A - A" "_text" "" "Podz. 1:1")

; zoom all
(command "_zoom" "_e")