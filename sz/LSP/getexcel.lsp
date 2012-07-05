;-------------------------------------------------------------------------------
; Program Name: GetExcel.lsp [GetExcel R4]
; Created By:   Terry Miller (Email: terrycadd@yahoo.com)
;               (URL: http://web2.airmail.net/terrycad)
; Date Created: 9-20-03
; Function:     Several functions to get and put values into Excel cells.
;-------------------------------------------------------------------------------
; Revision History
; Rev  By     Date    Description
;-------------------------------------------------------------------------------
; 1    TM   9-20-03   Initial version
; 2    TM   8-20-07   Rewrote GetExcel.lsp and added several new sub-functions
;                     including ColumnRow, Alpha2Number and Number2Alpha written
;                     by Gilles Chanteau from Marseille, France.
; 3    TM   12-1-07   Added several sub-functions written by Gilles Chanteau
;                     including Cell-p, Row+n, and Column+n. Also added his
;                     revision of the PutCell function.
; 4    GC   9-20-08   Revised the GetExcel argument MaxRange$ to accept a nil
;                     and get the current region from cell A1.
;-------------------------------------------------------------------------------
; Overview of Main functions
;-------------------------------------------------------------------------------
; GetExcel - Stores the values from an Excel spreadsheet into *ExcelData@ list
;   Syntax:  (GetExcel ExcelFile$ SheetName$ MaxRange$)
;   Example: (GetExcel "C:\\Folder\\Filename.xls" "Sheet1" "L30")
; GetCell - Returns the cell value from the *ExcelData@ list
;   Syntax:  (GetCell Cell$)
;   Example: (GetCell "H15")
; Function example of usage:
; (defun c:Get-Example ()
;   (GetExcel "C:\\Folder\\Filename.xls" "Sheet1" "L30");<-- Edit Filename.xls
;   (GetCell "H21");Or you can just use the global *ExcelData@ list
; );defun
;-------------------------------------------------------------------------------
; OpenExcel - Opens an Excel spreadsheet
;   Syntax:  (OpenExcel ExcelFile$ SheetName$ Visible)
;   Example: (OpenExcel "C:\\Folder\\Filename.xls" "Sheet1" nil)
; PutCell - Put values into Excel cells
;   Syntax:  (PutCell StartCell$ Data$) or (PutCell StartCell$ DataList@)
;   Example: (PutCell "A1" (list "GP093" 58.5 17 "Base" "3'-6 1/4\""))
; CloseExcel - Closes Excel session
;   Syntax:  (CloseExcel ExcelFile$)
;   Example: (CloseExcel "C:\\Folder\\Filename.xls")
; Function example of usage:
; (defun c:Put-Example ()
;   (OpenExcel "C:\\Folder\\Filename.xls" "Sheet1" nil);<-- Edit Filename.xls
;   (PutCell "A1" (list "GP093" 58.5 17 "Base" "3'-6 1/4\""));Repeat as required
;   (CloseExcel "C:\\Folder\\Filename.xls");<-- Edit Filename.xls
;   (princ)
; );defun
;-------------------------------------------------------------------------------
; Note: Review the conditions of each argument in the function headings
;-------------------------------------------------------------------------------
; GetExcel - Stores the values from an Excel spreadsheet into *ExcelData@ list
; Arguments: 3
;   ExcelFile$ = Path and filename
;   SheetName$ = Sheet name or nil for not specified
;   MaxRange$ = Maximum cell ID range to include or nil to get the current region from cell A1
; Syntax examples:
; (GetExcel "C:\\Temp\\Temp.xls" "Sheet1" "E19") = Open C:\Temp\Temp.xls on Sheet1 and read up to cell E19
; (GetExcel "C:\\Temp\\Temp.xls" nil "XYZ123") = Open C:\Temp\Temp.xls on current sheet and read up to cell XYZ123
;-------------------------------------------------------------------------------
(defun GetExcel (ExcelFile$ SheetName$ MaxRange$ / Column# ColumnRow@ Data@ ExcelRange^
  ExcelValue ExcelValue ExcelVariant^ MaxColumn# MaxRow# Range$ Row# Worksheet)
  (if (= (type ExcelFile$) 'STR)
    (if (not (findfile ExcelFile$))
      (progn
        (alert (strcat "Excel file " ExcelFile$ " not found."))
        (exit)
      );progn
    );if
    (progn
      (alert "Excel file not specified.")
      (exit)
    );progn
  );if
  (gc)
  (if (setq *ExcelApp% (vlax-get-object "Excel.Application"))
    (progn
      (alert "Close all Excel spreadsheets to continue!")
      (vlax-release-object *ExcelApp%)(gc)
    );progn
  );if
  (setq ExcelFile$ (findfile ExcelFile$))
  (setq *ExcelApp% (vlax-get-or-create-object "Excel.Application"))
  (vlax-invoke-method (vlax-get-property *ExcelApp% 'WorkBooks) 'Open ExcelFile$)
  (if SheetName$
    (vlax-for Worksheet (vlax-get-property *ExcelApp% "Sheets")
      (if (= (vlax-get-property Worksheet "Name") SheetName$)
        (vlax-invoke-method Worksheet "Activate")
      );if
    );vlax-for
  );if
  (if MaxRange$
    (progn
      (setq ColumnRow@ (ColumnRow MaxRange$))
      (setq MaxColumn# (nth 0 ColumnRow@))
      (setq MaxRow# (nth 1 ColumnRow@))
    );progn
    (progn
      (setq CurRegion (vlax-get-property (vlax-get-property
        (vlax-get-property *ExcelApp% "ActiveSheet") "Range" "A1") "CurrentRegion")
      );setq
      (setq MaxRow# (vlax-get-property (vlax-get-property CurRegion "Rows") "Count"))
      (setq MaxColumn# (vlax-get-property (vlax-get-property CurRegion "Columns") "Count"))
    );progn
  );if
  (setq *ExcelData@ nil)
  (setq Row# 1)
  (repeat MaxRow#
    (setq Data@ nil)
    (setq Column# 1)
    (repeat MaxColumn#
      (setq Range$ (strcat (Number2Alpha Column#)(itoa Row#)))
      (setq ExcelRange^ (vlax-get-property *ExcelApp% "Range" Range$))
      (setq ExcelVariant^ (vlax-get-property ExcelRange^ 'Value))
      (setq ExcelValue (vlax-variant-value ExcelVariant^))
      (setq ExcelValue
        (cond
          ((= (type ExcelValue) 'INT) (itoa ExcelValue))
          ((= (type ExcelValue) 'REAL) (rtosr ExcelValue))
          ((= (type ExcelValue) 'STR) (vl-string-trim " " ExcelValue))
          ((/= (type ExcelValue) 'STR) "")
        );cond
      );setq
      (setq Data@ (append Data@ (list ExcelValue)))
      (setq Column# (1+ Column#))
    );repeat
    (setq *ExcelData@ (append *ExcelData@ (list Data@)))
    (setq Row# (1+ Row#))
  );repeat
  (vlax-invoke-method (vlax-get-property *ExcelApp% "ActiveWorkbook") 'Close :vlax-False)
  (vlax-invoke-method *ExcelApp% 'Quit)
  (vlax-release-object *ExcelApp%)(gc)
  (setq *ExcelApp% nil)
  *ExcelData@
);defun GetExcel
;-------------------------------------------------------------------------------
; GetCell - Returns the cell value from the *ExcelData@ list
; Arguments: 1
;   Cell$ = Cell ID
; Syntax example: (GetCell "E19") = value of cell E19
;-------------------------------------------------------------------------------
(defun GetCell (Cell$ / Column# ColumnRow@ Return Row#)
  (setq ColumnRow@ (ColumnRow Cell$))
  (setq Column# (1- (nth 0 ColumnRow@)))
  (setq Row# (1- (nth 1 ColumnRow@)))
  (setq Return "")
  (if *ExcelData@
    (if (and (>= (length *ExcelData@) Row#)(>= (length (nth 0 *ExcelData@)) Column#))
      (setq Return (nth Column# (nth Row# *ExcelData@)))
    );if
  );if
  Return
);defun GetCell
;-------------------------------------------------------------------------------
; OpenExcel - Opens an Excel spreadsheet
; Arguments: 3
;   ExcelFile$ = Excel filename or nil for new spreadsheet
;   SheetName$ = Sheet name or nil for not specified
;   Visible = t for visible or nil for hidden
; Syntax examples:
; (OpenExcel "C:\\Temp\\Temp.xls" "Sheet2" t) = Opens C:\Temp\Temp.xls on Sheet2 as visible session
; (OpenExcel "C:\\Temp\\Temp.xls" nil nil) = Opens C:\Temp\Temp.xls on current sheet as hidden session
; (OpenExcel nil "Parts List" nil) =  Opens a new spreadsheet and creates a Part List sheet as hidden session
;-------------------------------------------------------------------------------
(defun OpenExcel (ExcelFile$ SheetName$ Visible / Sheet$ Sheets@ Worksheet)
  (if (= (type ExcelFile$) 'STR)
    (if (findfile ExcelFile$)
      (setq *ExcelFile$ ExcelFile$)
      (progn
        (alert (strcat "Excel file " ExcelFile$ " not found."))
        (exit)
      );progn
    );if
    (setq *ExcelFile$ "")
  );if
  (gc)
  (if (setq *ExcelApp% (vlax-get-object "Excel.Application"))
    (progn
      (alert "Close all Excel spreadsheets to continue!")
      (vlax-release-object *ExcelApp%)(gc)
    );progn
  );if
  (setq *ExcelApp% (vlax-get-or-create-object "Excel.Application"))
  (if ExcelFile$
    (if (findfile ExcelFile$)
      (vlax-invoke-method (vlax-get-property *ExcelApp% 'WorkBooks) 'Open ExcelFile$)
      (vlax-invoke-method (vlax-get-property *ExcelApp% 'WorkBooks) 'Add)
    );if
    (vlax-invoke-method (vlax-get-property *ExcelApp% 'WorkBooks) 'Add)
  );if
  (if Visible
    (vla-put-visible *ExcelApp% :vlax-true)
  );if
  (if (= (type SheetName$) 'STR)
    (progn
      (vlax-for Sheet$ (vlax-get-property *ExcelApp% "Sheets")
        (setq Sheets@ (append Sheets@ (list (vlax-get-property Sheet$ "Name"))))
      );vlax-for
      (if (member SheetName$ Sheets@)
        (vlax-for Worksheet (vlax-get-property *ExcelApp% "Sheets")
          (if (= (vlax-get-property Worksheet "Name") SheetName$)
            (vlax-invoke-method Worksheet "Activate")
          );if
        );vlax-for
        (vlax-put-property (vlax-invoke-method (vlax-get-property *ExcelApp% "Sheets") "Add") "Name" SheetName$)
      );if
    );progn
  );if
  (princ)
);defun OpenExcel
;-------------------------------------------------------------------------------
; PutCell - Put values into Excel cells
; Arguments: 2
;   StartCell$ = Starting Cell ID
;   Data@ = Value or list of values
; Syntax examples:
; (PutCell "A1" "PART NUMBER") = Puts PART NUMBER in cell A1
; (PutCell "B3" '("Dim" 7.5 "9.75")) = Starting with cell B3 put Dim, 7.5, and 9.75 across
;-------------------------------------------------------------------------------
(defun PutCell (StartCell$ Data@ / Cell$ Column# ExcelRange Row#)
  (if (= (type Data@) 'STR)
    (setq Data@ (list Data@))
  )
  (setq ExcelRange (vlax-get-property *ExcelApp% "Cells"))
  (if (Cell-p StartCell$)
    (setq Column# (car (ColumnRow StartCell$))
          Row# (cadr (ColumnRow StartCell$))
    );setq
    (if (vl-catch-all-error-p
          (setq Cell$ (vl-catch-all-apply 'vlax-get-property
            (list (vlax-get-property *ExcelApp% "ActiveSheet") "Range" StartCell$))
          );setq
        );vl-catch-all-error-p
        (alert (strcat "The cell ID \"" StartCell$ "\" is invalid."))
        (setq Column# (vlax-get-property Cell$ "Column")
              Row# (vlax-get-property Cell$ "Row")
        );setq
    );if
  );if
  (if (and Column# Row#)
    (foreach Item Data@
      (vlax-put-property ExcelRange "Item" Row# Column# (vl-princ-to-string Item))
      (setq Column# (1+ Column#))
    );foreach
  );if
  (princ)
);defun PutCell
;-------------------------------------------------------------------------------
; CloseExcel - Closes Excel spreadsheet
; Arguments: 1
;   ExcelFile$ = Excel saveas filename or nil to close without saving
; Syntax examples:
; (CloseExcel "C:\\Temp\\Temp.xls") = Saveas C:\Temp\Temp.xls and close
; (CloseExcel nil) = Close without saving
;-------------------------------------------------------------------------------
(defun CloseExcel (ExcelFile$ / Saveas)
  (if ExcelFile$
    (if (= (strcase ExcelFile$) (strcase *ExcelFile$))
      (if (findfile ExcelFile$)
        (vlax-invoke-method (vlax-get-property *ExcelApp% "ActiveWorkbook") "Save")
        (setq Saveas t)
      );if
      (if (findfile ExcelFile$)
        (progn
          (vl-file-delete (findfile ExcelFile$))
          (setq Saveas t)
        );progn
        (setq Saveas t)
      );if
    );if
  );if
  (if Saveas
    (vlax-invoke-method (vlax-get-property *ExcelApp% "ActiveWorkbook")
      "SaveAs" ExcelFile$ -4143 "" "" :vlax-false :vlax-false nil
    );vlax-invoke-method
  );if
  (vlax-invoke-method (vlax-get-property *ExcelApp% "ActiveWorkbook") 'Close :vlax-False)
  (vlax-invoke-method *ExcelApp% 'Quit)
  (vlax-release-object *ExcelApp%)(gc)
  (setq *ExcelApp% nil *ExcelFile$ nil)
  (princ)
);defun CloseExcel
;-------------------------------------------------------------------------------
; ColumnRow - Returns a list of the Column and Row number
; Function By: Gilles Chanteau from Marseille, France
; Arguments: 1
;   Cell$ = Cell ID
; Syntax example: (ColumnRow "ABC987") = '(731 987)
;-------------------------------------------------------------------------------
(defun ColumnRow (Cell$ / Column$ Char$ Row#)
  (setq Column$ "")
  (while (< 64 (ascii (setq Char$ (strcase (substr Cell$ 1 1)))) 91)
    (setq Column$ (strcat Column$ Char$)
          Cell$ (substr Cell$ 2)
    );setq
  );while
  (if (and (/= Column$ "") (numberp (setq Row# (read Cell$))))
    (list (Alpha2Number Column$) Row#)
    '(1 1);default to "A1" if there's a problem
  );if
);defun ColumnRow
;-------------------------------------------------------------------------------
; Alpha2Number - Converts Alpha string into Number
; Function By: Gilles Chanteau from Marseille, France
; Arguments: 1
;   Str$ = String to convert
; Syntax example: (Alpha2Number "ABC") = 731
;-------------------------------------------------------------------------------
(defun Alpha2Number (Str$ / Num#)
  (if (= 0 (setq Num# (strlen Str$)))
    0
    (+ (* (- (ascii (strcase (substr Str$ 1 1))) 64) (expt 26 (1- Num#)))
       (Alpha2Number (substr Str$ 2))
    );+
  );if
);defun Alpha2Number
;-------------------------------------------------------------------------------
; Number2Alpha - Converts Number into Alpha string
; Function By: Gilles Chanteau from Marseille, France
; Arguments: 1
;   Num# = Number to convert
; Syntax example: (Number2Alpha 731) = "ABC"
;-------------------------------------------------------------------------------
(defun Number2Alpha (Num# / Val#)
  (if (< Num# 27)
    (chr (+ 64 Num#))
    (if (= 0 (setq Val# (rem Num# 26)))
      (strcat (Number2Alpha (1- (/ Num# 26))) "Z")
      (strcat (Number2Alpha (/ Num# 26)) (chr (+ 64 Val#)))
    );if
  );if
);defun Number2Alpha
;-------------------------------------------------------------------------------
; Cell-p - Evaluates if the argument Cell$ is a valid cell ID
; Function By: Gilles Chanteau from Marseille, France
; Arguments: 1
;   Cell$ = String of the cell ID to evaluate
; Syntax examples: (Cell-p "B12") = t, (Cell-p "BT") = nil
;-------------------------------------------------------------------------------
(defun Cell-p (Cell$)
  (and (= (type Cell$) 'STR)
    (or (= (strcase Cell$) "A1")
      (not (equal (ColumnRow Cell$) '(1 1)))
    );or
  );and
);defun Cell-p
;-------------------------------------------------------------------------------
; Row+n - Returns the cell ID located a number of rows from cell
; Function By: Gilles Chanteau from Marseille, France
; Arguments: 2
;   Cell$ = Starting cell ID
;   Num# = Number of rows from cell
; Syntax examples: (Row+n "B12" 3) = "B15", (Row+n "B12" -3) = "B9"
;-------------------------------------------------------------------------------
(defun Row+n (Cell$ Num#)
  (setq Cell$ (ColumnRow Cell$))
  (strcat (Number2Alpha (car Cell$)) (itoa (max 1 (+ (cadr Cell$) Num#))))
);defun Row+n
;-------------------------------------------------------------------------------
; Column+n - Returns the cell ID located a number of columns from cell
; Function By: Gilles Chanteau from Marseille, France
; Arguments: 2
;   Cell$ = Starting cell ID
;   Num# = Number of columns from cell
; Syntax examples: (Column+n "B12" 3) = "E12", (Column+n "B12" -1) = "A12"
;-------------------------------------------------------------------------------
(defun Column+n (Cell$ Num#)
  (setq Cell$ (ColumnRow Cell$))
  (strcat (Number2Alpha (max 1 (+ (car Cell$) Num#))) (itoa (cadr Cell$)))
);defun Column+n
;-------------------------------------------------------------------------------
; rtosr - Used to change a real number into a short real number string
; stripping off all trailing 0's.
; Arguments: 1
;   RealNum~ = Real number to convert to a short string real number
; Returns: ShortReal$ the short string real number value of the real number.
;-------------------------------------------------------------------------------
(defun rtosr (RealNum~ / DimZin# ShortReal$)
  (setq DimZin# (getvar "DIMZIN"))
  (setvar "DIMZIN" 8)
  (setq ShortReal$ (rtos RealNum~ 2 8))
  (setvar "DIMZIN" DimZin#)
  ShortReal$
);defun rtosr
;-------------------------------------------------------------------------------
(princ);End of GetExcel.lsp
