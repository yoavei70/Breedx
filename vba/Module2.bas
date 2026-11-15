Attribute VB_Name = "Module2"
'====================================================================================
' modSeedOrder  -  Build an Inventory / Seed Order workbook from a Proforma Invoice PDF
'------------------------------------------------------------------------------------
' WHAT IT DOES
'   0. Reads run parameters from the "seedOrder" sheet of THIS workbook
'      (coding Breedx.xlsm) - falls back to the constants below if the sheet
'      is missing or a cell is empty.
'   1. Reads the line-item table of the input PDF (ProformaInvoice_<n>_source.pdf)
'      with Excel's PDF connector (modPepperLines) - Word is only a last resort.
'   2. Opens the variety inventory workbook (e.g. "\Inventory\<B7>.xlsx") and looks
'      up, per invoice Variety, the first matching Lot number, its Germination
'      validity (last 3 chars of column M) and its Weight per 1000 seeds (column I).
'   3. Opens the template (FO 013H Seed Order Form HE&EN.xlsx) and copies its
'      form sheet into a brand-new workbook.
'   4. Converts the template to English only, drops the Hebrew letterhead graphic
'      and drops in the breedx logo picture (logo.jpg) in cell A1.
'   5. Writes the header block (orderer / creation date / due date / country /
'      order no.) and one table row per "Pepper ..." invoice line
'      (Freight / Subtotal / VAT / Total lines are ignored).
'   6. Saves the result as InventoryOrder_<n>.xlsx and closes every file it opened.
'
' HOW TO USE
'   - In Excel: Alt+F11 -> File -> Import File... -> pick this .bas
'     and modPepperLines.bas (PDF reader used by step 1)
'   - Make sure "logo.jpg" sits in the input directory (already placed there).
'   - Run  CreateSeedOrder
'
' PARAMETERISE LATER
'   Everything marked  '<< PARAM  is currently hard-coded per the spec and is a
'   candidate for a real parameter / Office Form field in the future.
'====================================================================================
Option Explicit

'--- Fallback parameters (used only if the seedOrder sheet does not supply them) ---
Private Const DEF_INPUT_DIR   As String = "C:\Users\YoavEilenberg\Documents\Ordering"
Private Const DEF_EMPTY_FORM  As String = "FO 013H Seed Order Form HE&EN.xlsx"
Private Const DEF_ORDER_NO    As String = "90476"                         '<< PARAM
Private Const LOGO_FILE       As String = "logo.jpg"
Private Const DEF_INV_DIR     As String = "C:\Users\YoavEilenberg\Documents\Inventory"
Private Const DEF_INV_FILE    As String = "Peppers"                '<< PARAM - Inventory\Peppers.xlsx
Private Const DEF_INV_SHEET   As String = "Inventory Management"   ' matched case/space-insensitively

'--- Header mapping (hard-coded per spec - future parameters) ----------------------
Private Const ORDERER_NAME    As String = "Yvette Naftali"                '<< PARAM
Private Const COUNTRY_NAME    As String = "Mexico"                        '<< PARAM

'--- Per-row values still hard-coded per spec - the "Andy"/"Alex" placeholders are
'    gone now that Lot / Germination / Weight-per-1000 / net weight come from the
'    Inventory workbook; only these two remain genuinely hard-coded -------------
Private Const V_SEEDS_IN_PACK As String = "1000"      ' Seeds - in package  <<
Private Const V_ADDED_PCT     As String = "20%"       ' Added Seeds - %     <<

'--- layout constant -------------------------------------------------------------
Private Const FIRST_DATA_ROW As Long = 15             ' first row of the item table in the template

'--- inventory lookup result ------------------------------------------------------
Private Type TInvLookup
    Found      As Boolean
    Lot        As String
    GermPct    As String
    Weight1000 As Variant
    HasExpiry  As Boolean
    ExpiryDate As Date
End Type

'--- Column layout of the template form (1 = A) -----------------------------------
Private Const C_NUM     As Long = 1    ' A  #
Private Const C_SPECIES As Long = 2    ' B  Species
Private Const C_VARIETY As Long = 3    ' C  Variety / Catalog Code
Private Const C_LOT     As Long = 4    ' D  Lot (from Inventory lookup)
Private Const C_GERM    As Long = 5    ' E  Germination %
Private Const C_QTY     As Long = 6    ' F  Quantity of seeds (x1000)
Private Const C_NPACK   As Long = 7    ' G  Seeds - number of packages
Private Const C_INPACK  As Long = 8    ' H  Seeds - in package
Private Const C_W1000   As Long = 9    ' I  Seeds - weight gr of 1000
Private Const C_NETW    As Long = 10   ' J  Seeds - total net weight gr
Private Const C_ADDPCT  As Long = 11   ' K  Added seeds - %
Private Const C_ADDW    As Long = 12   ' L  Added seeds - weight gr   (formula)
Private Const C_TOTW    As Long = 13   ' M  Total seeds - yield gr    (formula)
Private Const C_TOTPACK As Long = 14   ' N  Total packages           (formula)
Private Const C_FINSEED As Long = 15   ' O  Final - seeds in package (formula)
Private Const C_FINW    As Long = 16   ' P  Final - weight gr in pack (formula)
Private Const C_LABELS  As Long = 17   ' Q  Labels
Private Const C_EXPIRY  As Long = 18   ' R  Expiration date  (added)
Private Const C_COMMENT As Long = 19   ' S  Comments         (added)

'--- Item record -----------------------------------------------------------------
Private Type TItem
    Species  As String
    variety  As String
    quantity As Double
End Type


'================================================================================
'  MAIN
'================================================================================
Public Sub CreateSeedOrder()

    Dim inputDir As String, emptyForm As String, orderNo As String
    Dim invDir As String, invVarietyFile As String, invSheetName As String
    Dim pdfPath As String, outPath As String, logoPath As String
    Dim items() As TItem, nItems As Long
    Dim lookups() As TInvLookup, warnings As String, notes As String
    Dim tplWb As Workbook, outWb As Workbook, ws As Worksheet
    Dim invWb As Workbook, invWs As Worksheet
    Dim i As Long, r As Long

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '--- 0. parameters --------------------------------------------------------
    ReadParams inputDir, emptyForm, orderNo, invDir, invVarietyFile, invSheetName
    If Right$(inputDir, 1) <> "\" Then inputDir = inputDir & "\"
    If Len(invDir) > 0 And Right$(invDir, 1) <> "\" Then invDir = invDir & "\"

    ' Sanity-check the folder itself before building any file path from it -
    ' a bad/garbled inputDir is the #1 cause of a cryptic "Bad file name or
    ' number" error further down, so fail loudly and specifically here.
    If Len(inputDir) = 0 Then
        Err.Raise vbObjectError + 1, , "Input directory is empty. Check seedOrder!B1 (or DEF_INPUT_DIR in the code)."
    End If
    If Not FolderExists(inputDir) Then
        Err.Raise vbObjectError + 1, , "Input directory does not exist or the path is invalid:" & vbLf & _
            "[" & inputDir & "]" & vbLf & vbLf & "Check seedOrder!B1 (or DEF_INPUT_DIR in the code) for stray quotes/spaces."
    End If
    If Len(emptyForm) = 0 Then
        Err.Raise vbObjectError + 2, , "Template file name is empty. Check seedOrder!B2 (or DEF_EMPTY_FORM in the code)."
    End If

    pdfPath = inputDir & "ProformaInvoice_" & orderNo & "_source.pdf"
    outPath = inputDir & "InventoryOrder_" & orderNo & ".xlsx"
    logoPath = inputDir & LOGO_FILE

    If Not FileExists(pdfPath) Then Err.Raise vbObjectError + 3, , "Input PDF not found:" & vbLf & "[" & pdfPath & "]"
    If Not FileExists(inputDir & emptyForm) Then Err.Raise vbObjectError + 4, , "Template not found:" & vbLf & "[" & inputDir & emptyForm & "]"

    '--- 1. read the invoice table from the PDF -----------------------------
    nItems = ParseInvoice(pdfPath, items)
    If nItems = 0 Then Err.Raise vbObjectError + 3, , "No 'Pepper' line items were found in the PDF."

    '--- 1b. look up Lot / Germination / Weight-per-1000 per Variety in the
    '        variety inventory workbook (Inventory\<invVarietyFile>.xlsx) -----
    ReDim lookups(1 To nItems)
    If Len(invVarietyFile) = 0 Or Not FolderExists(invDir) Then
        warnings = warnings & "Inventory lookup skipped - invDir/invVarietyFile not configured " & _
            "(seedOrder!B6/B7). Lot / Germination % / Weight-per-1000 left blank." & vbLf
    Else
        Dim invPath As String
        invPath = invDir & invVarietyFile
        If LCase$(Right$(invPath, 5)) <> ".xlsx" And LCase$(Right$(invPath, 5)) <> ".xlsm" Then invPath = invPath & ".xlsx"

        If Not FileExists(invPath) Then
            warnings = warnings & "Inventory workbook not found:" & vbLf & "[" & invPath & "]" & vbLf
        Else
            Set invWb = Workbooks.Open(FileName:=invPath, ReadOnly:=True, AddToMru:=False)
            Set invWs = FindSheetLike(invWb, invSheetName)
            If invWs Is Nothing Then
                warnings = warnings & "Sheet '" & invSheetName & "' not found in " & invVarietyFile & vbLf
            Else
                Dim nFoc As Long
                For i = 1 To nItems
                    If IsFocSample(items(i).variety) Then
                        ' FOC / trial sample (e.g. "SMKO92 FOC for trials") - per spec, these are
                        ' deliberately NOT looked up; Lot/Germination%/Weight stay blank.
                        nFoc = nFoc + 1
                    Else
                        lookups(i) = LookupInventory(invWs, items(i).variety)
                        If Not lookups(i).Found Then
                            warnings = warnings & "No inventory match for variety '" & items(i).variety & "'" & vbLf
                        End If
                    End If
                Next i
                If nFoc > 0 Then notes = notes & nFoc & " FOC/trial line(s) - inventory lookup intentionally skipped." & vbLf
            End If
            invWb.Close SaveChanges:=False
            Set invWb = Nothing
        End If
    End If

    '--- 2. new workbook from the template form sheet ----------------------
    Set tplWb = Workbooks.Open(FileName:=inputDir & emptyForm, ReadOnly:=True, AddToMru:=False)
    tplWb.Worksheets(1).Copy                      ' -> creates a new 1-sheet workbook
    Set outWb = ActiveWorkbook
    Set ws = outWb.Worksheets(1)
    ws.Name = "Seed Order"
    tplWb.Close SaveChanges:=False
    Set tplWb = Nothing

    '--- 3a. English only : left-to-right, strip Hebrew text + letterhead --
    On Error Resume Next
    ws.DisplayRightToLeft = False
    outWb.Windows(1).DisplayRightToLeft = False
    On Error GoTo Fail
    Anglicise ws
    ClearHeaderFooter ws

    '--- 3b. logo picture in a cell --------------------------------------
    InsertLogo ws, logoPath

    '--- 4a. header block ----------------------------------------------------
    ws.Range("A5").Value = "Internal order created by " & ORDERER_NAME
    ws.Range("A6").Value = "Internal Order Creation: " & Format$(Now, "dd/mm/yyyy, hh:nn:ss")
    On Error Resume Next
    ws.Range("A7:I7").Merge
    On Error GoTo Fail
    ws.Range("A7").Value = "Order Due Date: " & Format$(DateAdd("m", 1, Date), "d/mmm/yyyy")
    ws.Range("A9").Value = "Country: " & COUNTRY_NAME
    ws.Range("J6").Value = "Order no.: " & orderNo

    '--- 4b. add the two extra header columns not present on the HE&EN form
    ws.Cells(12, C_LABELS).Copy
    ws.Range(ws.Cells(12, C_EXPIRY), ws.Cells(FIRST_DATA_ROW + 11, C_COMMENT)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    ws.Cells(12, C_EXPIRY).Value = "Expiration date"
    ws.Cells(12, C_COMMENT).Value = "Comments"

    '--- 4c. one row per invoice line -------------------------------------
    For i = 1 To nItems
        r = FIRST_DATA_ROW + i - 1
        If i > 1 Then EnsureRowStyle ws, r          ' template only pre-formats 12 rows

        ws.Cells(r, C_NUM).Value = i                                  ' #  (counter from 1)
        ws.Cells(r, C_SPECIES).Value = items(i).Species               ' "Pepper"
        ws.Cells(r, C_VARIETY).Value = items(i).variety               ' SMK.... to end of description
        ws.Cells(r, C_LOT).Value = lookups(i).Lot                     ' Inventory: first matching Lot number
        ws.Cells(r, C_GERM).Value = lookups(i).GermPct                ' Inventory col M, last 3 chars
        ws.Cells(r, C_QTY).Value = items(i).quantity                  ' invoice Quantity
        ws.Cells(r, C_NPACK).Value = items(i).quantity                ' invoice Quantity (per spec, same value)
        ws.Cells(r, C_INPACK).Value = V_SEEDS_IN_PACK
        If lookups(i).Found Then
            ws.Cells(r, C_W1000).Value = lookups(i).Weight1000        ' Inventory col I
        Else
            ws.Cells(r, C_W1000).Value = ""
        End If
        ws.Cells(r, C_ADDPCT).Value = V_ADDED_PCT

        ' Formula columns.
        ' Net weight = number of packages x weight per 1000 seeds (both now
        ' real inventory-backed values) - per the updated spec (was hard-coded).
        ws.Cells(r, C_NETW).Formula = "=" & ws.Cells(r, C_NPACK).Address(False, False) & "*" & ws.Cells(r, C_W1000).Address(False, False)
        ' Remaining relationships unchanged from the original spec (=B26*B27 etc.,
        ' i.e. against the label rows of the seedOrder sheet; reproduced here
        ' against the real output cells).
        ws.Cells(r, C_ADDW).Formula = "=" & ws.Cells(r, C_NETW).Address(False, False) & "*" & ws.Cells(r, C_ADDPCT).Address(False, False)
        ws.Cells(r, C_TOTW).Formula = "=" & ws.Cells(r, C_ADDW).Address(False, False) & "+" & ws.Cells(r, C_NETW).Address(False, False)
        ws.Cells(r, C_TOTPACK).Formula = "=" & ws.Cells(r, C_NPACK).Address(False, False)
        ws.Cells(r, C_FINSEED).Formula = "=" & ws.Cells(r, C_INPACK).Address(False, False) & "*" & ws.Cells(r, C_ADDPCT).Address(False, False)
        ' spec literally says  Final weight = C24*C24  -> in-package seeds squared.
        ws.Cells(r, C_FINW).Formula = "=" & ws.Cells(r, C_INPACK).Address(False, False) & "*" & ws.Cells(r, C_INPACK).Address(False, False)

        ws.Cells(r, C_LABELS).Value = ""     ' Labels   - blank
        If lookups(i).HasExpiry Then
            ws.Cells(r, C_EXPIRY).Value = lookups(i).ExpiryDate   ' Inventory col M, first 5 chars -> last day of that month
            ws.Cells(r, C_EXPIRY).NumberFormat = "dd/mm/yyyy"
        Else
            ws.Cells(r, C_EXPIRY).Value = ""                      ' blank for FOC/trial lines or no match
        End If
        ws.Cells(r, C_COMMENT).Value = ""    ' Comments  - blank

        ' Row 1 of the template ships bold + oversized + a hard-coded red cell
        ' (a template artifact) - strip that back to the plain, non-bold,
        ' Calibri 11, black style every other row already has (leave column
        ' A's bold as-is; that's the template's normal "#" column style on
        ' every row).
        If i = 1 Then
            With ws.Range(ws.Cells(r, C_SPECIES), ws.Cells(r, C_COMMENT)).Font
                .Bold = False
                .Color = RGB(0, 0, 0)
                .Name = "Calibri"
                .Size = 11
            End With
        End If

        ' Quantity / Number of packages / Seeds-in-package / Net weight are
        ' shown in red - per spec, on every row (including row 1).
        HighlightRed ws.Cells(r, C_QTY)
        HighlightRed ws.Cells(r, C_NPACK)
        HighlightRed ws.Cells(r, C_INPACK)
        HighlightRed ws.Cells(r, C_NETW)
    Next i

    '--- clear the template's leftover pre-numbered rows below our data ---
    For r = FIRST_DATA_ROW + nItems To FIRST_DATA_ROW + 11
        ws.Cells(r, C_NUM).ClearContents
    Next r

    '--- Species / Variety / Lot (B:D) left-aligned instead of centered ---
    ws.Range(ws.Cells(12, C_SPECIES), ws.Cells(FIRST_DATA_ROW + 11, C_LOT)).HorizontalAlignment = xlLeft

    '--- tidy the two added columns --------------------------------------
    ws.Columns(C_EXPIRY).ColumnWidth = 14
    ws.Columns(C_COMMENT).ColumnWidth = 24
    On Error Resume Next
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(FIRST_DATA_ROW + 11, C_COMMENT)).Address
    On Error GoTo Fail

    '--- 5. save + close ---------------------------------------------------
    outWb.SaveAs FileName:=outPath, FileFormat:=xlOpenXMLWorkbook
    outWb.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Dim doneMsg As String
    doneMsg = "Done." & vbLf & vbLf & nItems & " item(s) written to:" & vbLf & outPath
    If Len(notes) > 0 Then doneMsg = doneMsg & vbLf & vbLf & "Notes:" & vbLf & notes
    If Len(warnings) > 0 Then doneMsg = doneMsg & vbLf & vbLf & "Warnings:" & vbLf & warnings
    MsgBox doneMsg, vbInformation
    Exit Sub

Fail:
    Dim msg As String
    msg = "CreateSeedOrder failed:" & vbLf & vbLf & Err.description
    On Error Resume Next
    If Not invWb Is Nothing Then invWb.Close SaveChanges:=False
    If Not tplWb Is Nothing Then tplWb.Close SaveChanges:=False
    If Not outWb Is Nothing Then outWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox msg, vbCritical
End Sub


'================================================================================
'  0. parameters - read B1:B8 of the seedOrder sheet if present
'================================================================================
Private Sub ReadParams(ByRef inputDir As String, ByRef emptyForm As String, ByRef orderNo As String, _
                        ByRef invDir As String, ByRef invVarietyFile As String, ByRef invSheetName As String)
    Dim sh As Worksheet
    inputDir = DEF_INPUT_DIR
    emptyForm = DEF_EMPTY_FORM
    orderNo = DEF_ORDER_NO
    invDir = DEF_INV_DIR
    invVarietyFile = DEF_INV_FILE
    invSheetName = DEF_INV_SHEET

    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("seedOrder")
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub

    On Error Resume Next
    Dim v1 As String, v2 As String, v6 As String, v7 As String, v8 As String
    v1 = SanitizeCellText(CStr(sh.Range("B1").Value))
    v2 = SanitizeCellText(CStr(sh.Range("B2").Value))
    v6 = SanitizeCellText(CStr(sh.Range("B6").Value))
    v7 = SanitizeCellText(CStr(sh.Range("B7").Value))
    v8 = SanitizeCellText(CStr(sh.Range("B8").Value))
    If Len(v1) > 0 Then inputDir = v1
    If Len(v2) > 0 Then emptyForm = v2
    If Len(v6) > 0 Then invDir = v6
    If Len(v7) > 0 Then invVarietyFile = v7
    If Len(v8) > 0 Then invSheetName = v8

    ' order number: pull the run number out of the inputFileName cell (B3)
    ' e.g. ProformaInvoice_90476_source.pdf
    Dim m As Object, re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d{4,6})"
    If re.Test(CStr(sh.Range("B3").Value)) Then
        Set m = re.Execute(CStr(sh.Range("B3").Value))
        orderNo = m(0).SubMatches(0)
    End If
    On Error GoTo 0
End Sub


'================================================================================
'  1b. variety inventory lookup
'================================================================================

' A workbook's sheet name is matched loosely: trimmed, case-insensitive
' ("Inventory Management" vs. the real "Inventory management").
Private Function FindSheetLike(ByVal wb As Workbook, ByVal wanted As String) As Worksheet
    Dim s As Worksheet, w As String
    w = LCase$(Trim$(wanted))
    For Each s In wb.Worksheets
        If LCase$(Trim$(s.Name)) = w Then
            Set FindSheetLike = s
            Exit Function
        End If
    Next s
    ' fall back to the only sheet, or the first one containing "inventory"
    If wb.Worksheets.Count = 1 Then
        Set FindSheetLike = wb.Worksheets(1)
        Exit Function
    End If
    For Each s In wb.Worksheets
        If InStr(1, LCase$(s.Name), "inventory") > 0 Then
            Set FindSheetLike = s
            Exit Function
        End If
    Next s
End Function

' FOC / trial samples (e.g. "SMKO92 FOC for trials", "SMK81FOC for trials")
' are, per spec, never looked up in the Inventory workbook - only the plain
' commercial lines (currently SMKO61 / SMK3 / SMKY23) are.
Private Function IsFocSample(ByVal variety As String) As Boolean
    IsFocSample = (InStr(1, variety, "FOC", vbTextCompare) > 0)
End Function

' Normalises a variety code for comparison against the Inventory sheet's
' group codes (column A), which are written with inconsistent spacing
' ("SMK 3", "SMKO61").
Private Function BaseVarietyCode(ByVal v As String) As String
    BaseVarietyCode = UCase$(Replace(v, " ", ""))
End Function

' Column A of "Inventory management" groups lot rows under a variety code
' (populated only on that group's first row); column B = Lot number,
' column I = Weight per thousand seeds (Gr), column M = Germination validity,
' e.g. "10/25- 93%" - the spec wants the last 3 characters ("93%") as the
' Germination %, and the first 5 characters ("10/25", i.e. MM/YY) turned into
' an actual Expiration date (the last day of that month).
Private Function LookupInventory(ByVal invWs As Worksheet, ByVal varietyCode As String) As TInvLookup
    Const COL_VARIETY As Long = 1   ' A
    Const COL_LOT     As Long = 2   ' B
    Const COL_WEIGHT  As Long = 9   ' I
    Const COL_GERM    As Long = 13  ' M

    Dim target As String, r As Long, lastRow As Long, cellA As String
    target = BaseVarietyCode(varietyCode)
    lastRow = invWs.Cells(invWs.Rows.Count, COL_VARIETY).End(xlUp).Row

    Dim result As TInvLookup
    For r = 1 To lastRow
        cellA = CStr(invWs.Cells(r, COL_VARIETY).Value)
        If Len(cellA) > 0 Then
            If BaseVarietyCode(cellA) = target Then
                result.Found = True
                result.Lot = CStr(invWs.Cells(r, COL_LOT).Value)
                result.Weight1000 = invWs.Cells(r, COL_WEIGHT).Value
                Dim germRaw As String
                germRaw = Trim$(CStr(invWs.Cells(r, COL_GERM).Value))
                If Len(germRaw) >= 3 Then
                    result.GermPct = Right$(germRaw, 3)
                Else
                    result.GermPct = germRaw
                End If

                ' first 5 chars, e.g. "10/25" -> MM=10, YY=25 -> last day of Oct 2025
                If Len(germRaw) >= 5 Then
                    Dim first5 As String, mm As String, yy As String
                    first5 = Left$(germRaw, 5)
                    mm = Left$(first5, 2)
                    yy = Right$(first5, 2)
                    If Mid$(first5, 3, 1) = "/" And IsNumeric(mm) And IsNumeric(yy) Then
                        If CInt(mm) >= 1 And CInt(mm) <= 12 Then
                            result.ExpiryDate = DateSerial(2000 + CInt(yy), CInt(mm) + 1, 0)  ' last day of MM/20YY
                            result.HasExpiry = True
                        End If
                    End If
                End If

                LookupInventory = result
                Exit Function
            End If
        End If
    Next r
    ' not found - Found stays False, other fields stay blank/empty
    LookupInventory = result
End Function


'================================================================================
'  formatting helper - red font for the columns the spec flags as red
'================================================================================
Private Sub HighlightRed(ByVal c As Range)
    c.Font.Color = RGB(255, 0, 0)
End Sub


'================================================================================
'  1. PDF -> line items
'      Primary: Excel's PDF connector (modPepperLines.ReadPepperLinesFromPdf).
'      Word's conversion of the Invoice4U PDFs exposes only the footer text,
'      so the text route below is kept only as a fallback.
'================================================================================
Private Function ParseInvoice(ByVal pdfPath As String, ByRef items() As TItem) As Long
    Dim lines As Collection, it As Variant
    Dim k As Long, n As Long
    Dim pdfErr As String

    On Error Resume Next
    Set lines = ReadPepperLinesFromPdf(pdfPath)
    If Err.Number <> 0 Then pdfErr = Err.description
    On Error GoTo 0

    If Not lines Is Nothing Then
        If lines.Count > 0 Then
            ReDim items(1 To lines.Count)
            For Each it In lines
                k = k + 1
                items(k).Species = "Pepper"
                ' "Pepper SMKO92 FOC for trials" -> "SMKO92 FOC for trials"
                items(k).variety = Trim$(Mid$(it(PL_DESCRIPTION), Len("Pepper") + 1))
                items(k).quantity = it(PL_QUANTITY)
            Next it
            ParseInvoice = k
            Exit Function
        End If
    End If

    ' Fallback: sidecar .txt / Python / Word text
    n = ParseInvoiceText(pdfPath, items)
    If n = 0 And Len(pdfErr) > 0 Then
        Err.Raise vbObjectError + 11, , "Reading the PDF with Excel's PDF connector failed:" & vbLf & pdfErr
    End If
    ParseInvoice = n
End Function

'--- fallback: text route ----------------------------------------------------
'      text is obtained by, in order of preference:
'        a) a plain-text sidecar  <same name>.txt  next to the PDF
'        b) Python + pypdf  (helper script  pdf_to_text.py , same folder as this file
'           OR the input directory) - fast and reliable
'        c) Word automation  (last resort - can be slow / may need the one-time
'           "convert PDF" confirmation)
'================================================================================
Private Function ParseInvoiceText(ByVal pdfPath As String, ByRef items() As TItem) As Long
    Dim raw As String, txt As String
    raw = GetInvoiceText(pdfPath)

    ' normalise Word table/cell marks and tabs to single spaces
    txt = raw
    txt = Replace(txt, Chr(7), " ")     ' cell mark
    txt = Replace(txt, Chr(9), " ")     ' tab
    txt = Replace(txt, Chr(11), vbLf)   ' line break
    txt = Replace(txt, Chr(160), " ")   ' nbsp

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    ' Pepper <description starting SMK...> <qty> <price> US$ ...
    re.Pattern = "Pepper\s+(SMK.*?)\s+(\d[\d,]*)\s+\d[\d,]*\s+US\$"

    Dim matches As Object, mm As Object
    Set matches = re.Execute(txt)

    Dim n As Long
    n = matches.Count

    ' Fallback: some Word conversions put every table cell on its own line.
    ' Flatten newlines to spaces and retry (still bounded by "US$" so it is safe).
    If n = 0 Then
        txt = Replace(Replace(txt, vbCr, " "), vbLf, " ")
        Set matches = re.Execute(txt)
        n = matches.Count
    End If

    If n = 0 Then
        ParseInvoiceText = 0
        Exit Function
    End If

    ReDim items(1 To n)
    Dim k As Long
    k = 0
    For Each mm In matches
        k = k + 1
        items(k).Species = "Pepper"
        items(k).variety = Trim$(mm.SubMatches(0))                 ' SMK... to end of description
        items(k).quantity = CDbl(Replace(mm.SubMatches(1), ",", ""))
    Next mm

    ParseInvoiceText = n
End Function

Private Function GetInvoiceText(ByVal pdfPath As String) As String
    Dim s As String

    ' a) sidecar .txt --------------------------------------------------------
    s = Left$(pdfPath, InStrRev(pdfPath, ".")) & "txt"
    If FileExists(s) Then
        GetInvoiceText = ReadTextFile(s)
        Exit Function
    End If

    ' b) Python + pypdf -----------------------------------------------------
    s = PdfTextViaPython(pdfPath)
    If Len(s) > 0 Then
        GetInvoiceText = s
        Exit Function
    End If

    ' c) Word -------------------------------------------------------------
    s = PdfTextViaWord(pdfPath)
    If Len(s) > 0 Then
        GetInvoiceText = s
        Exit Function
    End If

    Err.Raise vbObjectError + 10, , _
        "Could not extract text from the PDF (Python/pypdf and Word both failed)." & vbLf & vbLf & _
        "Work-around: open " & Mid$(pdfPath, InStrRev(pdfPath, "\") + 1) & " in a PDF reader," & vbLf & _
        "save / export it as plain text with the same name and a .txt extension" & vbLf & _
        "into the same folder, then run the macro again."
End Function

'--- (b) Python helper -------------------------------------------------------
Private Function PdfTextViaPython(ByVal pdfPath As String) As String
    Dim helper As String, outFile As String, cmd As String
    Dim sh As Object, rc As Long

    helper = FindHelperScript()
    If helper = "" Then Exit Function

    outFile = Environ$("TEMP") & "\seedorder_pdf_" & Format$(Now, "yyyymmddhhnnss") & ".txt"

    On Error Resume Next
    Set sh = CreateObject("WScript.Shell")

    Dim launchers As Variant, j As Long
    launchers = Array("python", "py -3", "python3")
    For j = LBound(launchers) To UBound(launchers)
        cmd = "cmd /c " & launchers(j) & " """ & helper & """ """ & pdfPath & """ """ & outFile & """"
        rc = sh.Run(cmd, 0, True)             ' hidden, wait
        If FileExists(outFile) Then Exit For
    Next j
    On Error GoTo 0

    If FileExists(outFile) Then
        PdfTextViaPython = ReadTextFile(outFile)
        On Error Resume Next
        Kill outFile
        On Error GoTo 0
    End If
End Function

Private Function FindHelperScript() As String
    Dim p As String
    On Error Resume Next

    If Len(ThisWorkbook.path) > 0 Then
        p = ThisWorkbook.path & "\pdf_to_text.py"
        If FileExists(p) Then FindHelperScript = p: Exit Function
    End If

    p = DEF_INPUT_DIR & "\pdf_to_text.py"
    If FileExists(p) Then FindHelperScript = p: Exit Function

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("seedOrder")
    If Not sh Is Nothing Then
        p = SanitizeCellText(CStr(sh.Range("B1").Value))
        If Len(p) > 0 Then
            If Right$(p, 1) <> "\" Then p = p & "\"
            p = p & "pdf_to_text.py"
            If FileExists(p) Then FindHelperScript = p
        End If
    End If
    On Error GoTo 0
End Function

'--- (c) Word last resort ---------------------------------------------------
Private Function PdfTextViaWord(ByVal pdfPath As String) As String
    Const msoAutomationSecurityForceDisable As Long = 3
    Dim wd As Object, doc As Object
    On Error GoTo CleanFail
    Set wd = CreateObject("Word.Application")
    wd.Visible = False
    On Error Resume Next
    wd.DisplayAlerts = 0
    wd.AutomationSecurity = msoAutomationSecurityForceDisable
    wd.Options.ConfirmConversions = False
    On Error GoTo CleanFail

    Set doc = wd.Documents.Open(FileName:=pdfPath, ConfirmConversions:=False, _
                                ReadOnly:=True, AddToRecentFiles:=False, Format:="PDF Files")
    PdfTextViaWord = doc.Content.text
    doc.Close SaveChanges:=False
    wd.Quit
    Set doc = Nothing: Set wd = Nothing
    Exit Function

CleanFail:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    If Not wd Is Nothing Then wd.Quit
    Set doc = Nothing: Set wd = Nothing
    PdfTextViaWord = ""
End Function

Private Function ReadTextFile(ByVal path As String) As String
    Dim f As Integer, s As String
    On Error GoTo Bad
    f = FreeFile
    Open path For Input As #f
    If LOF(f) > 0 Then s = Input$(LOF(f), f)
    Close #f
    ReadTextFile = s
    Exit Function
Bad:
    On Error Resume Next
    Close #f
    Err.Raise vbObjectError + 20, , "Could not read text file:" & vbLf & "[" & path & "]" & vbLf & Err.description
End Function

'--- path helpers: wrap Dir$ so a malformed path returns False instead of
'    raising the cryptic "Bad file name or number" from deep inside the code --
Private Function FileExists(ByVal p As String) As Boolean
    Dim r As String
    On Error Resume Next
    Err.Clear
    If Len(p) > 0 Then r = Dir$(p)
    FileExists = (Err.Number = 0) And (Len(r) > 0)
    On Error GoTo 0
End Function

Private Function FolderExists(ByVal p As String) As Boolean
    Dim r As String, chk As String
    On Error Resume Next
    Err.Clear
    chk = p
    If Right$(chk, 1) = "\" Then chk = Left$(chk, Len(chk) - 1)
    If Len(chk) > 0 Then r = Dir$(chk, vbDirectory)
    FolderExists = (Err.Number = 0) And (Len(r) > 0)
    On Error GoTo 0
End Function

' Strips characters that are easy to paste in by accident from a planning
' sheet (smart/straight quotes, tabs, newlines) and trims whitespace.
' ChrW$ is needed for the curly quotes: Chr$ only accepts 0-255 and raised
' an error that ReadParams' On Error Resume Next silently swallowed, so the
' seedOrder!B1:B8 values were never used.
Private Function SanitizeCellText(ByVal s As String) As String
    s = Replace(s, Chr$(34), "")       ' straight double quote
    s = Replace(s, ChrW$(8220), "")    ' left curly double quote
    s = Replace(s, ChrW$(8221), "")    ' right curly double quote
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    SanitizeCellText = Trim$(s)
End Function


'================================================================================
'  3a. English only - remove Hebrew characters from every text cell, then
'      re-label the header block with clean English.
'================================================================================
Private Sub Anglicise(ByVal ws As Worksheet)
    Dim c As Range, s As String
    For Each c In ws.UsedRange
        If VarType(c.Value) = vbString Then
            s = StripHebrew(CStr(c.Value))
            If s <> CStr(c.Value) Then c.Value = s
        End If
    Next c

    With ws
        .Range("A3").Value = "Seed Order Form"
        .Range("M3").Value = "Make sure bag is totally empty"
        .Range("M4").Value = ""
        .Range("N4").Value = ""
        .Range("P3").Value = "Warning:"

        .Cells(12, C_NUM).Value = "#"
        .Cells(12, C_SPECIES).Value = "Species"
        .Cells(12, C_VARIETY).Value = "Variety / Catalog Code"
        .Cells(13, C_LOT).Value = "Lot"
        .Cells(12, C_GERM).Value = "Germination %"
        .Cells(12, C_QTY).Value = "Total no. of seeds"
        .Cells(14, C_QTY).Value = "TH"
        .Cells(12, C_NPACK).Value = "Seeds"
        .Cells(13, C_NPACK).Value = "No. of packets"
        .Cells(14, C_NPACK).Value = "UNIT"
        .Cells(13, C_INPACK).Value = "Seeds in packet"
        .Cells(14, C_INPACK).Value = "UNIT"
        .Cells(13, C_W1000).Value = "1000 weight"
        .Cells(14, C_W1000).Value = "Gr"
        .Cells(13, C_NETW).Value = "Weight"
        .Cells(14, C_NETW).Value = "Gr"
        .Cells(12, C_ADDPCT).Value = "Added Seeds"
        .Cells(13, C_ADDPCT).Value = "%"
        .Cells(13, C_ADDW).Value = "Weight"
        .Cells(14, C_ADDW).Value = "Gr"
        .Cells(12, C_TOTW).Value = "Total"
        .Cells(13, C_TOTW).Value = "Weight"
        .Cells(14, C_TOTW).Value = "Gr"
        .Cells(13, C_TOTPACK).Value = "Packets"
        .Cells(14, C_TOTPACK).Value = "UNIT"
        .Cells(12, C_FINSEED).Value = "Final Package"
        .Cells(13, C_FINSEED).Value = "Seeds in packet"
        .Cells(14, C_FINSEED).Value = "UNIT"
        .Cells(13, C_FINW).Value = "Weight in packet"
        .Cells(14, C_FINW).Value = "Gr"
        .Cells(12, C_LABELS).Value = "Labels"
        .Cells(14, C_LABELS).Value = "Color"
    End With
End Sub

Private Function StripHebrew(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        If code >= 1424 And code <= 1535 Then
            ' Hebrew block -> drop
        Else
            out = out & ch
        End If
    Next i
    ' collapse whitespace / newlines
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "[\s\r\n]+"
    out = Trim$(re.Replace(out, " "))
    ' trim leftover separators from the ends
    Do While Len(out) > 0 And InStr("/\-_"" '", Left$(out, 1)) > 0
        out = Trim$(Mid$(out, 2))
    Loop
    Do While Len(out) > 0 And InStr("/\-_"" '", Right$(out, 1)) > 0
        out = Trim$(Left$(out, Len(out) - 1))
    Loop
    StripHebrew = out
End Function


'================================================================================
'  3a. remove the legacy Hebrew header / footer graphic
'================================================================================
Private Sub ClearHeaderFooter(ByVal ws As Worksheet)
    On Error Resume Next
    With ws.PageSetup
        .LeftHeader = "": .CenterHeader = "": .RightHeader = ""
        .LeftFooter = "": .CenterFooter = "": .RightFooter = ""
    End With
    On Error GoTo 0
End Sub


'================================================================================
'  3b. drop the breedx logo picture into cell A1
'================================================================================
Private Sub InsertLogo(ByVal ws As Worksheet, ByVal logoPath As String)
    If Not FileExists(logoPath) Then
        MsgBox "Logo picture not found - skipping:" & vbLf & "[" & logoPath & "]", vbExclamation
        Exit Sub
    End If
    Dim shp As Shape
    Set shp = ws.shapes.AddPicture(FileName:=logoPath, LinkToFile:=False, _
                                   SaveWithDocument:=True, _
                                   Left:=ws.Range("A1").Left + 2, Top:=ws.Range("A1").Top + 2, _
                                   Width:=-1, Height:=-1)
    With shp
        .LockAspectRatio = msoTrue
        .Height = Application.CentimetersToPoints(1.9)
        .Placement = xlMoveAndSize
        .Name = "logo"
    End With
End Sub


'================================================================================
'  copy row formatting when the invoice has more than the 12 pre-built rows
'================================================================================
Private Sub EnsureRowStyle(ByVal ws As Worksheet, ByVal r As Long)
    If r <= FIRST_DATA_ROW + 11 Then Exit Sub          ' template already styled
    ws.Rows(FIRST_DATA_ROW + 11).Copy
    ws.Rows(r).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub
