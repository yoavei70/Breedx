Attribute VB_Name = "modPepperLines"
'==============================================================================
' Reads the "Pepper" line items from a BreedX proforma invoice PDF.
'
' The PDF is read with Excel's own PDF connector (Power Query, Pdf.Tables -
' the same as Data > Get Data > From File > From PDF). Word is not used:
' Word's conversion of the Invoice4U PDFs does not expose the invoice table
' text to VBA.
'
' The invoice table is:  Code | Description | Quantity | Price | Total
' The Code column is empty and "Pepper ..." is in Description.
' A cell that starts with "Pepper" is a line item; the first stand-alone whole
' number after it (same cell, or the next non-empty cell in the row) is its
' Quantity. Freight and other non-Pepper rows are skipped.
'
' Requires Microsoft 365 / Excel 2021 or later (PDF connector).
'==============================================================================
Option Explicit

' Item layout returned by ReadPepperLinesFromPdf (each item is a Variant array)
Public Const PL_DESCRIPTION As Long = 0   ' "Pepper SMKO92 FOC for trials"
Public Const PL_CODE As Long = 1          ' "SMKO92"
Public Const PL_QUANTITY As Long = 2      ' 10

Private Const TEMP_QUERY As String = "tmpPepperPdf"

'------------------------------------------------------------------------------
' Returns a Collection of Array(Description, Code, Quantity) for every
' "Pepper" line item in the PDF. Empty collection when none.
'------------------------------------------------------------------------------
Public Function ReadPepperLinesFromPdf(ByVal pdfPath As String) As Collection
    Set ReadPepperLinesFromPdf = ParsePepperLines(PdfToText(pdfPath))
End Function

'------------------------------------------------------------------------------
' Loads all tables of the PDF (or, if it has none, the page text) into a
' temporary sheet with Power Query and returns the cells as text:
' one cell per line, an empty line between rows.
'------------------------------------------------------------------------------
Public Function PdfToText(ByVal pdfPath As String) As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim m As String
    Dim data As Variant
    Dim r As Long, c As Long
    Dim txt As String
    Dim errNum As Long, errDesc As String
    Dim prevAlerts As Boolean, prevScreen As Boolean

    If Len(Dir$(pdfPath)) = 0 Then Err.Raise vbObjectError + 600, , "PDF not found:" & vbLf & pdfPath

    Set wb = ThisWorkbook
    prevAlerts = Application.DisplayAlerts
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    RemoveTempQuery wb

    ' Tables of all pages; if the PDF has no detected tables, use the pages.
    m = "let" & vbLf & _
        "  Source = Pdf.Tables(File.Contents(""" & Replace(pdfPath, """", """""") & """))," & vbLf & _
        "  Tbls = Table.SelectRows(Source, each [Kind] = ""Table"")," & vbLf & _
        "  Pick = if Table.RowCount(Tbls) > 0 then Tbls else Table.SelectRows(Source, each [Kind] = ""Page"")," & vbLf & _
        "  Parts = List.Transform(Pick[Data], each Table.DemoteHeaders(_))," & vbLf & _
        "  Combined = Table.Combine(Parts)" & vbLf & _
        "in" & vbLf & _
        "  Combined"

    On Error GoTo Failed
    wb.Queries.Add Name:=TEMP_QUERY, Formula:=m

    Set ws = wb.Worksheets.Add
    With ws.ListObjects.Add(SourceType:=0, _
            Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;" & _
                    "Location=" & TEMP_QUERY & ";Extended Properties=""""", _
            Destination:=ws.Range("A1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & TEMP_QUERY & "]")
        .BackgroundQuery = False
        .Refresh BackgroundQuery:=False
    End With

    data = ws.UsedRange.Value
    If IsArray(data) Then
        For r = 2 To UBound(data, 1)              ' row 1 = Column1, Column2, ...
            For c = 1 To UBound(data, 2)
                If Not IsError(data(r, c)) Then txt = txt & CStr(data(r, c))
                txt = txt & vbLf
            Next c
            txt = txt & vbLf
        Next r
    End If

Failed:
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo -1                              ' leave the error-handler state
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not ws Is Nothing Then ws.Delete
    RemoveTempQuery wb
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    On Error GoTo 0

    If errNum <> 0 Then Err.Raise errNum, "PdfToText", "Reading the PDF failed:" & vbLf & errDesc
    PdfToText = txt
End Function

Private Sub RemoveTempQuery(ByVal wb As Workbook)
    Dim cn As WorkbookConnection
    On Error Resume Next
    For Each cn In wb.Connections
        If InStr(1, cn.Name, TEMP_QUERY, vbTextCompare) > 0 Then cn.Delete
    Next cn
    wb.Queries(TEMP_QUERY).Delete
End Sub

'------------------------------------------------------------------------------
' Pairs each "Pepper ..." piece with the first whole number that follows it.
' Pieces are separated by line feeds; an empty piece ends a table row.
'------------------------------------------------------------------------------
Private Function ParsePepperLines(ByVal text As String) As Collection
    Dim result As New Collection
    Dim pieces() As String
    Dim i As Long, j As Long
    Dim piece As String, description As String, qtyText As String
    Dim re As Object, m As Object

    ' Same piece: "Pepper SMKO61 171 1,300 US$ 222,300.00" -> first stand-alone
    ' number after the description is the quantity ("SMK3" is not stand-alone)
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^(Pepper.*?)\s+(\d[\d,]*)(?=\s|$)"

    text = Replace(text, vbCr, vbLf)
    text = Replace(text, vbTab, vbLf)
    text = Replace(text, ChrW(160), " ")
    pieces = Split(text, vbLf)

    For i = 0 To UBound(pieces)
        piece = Application.WorksheetFunction.Trim(pieces(i))
        If LCase$(piece) Like "pepper*" Then
            description = piece
            qtyText = ""

            If re.Test(piece) Then
                Set m = re.Execute(piece)(0)
                description = m.SubMatches(0)
                qtyText = m.SubMatches(1)
            Else
                ' Quantity = next non-empty piece if it is a whole number
                For j = i + 1 To UBound(pieces)
                    If Len(Trim$(pieces(j))) > 0 Then
                        If IsWholeNumber(pieces(j)) Then qtyText = pieces(j)
                        Exit For
                    End If
                Next j
            End If

            If Len(qtyText) > 0 Then AddItem result, description, qtyText
        End If
    Next i

    Set ParsePepperLines = result
End Function

Private Function IsWholeNumber(ByVal s As String) As Boolean
    s = Replace(Trim$(s), ",", "")
    IsWholeNumber = (Len(s) > 0 And Not s Like "*[!0-9]*")
End Function

Private Sub AddItem(ByVal result As Collection, ByVal description As String, _
                    ByVal qtyText As String)
    description = Application.WorksheetFunction.Trim(description)
    qtyText = Replace(Trim$(qtyText), ",", "")
    result.Add Array(description, CodeFromDescription(description), CDbl(qtyText))
End Sub

'------------------------------------------------------------------------------
' "Pepper SMKO61"                -> "SMKO61"
' "Pepper SMKO92 FOC for trials" -> "SMKO92"
' "Pepper SMK81FOC for trials"   -> "SMK81"
'------------------------------------------------------------------------------
Public Function CodeFromDescription(ByVal description As String) As String
    Dim parts() As String
    Dim code As String

    parts = Split(Application.WorksheetFunction.Trim(description), " ")
    If UBound(parts) < 1 Then Exit Function
    code = parts(1)
    If Len(code) > 3 And UCase$(Right$(code, 3)) = "FOC" Then code = Left$(code, Len(code) - 3)
    CodeFromDescription = code
End Function

'------------------------------------------------------------------------------
' Test: pick a proforma PDF and list the Pepper lines in a message box and
' the Immediate window (Ctrl+G). When nothing is found, the text read from
' the PDF is saved next to it as "<name>_extracted.txt".
' Does not change any workbook (the temporary sheet/query are removed).
'------------------------------------------------------------------------------
Public Sub TestReadPepperLines()
    Dim pdfPath As Variant
    Dim txt As String, msg As String, dumpPath As String
    Dim items As Collection
    Dim it As Variant
    Dim f As Integer

    pdfPath = Application.GetOpenFilename("PDF files (*.pdf),*.pdf", , "Select proforma invoice PDF")
    If VarType(pdfPath) = vbBoolean Then Exit Sub

    On Error GoTo Failed
    txt = PdfToText(CStr(pdfPath))
    Set items = ParsePepperLines(txt)

    msg = items.Count & " Pepper line(s) found:" & vbLf
    For Each it In items
        msg = msg & it(PL_CODE) & "   qty " & it(PL_QUANTITY) & "   (" & it(PL_DESCRIPTION) & ")" & vbLf
        Debug.Print it(PL_CODE), it(PL_QUANTITY), it(PL_DESCRIPTION)
    Next it

    If items.Count = 0 Then
        dumpPath = Left$(pdfPath, InStrRev(pdfPath, ".") - 1) & "_extracted.txt"
        f = FreeFile
        Open dumpPath For Output As #f
        Print #f, Replace(txt, vbLf, vbCrLf)
        Close #f
        msg = msg & vbLf & "The text read from the PDF was saved to:" & vbLf & dumpPath
    End If

    MsgBox msg, vbInformation, "Pepper lines"
    Exit Sub

Failed:
    MsgBox "Error: " & Err.Description, vbExclamation, "Pepper lines"
End Sub
