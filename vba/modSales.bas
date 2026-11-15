Attribute VB_Name = "modSales"
'==============================================================================
' BreedX Peppers - Sales entry
'
' Data layer for frmSalesEntry:
'   * Reads the Products / Company lists from sheet "Table" (columns A / B).
'   * Appends sales records to the last row of sheet "Master":
'       A Prodcut Catalog Number | B Company | C Delivery Date |
'       D Quantity of Seeds      | E Update Time
'   * The primary key of a record is (Product, Company, Delivery Date).
'   * The data workbook is DATA_FILE in DATA_FOLDER. It is used directly when
'     already open, otherwise it is opened from that folder.
'==============================================================================
Option Explicit

' --- data workbook location ---
Public Const DATA_FOLDER As String = "C:\Users\YoavEilenberg\Breedx LTD Dropbox\IT\Sales"
Public Const DATA_FILE As String = "BreedX Peppers sales examples.xlsm"
Public Const SAVE_AFTER_SUBMIT As Boolean = True   ' save after every record (Dropbox sync)

Public Const SHEET_MASTER As String = "Master"
Public Const SHEET_TABLE As String = "Table"

Public Const COL_PRODUCT As Long = 1      ' A
Public Const COL_COMPANY As Long = 2      ' B
Public Const COL_DELIVERY As Long = 3     ' C
Public Const COL_QUANTITY As Long = 4     ' D
Public Const COL_UPDATED As Long = 5      ' E

Public Const QTY_MIN As Long = 0
Public Const QTY_MAX As Long = 1000000

Private Const FIRST_DATA_ROW As Long = 2  ' row 1 holds the headers

'------------------------------------------------------------------------------
' Entry point - assign this macro to a button on the Master sheet.
'------------------------------------------------------------------------------
Public Sub ShowSalesForm()
    On Error GoTo NoData
    DataWorkbook
    On Error GoTo 0

    frmSalesEntry.Show
    Exit Sub

NoData:
    MsgBox Err.Description, vbExclamation, "BreedX Peppers - New Sale"
End Sub

'------------------------------------------------------------------------------
' Returns the sales data workbook: the open copy of DATA_FILE if there is one,
' otherwise DATA_FOLDER\DATA_FILE is opened. Raises an error when not found.
'------------------------------------------------------------------------------
Public Function DataWorkbook() As Workbook
    Dim wb As Workbook
    Dim fullPath As String

    On Error Resume Next
    Set wb = Workbooks(DATA_FILE)
    On Error GoTo 0

    If wb Is Nothing Then
        fullPath = DATA_FOLDER & "\" & DATA_FILE
        If Len(Dir$(fullPath)) = 0 Then
            Err.Raise vbObjectError + 513, "DataWorkbook", _
                      "Sales workbook not found:" & vbLf & fullPath
        End If
        Set wb = Workbooks.Open(fullPath)
    End If

    Set DataWorkbook = wb
End Function

'------------------------------------------------------------------------------
' Optional one-time helper: puts a "New Sale" button on the Master sheet
' that opens the form.
'------------------------------------------------------------------------------
Public Sub CreateLaunchButton()
    Dim ws As Worksheet
    Dim anchor As Range
    Dim btn As Button

    Set ws = DataWorkbook.Worksheets(SHEET_MASTER)
    On Error Resume Next
    ws.Buttons("btnNewSale").Delete
    On Error GoTo 0

    Set anchor = ws.Range("G2")
    Set btn = ws.Buttons.Add(anchor.Left, anchor.Top, 110, 28)
    btn.Name = "btnNewSale"
    btn.Caption = "New Sale..."
    btn.OnAction = "'" & ThisWorkbook.Name & "'!ShowSalesForm"
End Sub

'------------------------------------------------------------------------------
' Fills a combo box with the non-empty, distinct values of a column on the
' Table sheet (header in row 1 is skipped).
'------------------------------------------------------------------------------
Public Sub FillComboFromTable(ByVal cbo As MSForms.ComboBox, ByVal colIndex As Long)
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim item As String
    Dim seen As Object

    Set ws = DataWorkbook.Worksheets(SHEET_TABLE)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    cbo.Clear
    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
    For r = FIRST_DATA_ROW To lastRow
        item = Trim$(CStr(ws.Cells(r, colIndex).Value))
        If Len(item) > 0 Then
            If Not seen.Exists(item) Then
                seen.Add item, True
                cbo.AddItem item
            End If
        End If
    Next r
End Sub

'------------------------------------------------------------------------------
' Validates the quantity text: whole number between QTY_MIN and QTY_MAX.
'------------------------------------------------------------------------------
Public Function TryParseQuantity(ByVal text As String, ByRef qty As Long) As Boolean
    text = Trim$(text)
    If Len(text) = 0 Or Len(text) > 7 Then Exit Function
    If text Like "*[!0-9]*" Then Exit Function

    qty = CLng(text)
    TryParseQuantity = (qty >= QTY_MIN And qty <= QTY_MAX)
End Function

'------------------------------------------------------------------------------
' Returns the Master row numbers whose primary key (Product, Company,
' Delivery Date) matches the given values. Empty collection when none.
'------------------------------------------------------------------------------
Public Function FindRecordRows(ByVal product As String, ByVal company As String, _
                               ByVal deliveryDate As Date) As Collection
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim result As New Collection
    Dim key As Long

    Set ws = DataWorkbook.Worksheets(SHEET_MASTER)
    key = CLng(Int(deliveryDate))
    lastRow = LastDataRow(ws)

    For r = FIRST_DATA_ROW To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, COL_PRODUCT).Value)), product, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(ws.Cells(r, COL_COMPANY).Value)), company, vbTextCompare) = 0 Then
                If DateKey(ws.Cells(r, COL_DELIVERY).Value) = key Then
                    result.Add r
                End If
            End If
        End If
    Next r

    Set FindRecordRows = result
End Function

'------------------------------------------------------------------------------
' Appends a record as the last row of Master and stamps the update time.
' Any rows listed in rowsToReplace (existing records with the same primary
' key) are removed first, so the key stays unique.
' Returns the row number that was written.
'------------------------------------------------------------------------------
Public Function SaveSalesRecord(ByVal product As String, ByVal company As String, _
                                ByVal deliveryDate As Date, ByVal quantity As Long, _
                                Optional ByVal rowsToReplace As Collection) As Long
    Dim ws As Worksheet
    Dim newRow As Long
    Dim i As Long

    Set ws = DataWorkbook.Worksheets(SHEET_MASTER)
    If ws.Parent.ReadOnly Then
        Err.Raise vbObjectError + 514, "SaveSalesRecord", _
                  ws.Parent.FullName & vbLf & "is open read-only (someone else may be using it)."
    End If

    If Not rowsToReplace Is Nothing Then
        ' Delete bottom-up so earlier row numbers stay valid.
        For i = rowsToReplace.Count To 1 Step -1
            ws.Rows(rowsToReplace(i)).Delete
        Next i
    End If

    newRow = LastDataRow(ws) + 1

    With ws.Cells(newRow, COL_PRODUCT)
        .NumberFormat = "@"
        .Value = product
    End With
    With ws.Cells(newRow, COL_COMPANY)
        .NumberFormat = "@"
        .Value = company
    End With
    With ws.Cells(newRow, COL_DELIVERY)
        .NumberFormat = "mm/dd/yyyy"
        .Value = DateSerial(Year(deliveryDate), Month(deliveryDate), Day(deliveryDate))
    End With
    With ws.Cells(newRow, COL_QUANTITY)
        .NumberFormat = "#,##0"
        .Value = quantity
    End With
    With ws.Cells(newRow, COL_UPDATED)
        .NumberFormat = "mm/dd/yyyy hh:mm:ss"
        .Value = Now
    End With

    If SAVE_AFTER_SUBMIT Then ws.Parent.Save

    SaveSalesRecord = newRow
End Function

'------------------------------------------------------------------------------
' Last used row across the record columns (A:E); 1 when only headers exist.
'------------------------------------------------------------------------------
Private Function LastDataRow(ByVal ws As Worksheet) As Long
    Dim c As Long, r As Long, maxRow As Long

    maxRow = FIRST_DATA_ROW - 1
    For c = COL_PRODUCT To COL_UPDATED
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > maxRow Then maxRow = r
    Next c
    LastDataRow = maxRow
End Function

'------------------------------------------------------------------------------
' Converts a Delivery Date cell value (real date or date-looking text, as in
' the baseline row 2) to a whole-day serial number; -1 when not a date.
'------------------------------------------------------------------------------
Private Function DateKey(ByVal v As Variant) As Long
    DateKey = -1
    If IsEmpty(v) Then Exit Function
    If IsDate(v) Then
        DateKey = CLng(Int(CDbl(CDate(v))))
    ElseIf IsNumeric(v) Then
        DateKey = CLng(Int(CDbl(v)))
    End If
End Function
