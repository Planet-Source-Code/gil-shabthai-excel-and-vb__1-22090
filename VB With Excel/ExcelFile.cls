VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExcelFile"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim ExcelSheet      As Excel.Application

Dim LableNo         As Integer
Dim ExcelColNo      As Integer
Dim ExcelCel        As String
Dim ExcelRow        As Integer
Dim ColNoDB         As Integer
Dim RowNoDB         As Integer
Dim LineWidth       As Byte

Dim MakeDir         As FileSystemObject
Dim NumberOfColumns As Integer
Dim Counter         As Integer
Dim BackCounter     As Integer
Dim CaptionString   As String
Dim HeadColName     As String

Dim FieldsCounter   As Integer

Private Const ExcelColumn_B = 98    'Ascii Value For : B


Private Sub FillExcelSheet(ArrayValues() As String, FieldsCount As Integer)
' ******************************************************************************
' Routine:           FillExcelSheet
' Description:       set value to Excel Sheet
' Created by:        gil
' Machine:           GIL
' Date-Time:         25/06/00-16:09:00
' Last modification: last_modification_info_here
' ******************************************************************************

On Error GoTo ErrHandler

ExcelColNo = ExcelColumn_B
ExcelCel = Empty
ColNoDB = 0
RowNoDB = 0
ExcelRow = 2

'--- Fill All Cell(Row) In The First Column ---
'--- Then Fill SEcond Column And So On ---

'--- Move From Column To The Next ---
For ColNoDB = 0 To FieldsCount
       '--- Move From One Row To The Next ---
       For RowNoDB = LBound(ArrayValues) To UBound(ArrayValues)
            ExcelCel = UCase(Chr(ExcelColNo)) & 2 + ExcelRow
            ExcelSheet.Range(ExcelCel).Value = ArrayValues(RowNoDB, ColNoDB)
            ExcelRow = ExcelRow + 1
        Next RowNoDB
    ExcelRow = 2
    ExcelColNo = ExcelColNo + 1
    ExcelCel = Empty
Next ColNoDB


Exit Sub
ErrHandler:
     MsgBox Err.Number & vbCrLf & Err.Description
End Sub


Private Sub FillExcelLables(ArrayFields() As String)
' ******************************************************************************
' Routine:           FillExcelLables
' Description:       Set Lables To The Excel Sheet Columns
' Created by:        gil
' Machine:           GIL
' Date-Time:         25/06/00-16:08:13
' Last modification: last_modification_info_here
' ******************************************************************************

On Error GoTo ErrHandler

ExcelSheet.Workbooks.Add

BackCounter = 0

For LableNo = LBound(ArrayFields) To UBound(ArrayFields)
    '--- Make A Point To Excel Cel ---
    '--- Always In Line 3  ---
    ExcelCel = UCase(Chr(ExcelColNo)) & 3
    
    '--- Get Field Caption From DB ---
    HeadColName = ArrayFields(LableNo)
    
    '--- Insert Field Caption To Excel Cel ---
    ExcelSheet.Range(ExcelCel).Value = HeadColName
    
    '--- Increase Excel Column No ---
    ExcelColNo = ExcelColNo + 1
    BackCounter = BackCounter + 1
Next LableNo

    '--- Keep Number Of Fields For Later Use ---
    FieldsCounter = UBound(ArrayFields)

Exit Sub
ErrHandler:
     MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Public Function MakeExcelFile(FieldsArray() As String, ValuesArray() As String, FileNameToSave As String)
' ******************************************************************************
' Routine:           MakeExcelFile
' Description:       Save Excel File
' Created by:        gil
' Machine:           GIL
' Date-Time:         25/06/00-12:35:21
' Last modification: last_modification_info_here
' ******************************************************************************
    
    On Error GoTo ErrHandler
    
    Dim PathInRegesry As String
    
    Set ExcelSheet = CreateObject("excel.application")
    Set MakeDir = New FileSystemObject


    ExcelColNo = ExcelColumn_B
    LableNo = 0

    '--- get file name and directory name from the user ---
    If MakeDir.FolderExists("C:\ExcelFiles") = False Then
        MakeDir.CreateFolder "C:\ExcelFiles"
    End If
        
    Set MakeDir = Nothing
      
      
    '--- set lable to Excel Sheet ---
    Call FillExcelLables(FieldsArray)

    '--- set value to Excel Sheet
    Call FillExcelSheet(ValuesArray, UBound(FieldsArray))

    '--- Change Excel Sheet View ---
    Call SheetView

    '--- save the file ---
    ExcelSheet.AlertBeforeOverwriting = False
    ExcelSheet.ActiveWorkbook.SaveAs "C:\ExcelFiles\" & FileNameToSave
       
    ExcelSheet.Visible = True
    
    '--- end the Excel processes
    'ExcelSheet.Quit
    
    
    Set ExcelSheet = Nothing
    
    Exit Function
ErrHandler:

    ExcelSheet.Quit
    MsgBox Err.Number & vbCrLf & Err.Description
    Set ExcelSheet = Nothing
    
End Function

Private Function SheetView()
                     
' ******************************************************************************
' Routine:           OutOfBounds
' Description:       Change Sheet View , Font , Color
' Created by:        gil
' Machine:           GIL
' Date-Time:         03/07/00-15:38:40
' Last modification: last_modification_info_here
' ******************************************************************************
On Error GoTo ErrHandler
    
    Dim CellRange As String
    
    '--- Get Range To Change  - Lable Range ---
    CellRange = "B3:" & UCase(Chr(FieldsCounter + ExcelColumn_B)) & "3"

    '--- Change Fonts Property ---
  With ExcelSheet
    .Range(CellRange).Font.Bold = True
    .Range(CellRange).Font.Size = 13
    .Range(CellRange).Font.Color = vbRed
    .Range(CellRange).Font.Italic = True
    .Range(CellRange).Font.Underline = True

    ' --- Make a border -----
    .Range(CellRange).Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range(CellRange).Borders(xlEdgeLeft).Weight = xlMedium
    .Range(CellRange).Borders(xlEdgeLeft).ColorIndex = 32
    
    .Range(CellRange).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(CellRange).Borders(xlEdgeTop).Weight = xlMedium
    .Range(CellRange).Borders(xlEdgeTop).ColorIndex = 32
    
    .Range(CellRange).Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range(CellRange).Borders(xlEdgeBottom).Weight = xlMedium
    .Range(CellRange).Borders(xlEdgeBottom).ColorIndex = 32

    .Range(CellRange).Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range(CellRange).Borders(xlEdgeRight).Weight = xlMedium
    .Range(CellRange).Borders(xlEdgeRight).ColorIndex = 32

    '--- Aligment selection -----------------------------
    .Range(CellRange).HorizontalAlignment = xlRight
    .Range(CellRange).VerticalAlignment = xlBottom
    
    '--- Change Width For All Columns Automatic ---
    .Columns.AutoFit
    
    '--- Change Color Of CellRange - Lables Row ---
    .Range(CellRange).Interior.Color = vbYellow
    
    '--- Change with for (A1:A1)
    .Range("A3").Select
    .Columns("A:A").ColumnWidth = 20
     
  End With
  
Exit Function
ErrHandler:
    ExcelSheet.Quit
    MsgBox Err.Number & vbCrLf & Err.Description
    Set ExcelSheet = Nothing
End Function



