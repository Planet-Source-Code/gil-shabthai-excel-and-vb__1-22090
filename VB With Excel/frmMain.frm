VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Make Excel File"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn              As ADODB.Connection
Dim rs              As ADODB.Recordset

Dim FieldsName()    As String
Dim FieldsValue()   As String

Dim DBfields()      As String
Dim DBvaluse()      As String

Dim Counter         As Integer
Dim FieldsCounter   As Integer

Private Sub cmdExcel_Click()
 Call MakeExcel
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset

'--- Open ADODB Connection ---
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Db.mdb;Persist Security Info=False"

'--- Get A Recordset - Customer Table
rs.Open "select * from Customer", cn, adOpenDynamic, adLockOptimistic

'--- DBfields - Array To Hold Fields Name ---
'--- Get The Array From MakeArrayFields Function ---
DBfields = MakeArrayFields

'--- DBvaluse - Array To Hold Values In Table ---
'--- Get The Array From MakeArrayValues Function ---
DBvaluse = MakeArrayValues

End Sub
Private Function MakeArrayValues() As Variant
' ******************************************************************************
' Routine:           MakeArrayValues
' Description:       Get Values Of Table To An Array
' Created by:        gil
' Machine:           GIL
' Date-Time:         02/07/00-16:49:07
' Last modification: last_modification_info_here
' ******************************************************************************
On Error GoTo ErrHandler

    Dim RowCounter      As Integer
    Dim RowPlace        As Integer 'Row Number In Table
    Dim ColumnPlace     As Integer
    
    '--- Count Rows In Table ---
    Do Until rs.EOF = True
        RowCounter = RowCounter + 1
        rs.MoveNext
    Loop
    
    
    RowPlace = 1 ' Start In Row Number 1
    rs.MoveFirst ' Move REcordset To First The Record
    
    '--- Declare Array Size ---
    '--- First Rows Number Then Columns Number ---
    ReDim FieldsValue(0 To RowCounter - 1, 0 To FieldsCounter - 1)
    
    '--- Do This As The Number Of Rows ---
    For RowPlace = 0 To RowCounter - 1
        '--- Do This As The Number Of Fields ---
        For ColumnPlace = 0 To FieldsCounter - 1
            '--- Fill Array With Value ---
            FieldsValue(RowPlace, ColumnPlace) = rs.Fields(ColumnPlace).Value
        Next ColumnPlace
        '--- Move Recordset For The Next Record ---
        rs.MoveNext
    Next RowPlace
    
    '--- Function Return The full Array ---
    MakeArrayValues = FieldsValue()

Exit Function
ErrHandler:
    MsgBox Err.Number & vbCrLf & Err.Description
End Function

Private Function MakeArrayFields() As Variant
' ******************************************************************************
' Routine:           MakeArrayFields
' Description:       Get Fields Name To An Array
' Created by:        gil
' Machine:           GIL
' Date-Time:         02/07/00-17:05:27
' Last modification: last_modification_info_here
' ******************************************************************************
On Error GoTo ErrHandler

    Dim Counter As Integer
    
    '--- Count Fields In Table ---
    FieldsCounter = rs.Fields.Count
    
    '--- Declare Array Size ---
    '--- Size Will Be As FieldsCounter ---
    ReDim FieldsName(0 To FieldsCounter - 1)
    
    Counter = 0
    
    '--- Do This As The Number Of Fields ---
    For Counter = 0 To FieldsCounter - 1
        '--- Fill Array With Fields Name ---
        FieldsName(Counter) = rs.Fields.Item(Counter).Name
    Next Counter
    
    '--- Function Return The full Array ---
    MakeArrayFields = FieldsName()
    
Exit Function
ErrHandler:
    MsgBox Err.Number & vbCrLf & Err.Description
End Function

Public Sub MakeExcel()
' ******************************************************************************
' Routine:           MakeExcel
' Description:       Make The Excel File use DLL
' Created by:        gil
' Machine:           GIL
' Date-Time:         02/07/00-17:12:11
' Last modification: last_modification_info_here
' ******************************************************************************

On Error GoTo ErrHandler

    Dim ExcelFileName   As String
    Dim PrintToExcel    As ToExcelFile.ExcelFile
    
    Set PrintToExcel = New ToExcelFile.ExcelFile
    
    '--- File Name To Save As Excel File
    ExcelFileName = "xxxxx.xls"
    
    Screen.MousePointer = vbHourglass
    
    '--- Call Function From DLL File ---
    '--- To Make The Excel File ---
    '--- It Get 2 Array And Excel File Name ---
    '--- One Array For Fields Name ---
    '--- Second Array With Values ---
    Call PrintToExcel.MakeExcelFile(DBfields(), DBvaluse(), ExcelFileName)
    
    Screen.MousePointer = vbDefault
 
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number
End Sub

