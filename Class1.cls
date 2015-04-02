VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents appevent As Application
Attribute appevent.VB_VarHelpID = -1

Private Sub appevent_SheetActivate(ByVal Sh As Object)

    'MsgBox "change worksheet"
    Call SaveData
    
End Sub

Sub SaveData()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wbTarget As Workbook
    Dim wbThis As Workbook
    Dim zip As String
    Dim streetName As String
    Dim currentRow As Integer
    
    Set wbThis = ActiveWorkbook
    Set wbTarget = Workbooks.Open("H:\docs\VBA\Data\DataPoints.xlsx")
    zip = wbThis.Worksheets("Exp").Range("I7").Value
    streetName = wbThis.Worksheets("Exp").Range("E5").Value
    currentRow = FindProperty(zip, streetName, wbTarget)
    
    If currentRow <> 0 Then
        Call UpdateData(wbThis, wbTarget, currentRow)
    Else
        Call CopyData(wbThis, wbTarget)
    End If
    
    wbTarget.Save
    wbTarget.Close
    
    Set wbTarget = Nothing
    Set wbThis = Nothing

End Sub

Function FindProperty(zip As String, streetName As String, wbTarget As Workbook) As Integer

Dim address As String

FindProperty = 0

'Find a zip code and then check if the property has
'the same street address.
'Returns the row number of the first cell in the targeted range

With wbTarget.Worksheets(1).Range("I:I")
    Set c = .Find(zip, LookIn:=xlValues)
    If Not c Is Nothing Then
        firstAddress = c.address
        Do
            address = c.address(RowAbsolute:=False, ColumnAbsolute:=False)
            Range(address).Activate
            ActiveCell.Offset(-2, -4).Activate
            If ActiveCell.Value = streetName Then
                FindProperty = ActiveCell.Row - 4
            End If
            Set c = .FindNext(c)
        Loop While Not c Is Nothing And c.address <> firstAddress
    End If
End With

End Function

Sub UpdateData(wbThis As Workbook, wbTarget As Workbook, rowNumber As Integer)

    'Copies only the numbers from the active workbook to the target workbook
    
    'MsgBox "update"
    
    Application.CutCopyMode = False
        
    wbTarget.Sheets(1).Range("C" & rowNumber + 11).Value = wbThis.Sheets("Exp").Range("E22").Value
    wbTarget.Sheets(1).Range("K" & rowNumber).Activate
    wbTarget.Sheets(1).Range(ActiveCell, ActiveCell.Offset(20, 16)).Value = _
        wbThis.Sheets("Rent").Range("B3:S23").Value
    wbTarget.Sheets(1).Range("L" & rowNumber + 21).Value = wbThis.Sheets("Rent").Range("E42").Value
    
    wbTarget.Sheets(1).Range("AF" & rowNumber & ":AF" & rowNumber + 79).Value = wbThis.Sheets("Hist").Range("N10:N89").Value
    wbTarget.Sheets(1).Range("AF" & rowNumber + 79 & ":AF" & rowNumber + 79).NumberFormat = "#,##0.000"

    Application.CutCopyMode = False
End Sub

Sub CopyData(wbThis As Workbook, wbTarget As Workbook)
    
    'Copies both names and values from the active workbook to the target workbook
    
    'MsgBox "save"
    
    Application.CutCopyMode = False
    
    Dim rowNumber As Integer
    rowNumber = FindNextAvailableRange(wbTarget)
    
    wbTarget.Sheets(1).Range("B" & rowNumber & ":I" & rowNumber + 9).Value = wbThis.Sheets("Exp").Range("B1:I10").Value
    wbTarget.Sheets(1).Range("B" & rowNumber + 11).Value = wbThis.Sheets("Exp").Range("B22").Value
    wbTarget.Sheets(1).Range("C" & rowNumber + 11).Value = wbThis.Sheets("Exp").Range("E22").Value
    wbTarget.Sheets(1).Range("K" & rowNumber).Activate
    wbTarget.Sheets(1).Range(ActiveCell, ActiveCell.Offset(20, 16)).Value = _
        wbThis.Sheets("Rent").Range("B3:S23").Value
    wbTarget.Sheets(1).Range("K" & rowNumber + 21).Value = wbThis.Sheets("Rent").Range("B42").Value
    wbTarget.Sheets(1).Range("L" & rowNumber + 21).Value = wbThis.Sheets("Rent").Range("E42").Value
    wbTarget.Sheets(1).Range("AB" & rowNumber & ":AB" & rowNumber + 78).Value = wbThis.Sheets("Hist").Range("B10:B88").Value
    wbTarget.Sheets(1).Range("AF" & rowNumber & ":AF" & rowNumber + 79).Value = wbThis.Sheets("Hist").Range("N10:N89").Value
    wbTarget.Sheets(1).Range("AB" & rowNumber + 79).Value = wbThis.Sheets("Hist").Range("A89").Value
    wbTarget.Sheets(1).Range("AF" & rowNumber + 79 & ":AF" & rowNumber + 79).NumberFormat = "#,##0.000"

    Application.CutCopyMode = False
End Sub

Function FindNextAvailableRange(wbTarget As Workbook) As Integer

'Returns the first available row number for the next chunk of data

Dim i As Integer
i = 1

Do While Not IsEmpty(wbTarget.Sheets(1).Cells(i, 2))
    i = i + 81
Loop
    
FindNextAvailableRange = i

End Function

