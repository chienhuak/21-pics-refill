Attribute VB_Name = "Module1"
Option Explicit


Public Sub main()
Dim i, j, k As Integer
Dim MyColumn As Integer
Dim MyFlag As Boolean
Dim Temp As String
Dim Mystr As String
Dim strLength As Integer
Dim Pos As Integer
Dim MyPICsLen As Integer
Dim MyTableLen As Integer
Dim mystring(1000) As String
Dim LTEPicsCount As Integer
Dim NRPicsCount As Integer
Dim Temp1 As String
Dim MySpec As String
Dim Myitem As String
Dim MyCountS As Integer
Dim MyCountD As Integer
Dim MyErrorMessage As String

Set glSWWBK = ActiveWorkbook
Set glMain = glSWWBK.Sheets("Main")

Call LoadPICS(glMain.Range("Source"))
Set SWBK = ActiveWorkbook
Call LoadPICS(glMain.Range("Dest"))
Set DWBK = ActiveWorkbook

MyCountS = SWBK.Sheets.Count
MyCountD = DWBK.Sheets.Count
For i = 1 To MyCountS
    MyColumn = MyCol(SWBK.Sheets(i).Name)
    If MyColumn <> 0 Then
        For j = 1 To MyCountD
            If SWBK.Sheets(i).Name = DWBK.Sheets(j).Name Then
                MyCheck = True
                Call Fill_Pics(SWBK.Sheets(i).Name, MyColumn)
                Exit For
            Else
                MyCheck = False
                If j = MyCountD Then
                    MyErrorMessage = SWBK.Sheets(i).Name & " is not found." & vbCr & vbLf
                    SWBK.Sheets(i).Copy before:=DWBK.Sheets(1)
                End If
            End If
        Next j
    Else
        For k = 1 To MyCountD
            If SWBK.Sheets(i).Name = DWBK.Sheets(k).Name Then
                MyFlag = True
                Exit For
            Else
                MyFlag = False
            End If
        Next k
        If MyFlag = False Then
            SWBK.Sheets(i).Copy before:=DWBK.Sheets(1)
        End If
    End If
Next
If MyCheck = False Then
    MsgBox "Strange"
End If
Temp = Format(Date, "YYYYMMDD")
Temp = ActiveWorkbook.Path & "\Tempate_" & Temp & ".xlsx"
MsgBox ("Search is done")
glSWWBK.Save
SWBK.Close
DWBK.Activate
ActiveWorkbook.SaveAs Filename:= _
    Temp, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
ActiveWorkbook.Close




End Sub
Sub CheckEnd(MySheet, MyColumn)

Dim i As Integer

Dim Str1 As String
Dim Str2 As String
Dim Str3 As String
Dim StrTemp As String
Dim Pos As Integer



For i = 1 To 5000
    StrTemp = wb.Sheets(MySheet).Cells(i, MyColumn)
    Str1 = wb.Sheets(MySheet).Cells(i, 2)
    Str2 = wb.Sheets(MySheet).Cells(i + 1, 2)
    Str3 = wb.Sheets(MySheet).Cells(i + 2, 2)
    If Len(Str1) + Len(Str2) + Len(Str3) <> 0 Then
        StrTemp = wb.Sheets(MySheet).Cells(i, MyColumn - 1)
        If Len(Str1) <> 0 And Len(StrTemp) <> 0 Then
            If wb.Sheets(MySheet).Cells(i, MyColumn) = "" Then
                MyCheck = True
                Exit For
            Else
                MyCheck = False
            End If
        Else
            GoTo Here2
        End If
    Else
        Exit For
    End If
Here2:

Next i

End Sub

Public Sub Fill_Pics(Spec_Name As String, Col As Integer)
Dim i As Integer
Dim j As Integer
Dim Str1 As String
Dim Str2 As String
Dim Str3 As String
Dim StrTemp As String
Dim StrS1 As String
Dim StrS2 As String
Dim StrS3 As String
Dim Pos As Integer

For i = 13 To 5000
    StrTemp = SWBK.Sheets(Spec_Name).Cells(i, Col)
    StrS1 = SWBK.Sheets(Spec_Name).Cells(i, 2)
    StrS2 = SWBK.Sheets(Spec_Name).Cells(i + 1, 2)
    StrS3 = SWBK.Sheets(Spec_Name).Cells(i + 2, 2)
    Str1 = DWBK.Sheets(Spec_Name).Cells(i, 2)
    Str2 = DWBK.Sheets(Spec_Name).Cells(i + 1, 2)
    Str3 = DWBK.Sheets(Spec_Name).Cells(i + 2, 2)
    If Len(StrS1) + Len(StrS2) + Len(StrS3) <> 0 Then
'        StrTemp = wb.Sheets(MySheet).Cells(i, MyColumn - 1)
        If Len(StrS1) <> 0 And Len(StrTemp) <> 0 Then
            Call FindString(DWBK.Sheets(Spec_Name), DWBK.Sheets(Spec_Name).Range("b1:b5000"), StrS1)
            If MyCheck = True Then
                SWBK.Sheets(Spec_Name).Cells(i, Col).Copy
                DWBK.Sheets(Spec_Name).Cells(MyFound, Col).Select
                ActiveSheet.Paste
                GoTo Here2
            End If
        Else
            GoTo Here2
        End If
    Else
        Exit For
    End If
Here2:

Next i

End Sub

