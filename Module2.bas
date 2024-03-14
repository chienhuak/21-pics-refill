Attribute VB_Name = "Module2"
Option Explicit

Public wb As Workbook
Public SWBK As Workbook
Public DWBK As Workbook
Public glSWWBK As Workbook
Public glMain As Worksheet
Public MyCheck As Boolean
Public MyResult1 As String
Public MyResult2 As String
Public Mycount As Integer
Public MyFound As Integer
Sub FindString(MySheet As Worksheet, c As Range, MySearch As String)
MySheet.Activate
MySheet.Cells(1, 1).Select

With c

    Set c = .Find(MySearch, LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then
        MyFound = c.Row
        MyCheck = True
    Else
        MyCheck = False
    End If
End With
End Sub

Public Sub LoadPICS(MyFileName As String)
    Dim MyPath As String
    Dim i As Integer
    'Detect report exist error.

    MyPath = ActiveWorkbook.Path & "\"
    Workbooks.Open Filename:=MyPath & MyFileName
    Set wb = ActiveWorkbook
    i = wb.Sheets.Count
    For i = 1 To Mycount
        wb.Sheets(i).AutoFilterMode = False
    Next

End Sub

Public Function MyTrim(Str As String) As String
Dim i As Integer
Dim Str1 As String

For i = 1 To Len(Str)
    Str1 = Left(Str, 1)
    Str = Right(Str, Len(Str) - 1)
    If Str1 <> " " Then
        MyTrim = MyTrim & Str1
    End If
Next i

End Function



Public Function MyCol(MyName As String) As Integer

Select Case MyName
    Case Is = "3GPP TR 37.901"
        MyCol = 7
    Case Is = "3GPP TS 26.132 (Features)"
        MyCol = 8
    Case Is = "3GPP TS 31.121"
        MyCol = 7
    Case Is = "3GPP TS 31.124"
        MyCol = 8
    Case Is = "3GPP TS 34.121-2"
        MyCol = 6
    Case Is = "3GPP TS 34.123-2"
        MyCol = 7
    Case Is = "3GPP TS 34.171"
        MyCol = 6
    Case Is = "3GPP TS 34.229-2"
        MyCol = 8
    Case Is = "3GPP TS 36.521-2"
        MyCol = 7
    Case Is = "3GPP TS 36.523-2"
        MyCol = 7
    Case Is = "3GPP TS 37.571-3"
        MyCol = 7
    Case Is = "3GPP TS 38.508-2"
        MyCol = 7
    Case Is = "3GPP TS 51.010-2"
        MyCol = 8
    Case Is = "3GPP TS 51.010-4"
        MyCol = 8
    Case Is = "ETSI TS 102 230-1"
        MyCol = 7
    Case Is = "ETSI TS 102 384"
        MyCol = 7
    Case Is = "ETSI TS 102 694-1"
        MyCol = 7
    Case Is = "ETSI TS 102 695-1"
        MyCol = 7
    Case Is = "GSMA PRD TS.27"
        MyCol = 7
    Case Is = "GSMA SGP.23"
        MyCol = 7
    Case Is = "IMTC 3G-324M (Features)"
        MyCol = 0
    Case Is = "OMA-ETS-DM (ICS)"
        MyCol = 0
    Case Is = "OMA-ETS-FUMO (ICS)"
        MyCol = 0
    Case Is = "OMA-ETS-LPPe-V1_0"
        MyCol = 6
    Case Is = "OMA-ETS-MMS_CON (ICS)"
        MyCol = 0
    Case Is = "OMA-ETS-RCS-CON-V5_x (ICS)"
        MyCol = 0
    Case Is = "OMA-ETS-SUPL-V1 (ICS)"
        MyCol = 6
    Case Is = "OMA-ETS-SUPL-V2 (ICS)"
        MyCol = 6
    Case Is = "OMA-ETS-WCSS (ICS)"
        MyCol = 0
    Case Is = "OMA-ETS-XHTML (ICS)"
        MyCol = 0
    Case Is = "PTCRB AT Command Test Spec. cov"
        MyCol = 0
    Case Is = "PTCRB Bearer Agnostic TTY Test "
        MyCol = 0
    Case Is = "OMA-ETS-SCOMO (ICS)"
        MyCol = 0
    Case Is = "3GPP TS 34.229-3"
        MyCol = 0
    Case Is = "3GPP TS 36.523-3"
        MyCol = 0
    Case Is = "3GPP TS 36.579-4"
        MyCol = 0
    Case Is = "3GPP TS 51.010-3"
        MyCol = 0

End Select
End Function

