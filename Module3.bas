Attribute VB_Name = "Module3"
Option Explicit

Sub ����1()
Attribute ����1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����1 ����
'

'
    
    Range("D8").Copy
'    Selection.Copy
    Range("E8").Select
    ActiveSheet.Paste
End Sub
Sub ����2()
Attribute ����2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����2 ����
'

'
    Windows("pics_template_2021-12-13_07-23-42.xlsx").Activate
    ChDir "D:\Project folder\TomPics"
    ActiveWorkbook.SaveAs Filename:= _
        "D:\Project folder\TomPics\pics_template_2021-12-13_07-23-42-1.xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End Sub
