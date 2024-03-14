Attribute VB_Name = "Module3"
Option Explicit

Sub ⅷ떠1()
Attribute ⅷ떠1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ⅷ떠1 ⅷ떠
'

'
    
    Range("D8").Copy
'    Selection.Copy
    Range("E8").Select
    ActiveSheet.Paste
End Sub
Sub ⅷ떠2()
Attribute ⅷ떠2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ⅷ떠2 ⅷ떠
'

'
    Windows("pics_template_2021-12-13_07-23-42.xlsx").Activate
    ChDir "D:\Project folder\TomPics"
    ActiveWorkbook.SaveAs Filename:= _
        "D:\Project folder\TomPics\pics_template_2021-12-13_07-23-42-1.xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End Sub
