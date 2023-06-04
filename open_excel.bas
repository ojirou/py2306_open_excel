Attribute VB_Name = "Module1"
Option Explicit
Sub open_excel_test()
Call open_excel("C:\sample.xsx")
End Sub
Sub open_excel(ByVal FilePath As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(FilePath)
End Sub
