Sub ExportPRD()
Dim d As Document
Set d = Documents.Open("C:\Users\ncucc\Desktop\10-30\???????????_PRD.docx")
d.ExportAsFixedFormat OutputFileName:="C:\Users\ncucc\Desktop\10-30\prd_render.pdf", ExportFormat:=wdExportFormatPDF
d.Close SaveChanges:=False
Application.Quit
End Sub