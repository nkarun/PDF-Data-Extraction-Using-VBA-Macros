Sub PDFEXTRACT(PDF_FileFolderPath As String)

' This macro performs extraction of required fields and updates into respective columns of Template Sheet.
'
    Dim pageNumber As Byte, PDF_FilePath As String, PDF_PagesCount As Byte, page As String, PDF_FileName As String, lastRow1 As Integer, elements2, elements1, Count1, Count2, TotalCount
    Dim sheetTemp As Worksheet
    
    Application.ScreenUpdating = False
    
    Set sheetOutput = Sheets("Template")
    
    'Loop through all PDF files in a folder
    'PDF_FileFolderPath = "C:\Users\PROJECT\ARUN"
    PDF_FileName = Dir(PDF_FileFolderPath & "\*.pdf")
    
    While PDF_FileName <> ""
        IMGPageDetails = ""
        IMGComments = ""
        EXTPageDetails = ""
        EXTComments = ""
        PDF_FilePath = PDF_FileFolderPath & "\" & PDF_FileName
        'MsgBox ("""""" & PDF_FilePath & """""")
        
        PDF_PagesCount = Get_PDF_Pages_Count(PDF_FilePath)
        lastRow1 = sheetOutput.Range("B" & Rows.Count).End(xlUp).row + 1
        
        'MsgBox (PDF_PagesCount)
        
        For pageNumber = 1 To PDF_PagesCount
            If pageNumber > 0 And pageNumber < 10 Then
                page = "Page00" & pageNumber
            ElseIf pageNumber >= 10 And pageNumber < 100 Then
                page = "Page0" & pageNumber
            Else
                page = "Page" & pageNumber
            End If
            
            ExtractFlag = ""
            
            On Error Resume Next
                ActiveWorkbook.Queries(page).Delete
                
            For Each sheetTemp In ThisWorkbook.Worksheets
                If sheetTemp.Name <> "Template" And sheetTemp.Name <> "Controls" Then
                    Application.DisplayAlerts = False
                    sheetTemp.Delete
                    Application.DisplayAlerts = False
                End If
            Next
            
            Call PDF_To_Excel(page, PDF_FilePath, pageNumber)
            Call Extract_Fields(PDF_FileName, pageNumber)
    
        Next pageNumber
            
        'Set the PDF_FileName to the next PDF file
        PDF_FileName = Dir
    Wend
        
    Application.ScreenUpdating = False
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''Below code will export PDF to Excel ''''''''''''''''''''''''''''''''''''''''''''''''''

Function PDF_To_Excel(page, PDF_FilePath, pageNumber)
   
ActiveWorkbook.Queries.Add Name:=page, Formula:= _
    "let" & Chr(13) & "" & Chr(10) & _
    "    Source = Pdf.Tables(File.Contents(""" & PDF_FilePath & """), [Implementation=""1.3""])," & Chr(13) & "" & Chr(10) & _
    "    PageData = Source{[Id=""" & page & """]}[Data]," & Chr(13) & "" & Chr(10) & _
    "    #""Promoted Headers"" = Table.PromoteHeaders(PageData, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
    "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"", List.Transform(Table.ColumnNames(#""Promoted Headers""), each {_, type text}))" & Chr(13) & "" & Chr(10) & _
    "in" & Chr(13) & "" & Chr(10) & _
    "    #""Changed Type"""
        
        Create_QueryTableData_Sheet (page)      
        ActiveWorkbook.Queries(page).Delete
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''Below code will export PDFDDATABLE to Excel ''''''''''''''''''''''''''''''''''''''''''''''''''
Function PDFDDATABLE_To_Excel(page, PDF_FilePath, pageNumber)

ActiveWorkbook.Queries.Add Name:=page, Formula:= _
    "let" & Chr(13) & "" & Chr(10) & _
    "    Source = Pdf.Tables(File.Contents(""" & PDF_FilePath & """), [Implementation=""1.3""])," & Chr(13) & "" & Chr(10) & _
    "    PageData = Source{[Id=""" & page & """]}[Data]," & Chr(13) & "" & Chr(10) & _
    "    FilteredData = Table.SelectRows(PageData, each [ColumnName] <> null)," & Chr(13) & "" & Chr(10) & _
    "    #""Promoted Headers"" = Table.PromoteHeaders(FilteredData, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & _
    "    #""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"", List.Transform(Table.ColumnNames(#""Promoted Headers""), each {_, type text}))" & Chr(13) & "" & Chr(10) & _
    "in" & Chr(13) & "" & Chr(10) & _
    "    #""Changed Type"""

        Create_QueryTableData_Sheet (page)      
        ActiveWorkbook.Queries(page).Delete
End Function
