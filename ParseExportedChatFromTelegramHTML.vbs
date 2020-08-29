'29/08/2020 Alexey Zapolskiy aka alex; <alexeyzapolskiy at gmail dot com>
'Парсер файла.html - экспорта истории telegam-канала
'Сохраните как Parse.vbs и запустите

'Sub Parse()
    Set fso = CreateObject("Scripting.FileSystemObject") 'Dim fso As New FileSystemObject
'Dim FileNameIn As String, FileNameOut As String
'Dim strBody As String, result As String, objMatches As Object
    FileNameIn = "c:\Users\alex\Downloads\Telegram Desktop\ChatExport_2020-08-29\messages.html"
    FileNameOut = "c:\temp\ChatExport_2020-08-29.txt"
    Set tsInput = CreateObject("ADODB.Stream") 'Dim tsInput As New ADODB.Stream
    tsInput.Open
    tsInput.Type = 2 'text
    tsInput.Charset = "utf-8"
    tsInput.LoadFromFile FileNameIn
    strBody = tsInput.ReadText(fso.GetFile(FileNameIn).Size)
    Set objRegExp = CreateObject("VBScript.RegExp") 'Dim objRegExp As New RegExp
    objRegExp.Pattern = "<div class=""text"">\s*?(.+)\s*?<\/div>"
    objRegExp.IgnoreCase = True
    objRegExp.Global = True

    Set objMatches = objRegExp.Execute(strBody)
'Dim tsOutput As TextStream
    Set tsOutput = fso.CreateTextFile(FileNameOut, True)
'Dim ii As Integer
    For ii = 0 To objMatches.Count - 1
        tsOutput.WriteLine (objMatches(ii).SubMatches(0))
    Next
    MsgBox "Done"
'End Sub