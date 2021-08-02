Attribute VB_Name = "ParseInfo"
Public className As Collection


Public Sub main()
    Dim mainDict As Object, body As String
    Set className = New Collection
    Set mainDict = CreateObject("Scripting.Dictionary")
    mainDict.Add "class", parseMatrix(0)
    mainDict.Add "user", parseUser(0)
    body = ToJson(mainDict)
    save (body)
End Sub


 
 
Public Function containString(ByVal par As String, ByVal chi As String)
For Counter = 1 To Len(chi)
    If InStr(par, Mid(chi, Counter, 1)) = 0 Then
        containString = False
        Exit Function
    End If
Next
containString = True
End Function


Public Sub save(ByVal body As String)
    Dim FileSelected As String
    FileSelected = Application.GetSaveAsFilename(InitialFileName:="export", FileFilter:="JSON, *.json")
    
    If Not FileSelected <> "False" Then
    MsgBox "작업이 취소되었어요."
    Exit Sub
    End If
    
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2 'Specify stream type - we want To save text/string data.
    fsT.Charset = "utf-8" 'Specify charset For the source text data.
    fsT.Open 'Open the stream And write binary data To the object
    fsT.WriteText body
    fsT.SaveToFile FileSelected, 2 'Save binary data To disk
    
    MsgBox "저장이 완료되었어요."
    
End Sub

Function CToJson(ByVal col As Collection) As String
    Dim result As String
    result = "["
    For Each i In col
        result = result & IIf(Len(result) > 1, ",", "")
        result = result & ToJson(i)
    Next
    result = result & "]"
    CToJson = result
End Function


Function ToJson(ByVal dict As Object) As String
    Dim key As Variant, result As String, value As String
    result = "{"
    For Each key In dict.Keys
        result = result & IIf(Len(result) > 1, ",", "")

        If TypeName(dict(key)) = "Dictionary" Then
            value = ToJson(dict(key))
            ToJson = value
        ElseIf TypeName(dict(key)) = "Collection" Then
            value = CToJson(dict(key))
            ToJson = value
        Else
            value = """" & dict(key) & """"
        End If

        result = result & """" & key & """:" & value & ""
    Next key
    result = result & "}"
    ToJson = result
End Function

Public Function ExistsInCollection(ByVal col As Collection, ByVal key As Variant) As Boolean
     For Each i In col
        If i = key Then
            ExistsInCollection = True
            Exit Function
        End If
    Next
    ExistsInCollection = False
End Function

Public Function parseMatrix(id)
    Const blockHeight = 5
    Const beginX = 2
    Const beginY = 4
    Const classCount = 7
    Const day = 5
    
    Set parseMatrix = New Collection
    
    Dim sheet As Worksheet
    Set sheet = Worksheets("안내자료")
    sheet.Activate
    
    For dayI = 0 To day - 1
        For classI = 0 To classCount - 1
            For i = 0 To blockHeight - 1
            
            If Cells(beginY + 5 * dayI + i, beginX + 3 * classI).value = "" Then
                GoTo Continue
            End If
            
            If Cells(beginY + 5 * dayI + i, beginX + 3 * classI).MergeCells Then
                GoTo Continue
            End If
            cn = Trim(Cells(beginY + 5 * dayI + i, beginX + 3 * classI + 1).value)
            Set oneClassDict = CreateObject("Scripting.Dictionary")
            oneClassDict.Add "id", Cells(beginY + 5 * dayI + i, beginX + 3 * classI).value
            oneClassDict.Add "className", cn
            oneClassDict.Add "place", Cells(beginY + 5 * dayI + i, beginX + 3 * classI + 2).value
            oneClassDict.Add "day", dayI
            oneClassDict.Add "time", classI
            parseMatrix.Add oneClassDict
            
            If ExistsInCollection(className, cn) = False Then
                className.Add cn
            End If
            
Continue:
            Next
        Next
    Next
End Function

Public Function getFullName(ByVal shortName As String)
    sName = Replace(shortName, vbCr, Chr(19))
    sName = Replace(sName, vbLf, Chr(19))
    sName = Replace(sName, vbCrLf, Chr(19))
    sName = Split(sName, Chr(19))(0)
    sName = Split(sName, "신청")(0)
    sName = Split(sName, "분반")(0)
    sName = Trim(sName)
    For Each i In className
        If containString(i, sName) Then
            getFullName = i
            Exit Function
        End If
    Next
End Function

Public Function parseUser(id)
    Dim sheet As Worksheet
    Set sheet = Worksheets("수강신청 및 분반")
    sheet.Activate
    
    Const headerY = 9
    Const beginX = 5
    Dim wCount As Integer
    wCount = 0
    Set parseUser = New Collection
    Set classHeader = New Collection
    
    
    While Cells(headerY, wCount * 2 + beginX).value <> ""
        classHeader.Add getFullName(Cells(headerY, wCount * 2 + beginX).value)
        wCount = wCount + 1
    Wend
    
    i = 1
    While Cells(headerY + i, 1).value <> ""
        Set oneUserDict = CreateObject("Scripting.Dictionary")
        Set oneUserClass = New Collection
        
        For j = 0 To wCount - 1
            If Cells(headerY + i, j * 2 + beginX).value <> "" Then
                Set oneClassDict = CreateObject("Scripting.Dictionary")
                oneClassDict.Add "className", classHeader.Item(j + 1)
                oneClassDict.Add "id", Cells(headerY + i, j * 2 + beginX).value
                oneUserClass.Add oneClassDict
            End If
        Next
        
        oneUserDict.Add "name", Cells(headerY + i, 3).value
        oneUserDict.Add "id", Cells(headerY + i, 2).value
        oneUserDict.Add "data", oneUserClass
        parseUser.Add oneUserDict
        i = i + 1
    Wend
End Function

