Attribute VB_Name = "Module1"
Option Explicit
Sub selectFile()
    Dim preFileName
    Dim defaultFolderName
    
    preFileName = ThisWorkbook.Sheets(1).Cells(1, 3)
    defaultFolderName = ThisWorkbook.Sheets(1).Cells(2, 3)

    With Application.FileDialog(msoFileDialogOpen)
        If preFileName <> "" Then
            .InitialFileName = preFileName
            ThisWorkbook.Sheets(1).Cells(1, 3).ClearContents
        End If
        .Filters.Clear
        .Filters.Add "&img.plist", "*.plist"
        If .Show = True Then
            If InStr(.SelectedItems(1), "&img.plist") > 0 Then
                ThisWorkbook.Sheets(1).Cells(1, 3) = .SelectedItems(1)
            Else
                MsgBox ("Select '&img.plist' file")
                Exit Sub
            End If
        End If
    End With

End Sub
Sub loadImgPlist()
    Dim myDom As MSXML2.DOMDocument60
    Dim myNodeList As IXMLDOMNodeList
    Dim myNode As IXMLDOMNode
    Dim myChildNode As IXMLDOMNode
    Dim i As Integer
    Dim imgPlistPath
    Dim startRow
    Dim maxRow
    Dim subSeqCount
    Dim mainCategoryCount
    Dim subCategoryCount
    Dim array1 As Variant
    Dim myNode2
    
    With Sheets(1)
        imgPlistPath = .Cells(1, 3)
        If Dir(imgPlistPath) = "" Then
            MsgBox (imgPlistPath & " doesn't exist")
            Exit Sub
        End If
        Set myDom = New MSXML2.DOMDocument60
        With myDom
            .SetProperty "ProhibitDTD", False
            .async = False
            .resolveExternals = False
            .validateOnParse = False
            .Load xmlSource:=imgPlistPath
        End With
        Set myNodeList = myDom.SelectNodes("/plist")
        startRow = 20
        .Range(.Cells(startRow, 1), .Cells(1048576, 4)).Clear
        i = startRow
        subSeqCount = 0
        mainCategoryCount = 0
        subCategoryCount = 0
        For Each myNode In myNodeList
            array1 = Split(myNode.ChildNodes(0).Text, " ")
            For Each myNode2 In array1
                Select Case myNode2
                Case "mainCategory", "subCategory", "countStoredImages", "imageFile"
                    Select Case myNode2
                    Case "mainCategory"
                        .Cells(i, 1) = mainCategoryCount * 100
                        mainCategoryCount = mainCategoryCount + 1
                        subCategoryCount = 0
                    Case "subCategory"
                        .Cells(i, 1) = 1 + mainCategoryCount * 100 + subCategoryCount * 10
                        subCategoryCount = subCategoryCount + 1
                    Case "countStoredImages"
                        .Cells(i, 1) = 2 + mainCategoryCount * 100 + subCategoryCount * 10
                    Case "imageFile"
                        .Cells(i, 1) = 3 + mainCategoryCount * 100 + subCategoryCount * 10
                    End Select
                    .Cells(i, 2) = myNode2
                    subSeqCount = 0
                Case "items", "images"
                    subSeqCount = 0
                Case Else
                    subSeqCount = subSeqCount + 1
                    If subSeqCount = 1 Then
                        .Cells(i, 3) = myNode2
                        i = i + 1
                    Else
                        .Cells(i - 1, 3) = .Cells(i - 1, 3) & " " & myNode2
                    End If
                End Select
            Next
        Next
        'sort
        maxRow = .Cells(1048576, 1).End(xlUp).Row
        .Sort.SortFields.Clear
        .Sort.SortFields.Add2 Key:=.Range(.Cells(startRow, 1), .Cells(maxRow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange Range(Cells(startRow, 1), Cells(maxRow, 3))
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        'remove subCategory which has no photos
        For i = startRow To maxRow
            If .Cells(i, 2) = "countStoredImages" And .Cells(i, 3) = 0 Then
                .Range(.Cells(i - 1, 1), .Cells(i, 3)).ClearContents
            End If
        Next i
        .Sort.SortFields.Clear
        .Sort.SortFields.Add2 Key:=.Range(.Cells(startRow, 1), .Cells(maxRow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange Range(Cells(startRow, 1), Cells(maxRow, 3))
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        'remove mainCategory which has no subCategories
        maxRow = .Cells(1048576, 1).End(xlUp).Row
        For i = startRow To maxRow
            If .Cells(i, 2) = "mainCategory" Then
                If .Cells(i + 1, 1) = "" Or (.Cells(i + 1, 1) - .Cells(i, 1)) >= 100 Then
                    .Range(.Cells(i, 1), .Cells(i, 3)).ClearContents
                End If
            End If
        Next i
        .Sort.SortFields.Clear
        .Sort.SortFields.Add2 Key:=.Range(.Cells(startRow, 1), .Cells(maxRow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange Range(Cells(startRow, 1), Cells(maxRow, 3))
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
    
    MsgBox ("Completed")
End Sub
Sub unzipFile()
    Dim psCommand As String
    Dim WSH As Object
    Dim result As Integer
    Dim zipFilePath As String
    Dim toFolderPath
    Dim posFld
    
    With Sheets(1)
        zipFilePath = Replace(.Cells(1, 3), "&img.plist", ".zip")
        If Dir(zipFilePath) = "" Then
            MsgBox (zipFilePath & " doesn't exist")
            Exit Sub
        End If
        posFld = InStrRev(.Cells(1, 3), "\")
        toFolderPath = Mid(.Cells(1, 3), 1, posFld - 1)
        
        Set WSH = CreateObject("WScript.Shell")
        
        zipFilePath = Replace(zipFilePath, " ", "' '")
        zipFilePath = Replace(zipFilePath, "(", "'('")
        zipFilePath = Replace(zipFilePath, ")", "')'")
        toFolderPath = Replace(toFolderPath, " ", "' '")
        toFolderPath = Replace(toFolderPath, "(", "'('")
        toFolderPath = Replace(toFolderPath, ")", "')'")
        
        psCommand = "powershell -NoProfile -ExecutionPolicy Unrestricted Expand-Archive -Path """ & zipFilePath & """ -DestinationPath """ & toFolderPath & """ -Force"
        result = WSH.Run(psCommand, WindowStyle:=0, WaitOnReturn:=True)
    End With
    Set WSH = Nothing
    
    If result = 0 Then
        MsgBox ("Completed")
    Else
        MsgBox ("Failed to unzip file!!")
    End If

End Sub
Sub writeWordFile()
    Dim defaultPath
    Dim WORD As Object
    Dim DOC As Object
    Dim FSO As Object
    Dim PIC As Object
    Dim imageSize
    Dim FolderName
    Dim maxRow
    Dim startRowSheet1
    Dim imageFileCount
    Dim thisFilePath
    Dim thisExtension
    Dim wordFileName
    Dim i
    Dim realSize, picSize
    
    defaultPath = ThisWorkbook.Path
    Set WORD = CreateObject("Word.Application")
    WORD.Visible = True
    Set DOC = WORD.Documents.Open(defaultPath & "\" & "_new.doc")
    Set FSO = CreateObject("Scripting.FileSystemObject")
    imageFileCount = 0
    startRowSheet1 = 20
    imageSize = ThisWorkbook.Sheets(1).Cells(13, 2)
    realSize = 0
    
    With WORD.Selection
        FolderName = Replace(ThisWorkbook.Sheets(1).Cells(1, 3), "&img.plist", "") & "\"
        maxRow = ThisWorkbook.Sheets(1).Cells(1048576, 1).End(xlUp).Row
        For i = startRowSheet1 To maxRow
            Select Case ThisWorkbook.Sheets(1).Cells(i, 2)
            Case "imageFile"
                imageFileCount = imageFileCount + 1
                thisFilePath = FolderName & ThisWorkbook.Sheets(1).Cells(i, 3)
                thisExtension = LCase(FSO.GetExtensionName(thisFilePath))
                Select Case thisExtension
                Case "jpg", "jpeg"
                    If imageFileCount Mod 2 = 0 Then
                        'none
                    Else
                        If realSize + picSize >= 733.5 Then
                            .InsertBreak (wdPageBreak)
                            realSize = 0
                        End If
                    End If
                    Set PIC = DOC.Bookmarks("\EndOfDoc").Range.InlineShapes.AddPicture(thisFilePath)
                    With PIC
                        .LockAspectRatio = msoTrue
                        If .Width > .Height Then
                            .LockAspectRatio = msoTrue
                            .Width = imageSize
                            .Height = imageSize * 3 / 4
                        Else
                            .LockAspectRatio = msoTrue
                            .Width = imageSize * 3 / 4
                            .Height = imageSize
                        End If
                    End With
                    .EndKey Unit:=wdStory
                    .TypeText Text:=CStr(" ")
                    If imageFileCount Mod 2 = 0 Then
                        .TypeParagraph
                        .TypeParagraph
                        realSize = realSize + 12.25
                    Else
                        picSize = PIC.Height
                        realSize = realSize + picSize
                    End If
                    Set PIC = Nothing
                End Select
            Case "mainCategory"
                If ThisWorkbook.Sheets(1).Cells(i - 1, 2) = "imageFile" And imageFileCount Mod 2 = 1 Then
                    .TypeParagraph
                    .TypeParagraph
                    realSize = realSize + 12.25
                End If
                If realSize + 12.25 * 2 + picSize >= 733.5 Then
                    .InsertBreak (wdPageBreak)
                    realSize = 0
                End If
                .EndKey Unit:=wdStory
                .TypeText Text:=CStr(replaceLabel(ThisWorkbook.Sheets(1).Cells(i, 3))) & " :"
                .TypeParagraph
                realSize = realSize + 12.25
            Case "subCategory"
                If ThisWorkbook.Sheets(1).Cells(i - 1, 2) = "imageFile" And imageFileCount Mod 2 = 1 Then
                    .TypeParagraph
                    .TypeParagraph
                    realSize = realSize + 12.25
                End If
                If realSize + picSize >= 733.5 Then
                    .InsertBreak (wdPageBreak)
                    realSize = 0
                End If
                imageFileCount = 0
                .EndKey Unit:=wdStory
                .TypeText Text:=CStr("- " & replaceLabel(ThisWorkbook.Sheets(1).Cells(i, 3)))
                .TypeParagraph
                realSize = realSize + 12.25
            End Select
        Next i
    End With
    Set FSO = Nothing
    wordFileName = Replace(ThisWorkbook.Sheets(1).Cells(1, 3), "&img.plist", "") & ".doc"
    DOC.SaveAs Filename:=wordFileName
    DOC.Close
    Set DOC = Nothing
    WORD.Quit
    Set WORD = Nothing
    
    MsgBox ("Completed")

End Sub
Function replaceLabel(ByVal target)
    Dim maxRow, i, j
    Dim initialTarget
    Dim findStr, replaceStr
    Dim arr1 As Variant
    
    
    maxRow = Sheets("replace").Cells(1048576, 1).End(xlUp).Row
    initialTarget = target
    For i = 2 To maxRow
        findStr = Sheets("replace").Cells(i, 1)
        replaceStr = Sheets("replace").Cells(i, 2)
        arr1 = Split(target, "(")
        target = ""
        For j = 0 To UBound(arr1)
            If InStr(arr1(j), findStr) > 0 And InStr(replaceStr, "*") > 0 Then
                If findStr = "-" And IsNumeric(Mid(arr1(j), InStr(arr1(j), findStr) + 1, 1)) Then
                    target = Trim(arr1(j))
                Else
                    If j = 0 Then
                        target = Replace(replaceStr, "*", Trim(arr1(j)))
                    Else
                        target = target & " (" & Replace(replaceStr, "*", Trim(arr1(j)))
                    End If
                End If
            Else
                If j = 0 Then
                    target = Replace(Trim(arr1(j)), findStr, replaceStr)
                Else
                    target = target & " (" & Replace(Trim(arr1(j)), findStr, replaceStr)
                End If
            End If
        Next j
    Next i
    replaceLabel = target
End Function
