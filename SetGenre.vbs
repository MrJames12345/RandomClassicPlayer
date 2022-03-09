
 ' Variables
Set fsObj = CreateObject("Scripting.FileSystemObject")
Set oShell = WScript.CreateObject ("WScript.Shell")
thisFolderPath = fsObj.GetParentFolderName(WScript.ScriptFullName)
Set thisFolder = fsObj.GetFolder(thisFolderPath)
output = ""


 ' IF no argument, show help
If WScript.Arguments.Count = 0 Then
    
    output = HelpMessage()

 ' Else set genre
Else

    genreNum = WScript.Arguments.Item(0)

    ' Get genre folders list
    Set genreFoldersList = thisFolder.SubFolders

    ' If new genre num > num folders, reset to 0
    If Int(genreNum) > int(genreFoldersList.Count) Then
        genreNum = 0
        output = output & "The genre number you entered is not an option." & vbCrLf
    End If

    ' Set new genre num
    Set genreFile = fsObj.OpenTextFile(thisFolder & "\genre.txt", 2)
    genreFile.Write(genreNum)
    genreFile.Close
    Set genreFile = Nothing

    If genreNum = 0 Then
        genreName = "Random Genre"
    Else
        i = 1
        For Each genreFolder in genreFoldersList
        If Int(i) = Int(genreNum) Then
            genreName = genreFolder.Name
        End If
        i = i + 1
        Next
    End If

    output = output & "Your genre is set to " & Int(genreNum) & " (" & genreName & ")."

End If


' Output
Wscript.Echo output








' Functions

Function HelpMessage()

    Set genreFoldersList = thisFolder.SubFolders

    indent = "    "
    outString = vbCrLf & "Set the genre using one of these values:" & vbCrLf & indent & "0 - Random Genre"

    i = 1
    For Each folder in genreFoldersList
        outString = outString & vbCrLf & indent & i & " - " & folder.Name
        i = i + 1
    Next

    HelpMessage = outString

End Function

' Function GetGenreFromFile()

'   Set genreFile = fsObj.OpenTextFile("genre.txt", 1)
'   genreNum = genreFile.ReadAll()
'   genreFile.Close
'   Set genreFile = Nothing
'   GetGenreFromFile = genreNum

' End Function