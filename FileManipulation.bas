Attribute VB_Name = "FileManipulation"
'####################################################################################################################################
'NAME: Public Function hOpenFile(myPhylePath As String) As String
'DESCRIPTION: Opens a file and reads out the contents of that file
'RECEIVES: "myPhylePath" as a string containing the path and filename of the file to be opened
'RETURNS: a string containing the contents of the file
'REQUIRES: NOTHING
'####################################################################################################################################
Public Function hOpenFile(myPhylePath As String) As String
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Dim fs, f
    Dim fileContents As String
    Set fs = CreateObject("Scripting.FileSystemObject") 'Create our filesystemobject
    Set f = fs.OpenTextFile(myPhylePath, ForReading, TristateFalse) 'Open the text stream
    fileContents = f.readall 'Read all of the contents of the file
    f.Close 'Close the file
    hOpenFile = fileContents 'Return the contents of the file
End Function

'####################################################################################################################################
'NAME: Sub hWriteTextFile(whatFile As String, whatText As String)
'DESCRIPTION: Creates a text file (if file already exists, the contents are overwritten) with the string
'             passed in.  If the specified file does not exist, it is created.
'RECEIVES: "whatFile" as a string contains the path and filename of the file to be written;
'          "whatText" as a string contains the text that is to be written to that file
'RETURNS: NOTHING
'REQUIRES: NOTHING
'####################################################################################################################################
Sub hWriteTextFile(whatfile As String, whatText As String)
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CreateTextFile whatfile
    Set f = fs.OpenTextFile(whatfile, ForWriting, TristateFalse)
    f.Write whatText
    f.Close
End Sub

'####################################################################################################################################
'NAME: Sub hAppTextFile(whatFile As String, whatText As String)
'DESCRIPTION: Appends text to the specified text file without overwriting its contents.  If the specified
'             file does not exist, this subroutine will create it and then write to it.
'RECEIVES: "whatFile" as a string contains the path and filename of the file to be appended to;
'          "whatText" as a string contains the text that is to be appended to that file
'RETURNS: NOTHING
'REQUIRES: NOTHING
'####################################################################################################################################
Sub hAppTextFile(whatfile As String, whatText As String)
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim fs, f, ts
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FileExists(whatfile) Then fs.CreateTextFile whatfile
    Set f = fs.GetFile(whatfile)
    Set ts = f.OpenAsTextStream(ForAppending, TristateFalse)
    ts.Write whatText
    ts.Close
End Sub

Sub hWriteBinaryFile(whatfile As String, whatToWrite As String)

    Open whatfile For Binary Access Write As #1
    ' Close before reopening in another mode.
    Close #1

End Sub
