'Language:  VBA Office 2010
'Author:    TheAZAndyman
'Program Function:
'   Collect data from Word fillable forms
'Version: 1.0.0.0
Sub GetData()
'total # of controls
Dim NumCC As Integer

'used to refer to a contentControl
Dim CC As ContentControl 'currently not being used

'Total # of controls missing tags
Dim NumCCMT As Integer 'All content controls should have a tag
NumCCMT = 0

'tag not unique
Dim CCTNU As String 'All content controls should be unique

'Data of content control
Dim DataCC As String

'array for the tags
Dim tagArray() As Variant

'array for the record fields
Dim RcrdArray() As Variant

'for the string to print/write to the file
Dim tagText As String

'had problem reusing tagText so using record text
Dim RcrdText As String

'need a long to refer to index in array and looping iterations
Dim t As Long
Dim r As Long

'for use for freefile or open file instance.
Dim F As Integer
'sting for the filepath
Dim FilePath As String

'total # of controls
NumCC = ActiveDocument.ContentControls.Count
    'MsgBox ("number of content controls = " & NumCC)
    
'The data of the first content control'
'This returns the placeholder text if nothing entered, not ideal
DataCC = ActiveDocument.ContentControls(1).Range.Text
    'MsgBox ("The data of the first content control is " & DataCC)

DataCC = ActiveDocument.ContentControls(1).PlaceholderText.Value
    'MsgBox (" PlaceholderText of the first content control is " & DataCC)
'to determine whether the document is showing the placeholder text,
'compare the .PlaceholderText.Value with the ContentControl.Range.Text.
'if the two are equal do nothing or add null to the array as placeholder when creating the output file.
'else get the range.txt
'TODO Setup a boolean.  Are they equal if true don't get text if false get the range.text

'TODO this is validation check move out of this sub and call it with this sub.
'TODO check if all content controls have tags.  if not message to user and end macro
'check each content control and if no tag or a space exists in the tag then not valid.
'and ideally tell me which one(s) doesn't have a tag, not needed at this time.
For Each CC In ActiveDocument.ContentControls
    'MsgBox ("Tag is " & CC.Tag)
    If CC.Tag = "" Then
        'MsgBox (CC.Title & " is Null!")
        NumCCMT = NumCCMT + 1
    End If
    If InStr(CC.Tag, " ") > 0 Then
        'MsgBox (CC.Title & " has a space")
        NumCCMT = NumCCMT + 1
    End If
Next

'Missing or tag has a space in it so inform user.
If NumCCMT > 0 Then
    MsgBox ("Invalid document! Content Controls missing or invalid tags. " & NumCCMT & " missing.")
    'TODO if this is true can we end execution here.
End If

'end of validation checks that should be moved out of this sub.

'TODO check each tag and compare it to all other tags and if all are not unique message box
'If all tags are present and don't have spaces then we can check uniqueness.
'don't need to check uniqueness if missing or invalid tags.
'ideally which one is seen as not unique
'These are validation checks and should be in a method/function that is called and if fails validation stop processing
'should be able ot report which files were invalid
'to select the content control based on the tag.
'ActiveDocument.SelectContentControlsByTag

'create an array of all the tags on the content controls.
'with validation checks of file we have confirmed all content controls have a tag on them.
For Each CC In ActiveDocument.ContentControls
    ReDim Preserve tagArray(t)
    tagArray(t) = CC.Tag
    t = t + 1
Next

'Create an array of all the txt of the content controls

For Each CC In ActiveDocument.ContentControls
    ReDim Preserve RcrdArray(r)
    'boolean that compares placeholder text to range.text and only add to array if different else add empty index ("").
    If CC.Range.Text <> CC.PlaceholderText Then
    RcrdArray(r) = CC.Range.Text
    Else
        RcrdArray(r) = ""
    End If
    r = r + 1
Next

'loop through array and write each to a string variable with a tab between them.
For t = LBound(tagArray) To UBound(tagArray)
    If t = UBound(tagArray) Then
        tagText = tagText & tagArray(t)
    Else
        tagText = tagText & tagArray(t) & vbTab
    End If

Next t

'TODO now get the text for the record fields line
'loop through Rcrdarray and write each to a string variable with a tab between them.
For r = LBound(RcrdArray) To UBound(RcrdArray)
    If r = UBound(RcrdArray) Then
        RcrdText = RcrdText & RcrdArray(r)
    Else
        RcrdText = RcrdText & RcrdArray(r) & vbTab
   End If
    'MsgBox (RcrdText)
Next r

'Print the record, hopefully it goes on the next line.
'What is the file path and name for the new text file?
'TODO ask user to select the location
'TODO ask user to specify the filename.
FilePath = "D:\Friendswork\MyFile.txt"
'Determine the next file number available for use by the FileOpen function
F = FreeFile
'Open the text file which creates it.
Open FilePath For Output As F
    

    'Display the strings
    MsgBox (tagText)
    MsgBox (RcrdText)
    
'print the string variables to the file.
    Print #F, tagText
    Print #F, RcrdText
    Close F
    MsgBox ("file created")
End Sub

'This is a test of creating a textfile
'not sure why use this instead of print or write and uses the Scripting library.
Sub CreateAfile()
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("D:\Friendswork\testfile.txt", True)
    a.Writeline ("This is a test.") 'this is where we would spool out the array.
    a.Writeline ("This is another test.") 'this is where we would spool out the array.
    a.Close  'closes the file or it is seen as in use by filesystem.
    MsgBox ("file created")
End Sub




Sub TextFile_Create()
Dim TextFile As Integer
Dim FilePath As String

'What is the file path and name for the new text file?
  FilePath = "D:\Friendswork\MyFile.txt"

'Determine the next file number available for use by the FileOpen function
  TextFile = FreeFile

'Open the text file
  Open FilePath For Output As TextFile

'Write some lines of text
  Print #TextFile, "Hello Everyone!"
  Print #TextFile, "I created this file with VBA."
  Print #TextFile, "Goodbye"
  
'Save & Close Text File
  Close TextFile

End Sub















