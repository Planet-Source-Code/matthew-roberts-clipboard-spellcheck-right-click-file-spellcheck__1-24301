Attribute VB_Name = "modSpellCheck"

Option Explicit

'===============================================================================
'Purpose:       To easily access Word's spell checking library without having to open Word. To use it,
'                        compile the source, then create a shortcut to the .exe. Paste that shortcut into your
'                        Windows Profile -> All Users -> SendTo folder.
'
'                       Next, find a file to spell check. Right click on it and select SendTo -> SpellCheckFile
'                       This code will do the rest.
'
'
'Returns:         Corrected text will be IN THE WINDOWS CLIPBOARD when it is complete.
'
'Created By:    Matthew M.Roberts(M@)
'Date:              6/21/2001
'Comments:     NOTE: If you DO NOT pass in a file name (i.e. right-click and "send to",
'                      it will spell-check whatever is in the clipboard.
'===============================================================================

Private Sub Main()

    Dim objWord As New Word.Application
    Dim objDoc As New Word.Document
    Dim objSuggest
    Dim objCommandBar  As Object

    '   Create a document to work in
    Set objDoc = New Document
    
    '   No fancy error handling here!
     On Error Resume Next
    
    frmSplash.Show
   
    
    With objDoc
          
          If Command > "" Then
            .Content = GetFileContent(Command)
          Else
            .Content = Clipboard.GetText
          End If
          
          .Content.SetRange 1, Len(.Content)
              
'                   Hide the Word window

              Word.Application.Visible = False

'                   The loading of the object is complete, drop splash screen
            Unload frmSplash

          .Select
          
          If Selection.Range.SpellingErrors.Count = 0 Then
                MsgBox "All words are spelled correctly", vbInformation
          Else
          
                Selection.Range.CheckSpelling
'               Put the corrected text into the clipboard. This method preserves formatting.
            End If
            .Content.Copy
             Word.Application.WindowState = wdWindowStateMinimize
'               Close without saving
            .Close (False)
            
        
    End With
      
     
        Set objDoc = Nothing
        Word.Application.Quit
  
    
End Sub


Private Function GetFileContent(FileName As String) As String

Dim intFile As Integer
Dim strInput As String
Dim strContent As String

    intFile = FreeFile
    
    Open FileName For Input As #intFile
    
    While Not EOF(intFile)
        Line Input #intFile, strInput
        
        strContent = strContent & strInput & Chr(13)
           
    Wend
    
    GetFileContent = strContent
    
    
End Function

