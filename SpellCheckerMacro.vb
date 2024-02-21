

Sub SpellCheckerMacro()
'
' SpellCheckerMacro
'

'
' Prepared By M A Hafeez (Developer at Mahar Technologies)
' Contact: mahar.technologies1@gmail.com
'
  Dim limit As range
  
  Set docSource = ActiveDocument
  
  Set docNew = Documents.Add
  
  For Each limit In docSource.SpellingErrors
  
    limit.Font.Color = wdColorRed
    
    limit.Font.Bold = True
    
    docNew.range.InsertAfter limit.Text
    
  Next
  
'
' This macro will check the spell in current document.
' If found word with misspelled it will highlight in red.
' Misspeld words will be save in second unsaved file.
'
  
End Sub
  
