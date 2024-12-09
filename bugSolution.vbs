Function processInput(input)
  If IsNumeric(input) Then
    MsgBox "Numeric Input: " & input * 2
  Else
    MsgBox "String Input: " & UCase(input)
  End If
End Function

'Illustrate the issue: 
'processInput(10) 'This will work correctly (using the second definition of processInput)
'processInput("hello") 'This will work correctly (using the second definition of processInput) 


Function processInput(input)
  MsgBox "String Input: " & UCase(input) 
End Function

processInput(10) 'This will now only use the second version of the function
processInput("hello") 'This will work correctly (using the second definition of processInput)