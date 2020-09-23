Attribute VB_Name = "mosque"
'declaring some global variables

'database holder
Public donorInformation As Database

'sql statement holder
Public SQL As String

'line of string to be used in listbox additem method
Public strLine As String

'variable to hold my message boxes
Public msg As String

Public Function Selector(textSelector As TextBox)
    textSelector.SelStart = 0
    textSelector.SelLength = Len(textSelector.Text)
    textSelector.SetFocus
End Function

Public Sub deleteTextBoxes(passedForm As Form)
    Dim Control
        For Each Control In passedForm.Controls
        If TypeOf Control Is TextBox Then Control.Text = ""
    Next Control

End Sub
    
Public Sub UnlockTextBoxes(passedForm As Form)
    Dim Control
        For Each Control In passedForm.Controls
        If TypeOf Control Is TextBox Then Control.Locked = False
    Next Control

End Sub
Public Sub lockTextboxes(passedForm As Form)
    Dim Control
        For Each Control In passedForm.Controls
        If TypeOf Control Is TextBox Then Control.Locked = True
    Next Control
End Sub
