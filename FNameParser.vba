Dim rng As Range
Dim clickValue As Integer
Dim nameArray() As String
Dim origName As String
Dim newMiddle As String
Dim counter As Integer
Dim middleName As String

'Loop through each first name
'For counter = 975 To 11182
    'Get current text in the cell
    origName = Cells(rng.row, rng.Column)
    'Split it up, space delimited, into an array
    nameArray() = Split(origName)
    'Display message box to show original fname, new fname and new mname.
    clickValue = MsgBox("Original first name: " + origName + vbCrLf + "New first name: " + nameArray(0) + vbCrLf + "New Middle Name: " + nameArray(1) + vbCrLf + "Make Change?", vbYesNo, "Title")
      
    'If yes is clicked, fname, mname, and status are updated with info
    If test = 6 Then
        Cells(rng.row, rng.Column) = nameArray(0)
        Cells(rng.row, rng.Column + 2) = nameArray(1)
        Cells(rng.row, rng.Column - 2) = "U"
    'If no is clicked, Input box asks for middle name, then fname, mname, and status are updated.
    Else
        Cells(rng.row, rng.Column + 2) = InputBox("Change middle name to what?")
    Cells(rng.row, rng.Column) = nameArray(0)
        Cells(rng.row, rng.Column - 2) = "U"
        
    End If
    
Next counter
