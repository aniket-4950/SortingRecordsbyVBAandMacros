# SortingRecordsbyVBAandMacros
The data of the list can be sorted on various parameters just by the click of a button. Used VBA and Macros.

Code Explanation

Sorting by various columns :
Public Sub DivisonSort()
    Columns("A:F").Sort key1:=Range("A2"), order1:=xlDescending, Header:=xlYes
End Sub

Public Sub CategorySort()
    Columns("A:F").Sort key1:=Range("B2"), order1:=xlDescending, Header:=xlYes
End Sub

Public Sub TotalSort()
    Columns("A:F").Sort key1:=Range("F2"), order1:=xlDescending, Header:=xlYes
End Sub





Using user input and if statements to automate the sorting by the click of a button and also entertaining the wrong inputs.

Public Sub SortByChoice()
    Dim sortOrder As Integer
    Dim promptMSG As String
    Dim tryAgain As Integer

     On Error GoTo errhandler //for handling the errors
    
    
    promptMSG = "How would you like to sort your list" & vbCrLf & _
    "1 - Sort by Divison" & vbCrLf & _
    "2 - Sort by Category" & vbCrLf & _
    "3 - Sort by Total"

    sortOrder = InputBox(promptMSG, "Sort the table")
    
    If sortOrder = 1 Then
        DivisonSort
    ElseIf sortOrder = 2 Then
        CategorySort
    ElseIf sortOrder = 3 Then
        TotalSort
    Else
errhandler
        tryAgain = MsgBox("Invalid Value!, Do you want to try again", vbYesNo)
        If tryAgain = 6 Then
            SortByChoice
        End If
        
    End If
End Sub
&vbcrlf & _ : USED FOR BREAKING THE LINE AND CONTINUING THE TEXT FROM THE NEXT LINE

