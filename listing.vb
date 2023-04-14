'
' Macro to go through and parse library names
' Z Ashton 2023
' This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License along with this program. If not, see <https://www.gnu.org/licenses/>.

Option Explicit
Dim finalRow As Integer
Dim curRow As Integer
Public globalSplitCatagories As Variant



Private Sub acceptCatagoryFix_Click()
    Dim curCell As String
    ' Catagory Column Definition Here
    curCell = "Q" & curRow
    Debug.Print curCell
    ActiveSheet.Range(curCell).Select
    Dim i As Integer
    i = 0
    For i = LBound(globalSplitCatagories) To UBound(globalSplitCatagories)
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Value = globalSplitCatagories(i)
    Next i
    
End Sub

Private Sub AcceptChange_Click()
  Dim curCell As String
  ' saftey check lol
  ' Author Column Definition Here
  curCell = "E" & curRow
  ActiveSheet.Range(curCell).Select
  ActiveCell.Value = NewStyleBox.Value
  
End Sub



Private Sub Label1_Click()

End Sub

Private Sub NxtRow_Click()
 Dim myCell As String
 Dim splitString() As String
 Dim newString As String
 Dim splitStringLength As Integer
 
 curRow = curRow + 1
 ' Author Column Definition Here
 myCell = "E" & curRow
 ' MsgBox myCell
 ActiveSheet.Range(myCell).Select
 CurCelLBL.Caption = "Current Cell " & myCell
 CurFormatBox.Value = ActiveCell.Value
 
 newString = ReverseSentence(ActiveCell.Value)
 
 NewStyleBox.Value = newString
 
 SplitSubjects
 
End Sub


Sub SplitSubjects()
    'Erase globalSplitCatagories()
    Dim curCell As String
    Dim spiltCatagories() As String
    curCell = "Q" & curRow
    ActiveSheet.Range(curCell).Select
    
    spiltCatagories = Split(ActiveCell.Value & " ", delimiterValue.Value)
    subjectColShowUser.Caption = Join(spiltCatagories, vbNewLine)
    globalSplitCatagories = spiltCatagories
    curCell = "E" & curRow
    ActiveSheet.Range(curCell).Select
    
End Sub

Function ReverseSentence(inputString As String)
    
    Dim word As Variant
    Dim reversed As String
    reversed = StrReverse(inputString)
    
    Dim cWord As Integer
    cWord = 0
    For Each word In Split(reversed)
        If cWord = 1 Then
            ReverseSentence = Trim(ReverseSentence & " " & StrReverse(word))
        Else
            ReverseSentence = Trim(ReverseSentence & StrReverse(word) & ", ")
        End If
        cWord = cWord + 1
    Next
 
End Function
Private Sub startBtn_Click()

    Dim msgAns As Integer
    Dim placeholder As Integer
    

    msgAns = MsgBox("Before continuing, please ensure that the AUTHOR column is COLUMN E and that the subject column is COLUMN Q on the worksheet. Confirm?", vbQuestion + vbYesNo + vbDefaultButton2, "Saftey Check!")

    If msgAns = vbYes Then
        LoopThroughCells
    Else
        placeholder = MsgBox("Aborting processing of data.", vbExclamation + vbOKOnly + vbDefaultButton1, "Aborting")
    End If
End Sub

Private Sub UserForm_Click()
  
End Sub

Private Sub LoopThroughCells()


 ' Select cell A2, *first line of data*.
      
       
      
      Range("A2").Select
      
      ' Set Do loop to stop when an empty cell is reached.
      Do Until IsEmpty(ActiveCell)
         ' Insert your code here.
         ' Step down 1 row from present location.
         ActiveCell.Offset(1, 0).Select
         CurCelLBL.Caption = ActiveCell.Row
         finalRow = ActiveCell.Row
      Loop
      
      CurCelLBL.Caption = "Number of rows: " & finalRow
      curRow = 1
      
      MsgBox "Ready to start processing values. Rows: " & finalRow
      'globalSplitCatagories = "Hello"
      NxtRow_Click

 
 
End Sub



' from wellsr.com
Sub ReverseArray(vArray As Variant)
'Reverse the order of an array, so if it's already sorted
'from smallest to largest, it will now be sorted from
'largest to smallest.
Dim vTemp As Variant
Dim i As Long
Dim iUpper As Long
Dim iMidPt As Long
iUpper = UBound(vArray)
iMidPt = (UBound(vArray) - LBound(vArray)) \ 2 + LBound(vArray)
For i = LBound(vArray) To iMidPt
    vTemp = vArray(iUpper)
    vArray(iUpper) = vArray(i)
    vArray(i) = vTemp
    iUpper = iUpper - 1
Next i
End Sub
