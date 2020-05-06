VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormulaTraceForm 
   Caption         =   "Trace"
   ClientHeight    =   2400
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6756
   OleObjectBlob   =   "FormulaTraceForm.frx":0000
End
Attribute VB_Name = "FormulaTraceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   Michael Curley
'



' ----- local fields -----
Private parser As FormulaParser
Private helpMessage As String
Private theFormula As String
Private theValue As String
Private leftPad As Integer
Private textPercentBoundary As Double
Private newLineChr As String
Private paragraphChr As String
Private helpKey As Integer
Private formulaKey As Integer
Private valueKey As Integer
Private escapeKey As Integer
Private enterKey As Integer
Private leftShiftKey As Integer
Private rightShiftKey As Integer
Private resetSplitDataSelectionShiftKey As Integer
Private pixelShiftOnKey As Integer





' ----- constructor, initialize constants --------------------------------------

Private Sub UserForm_Initialize()
    ' create new parser for the form
    Set parser = New FormulaParser
    
    ' all following values are arbitrary, used in a global scope to maintain
    ' functionality within the constructor only
    theFormula = parser.formula
    theValue = ActiveCell.Text
    leftPad = 6
    textPercentBoundary = 2# / 5# ' 40%
    newLineChr = Chr(10)
    paragraphChr = Chr(182)
    helpKey = 104 ' h
    formulaKey = 102 ' f
    valueKey = 118 ' v
    escapeKey = 27
    enterKey = 13
    leftShiftKey = 62 ' <
    rightShiftKey = 60 ' >
    resetSplitDataSelectionShiftKey = 63 ' ?
    pixelShiftOnKey = 5
    pixelShiftPadding = pixelShiftOnKey * 5
    
    ' init the help message
    helpMessage = "" ' annoying but easier to see overall formatting
    helpMessage = helpMessage & "Select an option from ""Filtered By"" drop-down to modify what" & vbNewLine
    helpMessage = helpMessage & """Data"" is highlighted and linked to your worksheet." & vbNewLine & vbNewLine
    helpMessage = helpMessage & "Arrowing up/down will select different aspects of the displayed" & vbNewLine
    helpMessage = helpMessage & "equation and bring you to referenced ranges if applicable." & vbNewLine & vbNewLine
    helpMessage = helpMessage & "Double clicking on a data entry will expand that entry in its own" & vbNewLine
    helpMessage = helpMessage & "window in case you can't read the full text." & vbNewLine & vbNewLine
    helpMessage = helpMessage & "Keyboard Shortcuts:" & vbNewLine
    helpMessage = helpMessage & "   h" & vbTab & "Open this help window" & vbNewLine
    helpMessage = helpMessage & "   >" & vbTab & "Scroll the formula text to the right" & vbNewLine
    helpMessage = helpMessage & "   <" & vbTab & "Scroll the formula text to the left" & vbNewLine
    helpMessage = helpMessage & "   ?" & vbTab & "Scroll the formula text to its original position" & vbNewLine
    helpMessage = helpMessage & "   f" & vbTab & "Expand the full formula in a new window" & vbNewLine
    helpMessage = helpMessage & "   v" & vbTab & "Expand the final value in a new window" & vbNewLine
    helpMessage = helpMessage & "   enter" & vbTab & "Expand the selected data entry in a new window" & vbNewLine
    helpMessage = helpMessage & "   esc" & vbTab & "Close the current window"
    
    ' set the window in upper right corner
    Top = Application.Top + 50 ' top pad
    Left = Application.Left + Application.UsableWidth - Width - 50 ' right pad
    
    ' set title
    Caption = "Trace: " & theValue
    
    ' NOTE: the following must match indexes in FilterComboBox_Change()
    ' set the data in the combobox, select the first item, will trigger
    ' a full formatting of the UI
    FilterComboBox.AddItem ("All Arguments")
    FilterComboBox.AddItem ("References Only")
    FilterComboBox.AddItem ("All Delimiters")
    FilterComboBox.ListIndex = 0
End Sub





' ----- UserForm (FormulaTraceForm) --------------------------------------------

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    showOutputValueMsgBox
End Sub





' ----- FilterComboBox ---------------------------------------------------------

Private Sub FilterComboBox_Change()
    If FilterComboBox.ListIndex = 0 Then
        parser.splitByArguments
    ElseIf FilterComboBox.ListIndex = 1 Then
        parser.splitByReferences
    ElseIf FilterComboBox.ListIndex = 2 Then
        parser.splitByAll
    End If
    ' reset selection after splitting data
    resetSplitDataSelection
End Sub





' ----- KeyboardShortcutsCheckBox ----------------------------------------------

Private Sub KeyboardShortcutsCheckBox_Change()
    If Not KeyboardShortcutsCheckBox.Value Then
        alignFormulaLabels
    End If
    DataListBox.SetFocus
End Sub





' ----- HelpButton -------------------------------------------------------------

Private Sub HelpButton_Click()
    showHelpMsgBox
    DataListBox.SetFocus
End Sub





' ----- DataListBox ------------------------------------------------------------

Private Sub DataListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    showDataSelectionMsgBox
End Sub



Private Sub DataListBox_Change()
    Dim i As Integer
    Dim formula() As String
    For i = 0 To DataListBox.ListCount
        If DataListBox.Selected(i) Then
            formula = parser.gotoDataAndSplitFormula(i)
            LeftFormulaLabel = replaceNewLines(formula(0))
            MiddleFormulaLabel = replaceNewLines(formula(1))
            RightFormulaLabel = replaceNewLines(formula(2))
            alignFormulaLabels
            Exit For
        End If
    Next
End Sub



Private Sub DataListBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    If KeyAscii = helpKey Then
        showHelpMsgBox
    ElseIf KeyAscii = formulaKey Then
        showFormulaMsgBox
    ElseIf KeyAscii = valueKey Then
        showOutputValueMsgBox
    ElseIf KeyAscii = escapeKey Then
        Unload Me
    ElseIf KeyboardShortcutsCheckBox.Value Then
        If KeyAscii = enterKey Then
            showDataSelectionMsgBox
        ElseIf KeyAscii = leftShiftKey Then
            If RightFormulaLabel.Left + RightFormulaLabel.Width > Width * textPercentBoundary Then
                LeftFormulaLabel.Left = LeftFormulaLabel.Left - pixelShiftOnKey
                MiddleFormulaLabel.Left = MiddleFormulaLabel.Left - pixelShiftOnKey
                RightFormulaLabel.Left = RightFormulaLabel.Left - pixelShiftOnKey
            End If
        ElseIf KeyAscii = rightShiftKey Then
            If LeftFormulaLabel.Left < 0 Then
                LeftFormulaLabel.Left = LeftFormulaLabel.Left + pixelShiftOnKey
                MiddleFormulaLabel.Left = MiddleFormulaLabel.Left + pixelShiftOnKey
                RightFormulaLabel.Left = RightFormulaLabel.Left + pixelShiftOnKey
            End If
        ElseIf KeyAscii = resetSplitDataSelectionShiftKey Then
            alignFormulaLabels
        End If
    End If
End Sub





' ----- privates ---------------------------------------------------------------

' replaces newlines with paragraph characters
Private Function replaceNewLines(str As String) As String
    replaceNewLines = Replace(str, newLineChr, paragraphChr)
End Function



' shows the output of the formula in a message box
Private Function showFormulaMsgBox()
    MsgBox theFormula, , "Trace Formula"
End Function



' shows the output of the formula in a message box
Private Function showOutputValueMsgBox()
    MsgBox theValue, , "Trace Value"
End Function



' shows the help message box
Private Function showHelpMsgBox()
    MsgBox helpMessage, , "Help"
End Function



' shows whatever data is selected in a message box
Private Function showDataSelectionMsgBox()
    For i = 0 To DataListBox.ListCount
        If DataListBox.Selected(i) Then
            MsgBox DataListBox.List(i), , "Data Expansion"
            Exit For
        End If
    Next
End Function



' based on the current text of the formula labels, will shift the text left
' if the middle label extends beyone the textPercentBoundary
Private Function alignFormulaLabels()
    If leftPad + LeftFormulaLabel.Width < Width * textPercentBoundary Then
        LeftFormulaLabel.Left = leftPad
        If LeftFormulaLabel = "" Then
            LeftFormulaLabel.Left = LeftFormulaLabel.Left - LeftFormulaLabel.Width
        End If
        MiddleFormulaLabel.Left = LeftFormulaLabel.Left + LeftFormulaLabel.Width
        RightFormulaLabel.Left = MiddleFormulaLabel.Left + MiddleFormulaLabel.Width
    Else
        MiddleFormulaLabel.Left = (Width * textPercentBoundary) - (MiddleFormulaLabel.Width / 2)
        LeftFormulaLabel.Left = MiddleFormulaLabel.Left - LeftFormulaLabel.Width
        RightFormulaLabel.Left = MiddleFormulaLabel.Left + MiddleFormulaLabel.Width
    End If
End Function



' after the parser has had a splitBy call, will repolulate the DataListBox with
' the newly set data
Private Function resetSplitDataSelection()
    Dim first As Boolean
    Dim val As Boolean
    DataListBox.Clear
    first = True
    For Each dat In parser.data
        If first And dat = "=" Then
            DataListBox.AddItem ("= (ORIGIN)")
        Else
            DataListBox.AddItem (dat)
        End If
        first = False
    Next
    KeyboardShortcutsCheckBox.Enabled = Not first
    MiddleFormulaLabel.Visible = Not first
    RightFormulaLabel.Visible = Not first
    If first Then ' nothing was added
        LeftFormulaLabel.Left = leftPad
        LeftFormulaLabel = "No Data to Display for Selected Cell: " & ActiveCell.Address
    Else
        DataListBox.Selected(0) = True
        DataListBox.SetFocus
    End If
End Function





