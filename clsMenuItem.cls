VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Public Enum enInsertPoint
   ProcStart = 0
   Cursor = 1
   ProcEnd = 2
End Enum
'
Private strCaption As String
Private InsertionPoint As enInsertPoint
Private strCodeBlockName As String
Private cbCommandBarControl As Object
Private VBInstance As VBIDE.VBE
Public WithEvents MenuHandler As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1
'

Public Property Set VBIDEInstance(Item As Object)
   Set VBInstance = Item
End Property

Public Property Get VBIDEInstance() As Object
   Set VBIDEInstance = VBInstance
End Property

Public Property Set CommandBar(Item As Object)
   Set cbCommandBarControl = Item
End Property

Public Property Get CommandBar() As Object
   Set CommandBar = cbCommandBarControl
End Property

Public Property Let MenuCaption(value As String)
   strCaption = value
End Property

Public Property Get MenuCaption() As String
   MenuCaption = strCaption
End Property

Public Property Let MenuCodeInsertPoint(value As enInsertPoint)
   InsertionPoint = value
End Property

Public Property Get MenuCodeInsertPoint() As enInsertPoint
   MenuCodeInsertPoint = InsertionPoint
End Property

Public Property Let MenuCodeBlock(value As String)
   strCodeBlockName = value
End Property

Public Property Get MenuCodeBlock() As String
   MenuCodeBlock = strCodeBlockName
End Property

Public Sub Add()
   Dim cbCmdBarCtrl As Object
      'Add Sub-Menu
      Set cbCmdBarCtrl = cbCommandBarControl.Controls.Add(1)
      'Give menu Item a Caption
      cbCmdBarCtrl.Caption = strCaption
      'Link the MenuHandler to the Menu Item
      Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(cbCmdBarCtrl)
      Set cbCmdBarCtrl = Nothing
End Sub

Public Sub Remove()
   MenuCaption = Empty
   MenuCodeBlock = Empty
   MenuCodeInsertPoint = Empty
   Set MenuHandler = Nothing
End Sub

Private Sub Class_Terminate()
   Set VBInstance = Nothing
   Set cbCommandBarControl = Nothing
End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'Event runs when menu item is clicked
On Error Resume Next
   Dim strText As String
   '
      strText = GetStringFromFile(strCodeBlockName)
      Select Case InsertionPoint
      Case enInsertPoint.ProcStart
         Call VBInstance.ActiveCodePane.CodeModule.InsertLines(FirstLine, strText)
      Case enInsertPoint.Cursor
         Call VBInstance.ActiveCodePane.CodeModule.InsertLines(CurrentStartLine, strText)
      Case enInsertPoint.ProcEnd
         Call VBInstance.ActiveCodePane.CodeModule.InsertLines(LastLine, strText)
      End Select
End Sub

Private Function FirstLine() As Long
'Retrieve First Line number of Procedure
On Error Resume Next
   Dim strProcName As String
   Dim lngProcStartLine As Long
   '
      strProcName = VBInstance.ActiveCodePane.CodeModule.ProcOfLine(CurrentStartLine, vbext_pk_Proc)
      If strProcName = "" Then Exit Function
      lngProcStartLine = VBInstance.ActiveCodePane.CodeModule.ProcBodyLine(strProcName, vbext_pk_Proc)
      FirstLine = lngProcStartLine + 1
End Function

Private Function LastLine() As Long
'Retrieve Last Line number of Procedure
On Error Resume Next
   Dim strProcName As String
   Dim lngProcStartLine As Long
   Dim lngProcLineCount As Long
   '
      strProcName = VBInstance.ActiveCodePane.CodeModule.ProcOfLine(CurrentStartLine, vbext_pk_Proc)
      If strProcName = "" Then Exit Function
      lngProcStartLine = VBInstance.ActiveCodePane.CodeModule.ProcStartLine(strProcName, vbext_pk_Proc)
      lngProcLineCount = VBInstance.ActiveCodePane.CodeModule.ProcCountLines(strProcName, vbext_pk_Proc)
      LastLine = lngProcStartLine + lngProcLineCount - 1
End Function

Private Function CurrentStartLine() As Long
'Retrieve Line number of current cursor position or First Line number of selection if block selected Procedure
On Error Resume Next
   Dim lngStartLine As Long
   Dim lngStartColumn As Long
   Dim lngEndLine As Long
   Dim lngEndColumn As Long
   '
      Call VBInstance.ActiveCodePane.CodeModule.CodePane.GetSelection(lngStartLine, lngStartColumn, lngEndLine, lngEndColumn)
      CurrentStartLine = lngStartLine
End Function

Public Function GetStringFromFile(CodeBlockName As String) As String
'Read the code to insert from file
On Error GoTo ErrorHandler
   Dim strLineText As String
   Dim strText As String
   '
      Open App.Path & "\Code.txt" For Input As #1
      Do While Not EOF(1)
         Line Input #1, strLineText
         If strLineText = "[" & CodeBlockName & "]" Then
            Line Input #1, strLineText
            Do Until Left(strLineText, 1) = "[" And Right(strLineText, 1) = "]"
               strText = strText & strLineText & vbCrLf
               Line Input #1, strLineText
            Loop
            Exit Do
         End If
      Loop
      '
Exit_Routine:
On Error Resume Next
   'Remove vbCrLf from end of string
   GetStringFromFile = Left(strText, Len(strText) - 2)
   Close #1
   Exit Function
   '
ErrorHandler:
   Select Case Err.Number
   Case 62
      'do nothing - End of file error
   Case Else
      MsgBox "An Error has occured." & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation, "Insert Code Error"
   End Select
   Resume Exit_Routine
End Function
