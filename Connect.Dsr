VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   13485
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   15735
   _ExtentX        =   27755
   _ExtentY        =   23786
   _Version        =   393216
   Description     =   "Add-In Project Template"
   DisplayName     =   "Code Inserter"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSafe     =   -1  'True
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Public VBInstance As VBIDE.VBE
Private mcbMenuCommandBarCtrl As Object
Public WithEvents ErrorBlockMenuHandler As CommandBarEvents
Attribute ErrorBlockMenuHandler.VB_VarHelpID = -1
Public WithEvents OnErrorMenuHandler As CommandBarEvents
Attribute OnErrorMenuHandler.VB_VarHelpID = -1
Public WithEvents ErrorHandlerMenuHandler As CommandBarEvents
Attribute ErrorHandlerMenuHandler.VB_VarHelpID = -1
Private MenuItem() As clsMenuItem
'

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
'Event runs when AddIn is added to Instance of VB IDE
On Error GoTo ErrorHandler
   Dim cbMenu As Object
   Dim lngIndex As Long
   '
      'Link To Instance of VB IDE
      Set VBInstance = Application
      'Link to PopUp Menu for Code Editor Window
      Set cbMenu = VBInstance.CommandBars("Code Window")
      '
      If cbMenu Is Nothing Then
         Exit Sub
      End If
      '
      'Add My AddIn to Code Window PopUp Menu
      Set mcbMenuCommandBarCtrl = cbMenu.Controls.Add(10, , , 1)
      With mcbMenuCommandBarCtrl
         .Caption = "Insert Code"
      End With
      '
      'Load Menu Items List from File
      Call LoadMenuItemsFromFile
      '
      For lngIndex = LBound(MenuItem) To UBound(MenuItem)
         MenuItem(lngIndex).Add
      Next
      '
ExitRoutine:
On Error Resume Next
   Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
'Event runs when AddIn is removed from Instance of VB IDE
On Error Resume Next
Dim lngIndex As Long
'
   For lngIndex = LBound(MenuItem) To UBound(MenuItem)
      MenuItem(lngIndex).Remove
      Set MenuItem(lngIndex) = Nothing
   Next
   mcbMenuCommandBarCtrl.Delete
   Set VBInstance = Nothing
End Sub

Public Function LoadMenuItemsFromFile() As Collection
'Read the Menu Items from the file
On Error GoTo ErrorHandler
   Dim strLineText As String
   Dim lngArrayIndex As Long
   '
      lngArrayIndex = -1      'We initialize at -1 so that the first index will be 0
      '
      'Open File
      Open App.Path & "\MenuItems.txt" For Input As #1
      Do While Not EOF(1)
         'Line Input #1, strLineText
         'If we find [...] it indicates the start of a Menu Item
         If Left(strLineText, 1) = "[" And Right(strLineText, 1) = "]" Then
            'Get Array index and instantiate MenuItem Class
            lngArrayIndex = lngArrayIndex + 1
            ReDim Preserve MenuItem(lngArrayIndex) As clsMenuItem
            '
            'Set Top Level Properties
            Set MenuItem(lngArrayIndex) = New clsMenuItem
            Set MenuItem(lngArrayIndex).VBIDEInstance = VBInstance
            Set MenuItem(lngArrayIndex).CommandBar = mcbMenuCommandBarCtrl
            '
            Line Input #1, strLineText
            '
            'Loop until we get to the next menu item
            Do Until (Left(strLineText, 1) = "[" And Right(strLineText, 1) = "]")
               'Check that we have a Menu Item Property i.e.  Line must contain an = must not start with a '
               If InStr(1, strLineText, "=", vbTextCompare) <> 0 And Left(strLineText, 1) <> "'" Then
                  Select Case Left(strLineText, InStr(1, strLineText, "=", vbTextCompare) - 1)
                  Case "Caption"
                     MenuItem(lngArrayIndex).MenuCaption = Right(strLineText, Len(strLineText) - InStr(1, strLineText, "=", vbTextCompare))
                  Case "CodeBlockName"
                     MenuItem(lngArrayIndex).MenuCodeBlock = Right(strLineText, Len(strLineText) - InStr(1, strLineText, "=", vbTextCompare))
                  Case "InsertionPoint"
                     MenuItem(lngArrayIndex).MenuCodeInsertPoint = CLng(Right(strLineText, Len(strLineText) - InStr(1, strLineText, "=", vbTextCompare)))
                  End Select
               End If
               Line Input #1, strLineText
            Loop
         Else
            Line Input #1, strLineText
         End If
      Loop
      '
Exit_Routine:
On Error Resume Next
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

'
'Private Sub ErrorBlockMenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
''Event runs when menu item is clicked
'On Error Resume Next
'   Dim strText As String
'   '
'      strText = GetStringFromFile("OnError")
'      Call VBInstance.ActiveCodePane.CodeModule.InsertLines(FirstLine, strText)
'      strText = GetStringFromFile("ErrorHandler")
'      Call VBInstance.ActiveCodePane.CodeModule.InsertLines(LastLine, strText)
'End Sub
'
'Private Sub ErrorHandlerMenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
''Event runs when menu item is clicked
'On Error Resume Next
'   Dim strText As String
'   '
'      strText = GetStringFromFile("ErrorHandler")
'      Call VBInstance.ActiveCodePane.CodeModule.InsertLines(LastLine, strText)
'End Sub
'
'Private Sub OnErrorMenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
''Event runs when menu item is clicked
'On Error Resume Next
'   Dim strText As String
'   '
'      strText = GetStringFromFile("OnError")
'      Call VBInstance.ActiveCodePane.CodeModule.InsertLines(FirstLine, strText)
'End Sub

