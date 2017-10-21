VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' File   : clsButton
' Author : Anthony Malone
' Date   : 19/10/2017
' Purpose: Wrap incoming commandbutton, handle its events by drawing from tblForms
'---------------------------------------------------------------------------------------
'@Folder("Wrapper")
Option Compare Database
Option Explicit

Private WithEvents Cmd As Access.CommandButton
Attribute Cmd.VB_VarHelpID = -1
Private m_FormName As String

'---------------------------------------------------------------------------------------
' Method : FormName
' Author : Anthony Malone
' Date   : 19/10/2017
' Purpose: Outputs FormName (required by DLOOKUPs)
'---------------------------------------------------------------------------------------
Public Property Get FormName() As String

    FormName = m_FormName

End Property

'---------------------------------------------------------------------------------------
' Method : FormName
' Author : Anthony Malone
' Date   : 19/10/2017
' Purpose: Passes incoming Form Name to class variable (required by DLOOKUPs)
'---------------------------------------------------------------------------------------
Public Property Let FormName(ByVal strFormName As String)

    m_FormName = strFormName

End Property

'---------------------------------------------------------------------------------------
' Method : CmdButton
' Author : Anthony Malone
' Date   : 19/10/2017
' Purpose: Wraps Incoming CommandButton, securitises and adds tooltip
'---------------------------------------------------------------------------------------
Public Property Set CmdButton(cmdCallingButton As Access.CommandButton)

LoadButton:

    Set Cmd = cmdCallingButton
    
Securitisation:

    If DLookup("SecurityLevel", "tblCommandButtons", "Form='" & FormName & "' And ButtonName='" & Cmd.Name & "'") = 1 Then
    
        Cmd.Visible = True

SetToolTip:
    
        Cmd.ControlTipText = Nz(DLookup("DisplayThisToolTip", "tblCommandButtons", "Form='" & FormName & "' And ButtonName='" & Cmd.Name & "'"), "")
    
    Else
    
        Cmd.Visible = False
    
    End If
    
End Property

'---------------------------------------------------------------------------------------
' Method : Cmd_Click
' Author : Anthony Malone
' Date   : 19/10/2017
' Purpose: Handles incoming CommandButton click event
'---------------------------------------------------------------------------------------

Private Sub Cmd_Click()

    Dim strSourceForm As String, strButtonName As String
    
TrapBecauseWeMightCloseSourceFormDown:
    
    strSourceForm = FormName
    strButtonName = Cmd.Name

EitherOpenAForm:

    If Not IsNull(DLookup("OpenThisForm", "tblCommandButtons", "Form='" & strSourceForm & "' And ButtonName='" & strButtonName & "'")) Then
    
        DoCmd.Close
        DoCmd.OpenForm DLookup("OpenThisForm", "tblCommandButtons", "Form='" & strSourceForm & "' And ButtonName='" & strButtonName & "'")
    
OrRunSomeCode:
    
    ElseIf Not IsNull(DLookup("RunThisCode", "tblCommandButtons", "Form='" & strSourceForm & "' And ButtonName='" & strButtonName & "'")) Then
    
        Application.Run DLookup("RunThisCode", "tblCommandButtons", "Form='" & strSourceForm & "' And ButtonName='" & strButtonName & "'")
    
    Else
    
OtherwiseDisplayTBC:
    
        If Not IsNull(DLookup("PlaceholderAlertText", "tblCommandButtons", "Form='" & strSourceForm & "' And ButtonName='" & strButtonName & "'")) Then MsgBox DLookup("PlaceholderAlertText", "tblCommandButtons", "Form='" & FormName & "' And ButtonName='" & Cmd.Name & "'")
        
    End If
    
End Sub
