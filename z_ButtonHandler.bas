Attribute VB_Name = "z_ButtonHandler"
'---------------------------------------------------------------------------------------
' File   : z_ButtonHandler
' Author : Anthony Malone
' Date   : 19/10/2017
' Purpose: Loops through passed Form commandbuttons, adds to custom collection to handle events
'---------------------------------------------------------------------------------------
'@Folder("Wrapper")
Option Compare Database
Option Explicit
Dim myCMD As clsButton
Public colCommandButtons As New Collection

'---------------------------------------------------------------------------------------
' Method : HandleButtons
' Author : Anthony Malone
' Date   : 19/10/2017
' Purpose: Builds collection of custom commandbutton wrappers from form controls
'---------------------------------------------------------------------------------------
Sub HandleButtons(myForm As Form)

    Dim ctl As Control
    Dim strControlName As String
    
    On Error GoTo E_Handle
    
LoopThroughFormControls:
    
    For Each ctl In myForm.Controls
    
ActionOnlyCommandButtons:
        
        If ctl.ControlType = acCommandButton Then
        
StoreName:
            
            strControlName = ctl.Name
            
PassButtonToClassWrapper:
            
            Set myCMD = New clsButton
            myCMD.FormName = myForm.Name
            Set myCMD.CmdButton = myForm.Controls(strControlName)
            
AddMyButtonObjectToCollection:
    
            colCommandButtons.Add myCMD
            
PassEventHandlingToButtonClass:
            
            myForm.Controls(strControlName).OnClick = "[Event Procedure]"
        
        End If
        
    Next ctl
    
    Set myCMD = Nothing
        
sExit:

    On Error Resume Next
    Exit Sub
    
E_Handle:

    MsgBox Err.Description & vbCrLf & vbCrLf & "HandleButtons", vbOKOnly + vbCritical, "Error: " & Err.Number
    Resume sExit

End Sub

'---------------------------------------------------------------------------------------
' Method : ReleaseButtons
' Author : Anthony Malone
' Date   : 19/10/2017
' Purpose: Destroy command button collection object
'---------------------------------------------------------------------------------------

Sub ReleaseButtons()

    Set colCommandButtons = Nothing

End Sub
