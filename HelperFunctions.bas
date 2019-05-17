' StockReports
' @Author Jeffery Brown (daddyjab)
' @Date 2/2019
' @File HelperFunctions.bas

Attribute VB_Name = "HelperFunctions"

Function ScreenUpdating(Optional AFlag As Boolean = True) As Boolean
'Turn Screen Updating on (to help speed up macro execution) or off (to show screen changes)
'   AFlag = True if Screen Updating should be turned on (default)
'   AFlag = False if Screen Updating should be turned off
'   retval = Previous value of the ScreenUpdating property (before being set by this function)

    If (IsMissing(AFlag) Or IsEmpty(AFlag) Or AFlag) Then
        su_Flag = True
    Else
        su_Flag = False
    End If
    
    retval = Application.ScreenUpdating
    Application.ScreenUpdating = su_Flag
    ScreenUpdating = retval
End Function


Function StatusBar_Msg(Optional ByVal AMsg As String = "") As Boolean
'Turn Screen Updating on (to help speed up macro execution) or off (to show screen changes)
'   AMsg = The message to be displayed on the status bar
'           (If AMsg = "" or is missing then Status Bar is set to default mode)
'   retval = Previous value of the ScreenUpdating property (before being set by this function)

    If (IsMissing(AMsg) Or IsEmpty(AMsg) Or AMsg = "") Then
        sb_Msg = False
    Else
        sb_Msg = AMsg
    End If
        
    Application.DisplayStatusBar = True
    
    retval = Application.StatusBar
    'Ensure the Status Bar is turned on
    Application.StatusBar = sb_Msg
    ScreenUpdating_Off = retval

End Function

