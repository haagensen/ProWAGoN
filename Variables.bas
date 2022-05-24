Attribute VB_Name = "Variables"
Option Explicit

'===============================================================
' On my "professional" programs I use global variables the less
' that I can. I use the forms as classes, adding custom
' properties to it, etc. But come on, a program that is already
' a hack, using a function like VB's "SendKeys", can't be too
' serious. ;-)
'
' As a programming convention of mine, all my "global functions"
' have a "GF" appended to its name, so I can quickly know that
' it's declared as a public function. Likewise, global subs have
' a "GS" in its name.
'===============================================================

' Just a shortcut :-)
Public Const NLNL = vbNewLine & vbNewLine

' Name of the program
Public Const APPNAME = "ProWAGoN" ' "PROgram Without A GOod Name" :-)

' our INI file
Public Const INI_FILE = "PROWAGON.INI"

' Windows position
Public gsngLeft As Single
Public gsngTop As Single

'-----------------------------------------------------------------
' The following will store the program options on an INI file
'-----------------------------------------------------------------

' Name of the window with the ad-blocking tab
Public gsAdBlockingWindowName As String

' The shortcut to the "Add..." button on the
' "Ad Blocking" tab, on NIS Advanced Options window
Public gsAddShortcut As String

' The shortcut to the "Modify..." button on the
' "Ad Blocking" tab, on NIS Advanced Options window
Public gsModifyShortcut As String

' The shortcut to the "Remove" button on the
' "Ad Blocking" tab, on NIS Advanced Options window
Public gsRemoveShortcut As String

' Caption of the "Add New HTML String" dialog
Public gsAddNewHTMLStringCaption As String

' Caption of the "Modify HMTL String" dialog
Public gsModifyHMLStringCaption As String

'========================================================================================
' BackSlashGF
'
' Adds a "\" to the end of the string, if it doesn't have one, and
' returns that string
'========================================================================================
Public Function BackSlashGF(ByVal arq As String) As String
    If Right$(arq, 1) <> "\" Then
        BackSlashGF = arq & "\"
    Else
        BackSlashGF = arq
    End If
End Function

'========================================================================================
' SelectWholeTextGS
'
' Selects the contents of the given textbox
'========================================================================================
Public Sub SelectWholeTextGS(ByVal textboxname As TextBox)
    textboxname.SelStart = 0
    textboxname.SelLength = Len(textboxname.Text)
End Sub
