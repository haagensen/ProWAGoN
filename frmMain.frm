VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ProWAGoN"
   ClientHeight    =   6015
   ClientLeft      =   4935
   ClientTop       =   2760
   ClientWidth     =   4575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   4575
   Begin VB.Frame Frame2 
      Caption         =   "Configure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   4395
      Begin VB.CommandButton bt_configure 
         Caption         =   "C&onfigure..."
         Height          =   330
         Left            =   2925
         TabIndex        =   2
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   """Configure"" button."
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   17
         Top             =   675
         Width           =   1365
      End
      Begin VB.Label lbl_About 
         BackStyle       =   0  'Transparent
         Caption         =   "ProWAGoN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   450
         Width           =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "is not working correctly, please click the"
         Height          =   195
         Index           =   3
         Left            =   1050
         TabIndex        =   15
         Top             =   450
         Width           =   3120
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "If you haven't specified your configurations yet, or if"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "What do you want to do?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Index           =   0
      Left            =   75
      TabIndex        =   3
      Top             =   1222
      Width           =   4395
      Begin VB.OptionButton opt_WhatToDo 
         Caption         =   "&BACKUP the currently loaded ad-blocking list to a registry (.reg) file"
         Height          =   405
         Index           =   1
         Left            =   150
         TabIndex        =   5
         Top             =   975
         Width           =   4020
      End
      Begin VB.OptionButton opt_WhatToDo 
         Caption         =   "&ADD items to NIS (an AGNIS registry file, or a backup file)"
         Height          =   405
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   375
         Value           =   -1  'True
         Width           =   4095
      End
      Begin VB.OptionButton opt_WhatToDo 
         Caption         =   "&CLEAR (remove) the currently loaded ad-blocking list"
         Height          =   405
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   1575
         Width           =   4155
      End
   End
   Begin VB.CommandButton bt_DoIt 
      Caption         =   "&Start"
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   5550
      Width           =   1485
   End
   Begin VB.PictureBox Picture1 
      Height          =   465
      Left            =   6900
      Picture         =   "frmMain.frx":0E3A
      ScaleHeight     =   405
      ScaleWidth      =   360
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Frame Frame1 
      Caption         =   "Locate the AGNIS reg file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   1
      Left            =   75
      TabIndex        =   7
      Top             =   3300
      Width           =   4395
      Begin VB.CommandButton bt_browseforfolder 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3900
         Picture         =   "frmMain.frx":1144
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   750
         Width           =   375
      End
      Begin VB.TextBox txt_AgnisRegFile 
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Text            =   "C:\teste.reg"
         Top             =   750
         Width           =   3675
      End
      Begin VB.Label lbl_regfile 
         BackStyle       =   0  'Transparent
         Caption         =   "Please &locate the .reg file you want to add on the ad-blocking list"
         Height          =   420
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   4110
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Open the ""Ad-blocking"" window on your Norton firewall before clicking the button below."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   75
      TabIndex        =   14
      Top             =   4575
      Width           =   4380
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning: do not try to stop the process or switch to other windows while the job is running."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Index           =   0
      Left            =   75
      TabIndex        =   11
      Top             =   5025
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Option buttons
Private Const DO_ADD = 0
Private Const DO_BACKUP = 1
Private Const DO_REMOVE = 2

' Frames
Private Const FRAME_AGNISFILE = 1

' Default file name for the current
' add blocking list
Private Const CURRENTLISTDOTREG As String = "C:\MyCurrentNortonAdBlockingList.reg"

' Some APIs constants we'll need...
Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const VK_RETURN = &HD
Private Const WM_KEYDOWN = &H100

' Some APIs used.
' As a personal convention, I always add "API" in
' front of the function name.
Private Declare Function SendMessageStringAPI Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageAPI Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function PostMessageAPI Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowExAPI Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Sub SleepAPI Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

'===================================================================================================
' DialogClose
'
' Simulates an user clicking ENTER on the dialog, to close it.
'
' "lHandle" is the handle (identifier) from the window.
'===================================================================================================
Private Sub DialogClose(ByVal lHandle As Long)

    ' The WM_KEYUP is not needed here.
    Call PostMessageAPI(lHandle, WM_KEYDOWN, VK_RETURN, 0)
    
    DoEvents

End Sub

'===================================================================================================
' DialogSendText
'
' Sends the specified text to the textbox in the
' "Add New HTML String" dialog
'===================================================================================================
Private Function DialogSendText(ByVal sText As String) As Boolean

    Dim lHandleAddHTMLStringWindow As Long
    Dim lHandleTextbox As Long
    Dim lRetVal As Long
    
    ' Find the "Add New HTML String" dialog
    lHandleAddHTMLStringWindow = FindWindowExAPI(0, 0, _
                                                 vbNullString, _
                                                 gsAddNewHTMLStringCaption)
    If lHandleAddHTMLStringWindow = 0 Then
        MsgBox """" & gsAddNewHTMLStringCaption & _
               """ dialog was not found.", _
               vbCritical + vbMsgBoxSetForeground, "Error"
        Exit Function
    End If
    
    ' Find the textbox on this dialog
    lHandleTextbox = FindWindowExAPI(lHandleAddHTMLStringWindow, _
                                     0&, "Edit", vbNullString)
    If lHandleTextbox = 0 Then
        MsgBox "Edit window on " & gsAddNewHTMLStringCaption & _
               " dialog was not found.", _
               vbCritical + vbMsgBoxSetForeground, "Error"
        Exit Function
    End If
    
    ' Send the specified string to it
    lRetVal = SendMessageStringAPI(lHandleTextbox, WM_SETTEXT, _
                                   0&, ByVal sText)
    If lRetVal = 0 Then
        MsgBox "Could not send text to the edit window on " & _
               gsAddNewHTMLStringCaption & " dialog.", _
               vbCritical + vbMsgBoxSetForeground, "Error"
        Exit Function
    End If
    
    ' Close the dialog
    Call DialogClose(lHandleTextbox)
    
    ' If we reached this far, all's well
    DialogSendText = True

End Function

'===================================================================================================
' DialogGetText
'
' Gets the text that is on the textbox at the "Modify HTML
' String" dialog, and returns it
'===================================================================================================
Private Function DialogGetText() As String

    Dim lHandleModifyHTMLStringWindow As Long
    Dim lHandleTextbox As Long
    Dim lRetVal As Long
    Dim lLength As Long
    Dim sBuffer As String
    
    ' Find the "Modify HTML String" dialog
    lHandleModifyHTMLStringWindow = FindWindowExAPI(0, 0, _
                                                    vbNullString, _
                                                    gsModifyHMLStringCaption)
    If lHandleModifyHTMLStringWindow = 0 Then
        MsgBox """" & gsModifyHMLStringCaption & _
               """ dialog was not found.", _
               vbCritical + vbMsgBoxSetForeground, "Error"
        Exit Function
    End If
    
    ' Find the textbox on this dialog
    lHandleTextbox = FindWindowExAPI(lHandleModifyHTMLStringWindow, _
                                     0&, "Edit", vbNullString)
    If lHandleTextbox = 0 Then
        MsgBox "Edit window on " & gsModifyHMLStringCaption & _
               " dialog was not found.", _
               vbCritical + vbMsgBoxSetForeground, "Error"
        Exit Function
    End If
    
    ' Determine how much space is necessary for the buffer
    ' (1 is added for the terminating null character)
    lLength = SendMessageAPI(lHandleTextbox, WM_GETTEXTLENGTH, _
                             ByVal CLng(0), ByVal CLng(0)) + 1
    If lLength = 0 Then Exit Function
    
    ' Make enough room in the buffer to receive the text
    sBuffer = Space(lLength)
    
    ' Get the string on the textbox
    lRetVal = SendMessageStringAPI(lHandleTextbox, WM_GETTEXT, _
                                   lLength, ByVal sBuffer)
    If lRetVal = 0 Then
        MsgBox "Could not get text from the edit window on " & _
               gsModifyHMLStringCaption & " dialog.", _
               vbCritical + vbMsgBoxSetForeground, "Error"
        Exit Function
    End If
    
    ' Remove the terminating null and extra space from the buffer
    sBuffer = Left$(sBuffer, lRetVal)
    
    ' Close the dialog
    Call DialogClose(lHandleTextbox)
    
    ' Return value
    DialogGetText = sBuffer

End Function

'===================================================================================================
' SendKeystroke
'
' Sends the keystroke SKEYSTROKE and makes a little pause
'===================================================================================================
Private Sub SendKeystroke(ByVal sKeystroke As String)

    '----------------------------------------------------------
    ' Sends the keystroke.
    ' I believe "SendKeys" is based on the "keybd_event" API
    ' function (tip for you using Delphi and trying to mimic
    ' VB SendKeys). ;-)
    '
    ' NOTE:
    '
    ' The VB documentation states that the "Sendkeys" function
    ' can be used to send more than one keystroke at a time,
    ' using the format "{Key NumberOfTimes}", so instead of
    ' using, for example,
    '      SendKeys "{TAB}"
    '      SendKeys "{TAB}"
    ' we could use
    '      SendKeys "{TAB 2}"
    '
    ' I tried this on the program, but it just doesn't work,
    ' since there's no pause between each key sent. So I wrote
    ' the whole program to send only one at a time.
    '----------------------------------------------------------
    
    ' The "DoEvents" here is essential, it lets
    ' Windows process other tasks
    DoEvents
    
    SendKeys sKeystroke, True
    
    ' again...
    DoEvents

End Sub

'===================================================================================================
' IsAdBlockingWindowOpened
'
' Check if the "ad-blocking" window is open, and put the
' focus on the "ad blocking" tab.
'
' Returns FALSE if the window could not be found,
' TRUE otherwise.
'===================================================================================================
Private Function IsAdBlockingWindowOpened() As Boolean

    Const pause As Long = 600
    Dim lWindowHandle As Long
    Dim i As Integer

    '----------------------------------------------------------
    ' Look for NIS main window
    '----------------------------------------------------------
    lWindowHandle = FindWindowExAPI(0, 0, vbNullString, _
                                           gsAdBlockingWindowName)
    If lWindowHandle = 0 Then
        ' ops!
        Screen.MousePointer = vbDefault
        MsgBox "Couldn't find the window """ & _
               gsAdBlockingWindowName & """!" & NLNL & _
               "Make sure it is open, and try again.", _
               vbCritical + vbMsgBoxSetForeground, "Error"
        Exit Function
    End If

    '----------------------------------------------------------
    ' Activates it
    '----------------------------------------------------------
    On Error Resume Next
        AppActivate gsAdBlockingWindowName, True
    On Error GoTo 0

    '----------------------------------------------------------
    ' "Select" the ad-blocking tab
    '----------------------------------------------------------
    SendKeystroke "{TAB}"
    SendKeystroke "{TAB}"
    
    '----------------------------------------------------------
    ' If we reached this far, everything's ok
    '----------------------------------------------------------
    IsAdBlockingWindowOpened = True

End Function

'===================================================================================================
' AddEntries
'
' Opens the NIS window and sends each entry on the registry
' file (an AGNIS list, or a backup list) to the NIS "Ad
' Blocking" list.
'===================================================================================================
Private Sub AddEntries()

    Dim f As Integer              ' file number available for use by the Open statement
    Dim sTextLine As String       ' a line from AGNIS .reg file
    Dim bFound As Boolean         ' a sanity check flag
    Dim iPos As Integer           ' position where the string '"=' is found in the line
    Dim sngStartTime As Single    ' benchmarking: the start time
    Dim sngEndTime As Single      ' benchmarking: the end time
    Dim iRemEntries As Integer    ' counter: the number of commented entries on the ad-block list
    Dim iSentEntries As Integer   ' counter: the number of entries we sent to the ad-block list


    Screen.MousePointer = vbHourglass

    '----------------------------------------------------------
    ' Try to open the given file for input (reading)
    '----------------------------------------------------------
    On Error Resume Next
        f = FreeFile
        Open txt_AgnisRegFile.Text For Input As #f
        If Err.Number <> 0 Then
            ' ops!
            MsgBox "Error opening registry file: " & Error$, _
                   vbCritical + vbMsgBoxSetForeground, "Error"
            GoTo CleanUp ' yeah, I *KNOW* "GoTo" sux, see the end of this sub :-)
        End If
    On Error GoTo 0
    
    '----------------------------------------------------------
    ' Sanity check: we'll verify if the string
    ' "[HKEY_LOCAL_MACHINE\SOFTWARE\Symantec\IAM\HTTPConfig\Sites\(Defaults)\Block]"
    ' exists on the given file.
    '----------------------------------------------------------
    Do While Not EOF(f)
        Line Input #1, sTextLine
        If sTextLine = "[HKEY_LOCAL_MACHINE\SOFTWARE\Symantec\IAM\HTTPConfig\Sites\(Defaults)\Block]" Then
            bFound = True
            Exit Do
        End If
    Loop
    
    ' String found?
    If Not bFound Then
        MsgBox "The given file is not an ad-blocking " & _
               "registry file for Norton Internet Security!", _
               vbCritical + vbMsgBoxSetForeground, "Error"
        GoTo CleanUp
    End If

    '----------------------------------------------------------
    ' Make sure we have the right NIS window, then bring
    ' it to the front
    '----------------------------------------------------------
    If Not IsAdBlockingWindowOpened Then GoTo CleanUp
    
    '----------------------------------------------------------
    ' Now we'll read each line, pick the entry, format
    ' it, and put each one in the NIS window
    '----------------------------------------------------------
    sngStartTime = Timer
    
    Do While Not EOF(f)
        
        ' read an ad-block entry
        Line Input #1, sTextLine
        
        ' we'll ignore commented lines (a commented line
        ' has a ";" in front of it)
        If Left$(sTextLine, 1) = ";" Then
        
            iRemEntries = iRemEntries + 1
        
        Else
        
            ' find where the '"=hex:01' part is
            iPos = InStr(1, sTextLine, """=hex")
            
            ' Do the steps below only if it was found
            ' (maybe we're reading a blank line...)
            If iPos > 0 Then
            
                ' format the line, removing the initial '"' and
                ' the final '"=hex:01' on it (for you curious,
                ' the "hex=01" means "block" to NIS) :-)
                sTextLine = Mid$(sTextLine, 2, iPos - 2)
                
                ' "click" the "Add" button on the "ad blocking"
                ' panel on NIS
                SendKeystroke "%" & gsAddShortcut
                    
                ' Sends the ad-block entry to the textbox on that
                ' window, and close the dialog. The "block" radio
                ' button is already checked by NIS, so we don't
                ' have to bother with it.
                If Not DialogSendText(sTextLine) Then
                    ' some error happened
                    GoTo CleanUp
                End If
                
                ' increment counter
                iSentEntries = iSentEntries + 1
                    
                ' a little pause
                DoEvents
                
            End If ' iPos > 0
            
        End If ' Left$(1, sTextLine) <> ";"
        
    Loop
    
    sngEndTime = Timer
    
    '----------------------------------------------------------
    ' Done :-)
    '----------------------------------------------------------
    
    ' bring the focus back to ProWAGoN
    On Error Resume Next
        AppActivate Me.Caption
    On Error GoTo 0
    
    ' Alert the user
    MsgBox CStr(iSentEntries) & " registry entries sent to " & _
           "the firewall, " & _
           CStr(iRemEntries) & " commented lines ignored. " & _
           vbNewLine & "Time taken: " & _
           Format$(sngEndTime - sngStartTime, "###.##") & _
           " seconds." & NLNL & _
           "Press the ""OK"" button on your Norton firewall " & _
           "window to make the added entries permanent.", _
           vbInformation + vbMsgBoxSetForeground, "Done"

CleanUp:
    
    Screen.MousePointer = vbDefault

    '----------------------------------------------------------
    ' Close the file and bail out.
    ' <rant> We VBers had to live with those terrible "on
    ' error goto..." things, since VB5/VB6 don't have a
    ' decent error trapping. Fortunatelly VB7 ("VB.net") has
    ' a "try... catch... finally" block. </rant> But hey, what
    ' am I talking about -- this program is already a dirty
    ' hack, right? :-)
    '----------------------------------------------------------
    On Error Resume Next
        Close #f
    On Error GoTo 0

End Sub

'===================================================================================================
' BackupEntries
'
' Sends the current entries on NIS ad-block list to a REG file
'===================================================================================================
Private Sub BackupEntries()

    Dim f As Integer                 ' file number available for use by the Open statement
    Dim sCurrentEntry As String      ' a line on the ad-block list
    Dim sLastEntry As String         ' last line we read on the ad-block list
    Dim iTotalEntries As Integer     ' number of entries on the ad-block list
    Dim sngStartTime As Single       ' benchmarking: the start time
    Dim sngEndTime As Single         ' benchmarking: the end time
    

    Screen.MousePointer = vbHourglass

    '----------------------------------------------------------
    ' Make sure we have the right window, then
    ' bring it to the front
    '----------------------------------------------------------
    If Not IsAdBlockingWindowOpened Then Exit Sub
    
    '----------------------------------------------------------
    ' Select the first entry on the list, to enable the
    ' "modify" button (it's disabled)
    '----------------------------------------------------------
    SendKeystroke "{TAB}"
    SendKeystroke "{DOWN}"
    SendKeystroke "{HOME}"
    
    '----------------------------------------------------------
    ' Try to open the given file for output (writing)
    '----------------------------------------------------------
    On Error Resume Next
        f = FreeFile
        Open txt_AgnisRegFile.Text For Output Lock Read Write As #f
        If Err.Number <> 0 Then
            ' ops!
            MsgBox "Error creating the registry file: " & _
                   Error$, vbCritical + vbMsgBoxSetForeground, "Error"
            Exit Sub
        End If
    On Error GoTo 0
    
    '----------------------------------------------------------
    ' Write the header in our registry file
    '----------------------------------------------------------
    On Error GoTo ErrorWritingToTheFile
        Print #f, "REGEDIT4"
        Print #f, ""
        Print #f, "; This is the current ad block list for your Norton firewall,"
        Print #f, "; created by " & APPNAME & " on ";
        Print #f, Format$(Now, "Short Date") & ", " & Format$(Now, "Long Time")
        Print #f, ";"
        Print #f, "; ********************** WARNING **********************"
        Print #f, ";   -DO NOT- double click this file if you are using"
        Print #f, ";   NIS2002Pro, or any Norton 2003 firewall and"
        Print #f, ";   later versions, or you may SERIOUSLY DAMAGE your"
        Print #f, ";   firewall configurations!!!"
        Print #f, ";   To have these entries added on your Norton"
        Print #f, ";   firewall please use the " & APPNAME & " utility!"
        Print #f, "; *****************************************************"
        Print #f, ""
        Print #f, "[HKEY_LOCAL_MACHINE\SOFTWARE\Symantec\IAM\HTTPConfig\Sites\(Defaults)\Block]"
        Print #f, ""
    On Error GoTo 0
    
    '----------------------------------------------------------
    ' Now add each existing entry on it
    '----------------------------------------------------------
    sngStartTime = Timer

    Do
        
        ' We'll click the "Modify..." button to read an
        ' ad-block entry on NIS window
        SendKeystroke "%" & gsModifyShortcut
            
        ' Open the "Modify HTML string" dialog,
        ' gets its text and close the dialog
        sCurrentEntry = DialogGetText
        
        ' If we couldn't get the text, something's wrong
        If sCurrentEntry = "" Then
            Screen.MousePointer = vbDefault
            MsgBox "Couldn't get the entry from """ & _
                   gsModifyHMLStringCaption & """!", _
                   vbCritical + vbMsgBoxSetForeground, "Error"
            Close #f
            Exit Sub
        End If
        
        ' If this was read before, we already reached the end
        ' of the list, so we're done
        If sLastEntry = sCurrentEntry Then
            Exit Do
        End If
        
        ' Otherwise, put it on our registry file surrounded by
        ' quotes (the "hex=01" means "block" to NIS)
        Print #f, """" & sCurrentEntry & """=hex:01"
        
        ' Now the "current" line is the "last" read
        sLastEntry = sCurrentEntry
        
        ' Increment the counter
        iTotalEntries = iTotalEntries + 1
        
        ' Next item on the list
        SendKeystroke "{DOWN}"
        
        ' A little pause
        DoEvents
        
    Loop
    
    sngEndTime = Timer
    
    '----------------------------------------------------------
    ' Close the file
    '----------------------------------------------------------
    On Error Resume Next
        Close #f
    On Error GoTo 0
    
    '----------------------------------------------------------
    ' Done :-)
    '----------------------------------------------------------
    
    Screen.MousePointer = vbDefault
    
    ' Bring the focus back to this program
    On Error Resume Next
        AppActivate Me.Caption
    On Error GoTo 0
    
    ' Alert the user
    MsgBox "Current ad-blocking list (" & CStr(iTotalEntries) & _
           " entries) sent to the registry file, time taken: " & _
           Format$(sngEndTime - sngStartTime, "###.##") & _
           " seconds." & NLNL & _
           "Please press the ""OK"" button on your Norton " & _
           "firewall window to close it.", _
           vbInformation + vbMsgBoxSetForeground, "Done"

    Exit Sub
    
ErrorWritingToTheFile:
    Screen.MousePointer = vbDefault
    MsgBox "Error writing to the .reg file: " & _
           Error$ & " - operation terminated.", _
           vbCritical + vbMsgBoxSetForeground, "Error"

End Sub

'===================================================================================================
' RemoveEntries
'
' Opens the NIS window and removes ALL the entries on the
' "Ad Blocking" list.
'===================================================================================================
Private Sub RemoveEntries()

    Screen.MousePointer = vbHourglass
    
    '----------------------------------------------------------
    ' Make sure we have the right NIS window, then
    ' bring it to the front
    '----------------------------------------------------------
    If Not IsAdBlockingWindowOpened Then Exit Sub

    '----------------------------------------------------------
    ' Select an entry on the list, to enable the
    ' "remove" button (it's disabled)
    '----------------------------------------------------------
    SendKeystroke "{TAB}"
    SendKeystroke "{DOWN}"
    SendKeystroke "{HOME}"
    
    '----------------------------------------------------------
    ' Select all the entries on the ad-blocking list
    '----------------------------------------------------------
    SendKeystroke "+{END}"
    
    '----------------------------------------------------------
    ' "Click" the "Remove" button on the "ad blocking"
    ' panel on NIS
    '----------------------------------------------------------
    SendKeystroke "%" & gsRemoveShortcut
    
    '----------------------------------------------------------
    ' NIS will pop up a dialog asking for confirmation.
    ' ENTER means "yes".
    '----------------------------------------------------------
    SendKeystroke "{ENTER}"

    '----------------------------------------------------------
    ' Bring the focus back to this program
    '----------------------------------------------------------
    On Error Resume Next
        AppActivate Me.Caption
    On Error GoTo 0
    Screen.MousePointer = vbDefault

    '----------------------------------------------------------
    ' Alert the user
    '----------------------------------------------------------
    MsgBox "All the entries on the ad-blocking list were " & _
           "removed!" & NLNL & "Press the ""OK"" button on your " & _
           "Norton firewall window to make the clearing permanent.", _
           vbInformation + vbMsgBoxSetForeground, "Done"

End Sub

'===================================================================================================
' The "Browse For Folders" button - the button where the
' user will choose the directory/file where the registry
' file is/will be.
'===================================================================================================
Private Sub bt_browseforfolder_Click()

    Dim sOpen As SelectedFile
    Dim s As String, i As Integer
    
    If opt_WhatToDo(DO_BACKUP).Value = True Then

        ' User will backup his current ad-blocking list. Now
        ' he'll choose a directory/file where the registry
        ' file will be
        With FileDialog
            .sFilter = "Registry files (*.reg)" & Chr$(0) & "*.reg" & Chr$(0)
            .flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
            .sDlgTitle = "Please choose where to write the registry file with the backup"
            .sInitDir = BackSlashGF(App.Path)
            .sFile = CURRENTLISTDOTREG
            .sDefFileExt = "reg"
        End With

    Else

        ' User will choose an AGNIS registry file
        With FileDialog
            .sFilter = "Registry Files (*.reg)" & Chr$(0) & "*.reg" & Chr$(0) & _
                       "All Files (*.*)" & Chr$(0) & "*.*"
            .flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
            .sDlgTitle = "Please locate the registry file with the ad-blocking entries"
            .sInitDir = BackSlashGF(App.Path)
        End With

    End If

    ' Show the dialog
    sOpen = ShowOpen(Me.hWnd)

    ' If he not hit cancel...
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        txt_AgnisRegFile.Text = BackSlashGF(sOpen.sLastDirectory) & sOpen.sFiles(1)
    End If

End Sub

'===================================================================================================
' The "configure" button
'===================================================================================================
Private Sub bt_configure_Click()
    frmKeyboardShortcuts.Show 1
End Sub

'===================================================================================================
' "Do It" button
'===================================================================================================
Private Sub bt_DoIt_Click()

    ' disable this button so it can't be accidentally
    ' pressed again
    bt_DoIt.Enabled = False
    
    ' Do what we chose to do
    If opt_WhatToDo(DO_ADD).Value = True Then
        Call AddEntries
    ElseIf opt_WhatToDo(DO_REMOVE).Value = True Then
        Call RemoveEntries
    Else
        Call BackupEntries
    End If
    
    ' re-enable the button
    bt_DoIt.Enabled = True
    
End Sub

Private Sub Form_Load()

    Dim aux As String

    ' To make the "about" label behave like a "internet link",
    ' the mouse pointer will be a "hand" when over it
    With lbl_About
        .MousePointer = vbCustom
        .MouseIcon = Picture1.Picture
    End With
    
    ' Position the window
    Me.Move gsngLeft, gsngTop
    
    ' The default option
    Call opt_WhatToDo_Click(DO_ADD)
    
End Sub

Private Sub lbl_About_Click()
    ' OK - now you know, don't bother Eric about problems
    ' with THIS program. :-)
    frmAbout.Show vbModal
End Sub

Private Sub opt_WhatToDo_Click(Index As Integer)

    ' Hides the "choose dir/file" frame if the "clear"
    ' option is selected
    Frame1(FRAME_AGNISFILE).Visible = (Index <> DO_REMOVE)

    If Index = DO_BACKUP Then
        
        ' With the "backup" option, the textbox will provide
        ' a space where a registry file will be created
        Frame1(FRAME_AGNISFILE).Caption = "Choose directory/filename for backup"
        lbl_regfile.Caption = "Please choose a directory and a fi&le name for the backup file."
        txt_AgnisRegFile.Text = CURRENTLISTDOTREG
    
    Else
        
        ' With the "add" option, the textbox will provide
        ' a space where the user will choose an registry file to be
        ' added to NIS/NPF
        Frame1(FRAME_AGNISFILE).Caption = "Locate the registry file"
        lbl_regfile.Caption = "Please &locate the ad-block list file you wish to add."
        txt_AgnisRegFile.Text = "C:\nis-ads.reg"
    
    End If

End Sub

Private Sub txt_AgnisRegFile_GotFocus()
    SelectWholeTextGS txt_AgnisRegFile
End Sub
