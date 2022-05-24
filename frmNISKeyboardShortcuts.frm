VERSION 5.00
Begin VB.Form frmKeyboardShortcuts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ProWAGoN - firewall keyboard shortcuts and windows names"
   ClientHeight    =   6135
   ClientLeft      =   3090
   ClientTop       =   1980
   ClientWidth     =   8190
   Icon            =   "frmNISKeyboardShortcuts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8190
   Begin VB.CommandButton bt_back 
      Caption         =   "< &Back"
      Height          =   435
      Left            =   3225
      TabIndex        =   18
      Top             =   5580
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Step 3 of 3"
      Height          =   5505
      Index           =   2
      Left            =   8625
      TabIndex        =   24
      Top             =   5700
      Width           =   7920
      Begin VB.TextBox txt_ModifyHtmlStringWindow 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Text            =   "Modify HTML string"
         Top             =   3750
         Width           =   3945
      End
      Begin VB.TextBox txt_AddNewHtmlStringWindow 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Text            =   "Add new HTML string"
         Top             =   1275
         Width           =   3795
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   4
         Left            =   4200
         Picture         =   "frmNISKeyboardShortcuts.frx":0E3A
         Top             =   3450
         Width           =   2685
      End
      Begin VB.Image Image1 
         Height          =   1230
         Index           =   3
         Left            =   4125
         Picture         =   "frmNISKeyboardShortcuts.frx":1494
         Top             =   975
         Width           =   2880
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmNISKeyboardShortcuts.frx":1CE4
         Height          =   450
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   7560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "That's it! Click ""cancel"" to close the dialog. Close the other firewall windows, also."
         Height          =   195
         Index           =   11
         Left            =   150
         TabIndex        =   21
         Top             =   5100
         Width           =   5760
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmNISKeyboardShortcuts.frx":1DBC
         Height          =   495
         Index           =   8
         Left            =   150
         TabIndex        =   14
         Top             =   585
         Width           =   7560
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Step 2 of 3"
      Height          =   5505
      Index           =   1
      Left            =   8925
      TabIndex        =   23
      Top             =   150
      Width           =   7920
      Begin VB.TextBox txt_ModifyAdBlockingShortcut 
         Height          =   315
         Left            =   1515
         MaxLength       =   1
         TabIndex        =   10
         Text            =   "M"
         Top             =   3960
         Width           =   315
      End
      Begin VB.TextBox txt_RemoveAdBlockingShortcut 
         Height          =   315
         Left            =   1515
         MaxLength       =   1
         TabIndex        =   13
         Text            =   "v"
         Top             =   4815
         Width           =   315
      End
      Begin VB.TextBox txt_AddAdBlockingShortcut 
         Height          =   315
         Left            =   1515
         MaxLength       =   1
         TabIndex        =   7
         Text            =   "d"
         Top             =   3120
         Width           =   315
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmNISKeyboardShortcuts.frx":1E6A
         Height          =   570
         Index           =   12
         Left            =   150
         TabIndex        =   4
         Top             =   1875
         Width           =   7380
      End
      Begin VB.Image Image1 
         Height          =   1290
         Index           =   2
         Left            =   1575
         Picture         =   "frmNISKeyboardShortcuts.frx":1F3B
         Top             =   450
         Width           =   4620
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ALT +"
         Height          =   195
         Index           =   7
         Left            =   975
         TabIndex        =   9
         Top             =   3975
         Width           =   435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "The ""&Modify..."" button, on the ""Ad Blocking"" tab, can be accessed using this shortcut:"
         Height          =   195
         Index           =   6
         Left            =   795
         TabIndex        =   8
         Top             =   3660
         Width           =   6120
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ALT +"
         Height          =   195
         Index           =   5
         Left            =   975
         TabIndex        =   12
         Top             =   4875
         Width           =   435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ALT +"
         Height          =   195
         Index           =   3
         Left            =   975
         TabIndex        =   6
         Top             =   3180
         Width           =   435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "The ""A&dd..."" button,  on the ""Ad Blocking"" tab, can be accessed using this shortcut:"
         Height          =   195
         Index           =   2
         Left            =   825
         TabIndex        =   5
         Top             =   2820
         Width           =   6255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "The ""Remo&ve"" button, on the ""Ad Blocking"" tab, can be accessed using this shortcut:"
         Height          =   195
         Index           =   4
         Left            =   795
         TabIndex        =   11
         Top             =   4515
         Width           =   6120
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Step 1 of 3"
      Height          =   5505
      Index           =   0
      Left            =   150
      TabIndex        =   22
      Top             =   0
      Width           =   7920
      Begin VB.TextBox txt_AdBlockingWindow 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Text            =   "Advanced"
         Top             =   4500
         Width           =   3420
      End
      Begin VB.Image Image1 
         Height          =   1065
         Index           =   1
         Left            =   3750
         Picture         =   "frmNISKeyboardShortcuts.frx":2565
         Top             =   4125
         Width           =   4035
      End
      Begin VB.Image Image1 
         Height          =   2745
         Index           =   0
         Left            =   4275
         Picture         =   "frmNISKeyboardShortcuts.frx":2F2C
         Top             =   750
         Width           =   2700
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmNISKeyboardShortcuts.frx":873F
         Height          =   1395
         Index           =   10
         Left            =   150
         TabIndex        =   1
         Top             =   1500
         Width           =   3975
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&What is the name of the ad-blocking window? (not case-sensitive -- see the window below for an example)"
         Height          =   495
         Index           =   9
         Left            =   75
         TabIndex        =   2
         Top             =   3900
         Width           =   7680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please open the Ad-Blocking window on your Norton firewall."
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Width           =   7530
      End
   End
   Begin VB.CommandButton bt_cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   6525
      TabIndex        =   20
      Top             =   5580
      Width           =   1515
   End
   Begin VB.CommandButton bt_next 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Height          =   435
      Left            =   4785
      TabIndex        =   19
      Top             =   5580
      Width           =   1515
   End
End
Attribute VB_Name = "frmKeyboardShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' will store the current (visible) frame
Private miVisibleFrame As Integer

' API used to write to the INI file.
Private Declare Function WritePrivateProfileStringAPI Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'===================================================================================================
' WriteIni
'
' Just a shortcut to call the WritePrivateProfileString API...
'===================================================================================================
Private Sub WriteIni(ByVal secao As String, ByVal key As String, ByVal valor As String)
    Call WritePrivateProfileStringAPI(secao, key, valor, BackSlashGF(App.Path) & INI_FILE)
End Sub

'===================================================================================================
' SaveSettingsOnINI
'
' Write the our INI file the options the user chose
' on the buttons.
'
' This procedure is called from "bt_next_click", so when
' it this executed, the user already selected his/her NIS
' version
'===================================================================================================
Private Sub SaveSettingsOnINI()

    Call WriteIni("SETTINGS", "AdBlockingWindowName", gsAdBlockingWindowName)
    Call WriteIni("SETTINGS", "AddShortcut", gsAddShortcut)
    Call WriteIni("SETTINGS", "ModifyShortcut", gsModifyShortcut)
    Call WriteIni("SETTINGS", "RemoveShortcut", gsRemoveShortcut)
    Call WriteIni("SETTINGS", "AddNewHTMLString", gsAddNewHTMLStringCaption)
    Call WriteIni("SETTINGS", "ModifyHTMLString", gsModifyHMLStringCaption)

End Sub

'===================================================================================================
' ValidStep
'
' Returns TRUE if the required entries were filled by
' the user; otherwise alerts the user and returns FALSE.
'===================================================================================================
Private Function ValidStep() As Boolean

    Select Case miVisibleFrame
    
        '-----------------------------------------------------------------------------
        ' Step 1
        Case 0
        
            ' The step is "valid" if the user
            ' informed the NIS/NPF ad-blocking
            ' window caption

            If txt_AdBlockingWindow.Text = "" Then
                MsgBox "Please enter the ad-blocking main window title.", vbExclamation, "Error"
                txt_AdBlockingWindow.SetFocus
                Exit Function
            End If
            
        '-----------------------------------------------------------------------------
        ' Step 2
        Case 1
        
            ' This step is "valid" only if the user informed
            ' the shortcuts to the "add...", "remove" and
            ' "modify..." buttons
            If txt_AddAdBlockingShortcut.Text = "" Then
                MsgBox "Please enter the ""add..."" button shortcut.", vbExclamation, "Error"
                txt_AddAdBlockingShortcut.SetFocus
                Exit Function
            End If
            If txt_RemoveAdBlockingShortcut.Text = "" Then
                MsgBox "Please enter the ""remove"" button shortcut.", vbExclamation, "Error"
                txt_RemoveAdBlockingShortcut.SetFocus
                Exit Function
            End If
            If txt_ModifyAdBlockingShortcut.Text = "" Then
                MsgBox "Please enter the ""modify..."" button shortcut.", vbExclamation, "Error"
                txt_ModifyAdBlockingShortcut.SetFocus
                Exit Function
            End If
        
        '-----------------------------------------------------------------------------
        ' Step 3
        Case 2
        
            ' This step is "valid" only if the user entered
            ' the windows captions to the "Add new HTML string"
            ' and the "Modify HTML string" dialogs
            If txt_AddNewHtmlStringWindow.Text = "" Then
                MsgBox "Please enter the ""add new HTML string"" dialog title.", vbExclamation, "Error"
                txt_AddNewHtmlStringWindow.SetFocus
                Exit Function
            End If
            If txt_ModifyHtmlStringWindow.Text = "" Then
                MsgBox "Please enter the ""modify HTML string"" dialog title.", vbExclamation, "Error"
                txt_ModifyHtmlStringWindow.SetFocus
                Exit Function
            End If
        
    End Select
    
    ' If we reached this far, the step is valid
    ValidStep = True

End Function

'===================================================================================================
' The "Back" button
'===================================================================================================
Private Sub bt_back_Click()

    '------------------------------------------------------------
    ' Go to the previous frame
    '------------------------------------------------------------
    
    ' Since we are not on the last frame, change the button
    bt_next.Caption = "&Next >"
    
    ' Hide the actual frame
    Frame1(miVisibleFrame).Visible = False
    
    ' Decrement
    miVisibleFrame = miVisibleFrame - 1
    
    ' Show the correct frame
    Frame1(miVisibleFrame).Visible = True
        
    SendKeys "{TAB 3}", True
        
    ' If it's the first frame, disables the button
    If miVisibleFrame = 0 Then
        bt_back.Enabled = False
        Exit Sub
    End If
    

End Sub

'===================================================================================================
' The "Cancel" button.
' I think I don't need to comment what this does. Do I? :-)
'===================================================================================================
Private Sub bt_cancel_Click()
    Unload Me
End Sub

'===================================================================================================
' The "Next" button.
'===================================================================================================
Private Sub bt_next_Click()
    
    '------------------------------------------------------------
    ' Before going to the next step, we have to check if
    ' the user filled the options on the current step
    '------------------------------------------------------------
    If Not ValidStep Then Exit Sub
    
    
    If bt_next.Caption = "Finish" Then
        
        '------------------------------------------------------------
        ' If it's the "Finish" button, adjust properties and exit
        '------------------------------------------------------------

        gsAdBlockingWindowName = txt_AdBlockingWindow.Text
        gsAddShortcut = txt_AddAdBlockingShortcut.Text
        gsModifyShortcut = txt_ModifyAdBlockingShortcut.Text
        gsRemoveShortcut = txt_RemoveAdBlockingShortcut.Text
        gsAddNewHTMLStringCaption = txt_AddNewHtmlStringWindow.Text
        gsModifyHMLStringCaption = txt_ModifyHtmlStringWindow.Text
        
        Call SaveSettingsOnINI

        Unload Me
        
    Else
    
        '------------------------------------------------------------
        ' It's the "Next" button, go to the next frame
        '------------------------------------------------------------
        
        ' Hide the actual frame
        Frame1(miVisibleFrame).Visible = False
        
        ' Increment
        miVisibleFrame = miVisibleFrame + 1
        
        ' Show the correct frame
        Frame1(miVisibleFrame).Visible = True
        
        ' Since we're not on the first frame anymore,
        ' re-enable the "back" button
        bt_back.Enabled = True
        
        ' If it's the last frame, change the button
        If miVisibleFrame = 2 Then bt_next.Caption = "Finish"
        
        SendKeys "{TAB 2}", True
        
    End If
    
End Sub

Private Sub Form_Load()

    Dim i As Integer
    
    ' put all the frames over the first
    With Frame1(0)
        For i = 1 To 2
            Frame1(i).Move .Left, .Top, .Width, .Height
            Frame1(i).Visible = False
        Next
    End With
    
    ' Fill the textboxes with the current options (if any)
    If gsAdBlockingWindowName <> "" Then
        txt_AdBlockingWindow.Text = gsAdBlockingWindowName
        txt_AddAdBlockingShortcut.Text = gsAddShortcut
        txt_ModifyAdBlockingShortcut.Text = gsModifyShortcut
        txt_RemoveAdBlockingShortcut.Text = gsRemoveShortcut
        txt_AddNewHtmlStringWindow.Text = gsAddNewHTMLStringCaption
        txt_ModifyHtmlStringWindow.Text = gsModifyHMLStringCaption
    End If
    
    ' Position the window
    Me.Move gsngLeft, gsngTop

    ' The "back" button is initially disabled
    bt_back.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmKeyboardShortcuts = Nothing
End Sub

Private Sub txt_AdBlockingWindow_GotFocus()
    SelectWholeTextGS txt_AdBlockingWindow
End Sub

Private Sub txt_AddAdBlockingShortcut_GotFocus()
    SelectWholeTextGS txt_AddAdBlockingShortcut
End Sub

Private Sub txt_AddNewHtmlStringWindow_GotFocus()
    SelectWholeTextGS txt_AddNewHtmlStringWindow
End Sub

Private Sub txt_ModifyAdBlockingShortcut_GotFocus()
    SelectWholeTextGS txt_ModifyAdBlockingShortcut
End Sub

Private Sub txt_ModifyHtmlStringWindow_GotFocus()
    SelectWholeTextGS txt_ModifyHtmlStringWindow
End Sub

Private Sub txt_RemoveAdBlockingShortcut_GotFocus()
    SelectWholeTextGS txt_RemoveAdBlockingShortcut
End Sub
