VERSION 5.00
Begin VB.Form frmWelcome 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ProWAGoN"
   ClientHeight    =   4485
   ClientLeft      =   5160
   ClientTop       =   2715
   ClientWidth     =   4650
   Icon            =   "frmWelcome.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4650
   Begin VB.PictureBox Picture1 
      Height          =   465
      Left            =   5625
      Picture         =   "frmWelcome.frx":0E3A
      ScaleHeight     =   405
      ScaleWidth      =   360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton bt_Next 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Height          =   375
      Left            =   2925
      TabIndex        =   0
      Top             =   3900
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWelcome.frx":1144
      Height          =   1005
      Index           =   5
      Left            =   150
      TabIndex        =   7
      Top             =   2595
      Width           =   4350
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "can add the AGNIS list to your Norton"
      Height          =   195
      Index           =   4
      Left            =   1800
      TabIndex        =   6
      Top             =   2400
      Width           =   2670
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   900
      TabIndex        =   5
      Top             =   2400
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWelcome.frx":124C
      Height          =   885
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   480
      Width           =   4275
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWelcome.frx":130B
      Height          =   780
      Index           =   2
      Left            =   150
      TabIndex        =   3
      Top             =   1425
      Width           =   4365
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Until now."
      Height          =   195
      Index           =   3
      Left            =   150
      TabIndex        =   2
      Top             =   2400
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' API used to get values from the INI file.
Private Declare Function GetPrivateProfileStringAPI Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'===================================================================================================
' GetIni
'
' Just a (big) shortcut to call the GetPrivateProfileString API...
'===================================================================================================
Private Function GetIni(ByVal secao As String, ByVal key As String) As String
    Dim RetVal As String, ret As String, worked As Integer
    RetVal = Space$(255)
    worked = GetPrivateProfileStringAPI(secao, key, "", RetVal, Len(RetVal), BackSlashGF(App.Path) & INI_FILE)
    If worked = 0 Then
        ret = ""
    Else
        ret = Trim$(Left$(RetVal, worked))
    End If
    GetIni = ret
End Function

'===================================================================================================
' ReadSettingsFromINI
'
' Fill our working variables with the settings the
' user once chose
'===================================================================================================
Private Sub ReadSettingsFromINI()

    Dim aux As String

    ' If it's the first time we're running this program,
    ' there is no INI file. We'll keep the variables empty,
    ' so the user will be *forced* to choose a NIS version
    ' in the "Configurations" window.
    gsAdBlockingWindowName = GetIni("SETTINGS", "AdBlockingWindowName")
    If gsAdBlockingWindowName = "" Then
        ' No value? We'll assume there's no INI file,
        ' opens the "config" window.
        frmKeyboardShortcuts.Show 1
    End If
    
    ' If we reached this far, we have our settings.
    gsAdBlockingWindowName = GetIni("SETTINGS", "AdBlockingWindowName")
    gsAddShortcut = GetIni("SETTINGS", "AddShortcut")
    gsModifyShortcut = GetIni("SETTINGS", "ModifyShortcut")
    gsRemoveShortcut = GetIni("SETTINGS", "RemoveShortcut")
    gsAddNewHTMLStringCaption = GetIni("SETTINGS", "AddNewHTMLString")
    gsModifyHMLStringCaption = GetIni("SETTINGS", "ModifyHTMLString")
    
End Sub

Private Sub bt_next_Click()

    
    gsngLeft = Me.Left
    gsngTop = Me.Top
    
    ' Read previous values from the INI, if exists
    Call ReadSettingsFromINI
    
    
    frmMain.Show

    Unload Me
    
End Sub

Private Sub Form_Load()
    
    ' To make the "ProWAGoN" label behave like a "internet link",
    ' the mouse pointer will be a "hand" when over it
    With lbl_About
        .MousePointer = vbCustom
        .MouseIcon = Picture1.Picture
    End With
    
End Sub

Private Sub lbl_About_Click()
    ' OK - now you know, don't bother Eric about problems
    ' with THIS program. :-)
    frmAbout.Show vbModal
End Sub
