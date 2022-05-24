VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About ProWAGoN"
   ClientHeight    =   4170
   ClientLeft      =   3450
   ClientTop       =   2820
   ClientWidth     =   6555
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6555
   Begin VB.PictureBox pic_hand 
      Height          =   465
      Left            =   7200
      Picture         =   "frmAbout.frx":0E3A
      ScaleHeight     =   405
      ScaleWidth      =   360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2505
      Left            =   75
      Picture         =   "frmAbout.frx":1144
      ScaleHeight     =   2505
      ScaleWidth      =   6375
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1050
      Width           =   6375
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6750
      Top             =   75
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   -75
      ScaleHeight     =   855
      ScaleWidth      =   6705
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6765
      Begin VB.Image Image1 
         Height          =   735
         Left            =   75
         Picture         =   "frmAbout.frx":CE62
         Top             =   75
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ProWAGoN - Program Without A Good Name ;-)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   885
         TabIndex        =   3
         Top             =   300
         Width           =   5235
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4500
         TabIndex        =   2
         Top             =   675
         Width           =   2040
      End
   End
   Begin VB.CommandButton bt_OK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   5025
      TabIndex        =   0
      Top             =   3675
      Width           =   1350
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sngPosY As Single
Dim sngLastY As Single

Private Type TLink
    Caption As String
    URL     As String
    X1      As Single
    X2      As Single
    Y1      As Single
    Y2      As Single
End Type
Private Links(0 To 3) As TLink

Private Declare Function ShellExecuteAPI Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub bt_ok_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    
    On Error Resume Next
        pic1.Font.Name = "Tahoma"
        pic1.Font.Size = 10
    On Error GoTo 0
    sngPosY = pic1.Height

    Timer1.Interval = 50
    Timer1.Enabled = True

End Sub

Private Sub Form_Load()
    
    Label1(1).Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    ' centres the form
    Move ((Screen.Width - Width) / 2), ((Screen.Height - Height) / 2)
    
    Links(0).Caption = "vg4n2q2j02@sneakemail.com"
    Links(0).URL = "mailto:vg4n2q2j02@sneakemail.com"
    
    Links(1).Caption = "www.atletico.com.br"
    Links(1).URL = "http://www.atletico.com.br"
    
    Links(2).Caption = "http://www.staff.uiuc.edu/~ehowes"
    Links(2).URL = "http://www.staff.uiuc.edu/~ehowes"
    
    Links(3).Caption = "Symantec Corporation"
    Links(3).URL = "http://www.symantec.com"

    pic1.MouseIcon = pic_hand.Picture

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Set frmAbout = Nothing
End Sub

Private Sub ShellTo(ByVal sURL As String)
    Const SW_SHOWNORMAL As Long = 1
    If LenB(sURL) = 0 Then Exit Sub
    Call ShellExecuteAPI(0, "open", sURL, 0&, 0&, SW_SHOWNORMAL)
End Sub

Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim i As Integer
    
    pic1.MousePointer = vbDefault
    
    For i = 0 To 3
        If (X >= Links(i).X1 And X <= Links(i).X2) And (Y >= Links(i).Y1 And Y <= Links(i).Y2) Then
            pic1.MousePointer = vbCustom
        End If
    Next

End Sub

Private Sub pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer
    
    For i = 0 To 3
        If (X >= Links(i).X1 And X <= Links(i).X2) And (Y >= Links(i).Y1 And Y <= Links(i).Y2) Then
            Call ShellTo(Links(i).URL)
        End If
    Next

End Sub

Private Sub Timer1_Timer()
    If sngLastY < -pic1.TextHeight("O") Then
        sngPosY = pic1.Height
    End If
    sngPosY = sngPosY - 10
    Call Imprime(sngPosY)
End Sub

Private Sub ImprimeLink(ByVal iQualLink As Integer, ByVal bNewLineAfter As Boolean)

    Dim bPreviousFontUnderline As Boolean
    Dim bPreviousFontColour As Long

    bPreviousFontUnderline = pic1.Font.Underline
    bPreviousFontColour = pic1.ForeColor

    pic1.Font.Underline = True
    pic1.ForeColor = vbBlue

    Links(iQualLink).X1 = pic1.CurrentX
    Links(iQualLink).Y1 = pic1.CurrentY
    
    pic1.Print Links(iQualLink).Caption;
    
    Links(iQualLink).X2 = pic1.CurrentX
    Links(iQualLink).Y2 = pic1.CurrentY + pic1.TextHeight(Links(iQualLink).Caption)
    
    If bNewLineAfter Then pic1.Print

    pic1.Font.Underline = bPreviousFontUnderline
    pic1.ForeColor = bPreviousFontColour

End Sub

Private Sub Imprime(ByVal sngValor As Single)

    pic1.Cls
    pic1.CurrentY = sngValor
    pic1.ScaleLeft = -600
    
    pic1.Print "Created by Christian Haagensen Gontijo, © Copyright 2002-"
    pic1.Print "2003. Problems, bugs, suggestions? Contact me at"
    Call ImprimeLink(0, False)
    pic1.Print ". No spam, please!"
    pic1.Print ""
    
    pic1.Print "Icons from Clube Atlético Mineiro (";
    Call ImprimeLink(1, False)
    pic1.Print "),"
    pic1.Print "the best football team in the world! :-)"
    pic1.Print ""
    
    pic1.Print "Papyrus design by Christian Haagensen Gontijo and Marcelo"
    pic1.Print "Botti Ferri (aka ""Lelocop""), © Copyright 1993-2003."
    pic1.Print ""
    
    pic1.Print "Eric Howes' AGNIS list: © Copyright 2000-2003 Eric L."
    pic1.Print "Howes -- ";
    Call ImprimeLink(2, False)
    pic1.Print "."
    pic1.Print "My sincere thanks to him, for a great website, a terrific"
    pic1.Print "AGNIS utility, and donating his time and attention to test and"
    pic1.Print "give invaluable feedback about this program, and the"
    pic1.Print "patience to play with lots of different Norton trial products."
    pic1.Print "He also created a very nice readme file! :-)"
    pic1.Print ""
    
    pic1.Print "Norton Internet Security (NIS),  Norton Internet Security"
    pic1.Print "Professional (NIS Pro), and Norton Personal Firewall (NPF)"
    pic1.Print "are the trademarked property of ";
    Call ImprimeLink(3, False)
    pic1.Print " --"
    pic1.Print "© Copyright 2000-2003. Symantec has not endorsed or"
    pic1.Print "recommended the use of ProWAGoN with any version of"
    pic1.Print "Norton Internet Security or Norton Personal Firewall."
    pic1.Print ""
    
    pic1.Font.Bold = True: pic1.Print "Disclaimer And License": pic1.Font.Bold = False
    pic1.Print ""
    
    pic1.Print "This program is freeware. You are free to use and modify"
    pic1.Print "the code to suit your needs, but not to distribute modified"
    pic1.Print "versions of it. If you have made changes which you think"
    pic1.Print "are beneficial, or have bug reports, please email me and I"
    pic1.Print "will do my utmost to get a new version."
    pic1.Print ""
    
    pic1.Print "You can freely distribute the zips with the program and the"
    pic1.Print "source code to other ones, but you must distribute them"
    pic1.Font.Bold = True: pic1.Print "in their original state";: pic1.Font.Bold = False
    pic1.Print "."
    pic1.Print ""
    
    pic1.Print "While this program has been tested with success (and seems"
    pic1.Print "to work, actually!), I can't guarantee it will work correctly in"
    pic1.Print "any situation. In other words, ";
    pic1.Font.Italic = True: pic1.Print "use it at your own risk!";: pic1.Font.Italic = False
    
    sngLastY = pic1.CurrentY

End Sub

