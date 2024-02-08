VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "HTML Editor by ino|bambino"
   ClientHeight    =   6630
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9165
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   8895
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditBold 
         Caption         =   "&Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEditUnderline 
         Caption         =   "&Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuEditItalic 
         Caption         =   "&Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu ed0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditTextStyle 
         Caption         =   "&Text Style"
         Begin VB.Menu mnuEditCenter 
            Caption         =   "&Center"
            Shortcut        =   ^W
         End
         Begin VB.Menu ed2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditList 
            Caption         =   "&List"
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuEditQuote 
            Caption         =   "&Quote"
            Shortcut        =   ^Q
         End
         Begin VB.Menu mnuEditDefinition 
            Caption         =   "&Definition"
            Shortcut        =   ^D
         End
         Begin VB.Menu ed4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditEmphasize 
            Caption         =   "&Emphasize"
         End
         Begin VB.Menu mnuEditStrong 
            Caption         =   "&Strong"
         End
         Begin VB.Menu mnuEditCode 
            Caption         =   "C&ode"
            Shortcut        =   ^G
         End
         Begin VB.Menu ed3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditAdvanced 
            Caption         =   "Ad&vanced"
            Begin VB.Menu mnuEditCitation 
               Caption         =   "Citatio&n"
            End
            Begin VB.Menu mnuEditAddress 
               Caption         =   "&Address"
            End
            Begin VB.Menu mnuEditExample 
               Caption         =   "&Example"
            End
            Begin VB.Menu mnuEditKeyboard 
               Caption         =   "&Keyboard"
            End
            Begin VB.Menu mnuEditVariable 
               Caption         =   "&Variable"
            End
            Begin VB.Menu ed5 
               Caption         =   "-"
            End
            Begin VB.Menu mnuEditRevisionDelete 
               Caption         =   "Revision &Delete"
            End
            Begin VB.Menu mnuEditRevisionInsert 
               Caption         =   "Re&vision Insert"
            End
         End
      End
      Begin VB.Menu mnuEditCaseStyle 
         Caption         =   "&Case Style"
         Begin VB.Menu mnuEditCaseStypeUpperCase 
            Caption         =   "&Uppercase"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuEditCaseStypeLowerCase 
            Caption         =   "&Lowercase"
            Shortcut        =   +{F3}
         End
      End
      Begin VB.Menu mnuEditHeaderStyle 
         Caption         =   "&Header Style"
      End
      Begin VB.Menu ed1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "I&nsert"
      Begin VB.Menu mnuInsertLink 
         Caption         =   "&Link"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuInsertImage 
         Caption         =   "I&mage"
         Shortcut        =   ^M
      End
      Begin VB.Menu in0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertListCreator 
         Caption         =   "&List Creator"
      End
      Begin VB.Menu in1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertPageEntities 
         Caption         =   "&Page Entities"
         Begin VB.Menu mnuInsertParBreak 
            Caption         =   "&Paragraph Break"
            Shortcut        =   ^E
         End
      End
      Begin VB.Menu in2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertHorLine 
         Caption         =   "&Horizontal Line"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuExtra 
      Caption         =   "E&xtra"
      Begin VB.Menu mnuExtraOpenURL 
         Caption         =   "Open &URL"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuRGB2Hex 
         Caption         =   "&RGB to Hex"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
On Error Resume Next
    Text1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub mnuEditAddress_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<ADDRESS></ADDRESS>"
    Else
        Text1.SelText = "<ADDRESS>" & Text1.SelText & "</ADDRESS>"
    End If
End Sub

Private Sub mnuEditBold_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<B></B>"
    Else
        Text1.SelText = "<B>" & Text1.SelText & "</B>"
    End If
End Sub

Private Sub mnuEditCaseStypeLowerCase_Click()
    Text1.SelText = LCase$(Text1.SelText)
End Sub

Private Sub mnuEditCaseStypeUpperCase_Click()
    Text1.SelText = UCase$(Text1.SelText)
End Sub

Private Sub mnuEditCenter_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<CENTER></CENTER>"
    Else
        Text1.SelText = "<CENTER>" & Text1.SelText & "</CENTER>"
    End If
End Sub

Private Sub mnuEditCitation_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<CITE></CITE>"
    Else
        Text1.SelText = "<CITE>" & Text1.SelText & "</CITE>"
    End If
End Sub

Private Sub mnuEditCode_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<CODE></CODE>"
    Else
        Text1.SelText = "<CODE>" & Text1.SelText & "</CODE>"
    End If
End Sub

Private Sub mnuEditDefinition_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<DD></DD>"
    Else
        Text1.SelText = "<DD>" & Text1.SelText & "</DD>"
    End If
End Sub

Private Sub mnuEditEmphasize_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<EM></EM>"
    Else
        Text1.SelText = "<EM>" & Text1.SelText & "</EM>"
    End If
End Sub

Private Sub mnuEditExample_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<SAMP></SAMP>"
    Else
        Text1.SelText = "<SAMP>" & Text1.SelText & "</SAMP>"
    End If
End Sub

Private Sub mnuEditItalic_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<I></I>"
    Else
        Text1.SelText = "<I>" & Text1.SelText & "</I>"
    End If
End Sub

Private Sub mnuEditKeyboard_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<KBD></KBD>"
    Else
        Text1.SelText = "<KBD>" & Text1.SelText & "</KBD>"
    End If
End Sub

Private Sub mnuEditList_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<LI>"
    Else
        Text1.SelText = "<LI>" & Text1.SelText
    End If
End Sub

Private Sub mnuEditQuote_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<DL><DD><I></I></DD></DL>"
    Else
        Text1.SelText = "<DL><DD><I>" & Text1.SelText & "</I></DD></DL>"
    End If
End Sub

Private Sub mnuEditRevisionDelete_Click()
MsgBox CreateHashEx(32)
End Sub

Private Sub mnuEditRevisionInsert_Click()
ShellEx "http://www.google.com", essSW_SHOWNORMAL, , "c:\", , Me.hWnd
End Sub

Private Sub mnuEditSelectAll_Click()
    SendKeys "{home}+{end}"
End Sub

Private Sub mnuEditStrong_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<STRONG></STRONG>"
    Else
        Text1.SelText = "<STRONG>" & Text1.SelText & "</STRONG>"
    End If
End Sub

Private Sub mnuEditUnderline_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<U></U>"
    Else
        Text1.SelText = "<U>" & Text1.SelText & "</U>"
    End If
End Sub

Private Sub mnuEditVariable_Click()
    If Text1.SelLength = 0 Then
        Text1.SelText = "<VAR></VAR>"
    Else
        Text1.SelText = "<VAR>" & Text1.SelText & "</VAR>"
    End If
End Sub

Private Sub mnuExtraOpenURL_Click()
    frmOpenURL.Show 1
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuInsertHorLine_Click()
    Text1.SelText = "<HR>"
End Sub

Private Sub mnuInsertImage_Click()
    frmImage.Show 1
End Sub

Private Sub mnuInsertLink_Click()
    If Text1.SelLength = 0 Then
        frmLink.Show 1
    Else
        frmLink.Text3.Text = Text1.SelText
        frmLink.Show 1
    End If
    
End Sub

Private Sub mnuInsertParBreak_Click()
    Text1.SelText = "<!--BREAK-->"
End Sub
