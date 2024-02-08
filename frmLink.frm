VERSION 5.00
Begin VB.Form frmLink 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Link"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Open link in a new window"
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Alt:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&URL:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "frmLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim openNewWindow As String

    If Text1.Text = "" Then
        MsgBox "Please type an URL."
        Exit Sub
    End If
    
    If Text3.Text = "" Then
        MsgBox "Please type a name."
        Exit Sub
    End If
    
    If Check1.Value = 1 Then
        openNewWindow = " target=""_blank"""
    Else
        openNewWindow = ""
    End If
    
    frmMain.Text1.SelText = "<a href=""" & Text1.Text & """" & openNewWindow & ">" & Text3.Text & "</a>"
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
