VERSION 5.00
Begin VB.Form frmImage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Image"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
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
   ScaleHeight     =   3450
   ScaleWidth      =   5385
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&Height:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Width:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Alt:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   600
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
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim imgWidth As String
Dim imgHeight As String

    If Text1.Text = "" Then
        MsgBox "Please type a valid URL."
        Exit Sub
    End If
    
    If Text3.Text = "" Then
        imgWidth = ""
    Else
        imgWidth = " width=" & Val(Text3.Text)
    End If
    
    If Text4.Text = "" Then
        imgHeight = ""
    Else
        imgHeight = " height=" & Val(Text4.Text)
    End If
    
    frmMain.Text1.SelText = "<img src=""" & Text1.Text & """" & imgWidth & imgHeight & " />"
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
