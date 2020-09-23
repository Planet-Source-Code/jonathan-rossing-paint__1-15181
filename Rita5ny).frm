VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "New"
   ClientHeight    =   3690
   ClientLeft      =   2025
   ClientTop       =   5220
   ClientWidth     =   2580
   LinkTopic       =   "Form5"
   ScaleHeight     =   3690
   ScaleWidth      =   2580
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1170
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "3600"
      Top             =   210
      Width           =   765
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Left            =   1170
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "3600"
      Top             =   720
      Width           =   765
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   720
      TabIndex        =   2
      Top             =   2940
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Rita5ny).frx":0000
      Left            =   1230
      List            =   "Rita5ny).frx":0016
      TabIndex        =   1
      Text            =   "Color:"
      Top             =   1860
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1650
      TabIndex        =   0
      Top             =   2940
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Image Characteristis"
      Height          =   1185
      Left            =   0
      TabIndex        =   8
      Top             =   1410
      Width           =   2505
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "BackColor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Top             =   480
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Image Dimensions"
      Height          =   1215
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   2445
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Height:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Width:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   750
         Width           =   765
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

If Combo1.Text = "color:" Then
      MsgBox "Color?"
End If
If Combo1.Text = "Black" Then
      Form3.Pic1.BackColor = vbBlack
      Form3.Pic1.Picture = LoadPicture
End If
If Combo1.Text = "Blue" Then
      Form3.Pic1.BackColor = vbBlue
      Form3.Pic1.Picture = LoadPicture
End If
If Combo1.Text = "Red" Then
      Form3.Pic1.BackColor = vbRed
      Form3.Pic1.Picture = LoadPicture
End If
If Combo1.Text = "Green" Then
      Form3.Pic1.BackColor = vbGreen
      Form3.Pic1.Picture = LoadPicture
End If
If Combo1.Text = "White" Then
      Form3.Pic1.BackColor = &HFFFFFF
      Form3.Pic1.Picture = LoadPicture
      End If
   
  
   Form3.Pic1.Width = Text1
   Form3.Pic1.Height = Text2
     Form3.Width = Form3.Pic1.Width
     Form3.Height = Form3.Pic1.Height
   
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text2_Change()
If Text2.Text > 9825 Then
    Text2.Text = 3600
    MsgBox " Max 9825"
End If
End Sub

