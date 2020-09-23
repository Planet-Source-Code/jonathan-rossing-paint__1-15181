VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9915
   ClientLeft      =   0
   ClientTop       =   195
   ClientWidth     =   990
   LinkTopic       =   "form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9915
   ScaleWidth      =   990
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Index           =   10
      Left            =   30
      Picture         =   "Rita4ny).frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "#D rör"
      Top             =   9030
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Index           =   1
      Left            =   30
      Picture         =   "Rita4ny).frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "#D rör"
      Top             =   7230
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Index           =   2
      Left            =   30
      Picture         =   "Rita4ny).frx":056F
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Rita"
      Top             =   3630
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Index           =   3
      Left            =   30
      Picture         =   "Rita4ny).frx":068B
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Bakgrunds färg"
      Top             =   930
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Index           =   4
      Left            =   30
      Picture         =   "Rita4ny).frx":0797
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Rektangel"
      Top             =   1830
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Index           =   5
      Left            =   30
      Picture         =   "Rita4ny).frx":0926
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Hämta färg"
      Top             =   4530
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Index           =   6
      Left            =   30
      Picture         =   "Rita4ny).frx":0A62
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Linje fjäder"
      Top             =   5430
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Index           =   7
      Left            =   30
      Picture         =   "Rita4ny).frx":0B9E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Linje"
      Top             =   6330
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Index           =   8
      Left            =   30
      Picture         =   "Rita4ny).frx":0FE0
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Stjärn mönster"
      Top             =   8130
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Index           =   9
      Left            =   30
      Picture         =   "Rita4ny).frx":111C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Bubblor"
      Top             =   30
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Index           =   0
      Left            =   30
      Picture         =   "Rita4ny).frx":191D
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cirkel"
      Top             =   2730
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Move 0, 0
End Sub
Private Sub Option1_Click(Index As Integer)
If Option1(1).Value = True Then
    Form1.Frame1(0).Visible = True
    Form3.Pic1.FillStyle = 0
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame1(0).Visible = False
End If
If Option1(2).Value = True Then
    Form1.Frame2.Visible = True
    Form3.Pic1.FillStyle = 0
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame2.Visible = False
End If
If Option1(3).Value = True Then
    Form1.Frame3.Visible = True
    Form3.Pic1.FillStyle = 0
    Form3.Pic1.DrawWidth = 1
    Form3.Pic1.MousePointer = 1
Else
    Form1.Frame3.Visible = False
    Form3.Pic1.MousePointer = 99
End If
If Option1(4).Value = True Then
    Form1.Frame4.Visible = True
    Form3.Pic1.FillStyle = 1
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame4.Visible = False
End If
If Option1(5).Value = True Then
    Form1.Frame5.Visible = True
    Form3.Pic1.FillStyle = 0
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame5.Visible = False
End If
If Option1(6).Value = True Then
    Form1.Frame6.Visible = True
    Form3.Pic1.FillStyle = 0
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame6.Visible = False
End If
If Option1(7).Value = True Then
    Form1.Frame7.Visible = True
    Form3.Pic1.FillStyle = 1
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame7.Visible = False
End If
If Option1(8).Value = True Then
    Form1.Frame8.Visible = True
    Form3.Pic1.FillStyle = 0
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame8.Visible = False
End If
If Option1(9).Value = True Then
    Form1.Frame9.Visible = True
    Form3.Pic1.FillStyle = 0
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame9.Visible = False
End If
If Option1(0).Value = True Then
    Form1.Frame10.Visible = True
    Form3.Pic1.FillStyle = 1
    Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame10.Visible = False
End If
If Option1(10).Value = True Then
    'Form1.Frame12.Visible = True
    'Form3.Pic1.FillStyle = 1
    'Form3.Pic1.DrawWidth = 1
Else
    Form1.Frame12.Visible = False
End If
End Sub

