VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tool Box"
   ClientHeight    =   2940
   ClientLeft      =   6735
   ClientTop       =   2085
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   2760
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Eraiser"
      Height          =   2115
      Left            =   510
      TabIndex        =   47
      Top             =   540
      Width           =   1665
      Visible         =   0   'False
      Begin ComctlLib.Slider Slider5 
         Height          =   1455
         Left            =   930
         TabIndex        =   48
         Top             =   330
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   2566
         _Version        =   327682
         Orientation     =   1
         SmallChange     =   5
         Min             =   5
         Max             =   25
         SelStart        =   5
         Value           =   5
      End
      Begin VB.Shape Shape10 
         FillStyle       =   0  'Solid
         Height          =   585
         Left            =   180
         Shape           =   3  'Circle
         Top             =   1410
         Width           =   675
      End
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   210
         Shape           =   3  'Circle
         Top             =   780
         Width           =   525
      End
      Begin VB.Shape Shape8 
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   240
         Shape           =   3  'Circle
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   15
      Left            =   2700
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   44
      Top             =   240
      Width           =   15
      Visible         =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Background"
      Height          =   1815
      Left            =   510
      TabIndex        =   37
      Top             =   570
      Width           =   1695
      Visible         =   0   'False
      Begin VB.HScrollBar blueScroll 
         Height          =   135
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   40
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar greenScroll 
         Height          =   135
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   39
         Top             =   1080
         Width           =   855
      End
      Begin VB.HScrollBar redScroll 
         Height          =   135
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   38
         Top             =   600
         Width           =   855
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   1080
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3D pipe"
      Height          =   1215
      Index           =   0
      Left            =   480
      TabIndex        =   32
      Top             =   780
      Width           =   1695
      Visible         =   0   'False
      Begin VB.CommandButton Command4 
         Caption         =   "Color"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Fill color"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   750
         Width           =   1035
      End
      Begin VB.HScrollBar HS2 
         Height          =   255
         LargeChange     =   10
         Left            =   840
         Max             =   100
         Min             =   7
         TabIndex        =   33
         Top             =   480
         Value           =   7
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Penn"
      Height          =   1095
      Left            =   450
      TabIndex        =   28
      Top             =   960
      Width           =   1695
      Visible         =   0   'False
      Begin VB.CommandButton Command5 
         Caption         =   "Color"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
      Begin VB.HScrollBar HS1 
         Height          =   375
         Left            =   840
         Max             =   400
         Min             =   1
         TabIndex        =   29
         Top             =   480
         Value           =   1
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rectangle"
      Height          =   1425
      Left            =   540
      TabIndex        =   24
      Top             =   630
      Width           =   1695
      Visible         =   0   'False
      Begin VB.CommandButton Command7 
         Caption         =   "Color"
         Height          =   375
         Left            =   300
         TabIndex        =   26
         Top             =   930
         Width           =   1095
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   2
         Min             =   1
         Max             =   50
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Size"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get color"
      Height          =   1335
      Left            =   600
      TabIndex        =   22
      Top             =   840
      Width           =   1575
      Visible         =   0   'False
      Begin VB.CommandButton Command8 
         Caption         =   "Make it so"
         Height          =   495
         Left            =   150
         TabIndex        =   23
         Top             =   270
         Width           =   735
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   960
         Top             =   270
         Width           =   495
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   180
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sun feather"
      ClipControls    =   0   'False
      DragIcon        =   "Rita3ny).frx":0000
      DragMode        =   1  'Automatic
      Height          =   1455
      Left            =   570
      TabIndex        =   19
      Top             =   720
      Width           =   1515
      Visible         =   0   'False
      Begin VB.CommandButton Command1 
         Caption         =   "Color"
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   900
         Width           =   645
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   270
         Max             =   10
         Min             =   1
         TabIndex        =   20
         Top             =   540
         Value           =   1
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   255
         Left            =   270
         TabIndex        =   21
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Line"
      Height          =   1065
      Left            =   570
      TabIndex        =   16
      Top             =   1020
      Width           =   1605
      Visible         =   0   'False
      Begin VB.CommandButton Command10 
         Caption         =   "Line color"
         Height          =   285
         Left            =   180
         TabIndex        =   18
         Top             =   570
         Width           =   1245
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   255
         Left            =   210
         TabIndex        =   17
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   2
         Min             =   1
         Max             =   20
         SelStart        =   1
         Value           =   1
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stars"
      Height          =   1365
      Left            =   690
      TabIndex        =   12
      Top             =   870
      Width           =   1395
      Visible         =   0   'False
      Begin VB.CommandButton Command11 
         Caption         =   "Color"
         Height          =   375
         Left            =   180
         TabIndex        =   14
         Top             =   270
         Width           =   975
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   180
         Max             =   300
         Min             =   5
         TabIndex        =   13
         Top             =   930
         Value           =   5
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Size"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   690
         Width           =   975
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Air brush"
      Height          =   2775
      Left            =   270
      TabIndex        =   3
      Top             =   90
      Width           =   2175
      Visible         =   0   'False
      Begin VB.CommandButton Command12 
         Caption         =   "Fill color"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Color"
         Height          =   375
         Left            =   1260
         TabIndex        =   7
         Top             =   2280
         Width           =   615
      End
      Begin ComctlLib.Slider Slider4 
         Height          =   255
         Left            =   300
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   327682
         Min             =   1
         Max             =   70
         SelStart        =   1
         Value           =   1
      End
      Begin ComctlLib.Slider Slider3 
         Height          =   255
         Left            =   300
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   2
         Min             =   5
         Max             =   300
         SelStart        =   5
         Value           =   5
      End
      Begin ComctlLib.Slider S1 
         Height          =   255
         Left            =   300
         TabIndex        =   6
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   327682
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Size"
         Height          =   255
         Left            =   300
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Distribution"
         Height          =   255
         Left            =   300
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Amount "
         Height          =   255
         Left            =   300
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ovel"
      Height          =   1335
      Left            =   390
      TabIndex        =   0
      Top             =   810
      Width           =   1965
      Visible         =   0   'False
      Begin VB.CommandButton Command14 
         Caption         =   "Color"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   690
         Width           =   975
      End
      Begin ComctlLib.Slider s2 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   2
         Min             =   1
         Max             =   20
         SelStart        =   1
         Value           =   1
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      Height          =   2955
      Left            =   30
      TabIndex        =   45
      Top             =   0
      Width           =   2745
      Begin MSComDlg.CommonDialog CommonDialog3 
         Left            =   1140
         Top             =   210
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog4 
         Left            =   660
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   210
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Open..."
         Filter          =   "Image Files (*.bmp)|*.bmp|jpeg bilder(*.jpg)|*.jpg|gif Files (*.gif)|*.gif|"
      End
      Begin VB.Shape Shape7 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   3000
         Top             =   870
         Width           =   375
      End
      Begin VB.Shape Shape4 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   3000
         Top             =   1860
         Width           =   375
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   3000
         Top             =   2310
         Width           =   375
      End
      Begin VB.Shape Shape6 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2940
         Top             =   1380
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Färgen, fargen, Farg, farg2, linjen As Double
 Sub blueScroll_Change()
    Label5.Caption = blueScroll.Value
    Shape1.FillColor = RGB(redScroll.Value, greenScroll.Value, blueScroll.Value)

End Sub

Private Sub Form_Load()
Me.Move 1000, 3333
End Sub

Public Sub Frame11_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Public Sub greenScroll_Change()
    Label6.Caption = greenScroll.Value
    Shape1.FillColor = RGB(redScroll.Value, greenScroll.Value, blueScroll.Value)

End Sub

Public Sub redscroll_Change()
    Label4.Caption = redScroll.Value
    Shape1.FillColor = RGB(redScroll.Value, greenScroll.Value, blueScroll.Value)
End Sub

Public Sub pCol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    colpressed = True
    Shape8.FillColor = pCol.Point(X, Y)
  
End Sub

Public Sub HS1_Change()
    Label2.Caption = HS1.Value
End Sub

Public Sub HS2_Change()
    Label1.Caption = HS2.Value
End Sub

Public Sub jg_Click()

End Sub

Public Sub HScroll1_Change()
    Label14.Caption = HScroll1.Value
    Form3.Pic1.DrawWidth = HScroll1.Value
End Sub
Public Sub Command1_Click()
   CommonDialog1.ShowColor
End Sub

Private Sub Command10_Click()
On Error GoTo 5544
    CommonDialog1.ShowColor
   
5544
End Sub

Private Sub Command11_Click()
    CommonDialog3.ShowColor
    
End Sub

Private Sub Command12_Click()
 CommonDialog2.ShowColor
    
End Sub

Private Sub Command13_Click()
On Error GoTo 6655
    CommonDialog1.ShowColor
    
6655
End Sub

Private Sub Command14_Click()
On Error GoTo 5566
    CommonDialog1.ShowColor
   
5566
End Sub

Private Sub Command2_Click()
On Error GoTo 5665
    Pic1.Cls
    Pic1.Picture = LoadPicture
5665
End Sub

Private Sub Command3_Click()
    CommonDialog2.ShowColor
    
End Sub

Private Sub Command4_Click()
On Error GoTo 321
    CommonDialog1.ShowColor
    
321
End Sub

Private Sub Command5_Click()
    CommonDialog3.ShowColor
    
End Sub



Private Sub Command7_Click()
    CommonDialog3.ShowColor
    
End Sub


Public Sub Command8_Click()
    farg2 = Shape5.FillColor
End Sub

