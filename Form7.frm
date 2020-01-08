VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   ScaleHeight     =   12495
   ScaleWidth      =   22920
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   11000
      TabIndex        =   9
      Top             =   10000
      Width           =   2500
   End
   Begin VB.Frame Frame2 
      Caption         =   "圆角半径与余量"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8000
      Left            =   10000
      TabIndex        =   8
      Top             =   1500
      Width           =   12495
      Begin VB.Label Label6 
         Caption         =   $"Form7.frx":0000
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   5040
         TabIndex        =   12
         Top             =   2280
         Width           =   7000
      End
      Begin VB.Label Label5 
         Caption         =   $"Form7.frx":008A
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   5000
         TabIndex        =   11
         Top             =   600
         Width           =   7000
      End
      Begin VB.Image Image1 
         Height          =   2415
         Left            =   500
         Top             =   500
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "起模斜度"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8000
      Left            =   1440
      TabIndex        =   0
      Top             =   1500
      Width           =   8000
      Begin VB.CommandButton Command2 
         Caption         =   "确定"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   2800
         TabIndex        =   10
         Top             =   6000
         Width           =   2500
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2800
         TabIndex        =   6
         Top             =   3600
         Width           =   2500
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2800
         TabIndex        =   4
         Top             =   2400
         Width           =   2500
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2800
         TabIndex        =   2
         Top             =   1200
         Width           =   2500
      End
      Begin VB.Label Label4 
         Caption         =   "内起模斜度β可以按照外起模斜度α数值增大2°或3°"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   500
         TabIndex        =   7
         Top             =   4800
         Width           =   4500
      End
      Begin VB.Label Label3 
         Caption         =   "外起模斜度α"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   500
         TabIndex        =   5
         Top             =   3600
         Width           =   2000
      End
      Begin VB.Label Label2 
         Caption         =   "L/B"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   500
         TabIndex        =   3
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "H/B"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   500
         TabIndex        =   1
         Top             =   1200
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\6.jpg")
Image1.Stretch = False
Combo1.AddItem ("<=1")
Combo1.AddItem ("1-3")
Combo1.AddItem ("3-4.5")
Combo1.AddItem ("4.5-6.5")
Combo1.AddItem ("6.5-8")
Combo1.AddItem (">8")
Combo2.AddItem ("<=1.5")
Combo2.AddItem (">1.5")
End Sub
Private Sub Command2_click()
If Combo1.Text = "<=1" And Combo2.Text = "<=1.5" Then
Text1.Text = "5°"
ElseIf Combo1.Text = "<=1" And Combo2.Text = ">1.5" Then
Text1.Text = "5°"
ElseIf Combo1.Text = "1-3" And Combo2.Text = "<=1.5" Then
Text1.Text = "7°"
ElseIf Combo1.Text = "1-3" And Combo2.Text = ">1.5" Then
Text1.Text = "5°"
ElseIf Combo1.Text = "3-4.5" And Combo2.Text = "<=1.5" Then
Text1.Text = "10°"
ElseIf Combo1.Text = "3-4.5" And Combo2.Text = ">1.5" Then
Text1.Text = "7°"
ElseIf Combo1.Text = "4.5-6.5" And Combo2.Text = "<=1.5" Then
Text1.Text = "12°"
ElseIf Combo1.Text = "4.5-6.5" And Combo2.Text = ">1.5" Then
Text1.Text = "10°"
ElseIf Combo1.Text = "6.5-8" And Combo2.Text = "<=1.5" Then
Text1.Text = "15°"
ElseIf Combo1.Text = "6.5-8" And Combo2.Text = ">1.5" Then
Text1.Text = "12°"
ElseIf Combo1.Text = ">8" And Combo2.Text = "<=1.5" Then
Text1.Text = "15°"
ElseIf Combo1.Text = ">8" And Combo2.Text = ">1.5" Then
Text1.Text = "15°"
End If
End Sub
Private Sub Command1_Click()
form1.Show
form2.Hide
form3.Hide
form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide
End Sub
