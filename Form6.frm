VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   11100
      TabIndex        =   2
      Top             =   11000
      Width           =   2500
   End
   Begin VB.Frame Frame2 
      Caption         =   "锻件形状复杂系数"
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
      Left            =   500
      TabIndex        =   1
      Top             =   2000
      Width           =   12500
      Begin VB.TextBox Text3 
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
         Left            =   3600
         TabIndex        =   9
         Top             =   4200
         Width           =   2500
      End
      Begin VB.CommandButton Command1 
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
         Height          =   855
         Left            =   3720
         TabIndex        =   7
         Top             =   6120
         Width           =   2055
      End
      Begin VB.TextBox Text2 
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
         Left            =   3600
         TabIndex        =   6
         Top             =   2700
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
         Height          =   500
         Left            =   3600
         TabIndex        =   4
         Top             =   1200
         Width           =   2500
      End
      Begin VB.Image Image1 
         Height          =   1935
         Left            =   6400
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label Label3 
         Caption         =   "形状复杂系数"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   480
         TabIndex        =   8
         Top             =   4200
         Width           =   2505
      End
      Begin VB.Label Label2 
         Caption         =   "外廓包容体体积/mm^2"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   500
         TabIndex        =   5
         Top             =   2700
         Width           =   2500
      End
      Begin VB.Label Label1 
         Caption         =   "锻件体积/mm^2"
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
         Left            =   500
         TabIndex        =   3
         Top             =   1200
         Width           =   2500
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "锻件材质系数"
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
      Left            =   14000
      TabIndex        =   0
      Top             =   2000
      Width           =   8000
      Begin VB.Label Label7 
         Caption         =   "M3——不锈钢、高温耐热合金和钛合金"
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
         TabIndex        =   13
         Top             =   4800
         Width           =   6975
      End
      Begin VB.Label Label6 
         Caption         =   "M2——高碳高合金含量钢"
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
         TabIndex        =   12
         Top             =   3600
         Width           =   6735
      End
      Begin VB.Label Label5 
         Caption         =   "M1——低碳低合金钢"
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
         TabIndex        =   11
         Top             =   2400
         Width           =   6375
      End
      Begin VB.Label Label4 
         Caption         =   "M0——铝、镁合金"
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
         TabIndex        =   10
         Top             =   1200
         Width           =   3000
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\2.jpg")
Image1.Stretch = False
End Sub
Private Sub Command2_click()
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
Private Sub Text1_Change()
If IsNumeric(Text1.Text) = False Then
Text1.Text = ""
Text1.SetFocus
End If
End Sub
Private Sub Text2_Change()
If IsNumeric(Text2.Text) = False Then
Text2.Text = ""
Text2.SetFocus
End If
End Sub
Private Sub Command1_Click()
Dim a, b, c As Double
c = Val(Text1.Text) / Val(Text2.Text)
If c <= 0 Or c > 1 Then
Text3.Text = ""
MsgBox "无效参数!", , "警告"
ElseIf c > 0 And c <= 0.16 Then
Text3.Text = "S4,形状复杂"
ElseIf c > 0.16 And c <= 0.32 Then
Text3.Text = "S3,形状较复杂"
ElseIf c > 0.32 And c <= 0.63 Then
Text3.Text = "S2,形状一般"
ElseIf c > 0.63 And c <= 1 Then
Text3.Text = "S1,形状简单"
End If
End Sub
