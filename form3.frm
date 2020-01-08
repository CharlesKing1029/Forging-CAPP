VERSION 5.00
Begin VB.Form form3 
   Caption         =   "根据锻件长、宽、高确定机械加工余量与公差"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "根据锻件长、宽、高确定机械加工余量"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   1440
      TabIndex        =   2
      Top             =   1400
      Width           =   11000
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8500
         TabIndex        =   18
         Top             =   6000
         Width           =   2000
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8500
         TabIndex        =   16
         Top             =   4500
         Width           =   2000
      End
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
         Height          =   495
         Left            =   8500
         TabIndex        =   14
         Top             =   3000
         Width           =   2000
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
         Height          =   495
         Left            =   8500
         TabIndex        =   13
         Top             =   1500
         Width           =   2000
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
         Left            =   2500
         TabIndex        =   6
         Top             =   6000
         Width           =   2000
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   2500
         TabIndex        =   5
         Top             =   4500
         Width           =   2000
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
         Left            =   2500
         TabIndex        =   4
         Top             =   3000
         Width           =   2000
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
         Left            =   2500
         TabIndex        =   3
         Top             =   1500
         Width           =   2000
      End
      Begin VB.Label Label8 
         Caption         =   "水平尺寸下公差/mm"
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
         Left            =   5500
         TabIndex        =   17
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "水平尺寸上公差/mm"
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
         Left            =   5500
         TabIndex        =   15
         Top             =   4500
         Width           =   2500
      End
      Begin VB.Label Label6 
         Caption         =   "高度下公差/mm"
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
         Left            =   5500
         TabIndex        =   12
         Top             =   3000
         Width           =   2500
      End
      Begin VB.Label Label5 
         Caption         =   "高度上公差/mm"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5500
         TabIndex        =   11
         Top             =   1500
         Width           =   2500
      End
      Begin VB.Label Label4 
         Caption         =   "单边机械加工余量/mm"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   500
         TabIndex        =   10
         Top             =   6000
         Width           =   1600
      End
      Begin VB.Label Label3 
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
         Height          =   375
         Left            =   500
         TabIndex        =   9
         Top             =   4500
         Width           =   1600
      End
      Begin VB.Label Label2 
         Caption         =   "长度L/mm"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   500
         TabIndex        =   8
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "高度H/mm"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   495
         TabIndex        =   7
         Top             =   1500
         Width           =   1920
      End
   End
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
      Height          =   1000
      Left            =   8160
      TabIndex        =   1
      Top             =   9000
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
      Height          =   1000
      Left            =   3120
      TabIndex        =   0
      Top             =   9000
      Width           =   2500
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   12500
      Top             =   1600
      Width           =   10000
   End
End
Attribute VB_Name = "form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\5.jpg")
Image1.Stretch = False
Combo1.AddItem ("<30")
Combo1.AddItem ("30-60")
Combo1.AddItem ("60-100")
Combo1.AddItem ("100-150")
Combo1.AddItem (">150")     'combo1为高度
Combo2.AddItem ("<=50")
Combo2.AddItem ("50-120")
Combo2.AddItem ("120-260")
Combo2.AddItem ("260-360")
Combo2.AddItem ("360-500")
Combo2.AddItem ("500-800")
Combo2.AddItem ("800-1250")
Combo2.AddItem (">1250")      'combo2为长度
Combo3.AddItem ("<2")
Combo3.AddItem ("2-5")
Combo3.AddItem (">5")       'combo3为长宽比
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
Form9.Hide     '返回主界面
End Sub
Private Sub Command1_Click()
If Combo1.Text = "<30" And Combo2.Text = "<=50" And Combo3.Text = "<2" Then
Text1.Text = "1.0"
Text2.Text = "+0.8"
Text3.Text = "-0.4"
Text4.Text = "+1.0"
Text5.Text = "-0.5"
ElseIf Combo1.Text = "<30" And Combo2.Text = "50-120" And Combo3.Text = "<2" Then
Text1.Text = "1.25"
Text2.Text = "+0.9"
Text3.Text = "-0.5"
Text4.Text = "+1.5"
Text5.Text = "-0.7"
ElseIf Combo1.Text = "<30" And Combo2.Text = "120-260" And Combo3.Text = "<2" Then
Text1.Text = "1.5"
Text2.Text = "+1.0"
Text3.Text = "-0.5"
Text4.Text = "+2.0"
Text5.Text = "-1.0"
ElseIf Combo1.Text = "<30" And Combo2.Text = "260-360" And Combo3.Text = "<2" Then
Text1.Text = "1.75"
Text2.Text = "+1.2"
Text3.Text = "-0.6"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "<30" And Combo2.Text = "360-500" And Combo3.Text = "<2" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "<30" And Combo2.Text = "500-800" And Combo3.Text = "<2" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "<30" And Combo2.Text = "800-1250" And Combo3.Text = "<2" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "<30" And Combo2.Text = ">1250" And Combo3.Text = "<2" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
'<30,L/B <2



ElseIf Combo1.Text = "<30" And Combo2.Text = "<=50" And Combo3.Text = "2-5" Then
Text1.Text = "1.0"
Text2.Text = "+0.8"
Text3.Text = "-0.4"
Text4.Text = "+1.0"
Text5.Text = "-0.5"
ElseIf Combo1.Text = "<30" And Combo2.Text = "50-120" And Combo3.Text = "2-5" Then
Text1.Text = "1.0"
Text2.Text = "+0.8"
Text3.Text = "-0.4"
Text4.Text = "+1.5"
Text5.Text = "-0.7"
ElseIf Combo1.Text = "<30" And Combo2.Text = "120-260" And Combo3.Text = "2-5" Then
Text1.Text = "1.25"
Text2.Text = "+0.9"
Text3.Text = "-0.5"
Text4.Text = "+2.0"
Text5.Text = "-1.0"
ElseIf Combo1.Text = "<30" And Combo2.Text = "260-360" And Combo3.Text = "2-5" Then
Text1.Text = "1.5"
Text2.Text = "+1.0"
Text3.Text = "-0.5"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "<30" And Combo2.Text = "360-500" And Combo3.Text = "2-5" Then
Text1.Text = "1.75"
Text2.Text = "+1.2"
Text3.Text = "-0.6"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "<30" And Combo2.Text = "500-800" And Combo3.Text = "2-5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "<30" And Combo2.Text = "800-1250" And Combo3.Text = "2-5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "<30" And Combo2.Text = ">1250" And Combo3.Text = "2-5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
'<30,L/B 2-5



ElseIf Combo1.Text = "<30" And Combo2.Text = "<=50" And Combo3.Text = ">5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "<30" And Combo2.Text = "50-120" And Combo3.Text = ">5" Then
Text1.Text = "1.0"
Text2.Text = "+0.8"
Text3.Text = "-0.4"
Text4.Text = "+1.0"
Text5.Text = "-0.5"
ElseIf Combo1.Text = "<30" And Combo2.Text = "120-260" And Combo3.Text = ">5" Then
Text1.Text = "1.0"
Text2.Text = "+0.8"
Text3.Text = "-0.4"
Text4.Text = "+2.0"
Text5.Text = "-1.0"
ElseIf Combo1.Text = "<30" And Combo2.Text = "260-360" And Combo3.Text = ">5" Then
Text1.Text = "1.25"
Text2.Text = "+0.9"
Text3.Text = "-0.5"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "<30" And Combo2.Text = "360-500" And Combo3.Text = ">5" Then
Text1.Text = "1.5"
Text2.Text = "+2.5"
Text3.Text = "-1.5"
ElseIf Combo1.Text = "<30" And Combo2.Text = "500-800" And Combo3.Text = ">5" Then
Text1.Text = "1.75"
Text2.Text = "+3.0"
Text3.Text = "-2.0"
ElseIf Combo1.Text = "<30" And Combo2.Text = "800-1250" And Combo3.Text = ">5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "<30" And Combo2.Text = ">1250" And Combo3.Text = ">5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
'<30,L/B >5



ElseIf Combo1.Text = "30-60" And Combo2.Text = "<=50" And Combo3.Text = "<2" Then
Text1.Text = "1.25"
Text2.Text = "+0.9"
Text3.Text = "-0.5"
Text4.Text = "+1.0"
Text5.Text = "-0.5"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "50-120" And Combo3.Text = "<2" Then
Text1.Text = "1.5"
Text2.Text = "+1.0"
Text3.Text = "-0.5"
Text4.Text = "+1.5"
Text5.Text = "-0.7"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "120-260" And Combo3.Text = "<2" Then
Text1.Text = "1.75"
Text2.Text = "+1.2"
Text3.Text = "-0.6"
Text4.Text = "+2.0"
Text5.Text = "-1.0"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "260-360" And Combo3.Text = "<2" Then
Text1.Text = "2.0"
Text2.Text = "+1.4"
Text3.Text = "-0.7"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "360-500" And Combo3.Text = "<2" Then
Text1.Text = "2.25"
Text2.Text = "+1.6"
Text3.Text = "-0.8"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "500-800" And Combo3.Text = "<2" Then
Text1.Text = "2.5"
Text2.Text = "+1.8"
Text3.Text = "-0.9"
Text4.Text = "+3.0"
Text5.Text = "-2.0"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "800-1250" And Combo3.Text = "<2" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "30-60" And Combo2.Text = ">1250" And Combo3.Text = "<2" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
'30-60,L/B <2



ElseIf Combo1.Text = "30-60" And Combo2.Text = "<=50" And Combo3.Text = "2-5" Then
Text1.Text = "1.0"
Text2.Text = "+0.8"
Text3.Text = "-0.4"
Text4.Text = "+1.0"
Text5.Text = "-0.5"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "50-120" And Combo3.Text = "2-5" Then
Text1.Text = "1.25"
Text2.Text = "+0.9"
Text3.Text = "-0.5"
Text4.Text = "+1.5"
Text5.Text = "-0.7"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "120-260" And Combo3.Text = "2-5" Then
Text1.Text = "1.5"
Text2.Text = "+1.0"
Text3.Text = "-0.5"
Text4.Text = "+1.0"
Text5.Text = "-0.5"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "260-360" And Combo3.Text = "2-5" Then
Text1.Text = "1.75"
Text2.Text = "+1.2"
Text3.Text = "-0.6"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "360-500" And Combo3.Text = "2-5" Then
Text1.Text = "2.0"
Text2.Text = "+1.4"
Text3.Text = "-0.7"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "500-800" And Combo3.Text = "2-5" Then
Text1.Text = "2.25"
Text2.Text = "+1.6"
Text3.Text = "-0.8"
Text4.Text = "+3.0"
Text5.Text = "-2.0"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "800-1250" And Combo3.Text = "2-5" Then
Text1.Text = "2.5"
Text2.Text = "+1.8"
Text3.Text = "-0.9"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
ElseIf Combo1.Text = "30-60" And Combo2.Text = ">1250" And Combo3.Text = "2-5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
'30-60,L/B 2-5



ElseIf Combo1.Text = "30-60" And Combo2.Text = "<=50" And Combo3.Text = ">5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "30-60" And Combo2.Text = "50-120" And Combo3.Text = ">5" Then
Text1.Text = "1.0"
Text2.Text = "+0.8"
Text3.Text = "-0.4"
Text4.Text = "+1.0"
Text5.Text = "-0.5"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "120-260" And Combo3.Text = ">5" Then
Text1.Text = "1.25"
Text2.Text = "+0.9"
Text3.Text = "-0.5"
Text4.Text = "+2.0"
Text5.Text = "-1.0"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "260-360" And Combo3.Text = ">5" Then
Text1.Text = "1.5"
Text2.Text = "+1.0"
Text3.Text = "-0.5"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "360-500" And Combo3.Text = ">5" Then
Text1.Text = "1.75"
Text2.Text = "+1.2"
Text3.Text = "-0.6"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "500-800" And Combo3.Text = ">5" Then
Text1.Text = "2.0"
Text2.Text = "+1.4"
Text3.Text = "-0.7"
Text4.Text = "+3.0"
Text5.Text = "-2.0"
ElseIf Combo1.Text = "30-60" And Combo2.Text = "800-1250" And Combo3.Text = ">5" Then
Text1.Text = "2.25"
Text2.Text = "+0.8"
Text3.Text = "-0.4"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
ElseIf Combo1.Text = "30-60" And Combo2.Text = ">1250" And Combo3.Text = ">5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
'30-60,L/B >5



ElseIf Combo1.Text = "60-100" And Combo2.Text = "<=50" And Combo3.Text = "<2" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "60-100" And Combo2.Text = "50-120" And Combo3.Text = "<2" Then
Text1.Text = "1.75"
Text2.Text = "+1.2"
Text3.Text = "-0.6"
Text4.Text = "+1.5"
Text5.Text = "-0.7"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "120-260" And Combo3.Text = "<2" Then
Text1.Text = "2.0"
Text2.Text = "+1.4"
Text3.Text = "-0.7"
Text4.Text = "+2.0"
Text5.Text = "-1.0"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "260-360" And Combo3.Text = "<2" Then
Text1.Text = "2.25"
Text2.Text = "+1.6"
Text3.Text = "-0.8"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "360-500" And Combo3.Text = "<2" Then
Text1.Text = "2.5"
Text2.Text = "+1.8"
Text3.Text = "-0.9"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "500-800" And Combo3.Text = "<2" Then
Text1.Text = "2.75"
Text2.Text = "+2.0"
Text3.Text = "-1.0"
Text4.Text = "+3.0"
Text5.Text = "-2.0"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "800-1250" And Combo3.Text = "<2" Then
Text1.Text = "3.0"
Text2.Text = "+2.2"
Text3.Text = "-1.1"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
ElseIf Combo1.Text = "60-100" And Combo2.Text = ">1250" And Combo3.Text = "<2" Then
Text1.Text = "3.25"
Text2.Text = "+2.4"
Text3.Text = "-1.2"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
'60-100,L/B <2



ElseIf Combo1.Text = "60-100" And Combo2.Text = "<=50" And Combo3.Text = "2-5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "60-100" And Combo2.Text = "50-120" And Combo3.Text = "2-5" Then
Text1.Text = "1.5"
Text2.Text = "+1.1"
Text3.Text = "-0.5"
Text4.Text = "+1.5"
Text5.Text = "-0.7"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "120-260" And Combo3.Text = "2-5" Then
Text1.Text = "1.75"
Text2.Text = "+1.2"
Text3.Text = "-0.6"
Text4.Text = "+2.0"
Text5.Text = "-1.0"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "260-360" And Combo3.Text = "2-5" Then
Text1.Text = "2.0"
Text2.Text = "+1.4"
Text3.Text = "-0.7"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "360-500" And Combo3.Text = "2-5" Then
Text1.Text = "2.25"
Text2.Text = "+1.6"
Text3.Text = "-0.8"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "500-800" And Combo3.Text = "2-5" Then
Text1.Text = "2.5"
Text2.Text = "+1.8"
Text3.Text = "-0.9"
Text4.Text = "+3.0"
Text5.Text = "-2.0"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "800-1250" And Combo3.Text = "2-5" Then
Text1.Text = "2.75"
Text2.Text = "+2.0"
Text3.Text = "-1.0"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
ElseIf Combo1.Text = "60-100" And Combo2.Text = ">1250" And Combo3.Text = "2-5" Then
Text1.Text = "3.0"
Text2.Text = "+2.2"
Text3.Text = "-1.1"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
'60-100 ,L/B 2-5



ElseIf Combo1.Text = "60-100" And Combo2.Text = "<=50" And Combo3.Text = ">5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "60-100" And Combo2.Text = "50-120" And Combo3.Text = ">5" Then
Text1.Text = "1.25"
Text2.Text = "+0.9"
Text3.Text = "-0.5"
Text4.Text = "+1.5"
Text5.Text = "-0.7"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "120-260" And Combo3.Text = ">5" Then
Text1.Text = "1.5"
Text2.Text = "+1.0"
Text3.Text = "-0.5"
Text4.Text = "+2.0"
Text5.Text = "-1.0"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "260-360" And Combo3.Text = ">5" Then
Text1.Text = "1.75"
Text2.Text = "+1.2"
Text3.Text = "-0.6"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "360-500" And Combo3.Text = ">5" Then
Text1.Text = "2.0"
Text2.Text = "+1.4"
Text3.Text = "-0.7"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "500-800" And Combo3.Text = ">5" Then
Text1.Text = "2.25"
Text2.Text = "+1.6"
Text3.Text = "-0.8"
Text4.Text = "+3.0"
Text5.Text = "-2.0"
ElseIf Combo1.Text = "60-100" And Combo2.Text = "800-1250" And Combo3.Text = ">5" Then
Text1.Text = "2.5"
Text2.Text = "+1.8"
Text3.Text = "-0.9"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
ElseIf Combo1.Text = "60-100" And Combo2.Text = ">1250" And Combo3.Text = ">5" Then
Text1.Text = "2.75"
Text2.Text = "+2.0"
Text3.Text = "-1.0"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
'60-100,L/B >5



ElseIf Combo1.Text = "100-150" And Combo2.Text = "<=50" And Combo3.Text = "<2" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "100-150" And Combo2.Text = "50-120" And Combo3.Text = "<2" Then
Text1.Text = "2.0"
Text2.Text = "+1.4"
Text3.Text = "-0.7"
Text4.Text = "+1.5"
Text5.Text = "-0.7"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "120-260" And Combo3.Text = "<2" Then
Text1.Text = "2.25"
Text2.Text = "+1.6"
Text3.Text = "-0.8"
Text4.Text = "+2.0"
Text5.Text = "-1.0"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "260-360" And Combo3.Text = "<2" Then
Text1.Text = "2.5"
Text2.Text = "+1.8"
Text3.Text = "-0.9"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "360-500" And Combo3.Text = "<2" Then
Text1.Text = "2.75"
Text2.Text = "+2.0"
Text3.Text = "-1.0"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "500-800" And Combo3.Text = "<2" Then
Text1.Text = "3.0"
Text2.Text = "+2.2"
Text3.Text = "-1.1"
Text4.Text = "+3.0"
Text5.Text = "-2.0"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "800-1250" And Combo3.Text = "<2" Then
Text1.Text = "3.25"
Text2.Text = "+2.4"
Text3.Text = "-1.2"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
ElseIf Combo1.Text = "100-150" And Combo2.Text = ">1250" And Combo3.Text = "<2" Then
Text1.Text = "3.5"
Text2.Text = "+2.6"
Text3.Text = "-1.3"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
'100-150,L/B <2



ElseIf Combo1.Text = "100-150" And Combo2.Text = "<=50" And Combo3.Text = "2-5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "100-150" And Combo2.Text = "50-120" And Combo3.Text = "2-5" Then
Text1.Text = "1.75"
Text2.Text = "+1.2"
Text3.Text = "-0.6"
Text4.Text = "+1.5"
Text5.Text = "-0.7"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "120-260" And Combo3.Text = "2-5" Then
Text1.Text = "2.0"
Text2.Text = "+1.4"
Text3.Text = "-0.7"
Text4.Text = "+2.0"
Text5.Text = "-1.0"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "260-360" And Combo3.Text = "2-5" Then
Text1.Text = "2.25"
Text2.Text = "+1.6"
Text3.Text = "-0.8"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "360-500" And Combo3.Text = "2-5" Then
Text1.Text = "2.5"
Text2.Text = "+1.8"
Text3.Text = "-0.9"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "500-800" And Combo3.Text = "2-5" Then
Text1.Text = "2.75"
Text2.Text = "+2.0"
Text3.Text = "-1.0"
Text4.Text = "+3.0"
Text5.Text = "-2.0"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "800-1250" And Combo3.Text = "2-5" Then
Text1.Text = "3.0"
Text2.Text = "+2.2"
Text3.Text = "-1.1"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
ElseIf Combo1.Text = "100-150" And Combo2.Text = ">1250" And Combo3.Text = "2-5" Then
Text1.Text = "3.25"
Text2.Text = "+2.4"
Text3.Text = "-1.2"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
'100-150 ,L/B 2-5



ElseIf Combo1.Text = "100-150" And Combo2.Text = "<=50" And Combo3.Text = ">5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = "100-150" And Combo2.Text = "50-120" And Combo3.Text = ">5" Then
Text1.Text = "1.5"
Text2.Text = "+1.0"
Text3.Text = "-0.5"
Text4.Text = "+1.5"
Text5.Text = "-0.7"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "120-260" And Combo3.Text = ">5" Then
Text1.Text = "1.75"
Text2.Text = "+1.2"
Text3.Text = "-0.6"
Text4.Text = "+2.0"
Text5.Text = "-1.0"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "260-360" And Combo3.Text = ">5" Then
Text1.Text = "2.0"
Text2.Text = "+1.4"
Text3.Text = "-0.7"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "360-500" And Combo3.Text = ">5" Then
Text1.Text = "2.25"
Text2.Text = "+1.6"
Text3.Text = "-0.8"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "500-800" And Combo3.Text = ">5" Then
Text1.Text = "2.5"
Text2.Text = "+1.8"
Text3.Text = "-0.9"
Text4.Text = "+3.0"
Text5.Text = "-2.0"
ElseIf Combo1.Text = "100-150" And Combo2.Text = "800-1250" And Combo3.Text = ">5" Then
Text1.Text = "2.75"
Text2.Text = "+2.0"
Text3.Text = "-1.0"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
ElseIf Combo1.Text = "100-150" And Combo2.Text = ">1250" And Combo3.Text = ">5" Then
Text1.Text = "3.0"
Text2.Text = "+2.2"
Text3.Text = "-1.1"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
'100-150,L/B >5



ElseIf Combo1.Text = ">150" And Combo2.Text = "<=50" And Combo3.Text = "<2" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = ">150" And Combo2.Text = "50-120" And Combo3.Text = "<2" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = ">150" And Combo2.Text = "120-260" And Combo3.Text = "<2" Then
Text1.Text = "2.5"
Text2.Text = "+1.8"
Text3.Text = "-0.9"
Text4.Text = "+2.0"
Text5.Text = "-1.0"
ElseIf Combo1.Text = ">150" And Combo2.Text = "260-360" And Combo3.Text = "<2" Then
Text1.Text = "2.75"
Text2.Text = "+2.0"
Text3.Text = "-1.0"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = ">150" And Combo2.Text = "360-500" And Combo3.Text = "<2" Then
Text1.Text = "3.0"
Text2.Text = "+2.2"
Text3.Text = "-1.1"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = ">150" And Combo2.Text = "500-800" And Combo3.Text = "<2" Then
Text1.Text = "3.25"
Text2.Text = "+2.4"
Text3.Text = "-1.2"
Text4.Text = "+3.0"
Text5.Text = "-2.0"
ElseIf Combo1.Text = ">150" And Combo2.Text = "800-1250" And Combo3.Text = "<2" Then
Text1.Text = "3.5"
Text2.Text = "+2.6"
Text3.Text = "-1.3"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
ElseIf Combo1.Text = ">150" And Combo2.Text = ">1250" And Combo3.Text = "<2" Then
Text1.Text = "3.75"
Text2.Text = "+2.8"
Text3.Text = "-1.4"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
'>150 ,L/B <2



ElseIf Combo1.Text = ">150" And Combo2.Text = "<=50" And Combo3.Text = "2-5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = ">150" And Combo2.Text = "50-120" And Combo3.Text = "2-5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = ">150" And Combo2.Text = "120-260" And Combo3.Text = "2-5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = ">150" And Combo2.Text = "260-360" And Combo3.Text = "2-5" Then
Text1.Text = "2.5"
Text2.Text = "+1.8"
Text3.Text = "-0.9"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = ">150" And Combo2.Text = "360-500" And Combo3.Text = "2-5" Then
Text1.Text = "2.75"
Text2.Text = "+2.0"
Text3.Text = "-1.0"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = ">150" And Combo2.Text = "500-800" And Combo3.Text = "2-5" Then
Text1.Text = "3.0"
Text2.Text = "+2.2"
Text3.Text = "-1.1"
Text4.Text = "+3.0"
Text5.Text = "-2.0"
ElseIf Combo1.Text = ">150" And Combo2.Text = "800-1250" And Combo3.Text = "2-5" Then
Text1.Text = "3.25"
Text2.Text = "+2.4"
Text3.Text = "-1.2"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
ElseIf Combo1.Text = ">150" And Combo2.Text = ">1250" And Combo3.Text = "2-5" Then
Text1.Text = "3.5"
Text2.Text = "+2.6"
Text3.Text = "-1.3"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
'>150 ,L/B 2-5



ElseIf Combo1.Text = ">150" And Combo2.Text = "<=50" And Combo3.Text = ">5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = ">150" And Combo2.Text = "50-120" And Combo3.Text = ">5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = ">150" And Combo2.Text = "120-260" And Combo3.Text = ">5" Then
MsgBox "无效参数！", , "警告"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
ElseIf Combo1.Text = ">150" And Combo2.Text = "260-360" And Combo3.Text = ">5" Then
Text1.Text = "2.25"
Text2.Text = "+1.6"
Text3.Text = "-0.8"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = ">150" And Combo2.Text = "360-500" And Combo3.Text = ">5" Then
Text1.Text = "2.5"
Text2.Text = "+1.8"
Text3.Text = "-0.9"
Text4.Text = "+2.5"
Text5.Text = "-1.5"
ElseIf Combo1.Text = ">150" And Combo2.Text = "500-800" And Combo3.Text = ">5" Then
Text1.Text = "2.75"
Text2.Text = "+2.0"
Text3.Text = "-1.0"
Text4.Text = "+3.0"
Text5.Text = "-2.0"
ElseIf Combo1.Text = ">150" And Combo2.Text = "800-1250" And Combo3.Text = ">5" Then
Text1.Text = "3.0"
Text2.Text = "+2.2"
Text3.Text = "-1.1"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
ElseIf Combo1.Text = ">150" And Combo2.Text = ">1250" And Combo3.Text = ">5" Then
Text1.Text = "3.25"
Text2.Text = "+2.4"
Text3.Text = "-1.2"
Text4.Text = "+3.5"
Text5.Text = "-2.5"
'>150 ,L/B >5



End If
End Sub
