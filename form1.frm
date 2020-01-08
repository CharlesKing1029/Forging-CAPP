VERSION 5.00
Begin VB.Form form1 
   Caption         =   "锤上模锻件设计"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "应用举例"
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
      Left            =   18000
      TabIndex        =   6
      Top             =   10000
      Width           =   3800
   End
   Begin VB.CommandButton Command6 
      Caption         =   "常见技术要求"
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
      Left            =   18000
      TabIndex        =   5
      Top             =   7000
      Width           =   3800
   End
   Begin VB.CommandButton Command5 
      Caption         =   "确定锻件形状复杂系数与锻件材质系数"
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
      Left            =   18000
      TabIndex        =   4
      Top             =   1000
      Width           =   3800
   End
   Begin VB.CommandButton Command4 
      Caption         =   "确定锻件的脱模斜度与圆角半径"
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
      Left            =   18000
      TabIndex        =   3
      Top             =   5500
      Width           =   3800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "确定锻件内孔直径的机械加工单边余量"
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
      Left            =   18000
      TabIndex        =   2
      Top             =   4000
      Width           =   3800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "根据锻件长度、宽度、高度确定机械加工单边余量与公差"
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
      Left            =   18000
      TabIndex        =   1
      Top             =   2500
      Width           =   3800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定模锻锤吨位与毛边槽尺寸"
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
      Left            =   18000
      TabIndex        =   0
      Top             =   8500
      Width           =   3800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Top             =   900
      Width           =   1815
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\1.jpg")
Image1.Stretch = False
End Sub
Private Sub Command1_Click()
form1.Hide
form2.Show
form3.Hide
form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide
End Sub
Private Sub Command2_click()
form1.Hide
form2.Hide
form3.Show
form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide
End Sub
Private Sub command3_click()
form1.Hide
form2.Hide
form3.Hide
form4.Show
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide
End Sub
Private Sub Command5_Click()
form1.Hide
form2.Hide
form3.Hide
form4.Hide
Form5.Hide
Form6.Show
Form7.Hide
Form8.Hide
Form9.Hide
End Sub
Private Sub command6_click()
form1.Hide
form2.Hide
form3.Hide
form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Show
Form9.Hide
End Sub
Private Sub command4_click()
form1.Hide
form2.Hide
form3.Hide
form4.Hide
Form5.Hide
Form6.Hide
Form7.Show
Form8.Hide
Form9.Hide
End Sub
Private Sub command7_click()
form1.Hide
form2.Hide
form3.Hide
form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Show
End Sub
