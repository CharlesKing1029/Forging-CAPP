VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
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
      Left            =   11100
      TabIndex        =   0
      Top             =   9500
      Width           =   2500
   End
   Begin VB.Label Label6 
      Caption         =   "形式Ⅵ称为楔形毛边槽，其特点是终锻时水平方向金属流动愈来愈困难，适用形状更复杂的锻件，缺点是切除毛边困难。"
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
      Left            =   12000
      TabIndex        =   6
      Top             =   7200
      Width           =   8000
   End
   Begin VB.Label Label5 
      Caption         =   "形式Ⅴ只用于锻模局部，桥部增设阻尼沟，增加金属向仓部流动阻力，迫使金属流向型槽深处或枝芽处。"
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
      Left            =   12000
      TabIndex        =   5
      Top             =   6000
      Width           =   8000
   End
   Begin VB.Label Label4 
      Caption         =   "形式Ⅳ使用对象同形式Ⅲ，由于加宽下模毛边槽桥部，因而提高桥部强度，以避免桥部过快地磨损和过早地压塌。"
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
      Left            =   12000
      TabIndex        =   4
      Top             =   4800
      Width           =   8000
   End
   Begin VB.Label Label3 
      Caption         =   "形式Ⅲ适用于形状复杂，坯料体积不易计算准确而往往偏多的锻件，由于增大仓部容积，不至于发生上下模压不靠。"
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
      Left            =   12000
      TabIndex        =   3
      Top             =   3600
      Width           =   8000
   End
   Begin VB.Label Label2 
      Caption         =   $"Form5.frx":0000
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
      Left            =   12000
      TabIndex        =   2
      Top             =   2400
      Width           =   8000
   End
   Begin VB.Label Label1 
      Caption         =   "形式Ⅰ是使用最广泛的一种，其优点是桥部设在上模块，与坯料接触时间短，吸收热量少，因而温升少，能减轻桥部磨损或避免压塌。"
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
      Left            =   12000
      TabIndex        =   1
      Top             =   1200
      Width           =   8000
   End
   Begin VB.Image Image1 
      Height          =   2000
      Left            =   2000
      Top             =   2500
      Width           =   2500
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\4.jpg")
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
