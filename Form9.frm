VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
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
      Height          =   975
      Left            =   11000
      TabIndex        =   8
      Top             =   10000
      Width           =   2500
   End
   Begin VB.Frame Frame1 
      Caption         =   "设计流程"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8100
      Left            =   12000
      TabIndex        =   0
      Top             =   1000
      Width           =   10500
      Begin VB.Label Label7 
         Caption         =   "7、	确定毛边槽形式：h=1.6mm，h1=4mm，b=8mm，b1=25mm，r=1mm，投影面积126mm^2。"
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
         TabIndex        =   7
         Top             =   6200
         Width           =   9500
      End
      Begin VB.Label Label6 
         Caption         =   "6、	确定锻锤吨位：根据公式计算可知，G=αβF=720kg，选用1t锻锤。"
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
         TabIndex        =   6
         Top             =   5200
         Width           =   9500
      End
      Begin VB.Label Label5 
         Caption         =   $"Form9.frx":0000
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
         TabIndex        =   5
         Top             =   3800
         Width           =   9500
      End
      Begin VB.Label Label4 
         Caption         =   "4、	确定圆角半径：锻件高度余量为1.15mm，则需要倒角的圆角半径为3.15mm，取3mm，其余部位圆角半径取1.5mm。"
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
         TabIndex        =   4
         Top             =   3000
         Width           =   9500
      End
      Begin VB.Label Label3 
         Caption         =   "3、	确定模锻斜度：技术要求注明模锻斜度7°。"
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
         Top             =   2500
         Width           =   6735
      End
      Begin VB.Label Label2 
         Caption         =   $"Form9.frx":00E5
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
         TabIndex        =   2
         Top             =   1500
         Width           =   9500
      End
      Begin VB.Label Label1 
         Caption         =   "1、	确定分模位置：上下对称的直线分模"
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
         Top             =   1000
         Width           =   6000
      End
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   120
      Top             =   1200
      Width           =   8775
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\7.jpg")
Image1.Stretch = False
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
