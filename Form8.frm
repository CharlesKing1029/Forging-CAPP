VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
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
      TabIndex        =   8
      Top             =   10000
      Width           =   2500
   End
   Begin VB.Frame Frame1 
      Caption         =   "技术条件主要内容"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8200
      Left            =   8000
      TabIndex        =   0
      Top             =   1500
      Width           =   9000
      Begin VB.Label Label7 
         Caption         =   "（7）其他特殊要求，如锻件同心度、弯曲度。"
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
         TabIndex        =   7
         Top             =   7320
         Width           =   7000
      End
      Begin VB.Label Label6 
         Caption         =   "（6）需要取样进行金相组织检验和力学性能试验时，应注明在锻件上的取样位置。"
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
         Top             =   6000
         Width           =   7000
      End
      Begin VB.Label Label5 
         Caption         =   "（5）表面清理方法。"
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
         TabIndex        =   5
         Top             =   5000
         Width           =   5655
      End
      Begin VB.Label Label4 
         Caption         =   "（4）锻后热处理方法及硬度要求。"
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
         TabIndex        =   4
         Top             =   4000
         Width           =   6015
      End
      Begin VB.Label Label3 
         Caption         =   "（3）允许的表面缺陷深度。"
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
         Top             =   3000
         Width           =   7095
      End
      Begin VB.Label Label2 
         Caption         =   "（2）允许错移量和残余飞边的宽度。"
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
         TabIndex        =   2
         Top             =   2000
         Width           =   6615
      End
      Begin VB.Label Label1 
         Caption         =   "（1）未注明的模锻斜度和圆角半径。"
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
         Width           =   7995
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
