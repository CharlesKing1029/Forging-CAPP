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
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "��Բ"
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
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "��Բ"
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
         Caption         =   "7��	ȷ��ë�߲���ʽ��h=1.6mm��h1=4mm��b=8mm��b1=25mm��r=1mm��ͶӰ���126mm^2��"
         BeginProperty Font 
            Name            =   "��Բ"
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
         Caption         =   "6��	ȷ���ʹ���λ�����ݹ�ʽ�����֪��G=����F=720kg��ѡ��1t�ʹ���"
         BeginProperty Font 
            Name            =   "��Բ"
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
            Name            =   "��Բ"
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
         Caption         =   "4��	ȷ��Բ�ǰ뾶���ͼ��߶�����Ϊ1.15mm������Ҫ���ǵ�Բ�ǰ뾶Ϊ3.15mm��ȡ3mm�����ಿλԲ�ǰ뾶ȡ1.5mm��"
         BeginProperty Font 
            Name            =   "��Բ"
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
         Caption         =   "3��	ȷ��ģ��б�ȣ�����Ҫ��ע��ģ��б��7�㡣"
         BeginProperty Font 
            Name            =   "��Բ"
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
            Name            =   "��Բ"
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
         Caption         =   "1��	ȷ����ģλ�ã����¶ԳƵ�ֱ�߷�ģ"
         BeginProperty Font 
            Name            =   "��Բ"
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
