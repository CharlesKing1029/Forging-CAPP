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
      Height          =   1000
      Left            =   11100
      TabIndex        =   0
      Top             =   9500
      Width           =   2500
   End
   Begin VB.Label Label6 
      Caption         =   "��ʽ����ΪШ��ë�߲ۣ����ص����ն�ʱˮƽ��������������������ѣ�������״�����ӵĶͼ���ȱ�����г�ë�����ѡ�"
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
      Left            =   12000
      TabIndex        =   6
      Top             =   7200
      Width           =   8000
   End
   Begin VB.Label Label5 
      Caption         =   "��ʽ��ֻ���ڶ�ģ�ֲ����Ų��������ṵ�����ӽ�����ֲ�������������ʹ���������Ͳ����֦ѿ����"
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
      Left            =   12000
      TabIndex        =   5
      Top             =   6000
      Width           =   8000
   End
   Begin VB.Label Label4 
      Caption         =   "��ʽ��ʹ�ö���ͬ��ʽ�����ڼӿ���ģë�߲��Ų����������Ų�ǿ�ȣ��Ա����Ų������ĥ��͹����ѹ����"
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
      Left            =   12000
      TabIndex        =   4
      Top             =   4800
      Width           =   8000
   End
   Begin VB.Label Label3 
      Caption         =   "��ʽ����������״���ӣ�����������׼���׼ȷ������ƫ��Ķͼ�����������ֲ��ݻ��������ڷ�������ģѹ������"
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
      Left            =   12000
      TabIndex        =   3
      Top             =   3600
      Width           =   8000
   End
   Begin VB.Label Label2 
      Caption         =   $"Form5.frx":0000
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
      Left            =   12000
      TabIndex        =   2
      Top             =   2400
      Width           =   8000
   End
   Begin VB.Label Label1 
      Caption         =   "��ʽ����ʹ����㷺��һ�֣����ŵ����Ų�������ģ�飬�����ϽӴ�ʱ��̣����������٣���������٣��ܼ����Ų�ĥ������ѹ����"
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
