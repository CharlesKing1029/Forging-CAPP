VERSION 5.00
Begin VB.Form form1 
   Caption         =   "����ģ�ͼ����"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "Ӧ�þ���"
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
      Left            =   18000
      TabIndex        =   6
      Top             =   10000
      Width           =   3800
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��������Ҫ��"
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
      Left            =   18000
      TabIndex        =   5
      Top             =   7000
      Width           =   3800
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ȷ���ͼ���״����ϵ����ͼ�����ϵ��"
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
      Left            =   18000
      TabIndex        =   4
      Top             =   1000
      Width           =   3800
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ȷ���ͼ�����ģб����Բ�ǰ뾶"
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
      Left            =   18000
      TabIndex        =   3
      Top             =   5500
      Width           =   3800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ȷ���ͼ��ڿ�ֱ���Ļ�е�ӹ���������"
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
      Left            =   18000
      TabIndex        =   2
      Top             =   4000
      Width           =   3800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���ݶͼ����ȡ���ȡ��߶�ȷ����е�ӹ����������빫��"
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
      Left            =   18000
      TabIndex        =   1
      Top             =   2500
      Width           =   3800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��ģ�ʹ���λ��ë�߲۳ߴ�"
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
