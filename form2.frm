VERSION 5.00
Begin VB.Form form2 
   Caption         =   "ȷ���ʹ���λ��ë�߲۳ߴ�"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "��Բ"
      Size            =   42
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "ȷ��ģ�ʹ���λ"
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8000
      Left            =   1080
      TabIndex        =   17
      Top             =   2000
      Width           =   8415
      Begin VB.CommandButton Command4 
         Caption         =   "ȷ��"
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
         Left            =   3480
         TabIndex        =   27
         Top             =   6000
         Width           =   2500
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4000
         TabIndex        =   26
         Top             =   4200
         Width           =   3000
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4000
         TabIndex        =   23
         Top             =   3200
         Width           =   3000
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4000
         TabIndex        =   21
         Top             =   2200
         Width           =   3000
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4000
         TabIndex        =   18
         Top             =   1200
         Width           =   3000
      End
      Begin VB.Label Label11 
         Caption         =   "�ʹ���λ/t"
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   500
         TabIndex        =   25
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "������ë�ߵ�ģ�ͼ��ڷ�ģ���ϵ�ͶӰ���/mm^2"
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   480
         TabIndex        =   22
         Top             =   3200
         Width           =   3500
      End
      Begin VB.Label Label2 
         Caption         =   "�ͼ���״"
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
         TabIndex        =   20
         Top             =   2200
         Width           =   2000
      End
      Begin VB.Label Label1 
         Caption         =   "�ͼ�����"
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
         TabIndex        =   19
         Top             =   1200
         Width           =   2000
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "����ģ�ʹ���λȷ��ë�߲۳ߴ�"
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8000
      Left            =   10000
      TabIndex        =   1
      Top             =   2000
      Width           =   12500
      Begin VB.CommandButton Command1 
         Caption         =   "ȷ��"
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
         Left            =   3400
         TabIndex        =   24
         Top             =   6000
         Width           =   2500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "����ë�߲���ʽ��Ӧ��"
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
         Left            =   9000
         TabIndex        =   16
         Top             =   4700
         Width           =   2500
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   500
         TabIndex        =   8
         Top             =   1875
         Width           =   2500
      End
      Begin VB.TextBox Text3 
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
         Left            =   6000
         TabIndex        =   7
         Top             =   1200
         Width           =   2500
      End
      Begin VB.TextBox Text4 
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
         Left            =   6000
         TabIndex        =   6
         Top             =   2000
         Width           =   2500
      End
      Begin VB.TextBox Text5 
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
         Left            =   6000
         TabIndex        =   5
         Top             =   2800
         Width           =   2500
      End
      Begin VB.TextBox Text6 
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
         Left            =   6000
         TabIndex        =   4
         Top             =   3600
         Width           =   2500
      End
      Begin VB.TextBox Text7 
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
         Left            =   6000
         TabIndex        =   3
         Top             =   4400
         Width           =   2500
      End
      Begin VB.TextBox Text8 
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
         Left            =   6000
         TabIndex        =   2
         Top             =   5200
         Width           =   2500
      End
      Begin VB.Image Image1 
         Height          =   1815
         Left            =   8600
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "ѡ��ʹ���λ"
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   500
         TabIndex        =   15
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "h/mm"
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3400
         TabIndex        =   14
         Top             =   1200
         Width           =   2505
      End
      Begin VB.Label Label4 
         Caption         =   "h1/mm"
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3400
         TabIndex        =   13
         Top             =   2000
         Width           =   2505
      End
      Begin VB.Label Label5 
         Caption         =   "b/mm"
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3400
         TabIndex        =   12
         Top             =   2800
         Width           =   2505
      End
      Begin VB.Label Label6 
         Caption         =   "b1/mm"
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3400
         TabIndex        =   11
         Top             =   3600
         Width           =   2505
      End
      Begin VB.Label Label7 
         Caption         =   "R1/mm"
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3400
         TabIndex        =   10
         Top             =   4400
         Width           =   2505
      End
      Begin VB.Label Label8 
         Caption         =   "ë�߲۽����/mm^2"
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3400
         TabIndex        =   9
         Top             =   5200
         Width           =   2505
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   1080
      Left            =   11100
      TabIndex        =   0
      Top             =   10500
      Width           =   2500
   End
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Combo1.AddItem ("1tģ�ʹ�")
Combo1.AddItem ("2tģ�ʹ�")
Combo1.AddItem ("3tģ�ʹ�")
Combo1.AddItem ("5tģ�ʹ�")
Combo1.AddItem ("10tģ�ʹ�")
Combo2.AddItem ("�ṹ��")
Combo2.AddItem ("�����")
Combo2.AddItem ("���ȸ�")
Combo2.AddItem ("���Ͻ�")
Combo3.AddItem ("��״����")
Combo3.AddItem ("��״�ϸ���")
Combo3.AddItem ("��״һ��")
Combo3.AddItem ("��״��")
Image1.Picture = LoadPicture(App.Path & "\3.jpg")
Image1.Stretch = False
End Sub
Private Sub Text1_Change()
If IsNumeric(Text1.Text) = False Then
Text1.Text = ""
Text1.SetFocus
End If
End Sub
Private Sub Command1_Click()
If Combo1.Text = "1tģ�ʹ�" Then
Text3.Text = "1.0-1.6"
Text4.Text = "4"
Text5.Text = "8"
Text6.Text = "22-25"
Text7.Text = "1"
Text8.Text = "100-126"
ElseIf Combo1.Text = "2tģ�ʹ�" Then
Text3.Text = "1.8-2.2"
Text4.Text = "4"
Text5.Text = "10"
Text6.Text = "25-30"
Text7.Text = "1.5"
Text8.Text = "134-168"
ElseIf Combo1.Text = "3tģ�ʹ�" Then
Text3.Text = "2.5-3.0"
Text4.Text = "5"
Text5.Text = "12"
Text6.Text = "30-40"
Text7.Text = "1.5"
Text8.Text = "207-285"
ElseIf Combo1.Text = "5tģ�ʹ�" Then
Text3.Text = "3.0-4.0"
Text4.Text = "6"
Text5.Text = "12-14"
Text6.Text = "40-50"
Text7.Text = "2"
Text8.Text = "320-440"
ElseIf Combo1.Text = "10tģ�ʹ�" Then
Text3.Text = "4.0-6.0"
Text4.Text = "8"
Text5.Text = "14-16"
Text6.Text = "50-60"
Text7.Text = "2.5"
Text8.Text = "528-728"
ElseIf Combo1.Text = "ѡȡģ�ʹ�" Then
Exit Sub
End If
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
Private Sub command3_click()
form1.Hide
form2.Hide
form3.Hide
form4.Hide
Form5.Show
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide
End Sub
Private Sub command4_click()
Dim a, b, c, d As Double
c = Val(Text1.Text)
If Combo2.Text = "�ṹ��" Then
a = 1
ElseIf Combo2.Text = "�����" Then
a = 1.5
ElseIf Combo2.Text = "���ȸ�" Then
a = 2
ElseIf Combo2.Text = "���Ͻ�" Then
a = 0.8
End If

If Combo3.Text = "��״����" Or Combo3.Text = "��״�ϸ���" Then
b = 0.1
ElseIf Combo3.Text = "��״һ��" Then
b = 0.09
ElseIf Combo3.Text = "��״��" Then
b = 0.07
End If

 d = a * b * c

If d < 0 Then
MsgBox "�������ݺ����ԣ�", , "����"
ElseIf d > 0 And d <= 1000 Then
Text2.Text = 1
ElseIf d > 1000 And d <= 2000 Then
Text2.Text = 2
ElseIf d > 2000 And d <= 3000 Then
Text2.Text = 3
ElseIf d > 3000 And d <= 5000 Then
Text2.Text = 5
ElseIf d > 5000 And d <= 10000 Then
Text2.Text = 10
ElseIf d > 10000 Then
Text2.Text = ""
MsgBox "����ѡ��Χ��", , "����"
End If

End Sub
