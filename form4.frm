VERSION 5.00
Begin VB.Form form4 
   Caption         =   "锻件内孔直径的机械加工余量"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "幼圆"
      Size            =   14.25
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "锻件内孔直径机械加工余量"
      Height          =   5000
      Left            =   6000
      TabIndex        =   2
      Top             =   2500
      Width           =   11000
      Begin VB.ComboBox Combo2 
         Height          =   405
         Left            =   2400
         TabIndex        =   5
         Top             =   2400
         Width           =   2500
      End
      Begin VB.ComboBox Combo1 
         Height          =   405
         Left            =   2400
         TabIndex        =   4
         Top             =   1200
         Width           =   2500
      End
      Begin VB.TextBox Text1 
         Height          =   500
         Left            =   8000
         TabIndex        =   3
         Top             =   1200
         Width           =   2500
      End
      Begin VB.Label Label4 
         Caption         =   "余量/mm"
         Height          =   615
         Left            =   6400
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "余量/mm"
         Height          =   495
         Left            =   6600
         TabIndex        =   8
         Top             =   -360
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "孔深/mm"
         Height          =   500
         Left            =   800
         TabIndex        =   7
         Top             =   2400
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "孔径/mm"
         Height          =   500
         Left            =   800
         TabIndex        =   6
         Top             =   1200
         Width           =   1500
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   1080
      Left            =   12000
      TabIndex        =   1
      Top             =   8400
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   1080
      Left            =   7000
      TabIndex        =   0
      Top             =   8400
      Width           =   3000
   End
End
Attribute VB_Name = "form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Combo1.AddItem ("<25")
Combo1.AddItem ("25-40")
Combo1.AddItem ("40-63")
Combo1.AddItem ("63-100")
Combo1.AddItem ("100-160")
Combo1.AddItem ("160-250")  'combo1为孔径
Combo2.AddItem ("<63")
Combo2.AddItem ("63-100")
Combo2.AddItem ("100-140")
Combo2.AddItem ("140-200")
Combo2.AddItem ("200-280") 'combo2为孔深
End Sub
Private Sub Command2_click()
Text1.Text = ""
form1.Show
form2.Hide
form3.Hide
form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide          '返回主界面
End Sub
Private Sub Command1_Click()
If Combo1.Text = "<25" And Combo2.Text = "<63" Then
Text1.Text = "2.0"
ElseIf Combo1.Text = "<25" And Combo2.Text = "63-100" Then
Text1.Text = ""
MsgBox "无效参数！", , "警告"
ElseIf Combo1.Text = "<25" And Combo2.Text = "100-140" Then
Text1.Text = ""
MsgBox "无效参数！", , "警告"
ElseIf Combo1.Text = "<25" And Combo2.Text = "140-200" Then
Text1.Text = ""
MsgBox "无效参数！", , "警告"
ElseIf Combo1.Text = "<25" And Combo2.Text = "200-280" Then
Text1.Text = ""
MsgBox "无效参数！", , "警告"
ElseIf Combo1.Text = "25-40" And Combo2.Text = "<63" Then
Text1.Text = "2.0"
ElseIf Combo1.Text = "25-40" And Combo2.Text = "63-100" Then
Text1.Text = "2.6"
ElseIf Combo1.Text = "25-40" And Combo2.Text = "100-140" Then
Text1.Text = ""
MsgBox "无效参数！", , "警告"
ElseIf Combo1.Text = "25-40" And Combo2.Text = "140-200" Then
Text1.Text = ""
MsgBox "无效参数！", , "警告"
ElseIf Combo1.Text = "25-40" And Combo2.Text = "200-280" Then
Text1.Text = ""
MsgBox "无效参数！", , "警告"
ElseIf Combo1.Text = "40-63" And Combo2.Text = "<63" Then
Text1.Text = "2.0"
ElseIf Combo1.Text = "40-63" And Combo2.Text = "63-100" Then
Text1.Text = "2.6"
ElseIf Combo1.Text = "40-63" And Combo2.Text = "100-140" Then
Text1.Text = "3.0"
ElseIf Combo1.Text = "40-63" And Combo2.Text = "140-200" Then
Text1.Text = ""
MsgBox "无效参数！", , "警告"
ElseIf Combo1.Text = "40-63" And Combo2.Text = "200-280" Then
Text1.Text = ""
MsgBox "无效参数！", , "警告"
ElseIf Combo1.Text = "63-100" And Combo2.Text = "<63" Then
Text1.Text = "2.5"
ElseIf Combo1.Text = "63-100" And Combo2.Text = "63-100" Then
Text1.Text = "3.0"
ElseIf Combo1.Text = "63-100" And Combo2.Text = "100-140" Then
Text1.Text = "3.0"
ElseIf Combo1.Text = "63-100" And Combo2.Text = "140-200" Then
Text1.Text = "4.0"
ElseIf Combo1.Text = "63-100" And Combo2.Text = "200-280" Then
Text1.Text = ""
MsgBox "无效参数！", , "警告"
ElseIf Combo1.Text = "100-160" And Combo2.Text = "<63" Then
Text1.Text = "2.6"
ElseIf Combo1.Text = "100-160" And Combo2.Text = "63-100" Then
Text1.Text = "3.0"
ElseIf Combo1.Text = "100-160" And Combo2.Text = "100-140" Then
Text1.Text = "3.4"
ElseIf Combo1.Text = "100-160" And Combo2.Text = "140-200" Then
Text1.Text = "4.0"
ElseIf Combo1.Text = "100-160" And Combo2.Text = "200-280" Then
Text1.Text = "4.6"
ElseIf Combo1.Text = "160-250" And Combo2.Text = "<63" Then
Text1.Text = "3.0"
ElseIf Combo1.Text = "160-250" And Combo2.Text = "63-100" Then
Text1.Text = "3.0"
ElseIf Combo1.Text = "160-250" And Combo2.Text = "100-140" Then
Text1.Text = "3.4"
ElseIf Combo1.Text = "160-250" And Combo2.Text = "140-200" Then
Text1.Text = "4.0"
ElseIf Combo1.Text = "160-250" And Combo2.Text = "200-280" Then
Text1.Text = "4.6"
End If
End Sub

