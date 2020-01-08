VERSION 5.00
Begin VB.Form form2 
   Caption         =   "È·¶¨¶Í´¸¶ÖÎ»ÓëÃ«±ß²Û³ß´ç"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Ó×Ô²"
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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "È·¶¨Ä£¶Í´¸¶ÖÎ»"
      BeginProperty Font 
         Name            =   "Ó×Ô²"
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
         Caption         =   "È·¶¨"
         BeginProperty Font 
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
         Caption         =   "¶Í´¸¶ÖÎ»/t"
         BeginProperty Font 
            Name            =   "Ó×Ô²"
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
         Caption         =   "²»°üÀ¨Ã«±ßµÄÄ£¶Í¼þÔÚ·ÖÄ£ÃæÉÏµÄÍ¶Ó°Ãæ»ý/mm^2"
         BeginProperty Font 
            Name            =   "Ó×Ô²"
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
         Caption         =   "¶Í¼þÐÎ×´"
         BeginProperty Font 
            Name            =   "Ó×Ô²"
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
         Caption         =   "¶Í¼þ²ÄÁÏ"
         BeginProperty Font 
            Name            =   "Ó×Ô²"
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
      Caption         =   "¸ù¾ÝÄ£¶Í´¸¶ÖÎ»È·¶¨Ã«±ß²Û³ß´ç"
      BeginProperty Font 
         Name            =   "Ó×Ô²"
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
         Caption         =   "È·¶¨"
         BeginProperty Font 
            Name            =   "Ó×Ô²"
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
         Caption         =   "³£¼ûÃ«±ß²ÛÐÎÊ½ÓëÓ¦ÓÃ"
         BeginProperty Font 
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
         Caption         =   "Ñ¡Ôñ¶Í´¸¶ÖÎ»"
         BeginProperty Font 
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
            Name            =   "Ó×Ô²"
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
         Caption         =   "Ã«±ß²Û½ØÃæ»ý/mm^2"
         BeginProperty Font 
            Name            =   "Ó×Ô²"
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
      Caption         =   "·µ»Ø"
      BeginProperty Font 
         Name            =   "Ó×Ô²"
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
Combo1.AddItem ("1tÄ£¶Í´¸")
Combo1.AddItem ("2tÄ£¶Í´¸")
Combo1.AddItem ("3tÄ£¶Í´¸")
Combo1.AddItem ("5tÄ£¶Í´¸")
Combo1.AddItem ("10tÄ£¶Í´¸")
Combo2.AddItem ("½á¹¹¸Ö")
Combo2.AddItem ("²»Ðâ¸Ö")
Combo2.AddItem ("ÄÍÈÈ¸Ö")
Combo2.AddItem ("ÂÁºÏ½ð")
Combo3.AddItem ("ÐÎ×´¸´ÔÓ")
Combo3.AddItem ("ÐÎ×´½Ï¸´ÔÓ")
Combo3.AddItem ("ÐÎ×´Ò»°ã")
Combo3.AddItem ("ÐÎ×´¼òµ¥")
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
If Combo1.Text = "1tÄ£¶Í´¸" Then
Text3.Text = "1.0-1.6"
Text4.Text = "4"
Text5.Text = "8"
Text6.Text = "22-25"
Text7.Text = "1"
Text8.Text = "100-126"
ElseIf Combo1.Text = "2tÄ£¶Í´¸" Then
Text3.Text = "1.8-2.2"
Text4.Text = "4"
Text5.Text = "10"
Text6.Text = "25-30"
Text7.Text = "1.5"
Text8.Text = "134-168"
ElseIf Combo1.Text = "3tÄ£¶Í´¸" Then
Text3.Text = "2.5-3.0"
Text4.Text = "5"
Text5.Text = "12"
Text6.Text = "30-40"
Text7.Text = "1.5"
Text8.Text = "207-285"
ElseIf Combo1.Text = "5tÄ£¶Í´¸" Then
Text3.Text = "3.0-4.0"
Text4.Text = "6"
Text5.Text = "12-14"
Text6.Text = "40-50"
Text7.Text = "2"
Text8.Text = "320-440"
ElseIf Combo1.Text = "10tÄ£¶Í´¸" Then
Text3.Text = "4.0-6.0"
Text4.Text = "8"
Text5.Text = "14-16"
Text6.Text = "50-60"
Text7.Text = "2.5"
Text8.Text = "528-728"
ElseIf Combo1.Text = "Ñ¡È¡Ä£¶Í´¸" Then
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
If Combo2.Text = "½á¹¹¸Ö" Then
a = 1
ElseIf Combo2.Text = "²»Ðâ¸Ö" Then
a = 1.5
ElseIf Combo2.Text = "ÄÍÈÈ¸Ö" Then
a = 2
ElseIf Combo2.Text = "ÂÁºÏ½ð" Then
a = 0.8
End If

If Combo3.Text = "ÐÎ×´¸´ÔÓ" Or Combo3.Text = "ÐÎ×´½Ï¸´ÔÓ" Then
b = 0.1
ElseIf Combo3.Text = "ÐÎ×´Ò»°ã" Then
b = 0.09
ElseIf Combo3.Text = "ÐÎ×´¼òµ¥" Then
b = 0.07
End If

 d = a * b * c

If d < 0 Then
MsgBox "Çë¼ì²éÊý¾ÝºÏÀíÐÔ£¡", , "¾¯¸æ"
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
MsgBox "³¬³öÑ¡Ôñ·¶Î§£¡", , "¾¯¸æ"
End If

End Sub
