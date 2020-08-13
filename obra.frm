VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16440
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   16440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "SALIR"
      Height          =   735
      Left            =   11280
      TabIndex        =   11
      Top             =   6120
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RESET"
      Height          =   735
      Left            =   11280
      TabIndex        =   10
      Top             =   4800
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULAR CUOTA"
      Height          =   735
      Left            =   1800
      TabIndex        =   7
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   975
      Left            =   1320
      TabIndex        =   2
      Top             =   3720
      Width           =   13335
      Begin VB.OptionButton Option4 
         Caption         =   "45 A 55"
         Height          =   255
         Left            =   10080
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "25 A 35"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "35 A 45"
         Height          =   255
         Left            =   10080
         TabIndex        =   4
         Top             =   240
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         Caption         =   "18 A 25"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   5640
      TabIndex        =   1
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   10080
      TabIndex        =   15
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   10080
      TabIndex        =   14
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   5520
      TabIndex        =   13
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   5520
      TabIndex        =   12
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "COSTO ANUAL"
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "COSTO MENSUAL"
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CALCULADOR TARIFARIO OBRAS SOCIALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Select Case List1.ListIndex
Case 0
If Option1.Value = True Then
Label4.Caption = 100
Label5.Caption = 100 * 11
Label6.Caption = "Osde"
End If
If Option2.Value = True Then
Label4.Caption = 150
Label5.Caption = 150 * 11
Label6.Caption = "Medicus"
End If
If Option3.Value = True Then
Label4.Caption = 300
Label5.Caption = 300 * 11
Label6.Caption = "Galeno"
End If
If Option4.Value = True Then
Label4.Caption = 400
Label5.Caption = 400 * 11
Label6.Caption = "Accord Salud"
End If
Case 1
If Option1.Value = True Then
Label4.Caption = 120
Label5.Caption = 120 * 11
Label6.Caption = "Osde"
End If
If Option2.Value = True Then
Label4.Caption = 180
Label5.Caption = 180 * 11
Label6.Caption = "Medicus"
End If
If Option3.Value = True Then
Label4.Caption = 330
Label5.Caption = 330 * 11
Label6.Caption = "Galeno"
End If
If Option4.Value = True Then
Label4.Caption = 480
Label5.Caption = 480 * 11
Label6.Caption = "Accord Salud"
End If
Case 2
If Option1.Value = True Then
Label4.Caption = 200
Label5.Caption = 200 * 11
Label6.Caption = "Osde"
End If
If Option2.Value = True Then
Label4.Caption = 180
Label5.Caption = 180 * 11
Label6.Caption = "Medicus"
End If
If Option3.Value = True Then
Label4.Caption = 330
Label5.Caption = 330 * 11
Label6.Caption = "Galeno"
End If
If Option4.Value = True Then
Label4.Caption = 480
Label5.Caption = 480 * 11
Label6.Caption = "Accord Salud"
End If
Case 3
If Option1.Value = True Then
Label4.Caption = 200
Label5.Caption = 200 * 11
Label6.Caption = "Osde"
End If
If Option2.Value = True Then
Label4.Caption = 290
Label5.Caption = 290 * 11
Label6.Caption = "Medicus"
End If
If Option3.Value = True Then
Label4.Caption = 400
Label5.Caption = 400 * 11
Label6.Caption = "Galeno"
End If
If Option4.Value = True Then
Label4.Caption = 600
Label5.Caption = 600 * 11
Label6.Caption = "Accord Salud"
End If
End Select
End Sub

Private Sub Command2_Click()
Option1 = False
Option2 = False
Option3 = False
Option4 = False
Label4.Caption = ""
Label5.Caption = ""
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
List1.AddItem "1000"
List1.AddItem "2000"
List1.AddItem "3000"
List1.AddItem "4000"
End Sub
