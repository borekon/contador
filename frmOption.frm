VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preferencias"
   ClientHeight    =   1185
   ClientLeft      =   4695
   ClientTop       =   4200
   ClientWidth     =   2640
   Icon            =   "frmOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nº a sumar"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Left            =   840
         Max             =   1
         Min             =   20
         TabIndex        =   1
         Top             =   300
         Value           =   20
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
WriteINI App.Path & "\config.ini", "Suma", Text1.Text
Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
VScroll1.Value = ReadINI(App.Path & "\config.ini", "Suma")
Text1.Text = VScroll1.Value
End Sub
Private Sub VScroll1_Change()
Text1.Text = VScroll1.Value
End Sub
