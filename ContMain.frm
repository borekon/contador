VERSION 5.00
Begin VB.Form ContMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contador"
   ClientHeight    =   915
   ClientLeft      =   3720
   ClientTop       =   3900
   ClientWidth     =   3525
   Icon            =   "ContMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   3525
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sumar"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuExit 
         Caption         =   "S&alir"
      End
   End
   Begin VB.Menu mnuOpcion 
      Caption         =   "Opciones"
      Begin VB.Menu mnuConfig 
         Caption         =   "Preferencias"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "Acerca de..."
   End
End
Attribute VB_Name = "ContMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim hola As Long
hola = Label1.Caption
hola = hola + frmOption.Text1.Text
Label1.Caption = hola
End Sub

Private Sub Command2_Click()
WriteINI App.Path & "\config.ini", "Veces", Label1.Caption
End
End Sub
Private Sub Form_Load()
Label1.Caption = ReadINI(App.Path & "\config.ini", "Veces")
End Sub
Private Sub mnuShow_Click()

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuConfig_Click()
frmOption.Show
End Sub
Private Sub mnuExit_Click()
End
End Sub
