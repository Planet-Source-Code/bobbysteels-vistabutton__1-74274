VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "VistaButton Test"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4140
   LinkTopic       =   "Form2"
   ScaleHeight     =   3660
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin VistaButton_Test.VistaButton VistaButton1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      caption         =   "VistaButton1"
      forecolor       =   0
      font            =   "frmTest.frx":0000
      picture         =   "frmTest.frx":0024
      pictures        =   2
      usemaskcolor    =   -1  'True
      maskcolor       =   65280
      enabled         =   -1  'True
      nobackground    =   0   'False
      backcolor       =   -2147483633
      pictureoffset   =   0
   End
   Begin VistaButton_Test.VistaButton VistaButton3 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "ToolTip"
      Top             =   1200
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      caption         =   "VistaButton3"
      forecolor       =   0
      font            =   "frmTest.frx":0042
      picture         =   "frmTest.frx":0066
      pictures        =   2
      usemaskcolor    =   -1  'True
      maskcolor       =   65280
      enabled         =   0   'False
      nobackground    =   0   'False
      backcolor       =   -2147483633
      pictureoffset   =   0
   End
   Begin VistaButton_Test.VistaButton VistaButton2 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      caption         =   "VistaButton2"
      forecolor       =   0
      font            =   "frmTest.frx":06B8
      picture         =   "frmTest.frx":06DC
      pictures        =   1
      usemaskcolor    =   -1  'True
      maskcolor       =   65280
      enabled         =   -1  'True
      nobackground    =   0   'False
      backcolor       =   -2147483633
      pictureoffset   =   0
   End
   Begin VistaButton_Test.VistaButton VistaButton5 
      Height          =   1695
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   1815
      _extentx        =   3201
      _extenty        =   2990
      caption         =   "VistaButton5"
      forecolor       =   0
      font            =   "frmTest.frx":0A2E
      picture         =   "frmTest.frx":0A52
      pictures        =   2
      usemaskcolor    =   -1  'True
      maskcolor       =   65280
      enabled         =   -1  'True
      nobackground    =   0   'False
      backcolor       =   -2147483633
      pictureoffset   =   0
   End
   Begin VistaButton_Test.VistaButton VistaButton7 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "VistaButton7"
      Top             =   2160
      Width           =   375
      _extentx        =   661
      _extenty        =   661
      caption         =   ""
      forecolor       =   0
      font            =   "frmTest.frx":0A70
      picture         =   "frmTest.frx":0A94
      pictures        =   2
      usemaskcolor    =   -1  'True
      maskcolor       =   65280
      enabled         =   -1  'True
      nobackground    =   -1  'True
      backcolor       =   -2147483633
      pictureoffset   =   1
   End
   Begin VistaButton_Test.VistaButton VistaButton4 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
      _extentx        =   2778
      _extenty        =   450
      caption         =   "VistaButton4"
      forecolor       =   0
      font            =   "frmTest.frx":10E6
      picture         =   "frmTest.frx":110A
      pictures        =   2
      usemaskcolor    =   -1  'True
      maskcolor       =   65280
      enabled         =   -1  'True
      nobackground    =   0   'False
      backcolor       =   -2147483633
      pictureoffset   =   0
   End
   Begin VistaButton_Test.VistaButton VistaButton8 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      ToolTipText     =   "VistaButton8"
      Top             =   2160
      Width           =   375
      _extentx        =   661
      _extenty        =   661
      caption         =   ""
      forecolor       =   0
      font            =   "frmTest.frx":1128
      picture         =   "frmTest.frx":114C
      pictures        =   2
      usemaskcolor    =   -1  'True
      maskcolor       =   65280
      enabled         =   0   'False
      nobackground    =   0   'False
      backcolor       =   -2147483633
      pictureoffset   =   0
   End
   Begin VistaButton_Test.VistaButton VistaButton9 
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      ToolTipText     =   "VistaButton9"
      Top             =   2160
      Width           =   375
      _extentx        =   661
      _extenty        =   661
      caption         =   ""
      forecolor       =   0
      font            =   "frmTest.frx":179E
      picture         =   "frmTest.frx":17C2
      pictures        =   1
      usemaskcolor    =   -1  'True
      maskcolor       =   65280
      enabled         =   -1  'True
      nobackground    =   0   'False
      backcolor       =   -2147483633
      pictureoffset   =   0
   End
   Begin VistaButton_Test.VistaButton VistaButton10 
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      ToolTipText     =   "VistaButton10"
      Top             =   2160
      Width           =   375
      _extentx        =   661
      _extenty        =   661
      caption         =   ""
      forecolor       =   0
      font            =   "frmTest.frx":1B14
      picture         =   "frmTest.frx":1B38
      pictures        =   2
      usemaskcolor    =   -1  'True
      maskcolor       =   65280
      enabled         =   -1  'True
      nobackground    =   0   'False
      backcolor       =   -2147483633
      pictureoffset   =   0
   End
   Begin VistaButton_Test.VistaButton VistaButton6 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      caption         =   "VistaButton6"
      forecolor       =   12582912
      font            =   "frmTest.frx":218A
      picture         =   "frmTest.frx":21BA
      pictures        =   2
      usemaskcolor    =   -1  'True
      maskcolor       =   65280
      enabled         =   -1  'True
      nobackground    =   0   'False
      backcolor       =   -2147483633
      pictureoffset   =   0
   End
   Begin VistaButton_Test.VistaButton VistaButton11 
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   1575
      _extentx        =   2778
      _extenty        =   1085
      caption         =   "Button with a very long text"
      forecolor       =   0
      font            =   "frmTest.frx":21D8
      picture         =   "frmTest.frx":21FC
      pictures        =   2
      usemaskcolor    =   -1  'True
      maskcolor       =   65280
      enabled         =   -1  'True
      nobackground    =   0   'False
      backcolor       =   -2147483633
      pictureoffset   =   0
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub VistaButton2_Click()
    VistaButton3.Enabled = Not VistaButton3.Enabled
End Sub

Private Sub VistaButton5_MouseIn()
    VistaButton5.Caption = "MouseOver"

End Sub

Private Sub VistaButton5_MouseOut()
    VistaButton5.Caption = "MouseOut"
End Sub
