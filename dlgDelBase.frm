VERSION 5.00
Begin VB.Form dlgWin 
   BackColor       =   &H00AAD69B&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3900
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3615
   Icon            =   "dlgDelBase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   3615
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.Shape Shape2 
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         Height          =   2775
         Left            =   0
         Top             =   0
         Width           =   3375
      End
      Begin VB.Image btouiDown 
         Height          =   375
         Left            =   -120
         Picture         =   "dlgDelBase.frx":058A
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label msgconfirm 
         Alignment       =   2  'Center
         BackColor       =   &H00E2F5FF&
         Caption         =   "Confirmez-vous ?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1095
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Image btouiUp 
         Height          =   375
         Left            =   -240
         Picture         =   "dlgDelBase.frx":0DE2
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image btNonDown 
         Height          =   375
         Left            =   -120
         Picture         =   "dlgDelBase.frx":15DE
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image btNonUp 
         Height          =   375
         Left            =   -120
         Picture         =   "dlgDelBase.frx":1E36
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label msg 
         Alignment       =   2  'Center
         BackColor       =   &H00E2F5FF&
         Caption         =   "msg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF2E55&
         Height          =   1335
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Image btContinuer 
      Height          =   375
      Left            =   2280
      Picture         =   "dlgDelBase.frx":2632
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Image btAnnuler 
      Height          =   375
      Left            =   840
      Picture         =   "dlgDelBase.frx":2E2E
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   795
      Index           =   1
      Left            =   -240
      Picture         =   "dlgDelBase.frx":362A
      Top             =   3000
      Width           =   4470
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00D5F1FF&
      FillColor       =   &H00D5F1FF&
      FillStyle       =   0  'Solid
      Height          =   3195
      Left            =   0
      Top             =   -120
      Width           =   3615
   End
End
Attribute VB_Name = "dlgWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btAnnuler_Click()
  fMainForm.confirm = 0
  Unload Me
End Sub

Private Sub btAnnuler_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btAnnuler.Picture = btNonDown.Picture
End Sub
Private Sub btAnnuler_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btAnnuler.Picture = btNonUp.Picture
End Sub

Private Sub btContinuer_Click()
  fMainForm.confirm = 1
  Unload Me
End Sub

Private Sub btContinuer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btContinuer.Picture = btouiDown.Picture
End Sub
Private Sub btContinuer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btContinuer.Picture = btouiUp.Picture
End Sub

Private Sub Form_Load()
    fGbl.Enabled = False
    fMainForm.Frametop.Enabled = False
    Left = fGbl.Width / 2
    Top = 1000
   ' fMainForm.Enabled = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    fGbl.Enabled = True
    fMainForm.Frametop.Enabled = True
End Sub

Private Sub lbdel_Click(index As Integer)

End Sub

