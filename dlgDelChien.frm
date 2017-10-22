VERSION 5.00
Begin VB.Form dlgDelChien 
   BackColor       =   &H00AAD69B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supression d'un chien"
   ClientHeight    =   3180
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3615
   Icon            =   "dlgDelChien.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   3615
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.TextBox index 
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Text            =   "id"
         Top             =   2640
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         Height          =   2055
         Left            =   0
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E2F5FF&
         Caption         =   "Confirmez-vous la suppression ?"
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
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label lbdel 
         Alignment       =   2  'Center
         BackColor       =   &H00E2F5FF&
         Caption         =   "Chien"
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
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00E2F5FF&
         Caption         =   "vont être supprimés."
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
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label lbdel 
         Alignment       =   2  'Center
         BackColor       =   &H00E2F5FF&
         Caption         =   "et concours"
         ForeColor       =   &H00E03080&
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   600
         Width           =   3375
      End
      Begin VB.Image btouiUp 
         Height          =   375
         Left            =   240
         Picture         =   "dlgDelChien.frx":058A
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image btouiDown 
         Height          =   375
         Left            =   240
         Picture         =   "dlgDelChien.frx":0D86
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image btNonDown 
         Height          =   375
         Left            =   360
         Picture         =   "dlgDelChien.frx":15DE
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image btNonUp 
         Height          =   375
         Left            =   240
         Picture         =   "dlgDelChien.frx":1E36
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Image btContinuer 
      Height          =   375
      Left            =   2280
      Picture         =   "dlgDelChien.frx":2632
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image btAnnuler 
      Height          =   375
      Left            =   840
      Picture         =   "dlgDelChien.frx":2E2E
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   795
      Index           =   1
      Left            =   -240
      Picture         =   "dlgDelChien.frx":362A
      Top             =   2280
      Width           =   4470
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00D5F1FF&
      FillColor       =   &H00D5F1FF&
      FillStyle       =   0  'Solid
      Height          =   3075
      Left            =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "dlgDelChien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btAnnuler_Click()
    Unload Me
End Sub
Private Sub btAnnuler_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btAnnuler.Picture = btNonDown.Picture
End Sub
Private Sub btAnnuler_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btAnnuler.Picture = btNonUp.Picture
End Sub
Private Sub btContinuer_Click()
On Error GoTo DeleteErr
    'on se place sur le chien à supprimer
     fGbl.datPrimaryRSChi.Recordset.Move (index)
     
    'delete concours
    Call fGbl.delConcours(fGbl.idChien(0))
    
    'delete 1 chien
    fGbl.datPrimaryRSChi.Recordset.Delete
    fGbl.datPrimaryRSChi.Refresh
    fGbl.datPrimaryRSChi.Refresh
    
    ' Upd Proprio
    Call fGbl.datPrimaryRS.Recordset.Update("Nb_Chiens", fGbl.adhNbChien - 1)
    fGbl.datPrimaryRS.Refresh
    fGbl.datPrimaryRS.Refresh
        
    Call fGbl.affchienadh
    

    Unload Me
    
  Exit Sub
DeleteErr:
  MsgBox err.Description
End Sub
Private Sub btContinuer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btContinuer.Picture = btouiDown.Picture
End Sub
Private Sub btContinuer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btContinuer.Picture = btouiUp.Picture
End Sub
Private Sub Form_Load()
    fMainForm.Frametop.Enabled = False
    fGbl.Enabled = False
    Left = fGbl.Width / 2
    Top = 1000
    'fMainForm.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    fGbl.Enabled = True
    fMainForm.Frametop.Enabled = True
End Sub
