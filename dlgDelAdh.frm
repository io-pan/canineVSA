VERSION 5.00
Begin VB.Form dlgDelAdh 
   BackColor       =   &H00AAD69B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supression d'un adhérent"
   ClientHeight    =   3900
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3615
   Icon            =   "dlgDelAdh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   3615
   Begin VB.TextBox index 
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Text            =   "id"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
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
         Picture         =   "dlgDelAdh.frx":058A
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
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
         TabIndex        =   7
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Image btouiUp 
         Height          =   375
         Left            =   -240
         Picture         =   "dlgDelAdh.frx":0DE2
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image btNonDown 
         Height          =   375
         Left            =   -120
         Picture         =   "dlgDelAdh.frx":15DE
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image btNonUp 
         Height          =   375
         Left            =   -120
         Picture         =   "dlgDelAdh.frx":1E36
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lbdel 
         Alignment       =   2  'Center
         BackColor       =   &H00E2F5FF&
         Caption         =   "Adhérent,"
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
         TabIndex        =   6
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lbdel 
         Alignment       =   2  'Center
         BackColor       =   &H00E2F5FF&
         Caption         =   "chiens,"
         ForeColor       =   &H00E03080&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   5
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label lbdel 
         Alignment       =   2  'Center
         BackColor       =   &H00E2F5FF&
         Caption         =   "et concours"
         ForeColor       =   &H00E03080&
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   4
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label lbdel 
         Alignment       =   2  'Center
         BackColor       =   &H00E2F5FF&
         Caption         =   "cotisations,"
         ForeColor       =   &H00E03080&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lbdel 
         Alignment       =   2  'Center
         BackColor       =   &H00E2F5FF&
         Caption         =   "versements,"
         ForeColor       =   &H00E03080&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   2
         Top             =   840
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
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   1680
         Width           =   3375
      End
   End
   Begin VB.Image btContinuer 
      Height          =   375
      Left            =   2280
      Picture         =   "dlgDelAdh.frx":2632
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Image btAnnuler 
      Height          =   375
      Left            =   840
      Picture         =   "dlgDelAdh.frx":2E2E
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   795
      Index           =   1
      Left            =   -240
      Picture         =   "dlgDelAdh.frx":362A
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
Attribute VB_Name = "dlgDelAdh"
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
   
    'Suppression des cotis/ versements
    Call delCotis(fGbl.adhid, fGbl.datPrimaryRSCoti, fGbl.datPrimaryRSVers)
  
    'suppression des chiens
    With fGbl.datPrimaryRSChi.Recordset
        While .RecordCount > 0
            .MoveFirst
            'delete concours
            Call fGbl.delConcours(fGbl.idChien(0))
            .Delete
            fGbl.datPrimaryRSChi.Refresh
        Wend
    End With
    fGbl.datPrimaryRSChi.Refresh
    fGbl.datPrimaryRSChi.Refresh
    For i = 0 To maxChienAff - 1
        fGbl.frameChien(i).Visible = False
    Next
    
    'Suppression de l'adhérent
    fGbl.datPrimaryRS.Recordset.Delete
    fGbl.datPrimaryRS.Refresh
    fGbl.datPrimaryRS.Refresh
    fGbl.FrameAdh.Visible = False
    fGbl.Enabled = True
    fGbl.rechAdh.SetFocus
    Unload Me
  Exit Sub
DeleteErr:
  MsgBox err.Description
End Sub

Sub delCotis(NoAdh As Integer, datCoti, datVers)
    On Error GoTo DeleteErr
      
    ' suppression des cotis
    datCoti.RecordSource = "select AdhID,Cotisation,Echeance from EcheanceCotis where AdhID=" + CStr(NoAdh)
    datCoti.Refresh
    With datCoti.Recordset
        While .RecordCount > 0
            .MoveFirst
            .Delete
        Wend
    End With
    datCoti.Refresh
    datCoti.Refresh
   
    ' suppression des versements
    datVers.RecordSource = "select AdhID,Date,Montant,TypePayement,id from Versements where AdhID=" + CStr(NoAdh)
    datVers.Refresh
    With datVers.Recordset
        While .RecordCount > 0
            .MoveFirst
            .Delete
        Wend
    End With
    datVers.Refresh
    datVers.Refresh
    fGbl.FrameCoti.Visible = False
    
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
