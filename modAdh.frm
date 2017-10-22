VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form modAdh 
   BackColor       =   &H00AAD69B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modification d'un adhérent"
   ClientHeight    =   4200
   ClientLeft      =   3750
   ClientTop       =   1260
   ClientWidth     =   4590
   Icon            =   "modAdh.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4590
   Begin VB.Frame FrameAdh 
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Caption         =   "Nouvel Adhérent"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4335
      Begin VB.TextBox no_scra 
         DataField       =   "no_SCRA"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Ville"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   6
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Prenom"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtFields 
         DataField       =   "CodePostal"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   5
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox adhid 
         DataField       =   "id"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox nomAdh 
         DataField       =   "Nom"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Tel_Domicile"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "aA.DD.EE.DD.DD"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4108
            SubFormatType   =   0
         EndProperty
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   2
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Tel_Portable"
         DataSource      =   "datPrimaryRS"
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   4
         Left            =   2880
         TabIndex        =   3
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Adresse"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   4
         Top             =   2400
         Width           =   2895
      End
      Begin MSAdodcLib.Adodc datPrimaryRS 
         Height          =   330
         Left            =   2760
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=canine.MDB;"
         OLEDBString     =   "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=canine.MDB;"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   $"modAdh.frx":058A
         Caption         =   " "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "N° SCRA :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         Height          =   3015
         Left            =   0
         Top             =   0
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Modification d'un adhérent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Téléphone :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   27
         Left            =   360
         TabIndex        =   15
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Nom :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Prenom :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Domicile"
         ForeColor       =   &H80000015&
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   12
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Portable"
         ForeColor       =   &H80000015&
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   11
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Adresse :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   10
         Top             =   2400
         Width           =   855
      End
   End
   Begin VB.Image btFermer 
      Height          =   375
      Left            =   3000
      Picture         =   "modAdh.frx":0613
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Image btAnnulerAdh 
      Height          =   375
      Left            =   1560
      Picture         =   "modAdh.frx":106F
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   795
      Index           =   0
      Left            =   -240
      Picture         =   "modAdh.frx":1B47
      Top             =   3240
      Width           =   4470
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00D5F1FF&
      FillColor       =   &H00D5F1FF&
      FillStyle       =   0  'Solid
      Height          =   4035
      Left            =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "modAdh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btAnnulerAdh_Click()
     datPrimaryRS.Refresh
     Unload Me
End Sub






Private Function saveProprio() As Boolean

    On Error GoTo UpdateErr
   
    'check Saisie Proprio
    saveProprio = True
   
    If (StrComp(nomAdh, "") = 0) Then
        nomAdh.BackColor = &HF9D9D4
        saveProprio = False
        nomAdh.SetFocus
    Else
        nomAdh.BackColor = &HFFFFFF
    End If
    If (StrComp(txtFields(2), "") = 0) Then
        txtFields(2).BackColor = &HF9D9D4
        saveProprio = False
    Else
        txtFields(2).BackColor = &HFFFFFF
    End If

    If Not (saveProprio) Then
        err.Description = "Veuillez renseigner les champs sur-lignés"
        GoTo UpdateErr
    Else
        'sauve
        datPrimaryRS.Recordset.UpdateBatch adAffectAll
        datPrimaryRS.Refresh
        datPrimaryRS.Refresh
        datPrimaryRS.Recordset.MoveLast
        memId = adhid
        'select uniquement ce proprioproprio
        datPrimaryRS.RecordSource = "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Adresse,CodePostal,Ville,Nb_Chiens,No_SCRA from Proprietaire where id=" & memId & ""
        datPrimaryRS.Refresh
        datPrimaryRS.Refresh
    End If
Exit Function
UpdateErr:
    MsgBox err.Description
    saveProprio = False
End Function



Private Sub btFermer_Click()
    If saveProprio Then
        fGbl.Enabled = True
        Call fGbl.selAdh(adhid)
        Call fGbl.selCoti(adhid)
        nbrc = fGbl.selChienAdh(adhid)
        Unload Me
    End If
End Sub


Private Sub Form_Initialize()
ChDir (chemin)
' WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    Top = 0
    'Récupération de l'adh
    datPrimaryRS.RecordSource = "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Adresse,CodePostal,Ville,Nb_Chiens,No_SCRA from Proprietaire where id=" + fGbl.adhid
    datPrimaryRS.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
  fGbl.Enabled = True
  fMainForm.Frametop.Enabled = True
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then
  adStatus = adStatusCancel
  End If
End Sub





'
'Etats Boutons
'
Private Sub btFermer_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btFermer.Picture = fGbl.btFermerDown.Picture
End Sub
Private Sub btFermer_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btFermer.Picture = fGbl.btFermerUp.Picture
End Sub
Private Sub btAnnulerAdh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerAdh.Picture = fGbl.btAnnulerDown.Picture
End Sub
Private Sub btAnnulerAdh_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerAdh.Picture = fGbl.btAnnulerUp.Picture
End Sub

