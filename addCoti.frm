VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form addCoti 
   BackColor       =   &H00AAD69B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajout d'une cotisation"
   ClientHeight    =   2460
   ClientLeft      =   3750
   ClientTop       =   1260
   ClientWidth     =   4215
   Icon            =   "addCoti.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4215
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox montantCoti 
         Alignment       =   1  'Right Justify
         DataField       =   "Cotisation"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4108
            SubFormatType   =   1
         EndProperty
         DataSource      =   "datPrimaryRSCoti"
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox debiteurid 
         DataField       =   "AdhID"
         DataSource      =   "datPrimaryRSCoti"
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox echeance 
         DataField       =   "Echeance"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd.MM.yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4108
            SubFormatType   =   3
         EndProperty
         DataSource      =   "datPrimaryRSCoti"
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin MSAdodcLib.Adodc datPrimaryRSCoti 
         Height          =   330
         Left            =   2760
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         RecordSource    =   "select AdhID,Cotisation,Echeance from EcheanceCotis  where AdhID < 0 "
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
      Begin VB.Shape Shape2 
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         Height          =   1335
         Left            =   0
         Top             =   0
         Width           =   3975
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Montant :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   19
         Left            =   960
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E2F5FF&
         Caption         =   "€"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Echéance :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   18
         Left            =   2400
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Nouvelle cotisation"
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
         TabIndex        =   3
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Image btFermer 
      Height          =   375
      Left            =   2640
      Picture         =   "addCoti.frx":058A
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Image btAnnulerAdh 
      Height          =   375
      Left            =   1200
      Picture         =   "addCoti.frx":0FE6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image btFermerDown 
      Height          =   375
      Left            =   360
      Picture         =   "addCoti.frx":1ABE
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image btFermerUp 
      Height          =   375
      Left            =   240
      Picture         =   "addCoti.frx":2576
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   795
      Index           =   0
      Left            =   -240
      Picture         =   "addCoti.frx":2FD2
      Top             =   1560
      Width           =   4470
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00D5F1FF&
      FillColor       =   &H00D5F1FF&
      FillStyle       =   0  'Solid
      Height          =   2115
      Left            =   0
      Top             =   0
      Width           =   4230
   End
End
Attribute VB_Name = "addCoti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function saveCoti() As Boolean
    On Error GoTo UpdateErr
    saveCoti = False
    
   'Montant Cotisation
    montantCoti.SetFocus
    montantCoti.BackColor = &HFFFFFF
    On Error GoTo typmismatch
    s = FormatCurrency(montantCoti)
    If StrComp(montantCoti, "") = 0 Then
        GoTo typmismatch
    End If
    
    'Date d'échéance
    echeance.SetFocus
    echeance.BackColor = &HFFFFFF
    If Not IsDate(echeance) Then
            echeance.BackColor = &HF9D9D4
            MsgBox "La date d'échéance n'est pas une date valide"
    Else
        If (DateValue(echeance) <= Date) Then
            echeance.BackColor = &HF9D9D4
           MsgBox "L'échéance est dépassée"
        Else
            datPrimaryRSCoti.Recordset.UpdateBatch adAffectCurrent
            datPrimaryRSCoti.Refresh
            datPrimaryRSCoti.Refresh
                
            saveCoti = True
            
        End If
    End If

    Exit Function
    
typmismatch:
    montantCoti.BackColor = &HF9D9D4
    MsgBox "Le montant n'est pas valide"
    Exit Function
UpdateErr:
    MsgBox err.Description
End Function



Private Sub btAnnulerAdh_Click()
    datPrimaryRSCoti.Refresh
    Unload Me
End Sub

Private Sub btFermer_Click()
    If saveCoti Then
        fGbl.Enabled = True
       
        Call fGbl.selCoti(fGbl.adhid)
        Unload Me
    End If
End Sub



Private Sub datPrimaryRSCoti_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub



Private Sub echeance_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call btFermer_Click

End Sub


Private Sub Form_Initialize()
ChDir (chemin)
   
   ' WindowState = vbMaximized

End Sub

Private Sub Form_Load()

    datPrimaryRSCoti.Recordset.AddNew
    debiteurid = fGbl.adhid
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
  fGbl.Enabled = True
  fMainForm.Frametop.Enabled = True
  fGbl.nvVersMontant.SetFocus
End Sub



'
'Etats Boutons
'
Private Sub btFermer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btFermer.Picture = fGbl.btFermerDown.Picture
End Sub
Private Sub btFermer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btFermer.Picture = fGbl.btFermerUp.Picture
End Sub

Private Sub btAnnulerAdh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btAnnulerAdh.Picture = fGbl.btAnnulerDown.Picture
End Sub
Private Sub btAnnulerAdh_MouseUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btAnnulerAdh.Picture = fGbl.btAnnulerUp.Picture
End Sub

