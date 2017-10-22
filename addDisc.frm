VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form addDisc 
   BackColor       =   &H00AAD69B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajout d'une discipline"
   ClientHeight    =   2715
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   4215
   Icon            =   "addDisc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4215
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox id 
         DataField       =   "id"
         DataSource      =   "datDisc"
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Disc 
         DataField       =   "Label"
         DataSource      =   "datDisc"
         Height          =   285
         Left            =   960
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2880
      End
      Begin MSAdodcLib.Adodc datDisc 
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
         RecordSource    =   "select id,Label from TypeConcours"
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
         Height          =   1575
         Left            =   0
         Top             =   0
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Nouvelle discipline"
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
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Nom de la nouvelle discipline :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   2775
      End
   End
   Begin VB.Image btFermer 
      Height          =   375
      Left            =   2640
      Picture         =   "addDisc.frx":058A
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Image btAnnuler 
      Height          =   375
      Left            =   1200
      Picture         =   "addDisc.frx":0FE6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   795
      Index           =   1
      Left            =   -240
      Picture         =   "addDisc.frx":1ABE
      Top             =   1800
      Width           =   4470
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00D5F1FF&
      FillColor       =   &H00D5F1FF&
      FillStyle       =   0  'Solid
      Height          =   2595
      Left            =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "addDisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function saveDisc() As Boolean
  On Error GoTo UpdateErr

    saveDisc = False
    If Disc = "" Then
        MsgBox "Veuillez saisir le nom de la discipline"
    Else
        tmpC = fGbl.drpTypDisc.ListIndex
        datDisc.Recordset.UpdateBatch adAffectAll
        fGbl.labs.datlstDisc.Refresh
        fGbl.labs.datlstDisc.Refresh
        'on attend que l'update soit fait
        truc = fGbl.drpTypDisc.ListCount
        i = 0
        While truc = fGbl.drpTypDisc.ListCount And i < 500
            fGbl.labs.datlstDisc.Refresh
            Call fGbl.InitDropTypDisc
            i = i + 1
        Wend
        If truc = fGbl.drpTypDisc.ListCount Then GoTo UpdateErr

        saveDisc = True
        fGbl.drpTypDisc.ListIndex = tmpC
        'fGbl.drpTypCat.ListIndex = 0
    End If
    
  Exit Function
UpdateErr:
    MsgBox "L'enregistrement n'as pas eu lieu. Vérifiez que la discipline n'exite pas déjà."
  MsgBox err.Description
End Function

Private Sub btAnnuler_Click()
    fGbl.drpTypDisc.ListIndex = -1
    datDisc.Refresh
    
    Unload Me
End Sub
Private Sub btAnnuler_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btAnnuler.Picture = fGbl.btAnnulerDown.Picture
End Sub
Private Sub btAnnuler_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btAnnuler.Picture = fGbl.btAnnulerUp.Picture
End Sub

Private Sub btFermer_Click()
    If saveDisc Then Unload Me
End Sub
Private Sub btFermer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btFermer.Picture = fGbl.btFermerDown.Picture
End Sub
Private Sub btFermer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btFermer.Picture = fGbl.btFermerUp.Picture
End Sub

Private Sub Disc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call btFermer_Click
    End If
End Sub

Private Sub Form_Load()
  fMainForm.Frametop.Enabled = False
  fGbl.Enabled = False
  datDisc.Recordset.AddNew
  ChDir (chemin)
  'Disc.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
  fGbl.Enabled = True
  fMainForm.Frametop.Enabled = True
End Sub

Private Sub datDisc_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datDisc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  datDisc.Caption = "Record: " & CStr(datDisc.Recordset.AbsolutePosition)
End Sub

Private Sub datDisc_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datDisc.Recordset.AddNew

  Exit Sub
AddErr:
  MsgBox err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datDisc.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datDisc.Refresh
  Exit Sub
RefreshErr:
  MsgBox err.Description
End Sub

Private Sub cmdUpdate_Click()

End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

