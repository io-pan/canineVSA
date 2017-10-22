VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form addChien 
   BackColor       =   &H00AAD69B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajout d'un chien"
   ClientHeight    =   6525
   ClientLeft      =   3750
   ClientTop       =   1260
   ClientWidth     =   4560
   Icon            =   "addChien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   4560
   Begin VB.Frame frameChien 
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Caption         =   "Nouveau Chien"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox cnea 
         DataField       =   "no_CNEA"
         DataSource      =   "datPrimaryRSChien"
         Height          =   285
         Left            =   1320
         TabIndex        =   23
         Top             =   4680
         Width           =   2895
      End
      Begin VB.TextBox cun 
         DataField       =   "no_CUN"
         DataSource      =   "datPrimaryRSChien"
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   5040
         Width           =   2895
      End
      Begin VB.ComboBox drpSex 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox idRaceChien 
         BackColor       =   &H8000000F&
         DataField       =   "Race"
         DataSource      =   "datPrimaryRSChien"
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   20
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox editLabelRace 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox idRace 
         DataField       =   "ID"
         DataSource      =   "datPrimaryRSRace"
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   3600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox labelRace 
         DataField       =   "Label"
         DataSource      =   "datPrimaryRSRace"
         Height          =   285
         Left            =   360
         TabIndex        =   18
         Top             =   3600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ListBox lstRaces 
         Height          =   1230
         Left            =   1320
         TabIndex        =   17
         Top             =   3360
         Width           =   2895
      End
      Begin VB.CheckBox sex 
         Caption         =   "sex"
         DataField       =   "Sexe"
         DataSource      =   "datPrimaryRSChien"
         Height          =   255
         Left            =   3600
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox dateNaiss 
         DataField       =   "DateNaiss"
         DataSource      =   "datPrimaryRSChien"
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox tatou 
         DataField       =   "Tatouage"
         DataSource      =   "datPrimaryRSChien"
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox chienlof 
         DataField       =   "LOF"
         DataSource      =   "datPrimaryRSChien"
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox nomChien 
         DataField       =   "Nom"
         DataSource      =   "datPrimaryRSChien"
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox id 
         DataField       =   "id"
         DataSource      =   "datPrimaryRSChien"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Maitreid 
         BackColor       =   &H8000000F&
         DataField       =   "MaitreID"
         DataSource      =   "datPrimaryRSChien"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSAdodcLib.Adodc datPrimaryRSChien 
         Height          =   330
         Left            =   3000
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
         RecordSource    =   "select MaitreID,id,Nom,Race,Sexe,LOF,Tatouage,DateNaiss,no_CNEA,no_CUN  from Chiens  where id<1 Order by id"
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
      Begin MSAdodcLib.Adodc datPrimaryRSRace 
         Height          =   330
         Left            =   120
         Top             =   3360
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
         RecordSource    =   "select ID,Label from Races Order by Label"
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
         BackStyle       =   0  'Transparent
         Caption         =   "Licence CNEA :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Licence CUN :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   24
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         Height          =   5415
         Left            =   0
         Top             =   0
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Ajout d'un chien"
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
         TabIndex        =   21
         Top             =   120
         Width           =   3375
      End
      Begin VB.Image btNvRace 
         Height          =   375
         Left            =   1320
         Picture         =   "addChien.frx":058A
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Date Naiss :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   15
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Tatouage :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "LOF :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   13
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Sexe :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Race :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   13
         Left            =   360
         TabIndex        =   11
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Nom :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   14
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Mid:"
         Height          =   255
         Index           =   16
         Left            =   3600
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Image btFermerChien 
      Height          =   375
      Left            =   3000
      Picture         =   "addChien.frx":19E6
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Image btAnnulerChien 
      Height          =   375
      Left            =   1560
      Picture         =   "addChien.frx":2442
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Image btNvRaceDown 
      Height          =   375
      Left            =   0
      Picture         =   "addChien.frx":2F1A
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image btNvRaceUp 
      Height          =   375
      Left            =   0
      Picture         =   "addChien.frx":43CA
      Top             =   3240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image Image3 
      Height          =   795
      Index           =   1
      Left            =   -240
      Picture         =   "addChien.frx":5826
      Top             =   5640
      Width           =   4470
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00D5F1FF&
      FillColor       =   &H00D5F1FF&
      FillStyle       =   0  'Solid
      Height          =   6435
      Left            =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "addChien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub drpSex_Click()
    If drpSex.ListIndex = 0 Then
        sex = 1
    Else
        sex = 0
    End If
End Sub
Private Sub Form_Initialize()
   
   'WindowState = vbMaximized
ChDir (chemin)
End Sub

Private Sub Form_Load()
    Top = 0
    frameChien.Visible = True
    Call InitDropRaces
    
    On Error GoTo AddErr
    datPrimaryRSChien.Recordset.AddNew
    

    
    'Init champs
    Maitreid = fGbl.adhid
    drpSex.AddItem ("M")
    drpSex.AddItem ("F")
    drpSex.ListIndex = 0
    sex = 1

Exit Sub
AddErr:
    MsgBox err.Description
End Sub

Private Sub InitDropRaces()
    
    Call lstRaces.Clear
    datPrimaryRSRace.Recordset.MoveFirst
    While (Not (datPrimaryRSRace.Recordset.EOF))
        lstRaces.AddItem (labelRace)
        datPrimaryRSRace.Recordset.MoveNext
    Wend
    datPrimaryRSRace.Recordset.MoveFirst
    lstRaces.ListIndex = 0

End Sub


Private Sub btAnnulerChien_Click()
    datPrimaryRSChien.Refresh
    Unload Me
End Sub



Private Sub btNvRace_Click()
  On Error GoTo AddErr
  
    'ajout+sauvegarde nouvelle race
    datPrimaryRSRace.Recordset.AddNew
    labelRace = editLabelRace
    datPrimaryRSRace.Recordset.UpdateBatch adAffectAll
    datPrimaryRSRace.Refresh
    datPrimaryRSRace.Refresh
    
    'Recherche id dans la base
    datPrimaryRSRace.Recordset.MoveFirst
    While (StrComp(labelRace, editLabelRace) <> 0)
        datPrimaryRSRace.Recordset.MoveNext
    Wend
    idRaceChien = idRace
   
    'refresh sur la form principale
    While fGbl.datPrimaryRSRace.Recordset.RecordCount <> datPrimaryRSRace.Recordset.RecordCount
        fGbl.datPrimaryRSRace.Refresh
    Wend
    
    'refrech liste au cas ou le 2è chien(s'il y a) soit de la même race
    Call InitDropRaces
    Call editLabelRace_KeyUp(0, 0)
    
  Exit Sub
AddErr:
    MsgBox "L'enregistrement n'as pas eu lieu. Vérifiez que la race n'exite pas déjà."
        datPrimaryRSRace.Refresh
End Sub


Private Sub btFermerChien_Click()
    If saveChien() Then
        'refresh frame globale
        fGbl.datPrimaryRSChi.Refresh
        fGbl.datPrimaryRSChi.Refresh
        Call fGbl.affchienadh
        Unload Me
    End If
End Sub



Private Function saveChien() As Boolean
    On Error GoTo UpdateErr
    
    'check Saisie Chien
    saveChien = True
    dateNaiss.SetFocus
    If Not IsDate(dateNaiss) Then
        dateNaiss.BackColor = &HF9D9D4
        err.Description = "La date de naissance n'est pas une date valide"
        GoTo UpdateErr
    Else
        dateNaiss.BackColor = &HFFFFFF
        If (DateValue(dateNaiss) > date) Then
            dateNaiss.BackColor = &HF9D9D4
            err.Description = "Le chien n'est pas né !"
            GoTo UpdateErr
        End If
    End If
    If (StrComp(idRaceChien, "") = 0) Then
        editLabelRace.BackColor = &HF9D9D4
        saveChien = False
        editLabelRace.SetFocus
    Else
        editLabelRace.BackColor = &HFFFFFF
    End If '

    If (StrComp(tatou, "") = 0) Then
        tatou.BackColor = &HF9D9D4
        saveChien = False
        tatou.SetFocus
    Else
        tatou.BackColor = &HFFFFFF
    End If
    
  
    
    If (StrComp(nomChien, "") = 0) Then
        nomChien.BackColor = &HF9D9D4
        saveChien = False
        nomChien.SetFocus
    Else
        nomChien.BackColor = &HFFFFFF
    End If

    If Not (saveChien) Then
        err.Description = "Veuillez renseigner les champs sur-lignés"
        GoTo UpdateErr
    Else
      
        'sauve
        datPrimaryRSChien.Recordset.UpdateBatch adAffectAll
        datPrimaryRSChien.Refresh
        datPrimaryRSChien.Refresh
        
        ' Upd Proprio
        Call fGbl.datPrimaryRS.Recordset.Update("Nb_Chiens", fGbl.adhNbChien + 1)
         fGbl.datPrimaryRS.Refresh
         fGbl.datPrimaryRS.Refresh
    End If
    
Exit Function
UpdateErr:
    MsgBox err.Description
    saveChien = False
End Function


Private Sub editLabelRace_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KKup And lstRaces.ListIndex > 1 Then
        lstRaces.ListIndex = lstRaces.ListIndex - 1
    ElseIf KeyCode = KKdown And lstRaces.ListIndex < lstRaces.ListCount - 1 Then
        lstRaces.ListIndex = lstRaces.ListIndex + 1

    ElseIf KeyCode = KKenter Then
        btFermerChien_Click
    Else
        memInd = lstRaces.ListIndex
        lstRaces.ListIndex = 0
        selRace = lstRaces
        While (StrComp(LCase(editLabelRace), LCase(selRace)) > 0) And (lstRaces.ListIndex < lstRaces.ListCount - 1)
           lstRaces.ListIndex = lstRaces.ListIndex + 1
           selRace = lstRaces
        Wend
        If (StrComp(LCase(editLabelRace), LCase(selRace)) > 0) Then
            lstRaces.ListIndex = memInd
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    fGbl.Enabled = True
    fMainForm.Frametop.Enabled = True
End Sub






Private Sub lstRaces_DblClick()

    ' init idRace du Chien
    idRaceChien = idRace
    editLabelRace = lstRaces ' pour la visu

End Sub

Private Sub lstRaces_Click()
    On Error GoTo err
    
    ' init idRace du Chien
   
    datPrimaryRSRace.Recordset.MoveFirst
    datPrimaryRSRace.Recordset.Move (lstRaces.ListIndex)

    idRaceChien = idRace
    
  Exit Sub
err:
  MsgBox err.Description
End Sub




Private Sub sexM_Click()
sexF = False
sex = 1
End Sub
Private Sub sexF_Click()
sexM = False
sex = 0
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
Private Sub btNvRace_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvRace.Picture = btNvRaceDown.Picture
End Sub
Private Sub btNvRace_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvRace.Picture = btNvRaceUp.Picture
End Sub
Private Sub btFermerChien_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btFermerChien.Picture = fGbl.btFermerDown.Picture
End Sub
Private Sub btFermerChien_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
   btFermerChien.Picture = fGbl.btFermerUp.Picture
End Sub

Private Sub btAnnulerChien_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerChien.Picture = fGbl.btAnnulerDown.Picture
End Sub
Private Sub btAnnulerChien_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerChien.Picture = fGbl.btAnnulerUp.Picture
End Sub

