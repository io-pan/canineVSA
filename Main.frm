VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm fMain 
   BackColor       =   &H00D5F1FF&
   Caption         =   "Education canine Ville-sous-Anjou"
   ClientHeight    =   6450
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11160
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00AAD69B&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11100
      TabIndex        =   0
      Top             =   0
      Width           =   11160
      Begin VB.Frame Frametop 
         BackColor       =   &H00AAD69B&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   10815
         Begin VB.ComboBox typLst 
            Height          =   315
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   120
            Width           =   2175
         End
         Begin VB.Image btExpUp 
            Height          =   315
            Left            =   480
            Picture         =   "Main.frx":058A
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Image btExpDown 
            Height          =   315
            Left            =   240
            Picture         =   "Main.frx":0BBB
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Image btImpDown 
            Height          =   315
            Left            =   120
            Picture         =   "Main.frx":1241
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Image btImpUp 
            Height          =   315
            Left            =   0
            Picture         =   "Main.frx":189D
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Image btExp 
            Height          =   315
            Left            =   4680
            Picture         =   "Main.frx":1EA3
            Top             =   120
            Width           =   1215
         End
         Begin VB.Image btOpenLst 
            Height          =   315
            Left            =   8205
            Picture         =   "Main.frx":24D4
            Top             =   120
            Width           =   315
         End
         Begin VB.Image btImp 
            Height          =   315
            Left            =   3360
            Picture         =   "Main.frx":2764
            Top             =   120
            Width           =   1215
         End
         Begin VB.Image btMajSiteDown 
            Height          =   315
            Left            =   -1080
            Picture         =   "Main.frx":2D6A
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Image btMajSiteUp 
            Height          =   315
            Left            =   -960
            Picture         =   "Main.frx":3702
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Image btMajSite 
            Height          =   315
            Left            =   8760
            Picture         =   "Main.frx":406A
            Top             =   120
            Width           =   1575
         End
         Begin VB.Image btAddAdh 
            Height          =   315
            Left            =   1560
            Picture         =   "Main.frx":49D2
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.Image btNvadhDown 
         Height          =   315
         Left            =   240
         Picture         =   "Main.frx":6402
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Image btNvAdhUp 
         Height          =   315
         Left            =   120
         Picture         =   "Main.frx":7E32
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Image btOpenDown 
         Height          =   315
         Left            =   600
         Picture         =   "Main.frx":9862
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Image btOpenUp 
         Height          =   315
         Left            =   240
         Picture         =   "Main.frx":9B1E
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   720
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private WithEvents m_cZip As cZip

Private Sub btNvAdh_Click()
    Dim adAdh As New frmGlobale
    adAdh.Show
End Sub
Private Sub btNvAdh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
btNvAdh.Picture = btNvDown.Picture
End Sub
Private Sub btNvAdh_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
btNvAdh.Picture = btNvUp.Picture
End Sub



Private Sub btAddAdh_Click()
    Dim adah As New addAdh
    fermWin
    fGbl.imgBienvenu.Visible = False
    fGbl.Enabled = False
    fMainForm.Frametop.Enabled = False
    adah.Show
End Sub

Private Sub btAddAdh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAddAdh.Picture = btNvadhDown.Picture
End Sub

Private Sub btAddAdh_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAddAdh.Picture = btNvAdhUp.Picture
End Sub


Private Sub btExp_Click()

    bContinue = True
    fermWin
    
     On Error GoTo PathNotfound
    'création de la mise à jour zippée
    If Dir(destMAJ) <> "" Then ' il y a déjà une maj sur la disquette
        'on l'efface ou bien ?
        If MsgBox("Une mise à jour est présente sur la disquette. Elle va être remplacée par celle-ci." + Chr(13) + Chr(13) + "Confirmez-vous le remplacement ?", vbYesNo) <> vbYes Then
            bContinue = False
        End If
    End If
    
    If bContinue Then
    
     Unload fGbl
    
     Call FileCopy("canine.MDB", "maj.mdb")
    'Zipage
     Set m_cZip = New cZip
     m_cZip.ZipFile = chemin + "maj.zip"
     m_cZip.ClearFileSpecs
     m_cZip.AddFileSpec ("maj.mdb")
     m_cZip.Zip
    
    Call FileCopy("maj.zip", destMAJ)
    Kill ("MAJ.MDB")
    Kill ("MAJ.ZIP")
    
    Call MsgBox("Base de donnée exportée." + Chr(13) + Chr(13) + "Vous pouvez retirer la disquette.", vbInformation)
    
    Set fGbl = New frmGlobale
    fGbl.Show
    MDIForm_Resize
        
    End If
    
    Exit Sub
    
PathNotfound:
           MsgBox ("La mise à jour n'a pas été enregistrée." + Chr(13) + "Vérifiez que la disquette est bien insérée." + CStr(ret))
           Set fGbl = New frmGlobale
           fGbl.Show
           MDIForm_Resize
End Sub

Private Sub btImp_Click()
    On Error GoTo PathNotfound
    If (Dir(destMAJ)) = "" Then
        MsgBox ("La mise à jour est introuvable." + Chr(13) + "Vérifiez que la disquette est bien insérée et qu'elle contient une mise à jour.")
    Else
    
        fermWin
        
        'récupération de la mise à jour zippée
        Call FileCopy(destMAJ, chemin + "maj.zip")
        
        'déZip
         Set m_cUnzip = New cUnzip
         m_cUnzip.ZipFile = "maj.zip"
         m_cUnzip.Directory
           
        ' selected do entire zip:
        If Not bSel Then
           For iItem = 1 To m_cUnzip.FileCount
              m_cUnzip.FileSelected(iItem) = True
           Next iItem
        End If
        m_cUnzip.Unzip
             
        datMaj = FileDateTime("maj.mdb")
        
        If Dir("canine.mdb") <> "" Then
            datdb = FileDateTime("Canine.mdb")
        Else 'pas encore de base, c'est pas grave
            datdb = DateValue(dateMin)
        End If
        
        bContinue = True
        If datdb > datMaj Then
            If MsgBox("La base de données qui se trouve sur l'ordinateur est plus récente que la mise à jour." + Chr(13) + Chr(13) + "Confirmez-vous la mise à jour ?", vbYesNo) = vbNo Then
                bContinue = False
            End If
        End If
        If bContinue Then
        
            'MAJ
            If Not fGbl Is Nothing Then Unload fGbl
            Call FileCopy("maj.mdb", "canine.mdb")
            
            'Vérification qu'on a bien une base
            If Dir("canine.mdb") <> "" Then
                            Call MsgBox("Mise à jour effectuée." + Chr(13) + Chr(13) + "Vous pouvez retirer la disquette.", vbInformation)
                Set fGbl = New frmGlobale
                fGbl.Show
                fGbl.WindowState = 0
                MDIForm_Resize
            Else
                    If MsgBox("La base de données n'a pas été importée." + Chr(13) + Chr(13) + "Voulez-vous réessayer ?", vbYesNo) = vbYes Then
                       btImp_Click
                    Else
                        Unload Me
                    End If
            End If
            
            Call Kill("maj.zip")
            Call Kill("maj.mdb")
        End If
           
        
    End If 'MAJ dispo
    Exit Sub

PathNotfound:
           
    MsgBox ("La mise à jour n'a pas eu lieu." + Chr(13) + "Vérifiez que la disquette est bien insérée." + CStr(ret))
           Set fGbl = New frmGlobale
           fGbl.Show
           MDIForm_Resize
End Sub
Private Sub btImp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btImp.Picture = btImpDown.Picture
End Sub
Private Sub btImp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btImp.Picture = btImpUp.Picture
End Sub
Private Sub btExp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btExp.Picture = btExpDown.Picture
End Sub
Private Sub btExp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btExp.Picture = btExpUp.Picture
End Sub


Private Sub btMajSite_Click()

    fermWin
    
    
    fGbl.Visible = False
    fMainForm.Frametop.Visible = False
    
    Set fMajConc = New siteConc
    fMajConc.Top = 50
    fMajConc.WindowState = 0
    If Width > fMajConc.Width Then
        fMajConc.Left = Width / 2 - fMajConc.Width / 2 - 120
    Else
        fMajConc.Left = 0
    End If
    fMajConc.Show
    

End Sub
Private Sub btMajSite_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btMajSite.Picture = btMajSiteDown.Picture
End Sub
Private Sub btMajSite_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btMajSite.Picture = btMajSiteUp.Picture
End Sub

Private Sub btOpenLst_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btOpenLst.Picture = btOpenDown.Picture
End Sub

Private Sub btOpenLst_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btOpenLst.Picture = btOpenUp.Picture
End Sub






Private Sub Command1_Click()
Set m_cZip = New cZip

m_cZip.ZipFile = chemin + "maj.zip"
m_cZip.ClearFileSpecs
m_cZip.AddFileSpec ("maj.mdb")
m_cZip.Zip


End Sub


Private Sub MDIForm_Load()
    typLst.AddItem (" - Choisir une liste - ")
    typLst.AddItem ("Adhérents et Chiens")
    typLst.AddItem ("SCRA")
    typLst.AddItem ("Licences CNEA")
    typLst.AddItem ("Licences CUN")
    typLst.AddItem ("Chiens LOF")
    typLst.AddItem ("Chiens par Race")
    typLst.AddItem ("Chiens par Sexe")
    typLst.AddItem ("Concours")
    typLst.AddItem ("Concours par Chien")
    typLst.AddItem ("Echéances Cotisation")
    typLst.AddItem ("Echéances Versements")
    
    typLst.ListIndex = 0

    If Dir("canine.mdb") <> "" Then
        Set fGbl = New frmGlobale
        fGbl.Show
        fGbl.WindowState = 0
        MDIForm_Resize
    Else
            If MsgBox("La base de données introuvable." + Chr(13) + Chr(13) + "Voulez-vous en importer une ?", vbYesNo) = vbYes Then
               btImp_Click
            Else
                Unload Me
            End If
    End If

    
End Sub
Private Sub btOpenLst_Click()
    Dim sFIle As String

    With dlgCommonDialog
        .InitDir = cheminListes
        .DialogTitle = "Ouverture d'une liste"
        .CancelError = False
        .Filename = ""
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "Listes (*.rtf)|*.rtf"
        .ShowOpen
        If Len(.Filename) = 0 Then
            Exit Sub
        End If
        sFIle = .Filename

    End With
    ChDir (chemin)
    Dim d As New fListes
    
    d.Show
    fGbl.Enabled = False

    d.rtfText.LoadFile sFIle
    d.Caption = sFIle
    
End Sub
Private Sub MDIForm_Resize()
On Error GoTo fin
    fGbl.Top = 0
    If Width > fGbl.Width Then
        fGbl.Left = (Width / 2) - (fGbl.Width / 2) - 120
    Else
        fGbl.Left = 0
    End If
    Frametop.Left = fGbl.Left
    
    fMajConc.Top = 50
    If Width > fMajConc.Width Then
        fMajConc.Left = (Width / 2) - (fMajConc.Width / 2) - 120
    Else
        fMajConc.Left = 0
    End If
    
    
fin:
End Sub

Private Sub nvAdh_Click()
    Dim adAdh As New addAdh
    adAdh.Show
End Sub




Private Sub typLst_click()
    If typLst.ListIndex > 0 Then
        Dim d As New fListes
        fGbl.Enabled = False
        d.Show
        d.typLst.ListIndex = typLst.ListIndex
    End If
    typLst.ListIndex = 0
End Sub

