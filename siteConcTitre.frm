VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form siteConc 
   BackColor       =   &H00D5F1FF&
   BorderStyle     =   0  'None
   Caption         =   "Mise à jour du site"
   ClientHeight    =   7215
   ClientLeft      =   1050
   ClientTop       =   0
   ClientWidth     =   12030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7095
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.Frame fConc 
         BackColor       =   &H00D5F1FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H00AAD69B&
         Height          =   5895
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   8775
         Begin VB.Frame Frame3 
            BackColor       =   &H00D5F1FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   495
            Left            =   3000
            TabIndex        =   22
            Top             =   1200
            Width           =   3375
            Begin VB.Image ajout 
               Height          =   375
               Left            =   240
               Picture         =   "siteConcTitre.frx":0000
               Top             =   60
               Width           =   1215
            End
            Begin VB.Image suppr 
               Height          =   375
               Left            =   1680
               Picture         =   "siteConcTitre.frx":08C4
               Top             =   60
               Width           =   1215
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00E2F5FF&
               BorderColor     =   &H00AAD69B&
               FillColor       =   &H00E2F5FF&
               FillStyle       =   0  'Solid
               Height          =   495
               Left            =   120
               Top             =   0
               Width           =   2895
            End
         End
         Begin VB.CheckBox resultats 
            BackColor       =   &H00D5F1FF&
            Caption         =   "Mettre à jour"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E03080&
            Height          =   195
            Left            =   5280
            TabIndex        =   6
            Top             =   5520
            Width           =   1455
         End
         Begin VB.TextBox idConc 
            DataField       =   "id"
            DataSource      =   "datConc"
            Height          =   285
            Left            =   7320
            TabIndex        =   2
            Top             =   720
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSAdodcLib.Adodc datConc 
            Height          =   330
            Left            =   7560
            Top             =   4680
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
            RecordSource    =   "select id,Titre,Soustitre ,l1b,l1v,l2b,l2v,l3b,l3v,l4b,l4v,l5b,l5v  from siteConc  Order by id"
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
         Begin VB.Frame fManif 
            BackColor       =   &H00D5F1FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   3495
            Left            =   480
            TabIndex        =   8
            Top             =   1200
            Width           =   8055
            Begin VB.TextBox cpt 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   4108
                  SubFormatType   =   1
               EndProperty
               Height          =   285
               Left            =   7320
               TabIndex        =   21
               Text            =   "0"
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox txtViolet 
               BackColor       =   &H00FFFFFF&
               DataField       =   "l1v"
               DataSource      =   "datConc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00E03080&
               Height          =   315
               Index           =   1
               Left            =   4320
               TabIndex        =   20
               Text            =   "Texte"
               Top             =   1680
               Width           =   3375
            End
            Begin VB.TextBox txtBleu 
               BackColor       =   &H00FFFFFF&
               DataField       =   "l1b"
               DataSource      =   "datConc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF2E55&
               Height          =   285
               Index           =   1
               Left            =   600
               TabIndex        =   19
               Text            =   "Texte"
               Top             =   1680
               Width           =   3615
            End
            Begin VB.TextBox txtViolet 
               BackColor       =   &H00FFFFFF&
               DataField       =   "l2v"
               DataSource      =   "datConc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00E03080&
               Height          =   315
               Index           =   2
               Left            =   4320
               TabIndex        =   18
               Text            =   "Texte"
               Top             =   2040
               Width           =   3375
            End
            Begin VB.TextBox Titre 
               BackColor       =   &H00FFFFFF&
               DataField       =   "Titre"
               DataSource      =   "datConc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   405
               Left            =   600
               TabIndex        =   17
               Text            =   "Concours agility"
               Top             =   720
               Width           =   7095
            End
            Begin VB.TextBox sousTitre 
               BackColor       =   &H00FFFFFF&
               DataField       =   "Soustitre"
               DataSource      =   "datConc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00E03080&
               Height          =   285
               Left            =   0
               TabIndex        =   16
               Text            =   "Sous-titre"
               Top             =   1200
               Width           =   7695
            End
            Begin VB.TextBox txtBleu 
               BackColor       =   &H00FFFFFF&
               DataField       =   "l2b"
               DataSource      =   "datConc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF2E55&
               Height          =   285
               Index           =   2
               Left            =   600
               TabIndex        =   15
               Text            =   "Texte"
               Top             =   2040
               Width           =   3615
            End
            Begin VB.TextBox txtViolet 
               BackColor       =   &H00FFFFFF&
               DataField       =   "l3v"
               DataSource      =   "datConc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00E03080&
               Height          =   315
               Index           =   3
               Left            =   4320
               TabIndex        =   14
               Text            =   "Texte"
               Top             =   2400
               Width           =   3375
            End
            Begin VB.TextBox txtBleu 
               BackColor       =   &H00FFFFFF&
               DataField       =   "l3b"
               DataSource      =   "datConc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF2E55&
               Height          =   285
               Index           =   3
               Left            =   600
               TabIndex        =   13
               Text            =   "Texte"
               Top             =   2400
               Width           =   3615
            End
            Begin VB.TextBox txtViolet 
               BackColor       =   &H00FFFFFF&
               DataField       =   "l4v"
               DataSource      =   "datConc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00E03080&
               Height          =   315
               Index           =   4
               Left            =   4320
               TabIndex        =   12
               Text            =   "Texte"
               Top             =   2760
               Width           =   3375
            End
            Begin VB.TextBox txtBleu 
               BackColor       =   &H00FFFFFF&
               DataField       =   "l4b"
               DataSource      =   "datConc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF2E55&
               Height          =   285
               Index           =   4
               Left            =   600
               TabIndex        =   11
               Text            =   "Texte"
               Top             =   2760
               Width           =   3615
            End
            Begin VB.TextBox txtBleu 
               BackColor       =   &H00FFFFFF&
               DataField       =   "l5b"
               DataSource      =   "datConc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF2E55&
               Height          =   285
               Index           =   5
               Left            =   600
               TabIndex        =   10
               Text            =   "Texte"
               Top             =   3120
               Width           =   3615
            End
            Begin VB.TextBox txtViolet 
               BackColor       =   &H00FFFFFF&
               DataField       =   "l5v"
               DataSource      =   "datConc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00E03080&
               Height          =   315
               Index           =   5
               Left            =   4320
               TabIndex        =   9
               Text            =   "Texte"
               Top             =   3120
               Width           =   3375
            End
            Begin VB.Label lbltot 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D5F1FF&
               Caption         =   "/ 0"
               Height          =   255
               Left            =   7680
               TabIndex        =   23
               Top             =   410
               Width           =   375
            End
            Begin VB.Image btPev 
               Height          =   255
               Left            =   7800
               Picture         =   "siteConcTitre.frx":1188
               Stretch         =   -1  'True
               Top             =   720
               Width           =   255
            End
            Begin VB.Image btNext 
               Height          =   255
               Left            =   7800
               Picture         =   "siteConcTitre.frx":1470
               Stretch         =   -1  'True
               Top             =   1080
               Width           =   255
            End
            Begin VB.Image Image1 
               Height          =   450
               Left            =   0
               Picture         =   "siteConcTitre.frx":1758
               Top             =   720
               Width           =   600
            End
         End
         Begin VB.Image Image2 
            Height          =   435
            Index           =   1
            Left            =   2280
            Picture         =   "siteConcTitre.frx":1C68
            Top             =   5040
            Width           =   600
         End
         Begin VB.Label Label1 
            BackColor       =   &H00D5F1FF&
            Caption         =   "Résultats des sociétaires "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   495
            Index           =   2
            Left            =   2880
            TabIndex        =   5
            Top             =   5040
            Width           =   4215
         End
         Begin VB.Shape shapBordAdh 
            BackColor       =   &H00E2F5FF&
            BorderColor     =   &H00AAD69B&
            FillColor       =   &H00AAD69B&
            Height          =   5895
            Index           =   0
            Left            =   0
            Top             =   0
            Width           =   8775
         End
         Begin VB.Label Label1 
            BackColor       =   &H00D5F1FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Concours"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   855
            Index           =   1
            Left            =   7080
            TabIndex        =   4
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00D5F1FF&
            Caption         =   "Manifestations à venir"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   495
            Index           =   0
            Left            =   2760
            TabIndex        =   3
            Top             =   720
            Width           =   4215
         End
         Begin VB.Image Image2 
            Height          =   435
            Index           =   0
            Left            =   2160
            Picture         =   "siteConcTitre.frx":2170
            Top             =   720
            Width           =   600
         End
         Begin VB.Image Image3 
            Height          =   450
            Index           =   1
            Left            =   5040
            Picture         =   "siteConcTitre.frx":2678
            Top             =   120
            Width           =   4410
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2F5FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   6240
         Width           =   8655
         Begin VB.Image btPublier 
            Height          =   375
            Left            =   4680
            Picture         =   "siteConcTitre.frx":4990
            Top             =   360
            Width           =   1455
         End
         Begin VB.Image btFermer 
            Height          =   375
            Left            =   3000
            Picture         =   "siteConcTitre.frx":53EC
            Top             =   360
            Width           =   1455
         End
         Begin VB.Image Image4 
            Height          =   795
            Left            =   1920
            Picture         =   "siteConcTitre.frx":5E48
            Top             =   120
            Width           =   4470
         End
      End
      Begin VB.Shape shapBordAdh 
         BackColor       =   &H00E2F5FF&
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         Height          =   7095
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   9015
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00AAD69B&
      FillColor       =   &H00AAD69B&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   1560
      Top             =   6360
      Width           =   9015
   End
   Begin VB.Image btPublierDown 
      Height          =   375
      Left            =   0
      Picture         =   "siteConcTitre.frx":9CCC
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image btPublierUp 
      Height          =   375
      Left            =   120
      Picture         =   "siteConcTitre.frx":A784
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image btSuppDown 
      Height          =   375
      Left            =   0
      Picture         =   "siteConcTitre.frx":B1E0
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image btSuppUp 
      Height          =   375
      Left            =   0
      Picture         =   "siteConcTitre.frx":BB0C
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image btAddDown 
      Height          =   375
      Left            =   0
      Picture         =   "siteConcTitre.frx":C3D0
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image btAddUp 
      Height          =   375
      Left            =   0
      Picture         =   "siteConcTitre.frx":CCFC
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "siteConc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SHAPE {
'    select id,txt from siteConcTitres Order by id
'   } AS ParentCMD
'APPEND (
'   {select id,idConc,txtBleu,txtViolet from siteConcTxt Order by id }
'   AS ChildCMD RELATE id TO idConc)
'   AS ChildCMD

Private Sub ajout_Click()
  On Error GoTo AddErr
    
    datConc.Recordset.UpdateBatch adAffectAll
    datConc.Recordset.AddNew
    If (datConc.Recordset.RecordCount = 1) Then
        fManif.Visible = True
        suppr.Visible = True
        
    End If
    idConc = datConc.Recordset.RecordCount
    lbltot.Caption = "/ " + CStr(datConc.Recordset.RecordCount)
    Titre.SetFocus
  Exit Sub
AddErr:
  MsgBox err.Description
End Sub
Private Sub ajout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ajout.Picture = btAddDown.Picture
End Sub

Private Sub ajout_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ajout.Picture = btAddUp.Picture
End Sub
Private Sub cmdOpen_Click()


End Sub

Private Sub btFermer_Click()
    Unload Me
End Sub

Private Function htmlConc() As String
Dim lst As New fListes

On Error GoTo OpenError
    
    Open cheminSite + "concours.htm" For Output As #1
    
    htmlConc = ""
    
    'ENTETE
    temp$ = ""
    Open cheminSite + "concours_entete.txt" For Input As #11
    While Not EOF(11)
        Line Input #11, temp$
        Print #1, temp$
    Wend
    Close #11
    
    'MANIF
    If datConc.Recordset.RecordCount > 0 Then
        
        'entete
        temp$ = ""
        Open cheminSite + "concours_manif.txt" For Input As #11
        While Not EOF(11)
            Line Input #11, temp$
            Print #1, temp$
        Wend
        Close #11
        
        datConc.Recordset.MoveFirst
        While Not datConc.Recordset.EOF
            htmlConc = ""
            htmlConc = htmlConc + "         <BR><IMG SRC=""images/os.gif""  BORDER=0 HEIGHT=30 WIDTH=40 ALIGN=CENTER>" + Chr$(13)
            htmlConc = htmlConc + "         <B><FONT FACE=""arial, helvetica"" COLOR=""#CC3030"" SIZE=3>" + Chr$(13)
            htmlConc = htmlConc + Titre
            htmlConc = htmlConc + "         <br><FONT COLOR=""#8030E0"" SIZE=2>" + Chr$(13)
            htmlConc = htmlConc + sousTitre
            htmlConc = htmlConc + "         </FONT></FONT></B><BR>" + Chr$(13)
            htmlConc = htmlConc + "         <BR>" + Chr$(13)
            htmlConc = htmlConc + "         <UL  type=""square""><table>" + Chr$(13)
    
            Print #1, htmlConc
            htmlConc = ""
    
            For i = 1 To siteConcMaxLignes
                If txtBleu(i) <> "" Or txtViolet(i) <> "" Then
            
                    htmlConc = htmlConc + "             <tr><td><FONT FACE=""arial, helvetica"" COLOR=""#552EFF"" SIZE=2><b>" + Chr$(13)
                    htmlConc = htmlConc + txtBleu(i) + "<br></td>" + Chr$(13)
                    htmlConc = htmlConc + "             <td><FONT FACE=""arial, helvetica"" COLOR=""#8030E0"" SIZE=2><b>" + Chr$(13)
                    htmlConc = htmlConc + txtViolet(i) + "<br></td></tr>" + Chr$(13)
                End If
            Next
            htmlConc = htmlConc + "         </table></UL>" + Chr$(13)
            Print #1, htmlConc
    
             datConc.Recordset.MoveNext
        Wend 'parcours Manifs
        Print #1, "    </TD>"
        Print #1, "</TR>"
        
        datConc.Recordset.MoveFirst
    
    End If ' pas de manif
    
    'RESULTAS
    If resultats = 1 Then
    
        resu = lst.lstConchtml
        Unload lst
        If resu <> "" Then
        
            'entete
            Open cheminSite + "concours_result.txt" For Input As #12
            While Not EOF(12)
                Line Input #12, temp$
                Print #1, temp$
            Wend
            Close #12
            
            ''''
            'Resultats
            '''''
            Print #1, resu

            fGbl.Enabled = False
    
            Print #1, "    </TD>"
            Print #1, "</TR>"
        End If
    End If
    
    'Pied de page
    Open cheminSite + "concours_z.txt" For Input As #13
    While Not EOF(13)
        Line Input #13, temp$
        Print #1, temp$
    Wend
    Close #13
    Close #1
    
    Exit Function
OpenError:
    MsgBox "Error " & Format$(err.Number) & _
        " opening file." & vbCrLf & _
        err.Description
End Function

Private Sub btFermer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btFermer.Picture = fGbl.btFermerDown.Picture
End Sub



Private Sub btFermer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btFermer.Picture = fGbl.btFermerUp.Picture
End Sub

Private Sub btNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    datConc.Recordset.UpdateBatch adAffectCurrent
    If datConc.Recordset.AbsolutePosition < datConc.Recordset.RecordCount And datConc.Recordset.AbsolutePosition > 0 Then
        datConc.Recordset.MoveNext
    End If
    btNext = fGbl.btScrolldownUp.Picture

    Titre.SetFocus
End Sub

Private Sub btNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btNext = fGbl.btScrolldowndown.Picture
End Sub
Private Sub btPev_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    datConc.Recordset.UpdateBatch adAffectCurrent
    If datConc.Recordset.AbsolutePosition > 1 Then
        datConc.Recordset.MovePrevious
    End If
     btPev = fGbl.btScrollUpUp.Picture

     Titre.SetFocus
End Sub

Private Sub btPev_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btPev = fGbl.btScrollUpDown.Picture
End Sub




Private Sub btPublier_Click()
    'sauve
    If datConc.Recordset.RecordCount > 0 Then
        datConc.Recordset.UpdateBatch adAffectCurrent
        datConc.Refresh
        datConc.Refresh
    End If
    
    'génération de l'HTML et publication
    htmlConc
    publier ("Concours")
    
End Sub

Private Sub btPublier_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     btPublier.Picture = btPublierDown.Picture
End Sub

Private Sub btPublier_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     btPublier.Picture = btPublierUp.Picture
End Sub

Private Sub cpt_KeyUp(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = KKenter And cpt <> idConc Then
        'controle du nouveau
        On Error GoTo typepasbon

        nv = CInt(cpt)
        If nv > datConc.Recordset.RecordCount Then nv = datConc.Recordset.RecordCount
        If nv < 1 Then nv = 1
        
        On Error GoTo err
        If nv <> CInt(idConc) Then
            'change d'ordre
            If cpt > idConc Then
                mvt = 1
            Else
                mvt = -1
            End If
            totmvt = 1
            
            idConc = 0
            datConc.Recordset.Move (mvt)
            While idConc <> nv
                Call datConc.Recordset.Update("id", CInt(idConc) + mvt * -1)
                datConc.Recordset.Move (mvt)
                totmvt = totmvt + 1
            Wend
            
            Call datConc.Recordset.Update("id", CInt(idConc) + mvt * -1)
               
            datConc.Recordset.Move (totmvt * mvt * -1)
            Call datConc.Recordset.Update("id", nv)
            datConc.Refresh
            datConc.Refresh
            datConc.Recordset.Move (nv - 1)
        End If 'change d'ordre
        cpt = idConc
    End If ' enter
    Exit Sub
    
typepasbon:
    cpt = idConc
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Initialize()
    ChDir (chemin)
    datConc.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" + nomDB + ";"
    fManif.Visible = (datConc.Recordset.RecordCount > 0)
    suppr.Visible = fManif.Visible
    lbltot.Caption = "/ " + CStr(datConc.Recordset.RecordCount)
    'datConc.Recordset.MoveFirst
    resultats = 1
    cpt = 1
'    Top = 0
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'sauve
    If datConc.Recordset.RecordCount > 0 Then
        datConc.Recordset.UpdateBatch adAffectCurrent
        datConc.Refresh
    End If
    Screen.MousePointer = vbDefault
    fGbl.Visible = True
    fMainForm.Frametop.Visible = True
    Set fMajConc = Nothing
End Sub

Private Sub datConc_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datConc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  datConc.Caption = CStr(datConc.Recordset.AbsolutePosition)
End Sub

Private Sub datConc_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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

End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datConc.Refresh
  Set grdDataGrid.DataSource = datConc.Recordset("ChildCMD").UnderlyingValue
  Exit Sub
RefreshErr:
  MsgBox err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datConc.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub
       


Private Sub idConc_Change()
    cpt = idConc
End Sub



Private Sub suppr_Click()
    On Error GoTo DeleteErr
        
      With datConc.Recordset

        'on efface
        .Delete
        .MoveNext
        lbltot.Caption = "/ " + CStr(datConc.Recordset.RecordCount)
        If .EOF Then
        
            If .RecordCount = 0 Then
                fManif.Visible = False
                suppr.Visible = False
            Else
                .MoveLast
            End If
        
        Else
            's'il n'a pas délété le dernier, faut remetre les id
            tmpid = idConc
            totmvt = 0
            While Not .EOF
                Call datConc.Recordset.Update("id", CInt(idConc) - 1)
                .MoveNext
                totmvt = totmvt + 1
            Wend
            .Move (-totmvt)
        End If


      End With
        
  Exit Sub

DeleteErr:
  MsgBox err.Description
  MsgBox datConc.Recordset.RecordCount
End Sub

Private Sub suppr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    suppr.Picture = btSuppDown.Picture
End Sub

Private Sub suppr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    suppr.Picture = btSuppUp.Picture
End Sub
