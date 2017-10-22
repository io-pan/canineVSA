VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form flstGbl 
   Caption         =   "Liste"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   10320
   Begin VB.TextBox debiteurid 
      DataField       =   "AdhID"
      DataSource      =   "datPrimaryRSCoti"
      Height          =   285
      Left            =   1560
      TabIndex        =   41
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox nvCotiEch 
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
      Left            =   1320
      TabIndex        =   40
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox nvCotiMnt 
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
      Left            =   480
      TabIndex        =   39
      Top             =   4320
      Width           =   735
   End
   Begin MSAdodcLib.Adodc datPrimaryRSCoti 
      Height          =   330
      Left            =   480
      Top             =   3960
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
      RecordSource    =   "select AdhID,Cotisation,Echeance from EcheanceCotis Order by Echeance"
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
   Begin VB.ComboBox typLst 
      Height          =   315
      Left            =   7680
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   6360
      TabIndex        =   37
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox concDate 
      DataField       =   "Date"
      DataSource      =   "datPrimaryRSConcours"
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox concPoints 
      DataField       =   "Points"
      DataSource      =   "datPrimaryRSConcours"
      Height          =   285
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox concLieu 
      DataField       =   "Lieu"
      DataSource      =   "datPrimaryRSConcours"
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   6480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox concDisc 
      DataField       =   "Discipline"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   4108
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRSConcours"
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "0"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox concClas 
      Alignment       =   1  'Right Justify
      DataField       =   "Classement"
      DataSource      =   "datPrimaryRSConcours"
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox concidChien 
      DataField       =   "ChienID"
      DataSource      =   "datPrimaryRSConcours"
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox concCat 
      DataField       =   "Categorie"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   4108
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRSConcours"
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "0"
      Top             =   6480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox concApp 
      DataField       =   "Appreciation"
      DataSource      =   "datPrimaryRSConcours"
      Height          =   285
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "0"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox idRacepourchien 
      DataField       =   "ID"
      DataSource      =   "datPrimaryRSRace"
      Height          =   285
      Left            =   4200
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox labelracerace 
      DataField       =   "Label"
      DataSource      =   "datPrimaryRSRac"
      Height          =   285
      Left            =   1080
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox idRace 
      DataField       =   "ID"
      DataSource      =   "datPrimaryRSRac"
      Height          =   285
      Left            =   480
      TabIndex        =   26
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc datPrimaryRSRac 
      Height          =   330
      Left            =   120
      Top             =   360
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
      LockType        =   1
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
   Begin VB.Frame FrameAdh 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   4800
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox ville 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DataField       =   "Ville"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   24
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox adhNbChien 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         DataField       =   "Nb_Chiens"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         TabIndex        =   23
         Top             =   1560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox nomAdh 
         DataField       =   "Nom"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Text            =   "Nom"
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox codePost 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DataField       =   "CodePostal"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox adhid 
         DataField       =   "id"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   20
         Text            =   "id"
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox teldom 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Left            =   840
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox telport 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DataField       =   "Tel_Portable"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   220
         Left            =   3000
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox adr 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DataField       =   "Adresse"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox prenomAdh 
         DataField       =   "Prenom"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Text            =   "Prénom"
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox NomPrenom 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   4215
      End
      Begin MSAdodcLib.Adodc datPrimaryRS 
         Height          =   330
         Left            =   2400
         Top             =   1560
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
         BackColor       =   -2147483644
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
         RecordSource    =   "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Adresse,CodePostal,Ville,Nb_Chiens from Proprietaire where id <0 "
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
   End
   Begin VB.Frame FrameChien 
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2655
      Index           =   0
      Left            =   4800
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox labelRace 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         DataField       =   "Label"
         DataSource      =   "datPrimaryRSRace"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E03080&
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   25
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox idRaceChien 
         BackColor       =   &H8000000F&
         DataField       =   "Race"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4108
            SubFormatType   =   1
         EndProperty
         DataSource      =   "datPrimaryRSChi"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   13
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox dateNaiss 
         BackColor       =   &H8000000F&
         DataField       =   "DateNaiss"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd.MM.yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4108
            SubFormatType   =   3
         EndProperty
         DataSource      =   "datPrimaryRSChi"
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tatou 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         DataField       =   "Tatouage"
         DataSource      =   "datPrimaryRSChi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E03080&
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   11
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox chienlof 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         DataField       =   "LOF"
         DataSource      =   "datPrimaryRSChi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E03080&
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   10
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox nomChien 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2F5FF&
         BorderStyle     =   0  'None
         DataField       =   "Nom"
         DataSource      =   "datPrimaryRSChi"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H003030CC&
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   4215
      End
      Begin VB.CheckBox sex 
         Caption         =   "sex"
         DataField       =   "Sexe"
         DataSource      =   "datPrimaryRSChi"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   1800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox idChien 
         DataField       =   "id"
         DataSource      =   "datPrimaryRSChi"
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   7
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox age 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2F5FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E03080&
         Height          =   280
         Index           =   0
         Left            =   3720
         TabIndex        =   6
         Text            =   "0"
         Top             =   570
         Width           =   255
      End
      Begin VB.TextBox labelSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         DataSource      =   "datPrimaryRSChien"
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
         Height          =   280
         Index           =   0
         Left            =   3000
         TabIndex        =   5
         Top             =   570
         Width           =   735
      End
      Begin VB.TextBox Maitreid 
         BackColor       =   &H8000000F&
         DataField       =   "MaitreID"
         DataSource      =   "datPrimaryRSChi"
         Enabled         =   0   'False
         Height          =   285
         Left            =   0
         TabIndex        =   4
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox labelRace 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         DataField       =   "Label"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E03080&
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
      Begin MSAdodcLib.Adodc datPrimaryRSChi 
         Height          =   330
         Left            =   1800
         Top             =   1920
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
         LockType        =   1
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
         RecordSource    =   "select DateNaiss,id,LOF,MaitreID,Nom,Race,Sexe,Tatouage from Chiens where id <0 order by Nom"
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
   End
   Begin VB.TextBox cmd 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc datPrimaryRSRace 
      Height          =   330
      Left            =   4680
      Top             =   240
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
      LockType        =   1
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
      RecordSource    =   "select ID,Label from Races Order by ID"
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
   Begin MSAdodcLib.Adodc datPrimaryRSConcours 
      Height          =   330
      Left            =   840
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
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
      RecordSource    =   $"lstGlobale.frx":0000
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
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   7155
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   12621
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"lstGlobale.frx":008D
   End
End
Attribute VB_Name = "flstGbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub dateNaiss_Change(index As Integer)
        If IsDate(dateNaiss(index)) Then
            age(index) = Year(Date) - Year(DateValue(dateNaiss(index)))
            If (Month(Date) < Month(DateValue(dateNaiss(index)))) Or (Month(Date) = Month(DateValue(dateNaiss(index))) And Day(Date) < Day(DateValue(dateNaiss(index)))) Then
                age(index) = age(index) - 1
            End If
        End If
End Sub

Private Sub Form_Load()
    typLst.AddItem ("Générale")
    typLst.AddItem ("Chiens LOF")
    typLst.AddItem ("Chiens par race")
    typLst.AddItem ("Chiens par sexe")
    typLst.AddItem ("Concours")
    typLst.AddItem ("Echéances")
    typLst.AddItem ("Versements différéts")
End Sub

Private Sub Form_Resize()
Width = rtfText.Width + 125
rtfText.Height = Height - 475
End Sub
Private Sub lstConc()
    Dim labs As New frmLabels
     
    
    rtfText.TextRTF = ""
    Caption = "Liste des concours"
    
    txt = writeEntete(Caption)
    'Noms colones
    txt = txt + "\pard \fi142\ri-572\cf4\fs22\b0 "
    txt = txt + "NOM du chien \b0\fs20 et du Propriétaire"
    txt = txt + "\line"
    txt = txt + "\tx350\tx1600\tx4000\tx5200\tx6200\tx7400\tx8300"
    txt = txt + "\fs20 "
    txt = txt + "\tab Date\tab Lieu\tab Discipline\tab Catégorie\tab Appréciation \tab Points \tab Class."
    txt = txt + "\par "
    
    txt = txt + "\pard \par"
    
    
    'c'est parti
    totconcours = 0
    totChiens = 0
        
        
        
        
    'Récupération du chiens
    datPrimaryRSChi.RecordSource = "select MaitreID,id,Nom from Chiens order by nom,id"
    datPrimaryRSChi.Refresh
    
    If datPrimaryRSChi.Recordset.RecordCount < 1 Then
        ' pas de chien
        'nomChien(0) = ""
    Else

        'idChien(0) = 0
        datPrimaryRSChi.Recordset.MoveFirst
        While Not datPrimaryRSChi.Recordset.EOF

            'recupération des concours
            datPrimaryRSConcours.RecordSource = "select ChienID,Date,Lieu,Discipline,Categorie,Appreciation,Points,Classement from Concours where ChienID=" + idChien(0) + " order by Date desc,Lieu,Discipline,Categorie"
            datPrimaryRSConcours.Refresh
            If datPrimaryRSConcours.Recordset.RecordCount < 1 Then
                ' pas de concours, pour ce chien
            Else
         
                        'Récupération du maitre
                        datPrimaryRS.RecordSource = "select Nom,Prenom from Proprietaire where id=" + Maitreid
                        datPrimaryRS.Refresh
                        If datPrimaryRS.Recordset.RecordCount < 1 Then
                            ' adh pas trouvé
                            NomPrenom = ""
                        Else
                        End If
                       
                        ' affiche chien et adh
                        txt = txt + "\par "
                        txt = txt + "\pard \fi142\ri-572\cf3\fs24\b1 "
                        txt = txt + nomChien(0) + " \b0\cf3\fs20 à " + NomPrenom + "\Line "
        
        
                        totChiens = totChiens + 1
                    
                datPrimaryRSConcours.Recordset.MoveFirst
                While Not datPrimaryRSConcours.Recordset.EOF
         
                 
    
                    totconcours = totconcours + 1
                    'affichage concour
                    disci = labs.lblDisc(concDisc)
                    cate = labs.lblCat(concDisc, concCat)
                    appr = labs.lblApp(concApp)
                
                    txt = txt + "\tx350\tx1600\tx4000\tx5200\tx6400\tx7400\tx8300"
                    txt = txt + "\cf0\fs18 "
                    txt = txt + "\tab " + concDate + "\tab " + concLieu + "\tab " + disci + "\tab " + cate + "\tab " + appr + "\tab " + concPoints + "\tab " + concClas
                    txt = txt + "\par "
       
          
                    datPrimaryRSConcours.Recordset.MoveNext
                    totconcours = totconcours + 1
                    
                Wend
                
                totChiens = totChiens + 1
            End If 'pas conconcours pour ce chien
            datPrimaryRSChi.Recordset.MoveNext
        Wend
        End If 'pas de cjien
    'Total
    txt = txt + writeTotal(CStr(totChiens) + " chiens", CStr(totconcours) + " concours")
  
    
    txt = txt + "\par "
    txt = txt + "\pard \par"
    txt = txt + "}"
       
    Unload labs
    rtfText.TextRTF = txt
    rtfText.Refresh
    
End Sub
Private Sub lstChienSex()
   rtfText.TextRTF = ""
   
    Caption = "Liste des chiens par sexe"
    
    txt = writeEntete(Caption)
    'Noms colones
    txt = txt + "\pard \fi142\ri-572\cf4\fs22\b0 "
    txt = txt + "Sexe"
    txt = txt + "\line"
    txt = txt + "\tx350\tx800\tx1200\tx3580\tx5800"
    txt = txt + "\fs22 "
    txt = txt + "\tab NOM du chien \b0 et du Propriétaire"
    txt = txt + "\line "
    txt = txt + "\fs20"
    txt = txt + "\tab Age\tab Race\tab Tatouage\tab LOF"
    txt = txt + "\par "
    
    txt = txt + "\pard \par"
    
    
    'c'est parti
    nbchiens = 0
    totChiens = 0
    sexprec = 66
    
    'Récupération des chiens
    datPrimaryRSChi.RecordSource = "select MaitreID,id,Nom,Race,Sexe,LOF,Tatouage,DateNaiss from Chiens order by Sexe,Nom,MaitreID,id"
    datPrimaryRSChi.Refresh
    If datPrimaryRSChi.Recordset.RecordCount < 1 Then
        ' pas de chien
    Else
        datPrimaryRSChi.Recordset.MoveFirst

        labelPrec = labelSex(0)
        While Not datPrimaryRSChi.Recordset.EOF
    
            If sex(0) <> sexprec Then ' rupture
                
                If (sexprec <> 66) Then
                    txt = txt + "\pard \fi142\ri-572 "
                    txt = txt + "\tx350\tx600\tx900\tx3580\tx5800"
                    txt = txt + "\cf2\fs22 \b0 ________________________\line \tab Total " + labelPrec + " :  " + CStr(nbchiens) + " chiens"
                    txt = txt + "\par \page"
                End If
                
                labelPrec = labelSex(0)
                nbchiens = 0
                sexprec = sex(0)
                labelPrec = labelSex(0)
                txt = txt + "\pard \par"
                txt = txt + "\pard \fi142\ri-572 \cf3\fs22\b1 "
                txt = txt + labelSex(0) + "s"
                txt = txt + "\par"
            Else
                txt = txt + "\pard "
                txt = txt + "\par "
            End If
            'Récupération du maitre
            datPrimaryRS.RecordSource = "select Nom,Prenom from Proprietaire where id=" + Maitreid
            datPrimaryRS.Refresh
            
            txt = txt + "\pard \fi142\ri-572 "
            txt = txt + "\tx350\tx600\tx900\tx3580\tx5800"
            txt = txt + "\cf0\fs22\b1\tab "
            txt = txt + nomChien(0) + " \b0\cf3\fs20 à " + NomPrenom
            txt = txt + "\line"
            txt = txt + "\cf0\tab "
            txt = txt + age(0) + "\tab " + labelRace(0) + "\tab " + tatou(0) + "\tab " + chienlof(0)
            txt = txt + "\par"

            nbchiens = nbchiens + 1
            totChiens = totChiens + 1
            datPrimaryRSChi.Recordset.MoveNext
        Wend
    End If
    'Total
    txt = txt + writeTotal(CStr(totChiens) + " chiens", "")
        
    txt = txt + "\pard \par"
    txt = txt + "}"
    rtfText.TextRTF = txt
    rtfText.Refresh
    
End Sub
Sub lstGlb()
 
    rtfText.TextRTF = ""
    Caption = "Liste des Adhérents et des Chiens"
 
    nbadhs = 0
    nbchiens = 0
    txt = writeEntete(Caption)

    'Noms colones
    txt = txt + "\pard \fi142\ri-572 "
    txt = txt + "\cf4\fs22 \b0"
    
    txt = txt + "\tx1136\tx2272"
    txt = txt + "NOM et Pr\'e9nom de l'adh\'e9rent"
    txt = txt + "\fs20 "
    txt = txt + "\line "
    txt = txt + "Nb.Chiens\tab T\'e9l Domicile\tab T\'e9l Portable\tab Adresse"
    txt = txt + "\par "
    
    txt = txt + "\pard \fi142\ri-572 "
    txt = txt + "\tx350\tx800\tx1200\tx3580\tx5800"
    txt = txt + "fs20 "
    txt = txt + "\tab NOM du chien"
    txt = txt + "\line "
    txt = txt + "\fs18 "
    txt = txt + "\tab Sex\tab Age\tab Race\tab Tatouage\tab LOF"
    txt = txt + "\par "
    
    txt = txt + "\pard \par"
    
    
    'c'est parti
   
    'Récupération de tous les Adh
    datPrimaryRS.RecordSource = "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Adresse,CodePostal,Ville,Nb_Chiens from Proprietaire order by Nom"
    datPrimaryRS.Refresh
 
    If datPrimaryRS.Recordset.RecordCount > 0 Then
    
        datPrimaryRS.Recordset.MoveFirst
        While Not datPrimaryRS.Recordset.EOF
        
            txt = txt + "\pard \fi142\ri-572\f0\b1 "
            txt = txt + "\fs22 \cf0"
            txt = txt + NomPrenom
            txt = txt + "\b0"
            txt = txt + "\line \fs20 "
            txt = txt + "\tx350\tx1900 \tx3580\b1 "
            txt = txt + "" + adhNbChien + " \b0\tab " + teldom + "\tab " + telport + "\tab " + adr + " " + codePost + " " + ville
            txt = txt + "\par "
   
            'Récupération des chiens
            datPrimaryRSChi.RecordSource = "select MaitreID,id,Nom,Race,Sexe,LOF,Tatouage,DateNaiss from Chiens where MaitreID=" + adhid + " order by Nom,MaitreID,id"
            datPrimaryRSChi.Refresh
            
            If datPrimaryRSChi.Recordset.RecordCount > 0 Then
                datPrimaryRSChi.Recordset.MoveFirst
                While Not datPrimaryRSChi.Recordset.EOF
                
                    txt = txt + "\pard \fi142\ri-572 "
                    txt = txt + "\tx350\tx600\tx900\tx3580\tx5800"
                    txt = txt + "\cf3\fs20\b1 "
                    txt = txt + "-\tab " + nomChien(0)
                    txt = txt + "\line "
                    txt = txt + "\fs18 \b0"
                    txt = txt + "\tab " + Left(labelSex(0), 1) + "\tab " + age(0) + "\tab " + labelRace(0) + "\tab " + tatou(0) + "\tab " + chienlof(0)
                    
                    txt = txt + "\line"
                    nbchiens = nbchiens + 1
                    datPrimaryRSChi.Recordset.MoveNext
                Wend
                txt = txt + "\par "
            
            End If
            
            nbadhs = nbadhs + 1
            datPrimaryRS.Recordset.MoveNext
            
        Wend
    End If
    'Total
    txt = txt + writeTotal(CStr(nbadhs) + " adhérents", CStr(nbchiens) + " chiens")
    
    txt = txt + "\pard \par"
    txt = txt + "}"
    rtfText.TextRTF = txt
    rtfText.Refresh
    
End Sub 'lstGlb
Sub lstChienRace()

    rtfText.TextRTF = ""
    Caption = "Liste des chiens par race"
    

    txt = writeEntete(Caption)
    'Noms colones
    txt = txt + "\pard \fi142\ri-572\cf3\fs22\b0\cf4 "
    txt = txt + "Race"
    txt = txt + "\line"
    txt = txt + "\tx350\tx800\tx1200\tx3580\tx5800"
    txt = txt + "\fs22 "
    txt = txt + "\tab NOM du chien \b0 et du Propriétaire"
    txt = txt + "\line "
    txt = txt + "\fs20"
    txt = txt + "\tab Sex\tab Age\tab\tab Tatouage\tab LOF"
    txt = txt + "\par "
    
    txt = txt + "\pard \par"
    
    
    'c'est parti

    nbRaces = 0
    totChiens = 0

    'Récupération des Races
    datPrimaryRSRac.RecordSource = "select ID,Label from Races Order by Label"
    datPrimaryRSRac.Refresh
    If datPrimaryRSRac.Recordset.RecordCount < 1 Then
        'pas de races
    Else
            
        datPrimaryRSRac.Recordset.MoveFirst
        While Not datPrimaryRSRac.Recordset.EOF
            
            'Récupération des chiens
            datPrimaryRSChi.RecordSource = "select MaitreID,id,Nom,Race,Sexe,LOF,Tatouage,DateNaiss from Chiens where Race=" + idRace + " order by Nom,MaitreID,id"
            datPrimaryRSChi.Refresh
            If datPrimaryRSChi.Recordset.RecordCount < 1 Then
                'pas de chien pour cette race
            Else

                'Affiche race
                nbchiens = 0
                nbRaces = nbRaces + 1
                txt = txt + "\pard \par"
                txt = txt + "\pard \fi142\ri-572 \cf3\fs22\b1 "
                txt = txt + labelracerace
                txt = txt + "\par"
                        
                labelPrec = labelracerace
                datPrimaryRSChi.Recordset.MoveFirst
                While Not datPrimaryRSChi.Recordset.EOF
            
                    'Récupération du maitre
                    datPrimaryRS.RecordSource = "select Nom,Prenom from Proprietaire where id=" + Maitreid
                    datPrimaryRS.Refresh
                    
                    txt = txt + "\pard \fi142\ri-572 "
                    txt = txt + "\tx350\tx600\tx900\tx3580\tx5800"
                    txt = txt + "\cf0\fs22\b1\tab "
                    txt = txt + nomChien(0) + " \b0\cf3\fs20 à " + NomPrenom
                    txt = txt + "\line"
                    txt = txt + "\cf0\tab "
                    txt = txt + Left(labelSex(0), 1) + "\tab " + age(0) + "\tab \tab " + tatou(0) + "\tab " + chienlof(0)
                    txt = txt + "\par"
        
                    nbchiens = nbchiens + 1
                    totChiens = totChiens + 1
                    datPrimaryRSChi.Recordset.MoveNext
                Wend
                txt = txt + "\pard \fi142\ri-572 "
                txt = txt + "\tx350\tx600\tx900\tx3580\tx5800"
                txt = txt + "\cf2\fs22 \b0 ________________________\line \tab Total " + labelPrec + " :  " + CStr(nbchiens) + " chiens"
                txt = txt + "\par "
            End If
            datPrimaryRSRac.Recordset.MoveNext
        Wend


    End If
    
    'Total
    txt = txt + writeTotal(CStr(nbRaces) + " races", CStr(totChiens) + " chiens")
    
    txt = txt + "\pard \par"
    txt = txt + "}"
    rtfText.TextRTF = txt
    rtfText.Refresh

End Sub
Sub lstChienLOF()

    rtfText.TextRTF = ""
    Caption = "Liste des chiens LOF"
    
    

    txt = writeEntete(Caption)
    'Noms colones
    txt = txt + "\pard \fi142\ri-572\cf3\fs22\b0\cf4 "
    txt = txt + "\tx350\tx800\tx1200\tx3580\tx5800"
    txt = txt + "\fs22 "
    txt = txt + "NOM du chien \b0 et du Propriétaire"
    txt = txt + "\line "
    txt = txt + "\fs20"
    txt = txt + "Sex\tab Age\tab Race\tab Tatouage\tab LOF"
    txt = txt + "\par "
    
    txt = txt + "\pard \par"
    
    
    'c'est parti
    nbchiens = 0
    
    
    'Récupération des chiens
    datPrimaryRSChi.RecordSource = "select MaitreID,id,Nom,Race,Sexe,LOF,Tatouage,DateNaiss from Chiens where LOF is not null order by Nom,MaitreID,id"
    datPrimaryRSChi.Refresh
    If datPrimaryRSChi.Recordset.RecordCount < 1 Then
        'pas de chien LOF
    Else
    
        datPrimaryRSChi.Recordset.MoveFirst
        While Not datPrimaryRSChi.Recordset.EOF
    
            'Récupération du maitre
            datPrimaryRS.RecordSource = "select Nom,Prenom from Proprietaire where id=" + Maitreid
            datPrimaryRS.Refresh
            
            txt = txt + "\pard \fi142\ri-572 "
            txt = txt + "\tx350\tx600\tx900\tx3580\tx5800"
            txt = txt + "\cf0\fs22\b1 "
            txt = txt + nomChien(0) + " \b0\cf3\fs20 à " + NomPrenom
            txt = txt + "\line"
            txt = txt + "\cf0 "
            txt = txt + Left(labelSex(0), 1) + "\tab " + age(0) + "\tab " + labelRace(0) + "\tab " + tatou(0) + "\tab " + chienlof(0)
            txt = txt + "\line"
            txt = txt + "\par"
            
            nbchiens = nbchiens + 1
            datPrimaryRSChi.Recordset.MoveNext
        Wend
    End If
    'Total
    txt = txt + writeTotal(CStr(totChiens) + " chiens", "")
    txt = txt + "\pard \par"
    txt = txt + "}"
    rtfText.TextRTF = txt
    rtfText.Refresh

End Sub
Sub lstEch()

    rtfText.TextRTF = ""
    Caption = "Cotisations arrivées à échéance"
 
    nbadhs = 0
    nbchiens = 0
    txt = writeEntete(Caption)

    'Noms colones
    txt = txt + "\pard \fi142\ri-572 "
    txt = txt + "\cf4\fs22 \b0"
    
    txt = txt + "\tx1136\tx2272"
    txt = txt + "NOM et Pr\'e9nom de l'adh\'e9rent"
    txt = txt + "\fs20 "
    txt = txt + "\line "
    txt = txt + "Nb.Chiens\tab T\'e9l Domicile\tab T\'e9l Portable\tab Adresse"
    txt = txt + "\par "
    
    'txt = txt + "\pard \fi142\ri-572 "
    'txt = txt + "\tx350\tx800\tx1200\tx3580\tx5800"
    'txt = txt + "fs20 "
    'txt = txt + "\tab NOM du chien"
    'txt = txt + "\line "
    'txt = txt + "\fs18 "
    'txt = txt + "\tab Sex\tab Age\tab Race\tab Tatouage\tab LOF"
    'txt = txt + "\par "
    
    txt = txt + "\pard \par"
    
    
    'c'est parti
    nbEch = 0
    
    conditiondate = " Echeance <=" & datsql(Date)
    datPrimaryRSCoti.RecordSource = "select Echeance ,AdhID,Cotisation from EcheanceCotis where " + conditiondate + " Order by Echeance "
    datPrimaryRSCoti.Refresh
    

    If datPrimaryRSCoti.Recordset.RecordCount < 1 Then
        ' pas d'échéance dépassée
    Else
        
        datPrimaryRSCoti.Recordset.MoveFirst
        While Not datPrimaryRSCoti.Recordset.EOF
   
            txt = txt + "\pard \fi142\ri-572\f0\b0 "
            txt = txt + "\fs22 \cf0 "
            txt = txt + nvCotiEch


           
            'récupération de l'adh
            datPrimaryRS.RecordSource = "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Nb_Chiens from Proprietaire where id=" + debiteurid
            datPrimaryRS.Refresh
            If datPrimaryRS.Recordset.RecordCount < 1 Then
                ' pas normal
            Else

                txt = txt + "\tab \b1" + NomPrenom
                txt = txt + "\b0\line \fs20 "
                txt = txt + "\tx350\tx1900 \tx3580\b1 "
                txt = txt + "" + adhNbChien + " \b0\tab " + teldom + "\tab " + telport + "\tab " + adr + " " + codePost + " " + ville
                txt = txt + "\par "
            
            'du chiens
            'datPrimaryRSChi.RecordSource = "select MaitreID,id,Nom,Race,Sexe,LOF,Tatouage,DateNaiss from Chiens where MaitreID=" + adhid + " order by Nom"
            'datPrimaryRSChi.Refresh
    
            'If datPrimaryRSChi.Recordset.RecordCount > 0 Then
            '    datPrimaryRSChi.Recordset.MoveFirst
            '    While Not datPrimaryRSChi.Recordset.EOF
            '
            '        txt = txt + "\pard \fi142\ri-572 "
            '        txt = txt + "\tx350\tx600\tx900\tx3580\tx5800"
            '        txt = txt + "\cf3\fs20\b1 "
            '        txt = txt + "-\tab " + nomChien(0)
            '        txt = txt + "\line "
            '        txt = txt + "\fs18 \b0"
            '        txt = txt + "\tab " + Left(labelSex(0), 1) + "\tab " + age(0) + "\tab " + labelRace(0) + "\tab " + tatou(0) + "\tab " + chienlof(0)
                    
            '        txt = txt + "\line"
            '        nbchiens = nbchiens + 1
             '       datPrimaryRSChi.Recordset.MoveNext
             '   Wend
            End If
            
            nbEch = nbEch + 1
            datPrimaryRSCoti.Recordset.MoveNext
            
        Wend ' parcour echeances
    
    End If ' Ech dépassées
    
    'Total
    txt = txt + writeTotal(CStr(nbEch) + " cosisation(s) échue(s)", CStr(nbchiens) + " chiens")
   
    txt = txt + "\pard \par"
    txt = txt + "}"
    rtfText.TextRTF = txt
    rtfText.Refresh

End Sub



Private Sub idRaceChien_Change(index As Integer)
        If idRaceChien(index) <> "" Then
        datPrimaryRSRace.Recordset.MoveFirst
        While (idRacepourchien <> idRaceChien(index) And datPrimaryRSRace.Recordset.AbsolutePosition < datPrimaryRSRace.Recordset.RecordCount)
            datPrimaryRSRace.Recordset.MoveNext
        Wend
        End If
End Sub

Private Sub nomAdh_Change()
 NomPrenom = UCase(nomAdh) + " " + prenomAdh
End Sub

Private Sub prenomAdh_Change()
 NomPrenom = UCase(nomAdh) + " " + prenomAdh
End Sub

Private Sub sex_Click(index As Integer)
        If sex(0) = 0 Then
            labelSex(0) = "Mâle"
        ElseIf sex(0) = 1 Then
            labelSex(0) = "Femelle"
        End If
End Sub


Private Function writeTotal(tot1 As String, tot2 As String) As String
    
    txt = ""
    txt = txt + "\pard \fi142\ri-572 "
    txt = txt + "\tx350\tx600\tx900\tx3580\tx5800"
    txt = txt + "\cf2\fs22 \b0 \line ________________________ \line"
    
    txt = txt + "\cf3 "
    txt = txt + tot1
    txt = txt + "\line "
    txt = txt + "\cf0\fs20\tx1136\tx2272 "
    txt = txt + "\tab " + tot2
    txt = txt + "\par "
    writeTotal = txt
End Function
Private Function writeEntete(titre As String) As String
    
    txt = ""
    txt = "{\rtf1\ansi\ansicpg1252\deff0"
    txt = txt + "{\fonttbl"
    txt = txt + "{\f0\french\fprq2\fcharset0 Arial;}"
    txt = txt + "{\f1\french\fprq2\fcharset0 Times;}}"
    txt = txt + "{\colortbl ;\red0\green0\blue128;\red0\green0\blue0;\red0\green0\blue200;\red100\green100\blue200;}"
    txt = txt + "\viewkind4\uc1"
    'Titre
    txt = txt + "\pard\ri-572\sb100\sa100\qc\cf1\lang4108\b\f0\fs28 "
    txt = txt + titre + "\cf2\fs18\line " + CStr(Date)
    txt = txt + "\par "
    txt = txt + "\pard \par"
    
    writeEntete = txt
End Function



Private Sub typLst_Click()
    MousePointer = vbHourglass
    Select Case typLst.ListIndex
            Case 0: '
                Call lstGlb
            Case 1:
                Call lstChienLOF
            Case 2:
                Call lstChienRace
            Case 3:
                Call lstChienSex
            Case 4:
                Call lstConc
            Case 5:
                Call lstEch
            Case 6:
               ' Call lstVersDif
        End Select
        MousePointer = vbArrow
        rtfText.SetFocus
        
End Sub
