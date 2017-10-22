VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form fListes 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00AAD69B&
   Caption         =   "Liste"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   Icon            =   "Listes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   10635
   Begin MSComctlLib.Toolbar tbToolBar 
      Height          =   390
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Frame fDiscapp 
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   120
      TabIndex        =   48
      Top             =   1560
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox idDisc 
         DataField       =   "id"
         DataSource      =   "datlstDisc"
         Height          =   285
         Left            =   720
         TabIndex        =   53
         Top             =   0
         Width           =   3375
      End
      Begin VB.TextBox labelDisc 
         DataField       =   "label"
         DataSource      =   "datlstDisc"
         Height          =   285
         Left            =   720
         TabIndex        =   52
         Top             =   315
         Width           =   3375
      End
      Begin VB.TextBox idCat 
         DataField       =   "idCat"
         DataSource      =   "datlstCat"
         Height          =   285
         Left            =   2040
         TabIndex        =   51
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox idDiscCat 
         DataField       =   "idTypeConcours"
         DataSource      =   "datlstCat"
         Height          =   285
         Left            =   2040
         TabIndex        =   50
         Top             =   1635
         Width           =   3375
      End
      Begin VB.TextBox labelCat 
         DataField       =   "labelCat"
         DataSource      =   "datlstCat"
         Height          =   285
         Left            =   2040
         TabIndex        =   49
         Top             =   1965
         Width           =   3375
      End
      Begin MSAdodcLib.Adodc datlstDisc 
         Height          =   330
         Left            =   0
         Top             =   660
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   582
         ConnectMode     =   1
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
         RecordSource    =   "select id,label from TypeConcours order by id"
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
      Begin MSAdodcLib.Adodc datlstCat 
         Height          =   330
         Left            =   120
         Top             =   2280
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   582
         ConnectMode     =   1
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
         RecordSource    =   "select idCat,idTypeConcours,labelCat from TypeCategorie order by idTypeConcours,idCat"
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
   Begin VB.TextBox montantVers 
      Alignment       =   1  'Right Justify
      DataField       =   "Montant"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   4108
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRSVers"
      Height          =   285
      Left            =   840
      TabIndex        =   47
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox dateVers 
      DataField       =   "Date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.MM.yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   4108
         SubFormatType   =   3
      EndProperty
      DataSource      =   "datPrimaryRSVers"
      Height          =   285
      Left            =   1320
      TabIndex        =   46
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox payeurid 
      DataField       =   "AdhID"
      DataSource      =   "datPrimaryRSVers"
      Height          =   285
      Left            =   600
      TabIndex        =   45
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox typPay 
      DataField       =   "TypePayement"
      DataSource      =   "datPrimaryRSVers"
      Height          =   285
      Left            =   1080
      TabIndex        =   44
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSAdodcLib.Adodc datPrimaryRSVers 
      Height          =   330
      Left            =   600
      Top             =   2280
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   1
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
      RecordSource    =   "select AdhID,Date,Montant,TypePayement from Versements"
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
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00AAD69B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   10635
      TabIndex        =   42
      Top             =   7395
      Width           =   10635
      Begin VB.TextBox statusText 
         Appearance      =   0  'Flat
         BackColor       =   &H00AAD69B&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   20
         Width           =   8775
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10320
         Y1              =   0
         Y2              =   0
      End
   End
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      ConnectMode     =   1
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
      Enabled         =   0   'False
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   6360
      TabIndex        =   37
      Top             =   240
      Visible         =   0   'False
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   6360
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
      Top             =   6360
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
      Top             =   6600
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
      ConnectMode     =   1
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
      Left            =   5040
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox scra 
         DataField       =   "No_SCRA"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1080
         TabIndex        =   55
         Top             =   1560
         Width           =   1095
      End
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
         ConnectMode     =   1
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
         RecordSource    =   "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Adresse,CodePostal,Ville,Nb_Chiens,No_SCRA from Proprietaire where id <0 "
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
      Left            =   5040
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox no_cun 
         DataField       =   "no_CUN"
         DataSource      =   "datPrimaryRSChi"
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   57
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox no_cnea 
         DataField       =   "no_CNEA"
         DataSource      =   "datPrimaryRSChi"
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   56
         Top             =   1680
         Width           =   855
      End
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
         ConnectMode     =   1
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
         RecordSource    =   "select DateNaiss,id,LOF,MaitreID,Nom,Race,Sexe,Tatouage,no_CNEA,no_CUN from Chiens where id <0 order by Nom"
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
      ConnectMode     =   1
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
      ConnectMode     =   1
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
      RecordSource    =   $"Listes.frx":058A
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
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2760
      Top             =   1305
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Listes.frx":0617
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Listes.frx":0729
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Listes.frx":083B
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Listes.frx":094D
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Listes.frx":0A5F
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Listes.frx":0B71
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Listes.frx":0C83
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Listes.frx":0D95
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Listes.frx":0EA7
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Listes.frx":0FB9
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Listes.frx":10CB
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Listes.frx":11DD
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Listes.frx":12EF
            Key             =   "Align Right"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   7395
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   13044
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Listes.frx":1401
   End
End
Attribute VB_Name = "fListes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub dateNaiss_Change(index As Integer)
    ' age(index) = DateDiff("yyyy", DateValue(dateNaiss(index)), Date)
    If IsDate(dateNaiss(index)) Then
        age(index) = Year(date) - Year(DateValue(dateNaiss(index)))
        If (Month(date) < Month(DateValue(dateNaiss(index)))) Or (Month(date) = Month(DateValue(dateNaiss(index))) And Day(date) < Day(DateValue(dateNaiss(index)))) Then
            age(index) = age(index) - 1
        End If
    End If
End Sub






Private Sub Form_Initialize()
ChDir (chemin)
End Sub

Private Sub Form_Load()

    Top = 0
    typLst.AddItem (" - Choisir une liste - ")
    typLst.AddItem ("Adhérents et Chiens")
    typLst.AddItem ("SCRA")
    typLst.AddItem ("Licence CNEA")
    typLst.AddItem ("Licence CUN")
    typLst.AddItem ("Chiens LOF")
    typLst.AddItem ("Chiens par Race")
    typLst.AddItem ("Chiens par Sexe")
    typLst.AddItem ("Concours")
    typLst.AddItem ("Concours par Chien")
    typLst.AddItem ("Echéances Cotisation")
    typLst.AddItem ("Echéances Versements")
End Sub

Private Sub Form_Resize()
    On Error GoTo fin
    Width = rtfText.Width + 122
    rtfText.Height = Height - 747
fin:
End Sub
Private Sub lstConc()
    Dim labs As New frmLabels
     
    rtfText.TextRTF = ""
    Caption = "Concours"
    txt = writeEntete(Caption)

    'Noms colones
    txt = txt + "\pard \fi142\ri-572\cf4\fs22\b0 "
    txt = txt + "Date, lieu"
    txt = txt + "\line"
    txt = txt + "\tx350\tx1600\tx4000\tx5200\tx6200\tx7400\tx8300"
    txt = txt + "\fs20 "
    txt = txt + "\tab Catégorie\tab Participant\tab Appréciation\tab Points\tab Classement"
    txt = txt + "\par "
    txt = txt + "\pard\par"
    'c'est parti
    concPrec = "premierconcours"
    totconcours = 0
    totChiens = 0

    
    'recupération des catégories et Disciplines
    'datlstDisc.RecordSource = "select id,label from TypeConcours order by id"
    'datlstDisc.Refresh
    'datlstCat.RecordSource = "select idCat,idTypeConcours,labelCat from TypeCategorie order by idTypeConcours,idCat"
   ' datlstCat.Refresh
    
    'recupération des concours de l'année et de l'année précédente
     conditiondate = CStr(" Date>" + datsql(DateValue("01.01." + CStr(Year(DateValue(date)) - 1))))
  
    
    datPrimaryRSConcours.RecordSource = "select ChienID,Date,Lieu,Discipline,Categorie,Appreciation,Points,Classement from Concours where " + conditiondate + " order by Date desc, Lieu, Discipline,Categorie, Points desc, Appreciation, ChienID "
    datPrimaryRSConcours.Refresh
    
    If datPrimaryRSConcours.Recordset.RecordCount < 1 Then
        ' pas de concours,
    Else
    
        'parcours concours
        datPrimaryRSConcours.Recordset.MoveFirst
        While Not datPrimaryRSConcours.Recordset.EOF
                     
            If concPrec <> concDate + concLieu Then  'rupture sur concours
                
                If concPrec <> "premierconcours" Then
                    txt = txt + "\par "
                End If
                concPrec = concDate + concLieu
                totconcours = totconcours + 1
                
                'force rupture sur disci
                discprec = 0
                
                txt = txt + "\pard \fi142\ri-572\cf0\fs22\cf1 "
                txt = txt + "\tx350\tx1600\tx4000\tx5200\tx6400\tx7400\tx8300"
                txt = txt + "le " + concDate + "   à " + concLieu
            End If 'rupture concours
            
            'rupture discipline
            If discprec <> concDisc Then
                discprec = concDisc
                txt = txt + "\line\fs18 \b0 \tab " + labs.lblDisc(CInt(concDisc))
                txt = txt + "\line \fs20"
                'force rupture sur catégorie
                cateprec = 0
            Else 'on l'affiche pas
            End If
            
            'rupture catégorie
            If cateprec <> concCat Then
                cateprec = concCat
                txt = txt + "\tab " + labs.lblCat(CInt(concDisc), CInt(concCat))
            Else 'on l'affiche pas
                txt = txt + "\tab "
            End If
      
            'Récupération du chiens
            datPrimaryRSChi.RecordSource = "select id,MaitreID,Nom from Chiens where id=" + concidChien
            datPrimaryRSChi.Refresh
            
            If datPrimaryRSChi.Recordset.RecordCount < 1 Then
                ' pas de chien pas normal
                nomC = ""
            Else
                nomC = nomChien(0)
                'Récupération du maitre
                datPrimaryRS.RecordSource = "select Nom,Prenom from Proprietaire where id=" + Maitreid
                datPrimaryRS.Refresh
                If datPrimaryRS.Recordset.RecordCount < 1 Then
                    ' adh pas trouvé ' pas normal
                    nomA = ""
                Else
                    nomA = Left(prenomAdh, 1) + ". " + nomAdh
                End If
            End If
            
            txt = txt + "\tab " + nomC + " à " + nomA
           
            'résultat du chien
            txt = txt + "\tab " + labs.lblApp(concApp) + "\tab " + concPoints + "\tab " + concClas
            txt = txt + "\line "
            
            
            
            totChiens = totChiens + 1
            datPrimaryRSConcours.Recordset.MoveNext
            
        Wend
        txt = txt + "\par "
    End If 'pas de concours
    'Total
    txt = txt + writeTotal(CStr(totChiens) + " chiens", CStr(totconcours) + " concours")
    statusText = ""
    statusText.Refresh
    
    txt = txt + "\par "
    txt = txt + "\pard \par"
    txt = txt + "}"
       
    Unload labs
    rtfText.TextRTF = txt
    rtfText.Refresh
    
End Sub
Function lstConchtml() As String
    Dim labs As New frmLabels

    lstConc
    txt = ""


    'c'est parti
    concPrec = "premierconcours"
    totconcours = 0
    totChiens = 0

    'recupération des concours
     conditiondate = CStr(" Date>" + datsql(DateAdd("m", -6, date)))
     datPrimaryRSConcours.RecordSource = "select ChienID,Date,Lieu,Discipline,Categorie,Appreciation,Points,Classement from Concours where " + conditiondate + " order by Date desc, Lieu, Discipline,Categorie, Points desc, Appreciation, ChienID "
     datPrimaryRSConcours.Refresh
    
    If datPrimaryRSConcours.Recordset.RecordCount < 1 Then
        ' pas de concours,
    Else
    
        'parcours concours
        datPrimaryRSConcours.Recordset.MoveFirst
        While Not datPrimaryRSConcours.Recordset.EOF
                     
            If concPrec <> concDate + concLieu + concDisc Then 'rupture sur concours
                
                If concPrec <> "premierconcours" Then
                    txt = txt + "        <TR>"
                    txt = txt + "            <TD bgcolor=""#FFF1D5"" colspan=""3"" height=""30"">" + Chr$(13)
                    txt = txt + "            </TD>" + Chr$(13)
                    txt = txt + "       </TR>" + Chr$(13)
                Else
                    txt = txt + "<TABLE BORDER=0 CELLSPACING=2 CELLPADDING=3 WIDTH=""90%"" >" + Chr$(13)
                End If
                concPrec = concDate + concLieu + concDisc
                totconcours = totconcours + 1
                
                'force rupture sur disci
                discprec = 0
                txt = txt + "       <TR>" + Chr$(13)
                txt = txt + "            <TD align=""left"" bgcolor=""#9BD6AA"" colspan=""6"">" + Chr$(13)
                txt = txt + "                <FONT size=2 color=""#552EFF"" FACE=""arial, helvetica""><B>" + Chr$(13)
                txt = txt + "le " + concDate + "   à " + concLieu + Chr$(13)
                txt = txt + "                </FONT>" + Chr$(13)
                txt = txt + "            </TD>" + Chr$(13)
                txt = txt + "        </TR>" + Chr$(13)
                'labels colones
                txt = txt + "            <TR bgcolor=""#F7E5BC""><TD align=""center"" valign=""top""><FONT size=2 color=""#552EFF"" FACE=""arial, helvetica""><B>" + Chr$(13)
                txt = txt + labs.lblDisc(CInt(concDisc)) + Chr$(13)
                txt = txt + "            </FONT></TD>" + Chr$(13)
                txt = txt + "            <TD align=""center"" valign=""top""><FONT size=2 color=""#552EFF"" FACE=""arial, helvetica""><B>" + Chr$(13)
                txt = txt + "Participant" + Chr$(13)
                txt = txt + "            </FONT></TD>" + Chr$(13)
                txt = txt + "            <TD align=""center"" valign=""top""><FONT size=2 color=""#552EFF"" FACE=""arial, helvetica""><B>" + Chr$(13)
                txt = txt + "Appréciation" + Chr$(13)
                txt = txt + "            </FONT></TD>" + Chr$(13)
                txt = txt + "            <TD align=""center"" valign=""top""><FONT size=2 color=""#552EFF"" FACE=""arial, helvetica""><B>" + Chr$(13)
                txt = txt + "Points" + Chr$(13)
                txt = txt + "            </FONT></TD>" + Chr$(13)
                txt = txt + "            <TD align=""center"" valign=""top""><FONT size=2 color=""#552EFF"" FACE=""arial, helvetica""><B>" + Chr$(13)
                txt = txt + "Class." + Chr$(13)
                txt = txt + "            </FONT></TD></tr>" + Chr$(13)
            
            End If 'rupture concours
            

            
            'catégorie
            txt = txt + "            <TR bgcolor=""#F7E5BC""><TD align=""center"" valign=""top""><FONT size=2 color=""#552EFF"" FACE=""arial, helvetica""><B>" + Chr$(13)
            If cateprec <> concCat Then 'rupture
                cateprec = concCat
                txt = txt + labs.lblCat(CInt(concDisc), CInt(concCat)) + Chr$(13)
            Else 'on l'affiche pas
            End If
            txt = txt + "            </FONT></TD>" + Chr$(13)
            
            'Récupération du chiens
            datPrimaryRSChi.RecordSource = "select id,MaitreID,Nom from Chiens where id=" + concidChien
            datPrimaryRSChi.Refresh
            
            If datPrimaryRSChi.Recordset.RecordCount < 1 Then
                ' pas de chien pas normal
                nomC = ""
            Else
                nomC = nomChien(0)
                'Récupération du maitre
                datPrimaryRS.RecordSource = "select Nom,Prenom from Proprietaire where id=" + Maitreid
                datPrimaryRS.Refresh
                If datPrimaryRS.Recordset.RecordCount < 1 Then
                    ' adh pas trouvé ' pas normal
                    nomA = ""
                Else
                    nomA = Left(prenomAdh, 1) + ". " + nomAdh
                End If
            End If
            txt = txt + "            <TD align=""center"" valign=""top""><FONT size=2 color=""#552EFF"" FACE=""arial, helvetica""><B>" + Chr$(13)
            txt = txt + nomC + " à " + nomA + Chr$(13)
            txt = txt + "            </FONT></TD>" + Chr$(13)
            'résultat du chien
            txt = txt + "            <TD align=""center"" valign=""top""><FONT size=2 color=""#552EFF"" FACE=""arial, helvetica""><B>" + Chr$(13)
            txt = txt + labs.lblApp(concApp) + Chr$(13)
            txt = txt + "            </FONT></TD>" + Chr$(13)
            txt = txt + "            <TD align=""center"" valign=""top""><FONT size=2 color=""#552EFF"" FACE=""arial, helvetica""><B>" + Chr$(13)
            txt = txt + concPoints + Chr$(13)
            txt = txt + "            </FONT></TD>" + Chr$(13)
            txt = txt + "            <TD align=""center"" valign=""top""><FONT size=2 color=""#552EFF"" FACE=""arial, helvetica""><B>" + Chr$(13)
            txt = txt + concClas + Chr$(13)
            txt = txt + "            </FONT></TD></TR>" + Chr$(13)
            
            totChiens = totChiens + 1
            datPrimaryRSConcours.Recordset.MoveNext
            
        Wend
        txt = txt + "</TABLE>" + Chr$(13)
    End If 'pas de concours
     
  
    Unload labs
    lstConchtml = txt
    
    
End Function
Private Sub lstConcChien()
    Dim labs As New frmLabels
     
    
    rtfText.TextRTF = ""
    Caption = "Concours / Chiens"
    txt = writeEntete(Caption)
    'Noms colones
    txt = txt + "\pard \fi142\ri-572\cf4\fs22\b0 "
    txt = txt + "NOM du chien \b0\fs20 et du Propriétaire"
    txt = txt + "\line"
    txt = txt + "\tx350\tx1600\tx4000\tx5200\tx6200\tx7400\tx8300"
    txt = txt + "\fs20 "
    txt = txt + "\tab Date\tab Lieu\tab Discipline\tab Catégorie\tab Appréciation \tab Points \tab Class."
    txt = txt + "\par "
    
    'c'est parti
    totconcours = 0
    totChiens = 0
    totconcourschien = 0
        
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
            
            statusText = nomChien(0)
            statusText.Refresh
            
        
            'recupération des concours de l'année et de l'année précédente
            conditiondate = CStr(" Date>" + datsql(DateValue("01.01." + CStr(Year(DateValue(date)) - 1))))

            datPrimaryRSConcours.RecordSource = "select ChienID,Date,Lieu,Discipline,Categorie,Appreciation,Points,Classement from Concours where ChienID=" + idChien(0) + " and " + conditiondate + " order by Date desc,Lieu,Discipline,Categorie"
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
                txt = txt + "\pard \fi142\ri-572\cf0\fs22\b1 "
                txt = txt + nomChien(0) + " \b0\fs20 à " + NomPrenom + "\Line "

                datPrimaryRSConcours.Recordset.MoveFirst
                While Not datPrimaryRSConcours.Recordset.EOF
         
                    'affichage concour
                    disci = labs.lblDisc(concDisc)
                    cate = labs.lblCat(concDisc, concCat)
                    appr = labs.lblApp(concApp)
                
                    txt = txt + "\tx350\tx1600\tx4000\tx5200\tx6400\tx7400\tx8300"
                    txt = txt + "\cf1\fs18 "
                    txt = txt + "\tab " + concDate + "\tab " + concLieu + "\tab " + disci + "\tab " + cate + "\tab " + appr + "\tab " + concPoints + "\tab " + concClas
                    txt = txt + "\par "
       
          
                    datPrimaryRSConcours.Recordset.MoveNext
                    totconcours = totconcours + 1
                    totconcourschien = totconcourschien + 1
                Wend
                
                
                txt = txt + writeSousTot(CStr(totconcourschien) + " concours")
                totconcourschien = 0
                totChiens = totChiens + 1
            End If 'pas conconcours pour ce chien
            datPrimaryRSChi.Recordset.MoveNext
        Wend
        End If 'pas de cjien
    'Total
    txt = txt + writeTotal(CStr(totChiens) + " chiens", CStr(totconcours) + " concours")
    statusText = ""
    statusText.Refresh
    
    txt = txt + "\par "
    txt = txt + "\pard \par"
    txt = txt + "}"
       
    Unload labs
    rtfText.TextRTF = txt
    rtfText.Refresh
    
End Sub
Private Sub lstChienSex()
   rtfText.TextRTF = ""
   
    Caption = "Chiens par sexe"
    
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
    datPrimaryRSChi.RecordSource = "select DISTINCTROW  MaitreID,id,Nom,Race,Sexe,LOF,Tatouage,DateNaiss from Chiens order by Sexe,Nom,MaitreID,id"
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
                    txt = txt + "\cf2\fs18 \b0 ________________________\line \tab " + CStr(nbchiens) + " " + labelPrec + "s"
                    txt = txt + "\par \page"
                End If
                
                labelPrec = labelSex(0)
                nbchiens = 0
                sexprec = sex(0)
                labelPrec = labelSex(0)
                txt = txt + "\pard \par"
                txt = txt + "\pard \fi142\ri-572 \cf1\fs22\b1 "
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
            txt = txt + nomChien(0) + " \b0\cf1\fs20 à " + NomPrenom
            txt = txt + "\line"
            txt = txt + "\cf0\tab "
            txt = txt + age(0) + "\tab " + labelRace(0) + "\tab " + tatou(0) + "\tab " + chienlof(0)
            txt = txt + "\par"

            nbchiens = nbchiens + 1
            totChiens = totChiens + 1
            datPrimaryRSChi.Recordset.MoveNext
        Wend
        txt = txt + writeSousTot(CStr(nbchiens) + " " + labelPrec + "s")
        txt = txt + "\page "
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
    Caption = "Adhérents et Chiens"
 
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
    txt = txt + "\fs20 "
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
                    txt = txt + "\cf1\fs20\b1 "
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
    Caption = "Chiens par race"
    

    txt = writeEntete(Caption)
    'Noms colones
    txt = txt + "\pard \fi142\ri-572\cf1\fs22\b0\cf4 "
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
            
            statusText = labelracerace
            statusText.Refresh
            
            'Récupération des chiens
            datPrimaryRSChi.RecordSource = "select MaitreID,id,Nom,Race,Sexe,LOF,Tatouage,DateNaiss from Chiens where Race=" + idRace + " order by Nom"
            datPrimaryRSChi.Refresh
            If datPrimaryRSChi.Recordset.RecordCount < 1 Then
                'pas de chien pour cette race
            Else

                'Affiche race
                nbchiens = 0
                nbRaces = nbRaces + 1
                txt = txt + "\pard \par"
                txt = txt + "\pard \fi142\ri-572 \cf1\fs22\b1 "
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
                    txt = txt + nomChien(0) + " \b0\cf1\fs20 à " + NomPrenom
                    txt = txt + "\line"
                    txt = txt + "\cf0\tab "
                    txt = txt + Left(labelSex(0), 1) + "\tab " + age(0) + "\tab \tab " + tatou(0) + "\tab " + chienlof(0)
                    txt = txt + "\par"
        
                    nbchiens = nbchiens + 1
                    totChiens = totChiens + 1
                    datPrimaryRSChi.Recordset.MoveNext
                Wend
                txt = txt + writeSousTot(CStr(CStr(nbchiens) + " " + labelPrec))
            End If
            datPrimaryRSRac.Recordset.MoveNext
        Wend


    End If
    
    'Total
    txt = txt + writeTotal(CStr(nbRaces) + " races", CStr(totChiens) + " chiens")
    statusText = ""
    txt = txt + "\pard \par"
    txt = txt + "}"
    rtfText.TextRTF = txt
    rtfText.Refresh

End Sub

Sub lstSCRA()
 
    rtfText.TextRTF = ""
    Caption = "SCRA"
 
    nbadhs = 0
    txt = writeEntete(Caption)

    'Noms colones
    txt = txt + "\pard \fi142\ri-572 "
    txt = txt + "\cf4\fs22 \b0"
    
    txt = txt + "\tx1136\tx2272"
    txt = txt + "NOM Pr\'e9nom et Adresse de l'adh\'e9rent"
    txt = txt + "\fs20 "
    txt = txt + "\line "
    txt = txt + " N° SCRA"
    txt = txt + "\par "
    
    txt = txt + "\pard \par"
        
    'c'est parti
    'Récupération de tous les Adh
    datPrimaryRS.RecordSource = "select id,Nom,Prenom,Adresse,CodePostal,Ville,No_SCRA from Proprietaire where (No_SCRA is not null and No_SCRA<>"""") order by Nom"
    datPrimaryRS.Refresh
 
    If datPrimaryRS.Recordset.RecordCount > 0 Then
    
        datPrimaryRS.Recordset.MoveFirst
        While Not datPrimaryRS.Recordset.EOF
        
            txt = txt + "\pard \fi142\ri-572\f0"
            txt = txt + "\fs22 \cf0"
            txt = txt + "\b1" + NomPrenom + "\b0" + "\tab " + adr + " " + codePost + " " + ville + " "
            txt = txt + "\line \fs20 "
            txt = txt + "" + scra

            txt = txt + "\par "
            txt = txt + "\par "
               
            nbadhs = nbadhs + 1
            datPrimaryRS.Recordset.MoveNext
            
        Wend
    End If
    'Total
    txt = txt + writeTotal(CStr(nbadhs) + " adhérents", "")
    
    txt = txt + "\pard \par"
    txt = txt + "}"
    rtfText.TextRTF = txt
    rtfText.Refresh
End Sub ' scra
Sub lstCNEA()
    rtfText.TextRTF = ""
    Caption = "Licences CNEA"
 
    nbadhs = 0
    nbchiens = 0
    txt = writeEntete(Caption)

    'Noms colones
    txt = txt + "\pard \fi142\ri-572 "
    txt = txt + "\cf4\fs22 \b0"
    
    txt = txt + "\tx1136\tx2272"
    txt = txt + "NOM et Pr\'e9nom de l'adh\'e9rent"
    txt = txt + "\par "
    
    txt = txt + "\pard \fi142\ri-572 "
    txt = txt + "\tx350\tx800\tx1200\tx3580\tx5800"
    txt = txt + "\fs20 "
    txt = txt + "\tab NOM du chien \tab N° CNEA "
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
        

   
            'Récupération des chiens
            datPrimaryRSChi.RecordSource = "select MaitreID,id,Nom,Race,Sexe,LOF,Tatouage,DateNaiss,no_CNEA,no_CUN from Chiens where (MaitreID=" + adhid + "and no_CNEA is not null and no_CNEA <>"""") order by Nom,MaitreID,id"
            datPrimaryRSChi.Refresh
            
            If datPrimaryRSChi.Recordset.RecordCount > 0 Then
            
            txt = txt + "\pard \fi142\ri-572\f0\b1 "
            txt = txt + "\fs22 \cf0"
            txt = txt + NomPrenom
            txt = txt + "\b0"
            txt = txt + "\par "
            nbadhs = nbadhs + 1
            
                datPrimaryRSChi.Recordset.MoveFirst
                While Not datPrimaryRSChi.Recordset.EOF
                
                    txt = txt + "\pard \fi142\ri-572 "
                    txt = txt + "\tx350\tx600\tx900\tx3580\tx5800"
                    txt = txt + "\cf1\fs20\b1 "
                    txt = txt + "-\tab " + nomChien(0) + " \tab " + no_cnea(0)
                    txt = txt + "\line "
                    txt = txt + "\fs18 \b0"
                    txt = txt + "\tab " + Left(labelSex(0), 1) + "\tab " + age(0) + "\tab " + labelRace(0) + "\tab " + tatou(0) + "\tab " + chienlof(0)
                    
                    txt = txt + "\line"
                    nbchiens = nbchiens + 1
                    datPrimaryRSChi.Recordset.MoveNext
                Wend
                txt = txt + "\par "
            
            End If
            
            
            datPrimaryRS.Recordset.MoveNext
            
        Wend
    End If
    'Total
    txt = txt + writeTotal(CStr(nbadhs) + " adhérents", CStr(nbchiens) + " chiens")
    
    txt = txt + "\pard \par"
    txt = txt + "}"
    rtfText.TextRTF = txt
    rtfText.Refresh
End Sub ' CENA
Sub lstCUN()
  rtfText.TextRTF = ""
    Caption = "Licences CNEA"
 
    nbadhs = 0
    nbchiens = 0
    txt = writeEntete(Caption)

    'Noms colones
    txt = txt + "\pard \fi142\ri-572 "
    txt = txt + "\cf4\fs22 \b0"
    
    txt = txt + "\tx1136\tx2272"
    txt = txt + "NOM et Pr\'e9nom de l'adh\'e9rent"
    txt = txt + "\par "
    
    txt = txt + "\pard \fi142\ri-572 "
    txt = txt + "\tx350\tx800\tx1200\tx3580\tx5800"
    txt = txt + "\fs20 "
    txt = txt + "\tab NOM du chien \tab N° CUN "
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
        

   
            'Récupération des chiens
            datPrimaryRSChi.RecordSource = "select MaitreID,id,Nom,Race,Sexe,LOF,Tatouage,DateNaiss,no_CNEA,no_CUN from Chiens where (MaitreID=" + adhid + "and no_CUN is not null and no_CUN <>"""") order by Nom,MaitreID,id"
            datPrimaryRSChi.Refresh
            
            If datPrimaryRSChi.Recordset.RecordCount > 0 Then
            
            txt = txt + "\pard \fi142\ri-572\f0\b1 "
            txt = txt + "\fs22 \cf0"
            txt = txt + NomPrenom
            txt = txt + "\b0"
            txt = txt + "\par "
            nbadhs = nbadhs + 1
            
                datPrimaryRSChi.Recordset.MoveFirst
                While Not datPrimaryRSChi.Recordset.EOF
                
                    txt = txt + "\pard \fi142\ri-572 "
                    txt = txt + "\tx350\tx600\tx900\tx3580\tx5800"
                    txt = txt + "\cf1\fs20\b1 "
                    txt = txt + "-\tab " + nomChien(0) + " \tab " + no_cun(0)
                    txt = txt + "\line "
                    txt = txt + "\fs18 \b0"
                    txt = txt + "\tab " + Left(labelSex(0), 1) + "\tab " + age(0) + "\tab " + labelRace(0) + "\tab " + tatou(0) + "\tab " + chienlof(0)
                    
                    txt = txt + "\line"
                    nbchiens = nbchiens + 1
                    datPrimaryRSChi.Recordset.MoveNext
                Wend
                txt = txt + "\par "
            
            End If
            
            
            datPrimaryRS.Recordset.MoveNext
            
        Wend
    End If
    'Total
    txt = txt + writeTotal(CStr(nbadhs) + " adhérents", CStr(nbchiens) + " chiens")
    
    txt = txt + "\pard \par"
    txt = txt + "}"
    rtfText.TextRTF = txt
    rtfText.Refresh
End Sub 'CUN


Sub lstChienLOF()

    rtfText.TextRTF = ""
    Caption = "Chiens LOF"
    
    

    txt = writeEntete(Caption)
    'Noms colones
    txt = txt + "\pard \fi142\ri-572\cf1\fs22\b0\cf4 "
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
            txt = txt + nomChien(0) + " \b0\cf1\fs20 à " + NomPrenom
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
    txt = txt + writeTotal(CStr(nbchiens) + " chiens", "")
    txt = txt + "\pard \par"
    txt = txt + "}"
    rtfText.TextRTF = txt
    rtfText.Refresh

End Sub 'LOF
Sub lstEch()

    rtfText.TextRTF = ""
    Caption = "Echéances des cotisations"
 
    nbadhs = 0
    nbchiens = 0
    txt = writeEntete(Caption)

    'Noms colones
    txt = txt + "\pard \fi142\ri-572 "
    txt = txt + "\cf4\fs20 \b0"
    
    txt = txt + "Date d'Ech. NOM et Pr\'e9nom de l'adh\'e9rent"
    txt = txt + "\fs17 "
    txt = txt + "\line "
    txt = txt + "Nb.\tab Cotis.\tab T\'e9l Domicile\tab T\'e9l Portable"
    txt = txt + "\line "
    txt = txt + "Chiens\tab Annuelle"
    txt = txt + "\par "
    
    
    'c'est parti , par date d'échéance
    nbEch = 0
    moisprec = ""
    'Echéances > aujourd'hui
    conditiondate = " Echeance >= " & datsql(date)
    datPrimaryRSCoti.RecordSource = "select Echeance ,AdhID,Cotisation from EcheanceCotis where " + conditiondate + " Order by Echeance "
    datPrimaryRSCoti.Refresh

    If datPrimaryRSCoti.Recordset.RecordCount < 1 Then
        ' pas d'échéance
    Else
        
        datPrimaryRSCoti.Recordset.MoveFirst
        While Not datPrimaryRSCoti.Recordset.EOF

           'rupture sur le mois
           If (moisprec <> Month(nvCotiEch)) Then
                moisprec = Month(nvCotiEch)
                
                txt = txt + "\pard \par"
                txt = txt + "\pard \fi142\ri-572\f0\b1\cf4\fs20 "
                txt = txt + mois(Int(moisprec))
                txt = txt + " " + CStr(Year(nvCotiEch)) + " \par "
            
           End If
            'récupération de l'adh
            datPrimaryRS.RecordSource = "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Nb_Chiens from Proprietaire where id=" + debiteurid
            datPrimaryRS.Refresh
            datPrimaryRS.Refresh
            If datPrimaryRS.Recordset.RecordCount < 1 Then
                ' pas normal
                NomPrenom = ""
            Else
            End If
            txt = txt + "\pard \par"
            txt = txt + "\pard \fi142\ri-572\f0\b0 "
            txt = txt + "\fs20\cf0 "
            txt = txt + nvCotiEch
            txt = txt + "\fs18\tab \b1" + NomPrenom
            txt = txt + "\b0\fs18 \line "
            txt = txt + "\tx350\tx650\tx1400 \tx2580\b0 "
            txt = txt + adhNbChien + " ch." + "\tab " + nvCotiMnt + "\tab " + teldom + "\tab\tab " + telport
            txt = txt + "\par "
               
            nbEch = nbEch + 1
            datPrimaryRSCoti.Recordset.MoveNext
            
        Wend ' parcour echeances
    
    End If ' Ech dépassées
    
    'Total
    txt = txt + writeTotal(CStr(nbEch) + " cotisation(s) en cours", "")
   
    txt = txt + "\pard \par"
    txt = txt + "}"
    rtfText.TextRTF = txt
    rtfText.Refresh

End Sub

Sub lstVersDif()

    rtfText.TextRTF = ""
    Caption = "Echéances des versements différés"
 
    nbadhs = 0
    nbchiens = 0
    txt = writeEntete(Caption)

    'Noms colones
    txt = txt + "\pard \fi142\ri-572 "
    txt = txt + "\cf4\fs20 \b0"
    
    txt = txt + "NOM et Pr\'e9nom de l'adh\'e9rent"
    txt = txt + "\fs18 "
    txt = txt + "\line "
    txt = txt + "Date Vers.\tab Montant\tab \tab Nb.Chiens\tab T\'e9l Domicile\tab T\'e9l Portable"
    txt = txt + "\par "
    
    
    'c'est parti , par date de versement
    nbEch = 0
    moisprec = ""
    'Echéances > aujourd'hui
    conditiondate = " date >= " & datsql(date)
    datPrimaryRSVers.RecordSource = "select AdhID,date,montant from Versements where" + conditiondate + " Order by Date "
    datPrimaryRSVers.Refresh
    datPrimaryRSVers.Refresh

    If datPrimaryRSVers.Recordset.RecordCount < 1 Then
        ' pas d'échéance
    Else
        
        datPrimaryRSVers.Recordset.MoveFirst
        While Not datPrimaryRSVers.Recordset.EOF

           'rupture sur le mois
           If (moisprec <> Month(dateVers)) Then
                moisprec = Month(dateVers)
                
                txt = txt + "\pard \par"
                txt = txt + "\pard \fi142\ri-572\f0\b1\cf4\fs20 "
                txt = txt + mois(Int(moisprec))
                txt = txt + " " + CStr(Year(dateVers)) + " \par "
            
           End If
            'récupération de l'adh
            datPrimaryRS.RecordSource = "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Nb_Chiens from Proprietaire where id=" + payeurid
            datPrimaryRS.Refresh
            datPrimaryRS.Refresh
            If datPrimaryRS.Recordset.RecordCount < 1 Then
                ' pas normal
                NomPrenom = ""
            Else
            End If
            txt = txt + "\pard \par"
            txt = txt + "\pard \fi142\ri-572\f0\cf0 "
            txt = txt + "\fs18\b1" + NomPrenom
            txt = txt + "\line\fs20\b0 "
            txt = txt + dateVers + "\tab " + montantVers + " Euros \tab\fs18 "
            txt = txt + adhNbChien + " chien(s)" + "\tab " + teldom + "\tab " + telport
            txt = txt + "\par "
               
            nbEch = nbEch + 1
            datPrimaryRSVers.Recordset.MoveNext
            
        Wend ' parcour echeances
    
    End If ' Ech dépassées
    
    'Total
    txt = txt + writeTotal(CStr(nbEch) + " versement(s) en cours", "")
   
    txt = txt + "\pard \par"
    txt = txt + "}"
    rtfText.TextRTF = txt
    rtfText.Refresh
End Sub



Sub Form_Unload(Cancel As Integer)
    If nbWin <= 2 Then fGbl.Enabled = True
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
        If sex(index) = 1 Then
            labelSex(index) = "Mâle"
        ElseIf sex(index) = 0 Then
            labelSex(index) = "Femelle"
        End If
End Sub
Private Function writeSousTot(sousTot As String) As String
    txt = ""
    txt = txt + "\pard \fi142\ri-572 "
    txt = txt + "\tx350\tx600\tx900\tx3580\tx5800"
    txt = txt + "\cf2\fs18 \b0 ________________________\line \tab " + sousTot
    txt = txt + "\par "
    writeSousTot = txt
End Function

Private Function writeTotal(tot1 As String, tot2 As String) As String
    
    txt = ""
    txt = txt + "\pard \fi142\ri-572 "
    txt = txt + "\tx350\tx600\tx900\tx3580\tx5800"
    txt = txt + "\cf2\fs18 \b0 \line ________________________ \line"
    
    txt = txt + "\cf1 "
    txt = txt + tot1
    txt = txt + "\line "
    txt = txt + "\cf0\fs16\tx1136\tx2272 "
    txt = txt + "\tab " + tot2
    txt = txt + "\par "
    writeTotal = txt
End Function
Private Function writeEntete(Titre As String) As String
    
    
    txt = ""
    txt = "{\rtf1\ansi\ansicpg1252\deff0"
    txt = txt + "{\fonttbl"
    txt = txt + "{\f0\french\fprq2\fcharset0 Arial;}"
    txt = txt + "{\f1\french\fprq2\fcharset0 Times;}}"
    txt = txt + "{\colortbl ;\red0\green0\blue128;\red0\green0\blue0;\red0\green0\blue200;\red100\green100\blue200;}"
    txt = txt + "\viewkind4\uc1"
    'Titre
    txt = txt + "\pard\ri-572\sb100\sa100\qc\cf1\lang4108\b\f0\fs28 "
    txt = txt + Titre + "\cf2\fs18\line " + CStr(date)
    txt = txt + "\par "
    txt = txt + "\pard \par"
    
    writeEntete = txt
End Function



Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
    End Select
End Sub

Private Sub Text1_Change()

End Sub

Private Sub typLst_click()
    MousePointer = vbHourglass
    
    
    Select Case typLst.ListIndex
            Case 1: '
                Call lstGlb
            Case 2:
                Call lstSCRA
            Case 3:
                Call lstCNEA
            Case 4:
                Call lstCUN
            Case 5:
                Call lstChienLOF
            Case 6:
                Call lstChienRace
            Case 7:
                Call lstChienSex
            Case 8:
                Call lstConc
            Case 9:
                Call lstConcChien
            Case 10:
                Call lstEch
            Case 11:
                Call lstVersDif
        End Select
        MousePointer = vbArrow
        rtfText.SetFocus
End Sub


Private Sub mnuFileSave_Click()
    Dim sFIle As String

        With dlgCommonDialog
            .InitDir = cheminListes
            .DialogTitle = "Sauvegarde d'une liste"
            .CancelError = True
            'ToDo: set the flags and attributes of the common dialog control
            If InStr(Caption, ".rtf") = 0 Then
                sFIle = Caption + " " + Replace(CStr(date), ".", "-")
            Else
                sFIle = Caption
            End If
            
            .Filter = "Listes (*.rtf)|*.rtf"
            .Filename = sFIle
            .ShowSave
            
            If Len(.Filename) = 0 Then
                Exit Sub
            End If
            sFIle = .Filename
        End With
        rtfText.SaveFile sFIle
        ChDir (chemin)
End Sub

Private Sub mnuFileClose_Click()
    'ToDo: Add 'mnuFileClose_Click' code.
    MsgBox "Add 'mnuFileClose_Click' code."
End Sub
Private Sub mnuFilePrint_Click()
    On Error Resume Next
    
    

    With dlgCommonDialog
        .DialogTitle = "Imppression"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If rtfText.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If err <> MSComDlg.cdlCancel Then
            rtfText.SelPrint .hDC
        End If
    End With

End Sub
