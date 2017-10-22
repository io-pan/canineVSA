VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form addAdh 
   BackColor       =   &H00AAD69B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajout d'un adhérent"
   ClientHeight    =   7500
   ClientLeft      =   3750
   ClientTop       =   1260
   ClientWidth     =   4560
   Icon            =   "addAdh.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
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
      Height          =   6795
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   4335
      Begin VB.TextBox cun 
         DataField       =   "no_CUN"
         DataSource      =   "datPrimaryRSChien"
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   5160
         Width           =   2895
      End
      Begin VB.TextBox cnea 
         DataField       =   "no_CNEA"
         DataSource      =   "datPrimaryRSChien"
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   4800
         Width           =   2895
      End
      Begin VB.ComboBox drpSex 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox cpFermerchien 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5F1FF&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   5640
         Width           =   150
      End
      Begin VB.TextBox idRaceChien 
         BackColor       =   &H8000000F&
         DataField       =   "Race"
         DataSource      =   "datPrimaryRSChien"
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   41
         Top             =   2760
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox editLabelRace 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox idRace 
         DataField       =   "ID"
         DataSource      =   "datPrimaryRSRace"
         Height          =   285
         Left            =   360
         TabIndex        =   40
         Top             =   3240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox labelRace 
         DataField       =   "Label"
         DataSource      =   "datPrimaryRSRace"
         Height          =   285
         Left            =   240
         TabIndex        =   39
         Top             =   3600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ListBox lstRaces 
         Height          =   1425
         Left            =   1320
         TabIndex        =   38
         Top             =   3240
         Width           =   2895
      End
      Begin VB.CheckBox sex 
         Caption         =   "sex"
         DataField       =   "Sexe"
         DataSource      =   "datPrimaryRSChien"
         Height          =   255
         Left            =   3600
         TabIndex        =   28
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox dateNaiss 
         DataField       =   "DateNaiss"
         DataSource      =   "datPrimaryRSChien"
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox tatou 
         DataField       =   "Tatouage"
         DataSource      =   "datPrimaryRSChien"
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox chienlof 
         DataField       =   "LOF"
         DataSource      =   "datPrimaryRSChien"
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox nomChien 
         DataField       =   "Nom"
         DataSource      =   "datPrimaryRSChien"
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtFields 
         DataField       =   "id"
         DataSource      =   "datPrimaryRSChien"
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   0
         TabIndex        =   11
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Maitreid 
         BackColor       =   &H8000000F&
         DataField       =   "MaitreID"
         DataSource      =   "datPrimaryRSChien"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSAdodcLib.Adodc datPrimaryRSChien 
         Height          =   330
         Left            =   2880
         Top             =   2040
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
         RecordSource    =   "select MaitreID,id,Nom,Race,Sexe,LOF,Tatouage,DateNaiss,no_CNEA,no_CUN from Chiens Order by id"
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
         Left            =   240
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
         Caption         =   "Licence CUN:"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   15
         Left            =   160
         TabIndex        =   64
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Licence CNEA:"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   7
         Left            =   160
         TabIndex        =   63
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Image btFermerChien 
         Height          =   375
         Left            =   2760
         Picture         =   "addAdh.frx":058A
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Image btAnnulerChien 
         Height          =   375
         Left            =   1320
         Picture         =   "addAdh.frx":0FE6
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Image btNvChienChien 
         Height          =   375
         Left            =   2760
         Picture         =   "addAdh.frx":1ABE
         Top             =   5760
         Width           =   1455
      End
      Begin VB.Image Image3 
         Height          =   795
         Index           =   1
         Left            =   -240
         Picture         =   "addAdh.frx":251A
         Top             =   6000
         Width           =   4470
      End
      Begin VB.Shape shapBordAdh 
         BackColor       =   &H00E2F5FF&
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         Height          =   5535
         Left            =   0
         Top             =   0
         Width           =   4335
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00D5F1FF&
         FillColor       =   &H00D5F1FF&
         FillStyle       =   0  'Solid
         Height          =   1260
         Left            =   0
         Top             =   5520
         Width           =   4575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Nouveau chien"
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
         TabIndex        =   59
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image btNvRace 
         Height          =   375
         Left            =   1320
         Picture         =   "addAdh.frx":639E
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Naiss :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   27
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Tatouage :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   26
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "LOF :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   25
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Sexe :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   24
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Race :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   13
         Left            =   360
         TabIndex        =   23
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Nom :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   14
         Left            =   360
         TabIndex        =   22
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Mid:"
         Height          =   255
         Index           =   16
         Left            =   3120
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
   End
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
      Height          =   6375
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtFields 
         DataField       =   "no_scra"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   8
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox cpFermer 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2F5FF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   6960
         Width           =   150
      End
      Begin VB.Frame FrameCoti 
         BackColor       =   &H00E2F5FF&
         BorderStyle     =   0  'None
         Caption         =   "Cotisation"
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
         Height          =   2385
         Left            =   240
         TabIndex        =   42
         Top             =   3840
         Width           =   3975
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
            Left            =   2400
            TabIndex        =   56
            Text            =   "04.04.2002"
            Top             =   1680
            Width           =   975
         End
         Begin VB.ComboBox drpTypPay 
            Height          =   315
            ItemData        =   "addAdh.frx":77FA
            Left            =   1080
            List            =   "addAdh.frx":77FC
            TabIndex        =   54
            Text            =   "drpTypPay"
            Top             =   1680
            Width           =   1380
         End
         Begin VB.TextBox typPay 
            DataField       =   "TypePayement"
            DataSource      =   "datPrimaryRSVers"
            Height          =   285
            Left            =   360
            TabIndex        =   53
            Top             =   2040
            Visible         =   0   'False
            Width           =   495
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
            Left            =   360
            TabIndex        =   52
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox payeurid 
            DataField       =   "AdhID"
            DataSource      =   "datPrimaryRSVers"
            Height          =   285
            Left            =   2280
            TabIndex        =   51
            Top             =   2040
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtFields 
            DataField       =   "id"
            DataSource      =   "datPrimaryRSTypPay"
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   50
            Top             =   2520
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox labelTypPay 
            DataField       =   "Label"
            DataSource      =   "datPrimaryRSTypPay"
            Height          =   285
            Left            =   1440
            TabIndex        =   49
            Top             =   2880
            Visible         =   0   'False
            Width           =   1335
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
            TabIndex        =   45
            Top             =   840
            Width           =   975
         End
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
            Left            =   360
            TabIndex        =   44
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox debiteurid 
            DataField       =   "AdhID"
            DataSource      =   "datPrimaryRSCoti"
            Height          =   285
            Left            =   2280
            TabIndex        =   43
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSAdodcLib.Adodc datPrimaryRSCoti 
            Height          =   330
            Left            =   1440
            Top             =   120
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
         Begin MSAdodcLib.Adodc datPrimaryRSVers 
            Height          =   330
            Left            =   2520
            Top             =   2040
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
            RecordSource    =   "select AdhID,Date,Montant,TypePayement from Versements where AdhID <0"
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
         Begin MSAdodcLib.Adodc datPrimaryRSTypPay 
            Height          =   330
            Left            =   360
            Top             =   2880
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
            RecordSource    =   "select id,Label from TypeVersement Order by id"
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
         Begin VB.Shape Shape4 
            BackColor       =   &H00E2F5FF&
            BorderColor     =   &H00AAD69B&
            FillColor       =   &H00AAD69B&
            Height          =   2295
            Left            =   0
            Top             =   0
            Width           =   3975
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E2F5FF&
            Caption         =   "Cotisation"
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
            Index           =   2
            Left            =   120
            TabIndex        =   61
            Top             =   120
            Width           =   2295
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "Versement :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF2E55&
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   57
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "Date :"
            ForeColor       =   &H00FF2E55&
            Height          =   255
            Index           =   22
            Left            =   2400
            TabIndex        =   55
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "Echéance :"
            ForeColor       =   &H00FF2E55&
            Height          =   255
            Index           =   18
            Left            =   2400
            TabIndex        =   48
            Top             =   600
            Width           =   975
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
            Left            =   1080
            TabIndex        =   46
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "Montant :"
            ForeColor       =   &H00FF2E55&
            Height          =   255
            Index           =   19
            Left            =   360
            TabIndex        =   47
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Ville"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   6
         Top             =   2760
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
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox adhid 
         DataField       =   "id"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         TabIndex        =   30
         Top             =   0
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
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox adhNbChien 
         BackColor       =   &H80000014&
         DataField       =   "Nb_Chiens"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3240
         Width           =   615
      End
      Begin MSAdodcLib.Adodc datPrimaryRS 
         Height          =   330
         Left            =   1920
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
         RecordSource    =   "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Adresse,CodePostal,Ville,Nb_Chiens,no_scra from Proprietaire Order by id"
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
         TabIndex        =   62
         Top             =   1320
         Width           =   855
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E2F5FF&
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         Height          =   6375
         Left            =   0
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Nouvel adhérent"
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
         TabIndex        =   58
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image btNvChien 
         Height          =   375
         Left            =   2640
         Picture         =   "addAdh.frx":77FE
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Téléphone :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   27
         Left            =   360
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Nb Chiens :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   31
         Top             =   3240
         Width           =   855
      End
   End
   Begin VB.Image btFermer 
      Height          =   375
      Left            =   3000
      Picture         =   "addAdh.frx":80C2
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Image btAnnulerAdh 
      Height          =   375
      Left            =   1560
      Picture         =   "addAdh.frx":8B1E
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image btsavenvChiendown 
      Height          =   375
      Left            =   0
      Picture         =   "addAdh.frx":95F6
      Top             =   5520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image btsavenvChienup 
      Height          =   375
      Left            =   0
      Picture         =   "addAdh.frx":A0AE
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image btNvRaceDown 
      Height          =   375
      Left            =   240
      Picture         =   "addAdh.frx":AB0A
      Top             =   1920
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image btNvRaceUp 
      Height          =   375
      Left            =   240
      Picture         =   "addAdh.frx":BFBA
      Top             =   2400
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image btNvChienUp 
      Height          =   375
      Left            =   0
      Picture         =   "addAdh.frx":D416
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image btNvChienDown 
      Height          =   375
      Left            =   0
      Picture         =   "addAdh.frx":DCDA
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   795
      Index           =   0
      Left            =   -240
      Picture         =   "addAdh.frx":E5FA
      Top             =   6600
      Width           =   4470
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00D5F1FF&
      FillColor       =   &H00D5F1FF&
      FillStyle       =   0  'Solid
      Height          =   7395
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "addAdh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
Private Sub InitDropTypPay()
   
    Call drpTypPay.Clear
    datPrimaryRSTypPay.Recordset.MoveFirst
    While (Not (datPrimaryRSTypPay.Recordset.EOF))
        drpTypPay.AddItem (labelTypPay)
        datPrimaryRSTypPay.Recordset.MoveNext
    Wend
    drpTypPay.ListIndex = -1
    
End Sub

Private Sub adhNbChien_GotFocus()
    Call btNvChien_Click
End Sub

Private Sub btAnnulerAdh_Click()
   
    'suppression de l'adh
    If adhid <> "" Then

    
        'suppression des chiens
        datPrimaryRSChien.RecordSource = "select DateNaiss,id,LOF,MaitreID,Nom,Race,Sexe,Tatouage,no_CNEA,no_CUN from Chiens where MaitreID=" + adhid
        datPrimaryRSChien.Refresh
        With datPrimaryRSChien.Recordset
            While .RecordCount > 0
                .MoveFirst
                .Delete
            Wend
        End With
        datPrimaryRSChien.Refresh
        datPrimaryRSChien.Refresh
       
        'suppression de la coti
        datPrimaryRSCoti.RecordSource = "select AdhID,Cotisation,Echeance from EcheanceCotis  where AdhID=" + adhid
        datPrimaryRSCoti.Refresh
        If datPrimaryRSCoti.Recordset.RecordCount > 0 Then
            datPrimaryRSCoti.Recordset.Delete
            datPrimaryRSCoti.Refresh
            datPrimaryRSCoti.Refresh
        End If
        
        'suppression du versement
        datPrimaryRSVers.RecordSource = "select AdhID,Date,Montant,TypePayement from Versements where AdhID=" + adhid
        datPrimaryRSVers.Refresh
        If datPrimaryRSVers.Recordset.RecordCount > 0 Then
            datPrimaryRSVers.Recordset.Delete
            datPrimaryRSVers.Refresh
            datPrimaryRSVers.Refresh
        End If
        'supression adh
        datPrimaryRS.RecordSource = "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Adresse,CodePostal,Ville,Nb_Chiens,No_SCRA from Proprietaire where id=" + adhid
        datPrimaryRS.Refresh
        If (datPrimaryRS.Recordset.RecordCount > 0) Then
            datPrimaryRS.Recordset.Delete
            datPrimaryRS.Refresh
            datPrimaryRS.Refresh
        End If
    Else
        datPrimaryRSChien.Refresh
        datPrimaryRS.Refresh
    End If

    Unload Me
End Sub

Private Sub btAnnulerChien_Click()
    datPrimaryRSChien.Refresh
    frameChien.Visible = False
    FrameAdh.Enabled = True
    montantCoti.SetFocus
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
   
    'refrech liste au cas ou le 2è chien(s'il y a) soit de la même race
    Call InitDropRaces
    'refresh sur la form principale
    While fGbl.datPrimaryRSRace.Recordset.RecordCount <> datPrimaryRSRace.Recordset.RecordCount
        fGbl.datPrimaryRSRace.Refresh
    Wend
    Call editLabelRace_KeyUp(0, 0)
    
  Exit Sub
AddErr:
    MsgBox "L'enregistrement n'as pas eu lieu. Vérifiez que la race n'exite pas déjà."
        datPrimaryRSRace.Refresh
End Sub


'NOUVEAU CHIEN
Private Sub btNvChien_Click()
On Error GoTo AddErr
  
    If saveProprio() Then
        'nouveau chien
        frameChien.Visible = True
        frameChien.Enabled = True
        FrameAdh.Enabled = False
        
        datPrimaryRSChien.Recordset.AddNew
        'init chien
        Maitreid = adhid
        editLabelRace = ""
        sexM = False
        sexF = False
        nomChien.SetFocus
    End If
  Exit Sub
AddErr:
  MsgBox err.Description
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





Private Sub btNvVers_Click()
  On Error GoTo UpdateErr
  

      
  Exit Sub
UpdateErr:
  MsgBox err.Description
  
End Sub


Private Function saveCoti() As Boolean

    On Error GoTo UpdateErr
    saveCoti = True
    
   'Montant Cotisation
    montantCoti.SetFocus
    If (StrComp(montantCoti, "") = 0) Then
        montantCoti.BackColor = &HF9D9D4
        err.Description = "Le montant n'est pas valide"
        GoTo UpdateErr
    Else
        montantCoti.BackColor = &HFFFFFF
        On Error GoTo typmismatch
        If (FormatCurrency(montantCoti)) Then
        End If
        On Error GoTo UpdateErr
    End If
    
    'Date d'échéance
    echeance.SetFocus
    If Not IsDate(echeance) Then
            echeance.BackColor = &HF9D9D4
            err.Description = "La date d'échéance n'est pas une date valide"
            GoTo UpdateErr
    Else
        echeance.BackColor = &HFFFFFF
        If (DateValue(echeance) <= date) Then
            echeance.BackColor = &HF9D9D4
            err.Description = "L'échéance est dépassée"
            GoTo UpdateErr
        Else
            echeance.BackColor = &HFFFFFF
        End If
    End If



Exit Function
typmismatch:
    If montantCoti = "" And echeance = "" Then
    'rien n'est rendigné, on considère qu'il n'y en a pas
        datPrimaryRSCoti.Refresh
        saveCoti = True
    Else
        montantCoti.BackColor = &HF9D9D4
        saveCoti = False
        MsgBox "Le montant n'est pas valide"
    
    End If
  Exit Function
UpdateErr:
    If montantCoti = "" And echeance = "" Then
    'rien n'est rendigné, on considère qu'il n'y en a pas
        datPrimaryRSCoti.Refresh
        saveCoti = True
    Else
        MsgBox err.Description
        saveCoti = False
    End If
End Function

Private Sub btFermerVers_Click()
    If saveVers() Then
        FrameVers.Visible = False
        FrameCoti.Visible = False
        btNvCoti.Visible = False
        
    End If
End Sub
Private Function saveVers() As Boolean

   On Error GoTo UpdateErr
    
    If montantVers <> "" And dateVers <> "" And montantCoti = "" And echeance = "" Then
        ' pas de coti mais un versement, c'est pas normal
        saveVers = False
        MsgBox "Veuillez saisir une cotisation (ou pas de versement)"
        
        datPrimaryRSCoti.Recordset.AddNew
    Else
    
        saveVers = True
        
       'Montant Versement
        montantVers.SetFocus
        If (StrComp(montantVers, "") = 0) Then
            montantVers.BackColor = &HF9D9D4
            err.Description = "Le montant n'est pas valide"
            GoTo UpdateErr
        Else
            montantVers.BackColor = &HFFFFFF
            On Error GoTo typmismatch
            If (FormatCurrency(montantVers)) Then
            End If
            On Error GoTo UpdateErr
        End If
        
        'Date Versement
        dateVers.SetFocus
        If Not IsDate(dateVers) Then
                dateVers.BackColor = &HF9D9D4
                err.Description = "La date de versement n'est pas une date valide"
                GoTo UpdateErr
        Else
            dateVers.BackColor = &HFFFFFF
            If (DateValue(dateVers) > date) Then
                ' versement différé
                
                'dateVers.BackColor = &HF9D9D4
                'Err.Description = "Un tien vaut mieux que deux tu l'auras:" + Chr$(10) + "Veuillez saisir une date passée"
                'GoTo UpdateErr
            Else
                'dateVers.BackColor = &HFFFFFF
            End If
        End If
    
        'Type de payement
        
        'sauve
        datPrimaryRSVers.Recordset.UpdateBatch adAffectAll
        datPrimaryRSVers.Refresh
        datPrimaryRSVers.Refresh
    End If

    Exit Function
typmismatch:
    If montantVers = "" And dateVers = "" And drpTypPay = "" Then
    'rien n'est renseigné, on considère qu'il n'y en a pas
        saveVers = True
        datPrimaryRSVers.Refresh
    Else
        montantVers.BackColor = &HF9D9D4
        saveVers = False
        MsgBox "Le montant n'est pas valide"
    End If
  Exit Function
UpdateErr:
    If montantVers = "" And dateVers = "" And drpTypPay = "" Then
    'rien n'est renseigné, on considère qu'il n'y en a pas
         saveVers = True
        datPrimaryRSVers.Refresh
    Else
        MsgBox err.Description
        saveVers = False
    End If
End Function

Private Sub btNvAdh_Click()
On Error GoTo AddErr
  
    If sauve() Then
        Form_Load
    End If
  Exit Sub
AddErr:
  MsgBox err.Description
End Sub

Private Sub btNvChienChien_Click()
    On Error GoTo AddErr
    
    If saveChien() Then
        ' add
        datPrimaryRSChien.Recordset.AddNew
        'Ré-init champs
        Maitreid = adhid
        sexM = False
        sexF = False
        Call lstRaces_DblClick
        nomChien.SetFocus
    End If
Exit Sub
AddErr:
    MsgBox err.Description
End Sub
Private Sub btFermerChien_Click()
    If saveChien() Then
        frameChien.Visible = False
        FrameAdh.Enabled = True
        montantCoti.SetFocus
    End If
End Sub

Private Sub btFermer_Click()
    If sauve Then
        fGbl.Enabled = True
        Call fGbl.selAdh(adhid)
        Call fGbl.selCoti(adhid)
        nbrc = fGbl.selChienAdh(adhid)
        montantCoti.SetFocus
        fGbl.Enabled = False
        Unload Me
    End If
End Sub
Private Function sauve() As Boolean
    On Error GoTo UpdateErr
    sauve = False
    If saveProprio() Then
        
        'init les champs de coti
        debiteurid = adhid
        
        If saveCoti() Then
            'init les champs de versement
            payeurid = adhid
            If saveVers() Then
                'sauve coti
                datPrimaryRSCoti.Recordset.UpdateBatch adAffectAll
                datPrimaryRSCoti.Refresh
                datPrimaryRSCoti.Refresh
                
                sauve = True
            End If
        End If
    End If
    
Exit Function
UpdateErr:
    MsgBox err.Description
    sauve = False
End Function
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
        adhNbChien = adhNbChien + 1
        Call datPrimaryRS.Recordset.Update("Nb_Chiens", adhNbChien)
        datPrimaryRS.Refresh
        datPrimaryRS.Refresh
        
    End If
    
Exit Function
UpdateErr:
    MsgBox err.Description
    saveChien = False
End Function






Private Sub cpFermer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then btFermer_Click
End Sub

Private Sub cpFermerchien_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then btFermerChien_Click
End Sub



Private Sub drpSex_Click()
    If drpSex.ListIndex = 0 Then
        sex = 1
    Else
        sex = 0
    End If
End Sub

Private Sub drpTypPay_Validate(Cancel As Boolean)
    typPay = drpTypPay
End Sub




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

Private Sub editLabelRace_LostFocus()
    editLabelRace = lstRaces
End Sub




Private Sub Form_Initialize()
    ChDir (chemin)
End Sub

Private Sub Form_Load()
    
    Top = 0
    frameChien.Visible = False
    frameChien.Enabled = False
    
    datPrimaryRS.Recordset.AddNew
    datPrimaryRSCoti.Recordset.AddNew
    datPrimaryRSVers.Recordset.AddNew
        
    
    'Init Champs
    adhNbChien = 0
    
    Call InitDropRaces
    Call InitDropTypPay
    drpSex.AddItem ("M")
    drpSex.AddItem ("F")
    drpSex.ListIndex = 0
    sex = 1

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
Private Sub btNvChien_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvChien.Picture = btNvChienDown.Picture
End Sub
Private Sub btNvChien_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvChien.Picture = btNvChienUp.Picture
End Sub


Private Sub btNvAdh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvAdh.Picture = btNvadhDown.Picture
End Sub
Private Sub btNvAdh_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvAdh.Picture = btNvAdhUp.Picture
End Sub
Private Sub btFermer_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btFermer.Picture = fGbl.btFermerDown.Picture
    cpFermer.SetFocus
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
    cpFermerchien.SetFocus
End Sub
Private Sub btFermerChien_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
   btFermerChien.Picture = fGbl.btFermerUp.Picture
End Sub
Private Sub btNvChienChien_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvChienChien.Picture = btsavenvChiendown.Picture
End Sub
Private Sub btNvChienChien_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvChienChien.Picture = btsavenvChienup.Picture
End Sub


Private Sub btAnnulerAdh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerAdh.Picture = fGbl.btAnnulerDown.Picture
End Sub
Private Sub btAnnulerAdh_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerAdh.Picture = fGbl.btAnnulerUp.Picture
End Sub
Private Sub btAnnulerChien_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerChien.Picture = fGbl.btAnnulerDown.Picture
End Sub
Private Sub btAnnulerChien_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerChien.Picture = fGbl.btAnnulerUp.Picture
End Sub

