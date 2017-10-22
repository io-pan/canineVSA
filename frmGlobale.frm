VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmGlobale 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D5F1FF&
   BorderStyle     =   0  'None
   ClientHeight    =   9195
   ClientLeft      =   3705
   ClientTop       =   825
   ClientWidth     =   12030
   ForeColor       =   &H80000007&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1.022
   ScaleMode       =   0  'User
   ScaleWidth      =   1
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameConcoursF 
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   8400
      TabIndex        =   206
      Top             =   3720
      Width           =   2655
      Begin VB.Line Line2 
         BorderColor     =   &H00AAD69B&
         Index           =   1
         X1              =   0
         X2              =   2880
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00AAD69B&
         Index           =   0
         X1              =   2640
         X2              =   2640
         Y1              =   0
         Y2              =   800
      End
      Begin VB.Image btFermerConc 
         Height          =   255
         Left            =   240
         Picture         =   "frmGlobale.frx":0000
         Top             =   120
         Width           =   2250
      End
   End
   Begin VB.Frame FrameConcours 
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5055
      Left            =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   10575
      Begin VB.CheckBox concsixmois 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Seulement les 6 derniers mois"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Left            =   1440
         TabIndex        =   156
         Top             =   170
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox tbConcClas 
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
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   9480
         TabIndex        =   147
         Top             =   4440
         Width           =   495
      End
      Begin VB.TextBox tbConcPoints 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4108
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   8640
         TabIndex        =   146
         Top             =   4440
         Width           =   855
      End
      Begin VB.ComboBox drpTypApp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   145
         Top             =   4440
         Width           =   1335
      End
      Begin VB.ComboBox drpTypCat 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   144
         Top             =   4440
         Width           =   1455
      End
      Begin VB.ComboBox drpTypDisc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   143
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox tbConcdisc 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   12
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   104
         Top             =   3810
         Width           =   1455
      End
      Begin VB.TextBox tbConcCat 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   12
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   103
         Top             =   3810
         Width           =   1455
      End
      Begin VB.TextBox tbConcdisc 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   11
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   102
         Top             =   3540
         Width           =   1455
      End
      Begin VB.TextBox tbConcCat 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   11
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   101
         Top             =   3540
         Width           =   1455
      End
      Begin VB.TextBox tbConcdisc 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   10
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   3270
         Width           =   1455
      End
      Begin VB.TextBox tbConcCat 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   10
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   3270
         Width           =   1455
      End
      Begin VB.TextBox tbConcdisc 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   9
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox tbConcCat 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   9
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox tbConcdisc 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   8
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   2730
         Width           =   1455
      End
      Begin VB.TextBox tbConcCat 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   8
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   2730
         Width           =   1455
      End
      Begin VB.TextBox tbConcdisc 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   7
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   2460
         Width           =   1455
      End
      Begin VB.TextBox tbConcCat 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   7
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   2460
         Width           =   1455
      End
      Begin VB.TextBox tbConcdisc 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   6
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   2190
         Width           =   1455
      End
      Begin VB.TextBox tbConcCat 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   6
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   2190
         Width           =   1455
      End
      Begin VB.TextBox tbConcdisc 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   5
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox tbConcCat 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   5
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox tbConcdisc 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   4
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   1650
         Width           =   1455
      End
      Begin VB.TextBox tbConcCat 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   4
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   1650
         Width           =   1455
      End
      Begin VB.TextBox tbConcdisc 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   3
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   1380
         Width           =   1455
      End
      Begin VB.TextBox tbConcCat 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   3
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   1380
         Width           =   1455
      End
      Begin VB.TextBox tbConcdisc 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   2
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   1110
         Width           =   1455
      End
      Begin VB.TextBox tbConcCat 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   2
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   1110
         Width           =   1455
      End
      Begin VB.VScrollBar VScrollConcours 
         Height          =   3255
         LargeChange     =   4
         Left            =   10080
         TabIndex        =   71
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox tbConcDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   12
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   3810
         Width           =   1095
      End
      Begin VB.TextBox tbConcLieu 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   12
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   3810
         Width           =   3015
      End
      Begin VB.TextBox tbConcApp 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   12
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   3810
         Width           =   1335
      End
      Begin VB.TextBox tbConcPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   12
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   3810
         Width           =   855
      End
      Begin VB.TextBox tbConcClas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   12
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   3810
         Width           =   495
      End
      Begin VB.TextBox tbConcDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   11
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   3540
         Width           =   1095
      End
      Begin VB.TextBox tbConcLieu 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   11
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   3540
         Width           =   3015
      End
      Begin VB.TextBox tbConcApp 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   11
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   3540
         Width           =   1335
      End
      Begin VB.TextBox tbConcPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   11
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   3540
         Width           =   855
      End
      Begin VB.TextBox tbConcClas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   11
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   3540
         Width           =   495
      End
      Begin VB.TextBox tbConcDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   10
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   3270
         Width           =   1095
      End
      Begin VB.TextBox tbConcLieu 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   10
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   3270
         Width           =   3015
      End
      Begin VB.TextBox tbConcApp 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   10
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   3270
         Width           =   1335
      End
      Begin VB.TextBox tbConcPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   10
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   3270
         Width           =   855
      End
      Begin VB.TextBox tbConcClas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   10
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   3270
         Width           =   495
      End
      Begin VB.TextBox tbConcDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   9
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox tbConcLieu 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   9
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   3000
         Width           =   3015
      End
      Begin VB.TextBox tbConcApp 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   9
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox tbConcPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   9
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox tbConcClas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   9
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox tbConcDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   8
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   2730
         Width           =   1095
      End
      Begin VB.TextBox tbConcLieu 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   8
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   2730
         Width           =   3015
      End
      Begin VB.TextBox tbConcApp 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   8
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   2730
         Width           =   1335
      End
      Begin VB.TextBox tbConcPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   8
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   2730
         Width           =   855
      End
      Begin VB.TextBox tbConcClas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   8
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   2730
         Width           =   495
      End
      Begin VB.TextBox tbConcDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   7
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   2460
         Width           =   1095
      End
      Begin VB.TextBox tbConcLieu 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   7
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   2460
         Width           =   3015
      End
      Begin VB.TextBox tbConcApp 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   7
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   2460
         Width           =   1335
      End
      Begin VB.TextBox tbConcPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   7
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   2460
         Width           =   855
      End
      Begin VB.TextBox tbConcClas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   7
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2460
         Width           =   495
      End
      Begin VB.TextBox tbConcDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   6
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   2190
         Width           =   1095
      End
      Begin VB.TextBox tbConcLieu 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   6
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2190
         Width           =   3015
      End
      Begin VB.TextBox tbConcApp 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   6
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2190
         Width           =   1335
      End
      Begin VB.TextBox tbConcPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   6
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2190
         Width           =   855
      End
      Begin VB.TextBox tbConcClas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   6
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2190
         Width           =   495
      End
      Begin VB.TextBox tbConcDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   5
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox tbConcLieu 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   5
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox tbConcApp 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   5
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox tbConcPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   5
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox tbConcClas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   5
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox tbConcClas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   1
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox tbConcPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   1
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox tbConcApp 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   1
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox tbConcCat 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   1
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox tbConcdisc 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   1
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox tbConcLieu 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox tbConcDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   1
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox tbConcLieu 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1440
         TabIndex        =   142
         Top             =   4440
         Width           =   3015
      End
      Begin VB.TextBox tbConcDate 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd.MM.yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4108
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   360
         TabIndex        =   141
         Top             =   4440
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc datPrimaryRSConcours 
         Height          =   330
         Left            =   2040
         Top             =   240
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
         RecordSource    =   $"frmGlobale.frx":0BD8
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
      Begin VB.TextBox tbConcClas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   3
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1380
         Width           =   495
      End
      Begin VB.TextBox tbConcPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   3
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1380
         Width           =   855
      End
      Begin VB.TextBox tbConcApp 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   3
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1380
         Width           =   1335
      End
      Begin VB.TextBox tbConcLieu 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   3
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1380
         Width           =   3015
      End
      Begin VB.TextBox tbConcDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   3
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1380
         Width           =   1095
      End
      Begin VB.TextBox tbConcClas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   4
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1650
         Width           =   495
      End
      Begin VB.TextBox tbConcPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   4
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1650
         Width           =   855
      End
      Begin VB.TextBox tbConcApp 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   4
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1650
         Width           =   1335
      End
      Begin VB.TextBox tbConcLieu 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   4
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1650
         Width           =   3015
      End
      Begin VB.TextBox tbConcDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   4
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1650
         Width           =   1095
      End
      Begin VB.TextBox tbConcClas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   2
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1110
         Width           =   495
      End
      Begin VB.TextBox tbConcPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   2
         Left            =   8640
         TabIndex        =   17
         Top             =   1110
         Width           =   855
      End
      Begin VB.TextBox tbConcApp 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   2
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1110
         Width           =   1335
      End
      Begin VB.TextBox tbConcLieu 
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1110
         Width           =   3015
      End
      Begin VB.TextBox tbConcDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   260
         Index           =   2
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1110
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid grdConc 
         Bindings        =   "frmGlobale.frx":0C65
         Height          =   5655
         Left            =   4560
         TabIndex        =   132
         Top             =   1080
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   9975
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Date"
            Caption         =   "Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4108
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Lieu"
            Caption         =   "Lieu"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4108
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Discipline"
            Caption         =   "Discipline"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4108
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Categorie"
            Caption         =   "Categorie"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4108
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Appreciation"
            Caption         =   "Appreciation"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4108
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Points"
            Caption         =   "Points"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4108
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Classement"
            Caption         =   "Classement"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4108
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   0
            BeginProperty Column00 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   824.882
            EndProperty
         EndProperty
      End
      Begin VB.TextBox concApp 
         DataField       =   "Appreciation"
         DataSource      =   "datPrimaryRSConcours"
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   148
         Text            =   "0"
         Top             =   4440
         Visible         =   0   'False
         Width           =   1095
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   149
         Text            =   "0"
         Top             =   4440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox concidChien 
         DataField       =   "ChienID"
         DataSource      =   "datPrimaryRSConcours"
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   150
         Top             =   4440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox concClas 
         Alignment       =   1  'Right Justify
         DataField       =   "Classement"
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
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   151
         Top             =   4440
         Visible         =   0   'False
         Width           =   495
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
         TabIndex        =   152
         Text            =   "0"
         Top             =   4440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox concLieu 
         DataField       =   "Lieu"
         DataSource      =   "datPrimaryRSConcours"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   153
         Top             =   4440
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox concPoints 
         DataField       =   "Points"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4108
            SubFormatType   =   1
         EndProperty
         DataSource      =   "datPrimaryRSConcours"
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   154
         Top             =   4440
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox concDate 
         DataField       =   "Date"
         DataSource      =   "datPrimaryRSConcours"
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   155
         Top             =   4440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox nbLignes 
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   133
         Text            =   "Text1"
         Top             =   4440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E2F5FF&
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         Height          =   5415
         Index           =   5
         Left            =   0
         Top             =   0
         Width           =   10575
      End
      Begin VB.Image btSortConc 
         Height          =   255
         Index           =   0
         Left            =   360
         Picture         =   "frmGlobale.frx":0C88
         Top             =   540
         Width           =   1080
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Concours"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E03080&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Nouveau concours :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Left            =   240
         TabIndex        =   134
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Image btAddConcours 
         Height          =   375
         Left            =   10080
         Picture         =   "frmGlobale.frx":11FC
         Top             =   4440
         Width           =   375
      End
      Begin VB.Image btSortConc 
         Height          =   255
         Index           =   4
         Left            =   7300
         Picture         =   "frmGlobale.frx":1550
         Top             =   540
         Width           =   1335
      End
      Begin VB.Image btSortConcCla 
         Height          =   255
         Index           =   6
         Left            =   9490
         Picture         =   "frmGlobale.frx":1C18
         Top             =   540
         Width           =   480
      End
      Begin VB.Image btSortConcPt 
         Height          =   255
         Index           =   5
         Left            =   8640
         Picture         =   "frmGlobale.frx":1EE4
         Top             =   540
         Width           =   855
      End
      Begin VB.Image btSortConc 
         Height          =   255
         Index           =   2
         Left            =   4420
         Picture         =   "frmGlobale.frx":238C
         Top             =   540
         Width           =   2880
      End
      Begin VB.Image btSortConc 
         Height          =   255
         Index           =   1
         Left            =   1440
         Picture         =   "frmGlobale.frx":30F8
         Top             =   540
         Width           =   2985
      End
      Begin VB.Image Image5 
         Height          =   210
         Left            =   0
         Picture         =   "frmGlobale.frx":3EEC
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   10290
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00AAD69B&
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         FillStyle       =   0  'Solid
         Height          =   450
         Left            =   0
         Top             =   4890
         Width           =   10575
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00AAD69B&
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         FillStyle       =   0  'Solid
         Height          =   3555
         Left            =   210
         Top             =   600
         Width           =   10200
      End
   End
   Begin VB.Frame FrameCoti 
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Caption         =   "Cotisation"
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
      Height          =   4935
      Left            =   1080
      TabIndex        =   105
      Top             =   3720
      Width           =   4695
      Begin VB.TextBox TotVers 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "# ##0,00 ""€"""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   115
         Text            =   "0.00"
         Top             =   3720
         Width           =   735
      End
      Begin MSDataGridLib.DataGrid gridVers 
         Bindings        =   "frmGlobale.frx":6204
         Height          =   1935
         Left            =   1080
         TabIndex        =   131
         Top             =   1800
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         ColumnHeaders   =   0   'False
         HeadLines       =   0
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Montant"
            Caption         =   "Montant"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4108
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "AdhID"
            Caption         =   "AdhID"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4108
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4108
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "TypePayement"
            Caption         =   "TypePayement"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4108
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Date"
            Caption         =   "Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd.MM.yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4108
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   2
            RecordSelectors =   0   'False
            BeginProperty Column00 
               Alignment       =   1
               DividerStyle    =   0
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   74.835
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   74.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   975.118
            EndProperty
         EndProperty
      End
      Begin VB.TextBox nvVersDate 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd.MM.yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4108
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   108
         Top             =   4440
         Width           =   975
      End
      Begin VB.ComboBox drpTypPay 
         Height          =   315
         ItemData        =   "frmGlobale.frx":6223
         Left            =   1800
         List            =   "frmGlobale.frx":6225
         TabIndex        =   107
         Text            =   "drpTypPay"
         Top             =   4440
         Width           =   1380
      End
      Begin VB.TextBox nvVersMontant 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4108
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   106
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox debiteurid 
         DataField       =   "AdhID"
         DataSource      =   "datPrimaryRSCoti"
         Height          =   285
         Left            =   2760
         TabIndex        =   120
         Top             =   -120
         Visible         =   0   'False
         Width           =   495
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
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   119
         Text            =   "0.00"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox labelTypPay 
         DataField       =   "Label"
         DataSource      =   "datPrimaryRSTypPay"
         Height          =   285
         Left            =   240
         TabIndex        =   118
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         DataField       =   "id"
         DataSource      =   "datPrimaryRSTypPay"
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   117
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox typPay 
         DataField       =   "TypePayement"
         DataSource      =   "datPrimaryRSVers"
         Height          =   285
         Left            =   240
         TabIndex        =   116
         Top             =   3720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox EcheancePrec 
         Alignment       =   2  'Center
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   114
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Solde 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "0.00"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox payeurid 
         DataField       =   "AdhID"
         DataSource      =   "datPrimaryRSVers"
         Height          =   285
         Left            =   0
         TabIndex        =   112
         Top             =   3960
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
         Left            =   480
         TabIndex        =   111
         Top             =   3120
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
         Left            =   240
         TabIndex        =   110
         Top             =   3240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox echeance 
         Alignment       =   2  'Center
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
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   109
         Top             =   1080
         Width           =   975
      End
      Begin MSAdodcLib.Adodc datPrimaryRSCoti 
         Height          =   330
         Left            =   1560
         Top             =   -120
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
      Begin MSAdodcLib.Adodc datPrimaryRSTypPay 
         Height          =   330
         Left            =   0
         Top             =   2760
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
      Begin MSAdodcLib.Adodc datPrimaryRSVers 
         Height          =   330
         Left            =   0
         Top             =   2160
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
      Begin VB.Shape Shape1 
         BackColor       =   &H00E2F5FF&
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         Height          =   4935
         Index           =   6
         Left            =   0
         Top             =   0
         Width           =   4695
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00AAD69B&
         X1              =   4440
         X2              =   1065
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00AAD69B&
         X1              =   1065
         X2              =   1065
         Y1              =   3960
         Y2              =   1785
      End
      Begin VB.Label lblLabels 
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
         ForeColor       =   &H00E03080&
         Height          =   375
         Index           =   19
         Left            =   120
         TabIndex        =   121
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Date :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Left            =   3120
         TabIndex        =   136
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Nouveau Versement :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E03080&
         Height          =   255
         Left            =   360
         TabIndex        =   135
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Echéance :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   18
         Left            =   2160
         TabIndex        =   130
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label euro 
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
         ForeColor       =   &H00FF2E55&
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   129
         Top             =   600
         Width           =   135
      End
      Begin VB.Image btNvCoti 
         Height          =   375
         Left            =   2160
         Picture         =   "frmGlobale.frx":6227
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2F5FF&
         Caption         =   "Total :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Left            =   480
         TabIndex        =   128
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Anc. Ech. :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Left            =   2160
         TabIndex        =   127
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Solde :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Left            =   360
         TabIndex        =   126
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label euro 
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
         ForeColor       =   &H00FF2E55&
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   125
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label euro 
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
         ForeColor       =   &H00FF2E55&
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   124
         Top             =   960
         Width           =   135
      End
      Begin VB.Image btNvVers 
         Height          =   375
         Left            =   4200
         Picture         =   "frmGlobale.frx":71F3
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Montant :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Left            =   360
         TabIndex        =   123
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "Versements :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E03080&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   122
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Image btScrollUpEch 
         Height          =   255
         Left            =   4200
         Picture         =   "frmGlobale.frx":7547
         Top             =   720
         Width           =   255
      End
      Begin VB.Image btScrollDownEch 
         Height          =   255
         Left            =   4200
         Picture         =   "frmGlobale.frx":782F
         Top             =   1080
         Width           =   255
      End
   End
   Begin VB.Frame FrameCHIENS 
      BackColor       =   &H00D5F1FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7455
      Left            =   6000
      TabIndex        =   158
      Top             =   1680
      Width           =   4935
      Begin VB.Frame FrameChien 
         BackColor       =   &H00E2F5FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2415
         Index           =   1
         Left            =   0
         TabIndex        =   174
         Top             =   2520
         Visible         =   0   'False
         Width           =   4455
         Begin VB.TextBox cun 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00E03080&
            Height          =   285
            Index           =   1
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   213
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox cnea 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00E03080&
            Height          =   285
            Index           =   1
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   214
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox dateNaiss 
            BackColor       =   &H8000000F&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd.MM.yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4108
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   960
            TabIndex        =   185
            Top             =   2280
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox tatou 
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
            Height          =   285
            Index           =   1
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   184
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox chienlof 
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
            Height          =   285
            Index           =   1
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   183
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox nomChien 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003030CC&
            Height          =   330
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   182
            Top             =   120
            Width           =   4215
         End
         Begin VB.CheckBox sex 
            Caption         =   "sex"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   181
            Top             =   2280
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox idChien 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   480
            TabIndex        =   180
            Top             =   2280
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
            Index           =   1
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   179
            Text            =   "0"
            Top             =   570
            Width           =   255
         End
         Begin VB.TextBox labelSex 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
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
            Index           =   1
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   178
            Top             =   570
            Width           =   255
         End
         Begin VB.TextBox labelRace 
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
            Height          =   285
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   177
            Top             =   570
            Width           =   3255
         End
         Begin VB.TextBox idRaceChien 
            BackColor       =   &H8000000F&
            DataField       =   "Race"
            DataSource      =   "datPrimaryRSChi"
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   0
            TabIndex        =   176
            Top             =   2160
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00AAD69B&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   1
            Left            =   0
            ScaleHeight     =   450
            ScaleWidth      =   4935
            TabIndex        =   175
            Top             =   2000
            Width           =   4935
            Begin VB.Image modChien 
               Height          =   345
               Index           =   1
               Left            =   1080
               Picture         =   "frmGlobale.frx":7B17
               Top             =   0
               Width           =   1215
            End
            Begin VB.Image btSupChien 
               Height          =   255
               Index           =   1
               Left            =   120
               Picture         =   "frmGlobale.frx":7EEA
               Stretch         =   -1  'True
               Top             =   90
               Width           =   855
            End
            Begin VB.Image btAffConc 
               Height          =   255
               Index           =   1
               Left            =   2400
               Picture         =   "frmGlobale.frx":836E
               Top             =   90
               Width           =   1950
            End
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "Licence CUN :"
            ForeColor       =   &H00FF2E55&
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   216
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "Licence CNEA :"
            ForeColor       =   &H00FF2E55&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   215
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "N° LOF :"
            ForeColor       =   &H00FF2E55&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   188
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "Tatouage :"
            ForeColor       =   &H00FF2E55&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   187
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2F5FF&
            Caption         =   "ans"
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
            Height          =   285
            Index           =   1
            Left            =   3960
            TabIndex        =   186
            Top             =   570
            Width           =   375
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00AAD69B&
            FillColor       =   &H00AAD69B&
            Height          =   2415
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   4455
         End
         Begin VB.Image ImgChienFooter 
            Height          =   795
            Index           =   1
            Left            =   0
            Picture         =   "frmGlobale.frx":8CBA
            Top             =   1200
            Width           =   4470
         End
      End
      Begin VB.Frame FrameChien 
         BackColor       =   &H00E2F5FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2415
         Index           =   0
         Left            =   0
         TabIndex        =   159
         Top             =   0
         Width           =   4455
         Begin VB.TextBox cun 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
            BorderStyle     =   0  'None
            DataField       =   "no_CUN"
            DataSource      =   "datPrimaryRSChi"
            ForeColor       =   &H00E03080&
            Height          =   285
            Index           =   0
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   212
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox cnea 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
            BorderStyle     =   0  'None
            DataField       =   "no_CNEA"
            DataSource      =   "datPrimaryRSChi"
            ForeColor       =   &H00E03080&
            Height          =   285
            Index           =   0
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   211
            Top             =   1440
            Width           =   855
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
            Left            =   2400
            TabIndex        =   170
            Top             =   2160
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
            Left            =   1560
            TabIndex        =   169
            Top             =   2280
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox tatou 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
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
            Locked          =   -1  'True
            TabIndex        =   168
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox chienlof 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
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
            Locked          =   -1  'True
            TabIndex        =   167
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox nomChien 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
            BorderStyle     =   0  'None
            DataField       =   "Nom"
            DataSource      =   "datPrimaryRSChi"
            BeginProperty Font 
               Name            =   "Arial"
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
            Locked          =   -1  'True
            TabIndex        =   166
            Top             =   120
            Width           =   4215
         End
         Begin VB.CheckBox sex 
            Caption         =   "sex"
            DataField       =   "Sexe"
            DataSource      =   "datPrimaryRSChi"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   165
            Top             =   1920
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox idChien 
            DataField       =   "id"
            DataSource      =   "datPrimaryRSChi"
            Height          =   285
            Index           =   0
            Left            =   3240
            TabIndex        =   164
            Top             =   2280
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
            Height          =   285
            Index           =   0
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   163
            Text            =   "0"
            Top             =   570
            Width           =   255
         End
         Begin VB.TextBox labelSex 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
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
            Height          =   285
            Index           =   0
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   162
            Top             =   570
            Width           =   255
         End
         Begin VB.TextBox Maitreid 
            BackColor       =   &H8000000F&
            DataField       =   "MaitreID"
            DataSource      =   "datPrimaryRSChi"
            Enabled         =   0   'False
            Height          =   285
            Left            =   0
            TabIndex        =   161
            Top             =   2160
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox labelRace 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   160
            Top             =   570
            Width           =   3255
         End
         Begin MSAdodcLib.Adodc datPrimaryRSChi 
            Height          =   330
            Left            =   240
            Top             =   2160
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
            RecordSource    =   "select DateNaiss,id,LOF,MaitreID,Nom,Race,Sexe,Tatouage,no_CNEA,no_CUN from Chiens where id <0"
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
            Appearance      =   0  'Flat
            BackColor       =   &H00AAD69B&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   0
            Left            =   0
            ScaleHeight     =   450
            ScaleWidth      =   4935
            TabIndex        =   205
            Top             =   2000
            Width           =   4935
            Begin VB.Image modChien 
               Height          =   345
               Index           =   0
               Left            =   1080
               Picture         =   "frmGlobale.frx":8F50
               Top             =   0
               Width           =   1215
            End
            Begin VB.Image btAffConc 
               Height          =   255
               Index           =   0
               Left            =   2400
               Picture         =   "frmGlobale.frx":9323
               Top             =   90
               Width           =   1950
            End
            Begin VB.Image btSupChien 
               Height          =   255
               Index           =   0
               Left            =   120
               Picture         =   "frmGlobale.frx":9C6F
               Stretch         =   -1  'True
               Top             =   90
               Width           =   855
            End
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "Licence CNEA :"
            ForeColor       =   &H00FF2E55&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   210
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "Licence CUN :"
            ForeColor       =   &H00FF2E55&
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   209
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "N° LOF :"
            ForeColor       =   &H00FF2E55&
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   173
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "Tatouage :"
            ForeColor       =   &H00FF2E55&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   172
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2F5FF&
            Caption         =   "ans"
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
            Height          =   285
            Index           =   0
            Left            =   3960
            TabIndex        =   171
            Top             =   570
            Width           =   375
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E2F5FF&
            BorderColor     =   &H00AAD69B&
            FillColor       =   &H00AAD69B&
            Height          =   2415
            Index           =   0
            Left            =   0
            Top             =   0
            Width           =   4455
         End
         Begin VB.Image ImgChienFooter 
            Height          =   795
            Index           =   0
            Left            =   0
            Picture         =   "frmGlobale.frx":A0F3
            Top             =   1200
            Width           =   4470
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00AAD69B&
            FillColor       =   &H00AAD69B&
            FillStyle       =   0  'Solid
            Height          =   450
            Left            =   0
            Top             =   1635
            Width           =   4455
         End
      End
      Begin VB.VScrollBar VScrollChiens 
         Height          =   7455
         Left            =   4560
         TabIndex        =   204
         Top             =   0
         Width           =   255
      End
      Begin VB.Frame FrameChien 
         BackColor       =   &H00E2F5FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2415
         Index           =   2
         Left            =   0
         TabIndex        =   189
         Top             =   5040
         Visible         =   0   'False
         Width           =   4455
         Begin VB.TextBox cun 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00E03080&
            Height          =   285
            Index           =   2
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   217
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox cnea 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00E03080&
            Height          =   285
            Index           =   2
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   218
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox labelRace 
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
            Height          =   285
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   200
            Top             =   570
            Width           =   3255
         End
         Begin VB.TextBox labelSex 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
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
            Height          =   285
            Index           =   2
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   199
            Top             =   570
            Width           =   255
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
            Height          =   285
            Index           =   2
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   198
            Text            =   "0"
            Top             =   570
            Width           =   255
         End
         Begin VB.TextBox idChien 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   600
            TabIndex        =   197
            Top             =   2160
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CheckBox sex 
            Caption         =   "sex"
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   196
            Top             =   2280
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox nomChien 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E2F5FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003030CC&
            Height          =   330
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   195
            Top             =   120
            Width           =   4215
         End
         Begin VB.TextBox chienlof 
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
            Height          =   285
            Index           =   2
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   194
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox tatou 
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
            Height          =   285
            Index           =   2
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   193
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox dateNaiss 
            BackColor       =   &H8000000F&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd.MM.yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4108
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1440
            TabIndex        =   192
            Top             =   1800
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox idRaceChien 
            BackColor       =   &H8000000F&
            DataField       =   "Race"
            DataSource      =   "datPrimaryRSChi"
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   191
            Top             =   2040
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00AAD69B&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   2
            Left            =   0
            ScaleHeight     =   450
            ScaleWidth      =   4935
            TabIndex        =   190
            Top             =   2000
            Width           =   4935
            Begin VB.Image modChien 
               Height          =   345
               Index           =   2
               Left            =   1080
               Picture         =   "frmGlobale.frx":A389
               Top             =   0
               Width           =   1215
            End
            Begin VB.Image btSupChien 
               Height          =   255
               Index           =   2
               Left            =   120
               Picture         =   "frmGlobale.frx":A75C
               Stretch         =   -1  'True
               Top             =   90
               Width           =   855
            End
            Begin VB.Image btAffConc 
               Height          =   255
               Index           =   2
               Left            =   2400
               Picture         =   "frmGlobale.frx":ABE0
               Top             =   90
               Width           =   1950
            End
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "Licence CNEA :"
            ForeColor       =   &H00FF2E55&
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   220
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "Licence CUN :"
            ForeColor       =   &H00FF2E55&
            Height          =   285
            Index           =   12
            Left            =   120
            TabIndex        =   219
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2F5FF&
            Caption         =   "ans"
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
            Height          =   285
            Index           =   2
            Left            =   3960
            TabIndex        =   203
            Top             =   570
            Width           =   375
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "Tatouage :"
            ForeColor       =   &H00FF2E55&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   202
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00E2F5FF&
            Caption         =   "N° LOF :"
            ForeColor       =   &H00FF2E55&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   201
            Top             =   1080
            Width           =   975
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00AAD69B&
            FillColor       =   &H00AAD69B&
            Height          =   2415
            Index           =   2
            Left            =   0
            Top             =   0
            Width           =   4455
         End
         Begin VB.Image ImgChienFooter 
            Height          =   795
            Index           =   2
            Left            =   0
            Picture         =   "frmGlobale.frx":B52C
            Top             =   1200
            Width           =   4470
         End
      End
      Begin VB.Image btNextChien 
         Height          =   255
         Left            =   4560
         Picture         =   "frmGlobale.frx":B7C2
         Stretch         =   -1  'True
         Top             =   360
         Width           =   255
      End
      Begin VB.Image btPevChien 
         Height          =   255
         Left            =   4560
         Picture         =   "frmGlobale.frx":BAAA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.TextBox idRace 
      DataField       =   "ID"
      DataSource      =   "datPrimaryRSRace"
      Height          =   285
      Left            =   -480
      TabIndex        =   6
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstRaces 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      DataField       =   "Race"
      DataSource      =   "datPrimaryRSChien"
      Height          =   225
      Left            =   -480
      TabIndex        =   5
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frameCherchAdh 
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      Begin VB.TextBox rechAdh 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3135
      End
      Begin VB.TextBox totAdh 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2F5FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF2E55&
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   138
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E2F5FF&
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         Height          =   855
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label lblTotAdh 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2F5FF&
         Caption         =   "adh."
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Left            =   840
         TabIndex        =   137
         Top             =   480
         Width           =   375
      End
      Begin VB.Image btSelAdh 
         Height          =   255
         Left            =   2040
         Picture         =   "frmGlobale.frx":BD92
         Top             =   480
         Width           =   1215
      End
      Begin VB.Image btNextAdh 
         Height          =   255
         Left            =   1680
         Picture         =   "frmGlobale.frx":C3AE
         Stretch         =   -1  'True
         Top             =   480
         Width           =   255
      End
      Begin VB.Image btPrevAdh 
         Height          =   255
         Left            =   1320
         Picture         =   "frmGlobale.frx":C6EA
         Stretch         =   -1  'True
         Top             =   480
         Width           =   255
      End
   End
   Begin MSAdodcLib.Adodc datPrimaryRSRace 
      Height          =   330
      Left            =   -480
      Top             =   6720
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
   Begin VB.Frame FrameAdh 
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   360
      TabIndex        =   72
      Top             =   1080
      Width           =   7815
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2F5FF&
         BorderStyle     =   0  'None
         DataField       =   "No_SCRA"
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
         ForeColor       =   &H003030CC&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   207
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox NomPrenom 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2F5FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H003030CC&
         Height          =   330
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   120
         Width           =   4215
      End
      Begin VB.TextBox prenomAdh 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         ForeColor       =   &H003030CC&
         Height          =   285
         Left            =   3360
         TabIndex        =   82
         Text            =   "Prénom"
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox adr 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2F5FF&
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
         ForeColor       =   &H00FF2E55&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox telport 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2F5FF&
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
         ForeColor       =   &H00FF2E55&
         Height          =   220
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox teldom 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2F5FF&
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
         ForeColor       =   &H00FF2E55&
         Height          =   220
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox adhid 
         DataField       =   "id"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         TabIndex        =   78
         Text            =   "id"
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox codePost 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2F5FF&
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
         ForeColor       =   &H00FF2E55&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   1605
         Width           =   615
      End
      Begin VB.TextBox nomAdh 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         ForeColor       =   &H003030CC&
         Height          =   285
         Left            =   1320
         TabIndex        =   76
         Text            =   "Nom"
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox adhNbChien 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2F5FF&
         BorderStyle     =   0  'None
         DataField       =   "Nb_Chiens"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   4108
            SubFormatType   =   1
         EndProperty
         DataSource      =   "datPrimaryRS"
         ForeColor       =   &H00FF2E55&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   74
         Text            =   "0"
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox ville 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2F5FF&
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
         ForeColor       =   &H00FF2E55&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1605
         Width           =   3255
      End
      Begin MSAdodcLib.Adodc datPrimaryRS 
         Height          =   330
         Left            =   3480
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
         RecordSource    =   "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Adresse,CodePostal,Ville,Nb_Chiens,No_SCRA from Proprietaire where id <0"
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
      Begin VB.Image photoAdh 
         BorderStyle     =   1  'Fixed Single
         Height          =   1335
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00E2F5FF&
         Caption         =   "N° SCRA :"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   208
         Top             =   480
         Width           =   855
      End
      Begin VB.Image btModAdh 
         Height          =   345
         Left            =   2280
         Picture         =   "frmGlobale.frx":CAA2
         Top             =   2130
         Width           =   1215
      End
      Begin VB.Shape shapBordAdh 
         BackColor       =   &H00E2F5FF&
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         Height          =   2655
         Left            =   1200
         Top             =   0
         Width           =   6615
      End
      Begin VB.Image btSupAdh 
         Height          =   255
         Left            =   1320
         Picture         =   "frmGlobale.frx":CE75
         Top             =   2220
         Width           =   855
      End
      Begin VB.Image Image4 
         Height          =   450
         Left            =   1200
         Picture         =   "frmGlobale.frx":D2F9
         Top             =   1680
         Width           =   4410
      End
      Begin VB.Image Image3 
         Height          =   285
         Left            =   1320
         Picture         =   "frmGlobale.frx":F611
         Top             =   900
         Width           =   330
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   3480
         Picture         =   "frmGlobale.frx":F999
         Top             =   765
         Width           =   360
      End
      Begin VB.Image btAddChien 
         Height          =   375
         Left            =   6480
         Picture         =   "frmGlobale.frx":FE99
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblNbChien 
         BackColor       =   &H00E2F5FF&
         Caption         =   "chien"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Left            =   5880
         TabIndex        =   157
         Top             =   240
         Width           =   615
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         FillStyle       =   0  'Solid
         Height          =   900
         Left            =   1200
         Top             =   2115
         Width           =   6615
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00D5F1FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   2655
         Left            =   0
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame frameCherchChien 
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   6000
      TabIndex        =   4
      Top             =   120
      Width           =   3375
      Begin VB.TextBox totChiens 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2F5FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF2E55&
         Height          =   195
         Left            =   200
         Locked          =   -1  'True
         TabIndex        =   139
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox rechChien 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   3135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E2F5FF&
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         Height          =   855
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label lblTotChien 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2F5FF&
         Caption         =   "chiens"
         ForeColor       =   &H00FF2E55&
         Height          =   255
         Left            =   720
         TabIndex        =   140
         Top             =   480
         Width           =   495
      End
      Begin VB.Image btPevChienRech 
         Height          =   255
         Left            =   1320
         Picture         =   "frmGlobale.frx":1075D
         Stretch         =   -1  'True
         Top             =   480
         Width           =   255
      End
      Begin VB.Image btNextChienRech 
         Height          =   255
         Left            =   1680
         Picture         =   "frmGlobale.frx":10B15
         Stretch         =   -1  'True
         Top             =   480
         Width           =   255
      End
      Begin VB.Image btSelChien 
         Height          =   255
         Left            =   2040
         Picture         =   "frmGlobale.frx":10E51
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   150
   End
   Begin VB.Image btRappeldown 
      Height          =   255
      Left            =   120
      Picture         =   "frmGlobale.frx":1146D
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image btRappelup 
      Height          =   255
      Left            =   0
      Picture         =   "frmGlobale.frx":11B01
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image btModdown 
      Height          =   345
      Left            =   0
      Picture         =   "frmGlobale.frx":1211D
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image btModup 
      Height          =   345
      Left            =   120
      Picture         =   "frmGlobale.frx":1256A
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image btPtFermerD 
      Height          =   255
      Left            =   120
      Picture         =   "frmGlobale.frx":1293D
      Top             =   8760
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image btPtFermerU 
      Height          =   255
      Left            =   120
      Picture         =   "frmGlobale.frx":13579
      Top             =   8640
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image btAddChiendown 
      Height          =   375
      Left            =   120
      Picture         =   "frmGlobale.frx":14151
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image btAddChienup 
      Height          =   375
      Left            =   0
      Picture         =   "frmGlobale.frx":14A71
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image btSaveDown 
      Height          =   375
      Left            =   240
      Picture         =   "frmGlobale.frx":15335
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image btSaveUp 
      Height          =   375
      Left            =   120
      Picture         =   "frmGlobale.frx":156C5
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image btAffConcUp 
      Height          =   255
      Left            =   120
      Picture         =   "frmGlobale.frx":15A19
      Top             =   1800
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Image btAffConcDown 
      Height          =   255
      Left            =   0
      Picture         =   "frmGlobale.frx":16365
      Top             =   1800
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Image btsupprUp 
      Height          =   255
      Left            =   120
      Picture         =   "frmGlobale.frx":16D29
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   1
      Left            =   9510
      Picture         =   "frmGlobale.frx":171AD
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imgBienvenu 
      Height          =   4200
      Left            =   2160
      Picture         =   "frmGlobale.frx":17F61
      Top             =   1680
      Width           =   7350
   End
   Begin VB.Image btScrollRightDown 
      Height          =   255
      Left            =   240
      Picture         =   "frmGlobale.frx":1C41B
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image btScrollRightUp 
      Height          =   255
      Left            =   240
      Picture         =   "frmGlobale.frx":1C76F
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image btScrollLeftDown 
      Height          =   255
      Left            =   0
      Picture         =   "frmGlobale.frx":1CAAB
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image btScrollLeftUp 
      Height          =   255
      Left            =   0
      Picture         =   "frmGlobale.frx":1CDFB
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image btScrolldowndown 
      Height          =   255
      Left            =   480
      Picture         =   "frmGlobale.frx":1D1B3
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image btScrolldownUp 
      Height          =   255
      Left            =   480
      Picture         =   "frmGlobale.frx":1D497
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image btScrollUpDown 
      Height          =   255
      Left            =   720
      Picture         =   "frmGlobale.frx":1D77F
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image btScrollUpUp 
      Height          =   255
      Left            =   720
      Picture         =   "frmGlobale.frx":1DA63
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image btSortConcUp 
      Height          =   255
      Index           =   6
      Left            =   9720
      Picture         =   "frmGlobale.frx":1DD4B
      Top             =   8400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image btSortConcUp 
      Height          =   255
      Index           =   5
      Left            =   8880
      Picture         =   "frmGlobale.frx":1E017
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image btSortConcDown 
      Height          =   255
      Index           =   4
      Left            =   7440
      Picture         =   "frmGlobale.frx":1E4BF
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image btSortConcUp 
      Height          =   255
      Index           =   4
      Left            =   7800
      Picture         =   "frmGlobale.frx":1EC07
      Top             =   8640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image btSortConcDown 
      Height          =   255
      Index           =   2
      Left            =   5160
      Picture         =   "frmGlobale.frx":1F2CF
      Top             =   8280
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Image btSortConcUp 
      Height          =   255
      Index           =   2
      Left            =   4920
      Picture         =   "frmGlobale.frx":200B7
      Top             =   8400
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Image btSortConcDown 
      Height          =   255
      Index           =   1
      Left            =   4200
      Picture         =   "frmGlobale.frx":20E23
      Top             =   8640
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Image btSortConcUp 
      Height          =   255
      Index           =   1
      Left            =   3960
      Picture         =   "frmGlobale.frx":21C93
      Top             =   8640
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Image btSortConcDown 
      Height          =   255
      Index           =   0
      Left            =   3840
      Picture         =   "frmGlobale.frx":22A87
      Top             =   8400
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image btSortConcUp 
      Height          =   255
      Index           =   0
      Left            =   3600
      Picture         =   "frmGlobale.frx":23077
      Top             =   8400
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   0
      Left            =   1560
      Picture         =   "frmGlobale.frx":235EB
      Top             =   120
      Width           =   855
   End
   Begin VB.Image btAnnulerDown 
      Height          =   375
      Left            =   120
      Picture         =   "frmGlobale.frx":2439F
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image btAnnulerUp 
      Height          =   375
      Left            =   0
      Picture         =   "frmGlobale.frx":24ED3
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image btNvCotiDown 
      Height          =   375
      Left            =   120
      Picture         =   "frmGlobale.frx":259AB
      Top             =   3120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image btNvCotiUp 
      Height          =   375
      Left            =   0
      Picture         =   "frmGlobale.frx":269C7
      Top             =   3120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image btFermerDown 
      Height          =   375
      Left            =   120
      Picture         =   "frmGlobale.frx":27993
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image btFermerUp 
      Height          =   375
      Left            =   0
      Picture         =   "frmGlobale.frx":2844B
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image btsupprDown 
      Height          =   255
      Left            =   0
      Picture         =   "frmGlobale.frx":28EA7
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmGlobale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public labs As New frmLabels

Private Sub adhid_Change()
    If adhid <> "" Then
        On Error GoTo pasdephoto
        photoAdh.Picture = LoadPicture(chemin + "photos\" + adhid + ".jpg")
    End If

Exit Sub
pasdephoto:
photoAdh.Picture = LoadPicture(none)
    'photoAdh.Picture = LoadPicture(chemin + "photos\" + "PICT0005.JPG")

End Sub

Private Sub adhNbChien_Change()
 
    lblNbChien.Caption = "chien"
    If adhNbChien <> "" Then
       If adhNbChien > 1 Then
           lblNbChien.Caption = "chiens"
       End If
    End If
End Sub

Private Sub btAddAdh_Click()
    Dim adah As New addAdh

    fGbl.Enabled = False
    adah.Show
End Sub


Private Sub btAddChien_Click()
    Dim adch As New addChien
    btFermerConc_Click
    adch.Show
    Enabled = False
    fMainForm.Frametop.Enabled = False
End Sub

Private Sub btAddChien_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
 btAddChien.Picture = btAddChiendown
End Sub

Private Sub btAddChien_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
btAddChien.Picture = btAddChienup
End Sub

Private Sub btAddConcours_Click()

  On Error GoTo UpdateErr
  
    If CheckConc() Then
        'init
        datPrimaryRSConcours.Recordset.AddNew
        
        'passage des valeurs dans la base
        concidChien = idChien(0)
        
        concDisc = drpTypDisc.ListIndex + 1
        concCat = drpTypCat.ListIndex + 1
        concApp = drpTypApp.ListIndex + 1
        concDate = tbConcDate(0)
        concLieu = tbConcLieu(0)
        concPoints = tbConcPoints(0)
        concClas = tbConcClas(0)
        
        datPrimaryRSConcours.Recordset.UpdateBatch adAffectAll
        datPrimaryRSConcours.Refresh
        Call selConcours(idChien(0))
    End If
      
  Exit Sub
UpdateErr:
    MsgBox "L'enregistrement n'a pas eu lieu. Vérifiez que ce chien n'a pas déjà de concours pour" + Chr$(10) + "le " + tbConcDate(0) + Chr$(10) + "à " + tbConcLieu(0) + Chr$(10) + "en " + drpTypDisc + " / " + drpTypCat
    datPrimaryRSConcours.Refresh
  'MsgBox Err.Description
End Sub

Function CheckConc() As Boolean
   On Error GoTo UpdateErr
   CheckConc = True
    
    If (StrComp(drpTypApp, "") = 0) Then
        drpTypApp.BackColor = &HF9D9D4
        CheckConc = False
        drpTypApp.SetFocus
    Else
        drpTypApp.BackColor = &HFFFFFF
    End If
    
    If (StrComp(drpTypDisc, "") = 0) Then
        drpTypDisc.BackColor = &HF9D9D4
        CheckConc = False
        drpTypDisc.SetFocus
    Else
        drpTypDisc.BackColor = &HFFFFFF
        If (StrComp(drpTypCat, "") = 0) Then
            drpTypCat.BackColor = &HF9D9D4
            CheckConc = False
            drpTypCat.SetFocus
        Else
            drpTypCat.BackColor = &HFFFFFF
        End If
    End If
    If (StrComp(tbConcLieu(0), "") = 0) Then
        tbConcLieu(0).BackColor = &HF9D9D4
        CheckConc = False
        tbConcLieu(0).SetFocus
    Else
        tbConcLieu(0).BackColor = &HFFFFFF
    End If
    
    If Not CheckConc Then
        err.Description = "Veuillez renseigner les champs sur-lignés"
        GoTo UpdateErr
    Else

        If Not IsDate(tbConcDate(0)) Then
            tbConcDate(0).BackColor = &HF9D9D4
            err.Description = "La date n'est pas une date valide"
            tbConcDate(0).SetFocus
            GoTo UpdateErr
        Else
            tbConcDate(0).BackColor = &HFFFFFF
        End If
        
        ' champ pas obligatoire
        If (tbConcPoints(0) <> "") Then
            tbConcPoints(0).BackColor = &HFFFFFF
            On Error GoTo typmismatchPts
            If (FormatNumber(tbConcPoints(0))) Then
            End If
            On Error GoTo UpdateErr
        End If
            
        ' champ pas obligatoire
        If (tbConcClas(0) <> "") Then
            tbConcClas(0).BackColor = &HFFFFFF
            On Error GoTo typmismatchClas
            If (FormatNumber(tbConcClas(0))) Then
            End If
            On Error GoTo UpdateErr
        End If
    End If
    
    
    Exit Function
typmismatchPts:
    tbConcPoints(0).BackColor = &HF9D9D4
    CheckConc = False
    err.Description = "Le nombre de points n'est pas valide"
    MsgBox err.Description
    tbConcPoints(0).SetFocus
    Exit Function
    
typmismatchClas:
    tbConcClas(0).BackColor = &HF9D9D4
    CheckConc = False
    err.Description = "Le classement n'est pas valide"
    MsgBox err.Description
    tbConcClas(0).SetFocus
    Exit Function
    
UpdateErr:
  MsgBox err.Description
  CheckConc = False
End Function



Private Sub btAddConcours_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAddConcours.Picture = btSaveDown.Picture
End Sub

Private Sub btAddConcours_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAddConcours.Picture = btSaveUp.Picture
End Sub

Private Sub btAffConc_Click(index As Integer)

        idc = idChien(index)
        
        If index > 0 Then 'on se replace sur le sélectionné
            dpl = index
            datPrimaryRSChi.Recordset.Move (dpl)
        End If
        'on cache les autres
        i = 1
        While i < maxChienAff
            frameChien(i).Visible = False
            i = i + 1
        Wend

        Call selConcours(idc)
End Sub

Private Sub btAffConc_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAffConc(index).Picture = btAffConcDown.Picture
End Sub

Private Sub btAffConc_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAffConc(index).Picture = btAffConcUp.Picture
End Sub

Private Sub btAnnulerCoti_Click()
    FrameNvCoti.Visible = False
    datPrimaryRSCoti.Refresh
    FrameCoti.Visible = True
    nvVersMontant.SetFocus
End Sub

Sub btFermerConc_Click()
        drpTypApp.BackColor = &HFFFFFF
        drpTypCat.BackColor = &HFFFFFF
        drpTypDisc.BackColor = &HFFFFFF
        tbConcLieu(0).BackColor = &HFFFFFF
        nvVersDate.BackColor = &HFFFFFF
        tbConcPoints(0).BackColor = &HFFFFFF
        tbConcClas(0).BackColor = &HFFFFFF

        FrameConcours.Visible = False
            FrameConcoursF.Visible = False
            btPevChien.Visible = False
            btNextChien.Visible = False
        
        Call affchienadh
        FrameCoti.Visible = True
        
End Sub

Private Sub btFermerConc_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btFermerConc.Picture = btPtFermerD.Picture
End Sub

Private Sub btFermerConc_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btFermerConc.Picture = btPtFermerU.Picture
End Sub

Private Sub btFermerCoti_Click()
    If saveCoti() Then
        FrameNvCoti.Visible = False
        Call selCoti(adhid)
    End If
End Sub

Private Function saveCoti() As Boolean

    On Error GoTo UpdateErr
    saveCoti = True
    
   'Montant Cotisation
    nvCotiMnt.SetFocus
    If (StrComp(nvCotiMnt, "") = 0) Then
        nvCotiMnt.BackColor = &HF9D9D4
        err.Description = "Le montant n'est pas valide"
        GoTo UpdateErr
    Else
        nvCotiMnt.BackColor = &HFFFFFF
        On Error GoTo typmismatch
        If (FormatCurrency(nvCotiMnt)) Then
        End If
        On Error GoTo UpdateErr
    End If
    
    'Date d'échéance
    nvCotiEch.SetFocus
    If Not IsDate(nvCotiEch) Then
            nvCotiEch.BackColor = &HF9D9D4
            err.Description = "La date d'échéance n'est pas une date valide"
            GoTo UpdateErr
    Else
        nvCotiEch.BackColor = &HFFFFFF
        If (DateValue(nvCotiEch) <= date) Then
            nvCotiEch.BackColor = &HF9D9D4
            err.Description = "L'échéance est dépassée"
            GoTo UpdateErr
        Else
            nvCotiEch.BackColor = &HFFFFFF
        End If
    End If


    'sauve
    datPrimaryRSCoti.Recordset.UpdateBatch adAffectAll
    datPrimaryRSCoti.Refresh
    datPrimaryRSCoti.Refresh
    Call selCoti(adhid)
    
    Exit Function
typmismatch:
    montantCoti.BackColor = &HF9D9D4
    saveCoti = False
    MsgBox "Le montant n'est pas valide"
    Exit Function
UpdateErr:
    MsgBox err.Description
    saveCoti = False
End Function







Private Sub btModAdh_Click()
    Dim diamadh As New modAdh
    diamadh.Show
    Enabled = False
    fMainForm.Frametop.Enabled = False
End Sub
Private Sub btModAdh_Mouseup(Button As Integer, Shift As Integer, x As Single, Y As Single)
        btModAdh.Picture = btModup.Picture
End Sub
Private Sub btModAdh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        btModAdh.Picture = btModdown.Picture
End Sub


Private Sub btNvVers_Click()
  On Error GoTo UpdateErr
  
    If datPrimaryRSCoti.Recordset.RecordCount < 1 Then
        'pas de coti
        MsgBox "Veuillez entrer une cotisation avant un versement"
    Else
        If CheckVers() Then
            'init
            datPrimaryRSVers.Recordset.AddNew
            payeurid = adhid
            dateVers = nvVersDate
            montantVers = nvVersMontant
            typPay = drpTypPay
            
            datPrimaryRSVers.Recordset.UpdateBatch adAffectAll
            datPrimaryRSVers.Refresh
            datPrimaryRSVers.Refresh
            Call selVers(adhid, EcheancePrec, echeance)
        
        End If
    End If
  Exit Sub
UpdateErr:
  MsgBox err.Description
  
End Sub

'Concours Chien
Private Sub btPevChien_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btPevChien.Picture = btScrollUpDown.Picture
End Sub
Private Sub btPevChien_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If datPrimaryRSChi.Recordset.AbsolutePosition > 1 And datPrimaryRSChi.Recordset.AbsolutePosition <> adPosBOF Then
        MousePointer = vbHourglass
        datPrimaryRSChi.Recordset.MovePrevious
        Call selConcours(idChien(0))
        MousePointer = vbArrow
    End If
    btPevChien.Picture = btScrollUpUp.Picture
End Sub
Private Sub btNextChien_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNextChien.Picture = btScrolldowndown.Picture
End Sub
Private Sub btNextChien_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If datPrimaryRSChi.Recordset.AbsolutePosition < datPrimaryRSChi.Recordset.RecordCount And datPrimaryRSChi.Recordset.AbsolutePosition > 0 Then
        MousePointer = vbHourglass
        datPrimaryRSChi.Recordset.MoveNext
        Call selConcours(idChien(0))
        MousePointer = vbArrow
    End If
    btNextChien.Picture = btScrolldownUp.Picture
End Sub


'rech CHIEN
Private Sub btNextChienRech_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If datPrimaryRSChi.Recordset.AbsolutePosition < datPrimaryRSChi.Recordset.RecordCount And datPrimaryRSChi.Recordset.AbsolutePosition > 0 Then
        datPrimaryRSChi.Recordset.MoveNext
    End If
     btNextChienRech = btScrollRightUp.Picture
End Sub

Private Sub btNextChienRech_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNextChienRech = btScrollRightDown.Picture
End Sub

Private Sub btPevChienRech_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btPevChienRech = btScrollLeftDown.Picture
End Sub

Private Sub btPevChienRech_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If datPrimaryRSChi.Recordset.AbsolutePosition > 1 And datPrimaryRSChi.Recordset.AbsolutePosition > 0 Then
        datPrimaryRSChi.Recordset.MovePrevious
    End If
     btPevChienRech = btScrollLeftUp.Picture
End Sub



'rech ADH
Private Sub btPrevAdh_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If datPrimaryRS.Recordset.AbsolutePosition > 1 And datPrimaryRS.Recordset.AbsolutePosition > 0 Then
        datPrimaryRS.Recordset.MovePrevious
    End If
    btPrevAdh.Picture = btScrollLeftUp.Picture
End Sub
Private Sub btPrevAdh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btPrevAdh.Picture = btScrollLeftDown.Picture
End Sub
Private Sub btNextAdh_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If datPrimaryRS.Recordset.AbsolutePosition < datPrimaryRS.Recordset.RecordCount And datPrimaryRS.Recordset.AbsolutePosition > 0 Then
        datPrimaryRS.Recordset.MoveNext
    End If
    btNextAdh = btScrollRightUp.Picture
End Sub
Private Sub btNextAdh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
     btNextAdh.Picture = btScrollRightDown.Picture
End Sub

Private Sub btNvCoti_Click()
    Dim adcoti As New addCoti
    fMainForm.Frametop.Enabled = False
    fGbl.Enabled = False
    adcoti.Show
End Sub


Private Function CheckVers() As Boolean

   On Error GoTo UpdateErr
    CheckVers = True
    
   'Montant Versement
    nvVersMontant.SetFocus
    If (StrComp(nvVersMontant, "") = 0) Then
        nvVersMontant.BackColor = &HF9D9D4
        err.Description = "Le montant n'est pas valide"
        GoTo UpdateErr
    Else
        nvVersMontant.BackColor = &HFFFFFF
        On Error GoTo typmismatch
        If (FormatCurrency(nvVersMontant)) Then
        End If
        On Error GoTo UpdateErr
    End If
    
    'Date Versement

    nvVersDate.SetFocus
  
    If Not IsDate(nvVersDate) Then
            nvVersDate.BackColor = &HF9D9D4
            err.Description = "La date d'échéance n'est pas une date valide"
            GoTo UpdateErr
    Else
        nvVersDate.BackColor = &HFFFFFF
        
        'If (DateValue(nvVersDate) > Date) Then ' Versement différé
        '    nvVersDate.BackColor = &HF9D9D4
            'Err.Description = "Un tien vaut mieux que deux tu l'auras:" + Chr$(10) + "Veuillez saisir une date passée"
            'GoTo UpdateErr
        'Else
            'dateVers.BackColor = &HFFFFFF
        'End If
    End If


    Exit Function
typmismatch:
    montantVers.BackColor = &HF9D9D4
    CheckVers = False
    MsgBox "Le montant n'est pas valide"
    Exit Function
UpdateErr:
    MsgBox err.Description
    CheckVers = False
End Function

Private Sub btRechChien_Click()
    imgBienvenu.Visible = False
    btAddChien.Visible = False
    FrameCoti.Visible = False
    FrameConcours.Visible = False
            FrameConcoursF.Visible = False
            btPevChien.Visible = False
            btNextChien.Visible = False
    FrameAdh.Visible = False
    btSupAdh.Visible = False
    btModAdh.Visible = False
    VScrollChiens.Visible = False
    'Récupération de tous les chiens
    datPrimaryRSChi.RecordSource = "select MaitreID,id,Nom,Race,Sexe,LOF,Tatouage,DateNaiss,no_CNEA,no_CUN from Chiens order by Nom"
    datPrimaryRSChi.Refresh
    If datPrimaryRSChi.Recordset.RecordCount > 0 Then
        frameChien(0).Visible = True
        btSupChien(0).Visible = False
        modChien(0).Visible = False
        btModchien.Visible = False
        btAffConc(0).Visible = False
        For i = 1 To maxChienAff - 1
            frameChien(i).Visible = False
        Next
        totChiens = datPrimaryRSChi.Recordset.RecordCount
        rechChien = ""
        rechChien.SetFocus
    Else
        MsgBox "Pas de chien"
    End If
    
End Sub



Private Sub btScrollDownEch_Click()
    MousePointer = vbHourglass
    If datPrimaryRSCoti.Recordset.RecordCount > 0 Then
        'sauve Echéance
        tempEch = DateValue(echeance)
        With datPrimaryRSCoti.Recordset
            .MoveNext
            If .EOF Then 'on ETAIT sur la Dernière,
                .MoveLast
            Else '
                EcheancePrec = tempEch
                .MoveNext
                If .EOF Then 'on est sur la Dernière,
                    .MoveLast
                    Call selVers(adhid, EcheancePrec, CStr(dateMax))
                Else
                    .MovePrevious
                    Call selVers(adhid, EcheancePrec, echeance)
                End If
            End If
        End With
    End If
    
    MousePointer = vbArrow
End Sub
Private Sub btScrollDownEch_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btScrollDownEch.Picture = btScrolldowndown.Picture
End Sub
Private Sub btScrollDownEch_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btScrollDownEch.Picture = btScrolldownUp.Picture
End Sub

Private Sub btScrollUpEch_Click()
    MousePointer = vbHourglass
    If datPrimaryRSCoti.Recordset.RecordCount > 0 Then
 
        With datPrimaryRSCoti.Recordset
              .MovePrevious
              If .BOF Then 'on est sur la 1ère, on fait rien que s'y remètre
                  .MoveFirst
                  
              Else 'on regared si il y a un coti précédante
                  .MovePrevious
                  If .BOF Then 'Y en a pas
                      .MoveFirst
                      EcheancePrec = ""
                      Call selVers(adhid, CStr(dateMin), echeance)
                  Else ' y en a une
                      EcheancePrec = echeance
                      .MoveNext
                      Call selVers(adhid, EcheancePrec, echeance)
                  End If
          End If
        End With
    End If
    btScrollUpEch.Picture = btScrollUpUp.Picture
    MousePointer = vbArrow
End Sub
Private Sub btScrollUpEch_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btScrollUpEch.Picture = btScrollUpDown.Picture
End Sub
Private Sub btScrollUpEch_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btScrollUpEch.Picture = btScrollUpUp.Picture
End Sub




Private Sub btSelChien_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btSelChien.Picture = btRappeldown.Picture
End Sub
Private Sub btSelChien_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btSelChien.Picture = btRappelup.Picture
End Sub
Private Sub btSelAdh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btSelAdh.Picture = btRappeldown.Picture
End Sub

Private Sub btSelAdh_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btSelAdh.Picture = btRappelup.Picture
End Sub

Private Sub btSortConc_Click(index As Integer)
    
    If datPrimaryRSConcours.Recordset.RecordCount > 1 Then
        MousePointer = vbHourglass
        btSortConc(index).Picture = btSortConcDown(index).Picture
       
        VScrollConcours = 0
        refrechConcours (1)
    
    
        grdConc.Col = index
        
        grdConc.Row = 0
        tri = CStr(grdConc.Columns(index).Caption)
        preVal = grdConc
        grdConc.Row = datPrimaryRSConcours.Recordset.RecordCount - 1 'grdConc.VisibleRows - 1
        derVal = grdConc
                
        Select Case index
            Case 0: ' date
                
                If (DateValue(derVal) > DateValue(preVal)) Then
                    ascdesc = " desc"
                Else
                    ascdesc = " asc"
                End If
       
                tripar = ", Lieu,Discipline,Categorie,Appreciation"
            
            Case 1: ' lieu
                If (StrComp(derVal, preVal) > 0) Then
                    ascdesc = " desc"
                Else
                    ascdesc = " asc"
                End If
         
                tri2 = " ,Discipline,Categorie,Appreciation,Date"
                
            Case 2: ' disc 'label !!!
            ' If (StrComp(tbConcLieu(nbLignes), tbConcLieu(1)) > 0) Then
                If (derVal > preVal) Then
                    ascdesc = " desc"
                Else
                    ascdesc = " asc"
                End If
                tripar = " ,Discipline,Categorie,Appreciation,Lieu, Date"
                 
            Case 4: '  appréciation ! label mais par ordre
                If (derVal > preVal) Then
                    ascdesc = " desc"
                Else
                    ascdesc = " asc"
                End If
                tripar = " ,Appreciation,Discipline,Categorie,Lieu,Date"
                        
            'Case 5 ' points
            '    If (derVal > preVal) Then
            '        ascdesc = " desc"
            '    Else
            '        ascdesc = " asc"
            '    End If
            '    tripar = " ,Points,Discipline,Categorie,Lieu,Date"
            'Case 6: ' classement
            '    If (derVal > preVal) Then
            '        ascdesc = " desc"
            '    Else
            '        ascdesc = " asc"
            '    End If
            '    tripar = " ,Classement,Discipline,Categorie,Lieu,Date"
                        
        End Select
        If concsixmois Then
            conditiondate = CStr(" and Date>" + datsql(DateAdd("m", -6, date)))
        Else
            conditiondate = ""
        End If
    
        datPrimaryRSConcours.RecordSource = "select ChienID,Date,Lieu,Discipline,Categorie,Appreciation,Points,Classement from Concours  where ChienID =" + idChien(0) + conditiondate + " Order by " + tri + ascdesc + tripar
        datPrimaryRSConcours.Refresh
        Call refrechConcours(VScrollConcours)
        btSortConc(index).Picture = btSortConcUp(index).Picture
        MousePointer = vbArrow
    End If
End Sub
Private Sub btSortConc_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If datPrimaryRSConcours.Recordset.RecordCount > 0 Then
        btSortConc(index).Picture = btSortConcDown(index).Picture
        
    End If
End Sub


Private Sub btSupAdh_Click()
    Dim dia As New dlgDelAdh
    dia.index = index
    dia.Show
End Sub
Private Sub btSupAdh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btSupAdh.Picture = btsupprDown.Picture
End Sub
Private Sub btSupAdh_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btSupAdh.Picture = btsupprUp.Picture
End Sub

Private Sub btSupChien_Click(index As Integer)
Dim dia As New dlgDelChien
    'btFermerConc_Click
    dia.index = index
    dia.Show
End Sub



Private Sub btSupChien_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    btSupChien(index).Picture = btsupprDown.Picture
End Sub
Private Sub btSupChien_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    btSupChien(index).Picture = btsupprUp.Picture
End Sub



Private Sub concsixmois_Click()
    Call selConcours(idChien(0))
End Sub

Private Sub dateNaiss_Change(index As Integer)
        ' age(index) = DateDiff("yyyy", DateValue(dateNaiss(index)), Date)
        If IsDate(dateNaiss(index)) Then
            age(index) = Year(date) - Year(DateValue(dateNaiss(index)))
            If (Month(date) < Month(DateValue(dateNaiss(index)))) Or (Month(date) = Month(DateValue(dateNaiss(index))) And Day(date) < Day(DateValue(dateNaiss(index)))) Then
                age(index) = age(index) - 1
            End If
        End If
End Sub


Private Sub datPrimaryRSChi_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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




Private Sub drpTypApp_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KKenter Then tbConcPoints(0).SetFocus
End Sub

Private Sub drpTypCat_Click()

    If drpTypCat.ListIndex + 1 = drpTypCat.ListCount Then
        Dim addica As New addCat
        addica.Disc = drpTypDisc
        addica.idDisc = drpTypDisc.ListIndex + 1
        addica.idCat = drpTypCat.ListIndex + 1
        addica.Show
    Else
        concCat = drpTypCat.ListIndex + 1
       
    End If
End Sub

Private Sub drpTypCat_GotFocus()
    If (drpTypDisc = "") Then
        drpTypDisc.SetFocus
        drpTypCat.Enabled = False
    End If
End Sub




Private Sub drpTypCat_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KKenter Then drpTypApp.SetFocus
End Sub

Private Sub drpTypDisc_Click()
    If (drpTypDisc.ListIndex + 1 = drpTypDisc.ListCount) Then
    'nouvelle discipline
        Dim adDici As New addDisc
        adDici.id = drpTypDisc.ListCount
        adDici.Show
    Else
        'init les catégories
        concDisc = drpTypDisc.ListIndex + 1
        Call InitDropTypCat
        drpTypCat.Enabled = True

        
    End If
End Sub

Private Sub drpTypDisc_KeyUp(KeyCode As Integer, Shift As Integer)
       If KeyCode = KKenter Then drpTypCat.SetFocus
End Sub

Private Sub drpTypPay_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then nvVersDate.SetFocus
End Sub

Private Sub Form_Activate()
    If nbWin > 1 Then fGbl.ZOrder (1)
End Sub

Private Sub Form_Initialize()
   ' WindowState = vbMaximized
    
    ChDir (chemin)

End Sub


Private Sub Form_Load()
    btModAdh.Picture = btModup.Picture

    

    btNextAdh.Visible = False
    btPrevAdh.Visible = False
    btSelAdh.Visible = False
    btNextChienRech.Visible = False
    btPevChienRech.Visible = False
    btSelChien.Visible = False
    totChiens.Visible = False
    lblTotChien.Visible = False
    totAdh.Visible = False
    lblTotAdh.Visible = False
            
            
    btAddChien.Visible = False
    VScrollChiens.Enabled = False
    VScrollChiens.Visible = False
    btNextChien.Visible = False
    btPevChien.Visible = False
    FrameAdh.Visible = False
    btSupAdh.Visible = False
    btModAdh.Visible = False
    
    FrameCHIENS.Visible = False
    For i = 0 To maxChienAff - 1
        frameChien(i).Visible = False
    Next
    FrameCoti.Visible = False
    
    FrameConcours.Visible = False
        FrameConcoursF.Visible = False
        btPevChien.Visible = False
        btNextChien.Visible = False

    Call InitDropTypPay
    Call InitDropTypDisc
    Call InitDropTypApp
    
End Sub
Private Sub InitDropTypPay()
    Call drpTypPay.Clear
    datPrimaryRSTypPay.Recordset.MoveFirst
    While (Not (datPrimaryRSTypPay.Recordset.EOF))
        drpTypPay.AddItem (labelTypPay)
        datPrimaryRSTypPay.Recordset.MoveNext
    Wend
    datPrimaryRSTypPay.Recordset.MovePrevious
    'drpTypPay.ListIndex = 0
End Sub
 Sub InitDropTypDisc()
   
    Call drpTypDisc.Clear
    labs.datlstDisc.Recordset.MoveFirst
    While (Not (labs.datlstDisc.Recordset.EOF))
        drpTypDisc.AddItem (labs.labelDisc)
        labs.datlstDisc.Recordset.MoveNext
    Wend
    drpTypDisc.AddItem ("Nv. Discipline")
    labs.datlstDisc.Recordset.MovePrevious
    'drpTypPay.ListIndex = 0
End Sub
 Sub InitDropTypCat()
    
    Call drpTypCat.Clear
    If (drpTypDisc <> "") Then
        labs.datlstCat.Recordset.MoveFirst
        While (Not (labs.datlstCat.Recordset.EOF))
            If (labs.idDiscCat = concDisc) Then
                drpTypCat.AddItem (labs.labelCat)
            End If
            labs.datlstCat.Recordset.MoveNext
        Wend
        drpTypCat.AddItem ("Nv. Catégorie")
        labs.datlstCat.Recordset.MovePrevious

        'drpTypPay.ListIndex = 0
        drpTypCat.Enabled = True
    End If
    
End Sub
Private Sub InitDropTypApp()
   
    Call drpTypApp.Clear
    labs.datlstApp.Recordset.MoveFirst
    While (Not (labs.datlstApp.Recordset.EOF))
        drpTypApp.AddItem (labs.labelApp)
        labs.datlstApp.Recordset.MoveNext
    Wend
    labs.datlstCat.Recordset.MovePrevious
    'drpTypPay.ListIndex = 0
End Sub
Private Sub InitDropRaces()
    
    Call lstRaces.Clear
    datPrimaryRSRace.Recordset.MoveFirst
    While (Not (datPrimaryRSRace.Recordset.EOF))
        lstRaces.AddItem (labelRac)
     datPrimaryRSRace.Recordset.MoveNext
    Wend

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
  Unload labs
End Sub
Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
End Sub
Private Sub datPrimaryRSChi_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
    cptChien = datPrimaryRSChi.Recordset.AbsolutePosition & " /" & datPrimaryRSChi.Recordset.RecordCount
    
    If (idChien(0) = "") Then
        labelSex(0) = ""
        age(0) = ""
       ' datPrimaryRSRace.Recordset.MoveLast
       ' datPrimaryRSRace.Recordset.MoveNext
    Else
        'If sex(0) = 0 Then
        '    labelSex(0) = "Mâle"
        'ElseIf sex(0) = 1 Then
        '    labelSex(0) = "Femelle"
        'End If
        
       ' If IsDate(dateNaiss(0)) Then
        '    age(0) = Year(Date) - Year(DateValue(dateNaiss(0)))
        '    If (Month(Date) < Month(DateValue(dateNaiss(0)))) Or (Month(Date) = Month(DateValue(dateNaiss(0))) And Day(Date) < Day(DateValue(dateNaiss(0)))) Then
       '         age(0) = age(0) - 1
      '      End If
       ' End If
        

    End If
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


















Private Sub gridVers_LostFocus()
           datPrimaryRSVers.Recordset.UpdateBatch adAffectAll
End Sub

Private Sub idRaceChien_Change(index As Integer)
    datPrimaryRSRace.Recordset.MoveFirst
    While (idRace <> idRaceChien(index) And datPrimaryRSRace.Recordset.AbsolutePosition < datPrimaryRSRace.Recordset.RecordCount)
        datPrimaryRSRace.Recordset.MoveNext
    Wend
    '''        datPrimaryRSRace.Recordset.Move (idRaceChien(index) - 1)
End Sub






Private Sub lstGlb_Click()
    Dim d As New fListes
    d.Show
    d.typLst.ListIndex = 0
    
End Sub



Private Sub modChien_Click(index As Integer)
    Dim diam As New modChien
    diam.id = idChien(index)
    diam.Show
    Enabled = False
    fMainForm.Frametop.Enabled = False
End Sub
Private Sub modChien_Mouseup(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
        modChien(index).Picture = btModup.Picture
End Sub
Private Sub modChien_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
        modChien(index).Picture = btModdown.Picture
End Sub


Private Sub nomAdh_Change()
    NomPrenom = UCase(nomAdh) + " " + prenomAdh
End Sub



Private Sub nvVersDate_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then btNvVers_Click
End Sub

Private Sub nvVersMontant_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then drpTypPay.SetFocus
End Sub

Private Sub nvVersMontant_LostFocus()
    On Error GoTo err
    
        nvVersMontant = FormatNumber(nvVersMontant, 2)
    
  Exit Sub
err:
 nvVersMontant = ""
End Sub

Private Sub prenomAdh_Change()
    NomPrenom = UCase(nomAdh) + " " + prenomAdh
End Sub


Private Sub rechAdh_Change()
    Dim nbrc As Integer
    
    If (rechAdh <> "") Then
        If datPrimaryRS.Recordset.RecordCount < 1 Then
            MsgBox "Pas d'adhérent"
            rechAdh = ""
        Else
            If rech_nom(rechAdh, NomPrenom, datPrimaryRS) Then
                Call btSelAdh_Click
            End If
        End If
    End If
End Sub


Private Sub rechAdh_GotFocus()
        rechAdh = ""
        btNextChienRech.Visible = False
        btPevChienRech.Visible = False
        btSelChien.Visible = False
        totChiens.Visible = False
        lblTotChien.Visible = False
           
            
        imgBienvenu.Visible = False
        VScrollChiens.Visible = False
        FrameCoti.Visible = False
        FrameConcours.Visible = False
            FrameConcoursF.Visible = False
            btPevChien.Visible = False
            btNextChien.Visible = False
        btAddChien.Visible = False
        btSupAdh.Visible = False
        btModAdh.Visible = False
        FrameCHIENS.Visible = False
        For i = 0 To maxChienAff - 1
            frameChien(i).Visible = False
        Next
        
         'Récupération de tous les Adh
        datPrimaryRS.RecordSource = "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Adresse,CodePostal,Ville,Nb_Chiens,No_SCRA from Proprietaire order by Nom,prenom"
        datPrimaryRS.Refresh
        totAdh = datPrimaryRS.Recordset.RecordCount
        If datPrimaryRS.Recordset.RecordCount > 0 Then
        
            btNextAdh.Visible = True
            btPrevAdh.Visible = True
            btSelAdh.Visible = True
            totAdh.Visible = True
            lblTotAdh.Visible = True
    
            FrameAdh.Visible = True
        Else
        End If
End Sub

Private Sub rechAdh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call btSelAdh_Click
    End If
End Sub
Public Sub btSelAdh_Click()
    Dim nbrc As Integer
            btNextAdh.Visible = False
            btPrevAdh.Visible = False
            btSelAdh.Visible = False
            totAdh.Visible = False
            lblTotAdh.Visible = False
            
            
    Call selAdh(adhid)
    nbrc = selChienAdh(adhid)
End Sub
Private Sub btRechAdh_Click()
    imgBienvenu.Visible = False
    VScrollChiens.Visible = False
    FrameCoti.Visible = False
    FrameConcours.Visible = False
            FrameConcoursF.Visible = False
            btPevChien.Visible = False
            btNextChien.Visible = False
    btAddChien.Visible = False
    btSupAdh.Visible = False
    btModAdh.Visible = False
    FrameCHIENS.Visible = False
    For i = 0 To maxChienAff - 1
        frameChien(i).Visible = False
    Next

     'Récupération de tous les Adh
    datPrimaryRS.RecordSource = "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Adresse,CodePostal,Ville,Nb_Chiens,No_SCRA from Proprietaire order by Nom"
    datPrimaryRS.Refresh

    If datPrimaryRS.Recordset.RecordCount > 0 Then
    
        FrameAdh.Visible = True
        totAdh = datPrimaryRS.Recordset.RecordCount
        rechAdh = ""
        rechAdh.SetFocus
    Else
        MsgBox "Pas d'adhérent"
    End If
   
End Sub



Sub affchienadh()
    
    btAddChien.Visible = True
    VScrollChiens.Visible = False
    VScrollChiens.Enabled = False
    VScrollChiens = 0
    VScrollChiens.Max = 0
    If (datPrimaryRSChi.Recordset.RecordCount > 0) Then
        FrameCHIENS.Visible = True
        ' on se garde le premier de coté
        datPrimaryRSChi.Recordset.MoveFirst
        frameChien(0).Visible = True
        btSupChien(0).Visible = True
        modChien(0).Visible = True
        btAffConc(0).Visible = True
        i = 1
        'on affiche les autres
        While i < datPrimaryRSChi.Recordset.RecordCount And i < maxChienAff

            
            datPrimaryRSChi.Recordset.MoveNext
            idChien(i) = idChien(0)
            nomChien(i) = nomChien(0)
            labelSex(i) = labelSex(0)
            tatou(i) = tatou(0)
            chienlof(i) = chienlof(0)
            dateNaiss(i) = dateNaiss(0)
            labelRace(i) = labelRace(0)
            cnea(i) = cnea(0)
            cun(i) = cun(0)
            
            frameChien(i).Visible = True
           
            i = i + 1
        Wend
        datPrimaryRSChi.Recordset.MoveFirst
        'masque les autre chiens
        While i < maxChienAff
            frameChien(i).Visible = False
            i = i + i
        Wend
         
        If datPrimaryRSChi.Recordset.RecordCount > maxChienAff Then ' plus de chiens que l'on peut en afficher
            'affiche bouttons
                VScrollChiens.Visible = True
                VScrollChiens.Enabled = True
                VScrollChiens.Max = datPrimaryRSChi.Recordset.RecordCount - maxChienAff
        End If
    Else
        'masque tous les chiens
        FrameCHIENS.Visible = False
        For i = 1 To maxChienAff
            frameChien(i - 1).Visible = False
        Next
    End If
    FrameCoti.Visible = True
End Sub

Sub selAdh(ida)
    
    attend = 0
    datPrimaryRS.RecordSource = "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Adresse,CodePostal,Ville,Nb_Chiens,No_SCRA from Proprietaire where id=" + ida
    datPrimaryRS.Refresh
    While datPrimaryRS.Recordset.RecordCount < 1 And attend < 500
        datPrimaryRS.RecordSource = "select id,Nom,Prenom,Tel_Domicile,Tel_Portable,Adresse,CodePostal,Ville,Nb_Chiens,No_SCRA from Proprietaire where id=" + ida
        datPrimaryRS.Refresh
        attend = attend + 1
    Wend
    If attend > 1 Then MsgBox attend
        
    FrameAdh.Visible = True
    btSupAdh.Visible = True
    btModAdh.Visible = True
     selCoti (ida)
End Sub
''''' COTI ''''
Sub selCoti(ida As Integer)

    nvVersMontant = ""
    nvVersDate = ""
    drpTypPay = ""
    EcheancePrec = ""
    Solde = "0.00"
    TotVers = "0.00"
    datPrimaryRSCoti.RecordSource = "select AdhID,Cotisation,Echeance from EcheanceCotis where AdhID=" + CStr(ida) + " Order by Echeance"
    datPrimaryRSCoti.Refresh
    datPrimaryRSCoti.Refresh
     
      
      
    FrameCoti.Visible = True
    FrameConcours.Visible = False
        FrameConcoursF.Visible = False
        btPevChien.Visible = False
        btNextChien.Visible = False
        
    If datPrimaryRSCoti.Recordset.RecordCount > 0 Then
        
        datPrimaryRSCoti.Recordset.MoveLast
        
        's'il y en a avant, on l'affiche
        If (datPrimaryRSCoti.Recordset.RecordCount > 1) Then
            datPrimaryRSCoti.Recordset.MovePrevious
            EcheancePrec = echeance
            datPrimaryRSCoti.Recordset.MoveNext
            Call selVers(ida, EcheancePrec, CStr(dateMax))
        Else
            Call selVers(ida, CStr(dateMin), CStr(dateMax))
        End If
    Else 'pas de coti
        
        datPrimaryRSVers.RecordSource = "select AdhID,Date,Montant,TypePayement from Versements where (AdhID=" + CStr(ida) + ") order by Date"
        datPrimaryRSVers.Refresh
    End If
    nvVersMontant.SetFocus
    

End Sub

Sub selVers(ida As Integer, dateDeb As String, dateFin As String)
    
    'conditionDate = " and dateDeb > TO_DATE('"&2002-11-01 00:00:00','YYYY-MM-DD HH24:MI:SS')"
    If dateDeb = "" Then dateDeb = dateMin
    If dateFin = "" Then dateFin = dateMax
    conditiondate = " and date>" & datsql(dateDeb) & " and date<=" & datsql(dateFin)
    
    datPrimaryRSVers.RecordSource = "select AdhID,Date,Montant,TypePayement from Versements where (AdhID=" + CStr(ida) + conditiondate + ") order by Date"
    datPrimaryRSVers.Refresh
     
    
    'solde est total
    iTotVers = 0#
    iVers = 0#
    iSolde = 0#
    iSolde = -montantCoti '0#
    If (datPrimaryRSVers.Recordset.RecordCount > 0) Then

      iVers = 0#
      gridVers.Col = 0
      gridVers.Row = 0
      iSolde = iTotVers - montantCoti
       If gridVers.Row = 0 Then
          While gridVers.Row < gridVers.VisibleRows - 1
              iVers = gridVers
              iTotVers = iTotVers + iVers
              gridVers.Row = gridVers.Row + 1
          Wend
          iTotVers = iTotVers + gridVers
          iSolde = iTotVers - montantCoti
      End If
    End If
    TotVers = FormatNumber(iTotVers, 2)
    Solde = FormatNumber(iSolde, 2)
    
    nvVersDate = ""
    nvVersMontant = ""
    drpTypPay.ListIndex = -1
    nvVersMontant.SetFocus
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
    btFermer.Picture = btFermerDown.Picture
End Sub
Private Sub btFermer_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btFermer.Picture = btFermerUp.Picture
End Sub
Private Sub btNvRace_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvRace.Picture = btNvRaceDown.Picture
End Sub
Private Sub btNvRace_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvRace.Picture = btNvRaceUp.Picture
End Sub
Private Sub btFermerChien_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btFermerChien.Picture = btFermerDown.Picture
End Sub
Private Sub btFermerChien_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
   btFermerChien.Picture = btFermer.Picture
End Sub
Private Sub btNvChienChien_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvChienChien.Picture = btNvChienDown.Picture
End Sub
Private Sub btNvChienChien_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvChienChien.Picture = btNvChienUp.Picture
End Sub
Private Sub btNvCoti_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvCoti.Picture = btNvCotiDown.Picture
End Sub
Private Sub btNvCoti_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvCoti.Picture = btNvCotiUp.Picture
End Sub
Private Sub btNvVers_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btNvVers.Picture = btSaveDown.Picture
End Sub
Private Sub btNvVers_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
     btNvVers.Picture = btSaveUp.Picture
End Sub
Private Sub btFermerCoti_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btFermerCoti.Picture = btFermerDown.Picture
End Sub
Private Sub btFermerCoti_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btFermerCoti.Picture = btFermerUp.Picture
End Sub
Private Sub btFermerVers_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btFermerVers = btFermerDown.Picture
End Sub
Private Sub btFermerVers_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btFermerVers.Picture = btFermerUp.Picture
End Sub


Private Sub btAnnulerAdh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerAdh.Picture = btAnnulerDown.Picture
End Sub
Private Sub btAnnulerAdh_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerAdh.Picture = btAnnulerUp.Picture
End Sub
Private Sub btAnnulerChien_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerChien.Picture = btAnnulerDown.Picture
End Sub
Private Sub btAnnulerChien_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerChien.Picture = btAnnulerUp.Picture
End Sub
Private Sub btAnnulerCoti_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerCoti.Picture = btAnnulerDown.Picture
End Sub
Private Sub btAnnulerCoti_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerCoti.Picture = btAnnulerUp.Picture
End Sub
Private Sub btAnnulerVers_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerVers.Picture = btAnnulerDown.Picture
End Sub
Private Sub btAnnulerVers_MouseUP(Button As Integer, Shift As Integer, x As Single, Y As Single)
    btAnnulerVers.Picture = btAnnulerUp.Picture
End Sub



Private Function rech_nom(rech As Object, champ As Object, datrec As Adodc) As Boolean

    On Error GoTo err
    rech_nom = False
    memInd = datrec.Recordset.AbsolutePosition
    datrec.Recordset.MoveFirst
    
    'MsgBox KeyCode

    While (StrComp(LCase(rech), LCase(champ)) > 0) And (datrec.Recordset.AbsolutePosition < datrec.Recordset.RecordCount)
       datrec.Recordset.MoveNext
    Wend
    
    If (InStr(LCase(champ), LCase(rech)) = 1) Then
        If (datrec.Recordset.AbsolutePosition = datrec.Recordset.RecordCount) Then
             rech_nom = True
        Else
            datrec.Recordset.MoveNext
            If (InStr(LCase(champ), LCase(rech)) = 1) Then
                'Continue
            Else
                rech_nom = True
                rech = ""
            End If
            datrec.Recordset.MovePrevious
        End If
    Else
    
        'MsgBox "Aucun nom ne commence par """ + rech + """"
        'rech = Left(rech.Text, Len(rech.Text) - 1)
        'rech.SelStart = Len(rech.Text)
        'rech.SetFocus
        'Call rech_nom(rech, champ, datrec)
    End If
  
    
  Exit Function
err:
  MsgBox err.Description
End Function

Function selChienAdh(mtrid) As Integer
    datPrimaryRSChi.RecordSource = "select MaitreID,id,Nom,Race,Sexe,LOF,Tatouage,DateNaiss,no_CNEA,no_CUN from Chiens where Maitreid =" + mtrid + "  order by Nom"
    datPrimaryRSChi.Refresh
    selChienAdh = datPrimaryRSChi.Recordset.RecordCount
    FrameCHIENS.Visible = True
    Call affchienadh
End Function

Private Sub selConcours(idc)

    VScrollChiens.Visible = False
    FrameCoti.Visible = False
        
     btPevChien.Visible = datPrimaryRSChi.Recordset.RecordCount > 1
     btNextChien.Visible = btPevChien.Visible
        
    If concsixmois Then
        conditiondate = CStr(" and Date>" + datsql(DateAdd("m", -6, date)))
    Else
        conditiondate = ""
    End If
    
    datPrimaryRSConcours.RecordSource = "select Appreciation,Categorie,ChienID,Classement,Date,Discipline,Lieu,Points from Concours where (ChienID =" + idc + conditiondate + ") order by Date desc , Lieu, Discipline, Categorie"
    datPrimaryRSConcours.Refresh
        
      FrameConcours.Visible = True
      FrameConcoursF.Visible = True
      ' on laisse en place dans le cas de plusieurs saisies
                  'drpTypDisc.ListIndex = -1
                  'drpTypCat.ListIndex = -1
                  'drpTypApp.ListIndex = -1
                  'tbConcDate(0) = ""
                  'tbConcLieu(0) = ""
                  'tbConcPoints(0) = ""
                  'tbConcClas(0) = ""
      
      If datPrimaryRSConcours.Recordset.RecordCount >= maxConcours Then
          nbLignes = maxConcours
          VScrollConcours.Max = datPrimaryRSConcours.Recordset.RecordCount - maxConcours
          
          VScrollConcours.Enabled = True
          VScrollConcours = 0
      Else
          VScrollConcours.Enabled = False
          VScrollConcours.Max = 0
          nbLignes = datPrimaryRSConcours.Recordset.RecordCount
      End If

      'Affichage des enregistrement (cf. recordset_willchange)
      For i = 1 To maxConcours
          tbConcDate(i) = ""
          tbConcLieu(i) = ""
          tbConcdisc(i) = ""
          tbConcCat(i) = ""
          tbConcApp(i) = ""
          tbConcPoints(i) = ""
          tbConcClas(i) = ""
      Next
      Call refrechConcours(1)
       
End Sub
Sub refrechConcours(posDepart As Integer)

If (datPrimaryRSConcours.Recordset.RecordCount > 0) Then
        datPrimaryRSConcours.Recordset.MoveFirst
        For j = 2 To posDepart
            datPrimaryRSConcours.Recordset.MoveNext
        Next
         
        For j = 1 To nbLignes
            If concDate <> "" Then
                tbConcDate(j) = concDate
            Else
            
            End If
            If concLieu <> "" Then
                tbConcLieu(j) = concLieu
            Else
            
            End If
            If concDisc <> "" Then
                tbConcdisc(j) = labs.lblDisc(concDisc)
            Else
            
            End If
            If concDisc <> "" And concCat <> "" Then
                tbConcCat(j) = labs.lblCat(concDisc, concCat)
            Else
            
            End If
            If concApp <> "" Then
                tbConcApp(j) = labs.lblApp(concApp)
            Else
            
            End If
            tbConcPoints(j) = concPoints
            tbConcClas(j) = concClas
            
            datPrimaryRSConcours.Recordset.MoveNext
        Next
        End If
        tbConcDate(0).SetFocus
End Sub


Private Sub rechChien_Change()
    If rechChien <> "" Then
        If datPrimaryRSChi.Recordset.RecordCount < 1 Then
            rechChien = ""
            MsgBox "Pas de chien"
        Else
            If rech_nom(rechChien, nomChien(0), datPrimaryRSChi) Then
                btSelChien_Click
            End If
        End If
    End If
End Sub


Private Sub rechChien_GotFocus()

        FrameCHIENS.Visible = True
        imgBienvenu.Visible = False
        btAddChien.Visible = False
                totAdh.Visible = False
                lblTotAdh.Visible = False
                btNextAdh.Visible = False
                btPrevAdh.Visible = False
                btSelAdh.Visible = False
                
                
        FrameCoti.Visible = False
        FrameConcours.Visible = False
            FrameConcoursF.Visible = False
            btPevChien.Visible = False
            btNextChien.Visible = False
        FrameAdh.Visible = False
        btSupAdh.Visible = False
        btModAdh.Visible = False
        VScrollChiens.Visible = False
        

        'Récupération de tous les chiens
        datPrimaryRSChi.RecordSource = "select MaitreID,id,Nom,Race,Sexe,LOF,Tatouage,DateNaiss,no_CNEA,no_CUN from Chiens order by Nom"
        datPrimaryRSChi.Refresh
        totChiens = datPrimaryRSChi.Recordset.RecordCount

        If datPrimaryRSChi.Recordset.RecordCount > 0 Then
        
            btNextChienRech.Visible = True
            btPevChienRech.Visible = True
            btSelChien.Visible = True
            lblTotChien.Visible = True
            totChiens.Visible = True
            
            frameChien(0).Visible = True
            btSupChien(0).Visible = False
            modChien(0).Visible = False
            btAffConc(0).Visible = False
            For i = 1 To maxChienAff - 1
                frameChien(i).Visible = False
            Next
    
            
        Else
            'MsgBox "Pas de chien"
            
        End If

End Sub

Private Sub rechChien_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then btSelChien_Click
End Sub
Private Sub btSelChien_Click()
    idc = idChien(0)
    'Call selChien(idChien(0))
        btNextChienRech.Visible = False
        btPevChienRech.Visible = False
        btSelChien.Visible = False
        totChiens.Visible = False
        lblTotChien.Visible = False
        
    Call selAdh(Maitreid)
    Call selChienAdh(Maitreid)
    
    While idc <> idChien(0)
        datPrimaryRSChi.Recordset.MoveNext
    Wend
    Call selConcours(idChien(0))
        
End Sub
Private Sub sex_Click(index As Integer)
        If sex(index) = 1 Then
            labelSex(index) = "M" 'âle"
        ElseIf sex(index) = 0 Then
            labelSex(index) = "F" 'emelle"
        End If
End Sub


Private Sub tbConcClas_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If (KeyCode = KKenter) Then btAddConcours_Click
End Sub

Private Sub tbConcDate_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then tbConcLieu(0).SetFocus
End Sub

Private Sub tbConcDate_LostFocus(index As Integer)

    If (tbConcDate(index) <> "") Then
        If InStr(tbConcDate(index), ".") = 0 Then
             tbConcDate(index) = tbConcDate(index) & "." & Month(date)
         End If
        If IsDate(tbConcDate(index)) Then
             tbConcDate(index) = FormatDateTime(tbConcDate(index))
        End If
    End If
End Sub





Private Sub typLst_click()
    If typLst.ListIndex > 0 Then
        Dim d As New fListes
        d.Show
        d.typLst.ListIndex = typLst.ListIndex
    End If
End Sub


Private Sub tbConcLieu_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = KKenter Then drpTypDisc.SetFocus
End Sub

Private Sub tbConcPoints_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = KKenter Then tbConcClas(0).SetFocus
End Sub

Private Sub VScrollChiens_Change()
    pos = VScrollChiens
    
        ' on se garde le premier de coté
        datPrimaryRSChi.Recordset.MoveFirst
        datPrimaryRSChi.Recordset.Move (VScrollChiens)
        
        'on affiche les autres
        For i = 1 To maxChienAff - 1
            
            datPrimaryRSChi.Recordset.MoveNext
            idChien(i) = idChien(0)
            nomChien(i) = nomChien(0)
            tatou(i) = tatou(0)
            chienlof(i) = chienlof(0)
            dateNaiss(i) = dateNaiss(0)
            labelRace(i) = labelRace(0)
            labelSex(i) = labelSex(0)
            cun(i) = cun(0)
            cnea(i) = cnea(0)
        Next
        datPrimaryRSChi.Recordset.Move (-2)
            
End Sub


Private Sub VScrollConcours_Change()
    pos = VScrollConcours
    'Affichage des enregistrement (cf. recordset_willchange)
    Call refrechConcours(VScrollConcours + 1)
End Sub

Sub delConcours(noChi As Integer)
  On Error GoTo DeleteErr
    
    datPrimaryRSConcours.RecordSource = "select ChienID,Date,Lieu,Discipline,Categorie,Appreciation,Points,Classement from Concours where ChienID=" + CStr(noChi)
    datPrimaryRSConcours.Refresh
        
    With datPrimaryRSConcours.Recordset
            If .RecordCount > 0 Then
                .MoveFirst
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
            End If
    End With
    datPrimaryRSConcours.Refresh
    datPrimaryRSConcours.Refresh
    FrameConcours.Visible = False
        FrameConcoursF.Visible = False
        btPevChien.Visible = False
        btNextChien.Visible = False
  Exit Sub
DeleteErr:
  MsgBox err.Description
End Sub

