VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLabels 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6375
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   6210
   Icon            =   "Labels.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6210
   Begin VB.TextBox labelApp 
      DataField       =   "Appreciation"
      DataSource      =   "datlstApp"
      Height          =   285
      Left            =   1200
      TabIndex        =   20
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox posApp 
      DataField       =   "Position"
      DataSource      =   "datlstApp"
      Height          =   285
      Left            =   1200
      TabIndex        =   19
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Label"
      DataSource      =   "datlstRaces"
      Height          =   285
      Index           =   9
      Left            =   1920
      TabIndex        =   15
      Top             =   5355
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ID"
      DataSource      =   "datlstRaces"
      Height          =   285
      Index           =   8
      Left            =   1920
      TabIndex        =   13
      Top             =   5040
      Width           =   3375
   End
   Begin VB.TextBox labelCat 
      DataField       =   "labelCat"
      DataSource      =   "datlstCat"
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   2085
      Width           =   3375
   End
   Begin VB.TextBox idDiscCat 
      DataField       =   "idTypeConcours"
      DataSource      =   "datlstCat"
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   1755
      Width           =   3375
   End
   Begin VB.TextBox idCat 
      DataField       =   "idCat"
      DataSource      =   "datlstCat"
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox labelDisc 
      DataField       =   "label"
      DataSource      =   "datlstDisc"
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   435
      Width           =   3375
   End
   Begin VB.TextBox idDisc 
      DataField       =   "id"
      DataSource      =   "datlstDisc"
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc datlstDisc 
      Height          =   330
      Left            =   0
      Top             =   780
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   582
      ConnectMode     =   1
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
      Top             =   2400
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   582
      ConnectMode     =   1
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
   Begin MSAdodcLib.Adodc datlstRaces 
      Height          =   330
      Left            =   240
      Top             =   5700
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   582
      ConnectMode     =   1
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
      RecordSource    =   "select ID,Label from Races"
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
   Begin MSAdodcLib.Adodc datlstApp 
      Height          =   330
      Left            =   120
      Top             =   3705
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   582
      ConnectMode     =   1
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
      RecordSource    =   "select Appreciation,Position from TypAppreciation"
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
      Caption         =   "Position:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Appreciation:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   3525
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Races"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   "Label:"
      Height          =   255
      Index           =   9
      Left            =   0
      TabIndex        =   14
      Top             =   5355
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ID:"
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   12
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Catégories"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "labelCat:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   2445
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "idTypeConcours:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   2115
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "idCat:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Disciplines"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lblLabels 
      Caption         =   "label:"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   2
      Top             =   795
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "id:"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Initialize()
ChDir (chemin)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datlstTypVers_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datlstTypVers_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  datlstTypVers.Caption = "Record: " & CStr(datlstTypVers.Recordset.AbsolutePosition)
End Sub

Private Sub datlstTypVers_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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

Function lblCat(idDisc As Integer, id As Integer) As String
    lblCat = ""
    datlstCat.Recordset.MoveFirst
    While idDisc <> idDiscCat And datlstCat.Recordset.AbsolutePosition > 0 And datlstCat.Recordset.AbsolutePosition < datlstCat.Recordset.RecordCount
        datlstCat.Recordset.MoveNext
    Wend
    
    If (idDisc = idDiscCat) Then
        While id <> idCat And datlstCat.Recordset.AbsolutePosition > 0 And datlstCat.Recordset.AbsolutePosition < datlstCat.Recordset.RecordCount
            datlstCat.Recordset.MoveNext
        Wend
        If (id = idCat) Then
            lblCat = labelCat
        End If
    End If
End Function

Function lblDisc(id As Integer) As String
    lblDisc = ""
    datlstDisc.Recordset.MoveFirst
    While id <> idDisc And datlstDisc.Recordset.AbsolutePosition > 0 And datlstDisc.Recordset.AbsolutePosition < datlstDisc.Recordset.RecordCount
        datlstDisc.Recordset.MoveNext
    Wend

    If (id = idDisc) Then
        lblDisc = labelDisc
    End If
End Function
Function lblApp(id As Integer) As String
    lblApp = ""
    datlstApp.Recordset.MoveFirst
    While id <> posApp And datlstApp.Recordset.AbsolutePosition > 0 And datlstApp.Recordset.AbsolutePosition < datlstApp.Recordset.RecordCount
        datlstApp.Recordset.MoveNext
    Wend

    If (id = posApp) Then
        lblApp = labelApp
    End If
End Function

