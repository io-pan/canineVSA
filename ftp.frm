VERSION 5.00
Begin VB.Form fFtp 
   BackColor       =   &H00E2F5FF&
   Caption         =   "Publication de la page"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2F5FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2F5FF&
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
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00AAD69B&
         FillColor       =   &H00AAD69B&
         Height          =   735
         Left            =   0
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00D5F1FF&
      FillColor       =   &H00D5F1FF&
      FillStyle       =   0  'Solid
      Height          =   1035
      Left            =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "fFtp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActiveSession As Boolean
Dim hOpen As Long, hConnection As Long
Dim dwType As Long

    

Sub publier(page As String)

    Dim bRet As Boolean
    Dim szFileRemote As String, szDirRemote As String, szFileLocal As String
    Dim szTempString As String
    Dim txtServer As String
  
    lbl(0) = "Connection à internet"
    txtServer = "perso-ftp.wanadoo.fr"
    txtUser = "educationcanine.ville-sous-anjou.38"
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If hOpen = 0 Then ErrorOut err.LastDllError, "InternetOpen"
    
    lbl(0) = "Connection à Wanadoo"
    If Not bActiveSession And hOpen <> 0 Then
        hConnection = InternetConnect(hOpen, txtServer, INTERNET_INVALID_PORT_NUMBER, _
        txtUser, "k4h9be2", INTERNET_SERVICE_FTP, 0, 0)
        If hConnection = 0 Then
            bActiveSession = False
            ErrorOut err.LastDllError, "InternetConnect"
        Else
            bActiveSession = True
       End If
    End If
    
    'publish
    lbl(0) = "Publication de la page"
    If bActiveSession Then
        szDirRemote = szDirRemote + "/pub"
        szFileRemote = "concours_entete.txt"
        szFileLocal = "d:\dev\site\concours_entete.txt"
        

       Call rcd(szDirRemote, txtServer)
        
        bRet = FtpPutFile(hConnection, szFileLocal, szFileRemote, dwType, 0)
        If bRet = False Then
            ErrorOut err.LastDllError, "FtpPutFile"
            Exit Sub
        End If
        

   End If
   
   lbl(0) = "Déconnection à Wanadoo"
'disconect
    bDirEmpty = True
    If hConnection <> 0 Then InternetCloseHandle hConnection
    hConnection = 0
    

    bActiveSession = False

'close seeeion
    If hConnection <> 0 Then InternetCloseHandle (hConnection)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    hConnection = 0
    hOpen = 0
    bActiveSession = False

End Sub


Private Sub Form_Load()
    bActiveSession = False
    hOpen = 0
    hConnection = 0
    dwType = FTP_TRANSFER_TYPE_BINARY
    lbl(0) = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If hConnection <> 0 Then InternetCloseHandle (hConnection)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    hConnection = 0
    hOpen = 0
    If bActiveSession Then TreeView1.Nodes.Remove txtServer.Text
    bActiveSession = False
End Sub

Private Sub rcd(pszDir As String, txtServer As String)
    If pszDir = "" Then
        MsgBox "Please enter the directory to CD"
        Exit Sub
    Else
        Dim sPathFromRoot As String
        Dim bRet As Boolean
        If InStr(1, pszDir, txtServer) Then
        sPathFromRoot = Mid(pszDir, Len(txtServer) + 1, Len(pszDir) - Len(txtServer))
        Else
        sPathFromRoot = pszDir
        End If
        If sPathFromRoot = "" Then sPathFromRoot = "/"
        bRet = FtpSetCurrentDirectory(hConnection, sPathFromRoot)
        If bRet = False Then ErrorOut err.LastDllError, "rcd"
    End If
End Sub

Function ErrorOut(dError As Long, szCallFunction As String)
    Dim dwIntError As Long, dwLength As Long
    Dim strBuffer As String
    If dError = ERROR_INTERNET_EXTENDED_ERROR Then
        InternetGetLastResponseInfo dwIntError, vbNullString, dwLength
        strBuffer = String(dwLength + 1, 0)
        InternetGetLastResponseInfo dwIntError, strBuffer, dwLength
        
        MsgBox szCallFunction & " Extd Err: " & dwIntError & " " & strBuffer
       
        
    End If
    'close
        If hConnection Then InternetCloseHandle hConnection
        If hOpen Then InternetCloseHandle hOpen
        hConnection = 0
        hOpen = 0
        If bActiveSession Then TreeView1.Nodes.Remove txtServer.Text
        bActiveSession = False

End Function

