Attribute VB_Name = "Module1"
Option Explicit


Public fMainForm As fMain
Public fGbl As frmGlobale
Public fMajConc As siteConc
Public Const dateMin = "01.04.1977" ' date de création du club
Public Const dateMax = "31.12.2200"
Public Const KKenter = 13
Public Const KKtab = 9
Public Const KKechap = 8
Public Const KKup = 38
Public Const KKdown = 40
Public Const maxConcours = 12
Public Const maxChienAff = 3
Public Const destMAJ = "a:\maj.zip"
'Public Const destMAJ = "maj\maj.zip"
Public Const siteConcMaxLignes = 5

Public chemin As String
Public cheminSite As String
Public cheminListes As String
Public nomDB As String
'


Sub Main()
        On Error Resume Next
  
    
    chemin = CurDir$ + "\"
    nomDB = chemin + "canine.MDB"
    cheminSite = chemin + "Site\"
    cheminListes = chemin + "listes\"
    Set fMainForm = New fMain
    fMainForm.WindowState = vbMaximized
'    fMainForm.Show
    
End Sub


Function datsql(dtvb As String) As String
Dim dTmp As String
Dim mTmp As String
Dim yTmp As String
    dTmp = Left(dtvb, 2)
    mTmp = Mid(dtvb, 4, 2)
    yTmp = Mid(dtvb, 7, 4)
    datsql = "#" + mTmp + "/" + dTmp + "/" + yTmp + "#"
    
End Function

Function mois(no As Integer) As String
 Dim txt As String
 txt = ""
Select Case no
     Case 1: txt = "Janvier "
     Case 2: txt = "Février "
     Case 3: txt = "Mars "
     Case 4: txt = "Avril "
     Case 5: txt = "Mai "
     Case 6: txt = "Juin "
     Case 7: txt = "Juillet "
     Case 8: txt = "Aout "
     Case 9: txt = "Septembre "
     Case 10: txt = "Octobre "
     Case 11: txt = "Novembre "
     Case 12: txt = "Décembre "
 End Select
    mois = txt
End Function
Function nbWin() As Integer

    Dim frm As Form
    Dim is_child As Boolean
    Dim num_mdi As Integer
    
        For Each frm In Forms
            ' See if this form is an MDI child.
            On Error Resume Next
            is_child = frm.MDIChild
            If err.Number <> 0 Then is_child = False
            On Error GoTo 0
    
            If is_child Then num_mdi = num_mdi + 1
        Next frm
        nbWin = num_mdi
       
    
End Function
Function fermWin() As Integer

    Dim frm As Form
    Dim is_child As Boolean
    Dim num_mdi As Integer
    
        For Each frm In Forms
            ' See if this form is an MDI child.
            On Error Resume Next
            is_child = frm.MDIChild
            If err.Number <> 0 Then
                is_child = False
            Else
                If frm.hWnd <> fGbl.hWnd Then Unload frm
            End If
            On Error GoTo 0
    
            
        Next frm
   
       
    
End Function

