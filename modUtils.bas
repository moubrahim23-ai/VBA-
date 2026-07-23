Attribute VB_Name = "modUtils"
Option Explicit

'===============================================================================
' modUtils
' Fonctions utilitaires fichiers/dossiers + journalisation.
'
' Corrige l'audit du 2026-07-22 :
'   - CreateObject("Scripting.FileSystemObject") etait appele 13 fois dans le
'     projet. Une seule instance est desormais partagee (fonction GetFSO).
'   - LogAction avait une signature dont l'ordre des parametres ne
'     correspondait pas a l'ecriture interne dans les cellules ; ca ne
'     "marchait" que parce que les 2 sites d'appel du formulaire passaient
'     les arguments dans un ordre qui compensait exactement l'erreur. Toute
'     nouvelle utilisation avec le "bon" ordre aurait produit des donnees
'     fausses. La signature ci-dessous ecrit directement dans les colonnes
'     reelles de LOG_ACTIONS (Date, Action, Fichier, Domaine, Utilisateur,
'     Details, SousDomaine, Comment) sans intermediaire ambigu.
'   - LoAction (renommee LogActionNC) cherchait la feuille NC_LOG_ACTIONS
'     mais en creait une nommee LOG_ACTIONS si elle etait absente -> collision
'     de nom garantie avec la feuille deja creee par LogAction. Corrige.
'===============================================================================

Private mFSO As Object

' Instance FSO partagee - remplace les 13 CreateObject(...) individuels.
Public Function GetFSO() As Object
    If mFSO Is Nothing Then
        Set mFSO = CreateObject("Scripting.FileSystemObject")
    End If
    Set GetFSO = mFSO
End Function

' ------------------------------------------------------------------ Dossiers --

Public Function GetSubFolders(path As String) As Collection
    Dim col As New Collection
    Dim subFolder As Object

    If GetFSO().FolderExists(path) Then
        For Each subFolder In GetFSO().GetFolder(path).SubFolders
            col.Add subFolder.Name
        Next subFolder
    End If
    Set GetSubFolders = col
End Function

Public Function GetNonStatutFolders(path As String) As Collection
    Dim col As New Collection
    Dim subFolder As Object
    Dim statutFolders As Variant, i As Long, isStatut As Boolean

    statutFolders = modConfig.GetStatutFolders()

    If GetFSO().FolderExists(path) Then
        For Each subFolder In GetFSO().GetFolder(path).SubFolders
            isStatut = False
            For i = LBound(statutFolders) To UBound(statutFolders)
                If UCase$(subFolder.Name) = UCase$(statutFolders(i)) Then
                    isStatut = True
                    Exit For
                End If
            Next i
            If Not isStatut Then col.Add subFolder.Name
        Next subFolder
    End If
    Set GetNonStatutFolders = col
End Function

Public Sub LoadFilesToList(lst As MSForms.ListBox, folderPath As String)
    Dim file As Object
    lst.Clear
    If GetFSO().FolderExists(folderPath) Then
        For Each file In GetFSO().GetFolder(folderPath).Files
            lst.AddItem file.Name
        Next file
    End If
End Sub

Public Function FolderExists(path As String) As Boolean
    FolderExists = GetFSO().FolderExists(path)
End Function

Public Function FileExists(filePath As String) As Boolean
    FileExists = GetFSO().FileExists(filePath)
End Function

Public Sub CreateFolderIfNotExists(path As String)
    If Not GetFSO().FolderExists(path) Then GetFSO().CreateFolder path
End Sub

' -------------------------------------------------------------- Suivi (trace) --

' Renvoie la ligne existante pour fileName dans la feuille de suivi, ou la
' premiere ligne libre si le fichier n'y figure pas encore.
Public Function GetTraceRow(ws As Worksheet, fileName As String) As Long
    Dim lastRow As Long, i As Long
    With ws
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).row
        If lastRow = 1 And .Cells(1, 1).Value = "" Then
            GetTraceRow = 2
            Exit Function
        End If
        For i = 2 To lastRow
            If .Cells(i, 1).Value = fileName Then
                GetTraceRow = i
                Exit Function
            End If
        Next i
        GetTraceRow = lastRow + 1
    End With
End Function

' ---------------------------------------------------------------- Journalisation --

' Journalise un deplacement "normal" (workflow A TRAITER / EN COURS / TRAITES).
' Ecrit directement dans les colonnes reelles de LOG_ACTIONS :
'   1 Date | 2 Action | 3 Fichier | 4 Domaine | 5 Utilisateur | 6 Details | 7 SousDomaine | 8 Comment
Public Sub LogAction(action As String, fileName As String, domaine As String, _
                      sousDomaine As String, agent As String, _
                      statutSource As String, statutDestination As String, _
                      Optional comment As String = "")
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(modConfig.SHEET_LOG)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = modConfig.SHEET_LOG
        ws.Cells(1, 1).Value = "Date"
        ws.Cells(1, 2).Value = "Action"
        ws.Cells(1, 3).Value = "Fichier"
        ws.Cells(1, 4).Value = "Domaine"
        ws.Cells(1, 5).Value = "Utilisateur"
        ws.Cells(1, 6).Value = "Details"
        ws.Cells(1, 7).Value = "SousDomaine"
        ws.Cells(1, 8).Value = "Comment"
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(200, 220, 240)
    End If

    With ws
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).row + 1
        .Cells(lastRow, 1).Value = Now
        .Cells(lastRow, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Cells(lastRow, 2).Value = action
        .Cells(lastRow, 3).Value = fileName
        .Cells(lastRow, 4).Value = domaine
        .Cells(lastRow, 5).Value = agent
        .Cells(lastRow, 6).Value = statutSource & " " & modConfig.STATUS_SEP_NEW & " " & statutDestination
        .Cells(lastRow, 7).Value = sousDomaine
        .Cells(lastRow, 8).Value = comment
        .Columns.AutoFit
    End With
End Sub

' Journalise un deplacement NON CONFORME -> TRAITES.
' Cible TOUJOURS SHEET_LOG_NC, y compris a la creation (corrige le bug de
' l'ancienne LoAction qui creait "LOG_ACTIONS" au lieu de "NC_LOG_ACTIONS").
Public Sub LogActionNC(action As String, fileName As String, agent As String, details As String)
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(modConfig.SHEET_LOG_NC)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = modConfig.SHEET_LOG_NC
        ws.Cells(1, 1).Value = "Date"
        ws.Cells(1, 2).Value = "Action"
        ws.Cells(1, 3).Value = "Fichier"
        ws.Cells(1, 4).Value = "Utilisateur"
        ws.Cells(1, 5).Value = "Details"
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(200, 220, 240)
    End If

    With ws
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).row + 1
        .Cells(lastRow, 1).Value = Now
        .Cells(lastRow, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Cells(lastRow, 2).Value = action
        .Cells(lastRow, 3).Value = fileName
        .Cells(lastRow, 4).Value = agent
        .Cells(lastRow, 5).Value = details
        .Columns.AutoFit
    End With
End Sub
