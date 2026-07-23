Attribute VB_Name = "modInitialization"
Option Explicit

'===============================================================================
' modInitialization
'
' Corrige l'audit du 2026-07-22 :
'   - InitializeTraceSheet supprimait purement et simplement la feuille
'     SHEET_TRACE_NAME existante avant de la recreer avec un en-tete de 10
'     colonnes. Or la feuille reelle en a aujourd'hui 13 (Status_2,
'     Date_de_Retraitement, DMt_2, i ont ete ajoutees a la main). Si cette
'     procedure avait ete rappelee (elle l'est des que Main() ne trouve pas
'     "explicitement" la feuille), les 4 colonnes ajoutees auraient ete
'     perdues. Elle est desormais IDEMPOTENTE : si la feuille existe deja,
'     elle n'est pas touchee, sauf demande explicite (forceRecreate:=True).
'   - CheckFolderStructure / CreateMissingFolders utilisaient "CONTRTHEQUE"
'     (sans le A) alors que GetDomainStructure et l'interface utilisent
'     "CONTRATHEQUE" : ce domaine n'etait donc jamais reconnu comme
'     "Complexe" par la verification de structure. Corrige, et les listes de
'     sous-dossiers viennent maintenant de modConfig (source unique).
'===============================================================================

Public Sub InitializeTraceSheet(Optional forceRecreate As Boolean = False)
    Dim ws As Worksheet
    Dim headers As Variant
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(modConfig.SHEET_TRACE)
    On Error GoTo 0

    If Not ws Is Nothing And Not forceRecreate Then
        ' La feuille existe deja : on ne la recree pas (evite d'ecraser les
        ' colonnes ajoutees manuellement). On verifie juste que l'en-tete de
        ' base est present, sans rien supprimer.
        If Trim$(ws.Cells(1, 1).Value) = "" Then
            ws.Cells(1, 1).Value = "NomFichier"
        End If
        Exit Sub
    End If

    If Not ws Is Nothing And forceRecreate Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = modConfig.SHEET_TRACE

    headers = Array("NomFichier", "Domaine", "SousDomaine", "Statut", "Status_2", _
                    "Date_A_TRAITER", "Date_FIN", "Date_de_Retraitement", _
                    "Agent", "DMT", "DMt_2", "Commentaire", "i")

    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(198, 224, 180)
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns("A:A").ColumnWidth = 30
    ws.Columns("B:C").ColumnWidth = 25
    ws.Columns("D:E").ColumnWidth = 15
    ws.Columns("F:H").ColumnWidth = 20
    ws.Columns("I:I").ColumnWidth = 15
    ws.Columns("J:K").ColumnWidth = 12
    ws.Columns("L:L").ColumnWidth = 40

    ws.Columns("F:H").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Columns("J:K").NumberFormat = "[hh]:mm"

    ws.Range("A1:M1").AutoFilter

    MsgBox "Feuille de suivi initialisee avec succes !", vbInformation
End Sub

Public Function CheckFolderStructure() As Boolean
    Dim mainDomains As Variant, domain As Variant, domainPath As String

    mainDomains = Array("ACM", "CIC CIB ASSET SERVICING", "DIRCO", "CONTRATHEQUE")

    If Not modUtils.FolderExists(modConfig.ROOT_PATH) Then
        MsgBox "Dossier racine introuvable : " & modConfig.ROOT_PATH, vbCritical
        CheckFolderStructure = False
        Exit Function
    End If

    For Each domain In mainDomains
        domainPath = modConfig.ROOT_PATH & "\" & domain
        If Not modUtils.FolderExists(domainPath) Then
            MsgBox "Domaine manquant : " & domain, vbExclamation
            CheckFolderStructure = False
            Exit Function
        End If
    Next domain

    CheckFolderStructure = True
End Function

Public Sub CreateMissingFolders()
    CreateACMFolders
    CreateSimpleDomain "CIC CIB ASSET SERVICING"
    CreateComplexDomain "DIRCO", Array("FICHES CLIENTS")
    CreateContrTheque
    MsgBox "Structure des dossiers creee avec succes !", vbInformation
End Sub

Private Sub CreateACMFolders()
    Dim domainPath As String, aTraiterPath As String
    Dim subFolder As Variant, statut As Variant

    domainPath = modConfig.ROOT_PATH & "\ACM"
    modUtils.CreateFolderIfNotExists domainPath

    aTraiterPath = domainPath & "\" & modConfig.STATUT_A_TRAITER
    modUtils.CreateFolderIfNotExists aTraiterPath

    For Each subFolder In modConfig.GetSousDomaines("ACM")
        modUtils.CreateFolderIfNotExists aTraiterPath & "\" & subFolder
    Next subFolder

    For Each statut In modConfig.GetStatutFolders()
        If statut <> modConfig.STATUT_A_TRAITER Then
            modUtils.CreateFolderIfNotExists domainPath & "\" & statut
        End If
    Next statut
End Sub

Private Sub CreateSimpleDomain(domainName As String)
    Dim domainPath As String, statut As Variant
    domainPath = modConfig.ROOT_PATH & "\" & domainName
    modUtils.CreateFolderIfNotExists domainPath
    For Each statut In modConfig.GetStatutFolders()
        modUtils.CreateFolderIfNotExists domainPath & "\" & statut
    Next statut
End Sub

Private Sub CreateComplexDomain(domainName As String, categories As Variant)
    Dim domainPath As String, category As Variant, statut As Variant
    domainPath = modConfig.ROOT_PATH & "\" & domainName
    modUtils.CreateFolderIfNotExists domainPath
    For Each category In categories
        modUtils.CreateFolderIfNotExists domainPath & "\" & category
        For Each statut In modConfig.GetStatutFolders()
            modUtils.CreateFolderIfNotExists domainPath & "\" & category & "\" & statut
        Next statut
    Next category
End Sub

Private Sub CreateContrTheque()
    CreateComplexDomain "CONTRATHEQUE", modConfig.GetSousDomaines("CONTRATHEQUE")
End Sub
