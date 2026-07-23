Attribute VB_Name = "modMain"
Option Explicit

'===============================================================================
' modMain - Point d'entree
'
' Corrige l'audit du 2026-07-22 :
'   Les references a la feuille de suivi utilisaient le texte litteral
'   "SHEET_TRACE_NAME" a plusieurs endroits du projet au lieu de la constante
'   du meme nom (qui valait... "SHEET_TRACE_NAME", un exemple typique de
'   constante jamais vraiment renseignee). Desormais tout passe par
'   modConfig.SHEET_TRACE.
'===============================================================================

' Lanceur principal
Public Sub Main()
    On Error GoTo ErrorHandler

    If Not modInitialization.CheckFolderStructure() Then
        If MsgBox("Structure incomplete. Creer les dossiers ?", vbYesNo + vbExclamation) = vbYes Then
            modInitialization.CreateMissingFolders
        Else
            Exit Sub
        End If
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(modConfig.SHEET_TRACE)
    On Error GoTo 0

    If ws Is Nothing Then
        If MsgBox("Initialiser la feuille de tracabilite ?", vbYesNo + vbQuestion) = vbYes Then
            modInitialization.InitializeTraceSheet
        End If
    End If

    frmGestionDossier.Show vbModeless

    Exit Sub

ErrorHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

' Raccourci clavier
Public Sub LanceurRapide()
    Main
End Sub

' Menu principal (appele desormais depuis le vrai Workbook_Open, cf. ThisWorkbook.bas)
Public Sub ShowMainMenu()
    Dim choix As VbMsgBoxResult

    choix = MsgBox("OUTIL DE GESTION D'ACCESSIBILITE" & vbCrLf & vbCrLf & _
                  "Que voulez-vous faire ?" & vbCrLf & _
                  " OUI : Gerer les dossiers" & vbCrLf & _
                  " NON : Voir la tracabilite" & vbCrLf & _
                  " ANNULER : Quitter", _
                  vbYesNoCancel + vbInformation, "Menu")

    Select Case choix
        Case vbYes
            Main
        Case vbNo
            On Error Resume Next
            ThisWorkbook.Sheets(modConfig.SHEET_TRACE).Activate
            On Error GoTo 0
        Case vbCancel
            ' Rien
    End Select
End Sub
