Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

'===============================================================================
' Corrige l'audit du 2026-07-22 :
'   Le classeur original contenait un "Workbook_Open" declare par erreur dans
'   le module du formulaire frmGestionDossier au lieu du module ThisWorkbook.
'   Excel ne declenche cet evenement que s'il est ICI. Resultat : le menu ne
'   s'affichait jamais automatiquement a l'ouverture du fichier.
'===============================================================================

Private Sub Workbook_Open()
    On Error Resume Next
    modMain.ShowMainMenu
    On Error GoTo 0
End Sub
