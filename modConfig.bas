Attribute VB_Name = "modConfig"
Option Explicit

'===============================================================================
' modConfig
' Source UNIQUE de configuration : chemins, noms de feuilles, statuts,
' domaines et sous-domaines.
'
' Corrige l'audit du 2026-07-22 :
'   - la taxonomie des domaines/sous-domaines etait dupliquee a la main dans
'     au moins 5 endroits differents (CGO_Click, cboRacine_Change,
'     CreateACMFolders/CreateContrTheque, GetDomainStructure x2), avec des
'     divergences d'orthographe qui cassaient la reconnaissance de
'     CONTRATHEQUE ("CONTRTHEQUE" sans le A dans l'ancien modInitialization).
'   - Tout le reste du projet doit lire ses domaines/chemins ICI et nulle
'     part ailleurs.
'===============================================================================

' ---- Chemins reseau ----------------------------------------------------------
Public Const ROOT_PATH As String = "\\UF11-a03\eids-eaa-pdf-statiques\ACCESSIBILITE"

' NB audit : le dossier historique pointe vers un AUTRE serveur que ROOT_PATH
' dans le fichier original (\\uf11pu01 au lieu de \\UF11-a03). Je n'ai pas pu
' verifier lequel des deux est correct : verifiez ce chemin avant mise en
' production, puis ne le modifiez plus qu'ICI.
Public Const HISTO_PATH As String = "\\uf11pu01\ACCESSIBILITE_CONSEILLERS\ACCESSIBILITE\Historique"

' ---- Noms de feuilles ---------------------------------------------------------
' Le nom REEL de la feuille Excel n'est pas renomme (des TCD existants sur la
' feuille BUREAU s'appuient dessus) : seule cette constante VBA change de nom
' pour ne plus etre identique a sa propre valeur comme dans l'original
' (Public Const SHEET_TRACE_NAME As String = "SHEET_TRACE_NAME").
Public Const SHEET_TRACE As String = "SHEET_TRACE_NAME"
Public Const SHEET_LOG As String = "LOG_ACTIONS"
Public Const SHEET_LOG_NC As String = "NC_LOG_ACTIONS"
Public Const SHEET_SAISIE As String = "SAISIE_SARA"

' ---- Delimiteur pour stocker "statut source -> statut destination" -----------
' L'ancien code utilisait "?" (un caractere Unicode - probablement une fleche -
' corrompu en "?" litteral par l'editeur VBA, qui ne supporte que la page de
' code ANSI du poste). On le voit encore dans les 197 lignes historiques du
' journal ("A TRAITER ? EN COURS"). On garde la lecture compatible avec "?" ET
' "|", mais on n'ECRIT plus que "|" desormais.
Public Const STATUS_SEP_NEW As String = "|"
Public Const STATUS_SEP_LEGACY As String = "?"

' ---- Valeurs de statut ---------------------------------------------------------
Public Const STATUT_A_TRAITER As String = "A TRAITER"
Public Const STATUT_EN_COURS As String = "EN COURS"
Public Const STATUT_TRAITES As String = "TRAITES"
Public Const STATUT_NON_CONFORME As String = "NON CONFORME"

' ---- Structure des domaines -----------------------------------------------------
Public Enum StructureType
    SIMPLE = 1          ' 1 niveau : Domaine > Statut
    Complex = 2         ' 2 niveaux : Domaine > SousDomaine > Statut
    THREE_LEVEL = 3     ' 3 niveaux : Domaine > Statut > SousDomaine (ACM, DIRCO, CMO)
End Enum

' Retourne le type de structure d'un domaine. Domaine absent de la liste =>
' SIMPLE par defaut (comportement identique a l'original).
Public Function GetDomainStructure(domainName As String) As StructureType
    Select Case UCase$(Trim$(domainName))
        Case "ACM", "DIRCO", "CMO"
            GetDomainStructure = THREE_LEVEL
        Case "CONTRATHEQUE"
            GetDomainStructure = Complex
        Case Else
            GetDomainStructure = SIMPLE
    End Select
End Function

' Liste des domaines racine, dans l'ordre d'affichage voulu dans les combos.
Public Function GetDomaines() As Variant
    GetDomaines = Array("ACM", "Bouygues telecom", "CIC CIB ASSET SERVICING", _
        "CONTRATHEQUE", "DIRCO", "CMO", "CMMABN", "DIRCOM", "MONABANQUE")
End Function

' Sous-domaines connus pour chaque domaine. Remplace les listes dupliquees et
' divergentes de CGO_Click / cboRacine_Change / CreateACMFolders /
' CreateContrTheque. "GARANTIES" (pas "Granties"), "CONTRATHEQUE" (pas
' "CONTRTHEQUE") : orthographe corrigee et desormais unique.
Public Function GetSousDomaines(domainName As String) As Variant
    Select Case UCase$(Trim$(domainName))
        Case "ACM"
            GetSousDomaines = Array("ASS-BIENS", "ASS-COMPTA-CONTENTIEUX", "ASS-EMPRUNTEURS", _
                "ASS-INDEMN-BIENS-PERS", "ASS-PERS-COLL", "ASS-PERS-IND", _
                "ASS-PROTECTION-JUR", "ASS-VIE-EPARGNE-RET")
        Case "CMO"
            GetSousDomaines = Array("BAD", "Cartes bancaires", "Compte courant", _
                "Epargne bancaire", "Epargne financiere", "Packages", "Tarif")
        Case "CONTRATHEQUE"
            ' NB audit : l'original avait 3 listes divergentes pour ce domaine
            ' (CGO_Click, cboRacine_Change, CreateContrTheque), qui ne
            ' s'accordaient ni sur l'orthographe ("Granties" vs "GARANTIES")
            ' ni sur le contenu ("Conventions" absent de 2 des 3 listes).
            ' Ceci est l'union des trois, pour ne rien perdre. A valider avec
            ' vous : si "Conventions" ne doit pas exister comme sous-dossier
            ' CONTRATHEQUE, retirez-le simplement de cette liste.
            GetSousDomaines = Array("CONDITIONS GENERALES", "CONVENTIONS", _
                "DOCUMENT D'INFORMATION TARIFAIRE", "GARANTIES", "POLITIQUE DE PROTECTION")
        Case "DIRCO"
            GetSousDomaines = Array("CIC", "CM")
        Case Else
            GetSousDomaines = Array("[PAS DE SOUS-DOMAINE]")
    End Select
End Function

' Statuts "source" proposes lors d'un deplacement (ecran principal).
Public Function GetStatutsSource() As Variant
    GetStatutsSource = Array(STATUT_A_TRAITER, STATUT_EN_COURS)
End Function

' Statuts "destination" proposes lors d'un deplacement.
Public Function GetStatutsDestination() As Variant
    GetStatutsDestination = Array(STATUT_EN_COURS, STATUT_TRAITES, STATUT_NON_CONFORME)
End Function

' Sous-dossiers "statut" standard crees a la racine de chaque domaine simple/complexe.
Public Function GetStatutFolders() As Variant
    GetStatutFolders = Array(STATUT_A_TRAITER, STATUT_EN_COURS, STATUT_TRAITES, STATUT_NON_CONFORME)
End Function
