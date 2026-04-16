#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <File.au3>
#include <EditConstants.au3>
#include <GuiEdit.au3>
#include <Date.au3>
#include <ProgressConstants.au3>

Opt("WinTitleMatchMode", 2)

; =================================================================================
; 0. CONFIG
; =================================================================================
Global $g_sConfigFile = @ScriptDir & "\robot_v26_4.ini"
Global $g_sTempFolder = @TempDir & "\Temp_EDOC_Robot"
Global $g_sSep = Chr(30)

If Not FileExists($g_sTempFolder) Then DirCreate($g_sTempFolder)

; Timings
Global $g_iPollMs = 250
Global $g_iEdocPostEnterMs = 2200
Global $g_iUploadPostReadyMs = 1200
Global $g_iWordOpenMs = 1800
Global $g_iEtmsAfterPageUpMs = 400
Global $g_iEtmsAfterF2Ms = 1200

; ETMS constants
Global $g_sEtms16 = "N"
Global $g_sEtms22 = "G0731153"
Global $g_sEtms23 = "Y"
Global $g_sEtms24 = "Y"
Global $g_sEtms25 = "CDGH"
Global $g_sEtms26 = "CDGH"
Global $g_sEtms27 = "17:50"

; Colors
Global Const $CLR_BG = 0xF5F7FB
Global Const $CLR_ACCENT = 0x2563EB
Global Const $CLR_SUBTEXT = 0x64748B
Global Const $CLR_OK = 0x059669
Global Const $CLR_WARN = 0xD97706
Global Const $CLR_ERR = 0xDC2626

; COM error handler
Global $g_oComErr = ObjEvent("AutoIt.Error", "_ComErrHandler")
Global $g_bComErr = False

; Outlook
Global $g_oOutlook = ObjGet("", "Outlook.Application")
If @error Or Not IsObj($g_oOutlook) Then
    MsgBox(16, "Erreur", "Outlook doit être ouvert avant de lancer le robot.")
    Exit
EndIf
Global $g_oNamespace = $g_oOutlook.GetNamespace("MAPI")
If Not IsObj($g_oNamespace) Then
    MsgBox(16, "Erreur", "Impossible d'initialiser Outlook MAPI.")
    Exit
EndIf

; Saved settings
Global $g_sSavedEmail = IniRead($g_sConfigFile, "Settings", "TargetEmail", "service.export@cdghandling.fr")
Global $g_sSavedAccount = IniRead($g_sConfigFile, "Settings", "LastAccount", "")
Global $g_sSavedExcel = IniRead($g_sConfigFile, "Settings", "ExcelPath", "")
Global $g_sSavedPeriod = IniRead($g_sConfigFile, "Settings", "Period", "Aujourd'hui")
Global $g_sSavedSession = IniRead($g_sConfigFile, "Settings", "EtmsSession", "JASONB")
Global $g_iSavedAnticipe = Number(IniRead($g_sConfigFile, "Settings", "Anticipe", "0"))
Global $g_sSavedFolderPath = IniRead($g_sConfigFile, "Settings", "OutlookFolderPath", "")
Global $g_iSavedIncludeSub = Number(IniRead($g_sConfigFile, "Settings", "IncludeSubFolders", "1"))

; Stores list
Global $g_sStoresList = ""
For $oStore In $g_oNamespace.Stores
    $g_sStoresList &= $oStore.DisplayName & "|"
Next
If $g_sStoresList <> "" Then $g_sStoresList = StringTrimRight($g_sStoresList, 1)

; =================================================================================
; 1. GUI CONTROLS
; =================================================================================
Global $hMainGUI
Global $idComboMail, $idInpTargetEmail, $idInpExcel, $idBtnExcel, $idComboPeriod
Global $idInpEtmsSession, $idChkAnticipe
Global $idInpOutlookFolder, $idChkIncludeSubfolders
Global $idStatusLabel, $idConsoleLog, $idBtnStart, $idBtnTest, $idScanHistory
Global $idProgressBar, $idProgressText

; =================================================================================
; 2. GUI BUILD
; =================================================================================
$hMainGUI = GUICreate("ROBOT EDOC + ETMS V26.4", 1020, 900)
GUISetBkColor($CLR_BG)

Local $idHeader = GUICtrlCreateLabel("", 0, 0, 1020, 78)
GUICtrlSetBkColor($idHeader, $CLR_ACCENT)

Local $idTitle = GUICtrlCreateLabel("ROBOT EDOC + ETMS", 24, 16, 500, 24)
GUICtrlSetFont($idTitle, 14, 800, 0, "Segoe UI")
GUICtrlSetColor($idTitle, 0xFFFFFF)
GUICtrlSetBkColor($idTitle, $CLR_ACCENT)

Local $idSubtitle = GUICtrlCreateLabel("Analyse globale des PDF dans un dossier Outlook + rapport Excel", 24, 44, 900, 18)
GUICtrlSetFont($idSubtitle, 9, 400, 0, "Segoe UI")
GUICtrlSetColor($idSubtitle, 0xE0E7FF)
GUICtrlSetBkColor($idSubtitle, $CLR_ACCENT)

GUICtrlCreateGroup(" CONFIGURATION ", 18, 92, 984, 245)
GUICtrlSetColor(-1, $CLR_ACCENT)
GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")

GUICtrlCreateLabel("Compte Outlook", 32, 118, 180, 18)
GUICtrlSetColor(-1, $CLR_SUBTEXT)
$idComboMail = GUICtrlCreateCombo("", 32, 138, 950, 26)
GUICtrlSetData($idComboMail, $g_sStoresList, $g_sSavedAccount)

GUICtrlCreateLabel("Adresse e-mail cible", 32, 174, 180, 18)
GUICtrlSetColor(-1, $CLR_SUBTEXT)
$idInpTargetEmail = GUICtrlCreateInput($g_sSavedEmail, 32, 194, 950, 24)

GUICtrlCreateLabel("Fichier Excel", 32, 228, 180, 18)
GUICtrlSetColor(-1, $CLR_SUBTEXT)
$idInpExcel = GUICtrlCreateInput($g_sSavedExcel, 32, 248, 850, 24)
$idBtnExcel = GUICtrlCreateButton("Parcourir", 896, 246, 86, 28)
GUICtrlSetFont($idBtnExcel, 9, 700, 0, "Segoe UI")

GUICtrlCreateLabel("Dossier Outlook à scanner (ex: Boîte de réception\\EXPORT)", 32, 284, 460, 18)
GUICtrlSetColor(-1, $CLR_SUBTEXT)
$idInpOutlookFolder = GUICtrlCreateInput($g_sSavedFolderPath, 32, 304, 650, 24)

$idChkIncludeSubfolders = GUICtrlCreateCheckbox("Inclure les sous-dossiers", 700, 306, 240, 20)
If $g_iSavedIncludeSub = 1 Then GUICtrlSetState($idChkIncludeSubfolders, $GUI_CHECKED)

GUICtrlCreateGroup(" OPTIONS ", 18, 348, 984, 96)
GUICtrlSetColor(-1, $CLR_ACCENT)
GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")

GUICtrlCreateLabel("Période", 32, 372, 180, 18)
GUICtrlSetColor(-1, $CLR_SUBTEXT)
$idComboPeriod = GUICtrlCreateCombo("", 32, 392, 300, 25)
GUICtrlSetData($idComboPeriod, "Aujourd'hui|1 heure|2 heures|4 heures|8 heures|24 heures|7 jours|Depuis dernier scan", $g_sSavedPeriod)

GUICtrlCreateLabel("Session ETMS", 380, 372, 180, 18)
GUICtrlSetColor(-1, $CLR_SUBTEXT)
$idInpEtmsSession = GUICtrlCreateInput($g_sSavedSession, 380, 392, 180, 24)

$idChkAnticipe = GUICtrlCreateCheckbox("Mode anticipé (date ETMS = jour même)", 600, 394, 330, 20)
If $g_iSavedAnticipe = 1 Then GUICtrlSetState($idChkAnticipe, $GUI_CHECKED)

GUICtrlCreateGroup(" STATUT ", 18, 456, 984, 104)
GUICtrlSetColor(-1, $CLR_ACCENT)
GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")

$idStatusLabel = GUICtrlCreateLabel("En attente...", 34, 482, 950, 20)
GUICtrlSetFont($idStatusLabel, 11, 800, 0, "Segoe UI")
GUICtrlSetColor($idStatusLabel, $CLR_SUBTEXT)

$idProgressBar = GUICtrlCreateProgress(34, 512, 890, 18, $PBS_SMOOTH)
GUICtrlSetData($idProgressBar, 0)

$idProgressText = GUICtrlCreateLabel("0 %", 932, 510, 60, 18)
GUICtrlSetFont($idProgressText, 9, 700, 0, "Segoe UI")
GUICtrlSetColor($idProgressText, $CLR_SUBTEXT)

GUICtrlCreateGroup(" JOURNAL ", 18, 572, 984, 250)
GUICtrlSetColor(-1, $CLR_ACCENT)
GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")

$idConsoleLog = GUICtrlCreateEdit("", 32, 598, 956, 210, BitOR($ES_READONLY, $WS_VSCROLL, $ES_AUTOVSCROLL, $ES_MULTILINE))
GUICtrlSetFont($idConsoleLog, 9, 400, 0, "Segoe UI")
GUICtrlSetBkColor($idConsoleLog, 0xFFFFFF)

GUICtrlCreateGroup(" DERNIERS TRAITEMENTS ", 18, 832, 984, 52)
GUICtrlSetColor(-1, $CLR_ACCENT)
GUICtrlSetFont(-1, 9, 700, 0, "Segoe UI")

$idScanHistory = GUICtrlCreateEdit("", 32, 854, 956, 18, BitOR($ES_READONLY, $WS_VSCROLL, $ES_AUTOVSCROLL, $ES_MULTILINE))
GUICtrlSetFont($idScanHistory, 8, 400, 0, "Segoe UI")
GUICtrlSetBkColor($idScanHistory, 0xFFFFFF)

$idBtnTest = GUICtrlCreateButton("ANALYSER + TEST (sans F1)", 240, 888, 245, 30)
GUICtrlSetFont($idBtnTest, 10, 800, 0, "Segoe UI")

$idBtnStart = GUICtrlCreateButton("ANALYSER + TRAITER (avec F1)", 540, 888, 260, 30)
GUICtrlSetFont($idBtnStart, 10, 800, 0, "Segoe UI")

_RefreshScanHistoryControl()
GUISetState(@SW_SHOW)

; =================================================================================
; 3. MAIN LOOP
; =================================================================================
While 1
    Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE
            Exit

        Case $idBtnExcel
            Local $sFile = FileOpenDialog("Ouvrir Excel", @DesktopDir, "Excel (*.xlsx;*.xls)")
            If $sFile <> "" Then GUICtrlSetData($idInpExcel, $sFile)

        Case $idBtnTest
            _RunAnalyzeThenProcess(False)

        Case $idBtnStart
            _RunAnalyzeThenProcess(True)
    EndSwitch
WEnd

; =================================================================================
; 4. PIPELINE
; =================================================================================
Func _RunAnalyzeThenProcess($bValidateWithF1)
    Local $sExcel = StringStripWS(GUICtrlRead($idInpExcel), 3)
    Local $sTargetEmail = StringStripWS(GUICtrlRead($idInpTargetEmail), 3)
    Local $sStoreName = StringStripWS(GUICtrlRead($idComboMail), 3)
    Local $sPeriod = StringStripWS(GUICtrlRead($idComboPeriod), 3)
    Local $sEtmsSession = StringUpper(StringStripWS(GUICtrlRead($idInpEtmsSession), 3))
    Local $bAnticipe = (BitAND(GUICtrlRead($idChkAnticipe), $GUI_CHECKED) = $GUI_CHECKED)

    Local $sOutlookFolderPath = StringStripWS(GUICtrlRead($idInpOutlookFolder), 3)
    Local $bIncludeSubfolders = (BitAND(GUICtrlRead($idChkIncludeSubfolders), $GUI_CHECKED) = $GUI_CHECKED)

    If $sStoreName = "" Then
        MsgBox(16, "Erreur", "Sélectionne un compte Outlook.")
        Return
    EndIf
    If $sTargetEmail = "" Then
        MsgBox(16, "Erreur", "Renseigne l'adresse e-mail cible.")
        Return
    EndIf
    If Not FileExists($sExcel) Then
        MsgBox(16, "Erreur", "Fichier Excel invalide.")
        Return
    EndIf
    If $sEtmsSession = "" Then
        MsgBox(16, "Erreur", "Renseigne la session ETMS.")
        Return
    EndIf

    _SaveGuiSettings($sTargetEmail, $sStoreName, $sExcel, $sPeriod, $sEtmsSession, $bAnticipe, $sOutlookFolderPath, $bIncludeSubfolders)

    GUICtrlSetState($idBtnTest, $GUI_DISABLE)
    GUICtrlSetState($idBtnStart, $GUI_DISABLE)

    GUICtrlSetData($idConsoleLog, "")
    _SetProgress(0, "0 %")

    _LogInfo("--- ANALYSE ---")
    _LogInfo("Compte Outlook : " & $sStoreName)
    _LogInfo("Email cible : " & $sTargetEmail)
    _LogInfo("Période : " & _DescribePeriod($sPeriod))
    _LogInfo("Session ETMS : " & $sEtmsSession)
    _LogInfo("Mode anticipé : " & _BoolToText($bAnticipe))
    If $sOutlookFolderPath = "" Then
        _LogInfo("Dossier Outlook : Boîte de réception")
    Else
        _LogInfo("Dossier Outlook : " & $sOutlookFolderPath)
    EndIf
    _LogInfo("Sous-dossiers : " & _BoolToText($bIncludeSubfolders))

    _SetStatus("Chargement Excel...", $CLR_ACCENT)
    Local $oIndex = _ChargerIndexExcel($sExcel)
    If Not IsObj($oIndex) Then
        _LogErr("Impossible de charger l'index Excel.")
        _SetStatus("Erreur Excel", $CLR_ERR)
        _SetProgress(0, "0 %")
        GUICtrlSetState($idBtnTest, $GUI_ENABLE)
        GUICtrlSetState($idBtnStart, $GUI_ENABLE)
        Return
    EndIf
    _SetProgress(10, "10 %")

    _SetStatus("Analyse Outlook (PDF)...", $CLR_WARN)
    Local $sLastScanAt = IniRead($g_sConfigFile, "Settings", "LastScanAt", "")
    Local $aMatches = _AnalyserOutlookDepuisFolder($g_oNamespace, $sStoreName, $sOutlookFolderPath, $bIncludeSubfolders, $oIndex, $sTargetEmail, $sPeriod, $sLastScanAt)
    _SetProgress(55, "55 %")

    ; Préparer un tableau "report rows" (avec statut PENDING)
    Local $sMode = "TEST"
    If $bValidateWithF1 Then $sMode = "PROD"

    Local $sFolderShown = $sOutlookFolderPath
    If $sFolderShown = "" Then $sFolderShown = "Boîte de réception"

    ; Rapport unique (un seul fichier)
    Local $sReportFile = @ScriptDir & "\Rapport_Correspondances_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".xlsx"

    Local $aReportRows = _BuildReportRows($aMatches, $sMode, $sStoreName, $sFolderShown, $bIncludeSubfolders)

    If _DemanderConfirmationCorrespondances($aMatches) Then
        _SetStatus("Traitement...", $CLR_ACCENT)
        _TraiterCorrespondances($aMatches, $aReportRows, $sEtmsSession, $bAnticipe, $bValidateWithF1)
        _SetProgress(95, "95 %")
        _SetStatus("Génération du rapport Excel...", $CLR_ACCENT)
        _ExportReportToExcel($aReportRows, $sReportFile)
        _SetProgress(100, "100 %")
        _SetStatus("Terminé. Rapport généré.", $CLR_OK)
        _LogOk("Rapport : " & $sReportFile)
    Else
        _LogWarn("Traitement annulé.")
        _CleanupMatchesPDFs($aMatches)

        ; Marquer statut annulé dans le rapport et l'exporter quand même
        _MarkAllReportStatus($aReportRows, "ANNULÉ", "Traitement annulé par l'utilisateur.")
        _ExportReportToExcel($aReportRows, $sReportFile)
        _LogOk("Rapport (annulé) : " & $sReportFile)

        _SetStatus("Annulé", $CLR_WARN)
        _SetProgress(0, "0 %")
    EndIf

    GUICtrlSetState($idBtnTest, $GUI_ENABLE)
    GUICtrlSetState($idBtnStart, $GUI_ENABLE)
EndFunc
; =================================================================================
; OUTLOOK FOLDERS + ANALYSE
; =================================================================================

Func _AnalyserOutlookDepuisFolder($oNamespace, $sStoreName, $sFolderPath, $bIncludeSubfolders, $oExcelIndex, $sTargetSender, $sPeriod, $sLastScanAt)
    Local $aMatches[0]
    Local $oSeen = ObjCreate("Scripting.Dictionary")
    If Not IsObj($oSeen) Then Return $aMatches

    Local $oRoot = _GetStoreRootFolder($oNamespace, $sStoreName)
    If Not IsObj($oRoot) Then
        _LogErr("Impossible d'accéder au store Outlook.")
        Return $aMatches
    EndIf

    Local $oStartFolder = _GetFolderFromPath($oRoot, $sFolderPath)
    If Not IsObj($oStartFolder) Then
        _LogErr("Dossier Outlook introuvable : " & $sFolderPath)
        Return $aMatches
    EndIf

    ; Comptage des dossiers (pour progression analyse)
    Local $iFolderTotal = 0
    _CountFolders($oStartFolder, $bIncludeSubfolders, $iFolderTotal)
    If $iFolderTotal < 1 Then $iFolderTotal = 1

    Local $iFolderCounter = 0
    _AnalyserFolderRecursif($oStartFolder, $oExcelIndex, $sTargetSender, $sPeriod, $sLastScanAt, $aMatches, $oSeen, $bIncludeSubfolders, $iFolderCounter, $iFolderTotal)

    _LogInfo("Analyse terminée : " & UBound($aMatches) & " correspondance(s).")
    Return $aMatches
EndFunc

Func _CountFolders($oFolder, $bIncludeSubfolders, ByRef $iCount)
    If Not IsObj($oFolder) Then Return
    $iCount += 1
    If $bIncludeSubfolders Then
        For $oSub In $oFolder.Folders
            _CountFolders($oSub, $bIncludeSubfolders, $iCount)
        Next
    EndIf
EndFunc

Func _AnalyserFolderRecursif($oFolder, $oExcelIndex, $sTargetSender, $sPeriod, $sLastScanAt, ByRef $aMatches, ByRef $oSeen, $bIncludeSubfolders, ByRef $iFolderCounter, $iFolderTotal)
    If Not IsObj($oFolder) Then Return

    $iFolderCounter += 1
    Local $iPctFolder = 10 + Int(45 * ($iFolderCounter / $iFolderTotal))
    If $iPctFolder > 55 Then $iPctFolder = 55
    _SetProgress($iPctFolder, $iPctFolder & " %")
    _LogInfo("Dossier Outlook : " & String($oFolder.Name))

    Local $oItems = $oFolder.Items
    If IsObj($oItems) Then
        $oItems.Sort("[ReceivedTime]", True)

        Local $iCount = 0
        _ResetComErr()
        $iCount = $oItems.Count
        If $g_bComErr Then $iCount = 0

        Local $sCutoff = _GetCutoffNormalized($sPeriod, $sLastScanAt)
        Local $sToday = StringLeft(_NormalizeDateForCompare(_NowCalc()), 10)

        For $i = 1 To $iCount
            Local $oMail = $oItems.Item($i)
            If Not IsObj($oMail) Then ContinueLoop
            If $oMail.Class <> 43 Then ContinueLoop

            Local $sRecv = String($oMail.ReceivedTime)
            Local $sRecvNorm = _NormalizeDateForCompare($sRecv)
            If $sRecvNorm <> "" Then
                If Not _IsInTimeWindow($sRecvNorm, $sPeriod, $sCutoff, $sToday) Then ContinueLoop
            EndIf

            If Not _SenderMatchesTarget($oMail, $sTargetSender) Then ContinueLoop
            If Not _MailHasPdfAttachments($oMail) Then ContinueLoop

            Local $sSubject = String($oMail.Subject)
            Local $sEntryID = String($oMail.EntryID)
            Local $sNumRec = _ExtractReceptionFromSubject($sSubject)
            If $sNumRec = "" Then ContinueLoop

            _LogInfo("Mail : " & _Shorten($sSubject, 110))

            For $oAtt In $oMail.Attachments
                Local $sFileName = String($oAtt.FileName)
                If StringRight(StringLower($sFileName), 4) <> ".pdf" Then ContinueLoop

                Local $sTmpPdf = _MakeTempPdfPath($sFileName)
                $oAtt.SaveAsFile($sTmpPdf)
                If Not FileExists($sTmpPdf) Then ContinueLoop

                Local $aRefs = _ExtraireRefsDepuisPdf($sTmpPdf)
                If UBound($aRefs) = 0 Then
                    FileDelete($sTmpPdf)
                    ContinueLoop
                EndIf

                For $j = 0 To UBound($aRefs) - 1
                    Local $sFacture = $aRefs[$j]
                    If $oExcelIndex.Exists($sFacture) Then
                        Local $sExcelData = String($oExcelIndex.Item($sFacture))
                        Local $aSplit = StringSplit($sExcelData, $g_sSep, 2)
                        If UBound($aSplit) >= 2 Then
                            Local $sLigne = $aSplit[0]
                            Local $sDossier = $aSplit[1]
                            Local $sSeenKey = $sEntryID & "#" & $sFacture

                            If Not $oSeen.Exists($sSeenKey) Then
                                $oSeen.Item($sSeenKey) = 1

                                Local $sKeepPdf = _MakeTempPdfPath($sFileName)
                                FileCopy($sTmpPdf, $sKeepPdf, 9)

                                Local $sMatch = $sLigne & $g_sSep & _
                                                $sFacture & $g_sSep & _
                                                $sDossier & $g_sSep & _
                                                $sNumRec & $g_sSep & _
                                                $sSubject & $g_sSep & _
                                                $sKeepPdf & $g_sSep & _
                                                $sFileName & $g_sSep & _
                                                $sEntryID

                                _PushArray($aMatches, $sMatch)
                                _LogOk("Match : " & $sFacture & " -> " & $sDossier)
                            EndIf
                        EndIf
                    EndIf
                Next

                FileDelete($sTmpPdf)
            Next
        Next
    EndIf

    If $bIncludeSubfolders Then
        For $oSub In $oFolder.Folders
            _AnalyserFolderRecursif($oSub, $oExcelIndex, $sTargetSender, $sPeriod, $sLastScanAt, $aMatches, $oSeen, $bIncludeSubfolders, $iFolderCounter, $iFolderTotal)
        Next
    EndIf
EndFunc

Func _GetStoreRootFolder($oNamespace, $sStoreName)
    Local $oStoreTarget = 0
    For $oStore In $oNamespace.Stores
        If $oStore.DisplayName = $sStoreName Then
            $oStoreTarget = $oStore
            ExitLoop
        EndIf
    Next
    If IsObj($oStoreTarget) Then Return $oStoreTarget.GetRootFolder()
    Return 0
EndFunc

Func _GetFolderFromPath($oRootFolder, $sFolderPath)
    If Not IsObj($oRootFolder) Then Return 0

    ; Si vide -> Inbox
    If $sFolderPath = "" Then
        Return _GetInboxChildFolder($oRootFolder)
    EndIf

    Local $aParts = StringSplit($sFolderPath, "\", 2)
    Local $oCurrent = $oRootFolder

    For $i = 0 To UBound($aParts) - 1
        Local $sPart = StringStripWS($aParts[$i], 3)
        If $sPart = "" Then ContinueLoop

        Local $bFound = False
        For $oSub In $oCurrent.Folders
            If StringLower(String($oSub.Name)) = StringLower($sPart) Then
                $oCurrent = $oSub
                $bFound = True
                ExitLoop
            EndIf
        Next

        If Not $bFound Then Return 0
    Next

    Return $oCurrent
EndFunc

Func _GetInboxChildFolder($oRootFolder)
    For $oSub In $oRootFolder.Folders
        Local $sName = StringLower(String($oSub.Name))
        If $sName = "boîte de réception" Or $sName = "boite de reception" Or $sName = "inbox" Then
            Return $oSub
        EndIf
    Next
    Return 0
EndFunc


; =================================================================================
; CONFIRM + TRAITEMENT + REPORT (1 fichier)
; =================================================================================

Func _DemanderConfirmationCorrespondances(ByRef $aMatches)
    If UBound($aMatches) = 0 Then
        MsgBox(48, "Aucune correspondance", "Aucune correspondance trouvée sur la période choisie.")
        Return False
    EndIf

    Local $sMsg = UBound($aMatches) & " correspondance(s) trouvée(s) :" & @CRLF & @CRLF
    Local $iMax = UBound($aMatches) - 1
    If $iMax > 14 Then $iMax = 14

    For $i = 0 To $iMax
        Local $a = StringSplit($aMatches[$i], $g_sSep, 2)
        If UBound($a) >= 8 Then
            $sMsg &= "Ligne " & $a[0] & " | Ref " & $a[1] & " | Dossier " & $a[2] & " | Rec " & $a[3] & @CRLF
        EndIf
    Next

    If UBound($aMatches) > 15 Then
        $sMsg &= @CRLF & "... et " & (UBound($aMatches) - 15) & " autre(s)." & @CRLF
    EndIf

    $sMsg &= @CRLF & "Voulez-vous lancer le traitement ?"
    Local $iRep = MsgBox(36, "Confirmation", $sMsg)
    Return ($iRep = 6)
EndFunc

Func _TraiterCorrespondances(ByRef $aMatches, ByRef $aReportRows, $sEtmsSession, $bAnticipe, $bValidateWithF1)
    Local $sHistory = IniRead($g_sConfigFile, "History", "Processed", "|")
    Local $iTotal = UBound($aMatches)
    If $iTotal < 1 Then Return

    For $i = 0 To $iTotal - 1
        Local $a = StringSplit($aMatches[$i], $g_sSep, 2)
        If UBound($a) < 8 Then ContinueLoop

        Local $pct = 60 + Int((($i + 1) / $iTotal) * 35)
        If $pct > 95 Then $pct = 95
        _SetProgress($pct, $pct & " %")

        Local $sLigne  = $a[0]
        Local $sFacture= $a[1]
        Local $sDossier= $a[2]
        Local $sNumRec = $a[3]
        Local $sSujet  = $a[4]
        Local $sPdf    = $a[5]
        Local $sPdfName= $a[6]
        Local $sEntry  = $a[7]

        Local $sHistKey = $sEntry & "#" & $sFacture

        _SetStatus("Traitement dossier " & $sDossier & "...", $CLR_ACCENT)

        If StringInStr($sHistory, "|" & $sHistKey & "|") Then
            _LogWarn("Déjà traité : " & $sFacture)
            _ReportSetStatus($aReportRows, $sFacture, $sEntry, "SKIP", "Déjà traité")
            If FileExists($sPdf) Then FileDelete($sPdf)
            ContinueLoop
        EndIf

        If Not _PrePositionnerEDOC($sDossier) Then
            _LogErr("EDOC : impossible ouvrir dossier " & $sDossier)
            _ReportSetStatus($aReportRows, $sFacture, $sEntry, "ERREUR", "EDOC ouverture dossier")
            If FileExists($sPdf) Then FileDelete($sPdf)
            ContinueLoop
        EndIf

        If Not FileExists($sPdf) Then
            _LogErr("PDF introuvable : " & $sPdfName)
            _ReportSetStatus($aReportRows, $sFacture, $sEntry, "ERREUR", "PDF introuvable")
            ContinueLoop
        EndIf

        If Not _UploadPdfToEdoc($sNumRec, $sPdf) Then
            _LogErr("EDOC : upload échoué")
            _ReportSetStatus($aReportRows, $sFacture, $sEntry, "ERREUR", "Upload EDOC échoué")
            FileDelete($sPdf)
            ContinueLoop
        EndIf

        FileDelete($sPdf)

        If Not _RemplirETMS($sEtmsSession, $sDossier, $sNumRec, $bAnticipe, $bValidateWithF1) Then
            _LogErr("ETMS : remplissage échoué")
            _ReportSetStatus($aReportRows, $sFacture, $sEntry, "ERREUR", "ETMS échoué")
            ContinueLoop
        EndIf

        $sHistory &= $sHistKey & "|"
        IniWrite($g_sConfigFile, "History", "Processed", $sHistory)

        _LogOk("OK : " & $sFacture & " -> " & $sDossier)
        _ReportSetStatus($aReportRows, $sFacture, $sEntry, "OK", "")
    Next

    Local $sNow = _NowCalc()
    IniWrite($g_sConfigFile, "Settings", "LastScanAt", $sNow)
EndFunc


; =================================================================================
; REPORT DATA (1 file)
; =================================================================================

Func _BuildReportRows(ByRef $aMatches, $sMode, $sStoreName, $sFolderPath, $bSubfolders)
    ; Rows format:
    ; [0] timestamp [1] mode [2] store [3] folder [4] subfolders [5] ref [6] excelRow [7] dossier [8] numRec
    ; [9] pdfName [10] subject [11] entryID [12] status [13] message
    Local $aRows[0][14]
    Local $sNow = _NowCalc()
    Local $sSubs = "Non"
    If $bSubfolders Then $sSubs = "Oui"

    For $i = 0 To UBound($aMatches) - 1
        Local $a = StringSplit($aMatches[$i], $g_sSep, 2)
        If UBound($a) < 8 Then ContinueLoop

        Local $r = UBound($aRows)
        ReDim $aRows[$r + 1][14]

        $aRows[$r][0] = $sNow
        $aRows[$r][1] = $sMode
        $aRows[$r][2] = $sStoreName
        $aRows[$r][3] = $sFolderPath
        $aRows[$r][4] = $sSubs

        $aRows[$r][5] = $a[1] ; ref
        $aRows[$r][6] = $a[0] ; excel line
        $aRows[$r][7] = $a[2] ; dossier
        $aRows[$r][8] = $a[3] ; numRec

        $aRows[$r][9]  = $a[6] ; pdf name
        $aRows[$r][10] = $a[4] ; subject
        $aRows[$r][11] = $a[7] ; entryID

        $aRows[$r][12] = "PENDING"
        $aRows[$r][13] = ""
    Next

    Return $aRows
EndFunc

Func _ReportSetStatus(ByRef $aRows, $sRef, $sEntryID, $sStatus, $sMsg)
    For $i = 0 To UBound($aRows) - 1
        If $aRows[$i][5] = $sRef And $aRows[$i][11] = $sEntryID Then
            $aRows[$i][12] = $sStatus
            $aRows[$i][13] = $sMsg
            ExitLoop
        EndIf
    Next
EndFunc

Func _MarkAllReportStatus(ByRef $aRows, $sStatus, $sMsg)
    For $i = 0 To UBound($aRows) - 1
        $aRows[$i][12] = $sStatus
        $aRows[$i][13] = $sMsg
    Next
EndFunc

Func _ExportReportToExcel(ByRef $aRows, $sOutFile)
    Local $oExcel = ObjCreate("Excel.Application")
    If Not IsObj($oExcel) Then
        _LogErr("Export rapport impossible : Excel (COM) indisponible.")
        Return False
    EndIf

    $oExcel.DisplayAlerts = False
    $oExcel.Visible = False

    Local $oWB = $oExcel.Workbooks.Add()
    Local $oWS = $oWB.ActiveSheet
    $oWS.Name = "Correspondances"

    Local $aHeaders[14] = [ _
        "Horodatage","Mode","Compte Outlook","Dossier Outlook","Sous-dossiers", _
        "Référence (PDF)","Ligne Excel","Dossier (col G)","Numéro réception", _
        "Nom PDF","Sujet mail","EntryID","Statut","Message" _
    ]

    For $c = 0 To 13
        $oWS.Cells(1, $c + 1).Value = $aHeaders[$c]
        $oWS.Cells(1, $c + 1).Font.Bold = True
    Next

    For $i = 0 To UBound($aRows) - 1
        Local $r = $i + 2
        For $c = 0 To 13
            $oWS.Cells($r, $c + 1).Value = $aRows[$i][$c]
        Next
    Next

    $oWS.Columns("A:N").AutoFit()
    $oWS.Application.ActiveWindow.SplitRow = 1
    $oWS.Application.ActiveWindow.FreezePanes = True

    $oWB.SaveAs($sOutFile)
    $oWB.Close(True)
    $oExcel.Quit()

    Return True
EndFunc


; =================================================================================
; PDF / WORD / REFS
; =================================================================================

Func _PdfToTextViaWord($sPdfFile)
    If Not FileExists($sPdfFile) Then Return ""

    Local $oWord = ObjCreate("Word.Application")
    If Not IsObj($oWord) Then Return ""

    Local $oDoc = 0
    Local $sText = ""

    _ResetComErr()
    $oWord.Visible = False
    $oWord.DisplayAlerts = 0

    _ResetComErr()
    $oDoc = $oWord.Documents.Open($sPdfFile, False, True)

    If Not $g_bComErr And IsObj($oDoc) Then
        Sleep($g_iWordOpenMs)

        _ResetComErr()
        $sText = String($oDoc.Content.Text)
        If $g_bComErr Then $sText = ""

        _ResetComErr()
        $oDoc.Close(False)
    EndIf

    _ResetComErr()
    $oWord.Quit()

    Return $sText
EndFunc

Func _ExtraireRefsDepuisPdf($sPdfFile)
    Local $aResult[0]
    Local $sText = _PdfToTextViaWord($sPdfFile)
    If $sText = "" Then Return $aResult

    $sText = StringReplace($sText, @CRLF, " ")
    $sText = StringReplace($sText, @CR, " ")
    $sText = StringReplace($sText, @LF, " ")
    $sText = StringReplace($sText, @TAB, " ")

    ; 8 chiffres commençant par 1
    Local $aNums = StringRegExp($sText, "\b1\d{7}\b", 3)
    If @error Or Not IsArray($aNums) Then Return $aResult

    Local $oSeen = ObjCreate("Scripting.Dictionary")
    If Not IsObj($oSeen) Then Return $aResult

    For $i = 0 To UBound($aNums) - 1
        Local $sNum = _NormalizeRef($aNums[$i])
        If StringRegExp($sNum, "^1\d{7}$") Then
            If Not $oSeen.Exists($sNum) Then
                $oSeen.Item($sNum) = 1
                _PushArray($aResult, $sNum)
            EndIf
        EndIf
    Next

    Return $aResult
EndFunc


; =================================================================================
; EXCEL INDEX
; =================================================================================

Func _ChargerIndexExcel($sExcelFile)
    Local $oIndex = ObjCreate("Scripting.Dictionary")
    If Not IsObj($oIndex) Then Return 0

    Local $oExcel = ObjCreate("Excel.Application")
    If Not IsObj($oExcel) Then Return 0

    $oExcel.DisplayAlerts = False
    $oExcel.Visible = False

    Local $oWB = $oExcel.Workbooks.Open($sExcelFile)
    If Not IsObj($oWB) Then
        $oExcel.Quit()
        Return 0
    EndIf

    Local $oSheet = $oWB.ActiveSheet
    Local Const $xlUp = -4162
    Local $iLastRow = $oSheet.Cells($oSheet.Rows.Count, "B").End($xlUp).Row

    For $i = 2 To $iLastRow
        Local $sFacture = _NormalizeRef(String($oSheet.Cells($i, "B").Value))
        Local $sDossier = StringStripWS(String($oSheet.Cells($i, "G").Value), 3)

        If $sFacture = "" Then ContinueLoop
        If $sDossier = "" Then ContinueLoop

        If Not StringRegExp($sFacture, "^1\d{7}$") Then ContinueLoop

        If Not $oIndex.Exists($sFacture) Then
            $oIndex.Item($sFacture) = $i & $g_sSep & $sDossier
        EndIf
    Next

    $oWB.Close(False)
    $oExcel.Quit()

    _LogOk("Index Excel : " & $oIndex.Count & " réf.")
    Return $oIndex
EndFunc


; =================================================================================
; OUTLOOK SENDER / RECEPTION
; =================================================================================

Func _ExtractReceptionFromSubject($sSubject)
    Local $aMatch = StringRegExp(String($sSubject), "(?i)\bRECEPTION\s*(?:NO|N°|NUMERO)?\s*(\d{5,10})\b", 3)
    If @error Or Not IsArray($aMatch) Then Return ""
    Return $aMatch[0]
EndFunc

Func _SenderMatchesTarget($oMail, $sTargetSender)
    Local $sTarget = StringLower(StringStripWS($sTargetSender, 3))
    If $sTarget = "" Then Return False

    Local $sSMTP = StringLower(_GetMailSenderSmtp($oMail))
    Local $sEmailRaw = ""
    Local $sName = ""

    _ResetComErr()
    $sEmailRaw = StringLower(String($oMail.SenderEmailAddress))
    If $g_bComErr Then $sEmailRaw = ""

    _ResetComErr()
    $sName = StringLower(String($oMail.SenderName))
    If $g_bComErr Then $sName = ""

    If $sSMTP <> "" And StringInStr($sSMTP, $sTarget) Then Return True
    If $sEmailRaw <> "" And StringInStr($sEmailRaw, $sTarget) Then Return True
    If $sName <> "" And StringInStr($sName, $sTarget) Then Return True

    Return False
EndFunc

Func _GetMailSenderSmtp($oMail)
    Local $sSMTP = ""
    Local $sType = ""

    If Not IsObj($oMail) Then Return ""

    _ResetComErr()
    $sType = String($oMail.SenderEmailType)
    If $g_bComErr Then $sType = ""

    If StringUpper($sType) = "EX" Then
        _ResetComErr()
        Local $oSender = $oMail.Sender
        If Not $g_bComErr And IsObj($oSender) Then
            _ResetComErr()
            Local $oExUser = $oSender.GetExchangeUser()
            If Not $g_bComErr And IsObj($oExUser) Then
                _ResetComErr()
                $sSMTP = String($oExUser.PrimarySmtpAddress)
                If Not $g_bComErr And $sSMTP <> "" Then Return $sSMTP
            EndIf
        EndIf
    EndIf

    _ResetComErr()
    $sSMTP = String($oMail.SenderEmailAddress)
    If Not $g_bComErr And $sSMTP <> "" Then Return $sSMTP

    Return ""
EndFunc

Func _MailHasPdfAttachments($oMail)
    If Not IsObj($oMail) Then Return False
    Local $iCount = 0
    _ResetComErr()
    $iCount = $oMail.Attachments.Count
    If $g_bComErr Then Return False
    If $iCount <= 0 Then Return False

    For $oAtt In $oMail.Attachments
        Local $sFileName = String($oAtt.FileName)
        If StringRight(StringLower($sFileName), 4) = ".pdf" Then Return True
    Next
    Return False
EndFunc


; =================================================================================
; EDOC + ETMS
; =================================================================================

Func _PrePositionnerEDOC($sDoss)
    Local $hEdoc = _WaitWindowSmart("edoc Viewer CDG", "", 12, True)
    If $hEdoc = 0 Then Return False

    WinActivate($hEdoc)
    If Not WinWaitActive($hEdoc, "", 6) Then Return False

    If Not _SetControlTextRobust($hEdoc, "Edit1", $sDoss, 8) Then Return False
    ControlSend($hEdoc, "", "Edit1", "{ENTER}")

    Sleep($g_iEdocPostEnterMs)
    Return True
EndFunc

Func _UploadPdfToEdoc($sRec, $sFile)
    ShellExecute($sFile, "", "", "print")

    Local $hUpload = _WaitWindowSmart("Upload Documents CDG", "", 20, True)
    If $hUpload = 0 Then Return False

    WinActivate($hUpload)
    If Not WinWaitActive($hUpload, "", 6) Then Return False

    If Not _WaitControlExists($hUpload, "Edit2", 10) Then Return False
    If Not _WaitControlExists($hUpload, "TScaleEdit1", 10) Then Return False
    If Not _WaitControlExists($hUpload, "TButton2", 10) Then Return False

    Sleep($g_iUploadPostReadyMs)

    If Not _SetControlTextRobust($hUpload, "Edit2", "On-Hand Sheet", 8) Then Return False
    If Not _SetControlTextRobust($hUpload, "TScaleEdit1", $sRec, 8) Then Return False

    ControlClick($hUpload, "", "TButton2")
    Sleep(1200)

    Return True
EndFunc

Func _RemplirETMS($sSession, $sDossier, $sNumRec, $bAnticipe, $bValidate)
    Local $sEtmsTitle = "EXPORT CDG.CDG.EI " & StringUpper($sSession) & ".EXPORT"
    Local $hEtms = _WaitWindowSmart($sEtmsTitle, "", 12, True)
    If $hEtms = 0 Then Return False

    WinActivate($hEtms)
    If Not WinWaitActive($hEtms, "", 6) Then Return False

    Send("{PGUP}")
    Sleep($g_iEtmsAfterPageUpMs)

    Send($sDossier)
    Sleep(250)

    Send("{F2}")
    Sleep($g_iEtmsAfterF2Ms)

    If Not _WaitControlExists($hEtms, "TEIEdit20", 10) Then Return False

    Local $sDateEtms = _GetEtmsDate($bAnticipe)

    ; TEIEdit20 = num réception
    If Not _SetControlTextRobust($hEtms, "TEIEdit20", $sNumRec, 5) Then Return False

    ; ✅ AJOUT V26.4 : TEIEdit16 = N
    If Not _SetControlTextRobust($hEtms, "TEIEdit16", $g_sEtms16, 5) Then Return False

    If Not _SetControlTextRobust($hEtms, "TEIEdit22", $g_sEtms22, 5) Then Return False
    If Not _SetControlTextRobust($hEtms, "TEIEdit23", $g_sEtms23, 5) Then Return False
    If Not _SetControlTextRobust($hEtms, "TEIEdit24", $g_sEtms24, 5) Then Return False
    If Not _SetControlTextRobust($hEtms, "TEIEdit25", $g_sEtms25, 5) Then Return False
    If Not _SetControlTextRobust($hEtms, "TEIEdit26", $g_sEtms26, 5) Then Return False
    If Not _SetControlTextRobust($hEtms, "TEIEdit27", $g_sEtms27, 5) Then Return False
    If Not _SetControlTextRobust($hEtms, "TEIEdit28", $sDateEtms, 5) Then Return False

    If $bValidate Then
        Send("{F1}")
        Sleep(600)
    EndIf

    Return True
EndFunc


; =================================================================================
; SETTINGS SAVE
; =================================================================================
Func _SaveGuiSettings($sTargetEmail, $sStoreName, $sExcel, $sPeriod, $sEtmsSession, $bAnticipe, $sFolderPath, $bIncludeSub)
    IniWrite($g_sConfigFile, "Settings", "TargetEmail", $sTargetEmail)
    IniWrite($g_sConfigFile, "Settings", "LastAccount", $sStoreName)
    IniWrite($g_sConfigFile, "Settings", "ExcelPath", $sExcel)
    IniWrite($g_sConfigFile, "Settings", "Period", $sPeriod)
    IniWrite($g_sConfigFile, "Settings", "EtmsSession", $sEtmsSession)
    IniWrite($g_sConfigFile, "Settings", "Anticipe", _BoolToIni($bAnticipe))
    IniWrite($g_sConfigFile, "Settings", "OutlookFolderPath", $sFolderPath)
    IniWrite($g_sConfigFile, "Settings", "IncludeSubFolders", _BoolToIni($bIncludeSub))
EndFunc


; =================================================================================
; HELPERS / UI / FILES / DATES
; =================================================================================

Func _ComErrHandler($oError)
    $g_bComErr = True
    Return SetError(1, 0, 0)
EndFunc

Func _ResetComErr()
    $g_bComErr = False
EndFunc

Func _UpdateLog($sText)
    Local $sCurrent = GUICtrlRead($idConsoleLog)
    Local $sTime = StringFormat("%02d:%02d:%02d", @HOUR, @MIN, @SEC)
    GUICtrlSetData($idConsoleLog, $sCurrent & "[" & $sTime & "] " & $sText & @CRLF)
    _GUICtrlEdit_LineScroll($idConsoleLog, 0, _GUICtrlEdit_GetLineCount($idConsoleLog))
EndFunc

Func _LogInfo($sText)
    _UpdateLog($sText)
EndFunc
Func _LogOk($sText)
    _UpdateLog("OK - " & $sText)
EndFunc
Func _LogWarn($sText)
    _UpdateLog("ATTENTION - " & $sText)
EndFunc
Func _LogErr($sText)
    _UpdateLog("ERREUR - " & $sText)
EndFunc

Func _SetStatus($sText, $iColor = 0x000000)
    GUICtrlSetData($idStatusLabel, $sText)
    GUICtrlSetColor($idStatusLabel, $iColor)
EndFunc

Func _SetProgress($iValue, $sText = "")
    If $iValue < 0 Then $iValue = 0
    If $iValue > 100 Then $iValue = 100
    GUICtrlSetData($idProgressBar, $iValue)
    If $sText = "" Then $sText = $iValue & " %"
    GUICtrlSetData($idProgressText, $sText)
EndFunc

Func _NormalizeRef($sText)
    $sText = String($sText)
    $sText = StringRegExpReplace($sText, "\D", "")
    Return $sText
EndFunc

Func _PushArray(ByRef $aArray, $vValue)
    Local $iSize = UBound($aArray)
    ReDim $aArray[$iSize + 1]
    $aArray[$iSize] = $vValue
EndFunc

Func _Shorten($sText, $iMaxLen = 120)
    $sText = String($sText)
    If StringLen($sText) <= $iMaxLen Then Return $sText
    Return StringLeft($sText, $iMaxLen - 3) & "..."
EndFunc

Func _BoolToText($bValue)
    If $bValue Then Return "Oui"
    Return "Non"
EndFunc

Func _BoolToIni($bValue)
    If $bValue Then Return "1"
    Return "0"
EndFunc

Func _MakeTempPdfPath($sFileName)
    Local $sSafeName = StringReplace($sFileName, "\", "_")
    $sSafeName = StringReplace($sSafeName, "/", "_")
    $sSafeName = StringReplace($sSafeName, ":", "_")
    $sSafeName = StringReplace($sSafeName, "*", "_")
    $sSafeName = StringReplace($sSafeName, "?", "_")
    $sSafeName = StringReplace($sSafeName, '"', "_")
    $sSafeName = StringReplace($sSafeName, "<", "_")
    $sSafeName = StringReplace($sSafeName, ">", "_")
    $sSafeName = StringReplace($sSafeName, "|", "_")

    Return $g_sTempFolder & "\" & _
           @YEAR & @MON & @MDAY & "_" & _
           @HOUR & @MIN & @SEC & "_" & _
           Int(Random(1000, 9999, 1)) & "_" & $sSafeName
EndFunc

Func _CleanupMatchesPDFs(ByRef $aMatches)
    For $i = 0 To UBound($aMatches) - 1
        Local $a = StringSplit($aMatches[$i], $g_sSep, 2)
        If UBound($a) >= 6 Then
            If FileExists($a[5]) Then FileDelete($a[5])
        EndIf
    Next
EndFunc

Func _DescribePeriod($sPeriod)
    Switch $sPeriod
        Case "Aujourd'hui"
            Return "Aujourd'hui"
        Case "1 heure"
            Return "Dernière 1 heure"
        Case "2 heures"
            Return "Dernières 2 heures"
        Case "4 heures"
            Return "Dernières 4 heures"
        Case "8 heures"
            Return "Dernières 8 heures"
        Case "24 heures"
            Return "Dernières 24 heures"
        Case "7 jours"
            Return "Derniers 7 jours"
        Case "Depuis dernier scan"
            Return "Depuis le dernier scan"
        Case Else
            Return $sPeriod
    EndSwitch
EndFunc

Func _GetCutoffNormalized($sPeriod, $sLastScanAt)
    Local $sBase = _NowCalc()
    Switch $sPeriod
        Case "Aujourd'hui"
            Return StringLeft(_NormalizeDateForCompare($sBase), 10) & " 00:00:00"
        Case "1 heure"
            Return _NormalizeDateForCompare(_DateAdd("h", -1, $sBase))
        Case "2 heures"
            Return _NormalizeDateForCompare(_DateAdd("h", -2, $sBase))
        Case "4 heures"
            Return _NormalizeDateForCompare(_DateAdd("h", -4, $sBase))
        Case "8 heures"
            Return _NormalizeDateForCompare(_DateAdd("h", -8, $sBase))
        Case "24 heures"
            Return _NormalizeDateForCompare(_DateAdd("h", -24, $sBase))
        Case "7 jours"
            Return _NormalizeDateForCompare(_DateAdd("D", -7, $sBase))
        Case "Depuis dernier scan"
            If $sLastScanAt <> "" Then
                Return _NormalizeDateForCompare($sLastScanAt)
            Else
                Return StringLeft(_NormalizeDateForCompare($sBase), 10) & " 00:00:00"
            EndIf
        Case Else
            Return "1900/01/01 00:00:00"
    EndSwitch
EndFunc

Func _IsInTimeWindow($sRecvNorm, $sPeriod, $sCutoff, $sToday)
    If $sRecvNorm = "" Then Return True
    Switch $sPeriod
        Case "Aujourd'hui"
            Return (StringLeft($sRecvNorm, 10) = $sToday)
        Case "1 heure", "2 heures", "4 heures", "8 heures", "24 heures", "7 jours", "Depuis dernier scan"
            Return ($sRecvNorm >= $sCutoff)
        Case Else
            Return True
    EndSwitch
EndFunc

Func _NormalizeDateForCompare($sDate)
    $sDate = StringStripWS(String($sDate), 3)
    If $sDate = "" Then Return ""

    Local $a

    If StringRegExp($sDate, "^\d{14}$") Then
        Local $yyyy = Number(StringLeft($sDate, 4))
        Local $mm   = Number(StringMid($sDate, 5, 2))
        Local $dd   = Number(StringMid($sDate, 7, 2))
        Local $hh   = Number(StringMid($sDate, 9, 2))
        Local $mi   = Number(StringMid($sDate, 11, 2))
        Local $ss   = Number(StringRight($sDate, 2))
        Return StringFormat("%04d/%02d/%02d %02d:%02d:%02d", $yyyy, $mm, $dd, $hh, $mi, $ss)
    EndIf

    $a = StringRegExp($sDate, "^\s*(\d{4})\D(\d{1,2})\D(\d{1,2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s*$", 3)
    If Not @error And IsArray($a) Then
        Local $ss = 0
        If UBound($a) > 5 And $a[5] <> "" Then $ss = Number($a[5])
        Return StringFormat("%04d/%02d/%02d %02d:%02d:%02d", Number($a[0]), Number($a[1]), Number($a[2]), Number($a[3]), Number($a[4]), $ss)
    EndIf

    $a = StringRegExp($sDate, "^\s*(\d{1,2})\D(\d{1,2})\D(\d{4})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s*$", 3)
    If Not @error And IsArray($a) Then
        Local $ss2 = 0
        If UBound($a) > 5 And $a[5] <> "" Then $ss2 = Number($a[5])
        Return StringFormat("%04d/%02d/%02d %02d:%02d:%02d", Number($a[2]), Number($a[1]), Number($a[0]), Number($a[3]), Number($a[4]), $ss2)
    EndIf

    Return ""
EndFunc

Func _GetEtmsDate($bAnticipe)
    Local $sBase
    If $bAnticipe Then
        $sBase = _NormalizeDateForCompare(_NowCalc())
    Else
        $sBase = _NormalizeDateForCompare(_DateAdd("D", -1, _NowCalc()))
    EndIf

    If StringLen($sBase) >= 10 Then
        Local $yyyy = StringLeft($sBase, 4)
        Local $mm = StringMid($sBase, 6, 2)
        Local $dd = StringMid($sBase, 9, 2)
        Return $dd & "." & $mm & "." & StringRight($yyyy, 2)
    EndIf
    Return ""
EndFunc

; =================================================================================
; WINDOW / CONTROL HELPERS
; =================================================================================
Func _WaitWindowSmart($sTitleContains, $sProcContains = "", $iTimeoutSec = 15, $bVisibleOnly = True)
    Local $t0 = TimerInit()
    While TimerDiff($t0) < ($iTimeoutSec * 1000)
        Local $hWnd = _GetWindowSmart($sTitleContains, $sProcContains, $bVisibleOnly)
        If $hWnd <> 0 Then Return $hWnd
        Sleep($g_iPollMs)
    WEnd
    Return 0
EndFunc

Func _GetWindowSmart($sTitleContains, $sProcContains = "", $bVisibleOnly = True)
    Local $aWindows = WinList()
    For $i = 1 To $aWindows[0][0]
        Local $sTitle = $aWindows[$i][0]
        Local $hWnd = $aWindows[$i][1]
        If $sTitle = "" Then ContinueLoop
        If $bVisibleOnly And Not BitAND(WinGetState($hWnd), 2) Then ContinueLoop
        If $sTitleContains <> "" Then
            If Not StringInStr(StringLower($sTitle), StringLower($sTitleContains)) Then ContinueLoop
        EndIf
        Return $hWnd
    Next
    Return 0
EndFunc

Func _WaitControlExists($hWnd, $sControlID, $iTimeoutSec = 10)
    Local $t0 = TimerInit()
    While TimerDiff($t0) < ($iTimeoutSec * 1000)
        Local $hCtrl = ControlGetHandle($hWnd, "", $sControlID)
        If $hCtrl <> 0 Then Return True
        Sleep($g_iPollMs)
    WEnd
    Return False
EndFunc

Func _SetControlTextRobust($hWnd, $sControlID, $sValue, $iTimeoutSec = 8)
    If Not _WaitControlExists($hWnd, $sControlID, $iTimeoutSec) Then Return False

    ControlFocus($hWnd, "", $sControlID)
    Sleep(120)
    ControlSetText($hWnd, "", $sControlID, "")
    Sleep(100)
    ControlSetText($hWnd, "", $sControlID, $sValue)
    Sleep(200)

    Local $sCheck = ControlGetText($hWnd, "", $sControlID)
    If StringStripWS(String($sCheck), 3) = StringStripWS(String($sValue), 3) Then Return True

    ControlFocus($hWnd, "", $sControlID)
    Sleep(100)
    Send("^a")
    Sleep(80)
    Send($sValue)
    Sleep(220)

    $sCheck = ControlGetText($hWnd, "", $sControlID)
    If StringStripWS(String($sCheck), 3) = StringStripWS(String($sValue), 3) Then Return True
    Return False
EndFunc


; =================================================================================
; HISTORIQUE
; =================================================================================
Func _AppendScanHistory($sSummary)
    For $i = 9 To 1 Step -1
        Local $sPrev = IniRead($g_sConfigFile, "ScanHistory", "Item" & $i, "")
        If $sPrev <> "" Then IniWrite($g_sConfigFile, "ScanHistory", "Item" & ($i + 1), $sPrev)
    Next
    IniWrite($g_sConfigFile, "ScanHistory", "Item1", $sSummary)
EndFunc

Func _RefreshScanHistoryControl()
    Local $sText = ""
    For $i = 1 To 10
        Local $sLine = IniRead($g_sConfigFile, "ScanHistory", "Item" & $i, "")
        If $sLine <> "" Then $sText &= $sLine & @CRLF
    Next
    GUICtrlSetData($idScanHistory, $sText)
EndFunc
