Attribute VB_Name = "modTreatmentDiaryImporter"
Option Explicit

Private Const ONEDRIVE_WEB_PREFIX As String = "https://d.docs.live.net/"

Private Const MASTER_SHEET_NAME As String = "治療日誌マスタ"
Private Const REFERENCE_SHEET_NAME As String = "店舗リファレンス"
Private Const EXEC_LOG_SHEET_NAME As String = "取込実行ログ"
Private Const DETAIL_LOG_SHEET_NAME As String = "取込詳細ログ"
Private Const TRANSFER_SHEET_NAME As String = "転記"
Private Const NEW_PATIENT_SHEET_NAME As String = "新患分析"
Private Const NEW_PATIENT_ID_COLUMN As Long = 2
Private Const NEW_PATIENT_NAME_COLUMN As Long = 3

Private Const DATE_ROW_INDEX As Long = 1
Private Const DATA_START_ROW As Long = 4
Private Const DATA_END_ROW As Long = 53
Private Const BLOCK_WIDTH As Long = 10
Private Const LAST_SOURCE_COLUMN As Long = 110   ' DF列

Private Type AppState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Calculation As XlCalculation
End Type

Private Type CustomerResolution
    CustomerId As String
    StatusText As String
    WarningMessage As String
End Type

Public Sub InitializeTreatmentDiaryWorkbook()
    On Error GoTo ErrHandler
    Dim referenceSheet As Worksheet
    Dim storeId As String
    Dim storeName As String

    EnsureSheetWithHeaders MASTER_SHEET_NAME, MasterHeaders()
    EnsureSheetWithHeaders REFERENCE_SHEET_NAME, ReferenceHeaders()
    EnsureSheetWithHeaders EXEC_LOG_SHEET_NAME, ExecutionLogHeaders()
    EnsureSheetWithHeaders DETAIL_LOG_SHEET_NAME, DetailLogHeaders()

    Set referenceSheet = ThisWorkbook.Worksheets(REFERENCE_SHEET_NAME)
    If TryParseStoreInfoFromWorkbookName(ThisWorkbook.Name, storeId, storeName) Then
        PopulateReferenceSheet referenceSheet, storeId, storeName
        MsgBox "初期化が完了しました。" & vbCrLf & _
               "店舗リファレンスへ以下を自動設定しました。" & vbCrLf & _
               "店舗ID: " & storeId & vbCrLf & _
               "店舗名: " & storeName, vbInformation
    Else
        MsgBox "初期化が完了しました。" & vbCrLf & _
               "ブック名から店舗情報を取得できなかったため、""店舗リファレンス"" シートを確認してください。", vbExclamation
    End If
    Exit Sub

ErrHandler:
    MsgBox "初期化に失敗しました。" & vbCrLf & Err.Description, vbCritical
End Sub

Public Sub ImportTreatmentDiariesFromFolder()
    Dim selectedFolder As String
    Dim filePaths As Collection

    selectedFolder = PickFolder()
    If Len(selectedFolder) = 0 Then
        Exit Sub
    End If

    Set filePaths = New Collection
    CollectXlsmFilesRecursive selectedFolder, filePaths

    If filePaths.Count = 0 Then
        MsgBox "指定フォルダに .xlsm ファイルが見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    RunImport filePaths, "フォルダ指定"
End Sub

Public Sub ImportTreatmentDiariesFromFiles()
    Dim dialog As FileDialog
    Dim filePaths As Collection
    Dim selectedItem As Variant
    Dim normalizedSelectedPath As String

    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    With dialog
        .Title = "治療日誌ファイルを選択してください"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel Macro Workbook", "*.xlsm"

        If .Show <> -1 Then
            Exit Sub
        End If
    End With

    Set filePaths = New Collection
    For Each selectedItem In dialog.SelectedItems
        normalizedSelectedPath = NormalizeWorkbookPath(CStr(selectedItem))
        If Not IsTemporaryExcelFile(normalizedSelectedPath) Then
            If Not FileExists(normalizedSelectedPath) Then
                MsgBox "選択した取込対象ファイルが見つかりません。" & vbCrLf & normalizedSelectedPath, vbCritical
                Exit Sub
            End If
            filePaths.Add normalizedSelectedPath
        End If
    Next selectedItem

    If filePaths.Count = 0 Then
        MsgBox "有効な .xlsm ファイルが選択されていません。", vbExclamation
        Exit Sub
    End If

    RunImport filePaths, "複数ファイル指定"
End Sub

Private Sub RunImport(ByVal filePaths As Collection, ByVal sourceLabel As String)
    Dim state As AppState
    Dim storeMap As Object
    Dim patientMap As Object
    Dim importKeys As Object
    Dim startedAt As Date
    Dim filePath As Variant
    Dim outputCount As Long
    Dim successFiles As Long
    Dim failedFiles As Long
    Dim totalOutputRows As Long
    Dim analysisWorkbookPath As String
    Dim importSequence As Long
    Dim stepName As String
    Dim currentPath As String
    Dim settingsApplied As Boolean
    Dim errorMessage As String
    Dim archivedLogPath As String

    On Error GoTo ErrHandler

    stepName = "新患分析ファイル選択"
    analysisWorkbookPath = PickPatientAnalysisWorkbook()
    If Len(analysisWorkbookPath) = 0 Then
        Exit Sub
    End If

    currentPath = analysisWorkbookPath
    If Not FileExists(analysisWorkbookPath) Then
        Err.Raise vbObjectError + 1010, , "選択した新患分析ファイルが見つかりません。 path=" & analysisWorkbookPath
    End If

    stepName = "Excel高速化設定"
    SaveAppState state
    ApplyFastSettings
    settingsApplied = True

    stepName = "初期シート確認"
    EnsureSheetWithHeaders MASTER_SHEET_NAME, MasterHeaders()
    EnsureSheetWithHeaders REFERENCE_SHEET_NAME, ReferenceHeaders()
    EnsureSheetWithHeaders EXEC_LOG_SHEET_NAME, ExecutionLogHeaders()
    EnsureSheetWithHeaders DETAIL_LOG_SHEET_NAME, DetailLogHeaders()
    stepName = "既存ログ退避"
    currentPath = ThisWorkbook.FullName
    archivedLogPath = ArchiveAndClearImportLogs()

    stepName = "店舗リファレンス読込"
    Set storeMap = LoadStoreMap(ThisWorkbook.Worksheets(REFERENCE_SHEET_NAME))
    If storeMap.Count = 0 Then
        Err.Raise vbObjectError + 1000, , "店舗リファレンスに有効な店舗情報がありません。"
    End If

    stepName = "新患分析読込"
    Set patientMap = LoadPatientMap(analysisWorkbookPath)
    If patientMap.Count = 0 Then
        Err.Raise vbObjectError + 1006, , "新患分析に有効な患者情報がありません。"
    End If

    stepName = "既存取込キー読込"
    Set importKeys = LoadExistingImportKeys(ThisWorkbook.Worksheets(MASTER_SHEET_NAME))
    startedAt = Now
    If Len(archivedLogPath) > 0 Then
        AddExecutionLogInfo ThisWorkbook.Worksheets(EXEC_LOG_SHEET_NAME), startedAt, "前回ログを退避しました。 path=" & archivedLogPath
    End If

    stepName = "ファイル取込"
    For Each filePath In filePaths
        outputCount = 0
        currentPath = CStr(filePath)

        If ProcessSingleFile(CStr(filePath), startedAt, storeMap, patientMap, importKeys, importSequence, outputCount) Then
            successFiles = successFiles + 1
            totalOutputRows = totalOutputRows + outputCount
        Else
            failedFiles = failedFiles + 1
        End If
    Next filePath

    If settingsApplied Then
        RestoreAppState state
    End If

    MsgBox sourceLabel & " の取り込みが完了しました。" & vbCrLf & _
           "成功: " & successFiles & "件" & vbCrLf & _
           "失敗: " & failedFiles & "件" & vbCrLf & _
           "出力件数: " & totalOutputRows & "件", vbInformation
    Exit Sub

ErrHandler:
    If settingsApplied Then
        RestoreAppState state
    End If

    errorMessage = "取り込み処理を開始できませんでした。" & vbCrLf & _
                   "処理: " & stepName & vbCrLf
    If Len(currentPath) > 0 Then
        errorMessage = errorMessage & "path=" & currentPath & vbCrLf
    End If
    errorMessage = errorMessage & Err.Description

    MsgBox errorMessage, vbCritical
End Sub

Private Function ProcessSingleFile( _
    ByVal filePath As String, _
    ByVal startedAt As Date, _
    ByVal storeMap As Object, _
    ByVal patientMap As Object, _
    ByVal importKeys As Object, _
    ByRef importSequence As Long, _
    ByRef outputCount As Long) As Boolean

    Dim sourceBook As Workbook
    Dim fileName As String
    Dim logicalFileName As String
    Dim storeId As String
    Dim storeName As String
    Dim masterRows As Collection
    Dim detailLogs As Collection
    Dim execLogs As Collection
    Dim targetSheetNameList As Variant
    Dim sheetName As Variant
    Dim currentSheetName As String
    Dim fileYearMonth As String
    Dim rowsRead As Long
    Dim rowsOutput As Long
    Dim warningCount As Long
    Dim statusText As String
    Dim messageText As String
    Dim duplicateKey As String
    Dim duplicateMessage As String
    Dim failureLogs As Collection
    Dim stepName As String
    Dim normalizedFilePath As String
    Dim openedFilePath As String
    Dim tempOpenPath As String
    Dim physicalFileName As String
    Dim openErrorMessage As String
    Dim baseYear As Long
    Dim baseMonth As Long
    Dim expectedPatientCount As Long

    On Error GoTo ErrHandler

    Set masterRows = New Collection
    Set detailLogs = New Collection
    Set execLogs = New Collection

    normalizedFilePath = NormalizeWorkbookPath(filePath)
    stepName = "ファイル名解析"
    logicalFileName = GetLogicalFileName(normalizedFilePath)
    physicalFileName = GetFileName(normalizedFilePath)
    fileName = physicalFileName
    storeId = NormalizeStoreIdText(ExtractStoreId(logicalFileName))
    Call ExtractYearMonthFromPath(normalizedFilePath, baseYear, baseMonth)
    If Len(storeId) = 0 Then
        Err.Raise vbObjectError + 1001, , "ファイル名から店舗IDを取得できません。"
    End If

    stepName = "店舗ID検証"
    If Not storeMap.Exists(storeId) Then
        Err.Raise vbObjectError + 1002, , "店舗リファレンスに店舗ID " & storeId & " が存在しません。"
    End If

    storeName = CStr(storeMap(storeId))

    stepName = "ブックオープン前確認"
    If Not FileExists(normalizedFilePath) Then
        Err.Raise vbObjectError + 1003, , "取込対象ファイルが見つかりません。 path=" & normalizedFilePath
    End If

    If physicalFileName <> logicalFileName Then
        stepName = "一時コピー作成"
        tempOpenPath = CreateTempWorkbookCopy(normalizedFilePath, logicalFileName)
        openedFilePath = tempOpenPath
    Else
        openedFilePath = normalizedFilePath
    End If

    stepName = "ブックオープン"
    Set sourceBook = TryOpenWorkbookReadOnly(openedFilePath, openErrorMessage)
    If sourceBook Is Nothing And openedFilePath = normalizedFilePath Then
        stepName = "一時コピー作成"
        tempOpenPath = CreateTempWorkbookCopy(normalizedFilePath, logicalFileName)
        openedFilePath = tempOpenPath
        stepName = "一時コピーオープン"
        Set sourceBook = TryOpenWorkbookReadOnly(openedFilePath, openErrorMessage)
    End If
    If sourceBook Is Nothing Then
        Err.Raise vbObjectError + 1004, , "取込対象ファイルを開けません。 path=" & openedFilePath & " / detail=" & openErrorMessage
    End If

    stepName = "対象シート一覧取得"
    targetSheetNameList = TargetSheetNames()

    For Each sheetName In targetSheetNameList
        currentSheetName = CStr(sheetName)

        stepName = "シート存在確認"
        If Not WorksheetExists(sourceBook, currentSheetName) Then
            AddDetailLogRow detailLogs, startedAt, fileName, currentSheetName, vbNullString, vbNullString, vbNullString, "エラー", "対象シートが存在しません。"
            messageText = "対象シートが存在しません。"
            GoTo FileError
        End If

        rowsRead = 0
        rowsOutput = 0
        warningCount = 0
        messageText = vbNullString

        stepName = "シート処理"
        If Not ProcessSheet( _
            sourceBook.Worksheets(currentSheetName), _
            fileName, _
            normalizedFilePath, _
            storeId, _
            storeName, _
            baseYear, _
            baseMonth, _
            startedAt, _
            patientMap, _
            importKeys, _
            importSequence, _
            masterRows, _
            detailLogs, _
            fileYearMonth, _
            rowsRead, _
            rowsOutput, _
            warningCount, _
            messageText) Then

            GoTo FileError
        End If

        If warningCount > 0 Then
            statusText = "WARNING"
            messageText = "警告 " & warningCount & " 件"
        Else
            statusText = "SUCCESS"
            messageText = "正常終了"
        End If

        AddExecutionLogRow execLogs, startedAt, fileName, currentSheetName, statusText, rowsRead, rowsOutput, messageText
    Next sheetName

    stepName = "転記件数確認"
    currentSheetName = TRANSFER_SHEET_NAME
    If Not ValidateTransferredCount(sourceBook, CLng(masterRows.Count), expectedPatientCount, messageText) Then
        AddDetailLogRow detailLogs, startedAt, fileName, currentSheetName, vbNullString, vbNullString, vbNullString, "エラー", messageText
        GoTo FileError
    End If

    If Len(fileYearMonth) > 0 Then
        duplicateKey = BuildImportKey(fileName, fileYearMonth)
        importKeys(duplicateKey) = True
    End If

    stepName = "マスタ出力"
    AppendRows ThisWorkbook.Worksheets(MASTER_SHEET_NAME), masterRows, UBound(MasterHeaders()) + 1
    stepName = "実行ログ出力"
    AppendRows ThisWorkbook.Worksheets(EXEC_LOG_SHEET_NAME), execLogs, UBound(ExecutionLogHeaders()) + 1
    stepName = "詳細ログ出力"
    AppendRows ThisWorkbook.Worksheets(DETAIL_LOG_SHEET_NAME), detailLogs, UBound(DetailLogHeaders()) + 1

    outputCount = masterRows.Count
    ProcessSingleFile = True
    GoTo SafeExit

FileError:
    Set failureLogs = New Collection
    AddExecutionLogRow failureLogs, startedAt, fileName, currentSheetName, "ERROR", 0, 0, messageText
    AppendRows ThisWorkbook.Worksheets(EXEC_LOG_SHEET_NAME), failureLogs, UBound(ExecutionLogHeaders()) + 1
    AppendRows ThisWorkbook.Worksheets(DETAIL_LOG_SHEET_NAME), detailLogs, UBound(DetailLogHeaders()) + 1

    duplicateMessage = "エラー：過去にこのファイルを出力しています。（" & fileName & "）"
    If InStr(1, messageText, duplicateMessage, vbTextCompare) > 0 Then
        MsgBox duplicateMessage & vbCrLf & _
               "該当データを削除してから再実行してください。", vbExclamation
    Else
        MsgBox fileName & " にエラーがあります。" & vbCrLf & _
               "ログを確認してください。" & vbCrLf & _
               "該当ファイルはマスタに反映されていません。修正後に再実行してください。", vbExclamation
    End If

    ProcessSingleFile = False
    GoTo SafeExit

ErrHandler:
    messageText = Err.Description
    If Len(stepName) > 0 Then
        messageText = messageText & " / 処理=" & stepName
    End If
    If Len(currentSheetName) > 0 Then
        messageText = messageText & " / シート=" & currentSheetName
    End If
    If stepName = "ブックオープン" Then
        messageText = messageText & " / path=" & openedFilePath
    End If
    AddDetailLogRow detailLogs, startedAt, fileName, currentSheetName, vbNullString, vbNullString, vbNullString, "エラー", messageText
    Resume FileError

SafeExit:
    If Not sourceBook Is Nothing Then
        sourceBook.Close SaveChanges:=False
    End If
    DeleteFileIfExists tempOpenPath
End Function

Private Function ProcessSheet( _
    ByVal sourceSheet As Worksheet, _
    ByVal fileName As String, _
    ByVal filePath As String, _
    ByVal storeId As String, _
    ByVal storeName As String, _
    ByVal baseYear As Long, _
    ByVal baseMonth As Long, _
    ByVal startedAt As Date, _
    ByVal patientMap As Object, _
    ByVal importKeys As Object, _
    ByRef importSequence As Long, _
    ByVal masterRows As Collection, _
    ByVal detailLogs As Collection, _
    ByRef fileYearMonth As String, _
    ByRef rowsRead As Long, _
    ByRef rowsOutput As Long, _
    ByRef warningCount As Long, _
    ByRef messageText As String) As Boolean

    Dim sourceValues As Variant
    Dim blockStart As Variant
    Dim baseColumn As Long
    Dim rowIndex As Long
    Dim treatmentTime As Variant
    Dim targetDate As Date
    Dim category As String
    Dim staffName As String
    Dim patientName As String
    Dim totalAmount As Double
    Dim currentYearMonth As String
    Dim feeColumn As Long
    Dim stepName As String
    Dim customerResolution As CustomerResolution
    Dim internalImportId As String

    On Error GoTo ErrHandler

    stepName = "シート読込"
    sourceValues = sourceSheet.Range(sourceSheet.Cells(1, 1), sourceSheet.Cells(DATA_END_ROW, LAST_SOURCE_COLUMN)).Value2

    For Each blockStart In BlockStartColumns()
        baseColumn = CLng(blockStart)

        stepName = "ブロック判定"
        If BlockHasInputRows(sourceValues, baseColumn) Then
            stepName = "日付取得"
            If Not TryParseBlockDate(sourceSheet, DATE_ROW_INDEX, baseColumn, sourceValues(DATE_ROW_INDEX, baseColumn), baseYear, baseMonth, targetDate) Then
                messageText = "日付取得不可です。 " & DescribeCellForLog(sourceSheet, DATE_ROW_INDEX, baseColumn)
                AddDetailLogRow detailLogs, startedAt, fileName, sourceSheet.Name, vbNullString, vbNullString, vbNullString, "エラー", messageText
                GoTo FatalExit
            End If

            currentYearMonth = Format$(targetDate, "yyyymm")
            stepName = "年月検証"
            If Not ValidateFileYearMonth(fileName, currentYearMonth, fileYearMonth, importKeys, detailLogs, startedAt, sourceSheet.Name, messageText) Then
                GoTo FatalExit
            End If

            For rowIndex = DATA_START_ROW To DATA_END_ROW
                stepName = "施術時間取得"
                treatmentTime = sourceValues(rowIndex, baseColumn)

                If IsError(treatmentTime) Then
                    messageText = "施術時間セルにエラー値があります。"
                    AddDetailLogRow detailLogs, startedAt, fileName, sourceSheet.Name, targetDate, vbNullString, vbNullString, "エラー", messageText
                    GoTo FatalExit
                End If

                If Not IsBlankCellValue(treatmentTime) Then
                    rowsRead = rowsRead + 1
                    stepName = "基本項目取得"
                    category = NormalizeText(sourceValues(rowIndex, baseColumn + 1))
                    staffName = NormalizeText(sourceValues(rowIndex, baseColumn + 2))
                    patientName = NormalizeText(sourceValues(rowIndex, baseColumn + 3))

                    If Len(staffName) = 0 Then
                        warningCount = warningCount + 1
                        AddDetailLogRow detailLogs, startedAt, fileName, sourceSheet.Name, targetDate, patientName, staffName, "警告", "担当が空欄です"
                    End If

                    If Len(patientName) = 0 Then
                        warningCount = warningCount + 1
                        AddDetailLogRow detailLogs, startedAt, fileName, sourceSheet.Name, targetDate, patientName, staffName, "警告", "患者名が空欄です"
                    End If

                    totalAmount = 0
                    For feeColumn = baseColumn + 4 To baseColumn + 9
                        stepName = "料金集計"
                        If Not AccumulateAmount(sourceValues(rowIndex, feeColumn), totalAmount) Then
                            warningCount = warningCount + 1
                            AddDetailLogRow detailLogs, startedAt, fileName, sourceSheet.Name, targetDate, patientName, staffName, "警告", "料金に文字が含まれています"
                        End If
                    Next feeColumn

                    stepName = "顧客ID判定"
                    importSequence = importSequence + 1
                    internalImportId = BuildInternalImportId(startedAt, importSequence)
                    customerResolution = ResolveCustomerId(patientMap, patientName, internalImportId)
                    If Len(customerResolution.WarningMessage) > 0 Then
                        warningCount = warningCount + 1
                        AddDetailLogRow detailLogs, startedAt, fileName, sourceSheet.Name, targetDate, patientName, staffName, "警告", customerResolution.WarningMessage
                    End If

                    stepName = "マスタ行追加"
                    AddMasterRow masterRows, _
                                 customerResolution.CustomerId, _
                                 customerResolution.StatusText, _
                                 internalImportId, _
                                 targetDate, _
                                 currentYearMonth, _
                                 NormalizeOutputValue(treatmentTime), _
                                 category, _
                                 staffName, _
                                 patientName, _
                                 totalAmount, _
                                 storeId, _
                                 storeName, _
                                 fileName, _
                                 filePath, _
                                 sourceSheet.Name, _
                                 startedAt

                    rowsOutput = rowsOutput + 1
                End If
            Next rowIndex
        End If
    Next blockStart

    ProcessSheet = True
    Exit Function

FatalExit:
    ProcessSheet = False
    Exit Function

ErrHandler:
    messageText = "致命的な構造不整合: " & Err.Description
    If Len(stepName) > 0 Then
        messageText = messageText & " / 処理=" & stepName
    End If
    If baseColumn > 0 Then
        messageText = messageText & " / ブロック列=" & ColumnNumberToLetter(baseColumn)
    End If
    If rowIndex > 0 Then
        messageText = messageText & " / 行=" & rowIndex
    End If
    AddDetailLogRow detailLogs, startedAt, fileName, sourceSheet.Name, vbNullString, vbNullString, vbNullString, "エラー", messageText
    ProcessSheet = False
End Function

Private Function ValidateFileYearMonth( _
    ByVal fileName As String, _
    ByVal currentYearMonth As String, _
    ByRef fileYearMonth As String, _
    ByVal importKeys As Object, _
    ByVal detailLogs As Collection, _
    ByVal startedAt As Date, _
    ByVal sheetName As String, _
    ByRef messageText As String) As Boolean

    Dim duplicateKey As String

    If Len(fileYearMonth) = 0 Then
        duplicateKey = BuildImportKey(fileName, currentYearMonth)
        If importKeys.Exists(duplicateKey) Then
            messageText = "エラー：過去にこのファイルを出力しています。（" & fileName & "）"
            AddDetailLogRow detailLogs, startedAt, fileName, sheetName, vbNullString, vbNullString, vbNullString, "エラー", messageText
            ValidateFileYearMonth = False
            Exit Function
        End If

        fileYearMonth = currentYearMonth
        ValidateFileYearMonth = True
        Exit Function
    End If

    If fileYearMonth <> currentYearMonth Then
        messageText = "同一ファイル内に複数年月が存在します。"
        AddDetailLogRow detailLogs, startedAt, fileName, sheetName, vbNullString, vbNullString, vbNullString, "エラー", messageText
        ValidateFileYearMonth = False
        Exit Function
    End If

    ValidateFileYearMonth = True
End Function

Private Function ValidateTransferredCount( _
    ByVal sourceBook As Workbook, _
    ByVal outputCount As Long, _
    ByRef expectedCount As Long, _
    ByRef messageText As String) As Boolean

    Dim transferSheet As Worksheet
    Dim rawValue As Variant

    If Not WorksheetExists(sourceBook, TRANSFER_SHEET_NAME) Then
        messageText = "転記シートが存在しません。"
        Exit Function
    End If

    Set transferSheet = sourceBook.Worksheets(TRANSFER_SHEET_NAME)
    rawValue = transferSheet.Range("B35").Value

    If IsError(rawValue) Or IsBlankCellValue(rawValue) Then
        messageText = "転記シートB35の件数を取得できません。"
        Exit Function
    End If

    If Not IsNumeric(rawValue) Then
        messageText = "転記シートB35が数値ではありません。 値=""" & CStr(rawValue) & """"
        Exit Function
    End If

    expectedCount = CLng(rawValue)
    If outputCount <> expectedCount Then
        messageText = "転記シートB35の件数と出力件数が一致しません。 B35=" & expectedCount & " / 出力件数=" & outputCount
        Exit Function
    End If

    ValidateTransferredCount = True
End Function

Private Function LoadStoreMap(ByVal referenceSheet As Worksheet) As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim storeMap As Object
    Dim storeId As String
    Dim storeName As String
    Dim activeFlag As Variant

    Set storeMap = CreateObject("Scripting.Dictionary")
    lastRow = LastDataRow(referenceSheet, 1)

    For rowIndex = 2 To lastRow
        storeId = NormalizeStoreIdValue(referenceSheet.Cells(rowIndex, 1).Value)
        storeName = NormalizeText(referenceSheet.Cells(rowIndex, 2).Value)
        activeFlag = referenceSheet.Cells(rowIndex, 3).Value

        If Len(storeId) > 0 And Len(storeName) > 0 Then
            If IsReferenceActive(activeFlag) Then
                storeMap(storeId) = storeName
            End If
        End If
    Next rowIndex

    Set LoadStoreMap = storeMap
End Function

Private Sub PopulateReferenceSheet(ByVal referenceSheet As Worksheet, ByVal storeId As String, ByVal storeName As String)
    referenceSheet.Cells(2, 1).NumberFormatLocal = "@"
    referenceSheet.Cells(2, 1).Value2 = storeId
    referenceSheet.Cells(2, 2).Value = storeName
    referenceSheet.Cells(2, 3).Value = vbNullString
End Sub

Private Function LoadExistingImportKeys(ByVal masterSheet As Worksheet) As Object
    Dim importKeys As Object
    Dim lastRow As Long
    Dim sourceValues As Variant
    Dim rowIndex As Long
    Dim fileName As String
    Dim yearMonth As String

    Set importKeys = CreateObject("Scripting.Dictionary")
    lastRow = LastDataRow(masterSheet, 1)

    If lastRow < 2 Then
        Set LoadExistingImportKeys = importKeys
        Exit Function
    End If

    sourceValues = masterSheet.Range(masterSheet.Cells(2, 1), masterSheet.Cells(lastRow, UBound(MasterHeaders()) + 1)).Value2

    For rowIndex = 1 To UBound(sourceValues, 1)
        fileName = NormalizeText(sourceValues(rowIndex, 13))
        yearMonth = NormalizeText(sourceValues(rowIndex, 5))

        If Len(fileName) > 0 And Len(yearMonth) > 0 Then
            importKeys(BuildImportKey(fileName, yearMonth)) = True
        End If
    Next rowIndex

    Set LoadExistingImportKeys = importKeys
End Function

Private Sub AddMasterRow( _
    ByVal rows As Collection, _
    ByVal customerId As String, _
    ByVal customerStatus As String, _
    ByVal internalImportId As String, _
    ByVal targetDate As Date, _
    ByVal yearMonth As String, _
    ByVal treatmentTime As Variant, _
    ByVal category As String, _
    ByVal staffName As String, _
    ByVal patientName As String, _
    ByVal totalAmount As Double, _
    ByVal storeId As String, _
    ByVal storeName As String, _
    ByVal fileName As String, _
    ByVal filePath As String, _
    ByVal sheetName As String, _
    ByVal importedAt As Date)

    rows.Add Array( _
        customerId, _
        customerStatus, _
        internalImportId, _
        targetDate, _
        yearMonth, _
        treatmentTime, _
        category, _
        staffName, _
        patientName, _
        totalAmount, _
        storeId, _
        storeName, _
        fileName, _
        filePath, _
        sheetName, _
        importedAt)
End Sub

Private Sub AddExecutionLogRow( _
    ByVal rows As Collection, _
    ByVal executedAt As Date, _
    ByVal fileName As String, _
    ByVal sheetName As String, _
    ByVal statusText As String, _
    ByVal readCount As Long, _
    ByVal outputCount As Long, _
    ByVal messageText As String)

    rows.Add Array( _
        executedAt, _
        fileName, _
        sheetName, _
        statusText, _
        readCount, _
        outputCount, _
        messageText)
End Sub

Private Sub AddDetailLogRow( _
    ByVal rows As Collection, _
    ByVal executedAt As Date, _
    ByVal fileName As String, _
    ByVal sheetName As String, _
    ByVal targetDate As Variant, _
    ByVal patientName As String, _
    ByVal staffName As String, _
    ByVal levelText As String, _
    ByVal messageText As String)

    rows.Add Array( _
        executedAt, _
        fileName, _
        sheetName, _
        targetDate, _
        patientName, _
        staffName, _
        levelText, _
        messageText)
End Sub

Private Sub AppendRows(ByVal targetSheet As Worksheet, ByVal rows As Collection, ByVal columnCount As Long)
    Dim nextRow As Long
    Dim outputValues() As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim rowData As Variant

    If rows Is Nothing Then
        Exit Sub
    End If

    If rows.Count = 0 Then
        Exit Sub
    End If

    nextRow = LastDataRow(targetSheet, 1) + 1
    If nextRow < 2 Then
        nextRow = 2
    End If

    ReDim outputValues(1 To rows.Count, 1 To columnCount)

    For rowIndex = 1 To rows.Count
        rowData = rows(rowIndex)
        For columnIndex = 0 To UBound(rowData)
            outputValues(rowIndex, columnIndex + 1) = rowData(columnIndex)
        Next columnIndex
    Next rowIndex

    targetSheet.Cells(nextRow, 1).Resize(rows.Count, columnCount).Value = outputValues
End Sub

Private Sub EnsureSheetWithHeaders(ByVal sheetName As String, ByVal headers As Variant)
    Dim targetSheet As Worksheet
    Dim columnIndex As Long
    Dim existingValue As String
    Dim canResetHeaders As Boolean
    Dim hasMismatch As Boolean

    Set targetSheet = GetOrCreateWorksheet(sheetName)
    canResetHeaders = Not SheetHasDataRows(targetSheet)

    If Application.WorksheetFunction.CountA(targetSheet.Rows(1)) = 0 Then
        For columnIndex = LBound(headers) To UBound(headers)
            targetSheet.Cells(1, columnIndex + 1).Value = headers(columnIndex)
        Next columnIndex
        targetSheet.Rows(1).Font.Bold = True
        Exit Sub
    End If

    For columnIndex = LBound(headers) To UBound(headers)
        existingValue = NormalizeText(targetSheet.Cells(1, columnIndex + 1).Value)
        If existingValue <> CStr(headers(columnIndex)) Then
            hasMismatch = True
            Exit For
        End If
    Next columnIndex

    If hasMismatch Then
        If canResetHeaders Then
            targetSheet.Rows(1).ClearContents
            For columnIndex = LBound(headers) To UBound(headers)
                targetSheet.Cells(1, columnIndex + 1).Value = headers(columnIndex)
            Next columnIndex
            targetSheet.Rows(1).Font.Bold = True
            Exit Sub
        End If

        Err.Raise vbObjectError + 1100, , """" & sheetName & """ シートのヘッダーが想定と異なります。"
    End If
End Sub

Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    Dim worksheetItem As Worksheet

    For Each worksheetItem In ThisWorkbook.Worksheets
        If worksheetItem.Name = sheetName Then
            Set GetOrCreateWorksheet = worksheetItem
            Exit Function
        End If
    Next worksheetItem

    Set worksheetItem = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    worksheetItem.Name = sheetName
    Set GetOrCreateWorksheet = worksheetItem
End Function

Private Function WorksheetExists(ByVal targetBook As Workbook, ByVal sheetName As String) As Boolean
    Dim worksheetItem As Worksheet

    For Each worksheetItem In targetBook.Worksheets
        If worksheetItem.Name = sheetName Then
            WorksheetExists = True
            Exit Function
        End If
    Next worksheetItem
End Function

Private Sub CollectXlsmFilesRecursive(ByVal folderPath As String, ByVal filePaths As Collection)
    Dim fso As Object
    Dim folderObject As Object
    Dim subFolder As Object
    Dim fileObject As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folderObject = fso.GetFolder(folderPath)

    For Each fileObject In folderObject.Files
        If LCase$(fso.GetExtensionName(fileObject.Name)) = "xlsm" Then
            If Not IsTemporaryExcelFile(CStr(fileObject.Path)) Then
                filePaths.Add CStr(fileObject.Path)
            End If
        End If
    Next fileObject

    For Each subFolder In folderObject.SubFolders
        CollectXlsmFilesRecursive CStr(subFolder.Path), filePaths
    Next subFolder
End Sub

Private Function PickFolder() As String
    Dim dialog As FileDialog

    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    With dialog
        .Title = "治療日誌フォルダを選択してください"
        If .Show <> -1 Then
            Exit Function
        End If

        PickFolder = .SelectedItems(1)
    End With
End Function

Private Function PickPatientAnalysisWorkbook() As String
    Dim dialog As FileDialog
    Dim selectedPath As String

    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    With dialog
        .Title = "新患分析 / id店舗名.xlsx を選択してください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Workbook", "*.xlsx;*.xlsm;*.xls"

        If .Show <> -1 Then
            Exit Function
        End If

        selectedPath = NormalizeWorkbookPath(CStr(.SelectedItems(1)))
        PickPatientAnalysisWorkbook = selectedPath
    End With
End Function

Private Function LoadPatientMap(ByVal workbookPath As String) As Object
    Dim patientMap As Object
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim normalizedPath As String
    Dim openedFilePath As String
    Dim tempOpenPath As String
    Dim openErrorMessage As String
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim patientName As String
    Dim customerId As String
    Dim physicalFileName As String
    Dim logicalFileName As String
    Dim raisedNumber As Long
    Dim raisedDescription As String

    On Error GoTo ErrHandler

    Set patientMap = CreateObject("Scripting.Dictionary")
    normalizedPath = NormalizeWorkbookPath(workbookPath)
    If Not FileExists(normalizedPath) Then
        Err.Raise vbObjectError + 1007, , "新患分析ファイルが見つかりません。 path=" & normalizedPath
    End If

    physicalFileName = GetFileName(normalizedPath)
    logicalFileName = GetLogicalFileName(normalizedPath)
    If physicalFileName <> logicalFileName Then
        tempOpenPath = CreateTempWorkbookCopy(normalizedPath, logicalFileName)
        openedFilePath = tempOpenPath
    Else
        openedFilePath = normalizedPath
    End If

    Set sourceBook = TryOpenWorkbookReadOnly(openedFilePath, openErrorMessage)
    If sourceBook Is Nothing And openedFilePath = normalizedPath Then
        tempOpenPath = CreateTempWorkbookCopy(normalizedPath, logicalFileName)
        openedFilePath = tempOpenPath
        Set sourceBook = TryOpenWorkbookReadOnly(openedFilePath, openErrorMessage)
    End If
    If sourceBook Is Nothing Then
        Err.Raise vbObjectError + 1008, , "新患分析ファイルを開けません。 path=" & openedFilePath & " / detail=" & openErrorMessage
    End If

    If Not WorksheetExists(sourceBook, NEW_PATIENT_SHEET_NAME) Then
        Err.Raise vbObjectError + 1009, , "新患分析ファイルに """ & NEW_PATIENT_SHEET_NAME & """ シートが存在しません。"
    End If

    Set sourceSheet = sourceBook.Worksheets(NEW_PATIENT_SHEET_NAME)
    lastRow = LastDataRow(sourceSheet, NEW_PATIENT_NAME_COLUMN)
    For rowIndex = 2 To lastRow
        patientName = NormalizeText(sourceSheet.Cells(rowIndex, NEW_PATIENT_NAME_COLUMN).Value)
        customerId = NormalizeCustomerIdValue(sourceSheet.Cells(rowIndex, NEW_PATIENT_ID_COLUMN).Value)
        AddPatientMapEntry patientMap, patientName, customerId
    Next rowIndex

SafeExit:
    If Not sourceBook Is Nothing Then
        sourceBook.Close SaveChanges:=False
    End If
    DeleteFileIfExists tempOpenPath
    Set LoadPatientMap = patientMap
    Exit Function

ErrHandler:
    raisedNumber = Err.Number
    raisedDescription = Err.Description
    On Error Resume Next
    If Not sourceBook Is Nothing Then
        sourceBook.Close SaveChanges:=False
    End If
    DeleteFileIfExists tempOpenPath
    Err.Raise raisedNumber, , raisedDescription
End Function

Private Sub AddPatientMapEntry(ByVal patientMap As Object, ByVal patientName As String, ByVal customerId As String)
    Dim patientKey As String
    Dim patientInfo As Object
    Dim candidateIds As Object
    Dim numericCustomerId As Double

    patientKey = NormalizePatientNameKey(patientName)
    If Len(patientKey) = 0 Then
        Exit Sub
    End If

    If patientMap.Exists(patientKey) Then
        Set patientInfo = patientMap(patientKey)
    Else
        Set patientInfo = CreateObject("Scripting.Dictionary")
        patientInfo("row_count") = 0
        patientInfo("latest_id") = vbNullString
        patientInfo("latest_id_value") = -1#
        Set candidateIds = CreateObject("Scripting.Dictionary")
        patientInfo.Add "candidate_ids", candidateIds
        patientMap.Add patientKey, patientInfo
    End If

    patientInfo("row_count") = CLng(patientInfo("row_count")) + 1

    If Len(customerId) = 0 Then
        Exit Sub
    End If

    Set candidateIds = patientInfo("candidate_ids")
    If Not candidateIds.Exists(customerId) Then
        candidateIds.Add customerId, CDbl(customerId)
    End If

    numericCustomerId = CDbl(customerId)
    If numericCustomerId >= CDbl(patientInfo("latest_id_value")) Then
        patientInfo("latest_id_value") = numericCustomerId
        patientInfo("latest_id") = customerId
    End If
End Sub

Private Function ResolveCustomerId(ByVal patientMap As Object, ByVal patientName As String, ByVal internalImportId As String) As CustomerResolution
    Dim patientKey As String
    Dim patientInfo As Object
    Dim candidateIds As Object
    Dim latestId As String

    patientKey = NormalizePatientNameKey(patientName)

    If Len(patientKey) = 0 Then
        ResolveCustomerId.CustomerId = BuildTemporaryCustomerId(internalImportId)
        ResolveCustomerId.StatusText = "仮"
        ResolveCustomerId.WarningMessage = "顧客ID（カルテNo）が取得できません。治療日誌の患者名が空欄です。"
        Exit Function
    End If

    If Not patientMap.Exists(patientKey) Then
        ResolveCustomerId.CustomerId = BuildTemporaryCustomerId(internalImportId)
        ResolveCustomerId.StatusText = "仮"
        ResolveCustomerId.WarningMessage = "顧客ID（カルテNo）が取得できません。治療日誌の名前と新患分析の名前が一致していません。新患分析の名前を確認してください。治療日誌の名前: " & patientName
        Exit Function
    End If

    Set patientInfo = patientMap(patientKey)
    Set candidateIds = patientInfo("candidate_ids")
    If candidateIds.Count = 0 Then
        ResolveCustomerId.CustomerId = BuildTemporaryCustomerId(internalImportId)
        ResolveCustomerId.StatusText = "仮"
        ResolveCustomerId.WarningMessage = "顧客ID（カルテNo）が未登録です。新患分析の顧客Noを確認してください。治療日誌の名前: " & patientName
        Exit Function
    End If

    latestId = CStr(patientInfo("latest_id"))
    ResolveCustomerId.CustomerId = latestId

    If CLng(patientInfo("row_count")) > 1 Then
        ResolveCustomerId.StatusText = "警告"
        If candidateIds.Count > 1 Then
            ResolveCustomerId.WarningMessage = "同じ名前の方がいます。候補: " & JoinCandidateIds(candidateIds) & "。最新の顧客ID No" & latestId & " を採用しました。"
        Else
            ResolveCustomerId.WarningMessage = "同じ名前の方がいます。顧客ID No" & latestId & " を採用しました。"
        End If
    Else
        ResolveCustomerId.StatusText = "確定"
    End If
End Function

Private Function BlockHasInputRows(ByVal sourceValues As Variant, ByVal baseColumn As Long) As Boolean
    Dim rowIndex As Long

    For rowIndex = DATA_START_ROW To DATA_END_ROW
        If Not IsBlankCellValue(sourceValues(rowIndex, baseColumn)) Then
            BlockHasInputRows = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function TryParseExcelDate(ByVal value As Variant, ByRef parsedDate As Date) As Boolean
    If IsError(value) Then
        Exit Function
    End If

    If IsBlankCellValue(value) Then
        Exit Function
    End If

    If IsDate(value) Then
        parsedDate = DateValue(CDate(value))
        TryParseExcelDate = True
    End If
End Function

Private Function TryParseBlockDate( _
    ByVal sourceSheet As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long, _
    ByVal rawValue As Variant, _
    ByVal baseYear As Long, _
    ByVal baseMonth As Long, _
    ByRef parsedDate As Date) As Boolean

    Dim displayText As String

    displayText = NormalizeText(sourceSheet.Cells(rowIndex, columnIndex).Text)

    If TryBuildDateFromFullText(displayText, parsedDate) Then
        TryParseBlockDate = True
        Exit Function
    End If

    If baseYear > 0 And baseMonth > 0 Then
        If TryBuildDateFromMonthDayText(displayText, baseYear, parsedDate) Then
            TryParseBlockDate = True
            Exit Function
        End If

        If TryBuildDateFromDayValue(rawValue, displayText, baseYear, baseMonth, parsedDate) Then
            TryParseBlockDate = True
            Exit Function
        End If
    End If

    If TryParseExcelDate(rawValue, parsedDate) Then
        TryParseBlockDate = True
        Exit Function
    End If

    If TryParseExcelDate(displayText, parsedDate) Then
        TryParseBlockDate = True
    End If
End Function

Private Function TryBuildDateFromFullText(ByVal textValue As String, ByRef parsedDate As Date) As Boolean
    Dim regex As Object
    Dim matches As Object
    Dim yearNumber As Long
    Dim monthNumber As Long
    Dim dayNumber As Long

    If Len(textValue) = 0 Then
        Exit Function
    End If

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^\s*(\d{4})\D+(\d{1,2})\D+(\d{1,2})"
    regex.Global = False

    If regex.Test(textValue) Then
        Set matches = regex.Execute(textValue)
        yearNumber = CLng(matches(0).SubMatches(0))
        monthNumber = CLng(matches(0).SubMatches(1))
        dayNumber = CLng(matches(0).SubMatches(2))

        If monthNumber >= 1 And monthNumber <= 12 And dayNumber >= 1 And dayNumber <= 31 Then
            parsedDate = DateSerial(yearNumber, monthNumber, dayNumber)
            TryBuildDateFromFullText = True
        End If
    End If
End Function

Private Function TryBuildDateFromMonthDayText(ByVal textValue As String, ByVal baseYear As Long, ByRef parsedDate As Date) As Boolean
    Dim regex As Object
    Dim matches As Object
    Dim monthNumber As Long
    Dim dayNumber As Long

    If Len(textValue) = 0 Then
        Exit Function
    End If

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^\s*(\d{1,2})\D+(\d{1,2})"
    regex.Global = False

    If regex.Test(textValue) Then
        Set matches = regex.Execute(textValue)
        monthNumber = CLng(matches(0).SubMatches(0))
        dayNumber = CLng(matches(0).SubMatches(1))

        If monthNumber >= 1 And monthNumber <= 12 And dayNumber >= 1 And dayNumber <= 31 Then
            parsedDate = DateSerial(baseYear, monthNumber, dayNumber)
            TryBuildDateFromMonthDayText = True
        End If
    End If
End Function

Private Function TryBuildDateFromDayValue( _
    ByVal rawValue As Variant, _
    ByVal textValue As String, _
    ByVal baseYear As Long, _
    ByVal baseMonth As Long, _
    ByRef parsedDate As Date) As Boolean

    Dim dayNumber As Long

    If TryExtractDayNumber(rawValue, textValue, dayNumber) Then
        parsedDate = DateSerial(baseYear, baseMonth, dayNumber)
        TryBuildDateFromDayValue = True
    End If
End Function

Private Function TryExtractDayNumber(ByVal rawValue As Variant, ByVal textValue As String, ByRef dayNumber As Long) As Boolean
    Dim regex As Object
    Dim matches As Object

    If Not IsError(rawValue) Then
        If IsNumeric(rawValue) Then
            dayNumber = CLng(rawValue)
            If dayNumber >= 1 And dayNumber <= 31 Then
                TryExtractDayNumber = True
                Exit Function
            End If
        End If
    End If

    If Len(textValue) = 0 Then
        Exit Function
    End If

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^\s*(\d{1,2})(?:\D.*)?$"
    regex.Global = False

    If regex.Test(textValue) Then
        Set matches = regex.Execute(textValue)
        dayNumber = CLng(matches(0).SubMatches(0))
        If dayNumber >= 1 And dayNumber <= 31 Then
            TryExtractDayNumber = True
        End If
    End If
End Function

Private Function AccumulateAmount(ByVal value As Variant, ByRef totalAmount As Double) As Boolean
    If IsError(value) Then
        AccumulateAmount = False
        Exit Function
    End If

    If IsBlankCellValue(value) Then
        AccumulateAmount = True
        Exit Function
    End If

    Select Case VarType(value)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            totalAmount = totalAmount + CDbl(value)
            AccumulateAmount = True
        Case Else
            AccumulateAmount = False
    End Select
End Function

Private Function NormalizeText(ByVal value As Variant) As String
    Dim textValue As String

    If IsError(value) Then
        Exit Function
    End If

    If IsBlankCellValue(value) Then
        Exit Function
    End If

    textValue = CStr(value)
    textValue = Replace(textValue, ChrW$(12288), " ")
    NormalizeText = Trim$(textValue)
End Function

Private Function NormalizePatientNameKey(ByVal value As Variant) As String
    Dim normalized As String

    normalized = NormalizeText(value)
    normalized = Replace(normalized, " ", vbNullString)
    normalized = Replace(normalized, ChrW$(12288), vbNullString)
    NormalizePatientNameKey = normalized
End Function

Private Function NormalizeCustomerIdValue(ByVal value As Variant) As String
    Dim normalized As String

    If IsError(value) Or IsBlankCellValue(value) Then
        Exit Function
    End If

    normalized = NormalizeText(value)
    If Len(normalized) = 0 Then
        Exit Function
    End If

    If IsNumeric(normalized) Then
        If CDbl(normalized) = Fix(CDbl(normalized)) Then
            NormalizeCustomerIdValue = CStr(Fix(CDbl(normalized)))
        End If
    End If
End Function

Private Function NormalizeOutputValue(ByVal value As Variant) As Variant
    If IsError(value) Or IsBlankCellValue(value) Then
        NormalizeOutputValue = vbNullString
    Else
        NormalizeOutputValue = value
    End If
End Function

Private Function IsBlankCellValue(ByVal value As Variant) As Boolean
    If IsError(value) Then
        Exit Function
    End If

    If IsEmpty(value) Or IsNull(value) Then
        IsBlankCellValue = True
        Exit Function
    End If

    If VarType(value) = vbString Then
        IsBlankCellValue = Len(Trim$(Replace(CStr(value), ChrW$(12288), " "))) = 0
        Exit Function
    End If

    IsBlankCellValue = False
End Function

Private Function ExtractStoreId(ByVal fileName As String) As String
    Dim baseName As String
    Dim characterIndex As Long
    Dim currentChar As String

    baseName = GetLogicalFileName(fileName)
    If LCase$(Right$(baseName, 5)) = ".xlsm" Then
        baseName = Left$(baseName, Len(baseName) - 5)
    End If
    baseName = NormalizeText(baseName)

    For characterIndex = 1 To Len(baseName)
        currentChar = Mid$(baseName, characterIndex, 1)
        If currentChar Like "#" Then
            ExtractStoreId = ExtractStoreId & currentChar
        Else
            Exit For
        End If
    Next characterIndex
End Function

Private Function NormalizeStoreIdValue(ByVal value As Variant) As String
    If IsError(value) Or IsBlankCellValue(value) Then
        Exit Function
    End If

    NormalizeStoreIdValue = NormalizeStoreIdText(CStr(value))
End Function

Private Function NormalizeStoreIdText(ByVal value As String) As String
    Dim normalized As String

    normalized = NormalizeText(value)
    If Len(normalized) = 0 Then
        Exit Function
    End If

    If normalized Like String$(Len(normalized), "#") Then
        NormalizeStoreIdText = Format$(CLng(normalized), "00")
    Else
        NormalizeStoreIdText = normalized
    End If
End Function

Private Function TryParseStoreInfoFromWorkbookName(ByVal workbookName As String, ByRef storeId As String, ByRef storeName As String) As Boolean
    Dim baseName As String
    Dim regex As Object
    Dim matches As Object

    baseName = workbookName
    If InStrRev(baseName, ".") > 0 Then
        baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    End If

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^.*_([^_]+)_([^_]+)$"
    regex.Global = False

    If Not regex.Test(baseName) Then
        Exit Function
    End If

    Set matches = regex.Execute(baseName)
    storeId = NormalizeStoreIdText(matches(0).SubMatches(0))
    storeName = NormalizeText(matches(0).SubMatches(1))

    If Len(storeId) = 0 Or Len(storeName) = 0 Then
        Exit Function
    End If

    TryParseStoreInfoFromWorkbookName = True
End Function

Private Function IsReferenceActive(ByVal value As Variant) As Boolean
    Dim normalized As String

    If IsBlankCellValue(value) Then
        IsReferenceActive = True
        Exit Function
    End If

    normalized = UCase$(NormalizeText(value))

    Select Case normalized
        Case "0", "FALSE", "NO", "N", "無効", "停止"
            IsReferenceActive = False
        Case Else
            IsReferenceActive = True
    End Select
End Function

Private Function BuildImportKey(ByVal fileName As String, ByVal yearMonth As String) As String
    BuildImportKey = UCase$(fileName) & "|" & yearMonth
End Function

Private Function BuildInternalImportId(ByVal startedAt As Date, ByVal importSequence As Long) As String
    BuildInternalImportId = Format$(startedAt, "yyyymmddhhnnss") & Format$(importSequence, "0000")
End Function

Private Function BuildTemporaryCustomerId(ByVal internalImportId As String) As String
    BuildTemporaryCustomerId = "-" & internalImportId
End Function

Private Function JoinCandidateIds(ByVal candidateIds As Object) As String
    Dim keys As Variant
    Dim indexA As Long
    Dim indexB As Long
    Dim swapValue As Variant

    keys = candidateIds.Keys
    For indexA = LBound(keys) To UBound(keys) - 1
        For indexB = indexA + 1 To UBound(keys)
            If CDbl(keys(indexA)) > CDbl(keys(indexB)) Then
                swapValue = keys(indexA)
                keys(indexA) = keys(indexB)
                keys(indexB) = swapValue
            End If
        Next indexB
    Next indexA

    For indexA = LBound(keys) To UBound(keys)
        If indexA > LBound(keys) Then
            JoinCandidateIds = JoinCandidateIds & ", "
        End If
        JoinCandidateIds = JoinCandidateIds & "No" & CStr(keys(indexA))
    Next indexA
End Function

Private Sub ExtractYearMonthFromPath(ByVal filePath As String, ByRef baseYear As Long, ByRef baseMonth As Long)
    Dim regex As Object
    Dim matches As Object
    Dim lastMatch As Object

    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(\d{4})年(\d{1,2})月"
    regex.Global = True

    If regex.Test(filePath) Then
        Set matches = regex.Execute(filePath)
        Set lastMatch = matches(matches.Count - 1)
        baseYear = CLng(lastMatch.SubMatches(0))
        baseMonth = CLng(lastMatch.SubMatches(1))
    End If
End Sub

Private Function NormalizeWorkbookPath(ByVal filePath As String) As String
    Dim normalizedPath As String

    normalizedPath = Trim$(filePath)
    normalizedPath = Replace(normalizedPath, "/", Application.PathSeparator)

    Do While InStr(normalizedPath, Application.PathSeparator & Application.PathSeparator) > 0
        normalizedPath = Replace(normalizedPath, Application.PathSeparator & Application.PathSeparator, Application.PathSeparator)
    Loop

    NormalizeWorkbookPath = normalizedPath
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    If Len(filePath) = 0 Then
        Exit Function
    End If

    FileExists = Len(Dir$(filePath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)) > 0
End Function

Private Function DescribeCellForLog(ByVal sourceSheet As Worksheet, ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    Dim displayText As String
    Dim rawValue As Variant
    Dim rawText As String

    rawValue = sourceSheet.Cells(rowIndex, columnIndex).Value
    displayText = sourceSheet.Cells(rowIndex, columnIndex).Text

    If IsError(rawValue) Then
        rawText = "#ERROR"
    ElseIf IsEmpty(rawValue) Or IsNull(rawValue) Then
        rawText = "(empty)"
    Else
        rawText = CStr(rawValue)
    End If

    DescribeCellForLog = "cell=" & ColumnNumberToLetter(columnIndex) & rowIndex & _
                         " / 表示=""" & displayText & """" & _
                         " / 値=""" & rawText & """"
End Function

Private Function TryOpenWorkbookReadOnly(ByVal filePath As String, ByRef errorMessage As String) As Workbook
    On Error GoTo OpenFailed

    errorMessage = vbNullString
    Set TryOpenWorkbookReadOnly = Workbooks.Open( _
        Filename:=filePath, _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True, _
        AddToMru:=False)
    Exit Function

OpenFailed:
    errorMessage = CStr(Err.Number) & ":" & Err.Description
    Set TryOpenWorkbookReadOnly = Nothing
End Function

Private Function CreateTempWorkbookCopy(ByVal sourcePath As String, ByVal logicalFileName As String) As String
    Dim tempFolder As String
    Dim tempFileName As String
    Dim destinationPath As String

    tempFolder = Environ$("TEMP")
    If Len(tempFolder) = 0 Then
        Err.Raise vbObjectError + 1005, , "一時フォルダを取得できません。"
    End If

    tempFileName = Format$(Now, "yyyymmddhhnnss") & "_" & SanitizeFileName(logicalFileName)
    destinationPath = tempFolder & Application.PathSeparator & tempFileName

    FileCopy sourcePath, destinationPath
    CreateTempWorkbookCopy = destinationPath
End Function

Private Sub DeleteFileIfExists(ByVal filePath As String)
    On Error GoTo DeleteFailed

    If Len(filePath) = 0 Then
        Exit Sub
    End If

    If FileExists(filePath) Then
        Kill filePath
    End If
    Exit Sub

DeleteFailed:
End Sub

Private Function SanitizeFileName(ByVal fileName As String) As String
    Dim invalidChars As Variant
    Dim characterItem As Variant
    Dim sanitized As String

    sanitized = Trim$(fileName)
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For Each characterItem In invalidChars
        sanitized = Replace(sanitized, CStr(characterItem), "_")
    Next characterItem

    If Len(sanitized) = 0 Then
        sanitized = "source.xlsm"
    End If

    SanitizeFileName = sanitized
End Function

Private Function ColumnNumberToLetter(ByVal columnNumber As Long) As String
    Dim dividend As Long
    Dim modulo As Long

    dividend = columnNumber

    Do While dividend > 0
        modulo = (dividend - 1) Mod 26
        ColumnNumberToLetter = Chr$(65 + modulo) & ColumnNumberToLetter
        dividend = (dividend - modulo - 1) \ 26
    Loop
End Function

Private Function GetFileName(ByVal filePath As String) As String
    Dim normalizedPath As String
    Dim separatorIndex As Long

    normalizedPath = NormalizeWorkbookPath(filePath)

    separatorIndex = InStrRev(normalizedPath, Application.PathSeparator)
    If separatorIndex = 0 Then
        GetFileName = normalizedPath
    Else
        GetFileName = Mid$(normalizedPath, separatorIndex + 1)
    End If
End Function

Private Function GetLogicalFileName(ByVal filePath As String) As String
    Dim physicalFileName As String
    Dim separatorIndex As Long

    physicalFileName = GetFileName(filePath)

    separatorIndex = InStrRev(physicalFileName, "／")
    If separatorIndex = 0 Then
        separatorIndex = InStrRev(physicalFileName, "/")
    End If

    If separatorIndex > 0 Then
        GetLogicalFileName = Mid$(physicalFileName, separatorIndex + 1)
    Else
        GetLogicalFileName = physicalFileName
    End If
End Function

Private Function IsTemporaryExcelFile(ByVal filePath As String) As Boolean
    IsTemporaryExcelFile = Left$(GetFileName(filePath), 2) = "~$"
End Function

Private Function LastDataRow(ByVal targetSheet As Worksheet, ByVal keyColumn As Long) As Long
    Dim lastRow As Long

    lastRow = targetSheet.Cells(targetSheet.Rows.Count, keyColumn).End(xlUp).Row
    If lastRow < 1 Then
        lastRow = 1
    End If

    LastDataRow = lastRow
End Function

Private Function SheetHasDataRows(ByVal targetSheet As Worksheet) As Boolean
    Dim usedLastRow As Long
    Dim usedLastColumn As Long

    On Error GoTo SafeExit

    usedLastRow = targetSheet.Cells.Find(What:="*", After:=targetSheet.Cells(1, 1), LookIn:=xlFormulas, _
                                         LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
    usedLastColumn = targetSheet.Cells.Find(What:="*", After:=targetSheet.Cells(1, 1), LookIn:=xlFormulas, _
                                            LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
    SheetHasDataRows = (usedLastRow > 1) Or (usedLastRow = 1 And usedLastColumn > 0 And Application.WorksheetFunction.CountA(targetSheet.Rows(1)) = 0)
    Exit Function

SafeExit:
    SheetHasDataRows = False
End Function

Private Function ArchiveAndClearImportLogs() As String
    Dim execSheet As Worksheet
    Dim detailSheet As Worksheet

    Set execSheet = ThisWorkbook.Worksheets(EXEC_LOG_SHEET_NAME)
    Set detailSheet = ThisWorkbook.Worksheets(DETAIL_LOG_SHEET_NAME)

    ArchiveAndClearImportLogs = ArchiveImportLogs(execSheet, detailSheet, Now)
    ClearSheetDataRows execSheet
    ClearSheetDataRows detailSheet
End Function

Private Function ArchiveImportLogs(ByVal execSheet As Worksheet, ByVal detailSheet As Worksheet, ByVal archivedAt As Date) As String
    Dim hasExecLogs As Boolean
    Dim hasDetailLogs As Boolean
    Dim fso As Object
    Dim logStream As Object
    Dim outputPath As String
    Dim raisedNumber As Long
    Dim raisedDescription As String

    On Error GoTo ErrHandler

    hasExecLogs = SheetHasDataRows(execSheet)
    hasDetailLogs = SheetHasDataRows(detailSheet)
    If Not hasExecLogs And Not hasDetailLogs Then
        Exit Function
    End If

    outputPath = BuildImportLogFilePath(archivedAt)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logStream = fso.CreateTextFile(outputPath, True, True)

    logStream.WriteLine "治療日誌取込ログ"
    logStream.WriteLine "出力日時: " & Format$(archivedAt, "yyyy/mm/dd hh:nn:ss")
    logStream.WriteLine "ブック: " & ThisWorkbook.Name
    logStream.WriteLine String$(80, "=")

    WriteLogSection logStream, execSheet, EXEC_LOG_SHEET_NAME, UBound(ExecutionLogHeaders()) + 1
    logStream.WriteLine String$(80, "=")
    WriteLogSection logStream, detailSheet, DETAIL_LOG_SHEET_NAME, UBound(DetailLogHeaders()) + 1
    logStream.Close
    ArchiveImportLogs = outputPath
    Exit Function

ErrHandler:
    raisedNumber = Err.Number
    raisedDescription = Err.Description
    On Error Resume Next
    If Not logStream Is Nothing Then
        logStream.Close
    End If
    Err.Raise raisedNumber, , "ログファイルを作成できません。 path=" & outputPath & " / detail=" & raisedDescription
End Function

Private Sub AddExecutionLogInfo(ByVal targetSheet As Worksheet, ByVal executedAt As Date, ByVal messageText As String)
    Dim rows As Collection

    Set rows = New Collection
    AddExecutionLogRow rows, executedAt, vbNullString, vbNullString, "INFO", 0, 0, messageText
    AppendRows targetSheet, rows, UBound(ExecutionLogHeaders()) + 1
End Sub

Private Sub WriteLogSection(ByVal logStream As Object, ByVal targetSheet As Worksheet, ByVal sectionName As String, ByVal columnCount As Long)
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim lineText As String

    logStream.WriteLine "[" & sectionName & "]"

    lastRow = LastDataRow(targetSheet, 1)
    If lastRow < 2 Then
        logStream.WriteLine "(データなし)"
        logStream.WriteLine vbNullString
        Exit Sub
    End If

    For rowIndex = 1 To lastRow
        lineText = vbNullString
        For columnIndex = 1 To columnCount
            If columnIndex > 1 Then
                lineText = lineText & vbTab
            End If
            lineText = lineText & EscapeLogValue(targetSheet.Cells(rowIndex, columnIndex).Value)
        Next columnIndex
        logStream.WriteLine lineText
    Next rowIndex

    logStream.WriteLine vbNullString
End Sub

Private Sub ClearSheetDataRows(ByVal targetSheet As Worksheet)
    Dim lastRow As Long

    lastRow = LastDataRow(targetSheet, 1)
    If lastRow < 2 Then
        Exit Sub
    End If

    targetSheet.Rows("2:" & CStr(lastRow)).ClearContents
End Sub

Private Function BuildImportLogFilePath(ByVal archivedAt As Date) As String
    Dim baseFolder As String
    Dim logFolder As String

    baseFolder = ResolveLogBaseFolder()

    logFolder = baseFolder & Application.PathSeparator & "log"
    EnsureFolderExists logFolder
    BuildImportLogFilePath = logFolder & Application.PathSeparator & "import_" & Format$(archivedAt, "yyyymmdd_hhnnss") & ".log"
End Function

Private Function ResolveLogBaseFolder() As String
    Dim candidateFolder As String
    Dim localWorkbookPath As String

    candidateFolder = ThisWorkbook.Path
    If IsUsableLocalFolder(candidateFolder) Then
        ResolveLogBaseFolder = candidateFolder
        Exit Function
    End If

    localWorkbookPath = ResolveOneDriveLocalPath(ThisWorkbook.FullName)
    If Len(localWorkbookPath) > 0 Then
        candidateFolder = GetParentFolderPath(localWorkbookPath)
    Else
        candidateFolder = ResolveOneDriveLocalFolder(ThisWorkbook.Path)
    End If
    If IsUsableLocalFolder(candidateFolder) Then
        ResolveLogBaseFolder = candidateFolder
        Exit Function
    End If

    candidateFolder = Application.DefaultFilePath
    If IsUsableLocalFolder(candidateFolder) Then
        ResolveLogBaseFolder = candidateFolder
        Exit Function
    End If

    candidateFolder = Environ$("USERPROFILE")
    If IsUsableLocalFolder(candidateFolder) Then
        ResolveLogBaseFolder = candidateFolder
        Exit Function
    End If

    candidateFolder = Environ$("TEMP")
    If IsUsableLocalFolder(candidateFolder) Then
        ResolveLogBaseFolder = candidateFolder
        Exit Function
    End If

    Err.Raise vbObjectError + 1011, , _
              "ログ出力基準フォルダが見つかりません。 workbookPath=" & ThisWorkbook.Path & _
              " / workbookFullName=" & ThisWorkbook.FullName & _
              " / defaultPath=" & Application.DefaultFilePath
End Function

Private Function IsUsableLocalFolder(ByVal folderPath As String) As Boolean
    Dim normalizedPath As String

    normalizedPath = Trim$(folderPath)
    If Len(normalizedPath) = 0 Then
        Exit Function
    End If

    If LCase$(Left$(normalizedPath, 7)) = "http://" Or LCase$(Left$(normalizedPath, 8)) = "https://" Then
        Exit Function
    End If

    IsUsableLocalFolder = FolderExists(normalizedPath)
End Function

Private Function ResolveOneDriveLocalFolder(ByVal workbookPath As String) As String
    Dim normalizedPath As String
    Dim relativePath As String
    Dim oneDriveRoot As String
    Dim personalRoot As String

    normalizedPath = Trim$(workbookPath)
    If LCase$(Left$(normalizedPath, Len(ONEDRIVE_WEB_PREFIX))) <> ONEDRIVE_WEB_PREFIX Then
        Exit Function
    End If

    relativePath = Mid$(normalizedPath, Len(ONEDRIVE_WEB_PREFIX) + 1)
    If InStr(relativePath, "/") = 0 Then
        Exit Function
    End If

    relativePath = Mid$(relativePath, InStr(relativePath, "/") + 1)
    relativePath = Replace(relativePath, "/", Application.PathSeparator)

    oneDriveRoot = Environ$("OneDrive")
    If Len(oneDriveRoot) > 0 Then
        personalRoot = NormalizeWorkbookPath(oneDriveRoot & Application.PathSeparator & relativePath)
        If FolderExists(personalRoot) Then
            ResolveOneDriveLocalFolder = personalRoot
            Exit Function
        End If
    End If

    personalRoot = Environ$("UserProfile")
    If Len(personalRoot) > 0 Then
        personalRoot = NormalizeWorkbookPath(personalRoot & Application.PathSeparator & "OneDrive" & Application.PathSeparator & relativePath)
        If FolderExists(personalRoot) Then
            ResolveOneDriveLocalFolder = personalRoot
        End If
    End If
End Function

Private Function ResolveOneDriveLocalPath(ByVal workbookFullName As String) As String
    Dim normalizedPath As String
    Dim relativePath As String
    Dim oneDriveRoot As String
    Dim personalPath As String

    normalizedPath = Trim$(workbookFullName)
    If LCase$(Left$(normalizedPath, Len(ONEDRIVE_WEB_PREFIX))) <> ONEDRIVE_WEB_PREFIX Then
        Exit Function
    End If

    relativePath = Mid$(normalizedPath, Len(ONEDRIVE_WEB_PREFIX) + 1)
    If InStr(relativePath, "/") = 0 Then
        Exit Function
    End If

    relativePath = Mid$(relativePath, InStr(relativePath, "/") + 1)
    relativePath = Replace(relativePath, "/", Application.PathSeparator)

    oneDriveRoot = Environ$("OneDrive")
    If Len(oneDriveRoot) > 0 Then
        personalPath = NormalizeWorkbookPath(oneDriveRoot & Application.PathSeparator & relativePath)
        If FileExists(personalPath) Then
            ResolveOneDriveLocalPath = personalPath
            Exit Function
        End If
    End If

    personalPath = Environ$("UserProfile")
    If Len(personalPath) > 0 Then
        personalPath = NormalizeWorkbookPath(personalPath & Application.PathSeparator & "OneDrive" & Application.PathSeparator & relativePath)
        If FileExists(personalPath) Then
            ResolveOneDriveLocalPath = personalPath
        End If
    End If
End Function

Private Function GetParentFolderPath(ByVal filePath As String) As String
    Dim fso As Object

    If Len(filePath) = 0 Then
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    GetParentFolderPath = fso.GetParentFolderName(filePath)
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fso As Object
    Dim parentFolder As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        parentFolder = fso.GetParentFolderName(folderPath)
        If Len(parentFolder) = 0 Or Not fso.FolderExists(parentFolder) Then
            Err.Raise vbObjectError + 1012, , "ログ出力先の親フォルダが見つかりません。 parent=" & parentFolder & " / target=" & folderPath
        End If
        fso.CreateFolder folderPath
    End If
End Sub

Private Function FolderExists(ByVal folderPath As String) As Boolean
    Dim fso As Object

    If Len(folderPath) = 0 Then
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(folderPath)
End Function

Private Function EscapeLogValue(ByVal value As Variant) As String
    Dim textValue As String

    If IsError(value) Then
        EscapeLogValue = "#ERROR"
        Exit Function
    End If

    If IsEmpty(value) Or IsNull(value) Then
        Exit Function
    End If

    If IsDate(value) Then
        textValue = Format$(CDate(value), "yyyy/mm/dd hh:nn:ss")
    Else
        textValue = CStr(value)
    End If

    textValue = Replace(textValue, vbCrLf, " ")
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    textValue = Replace(textValue, vbTab, " ")
    EscapeLogValue = textValue
End Function

Private Sub SaveAppState(ByRef state As AppState)
    state.ScreenUpdating = Application.ScreenUpdating
    state.EnableEvents = Application.EnableEvents
    state.DisplayAlerts = Application.DisplayAlerts
    state.Calculation = Application.Calculation
End Sub

Private Sub ApplyFastSettings()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub RestoreAppState(ByRef state As AppState)
    Application.ScreenUpdating = state.ScreenUpdating
    Application.EnableEvents = state.EnableEvents
    Application.DisplayAlerts = state.DisplayAlerts
    Application.Calculation = state.Calculation
End Sub

Private Function TargetSheetNames() As Variant
    TargetSheetNames = Array("～10日", "11～20日", "21日～")
End Function

Private Function BlockStartColumns() As Variant
    BlockStartColumns = Array(1, 11, 21, 31, 41, 51, 61, 71, 81, 91, 101)
End Function

Private Function MasterHeaders() As Variant
    MasterHeaders = Array( _
        "顧客No", _
        "顧客No判定状態", _
        "内部取込ID", _
        "日付", _
        "年月", _
        "施術時間", _
        "分類", _
        "担当", _
        "患者名", _
        "料金合計", _
        "店舗ID", _
        "店舗名", _
        "元ファイル名", _
        "元ファイルパス", _
        "元シート名", _
        "取込日時")
End Function

Private Function ReferenceHeaders() As Variant
    ReferenceHeaders = Array("店舗ID", "店舗名", "使用中フラグ")
End Function

Private Function ExecutionLogHeaders() As Variant
    ExecutionLogHeaders = Array("実行日時", "ファイル名", "シート名", "ステータス", "読取件数", "出力件数", "メッセージ")
End Function

Private Function DetailLogHeaders() As Variant
    DetailLogHeaders = Array("実行日時", "ファイル名", "シート名", "日付", "患者名", "担当", "種別", "内容")
End Function
