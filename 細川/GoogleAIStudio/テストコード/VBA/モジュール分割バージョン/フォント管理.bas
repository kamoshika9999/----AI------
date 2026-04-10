Option Explicit

Public Function GetOrCreateFontScratchSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SCRATCH_SHEET_FONT)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error Resume Next
        ws.Name = SCRATCH_SHEET_FONT
        On Error GoTo 0
        ws.Range("A1").Value = "（フォント選択用・削除しないでください）"
        ws.Visible = xlSheetVeryHidden
    End If
    Set フォント作業用シートを取得または作成 = ws
End Function

Public Sub RestoreCellFontProps(ByVal r As Range, ByVal oldName As String, _
    ByVal oldSize As Variant, ByVal oldBold As Variant, ByVal oldItalic As Variant, _
    ByVal oldUnderline As Variant, ByVal oldColor As Variant, ByVal oldStrike As Variant)
    On Error Resume Next
    With r.Font
        .Name = oldName
        If Not IsEmpty(oldSize) Then .Size = oldSize
        If Not IsEmpty(oldBold) Then .Bold = oldBold
        If Not IsEmpty(oldItalic) Then .Italic = oldItalic
        If Not IsEmpty(oldUnderline) Then .Underline = oldUnderline
        If Not IsEmpty(oldColor) Then .Color = oldColor
        If Not IsEmpty(oldStrike) Then .Strikethrough = oldStrike
    End With
    On Error GoTo 0
End Sub

Public Function FontNameExistsInExcel(ByVal fontName As String) As Boolean
    Dim i As Long
    For i = 1 To Application.FontNames.Count
        If StrComp(Application.FontNames(i), fontName, vbTextCompare) = 0 Then
            フォント名がExcel一覧に存在するか = True
            Exit Function
        End If
    Next i
End Function

Public Sub ApplyFontToAllSheetCells(ByVal fontName As String, ByRef skippedOut As String)
    Dim ws As Worksheet
    Dim ur As Range
    Dim rangeErr As Boolean
    
    skippedOut = ""
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        rangeErr = False
        Set ur = Nothing
        Set ur = ws.UsedRange
        If Err.Number <> 0 Then
            skippedOut = skippedOut & "・" & ws.Name & "（UsedRange: " & Err.Description & "）" & vbCrLf
            Err.Clear
            rangeErr = True
        End If
        If Not rangeErr Then
            ur.Font.Name = fontName
            If Err.Number <> 0 Then
                skippedOut = skippedOut & "・" & ws.Name & "（Font.Name: " & Err.Description & "）" & vbCrLf
                Err.Clear
            End If
        End If
        On Error GoTo 0
    Next ws
End Sub

' 段階1・段階2のコア処理が成功した直後に呼ぶ。MsgBox なしで BIZ UDPゴシックを全シート UsedRange に付与しメインを AutoFit。
Public Sub 配台_全シートフォントBIZ_UDP_自動適用()
    Dim skipped As String
    On Error Resume Next
    全シートのセルにフォント適用 BIZ_UDP_GOTHIC_FONT_NAME, skipped
    メインシート_AからK列_AutoFit
    結果_主要4結果シート_列オートフィット
    On Error GoTo 0
End Sub

' Excel 標準の［フォント］ダイアログで選んで全シートに適用
Public Sub 全シートフォントをリストから選択して統一()
    Dim wsScratch As Worksheet
    Dim r As Range
    Dim prevWs As Worksheet
    Dim prevVis As XlSheetVisibility
    Dim oldName As String
    Dim oldSize As Variant
    Dim oldBold As Variant
    Dim oldItalic As Variant
    Dim oldUnderline As Variant
    Dim oldColor As Variant
    Dim oldStrike As Variant
    Dim picked As String
    Dim skipped As String
    
    Set prevWs = ActiveSheet
    Set wsScratch = フォント作業用シートを取得または作成()
    prevVis = wsScratch.Visible
    
    Set r = wsScratch.Range("A1")
    With r.Font
        oldName = .Name
        oldSize = .Size
        oldBold = .Bold
        oldItalic = .Italic
        oldUnderline = .Underline
        oldColor = .Color
        oldStrike = .Strikethrough
    End With
    
    wsScratch.Visible = xlSheetVisible
    wsScratch.Activate
    r.Select
    
    If Not Application.Dialogs(xlDialogFormatFont).Show Then
        セルのフォント属性を復元 r, oldName, oldSize, oldBold, oldItalic, oldUnderline, oldColor, oldStrike
        wsScratch.Visible = prevVis
        On Error Resume Next
        prevWs.Activate
        On Error GoTo 0
        Exit Sub
    End If
    
    picked = r.Font.Name
    セルのフォント属性を復元 r, oldName, oldSize, oldBold, oldItalic, oldUnderline, oldColor, oldStrike
    wsScratch.Visible = prevVis
    On Error Resume Next
    prevWs.Activate
    On Error GoTo 0
    
    配台マクロ_全シート保護を試行解除
    スプラッシュ_手順文を設定 "フォント「" & picked & "」を全シートへ適用しています…"
    On Error GoTo Fail
    全シートのセルにフォント適用 picked, skipped
    On Error GoTo 0
    
    If Len(skipped) = 0 Then
        スプラッシュ_手順文を設定 "全シートのフォントを「" & picked & "」に設定しました。"
        m_animMacroSucceeded = True
    Else
        スプラッシュ_手順文を設定 "フォントは適用しましたが、一部シートをスキップしました（ダイアログで詳細を確認してください）。"
        MsgBox "フォント「" & picked & "」を設定しました。スキップしたシート:" & vbCrLf & vbCrLf & skipped, vbExclamation
        m_animMacroSucceeded = True
    End If
    On Error Resume Next
    メインシート_AからK列_AutoFit
    結果_主要4結果シート_列オートフィット
    配台マクロ_対象シートを条件どおりに保護
    On Error GoTo 0
    Exit Sub
    
Fail:
    On Error Resume Next
    配台マクロ_対象シートを条件どおりに保護
    On Error GoTo 0
    MsgBox "フォント設定でエラー: " & Err.Description, vbCritical
End Sub

Public Sub 全シートフォントを選択して統一()
    Call アニメ付き_全シートフォントをリストから選択して統一
End Sub

Public Sub 全シートフォントを手入力で統一()
    Dim v As Variant
    Dim fontName As String
    Dim skipped As String
    
    v = Application.InputBox( _
        "適用するフォント名を入力してください。" & vbCrLf & _
        "（ホームのフォントボックスと同じ表記）", _
        "全シートのフォント統一（手入力）", _
        BIZ_UDP_GOTHIC_FONT_NAME, _
        Type:=2)
    If VarType(v) = vbBoolean Then Exit Sub
    
    fontName = Trim$(CStr(v))
    If Len(fontName) = 0 Then
        MsgBox "フォント名が空のため中止しました。", vbExclamation
        Exit Sub
    End If
    
    If Not フォント名がExcel一覧に存在するか(fontName) Then
        If MsgBox( _
            "フォント「" & fontName & "」が一覧に見つかりませんでした。" & vbCrLf & _
            "このまま適用を試みますか？", _
            vbQuestion Or vbYesNo, "確認") = vbNo Then
            Exit Sub
        End If
    End If
    
    配台マクロ_全シート保護を試行解除
    スプラッシュ_手順文を設定 "フォント「" & fontName & "」を全シートへ適用しています…"
    On Error GoTo FailHand
    全シートのセルにフォント適用 fontName, skipped
    On Error GoTo 0
    
    If Len(skipped) = 0 Then
        スプラッシュ_手順文を設定 "全シートのフォントを「" & fontName & "」に設定しました。"
        m_animMacroSucceeded = True
    Else
        スプラッシュ_手順文を設定 "フォントは適用しましたが、一部シートをスキップしました（ダイアログで詳細を確認してください）。"
        MsgBox "フォント「" & fontName & "」を設定しました。スキップ:" & vbCrLf & vbCrLf & skipped, vbExclamation
        m_animMacroSucceeded = True
    End If
    On Error Resume Next
    メインシート_AからK列_AutoFit
    結果_主要4結果シート_列オートフィット
    配台マクロ_対象シートを条件どおりに保護
    On Error GoTo 0
    Exit Sub
    
FailHand:
    On Error Resume Next
    配台マクロ_対象シートを条件どおりに保護
    On Error GoTo 0
    MsgBox "フォント設定でエラー: " & Err.Description, vbCritical
End Sub

Public Sub 全シートフォント_BIZ_UDPゴシックに統一()
    Dim skipped As String
    スプラッシュ_手順文を設定 "全シートのフォントを「" & BIZ_UDP_GOTHIC_FONT_NAME & "」へ適用しています…"
    配台マクロ_全シート保護を試行解除
    On Error GoTo FailB
    全シートのセルにフォント適用 BIZ_UDP_GOTHIC_FONT_NAME, skipped
    On Error GoTo 0
    
    If Len(skipped) = 0 Then
        スプラッシュ_手順文を設定 "全シートのフォントを「" & BIZ_UDP_GOTHIC_FONT_NAME & "」に設定しました。"
        m_animMacroSucceeded = True
    Else
        スプラッシュ_手順文を設定 "フォントは適用しましたが、一部シートをスキップしました（ダイアログで詳細を確認してください）。"
        MsgBox "フォントを設定しました。スキップ:" & vbCrLf & vbCrLf & skipped, vbExclamation
        m_animMacroSucceeded = True
    End If
    On Error Resume Next
    メインシート_AからK列_AutoFit
    結果_主要4結果シート_列オートフィット
    配台マクロ_対象シートを条件どおりに保護
    On Error GoTo 0
    Exit Sub
    
FailB:
    On Error Resume Next
    配台マクロ_対象シートを条件どおりに保護
    On Error GoTo 0
    MsgBox "フォント設定でエラー: " & Err.Description, vbCritical
End Sub

'==============================================================================
' 列設定_結果_タスク一覧 ? 表示列(B)と連動するフォームのチェックボックスを配置
' 開発タブ → 挿入 → フォーム コントロールの「チェックボックス」と同等。
' 再実行すると既存のチェックボックスを削除して付け直します（セルの TRUE/FALSE は保持）。
'==============================================================================
Public Sub 列設定_結果_タスク一覧_チェックボックスを配置()
    Dim ws As Worksheet
    Dim r As Long
    Dim lastR As Long
    Dim cb As CheckBox
    Dim rng As Range
    Dim linkAddr As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_COL_CONFIG_RESULT_TASK)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_COL_CONFIG_RESULT_TASK & "」が見つかりません。", vbExclamation, "列設定"
        Exit Sub
    End If

    On Error GoTo FailChk
    Do While ws.CheckBoxes.Count > 0
        ws.CheckBoxes(1).Delete
    Loop

    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then
        MsgBox "データ行がありません（1行目は見出し、2行目以降に列名を入れてください）。", vbInformation
        Exit Sub
    End If

    linkAddr = "'" & Replace(ws.Name, "'", "''") & "'!"

    For r = 2 To lastR
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then GoTo NextLoop

        Set rng = ws.Cells(r, 2)
        If Len(Trim$(CStr(rng.Value))) = 0 Then
            rng.Value = True
        End If

        Set cb = ws.CheckBoxes.Add(rng.Left + 2, rng.Top + 0.5, 50, 14)
        With cb
            .LinkedCell = linkAddr & rng.Address(True, True)
            .Caption = ""
        End With
NextLoop:
    Next r

    MsgBox "チェックボックスを配置しました。" & vbCrLf & _
        "表示列(B)の TRUE/FALSE と連動します。", vbInformation
    Exit Sub

FailChk:
    MsgBox "チェックボックス配置でエラー: " & Err.Description, vbCritical
End Sub

'==============================================================================
' 列設定_結果_タスク一覧 → 結果_タスク一覧 へ列順・列非表示を適用（Python / xlwings）
' 図形のマクロ: 「アニメ付き_列設定_結果_タスク一覧_列順表示をPython適用」（押下アニメ付き）。
'   重複列名の整理のみ: 「アニメ付き_列設定_結果_タスク一覧_重複列名を整理」。
'   本体を直指定すると ボタン押下アニメーション が動かない。
' ・事前に「列設定」シートで列名・表示を編集してから実行。
' ・Excel で本ブックを開いたまま（xlwings が接続）。保存してから実行推奨。
'==============================================================================
Public Sub 列設定_結果_タスク一覧_列順表示をPython適用()
    Dim wsh As Object
    Dim runBat As String
    Dim targetDir As String
    Dim exitCode As Long
    Dim wsRes As Worksheet
    Dim wsCfg As Worksheet
    Dim prevScreen As Boolean

    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then
        MsgBox "先にこの Excel ファイルを保存してください。", vbExclamation, "列設定の適用"
        Exit Sub
    End If

    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets(SHEET_RESULT_TASK_LIST)
    Set wsCfg = ThisWorkbook.Worksheets(SHEET_COL_CONFIG_RESULT_TASK)
    On Error GoTo 0
    If wsRes Is Nothing Then
        MsgBox "シート「" & SHEET_RESULT_TASK_LIST & "」がありません。", vbExclamation, "列設定の適用"
        Exit Sub
    End If
    If wsCfg Is Nothing Then
        MsgBox "シート「" & SHEET_COL_CONFIG_RESULT_TASK & "」がありません。", vbExclamation, "列設定の適用"
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Set wsh = CreateObject("WScript.Shell")
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName

    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    スプラッシュ_手順文を設定 "列設定: Python で結果タスク一覧の列順・表示を適用しています…"
    runBat = "@echo off" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & "chcp 65001>nul" & vbCrLf & _
             "py -3 -u python\apply_result_task_column_layout.py" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo [column-layout] ERRORLEVEL=%ERRORLEVEL%" & vbCrLf & _
             "exit /b %ERRORLEVEL%"
    exitCode = 一時CMDをコンソールレイアウト付きで実行(wsh, runBat)
    Application.ScreenUpdating = prevScreen

    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets(SHEET_RESULT_TASK_LIST)
    If Not wsRes Is Nothing Then
        結果シート_列幅_AutoFit非表示を維持 wsRes
    End If
    On Error GoTo 0

    If exitCode <> 0 Then
        MsgBox "Python の終了コードが " & CStr(exitCode) & " です。" & vbCrLf _
            & "log\execution_log.txt を確認してください。", vbExclamation, "列設定の適用"
    Else
        スプラッシュ_手順文を設定 "「" & SHEET_RESULT_TASK_LIST & "」の列順・表示を「" & SHEET_COL_CONFIG_RESULT_TASK & "」に合わせました。"
        m_animMacroSucceeded = True
    End If
End Sub

'==============================================================================
' 列設定_結果_タスク一覧: 列名の重複行を除き A:B を書き直す（Python）。結果_タスク一覧は変更しない。
' チェックボックスを B 列に付けている場合は、実行後に「列設定_結果_タスク一覧_チェックボックスを配置」を推奨。
'==============================================================================
Public Sub 列設定_結果_タスク一覧_重複列名を整理()
    Dim wsh As Object
    Dim runBat As String
    Dim targetDir As String
    Dim exitCode As Long
    Dim wsCfg As Worksheet
    Dim prevScreen As Boolean

    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then
        MsgBox "先にこの Excel ファイルを保存してください。", vbExclamation, "列設定の整理"
        Exit Sub
    End If

    On Error Resume Next
    Set wsCfg = ThisWorkbook.Worksheets(SHEET_COL_CONFIG_RESULT_TASK)
    On Error GoTo 0
    If wsCfg Is Nothing Then
        MsgBox "シート「" & SHEET_COL_CONFIG_RESULT_TASK & "」がありません。", vbExclamation, "列設定の整理"
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Set wsh = CreateObject("WScript.Shell")
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName

    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    スプラッシュ_手順文を設定 "列設定: Python で重複列名を整理しています…"
    runBat = "@echo off" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & "chcp 65001>nul" & vbCrLf & _
             "py -3 -u python\dedupe_result_task_column_config_sheet.py" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo [dedupe-column-config] ERRORLEVEL=%ERRORLEVEL%" & vbCrLf & _
             "exit /b %ERRORLEVEL%"
    exitCode = 一時CMDをコンソールレイアウト付きで実行(wsh, runBat)
    Application.ScreenUpdating = prevScreen

    If exitCode <> 0 Then
        MsgBox "Python の終了コードが " & CStr(exitCode) & " です。" & vbCrLf _
            & "log\execution_log.txt を確認してください。", vbExclamation, "列設定の整理"
    Else
        スプラッシュ_手順文を設定 "「" & SHEET_COL_CONFIG_RESULT_TASK & "」の重複列名を除き A:B を更新しました。（チェックボックス利用時は配置マクロの再実行を推奨）"
        m_animMacroSucceeded = True
    End If
End Sub

'==============================================================================
' 配台計画_タスク入力 → 「配台試行順番」を段階2と同じ基準で再計算（Python / xlwings）
' 図形のマクロ: 「アニメ付き_配台試行順番_タスク入力をPythonで更新」（業務ロジック.bas。押下アニメ＋スプラッシュ付き）。
'   本体を直指定するとボタン押下アニメーションが動かない場合があります。
' ・配台不要を手動でクリアしたあとなど、試行順を付け直したいときに使用。
' ・段階2と同様、読込後に設定シート・分割行・配台不要ルールを DataFrame に反映してから試行順を付け直す（planning_core と同一）。
' ・Excel で本ブックを開いたまま。保存してから実行推奨。
' ・ボタン自動配置: マクロ「配台計画_タスク入力_配台試行順番更新ボタンを配置」（業務ロジック.bas）
'==============================================================================
Public Sub 配台試行順番_配台計画タスク入力をPythonで更新()
    Dim wsh As Object
    Dim runBat As String
    Dim targetDir As String
    Dim exitCode As Long
    Dim wsPlan As Worksheet
    Dim prevScreen As Boolean
    Dim didUnlock As Boolean

    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then
        MsgBox "先にこの Excel ファイルを保存してください。", vbExclamation, "配台試行順番の更新"
        Exit Sub
    End If

    On Error Resume Next
    Set wsPlan = ThisWorkbook.Worksheets("配台計画_タスク入力")
    On Error GoTo 0
    If wsPlan Is Nothing Then
        MsgBox "シート「配台計画_タスク入力」がありません。", vbExclamation, "配台試行順番の更新"
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    didUnlock = False
    On Error Resume Next
    配台マクロ_全シート保護を試行解除
    didUnlock = True
    On Error GoTo 0

    Set wsh = CreateObject("WScript.Shell")
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName

    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    スプラッシュ_手順文を設定 "配台試行順番: Python で配台計画_タスク入力を再計算しています…"
    runBat = "@echo off" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & "chcp 65001>nul" & vbCrLf & _
             "py -3 -u python\refresh_plan_input_dispatch_trial_order.py" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo [plan-dispatch-trial-order] ERRORLEVEL=%ERRORLEVEL%" & vbCrLf & _
             "exit /b %ERRORLEVEL%"
    exitCode = 一時CMDをコンソールレイアウト付きで実行(wsh, runBat)
    Application.ScreenUpdating = prevScreen

    If didUnlock Then
        On Error Resume Next
        配台マクロ_対象シートを条件どおりに保護 targetDir
        On Error GoTo 0
    End If

    If exitCode <> 0 Then
        MsgBox "Python の終了コードが " & CStr(exitCode) & " です。" & vbCrLf _
            & "log\execution_log.txt を確認してください。", vbExclamation, "配台試行順番の更新"
    Else
        スプラッシュ_手順文を設定 "「配台計画_タスク入力」の配台試行順番を更新し、行を試行順の昇順に並べ替えました。"
        m_animMacroSucceeded = True
    End If
End Sub

'==============================================================================
' デバッグ: ブックを開いたまま「どのシートが COM 的に触りにくいか」を一覧する
' 開発タブ → マクロ → 「COM操作テスト_全シートをログに出す」を実行。
' シート「COM操作テストログ」を末尾に作成し、シートごとの OK/NG を出します。
' （Excel 本体からの操作の目安。外部 Python/pywin32 の COM とはプロセスが異なる場合があります）
'==============================================================================
