Option Explicit

' =========================================================
' Windows APIの宣言（ボタンの待機・アニメーション処理用）
' =========================================================
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As LongPtr, ByVal dwFlags As Long) As Long
    Public Declare PtrSafe Function PlaySoundW Lib "winmm.dll" (ByVal pszSound As LongPtr, ByVal hmod As LongPtr, ByVal fdw As Long) As Long
    Public Declare PtrSafe Function mciSendStringW Lib "winmm.dll" (ByVal lpstrCommand As LongPtr, ByVal lpstrReturnString As LongPtr, ByVal uReturnLength As Long, ByVal hwndCallback As LongPtr) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As LongPtr, ByVal lpWindowName As String) As LongPtr
    Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
    Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Public Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As LongPtr) As Long
    Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
    Public Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hwnd As LongPtr, lpPoint As POINTAPI) As Long
    Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    #If Win64 Then
    Public Declare PtrSafe Function SplashGetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrW" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Public Declare PtrSafe Function SplashSetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrW" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
    Public Declare PtrSafe Function SplashGetWindowLongPtr Lib "user32" Alias "GetWindowLongW" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function SplashSetWindowLongPtr Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    #End If
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
    Public Declare Function PlaySoundW Lib "winmm.dll" (ByVal pszSound As Long, ByVal hmod As Long, ByVal fdw As Long) As Long
    Public Declare Function mciSendStringW Lib "winmm.dll" (ByVal lpstrCommand As Long, ByVal lpstrReturnString As Long, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As String) As Long
    Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
    Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
    Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If
Public Const SW_HIDE As Long = 0
Public Const SND_ASYNC As Long = &H1
Public Const SND_ALIAS As Long = &H10000
Public Const SND_FILENAME As Long = &H20000
' 完了チャイム WAV（ブックと同じフォルダ配下）。無ければ下記 URL から取得（リポジトリが public かつブランチ一致が前提）
Public Const MACRO_COMPLETE_CHIME_REL_DIR As String = "sounds"
Public Const MACRO_COMPLETE_CHIME_FILE_NAME As String = "macro_complete_chime.wav"
Public Const MACRO_COMPLETE_CHIME_DOWNLOAD_URL As String = "https://raw.githubusercontent.com/kamoshika9999/----AI------/xlwings%E7%A9%8D%E6%A5%B5%E9%81%A9%E7%94%A8/%E7%B4%B0%E5%B7%9D/GoogleAIStudio/%E3%83%86%E3%82%B9%E3%83%88%E3%82%B3%E3%83%BC%E3%83%89/sounds/macro_complete_chime.wav"
' 完了 MP3（sounds 配下）。トラックは「設定」D4 に 1?4（空・不正は 1）。実ファイル名を変えたら定数だけ合わせる。
Public Const MACRO_COMPLETE_MP3_1 As String = "穏やかな応答1.mp3"
Public Const MACRO_COMPLETE_MP3_2 As String = "穏やかな応答2.mp3"
Public Const MACRO_COMPLETE_MP3_3 As String = "穏やかな応答3.mp3"
Public Const MACRO_COMPLETE_MP3_4 As String = "穏やかな応答4.mp3"
' スプラッシュ中 BGM（sounds 配下）。MCI 固定 alias でループ再生し、終了時はフェードアウト後 close。
Public Const MACRO_START_BGM_FILENAME As String = "Glass_Architecture1.mp3"
Public Const MACRO_START_BGM_ALIAS As String = "pm_ai_glassbg"
' frmMacroSplash を Excel メイン枠の下端・水平中央に合わせるときの下端からの余白（px）
Public Const SPLASH_EXCEL_BOTTOM_GAP_PX As Long = 0
' 段階1/2 の cmd コンソールをプライマリ画面 左上・全幅・高さ 1/4 に置くときに使用
Public Const SM_CXSCREEN As Long = 0
Public Const SM_CYSCREEN As Long = 1
Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_SHOWWINDOW As Long = &H40
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_FRAMECHANGED As Long = &H20
Public Const SWP_NOACTIVATE As Long = &H10
Public Const LOGPIXELSX As Long = 88
Public Const LOGPIXELSY As Long = 90
Public Const GWL_STYLE As Long = -16
' コンソール枠除去用（Caption/ThickFrame/MinMax/SysMenu）。環境によっては conhost が無視する場合あり
Public Const WS_CONSOLE_OVERLAY_STRIP As Long = &HCF0000
' D3=false オーバーレイ待機（ms）。短すぎると Python→xlwings の COM 同期と Excel 主スレッドが奪い合い「応答していません」になりやすいため SPLASH_LOG と同程度にする
Public Const STAGE12_CMD_OVERLAY_POLL_MS As Long = 1200
' ログ枠（TextBox）の画面矩形よりコンソール外周を四辺いくら縮めるか（px）。conhost の内側の黒余白で「枠が大きい」と感じる場合に調整
Public Const STAGE12_CMD_OVERLAY_RECT_INSET_PX As Long = 10
' True=上記オーバーレイ時にタイトルバー等を API で外す（うまくいかない PC では False）
Public Const STAGE12_CMD_OVERLAY_BORDERLESS As Boolean = True
' D3=false かつスプラッシュ表示中の cmd 見た目をログ枠に重ねる（Exec＋FindWindow＋SetWindowPos／conhost／枠除去など）。
' ★ True にすると段階1/2 の Python が同一 Excel に xlwings で COM 同期するとき、VBA 主スレッドが待機ループ＋HWND 操作で奪い合いになりやすく、処理が固まったように見えることが多い。実用では False 推奨。
' False=同期 Run のみ（cmd は OS 既定表示）。xlwings 同期・応答なし疑い時は必ず False。
Public Const STAGE12_D3FALSE_SPLASH_CONSOLE_LAYOUT As Boolean = False
' 段階1/2 の cmd: True=ウィンドウ非表示。進捗は UserForm（txtExecutionLog）＝ execution_log.txt を Exec 待機中にポーリング。py の余剰 stdout/stderr は nul へ。False=画面上部にコンソール（1/4 高さ・全幅）
' 実効値は Stage12CmdHideWindowEffective（シート「設定_環境変数」の STAGE12_CMD_HIDE_WINDOW → OS 環境変数同名 → 本定数）
Public Const STAGE12_CMD_HIDE_WINDOW As Boolean = True
' True=スプラッシュに実行ログを「処理中」に表示したい → 同期 RunPython は使えないため常に cmd+Exec+ポーリング（本 True のとき STAGE12_USE_XLWINGS_RUNPYTHON は実質無視）。False=下記 RunPython の可否に従う
Public Const STAGE12_USE_XLWINGS_SPLASH_LOG As Boolean = True
' True かつ STAGE12_USE_XLWINGS_SPLASH_LOG=False のときのみ xlwings.RunPython+runpy.run_path（Tools→参照に xlwings）。実行中はログ枠はほぼ更新されず終了後に一括表示。進捗表示優先なら SPLASH_LOG=True のまま（cmd になる）
Public Const STAGE12_USE_XLWINGS_RUNPYTHON As Boolean = True

' 全シートフォント統一マクロで使用（PC にフォントが入っている必要があります）
Public Const BIZ_UDP_GOTHIC_FONT_NAME As String = "BIZ UDPゴシック"
' フォント選択ダイアログ用の一時シート（VeryHidden）
Public Const SCRATCH_SHEET_FONT As String = "_FontPick"
' 全シートフォント統一で保護シートを一時解除するときのパスワード。空＝パスワード無しのシートのみ解除可。パスワード付きで統一したい場合はここに同じパスワードを設定（再保護にも使用）。
Public Const SHEET_FONT_UNPROTECT_PASSWORD As String = ""
' planning_core の COLUMN_CONFIG_SHEET_NAME / RESULT_TASK_SHEET_NAME と一致させる
Public Const SHEET_COL_CONFIG_RESULT_TASK As String = "列設定_結果_タスク一覧"
Public Const SHEET_RESULT_TASK_LIST As String = "結果_タスク一覧"
Public Const SHEET_RESULT_EQUIP_SCHEDULE As String = "結果_設備毎の時間割"
Public Const SHEET_RESULT_EQUIP_BY_MACHINE As String = "結果_設備毎の時間割_機械名毎"
Public Const SHEET_RESULT_CALENDAR_ATTEND As String = "結果_カレンダー(出勤簿)"
Public Const SHEET_RESULT_EQUIP_GANTT As String = "結果_設備ガント"
Public Const SHEET_SETTINGS As String = "設定"
' planning_core.EXCLUDE_RULES_SHEET_NAME / EXCLUDE_RULE_COL_* と見出しを一致させる（シート作成は VBA、行同期は Python）
Public Const SHEET_EXCLUDE_ASSIGNMENT As String = "設定_配台不要工程"
' workbook_env_bootstrap.WORKBOOK_ENV_SHEET_NAME と一致（A=変数名・B=値・C=説明）
Public Const SHEET_WORKBOOK_ENV As String = "設定_環境変数"
' シートのタブ表示と並び順を一覧・適用する（VBA のみ。Python 連携なし）
Public Const SHEET_SHEET_VISIBILITY As String = "設定_シート表示"
' Ctrl+Shift+テンキー - → メインシートへ（Application.OnKey）。^=Ctrl、+=Shift、{109}=テンキー -（vbKeySubtract）。{SUBTRACT} は環境により OnKey が 1004 で失敗するため数値コードを使用
' 起動ショートカット.bas 等から OnKey 登録で参照するため Public
Public Const SHORTCUT_MAIN_SHEET_ONKEY As String = "^+{109}"
' Gemini: 暗号化時のパスフレーズはマクロで入力（社内手順の値）。復号は planning_core のソース内定数のみ（当ファイル・シートにパスフレーズを書かない）。
' planning_core.MASTER_FILE / SHEET_MACHINE_CALENDAR と一致
Public Const MASTER_WORKBOOK_FILE As String = "master.xlsm"
Public Const SHEET_MACHINE_CALENDAR As String = "機械カレンダー"

' ★ 本ファイルは「生産管理_AI配台テスト.xlsm」の標準モジュール用テキストバックアップ（master.xlsm 用は master_xlsm_VBA.txt）
' ★ planning_core は同フォルダの master.xlsm を MASTER として読む。上書き用アプリ JSON は json\（API 料金累計は API_Payment\）。
' ★「設定」D3=スプラッシュの txtExecutionLog へ execution_log を反映するか（true/空=する・false=しない）。false のときポーリング無し。D3=false 時の cmd は既定で同期 Run。STAGE12_D3FALSE_SPLASH_CONSOLE_LAYOUT=True はログ枠重ね（実験用）だが xlwings 同期と併用で固まりやすいので False 推奨。
' ★「設定」B4=配台結果 xlsx に Python が埋め込むフォント名（空なら書体名を付けず、取り込み後も「全シートフォント」が維持されやすい）。B5=ポイント（B4 指定時、空なら 11）。
' ★ TASK_INPUT_WORKBOOK には本ブック（ThisWorkbook）を渡す。
' ★ メインに「master.xlsm を開く」ボタン: 開発タブ→マクロ→「メインシート_master開くボタンを配置」を1回実行（既存ボタンと重なる場合は位置をドラッグ調整）
' ★ 「配台計画_タスク入力」に配台試行順番の再計算ボタン: マクロ「配台計画_タスク入力_配台試行順番更新ボタンを配置」（同名図形は削除して付け直し）
' ★「設定_配台不要工程」は Python で新規作成しない。段階1・段階2 の先頭で 設定_配台不要工程_シートを確保（見出し・表示 xlSheetVisible）。
' ★「設定_環境変数」は workbook_env_bootstrap が import 前に読む。段階1・段階2 先頭で 設定_環境変数_シートを確保（見出し・不足キーのみ追記。既存行は上書きしない）。
' ★「設定_シート表示」は A=並び順（1 始まり・小さいほど左のタブ）・B=シート名・C=表示（ドロップダウンはインライン一覧。F2:F4 は候補の目安）。マクロ「設定_シート表示_一覧をブックから再取得」「設定_シート表示_ブックへ適用」。段階1/2 成功完了時は一覧更新のあと「ブックへ適用」まで自動実行（適用末尾で再び一覧同期）。当シートは常に表示。
' ★ アニメ付き_* マクロは処理中に UserForm「frmMacroSplash」を表示する。作成手順は frmMacroSplash_VBA.txt。
'   ・表示位置は Application.hwnd のウィンドウ矩形に対し下端・水平中央（SPLASH_EXCEL_BOTTOM_GAP_PX）。スプラッシュ_フォームを最前面へ のたびに再配置（長時間処理中の Excel 移動に追従）。
'   ・STAGE12_USE_XLWINGS_SPLASH_LOG=True … 段階1/2 の Python は必ず cmd+Exec。待機中 スプラッシュ_実行ログ枠を更新 で execution_log.txt をポーリング（固まったように見えない）。
'   ・STAGE12_USE_XLWINGS_SPLASH_LOG=False かつ STAGE12_USE_XLWINGS_RUNPYTHON=True … 同期 xlwings.RunPython（終了後 スプラッシュ_実行ログをパスから読込 で一括）。実行中はログ枠はほぼ動かない。
' ★ 段階1/2: PowerShell は起動しない（cmd＋conhost --headless または cmd）。Exec 待機中に execution_log を UserForm へ。STAGE12_CMD_HIDE_WINDOW（シートまたは OS 環境・既定1）=True で WT 経由の黒画面を避ける。終了コードは log\stage_vba_exitcode.txt 優先。False なら cmd＋SetWindowPos。非表示時 py は 1>nul 2>&1。
'
' ★ xlwings「Show Console」で cmd なしに Python ログを見る（任意・要 xlwings アドイン・参照設定）
'   ・import 解決のため xlwings.RunPython に渡す文字列は runpy.run_path で python\xlwings_console_runner.py を実行する形式を使う。
'       xlwings.RunPython "import os, runpy, xlwings as xw; wb=xw.Book.caller(); p=os.path.join(os.path.dirname(str(wb.fullname)), 'python', 'xlwings_console_runner.py'); ns=runpy.run_path(p); ns['run_stage1_for_xlwings']()"
'   ・本番は SPLASH_LOG=False かつ RUNPYTHON=True のときだけ Xlwings_コンソールランナー実行。SPLASH_LOG=True のときは cmd（進捗優先）。
'   ・補助: 同フォルダ xlwings.conf.json（PYTHONPATH=python）。runner は log\stage_vba_exitcode.txt に終了コードを書く。

' InstallComponents: winget 失敗時に使う公式 amd64 インストーラ URL（必要なら 3.12 のパッチ版に更新）
Public Const PY_OFFICIAL_INSTALLER_URL As String = "https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe"

' True ならマクロ先頭の ThisWorkbook.RefreshAll をスキップ（接続更新で固まる場合の緊急回避）
Public Const SKIP_WORKBOOK_REFRESH_ALL As Boolean = False
' Power Query 更新前の到達確認（接続先と同一想定の IP）。応答なしなら RefreshAll を行わず処理は継続する。
' 接続先の IP アドレスを指定（Power Query のデータソースと揃える）
Public Const PQ_REFRESH_PING_HOST As String = "192.168.0.101"
' ping -w に渡すタイムアウト（ミリ秒）。合計待ちはあともう数百 ms 程度のことがある。
Public Const PQ_REFRESH_PING_TIMEOUT_MS As Long = 2000

' 段階1コアの最終 Python exitCode（0=成功）。VBAエラー・事前失敗は -1。
Public m_lastStage1ExitCode As Long
Public m_lastStage1ErrMsg As String
' 段階2コア用（ダイアログは呼び出し元で出す）
Public m_lastStage2ErrMsg As String
Public m_lastStage2ExitCode As Long
Public m_stage2PlanImported As Boolean
Public m_stage2MemberImported As Boolean
' TryRefreshWorkbookQueries 失敗時の詳細（MsgBox なし。段階1・2の ErrMsg に連結）
Public m_lastRefreshQueriesErrMsg As String
' スプラッシュ表示中（UserForm「frmMacroSplash」。未インポート時は何も出ずエラーも抑止）
' 業務ロジック・スプラッシュ表示 等の標準モジュールから参照するため Public
Public m_macroSplashShown As Boolean
' MacroSplash_Show で Application.Interactive=False を立てたときだけ Hide で True に戻す
Public m_macroSplashLockedExcel As Boolean
' アニメ付き_スプラッシュ付きで実行 の成功終了時のみチャイム（各処理が True に設定）
Public m_animMacroSucceeded As Boolean
' True のときのみ BGM・完了チャイムを許可（段階1／段階2のスプラッシュ起動マクロが立てる）
Public m_splashAllowMacroSound As Boolean
' スプラッシュ用 BGM（Glass_Architecture.mp3）を MCI で開いているとき True
Public m_macroStartBgmOpen As Boolean
' 段階1/2 中のみ: execution_log.txt のフルパス（スプラッシュ txtExecutionLog ポーリング用。xlwings 時は空）
Public m_splashExecutionLogPath As String
' xlwings 有効で上記が空のとき、stage_vba_exitcode.txt の所在（log フォルダ）だけ渡す
Public m_stageVbaExitCodeLogDir As String
Public Const SPLASH_LOG_MAX_DISPLAY_CHARS As Long = 120000
' UserForm 前面化用（Caption を一意にし FindWindow で HWND を得る。フォームの Caption プロパティも同じ文字列にするとよい）
Public Const SPLASH_FORM_WINDOW_TITLE As String = "PM_AI_MACRO_SPLASH"
' D3=true 時の Exec 待機ループの間隔（ms）。短すぎると xlwings COM と Excel 主スレッドが奪い合いやすい
' xlwings が同一 Excel で COM 同期する間、短い間隔だと主スレッド奪い合いで極端に遅くなることがある（400→1200）
Public Const SPLASH_LOG_POLL_INTERVAL_MS As Long = 1200
' execution_log の直近成功読み取り時の FileLen。サイズ不変なら UTF-8 全文読みをスキップ（負荷・COM 競合軽減）
Public m_splashPollLastFileLen As Long
Public m_splashPollHaveCachedFileLen As Boolean
' execution_log ポーリングで「ファイルはあるが読めない」と判定したとき txtExecutionLog 先頭に1回だけ付与
Public m_splashReadErrShown As Boolean
' 直前に表示したログ全文と同じなら Text 再代入しない（ちらつき防止）
Public m_splashLastLogSnapshot As String
' D3=false オーバーレイ中に txtExecutionLog を非表示にしているとき True（必ず End で戻す）
Public m_splashConsoleOverlayActive As Boolean
' 結果_設備ガント：選択中データ行の太枠（Interior は触らない）
Public mGanttHL_SheetName As String
Public mGanttHL_Row As Long
Public mGanttHL_LastCol As Long

' 段階1/2 cmd 非表示: シート「設定_環境変数」A 列=STAGE12_CMD_HIDE_WINDOW かつ B 非空 → その値。未設定なら Environ("STAGE12_CMD_HIDE_WINDOW")。どちらも空なら STAGE12_CMD_HIDE_WINDOW 定数。
