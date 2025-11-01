#Requires AutoHotkey v2.0
#SingleInstance Off
Persistent(True)

DEBUG_MODE := true
LOG_FILE := A_ScriptDir "\CSP_dialog_center_debug.log"

if DEBUG_MODE
{
    try FileDelete(LOG_FILE)
}

TARGET_PROCESS := "clipstudiopaint.exe"
DIALOG_RATIO_LIMIT := 0.9
MAX_MOVE_ATTEMPTS := 10
MIN_MAIN_DIMENSION := 240
MIN_DIALOG_DIMENSION := 80

gShellHookGui := 0
gShellHookMsg := 0
gPendingRelocate := Map()
gKnownWindows := Map()
gScannerStarted := false

currentPid := DllCall("GetCurrentProcessId", "uint")
DebugLog(Format("CSP dialog center started. PID={}", currentPid))

gMutex := AcquireSingleInstanceMutex("CSPDialogCenter")
if !gMutex
{
    DebugLog("Another instance is already running; exiting.")
    ExitApp
}
DebugLog(Format("Acquired mutex. Handle=0x{:X}", gMutex))

OnExit(ReleaseResources)
SetupTrayMenu()
StartShellHook()
InitializeExistingDialogs()

return

AcquireSingleInstanceMutex(baseName)
{
    prefixes := ["Global\", ""]
    for prefix in prefixes
    {
        fullName := prefix . baseName
        hMutex := DllCall("CreateMutex", "ptr", 0, "int", 1, "str", fullName, "ptr")
        err := A_LastError
        DebugLog(Format("CreateMutex called. name='{}' handle=0x{:X} err={}", fullName, hMutex, err))
        if !hMutex
        {
            ; Try the next prefix if available.
            continue
        }
        if (err = 183)  ; ERROR_ALREADY_EXISTS
        {
            DllCall("CloseHandle", "ptr", hMutex)
            DebugLog(Format("Mutex already exists. name='{}'", fullName))
            return 0
        }
        return hMutex
    }
    DebugLog("Failed to create mutex.")
    return 0
}

SetupTrayMenu()
{
    DebugLog("Initializing tray menu.")
    A_IconTip := "CSP dialog center"
    try A_TrayMenu.Delete()
    catch
    {
        ; Ignore if clearing the tray menu fails (e.g. already empty or localized default items).
        DebugLog("Failed to clear default tray menu entries; continuing.")
    }
    A_TrayMenu.Add("Exit", ExitScript)
    DebugLog("Tray menu ready.")
}

ExitScript(*)
{
    ExitApp
}

StartShellHook()
{
    global gShellHookGui, gShellHookMsg

    ; 隠しウィンドウの作成。+ToolWindow -Caption等で枠やタイトルバーのない軽量なウィンドウを作成して、隠し状態で表示
    gShellHookGui := Gui("+ToolWindow -Caption +OwnDialogs")
    gShellHookGui.Show("Hide")
    DebugLog(Format("Created hidden GUI for shell hook. hwnd=0x{:X}", gShellHookGui.Hwnd))

    ; この登録を行うと、作成した隠しウィンドウが、システム全体で新しいウィンドウが開いたりアクティブ化したときに通知を受け取れるようになる
    if !DllCall("RegisterShellHookWindow", "ptr", gShellHookGui.Hwnd, "int")
    {
        DebugLog("RegisterShellHookWindow failed; exiting.")
        MsgBox("Failed to register the shell hook window.", "CSP dialog center", "Iconx")
        ExitApp
    }

    ; システムから届く"SHELLHOOK"メッセージの数値IDを記憶する。数値IDで判定することにより"SHELLHOOK"か否か判定する
    ; "SHELLHOOK"のように新しい機能の場合固定の数値IDではなく、"SHELLHOOK"などの文字列で登録する方式になっている
    ; OSの再起動などでRegisterShellHookWindow("SHELLHOOK")で初回登録されたとき、新たに数値IDが割り当てられるので都度数値が変わると思った方がよい(同一セッション内で同値になるらしい)
    ; Windowsがシステムイベントを通知してくるときは、この数値IDを使ってメッセージを送信してくるので、目的のメッセージだけ判別可能になる
    gShellHookMsg := DllCall("RegisterWindowMessage", "str", "SHELLHOOK", "uint")
    DebugLog(Format("Registered SHELLHOOK message ID. msg=0x{:X}", gShellHookMsg))
    ; 上で得た"SHELLHOOK"メッセージの数値IDが到着したときは、ShellMessage関数を呼ぶように登録
    OnMessage(gShellHookMsg, ShellMessage)
    DebugLog("Shell hook initialization finished.")
}

InitializeExistingDialogs()
{
    handles := WinGetList()
    DebugLog(Format("Scanning existing windows. total={}", handles.Length))
    for hwnd in handles
    {
        if !IsCSPWindow(hwnd)
            continue
        DebugLog(Format("Existing CSP window detected. hwnd=0x{:X}", hwnd))
        ScheduleRelocate(hwnd)
    }
    StartDialogScanner()
}

; 「移動対象として見つけたウィンドウ」を、少し遅れて処理するキューに登録する係
ScheduleRelocate(hwnd)
{
    ; 特定のウィンドウハンドル(hwnd)が「処理中」であると管理するためのMapで
    ; gPendingRelocateが現在時刻、gKnownWindowsがここまでは再スキャン不要との時刻を示す
    global gPendingRelocate, gKnownWindows

    ; メインウィンドウなら何もせず帰る
    if IsMainCspWindow(hwnd)
    {
        DebugLog(Format("Main window detected (0x{:X}); skipping scheduling.", hwnd))
        return
    }

    ; そのウィンドウが「移動対象として扱ってよい」ものかをチェック。メニュー類など除外対象ならここで帰る
    if !IsWindowEligible(hwnd)
    {
        return
    }

    now := A_TickCount  ; 現在時刻(ミリ秒)取得
    if gPendingRelocate.Has(hwnd)
    {
        ; すでに同じウィンドウで予約済みかを調べる。200ms 以内に再度リクエストが来た場合は「まだ処理待ちだから不要」と判断し、
        ; ログを残してリターンします。通知の連打による無駄な再処理を防ぐ仕掛け。
        ; 200ms以内は"登録自体せず捨てる"
        previous := gPendingRelocate[hwnd]
        if (now - previous < 200)
        {
            DebugLog(Format("Relocation already pending for hwnd=0x{:X}; skipping duplicate schedule.", hwnd))
            return
        }
    }

    ; hwndについて今から500msはスキャン不要とする。これは「hwndが処理待ち」とマーク(キューに入って待っている)している。
    gPendingRelocate[hwnd] := now
    gKnownWindows[hwnd] := now + 500
    DebugLog(Format("Queueing relocation attempt for window 0x{:X}.", hwnd))

    ; 実際の移動処理 TryRelocateWindow を 50ms 後に一度だけ実行するよう予約します。
    ; 少し待ってから動かすのは、ウィンドウが表示されてサイズが確定(確定前はサイズが1x1とかになる)するまで短い時間を与えるため。
    SetTimer(() => TryRelocateWindow(hwnd), -50)
}

; 初期化処理のOnMessageで登録したので、"SHELLHOOK"メッセージがシステムから届くたびにコールバックされる
ShellMessage(wParam, lParam, msg, hwnd)
{
    targetHwnd := lParam  ; lParamに通知対象のウィンドウハンドルが入っている
    if (targetHwnd && IsMainCspWindow(targetHwnd))
    {
        DebugLog(Format("Shell event targeted main window (0x{:X}); ignoring.", targetHwnd))
        return
    }

    ; 対象のイベントかどうかを確認。Mapにしてるのはhandled.Has(wParam)と調べられるようにするため。
    ; 1 → HSHELL_WINDOWCREATED = ウィンドウが生成された
    ; 4 → HSHELL_WINDOWACTIVATED = ウィンドウがアクティブになった
    ; 5 → HSHELL_GETMINRECT = 最小化などで枠が決まるタイミング
    ; 6 → HSHELL_REDRAW = 表示が更新された
    ; 0x8004 → HSHELL_RUDEAPPACTIVATED = 強制的に前面化された
    static handled := Map(1, true, 4, true, 5, true, 6, true, 0x8004, true)
    if !handled.Has(wParam)
    {
        ; 今回対象以外のイベントだったら無視する
        DebugLog(Format("Ignored shell event. code=0x{:X} hwnd=0x{:X}", wParam, lParam))
        return
    }

    DebugLog(Format("Shell event received. event={} (0x{:X}) hwnd=0x{:X}", ShellEventName(wParam), wParam, targetHwnd))
    if !IsWindow(targetHwnd)
    {
        ; すでにハンドルが無効(=ウィンドウが閉じられている)の場合、何もせず帰る
        DebugLog(Format("hwnd=0x{:X} is not a valid window anymore.", targetHwnd))
        return
    }

    if !IsCSPWindow(targetHwnd, &processName)
    {
        ; そのウィンドウがCLIP STUDIO PAINTのものかどうかを判定
        DebugLog(Format("Process name mismatch. proc='{}' hwnd=0x{:X}", processName, targetHwnd))
        return
    }

    ; ここまでくればクリスタの対象イベントなので、「中央に寄せる処理」を行うキューに登録する
    ScheduleRelocate(targetHwnd)
}

; 移動対象のウィンドウ hwnd を中央に寄せようとする関数。attempt は何回目の試行かを示すカウンタ
TryRelocateWindow(hwnd, attempt := 1)
{
    ; 試行回数の上限、ダイアログを動かすときの最小のサイズ、試行待ちリスト
    global MAX_MOVE_ATTEMPTS, MIN_MAIN_DIMENSION, MIN_DIALOG_DIMENSION, gPendingRelocate, gKnownWindows
    
    ; 「hwndが処理待ち」とマーク(キューに入って待っている)なら、マークを解除(キューから削除)する
    if gPendingRelocate.Has(hwnd)
    {
        gPendingRelocate.Delete(hwnd)
    }

    DebugLog(Format("TryRelocateWindow start. hwnd=0x{:X} attempt={}", hwnd, attempt))
    if !IsWindow(hwnd)
    {
        ; ウィンドウ自体すでに存在しなかった場合は何もせず帰る。記憶していた情報があればそれも削除する
        DebugLog(Format("hwnd=0x{:X} no longer exists; aborting.", hwnd))
        if gKnownWindows.Has(hwnd)
        {
            gKnownWindows.Delete(hwnd)
        }
        return
    }

    if IsMainCspWindow(hwnd)
    {
        DebugLog("Main window detected during relocation; skipping.")
        DeferWindowRetry(hwnd, 2000)
        return
    }

    if !IsCSPWindow(hwnd, &processName)
    {
        ; クリスタのウィンドウの通知でなければ、1.5秒間くらいはすべてのメッセージを無視する
        ; 無駄なメッセージが多いので処理負荷を下げる＆クリスタなのに稀にクリスタでないと判定できてしまうことがあるため。
        DebugLog(Format("Process name mismatch ({}); aborting.", processName))
        DeferWindowRetry(hwnd, 1500)
        return
    }

    if !DllCall("IsWindowVisible", "ptr", hwnd, "int")
    {
        ; まだ表示すらされていない状況。表示されるまで80msあけてリトライする
        DebugLog(Format("Window 0x{:X} is not yet visible.", hwnd))
        if (attempt < MAX_MOVE_ATTEMPTS)
        {
            DebugLog("Queueing retry.")
            SetTimer(() => TryRelocateWindow(hwnd, attempt + 1), -80)
        }
        DeferWindowRetry(hwnd, 200)
        return
    }

    ; ウィンドウの状態（通常／最小化など）を取得。失敗したり最小化されている場合は一時停止扱いにして後でやり直す
    try state := WinGetMinMax("ahk_id " hwnd)
    catch error
    {
        DebugLog(Format("Exception while calling WinGetMinMax: {}", ErrorText(error)))
        DeferWindowRetry(hwnd, 500)
        return
    }
    if (state = -1)
    {
        DebugLog("Target window is minimized; skipping.")
        DeferWindowRetry(hwnd, 2000)
        return
    }

    ; ウィンドウスタイルを取得
    try style := WinGetStyle("ahk_id " hwnd)
    catch error
    {
        DebugLog(Format("Exception while calling WinGetStyle: {}", ErrorText(error)))
        DeferWindowRetry(hwnd, 500)
        return
    }

    ; WS_CHILDは親に埋め込まれるタイプの子ウィンドウのため対象外で即帰る
    if (style & 0x40000000)  ; WS_CHILD
    {
        DebugLog("Window is a child; skipping.")
        DeferWindowRetry(hwnd, 2000)
        return
    }

    ; ウィンドウサイズを取得
    ; ウィンドウ出来立ての頃は1x1などのサイズでとれることがある。少し時間を置いて再試行する。
    rect := GetWindowRectInfo(hwnd)
    if !IsObject(rect)
    {
        if (attempt < MAX_MOVE_ATTEMPTS)
        {
            DebugLog("Window bounds unavailable; scheduling retry.")
            SetTimer(() => TryRelocateWindow(hwnd, attempt + 1), -120)
        }
        DeferWindowRetry(hwnd, 300)
        return
    }

    if (rect.w < MIN_DIALOG_DIMENSION || rect.h < MIN_DIALOG_DIMENSION)
    {
        DebugLog(Format("Window 0x{:X} is still too small ({}x{}); retrying.", hwnd, rect.w, rect.h))
        if (attempt < MAX_MOVE_ATTEMPTS)
            SetTimer(() => TryRelocateWindow(hwnd, attempt + 1), -120)
        DeferWindowRetry(hwnd, 250)
        return
    }

    ; CLIP STUDIO PAINTのメインウィンドウを探してサイズを確認。最小化中などで判定できない場合は後回し。
    main := GetMainWindowInfo()
    if !IsObject(main)
    {
        DebugLog("Failed to identify CSP main window.")
        DeferWindowRetry(hwnd, 500)
        return
    }
    if (main.w < MIN_MAIN_DIMENSION || main.h < MIN_MAIN_DIMENSION)
    {
        DebugLog("Main window does not meet minimum size requirements.")
        DeferWindowRetry(hwnd, 2000)
        return
    }

    ; サイズ比やタイトルバーの有無など、中央に寄せる対象として条件を満たしているかをチェック。
    ; メニュー(コンテキストメニュー)などはここで除外される。
    if !IsDialogCandidate(hwnd, main, rect, style)
    {
        DebugLog("Window rejected as dialog candidate.")
        DeferWindowRetry(hwnd, 2000)
        return
    }

    ; 実際に中心位置へ移動し、成功したら「次に触るまで最低5秒あける」よう記録して終了
    if MoveDialogToCenter(hwnd, main, rect)
    {
        DebugLog("Dialog moved successfully.")
        DeferWindowRetry(hwnd, 5000)
        return
    }

    ; 移動に失敗した場合は、（回数上限まで）80ms 後に再チャレンジを予約。併せて「400ms は再スキャン不要」と記憶する。
    if (attempt < MAX_MOVE_ATTEMPTS)
    {
        DebugLog("Move failed; scheduling retry.")
        SetTimer(() => TryRelocateWindow(hwnd, attempt + 1), -80)
    }
    DeferWindowRetry(hwnd, 400)
}

; ウィンドウhwndが、メインウィンドウの中央へ移動する対象となるかどうかを返す
; メインウィンドウである、タイトルバーがない(メニューのコンテキスト)、
IsDialogCandidate(hwnd, main, dialogRect := 0, style := "")
{
    global DIALOG_RATIO_LIMIT
    if (hwnd = main.hwnd)
    {
        DebugLog("Window is the main window; skipping.")
        return false
    }

    ; ウィンドウ属性（スタイル）が取得できなければ返る
    if (style = "")
    {
        try style := WinGetStyle("ahk_id " hwnd)
        catch error
        {
            DebugLog(Format("Failed to retrieve style for candidate: {}", ErrorText(error)))
            return false
        }
    }

    ; タイトルバーがついていないウィンドウは対象外（メニューバーのコンテキストメニューなど）
    if !(style & 0x00C00000) ; WS_CAPTION
    {
        DebugLog("Candidate lacks caption/title bar; excluded.")
        return false
    }

    if IsObject(dialogRect)
    {
        x := dialogRect.x
        y := dialogRect.y
        w := dialogRect.w
        h := dialogRect.h
    }
    else
    {
        try {
            WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
        } catch {
            DebugLog("Failed to get window position.")
            return false
        }
    }

    ; サイズが正常に取得できない場合は対象外
    if (w <= 0 || h <= 0)
    {
        DebugLog(Format("Invalid window dimensions. w={} h={}", w, h))
        return false
    }

    ; 面積計算できなければ（メインウィンドウも含め）対象外
    area := w * h
    if (area <= 0 || main.area <= 0)
    {
        DebugLog(Format("Invalid area calculation. area={} main.area={}", area, main.area))
        return false
    }

    ; ダイアログが最小化されていたら対象外
    try subState := WinGetMinMax("ahk_id " hwnd)
    catch error
    {
        DebugLog(Format("Failed to get dialog MinMax: {}", ErrorText(error)))
        return false
    }
    if (subState = -1)
    {
        DebugLog("Sub window is minimized; skipping.")
        return false
    }

    ; メインウィンドウと比較して大きければ対象外
    ratio := area / main.area
    if (ratio >= DIALOG_RATIO_LIMIT)
    {
        DebugLog(Format("Area ratio above limit. ratio={:.2f} limit={:.2f}", ratio, DIALOG_RATIO_LIMIT))
        return false
    }

    ; Exclude menu popups (class "#32768") unless they are anchored and large enough.
    ; クラス名を調べ、ポップアップメニューなど除外対象であれば帰る
    class := GetClassName(hwnd)
    if !IsWindowEligible(hwnd, class)
    {
        return false
    }

    DebugLog(Format("Accepting as dialog candidate. ratio={:.2f}", ratio))
    
    ; メインウィンドウの中心にもってくるダイアログ（対象）であると判定
    return true
}

MoveDialogToCenter(hwnd, main, dialogRect := 0)
{
    if IsObject(dialogRect)
    {
        dx := dialogRect.x
        dy := dialogRect.y
        dw := dialogRect.w
        dh := dialogRect.h
    }
    else
    {
        try {
            WinGetPos(&dx, &dy, &dw, &dh, "ahk_id " hwnd)
        } catch {
            DebugLog("Failed to get current dialog position.")
            return false
        }
    }
    targetX := Round(main.x + (main.w - dw) / 2)
    targetY := Round(main.y + (main.h - dh) / 2)
    DebugLog(Format("Moving window 0x{:X} to center. current=({}, {}) target=({}, {}) size={}x{}", hwnd, dx, dy, targetX, targetY, dw, dh))
    return AttemptWindowMove(hwnd, targetX, targetY)
}

AttemptWindowMove(hwnd, x, y)
{
    flags := 0x0001 | 0x0004 | 0x0010 | 0x0200 | 0x0400
    success := DllCall("SetWindowPos", "ptr", hwnd, "ptr", 0, "int", x, "int", y, "int", 0, "int", 0, "uint", flags)
    if !success
    {
        DebugLog(Format("SetWindowPos failed. err={}", A_LastError))
        return false
    }
    Sleep(1)
    try {
        WinGetPos(&curX, &curY, , , "ahk_id " hwnd)
    } catch {
        DebugLog("Failed to read position after move.")
        return false
    }
    if (Abs(curX - x) <= 2 && Abs(curY - y) <= 2)
    {
        DebugLog(Format("Window reached target position. current=({}, {})", curX, curY))
        return true
    }
    DebugLog(Format("Window did not reach target position. current=({}, {}) target=({}, {})", curX, curY, x, y))
    return false
}

GetMainWindowInfo(log := true)
{
    global MIN_MAIN_DIMENSION
    handles := WinGetList()
    maxArea := 0
    main := 0
    if log
    {
        DebugLog(Format("Searching for main window. candidates={}", handles.Length))
    }

    for hwnd in handles
    {
        if !IsCSPWindow(hwnd)
        {
            continue
        }

        try style := WinGetStyle("ahk_id " hwnd)
        catch error
        {
            if log
            {
                DebugLog(Format("Failed to get style. hwnd=0x{:X} err={}", hwnd, ErrorText(error)))
            }
            continue
        }

        if (style & 0x40000000)
        {
            if log
            {
                DebugLog(Format("Skipping child window. hwnd=0x{:X}", hwnd))
            }
            continue
        }

        try state := WinGetMinMax("ahk_id " hwnd)
        catch error
        {
            if log
            {
                DebugLog(Format("Failed to get MinMax. hwnd=0x{:X} err={}", hwnd, ErrorText(error)))
            }
            continue
        }

        if (state = -1)
        {
            if log
            {
                DebugLog(Format("Skipping minimized window. hwnd=0x{:X}", hwnd))
            }
            continue
        }

        try
        {
            WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
        } 
        catch
        {
            if log
            {
                DebugLog(Format("Failed to get window position. hwnd=0x{:X}", hwnd))
            }
            continue
        }

        if (w < MIN_MAIN_DIMENSION || h < MIN_MAIN_DIMENSION)
        {
            if log
                DebugLog(Format("Window too small. hwnd=0x{:X} size={}x{}", hwnd, w, h))
            continue
        }

        area := w * h
        if (area > maxArea)
        {
            maxArea := area
            main := { hwnd: hwnd, x: x, y: y, w: w, h: h, area: area, state: state }
            if log
                DebugLog(Format("Main window candidate updated. hwnd=0x{:X} area={} size={}x{}", hwnd, area, w, h))
        }
    }

    if IsObject(main)
    {
        if log
        {
            DebugLog(Format("Main window selected. hwnd=0x{:X} size={}x{}", main.hwnd, main.w, main.h))
        }
    }
    else if log
    {
        DebugLog("No main window found.")
    }
    
    return main
}

IsCSPWindow(hwnd, &processName := "")
{
    global TARGET_PROCESS
    if !IsWindow(hwnd)
        return false
    processName := GetProcessNameByHwnd(hwnd)
    if (processName = "")
        return false
    result := (StrLower(processName) = TARGET_PROCESS)
    if result
        DebugLog(Format("CSP window confirmed. hwnd=0x{:X}", hwnd))
    return result
}

GetLayeredAlpha(hwnd)
{
    alpha := 255
    has := 0
    ; BOOL GetLayeredWindowAttributes(HWND, COLORREF*, BYTE*, DWORD*)
    if DllCall("GetLayeredWindowAttributes"
        , "ptr", hwnd, "uint*", 0, "uchar*", &alpha, "uint*", &has) {
        ; alpha は 0..255。取得できなければ 255 のまま
    }
    return alpha
}

IsMainCspWindow(hwnd)
{
    ; ウィンドウが存在しなければ対象外
    ;if !WinExist("ahk_id " hwnd)
    ;{
    ;    return false
    ;}

    ; ウィンドウの存在とプロセス名の一致を確認する
    if !IsCSPWindow(hwnd)
    {
        return false
    }

    ; タイトルバーの文字列に「CLIP STUDIO PAINT」が含まれるかどうか調べる
    title := WinGetTitle("ahk_id " hwnd)
    ; 例: ローカライズを考慮しつつ基本は含有チェック
    if !(title ~= "i)CLIP\s*STUDIO\s*PAINT")
    {
        return false
    }

    ; トップレベル & 非最小化
    mm := WinGetMinMax("ahk_id " hwnd)  ; -1:min, 0:normal, 1:max
    if (mm = -1)
    {
        return false
    }
    owner := DllCall("GetWindow", "ptr", hwnd, "uint", 4, "ptr") ; GW_OWNER=4
    if (owner)
    {
        return false
    }

    ; 標準枠を持ち、ツールウィンドウでない
    style  := WinGetStyle("ahk_id " hwnd)
    exstyle := WinGetExStyle("ahk_id " hwnd)
    ; WS_CAPTION=0x00C00000, WS_EX_TOOLWINDOW=0x00000080
    if !(style & 0x00C00000)
    {
        return false
    }
    if (exstyle & 0x00000080)  ; TOOLWINDOW は除外
    {
        return false
    }

    ; 透明オーバーレイの除外
    ; WS_EX_LAYERED=0x00080000, WS_EX_TRANSPARENT=0x00000020
    if (exstyle & 0x00080000) 
    {
        alpha := GetLayeredAlpha(hwnd)  ; 下の補助関数参照（なければ 255 扱い）
        if ((exstyle & 0x20) || alpha < 250) ; ほぼ不透明以外を除外（閾値は調整可）
        {
            return false
        }
    }

    ; サイズの下限と全面オーバーレイ除外
    WinGetPos(&x,&y,&w,&h, "ahk_id " hwnd)
    if (w < 400 || h < 300)
    {
        return false
    }

    ; 仮想スクリーン全面の 90% 超を除外（透明全画面オーバーレイ対策）
    vL := SysGet(76), vT := SysGet(77), vW := SysGet(78), vH := SysGet(79) ; SM_XVIRTUALSCREEN...
    if (w*h >= 0.90 * vW * vH)
    {
        ; ただし最大化中の正規メインを誤排除しないよう、キャプション・不透明・非TRANSPARENTを満たすか再確認
        if (exstyle & 0x20)  ; TRANSPARENT が立っていたら確実に除外
        {
            return false
        }
        
        if (exstyle & 0x00080000) 
        {
            alpha := GetLayeredAlpha(hwnd)
            if (alpha < 250)
            {
                return false
            }
        }

        ; 上の条件をすべてクリアしていれば「本物の最大化メイン」とみなして通過
    }

    return true

    ; 旧処理いったんやめ。
    ; main := GetMainWindowInfo(DEBUG_MODE)
    ; return (IsObject(main) && main.hwnd = hwnd)
}

IsWindow(hwnd)
{
    return (hwnd && DllCall("IsWindow", "ptr", hwnd, "int"))
}

GetClassName(hwnd)
{
    try return WinGetClass("ahk_id " hwnd)
    catch error
    {
        DebugLog(Format("WinGetClass failed for hwnd=0x{:X}. err={}", hwnd, ErrorText(error)))
        return ""
    }
}

GetWindowRectInfo(hwnd)
{
    try {
        WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
        if (w <= 1 || h <= 1)
        {
            ext := GetWindowRectFromDwm(hwnd)
            if IsObject(ext)
            {
                DebugLog(Format("Using DWM bounds for hwnd=0x{:X}: size={}x{}", hwnd, ext.w, ext.h))
                return ext
            }
        }
        return { x: x, y: y, w: w, h: h }
    } catch error {
        DebugLog(Format("WinGetPos failed for hwnd=0x{:X}. err={}", hwnd, ErrorText(error)))
        return 0
    }
}

GetWindowRectFromDwm(hwnd)
{
    static DWMWA_EXTENDED_FRAME_BOUNDS := 9
    rectBuf := Buffer(16, 0)
    try result := DllCall("dwmapi\DwmGetWindowAttribute", "ptr", hwnd, "int", DWMWA_EXTENDED_FRAME_BOUNDS, "ptr", rectBuf, "int", rectBuf.Size, "int")
    catch error
    {
        DebugLog(Format("DwmGetWindowAttribute failed for hwnd=0x{:X}. err={}", hwnd, ErrorText(error)))
        return 0
    }
    if (result != 0)
    {
        DebugLog(Format("DwmGetWindowAttribute returned hr=0x{:X} for hwnd=0x{:X}", result, hwnd))
        return 0
    }
    left := NumGet(rectBuf, 0, "int")
    top := NumGet(rectBuf, 4, "int")
    right := NumGet(rectBuf, 8, "int")
    bottom := NumGet(rectBuf, 12, "int")
    w := right - left
    h := bottom - top
    if (w <= 0 || h <= 0)
    {
        DebugLog(Format("DWM bounds invalid for hwnd=0x{:X}: {}x{}", hwnd, w, h))
        return 0
    }
    return { x: left, y: top, w: w, h: h }
}

GetProcessNameByHwnd(hwnd)
{
    try pid := WinGetPID("ahk_id " hwnd)
    catch error
    {
        DebugLog(Format("WinGetPID failed for hwnd=0x{:X}. err={}", hwnd, ErrorText(error)))
        return ""
    }
    if (pid <= 0)
    {
        DebugLog(Format("WinGetPID returned invalid PID ({}) for hwnd=0x{:X}", pid, hwnd))
        return ""
    }
    try name := ProcessGetName(pid)
    catch error
    {
        DebugLog(Format("ProcessGetName failed for pid={} hwnd=0x{:X}. err={}", pid, hwnd, ErrorText(error)))
        name := ""
    }
    if (name = "")
    {
        try path := ProcessGetPath(pid)
        catch error
        {
            DebugLog(Format("ProcessGetPath failed for pid={} hwnd=0x{:X}. err={}", pid, hwnd, ErrorText(error)))
            path := ""
        }
        if (path != "")
        {
            SplitPath path, &base
            name := base
        }
    }
    return name
}

ErrorText(value)
{
    if !IsObject(value)
        return value
    try {
        if HasProp(value, "Message")
            return value.Message
    }
    catch
    {
    }
    for name in ["What", "Description", "Value", "Reason", "Extra"]
    {
        try {
            if HasProp(value, name)
                return value.%name%
        } catch {
        }
    }
    try return value.ToString()
    catch
        return Type(value)
}

ShellEventName(code)
{
    switch code
    {
        case 1:
            return "HSHELL_WINDOWCREATED"
        case 2:
            return "HSHELL_WINDOWDESTROYED"
        case 4:
            return "HSHELL_WINDOWACTIVATED"
        case 5:
            return "HSHELL_GETMINRECT"
        case 6:
            return "HSHELL_REDRAW"
        case 0x8004:
            return "HSHELL_RUDEAPPACTIVATED"
        default:
            return Format("UNKNOWN(0x{:X})", code)
    }
}

DebugLog(message)
{
    global DEBUG_MODE, LOG_FILE
    if !DEBUG_MODE
        return
    try time := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    catch error
    {
        time := A_Now
    }
    ms := A_MSec + 0
    line := Format("[{1}.{2:03}] {3}`n", time, ms, message)
    try FileAppend(line, LOG_FILE, "UTF-8")
    catch error
    {
        ; Ignore logging failures
    }
}

ReleaseResources(*)
{
    global gShellHookGui, gShellHookMsg, gMutex, gScannerStarted
    DebugLog("Beginning resource cleanup.")
    if gShellHookMsg
    {
        DebugLog("Removing shell hook message handler.")
        OnMessage(gShellHookMsg, ShellMessage, 0)
    }
    if IsObject(gShellHookGui)
    {
        DebugLog("Unregistering shell hook window.")
        DllCall("DeregisterShellHookWindow", "ptr", gShellHookGui.Hwnd)
        gShellHookGui.Destroy()
    }
    if gMutex
    {
        DebugLog("Releasing mutex.")
        DllCall("ReleaseMutex", "ptr", gMutex)
        DllCall("CloseHandle", "ptr", gMutex)
        gMutex := 0
    }
    if gScannerStarted
    {
        SetTimer(ScanForDialogs, 0)
        gScannerStarted := false
        DebugLog("Dialog watchdog timer stopped.")
    }
    DebugLog("Resource cleanup finished.")
}

DeferWindowRetry(hwnd, delay := 1000)
{
    global gKnownWindows
    if delay <= 0
        delay := 100
    gKnownWindows[hwnd] := A_TickCount + delay
}

StartDialogScanner()
{
    global gScannerStarted
    if gScannerStarted
        return
    DebugLog("Starting dialog watchdog timer.")
    SetTimer(ScanForDialogs, 500)
    gScannerStarted := true
}

ScanForDialogs()
{
    global TARGET_PROCESS, gKnownWindows, gPendingRelocate
    handles := []
    try handles := WinGetList("ahk_exe " TARGET_PROCESS)
    catch error
    {
        DebugLog(Format("WinGetList with filter failed: {}", ErrorText(error)))
    }
    if (handles.Length = 0)
    {
        try handles := WinGetList()
        catch error
        {
            DebugLog(Format("WinGetList without filter failed: {}", ErrorText(error)))
            return
        }
    }
    now := A_TickCount
    for hwnd in handles
    {
        if !IsWindow(hwnd)
            continue
        if IsMainCspWindow(hwnd)
        {
            DebugLog(Format("Watchdog detected main window (0x{:X}); skipping.", hwnd))
            continue
        }
        if gPendingRelocate.Has(hwnd)
            continue
        if gKnownWindows.Has(hwnd) && (now < gKnownWindows[hwnd])
            continue
        if !DllCall("IsWindowVisible", "ptr", hwnd, "int")
            continue
        if !IsCSPWindow(hwnd)
            continue
        gKnownWindows[hwnd] := now + 500
        DebugLog(Format("Watchdog detected candidate hwnd=0x{:X}", hwnd))
        ScheduleRelocate(hwnd)
    }
    ; prune defunct entries periodically
    for hwnd, ts in gKnownWindows.Clone()
    {
        if !IsWindow(hwnd)
            gKnownWindows.Delete(hwnd)
    }
}

IsWindowEligible(hwnd, class := "")
{
    global TARGET_PROCESS, MIN_DIALOG_DIMENSION
    if (class = "")
        class := GetClassName(hwnd)
    if (class = "#32768")
    {
        DebugLog("Menu popup (#32768) excluded from relocation.")
        return false
    }
    return true
}
