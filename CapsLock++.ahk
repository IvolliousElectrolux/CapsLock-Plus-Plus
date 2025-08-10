#Requires AutoHotkey v2.0
#SingleInstance Force

; =====================================================================
; CapsLock++ v1.01
; 增强版 CapsLock 功能脚本
; =====================================================================

; =====================================================================
; 初始化
; =====================================================================

; 初始化全局变量
InitializeGlobalVariables() {
    global showDebugTooltips := false
    global useTaskbarOrder := false
    global includeMultipleInstances := true
    
    ; 虚拟环境相关
    global virtualEnvEnabled := false
    global virtualEnvWindows := []
    
    ; 首次运行标志
    global isFirstRun := true

    ; 跳转模式全局变量 - 提前声明确保在任何引用之前初始化
    global jumpMode := ""           ; 存储当前跳转模式
    global jumpActive := false      ; 是否激活跳转模式
    global jumpBuffer := ""         ; 存储输入的数字
    global jumpPosition := {x: 0, y: 0}  ; 存储位置
    global g_inputHook := {}        ; 存储输入钩子对象
    global isWordJump := false      ; 标识是否是单词级跳转
}

; 设置自定义任务托盘图标
SetCustomTrayIcon() {
    ; 图标文件路径
    iconPath := A_ScriptDir . "\Icon\CapsLock++.ico"
    
    ; 检查图标文件是否存在
    if (FileExist(iconPath)) {
        ; 设置任务托盘图标和提示文本
        TraySetIcon(iconPath)
        A_IconTip := "CapsLock++"
    } else {
        ; 如果图标文件不存在，显示提示
        ToolTip("未找到自定义图标文件: " . iconPath)
        SetTimer () => ToolTip(), -3000
    }
}

; 自定义任务托盘菜单
CustomizeTrayMenu() {
    ; 获取当前任务托盘菜单对象
    trayMenu := A_TrayMenu
    
    ; 清除默认菜单项
    trayMenu.Delete()
    
    ; 添加临时禁用选项
    trayMenu.Add("临时禁用", DisableCapsLockPlusPlus)
    
    ; 添加分隔线
    trayMenu.Add()
    
    ; 添加退出选项
    trayMenu.Add("退出", (*) => ExitApp())
}

; 临时禁用CapsLock++的函数
DisableCapsLockPlusPlus(*) {
    ; 检查文件是否存在并决定使用哪个扩展名
    capsLockNative := A_ScriptDir "\CapsLockNative.ahk"
    if (!FileExist(capsLockNative))
        capsLockNative := A_ScriptDir "\CapsLockNative.exe"
    
    ; 启动原生CapsLock脚本
    Try {
        Run(capsLockNative)
        ExitApp  ; 切换后退出当前脚本
    } Catch as err {
        ShowTooltip("无法启动原生CapsLock模式: " err.Message)
    }
}

; 主程序初始化
InitializeApp() {
    ; 初始化全局变量
    InitializeGlobalVariables()
    
    ; 确保CapsLock始终处于关闭状态，防止状态指示器闪烁
    SetCapsLockState("AlwaysOff")
    
    ; 创建定时器，定期检查并确保CapsLock处于关闭状态
    SetTimer(CheckCapsLockState, 2000)
    
    ; 创建定时器，定期检查虚拟环境窗口是否存在
    SetTimer(CleanupVirtualEnvWindows, 5000)  ; 每5秒检查一次
    
    ; 设置自定义任务托盘图标
    SetCustomTrayIcon()
    
    ; 自定义任务托盘菜单
    CustomizeTrayMenu()

    ; 开始时显示欢迎提示
    ToolTip("任务栏应用切换工具已启动`n使用Alt+鼠标滚轮切换窗口")
    SetTimer () => ToolTip(), -3000
}

; 检查并维持CapsLock状态
CheckCapsLockState() {
    global capsLockManuallyEnabled
    
    ; 如果CapsLock已关闭，不需要操作
    if !GetKeyState("CapsLock", "T")
        return
    
    ; 如果CapsLock开启，但不是手动开启的，则强制关闭
    if !capsLockManuallyEnabled {
        SetCapsLockState("AlwaysOff")
    }
}

; 启动应用
InitializeApp() 


; 请求管理员权限
if (!A_IsAdmin) {
    Run("*RunAs " A_ScriptFullPath)
    ExitApp
}

; 临时禁用/启用CapsLock++热键 - 改为切换到原生CapsLock脚本
#HotIf GetKeyState("CapsLock", "P")
Escape::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    ; 获取当前菜单窗口信息
    global currentMenuGui
    
    ; 如果当前菜单窗口存在，则关闭菜单
    if (currentMenuGui && WinExist("ahk_id " currentMenuGui)) {
        CloseMenu()
    }
    
    ; 检查文件是否存在并决定使用哪个扩展名
    capsLockNative := A_ScriptDir "\CapsLockNative.ahk"
    if (!FileExist(capsLockNative))
        capsLockNative := A_ScriptDir "\CapsLockNative.exe"
    
    ; 启动原生CapsLock脚本
    Try {
        Run(capsLockNative)
    } Catch as err {
        ShowTooltip("无法启动原生CapsLock模式: " err.Message)
        Return
    }
    
    ; 退出当前脚本
    ExitApp
}
#HotIf

; 任务栏应用切换
; 功能：
;   1. 任务栏窗口切换（两种排序方式：优先级与任务栏顺序）
;   2. 虚拟环境窗口管理（添加/删除/查看/清空）
;   3. 支持同一进程多窗口（如多个记事本、多个浏览器窗口）
;   4. 监视器静默关闭特定窗口
;   5. 窗口置顶功能 (CapsLock+右键)
;   6. 工作区清理功能

; =====================================================================
; 全局设置
; =====================================================================

; 选择窗口排序方式
global useTaskbarOrder := false  ; false = 优先级排序（默认），true = 任务栏顺序排序
; 是否显示调试信息
global showDebugTooltips := false  ; 开启调试信息显示

; 提示信息显示时间（毫秒）
global tipDuration := 2000        ; 普通提示
global debugTipDuration := 3000   ; 调试提示
global longTipDuration := 5000    ; 长提示（如窗口列表）

; CapsLock 状态监控相关变量 
global capsLockManuallyEnabled := false            ; CapsLock是否被手动启用
global capsLockKeyPressed := false                 ; 记录CapsLock是否被按下
global capsLockPressTime := 0                      ; 记录CapsLock按下时间
global capsLockReleaseTime := 0                    ; 记录CapsLock释放时间
global otherKeyPressed := false                    ; 记录是否按下了其他键

; 显示临时提示
ShowTooltip(text, duration := 2000) {
    ToolTip(text)
    SetTimer () => ToolTip(), -duration
}

; 确保CapsLock始终关闭
SetCapsLockState("AlwaysOff")

; =====================================================================
; 调试功能
; =====================================================================
; Ctrl+Win+I: 切换调试信息显示
^!i::ToggleDebugTooltips()

; =====================================================================
; 黑名单设置
; =====================================================================
; 黑名单程序路径，这些程序不会出现在切换列表中
; 黑名单程序路径读取 - 使用循环读取所有条目
blacklist := []
i := 1
Loop {
    key := "black" . i
    value := IniRead("CapsLock++.ini", "blacklist_virtual_env", key, "")
    if (value = "")
        break
    blacklist.Push(value)
    i++
}

; 黑名单窗口类名读取 - 使用循环读取所有条目
blacklistClasses := []
i := 1
Loop {
    key := "blackclasses" . i
    value := IniRead("CapsLock++.ini", "blacklist_classes_virtual_env", key, "")
    if (value = "")
        break
    blacklistClasses.Push(value)
    i++
}

; 虚拟环境相关变量
virtualEnvEnabled := false      ; 是否启用虚拟环境
virtualEnvWindows := []         ; 存储虚拟环境中的窗口信息（包含hwnd, title, processName等）

; 为了提高效率，预先创建黑名单进程名称的缓存
blacklistProcessNames := []
for path in blacklist {
    SplitPath(path, &processName)
    blacklistProcessNames.Push(processName)
}

; =====================================================================
; 热键列表
; =====================================================================
; CapsLock键行为重定义
; --------------------------------------------------------------------
; 在脚本启动时立即将CapsLock键功能禁用
; 确保始终关闭
SetCapsLockState("AlwaysOff")

; 格式化数字，保留2位小数
FormatNumber(num) {
    return Round(num, 2)
}

; =============== CapsLock处理 - 纯热键方式 ===============

; 全局变量
global capsLockPressTime := 0
global otherKeyPressed := false
global capsLockIsDown := false  ; 新增：跟踪CapsLock是否已经处于按下状态

; CapsLock按下时的处理
~CapsLock::
{
    ; 声明所有需要的全局变量
    global capsLockIsDown, capsLockPressTime, otherKeyPressed, capsLockManuallyEnabled, showDebugTooltips
    
    ; 只在首次按下时记录时间
    if (!capsLockIsDown) {
        capsLockIsDown := true
        capsLockPressTime := A_TickCount
        otherKeyPressed := false
        
        if (showDebugTooltips) {
            ToolTip("CapsLock按下，开始计时: " capsLockPressTime)
            SetTimer () => ToolTip(), -1000
        }
    }
    
    ; 确保CapsLock保持关闭状态（除非手动启用）
    if (!capsLockManuallyEnabled)
        SetCapsLockState("AlwaysOff")
}

; CapsLock释放时的处理
~CapsLock Up::
{
    ; 声明所有需要的全局变量
    global capsLockIsDown, capsLockPressTime, otherKeyPressed, capsLockManuallyEnabled, showDebugTooltips
    
    ; 只有当我们记录了按下状态时才处理释放
    if (capsLockIsDown) {
        ; 计算持续时间
        pressDuration := A_TickCount - capsLockPressTime
        
        ; 重置按下状态
        capsLockIsDown := false
        
        ; 确保CapsLock保持在预期状态
        if (!capsLockManuallyEnabled)
            SetCapsLockState("AlwaysOff")
        
        ; 如果没有按下其他键且持续时间小于阈值，发送Esc
        if (!otherKeyPressed && pressDuration < 300) {
            SendInput("{Esc}")
            
            if (showDebugTooltips) {
                ToolTip("CapsLock单击，发送Esc (" pressDuration "ms)")
                SetTimer () => ToolTip(), -1000
            }
        } else if (showDebugTooltips) {
            ToolTip("CapsLock释放，" 
                  . (otherKeyPressed ? "按下了其他键" : "超过阈值") 
                  . " (" pressDuration "ms)")
            SetTimer () => ToolTip(), -1000
        }
        
        ; 重置其他状态
        otherKeyPressed := false
    }
}

; 手动切换CapsLock状态 (Ctrl+CapsLock)
^CapsLock::
{
    global capsLockManuallyEnabled, otherKeyPressed
    
    ; 标记为按下了其他键，避免触发Esc
    otherKeyPressed := true
    
    ; 切换CapsLock状态
    if GetKeyState("CapsLock", "T") {
        SetCapsLockState("AlwaysOff")
        capsLockManuallyEnabled := false
        state := "关闭"
    } else {
        SetCapsLockState("AlwaysOn")
        capsLockManuallyEnabled := true
        state := "开启"
    }
    
    ; 显示当前CapsLock状态
    ToolTip("大写锁定: " state)
    SetTimer () => ToolTip(), -1000
}

; 在脚本退出时确保CapsLock设置还原
OnExit((*) => SetCapsLockState("AlwaysOff"))

; 窗口切换
; --------------------------------------------------------------------
; 基本窗口切换功能 - CapsLock 替代 Alt
#HotIf GetKeyState("CapsLock", "P")
WheelDown::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    ; 检查ALT键状态，实现CapsLock+Alt+WheelDown的效果
    if GetKeyState("LAlt", "P") {
        ; Alt按下时调整鼠标速度
        global mouseSpeedValue
        mouseSpeedValue := Max(mouseSpeedValue - 1, 1)
        ShowTooltip("鼠标速度已调整为: " mouseSpeedValue)
    } else {
        ; 否则执行正常的窗口切换功能
        SwitchTaskbarWindow(1)  ; 正向切换窗口
        
        ; 显示调试信息
        if (showDebugTooltips) {
            ShowTooltip("CapsLock+WheelDown: 正向切换窗口")
        }
    }
}

WheelUp::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    ; 检查ALT键状态，实现CapsLock+Alt+WheelUp的效果
    if GetKeyState("LAlt", "P") {
        ; Alt按下时调整鼠标速度
        global mouseSpeedValue
        mouseSpeedValue := Min(mouseSpeedValue + 1, 20)
        ShowTooltip("鼠标速度已调整为: " mouseSpeedValue)
    } else {
        ; 否则执行正常的窗口切换功能
        SwitchTaskbarWindow(-1)  ; 反向切换窗口
        
        ; 显示调试信息
        if (showDebugTooltips) {
            ShowTooltip("CapsLock+WheelUp: 反向切换窗口")
        }
    }
}
#HotIf

XButton1::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    SwitchTaskbarWindow(1)  ; 正向切换窗口
        
        ; 显示调试信息
    if (showDebugTooltips) {
        ShowTooltip("XButton1: 正向切换窗口")
    }
}

XButton2::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    SwitchTaskbarWindow(-1)  ; 反向切换窗口
        
        ; 显示调试信息
    if (showDebugTooltips) {
        ShowTooltip("XButton2: 反向切换窗口")
    }
}

; 其他窗口切换热键
!Escape::SwitchTaskbarWindow(1)                       ; Ctrl+Alt+Esc：正序切换
!+Escape::SwitchTaskbarWindow(-1)                     ; Ctrl+Alt+Shift+Esc：逆序切换

; =====================================================================
; 标签页切换
; =====================================================================
#HotIf GetKeyState("RButton", "P")
WheelDown::
{
    global otherKeyPressed := true
    
    SendInput("^{Tab}")
}

WheelUp::
{
    global otherKeyPressed := true
    
    SendInput("^+{Tab}")
}
#HotIf

; =====================================================================
; 虚拟环境管理
; =====================================================================
; CapsLock+中键: 智能添加/移除当前窗口到虚拟环境(与Alt+中键相同)

#HotIf GetKeyState("CapsLock", "P")
MButton::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    ; 检查Alt键是否被按下
    if (GetKeyState("Alt", "P")) {
        ; Alt被按下，清空虚拟环境
        ClearVirtualEnv()
    } else {
        ; Alt未被按下，执行智能添加/移除功能
        SmartVirtualEnvToggle()
    }
}

; CapsLock+右键: 置顶/取消置顶光标所在窗口
RButton::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    ToggleWindowPinned()
}

; CapsLock+左键: 先选中文件/文件夹，然后重命名
LButton::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    local IsRename := false
    
    ; 检查当前激活窗口和鼠标下的窗口
    try {
        activeWin := WinGetID("A")
        activeWinClass := WinGetClass("ahk_id " activeWin)
    } catch {
        ; 如果无法获取活动窗口，设置为空值
        activeWin := 0
        activeWinClass := ""
    }
    
    ; 获取鼠标下的窗口
    try {
        MouseGetPos(, , &mouseWin)
        mouseWinClass := WinGetClass("ahk_id " mouseWin)
    } catch {
        mouseWin := 0
        mouseWinClass := ""
    }
    
    ; 检查是否需要先激活窗口（针对ExplorerBrowserOwner的特殊处理）
    needActivate := IsExplorerBrowserOwnerCase(activeWinClass, mouseWinClass)
    
    ; 如果需要先激活窗口，先激活再点击
    if (needActivate) {
        WinActivate("ahk_id " mouseWin)
        ; 使用定时器而不是Sleep阻塞
        SetTimer(PerformClick, -40)
        return
    }
    
    ; 执行普通的左键点击
    Click("Left")
    
    SetTimer (LButtonRenamer), -20
}
#HotIf

; 特殊情况下延迟点击的函数
PerformClick() {
    Click("Left")
    SetTimer(LButtonRenamer, -20)
}

; 特殊情况判断：ExplorerBrowserOwner相关窗口的处理
IsExplorerBrowserOwnerCase(activeWinClass, mouseWinClass) {
    ; 如果有一个窗口是ExplorerBrowserOwner，而另一个不是，需要先激活
    if ((activeWinClass = "ExplorerBrowserOwner" && mouseWinClass != "ExplorerBrowserOwner") ||
        (activeWinClass != "ExplorerBrowserOwner" && mouseWinClass = "ExplorerBrowserOwner")) {
        return true
    }
    return false
}

LButtonRenamer(){
    ; 检查光标下的窗口是否为文件资源管理器
    MouseGetPos(, , &mouseWin)
    if (mouseWin) {
        activeClass := WinGetClass("ahk_id " mouseWin)
    } else {
        activeClass := ""
    }
    
    ; 如果是文件资源管理器相关窗口，则发送F2进行重命名
    if (activeClass = "CabinetWClass" || activeClass = "ExploreWClass" || 
        activeClass = "Progman" || activeClass = "WorkerW" || activeClass = "ExplorerBrowserOwner") {
        Send("{F2}")
    }
}

; 工作区清理功能
; --------------------------------------------------------------------
; Ctrl+Win+Z: 清理工作区 (最小化非虚拟环境窗口或非当前窗口)
; 注意：由于AHK2.0的热键处理方式，确保这里没有注释混淆

; 全局变量，用于存储最小化的窗口信息
global minimizedWindows := []         ; 存储被最小化的窗口信息
global lastWorkspaceCleanupTime := 0  ; 上次清理工作区的时间
global workspaceCleanupMode := "minimize"  ; 当前模式：minimize或restore

; 工作区清理功能 - 确保热键定义正确且唯一
#HotIf GetKeyState("CapsLock", "P")
#z::CleanupWorkspaceHotkey()
#HotIf

; ; 使用Hotkey()函数显式注册热键作为备选方案
; try {
;     Hotkey("^#z", (*) => CleanupWorkspaceHotkey(), "On")
; } catch as e {
;     ToolTip("注册工作区清理热键失败: " e.Message)
;     SetTimer () => ToolTip(), -3000
; }

; 清理工作区热键处理函数
; 功能：
;   1. 当虚拟环境启用时：最小化除虚拟环境窗口外的所有窗口
;   2. 当虚拟环境未启用时：最小化除光标下窗口外的所有窗口
;   3. 再次按下时，恢复之前最小化的窗口（如果窗口状态未被手动改变）
;   4. 黑名单窗口和特殊UI元素不会被处理
CleanupWorkspaceHotkey() {
    global minimizedWindows, lastWorkspaceCleanupTime, workspaceCleanupMode
    
    ; 显示触发提示以确认热键正常工作
    ToolTip("工作区清理功能已触发")
    SetTimer () => ToolTip(), -1000
    
    ; 获取当前时间
    currentTime := A_TickCount
    
    ; 如果有最小化的窗口记录，尝试恢复它们
    if (minimizedWindows.Length > 0) {
        ; 简单检查：至少有一个窗口已经不是最小化状态了
        windowsChanged := false
        validWindowCount := 0
        
        for _, winInfo in minimizedWindows {
            ; 检查窗口是否存在
            if (WinExist("ahk_id " winInfo.hwnd)) {
                validWindowCount++
                ; 检查窗口是否仍然是最小化状态
                currentMinState := WinGetMinMax("ahk_id " winInfo.hwnd) = -1
                ; 如果窗口状态与记录不符，标记为已变化
                if (currentMinState != true) {  ; 应该是最小化的
                    windowsChanged := true
                    break
                }
            }
        }
        
        ; 如果没有有效窗口或窗口状态已变化，清空记录并执行最小化
        if (validWindowCount = 0 || windowsChanged) {
            ToolTip("窗口状态已变化或无有效窗口，`n执行最小化操作")
            SetTimer () => ToolTip(), -2000
            minimizedWindows := []
            MinimizeWorkspaceWindows()
        } else {
            ; 否则恢复窗口
            RestoreMinimizedWindows()
        }
    } else {
        ; 没有最小化的窗口记录，执行最小化操作
        MinimizeWorkspaceWindows()
    }
    
    ; 更新上次清理时间
    lastWorkspaceCleanupTime := currentTime
}

; 检查窗口状态是否发生变化
CheckWindowStateChanged() {
    global minimizedWindows, lastWorkspaceCleanupTime
    
    ; 如果没有最小化的窗口记录，返回false
    if (minimizedWindows.Length = 0)
        return false
    
    ; 检查每个记录的窗口状态是否与记录时一致
    for i, winInfo in minimizedWindows {
        ; 检查窗口是否存在
        if (!WinExist("ahk_id " winInfo.hwnd))
            return true
        
        ; 检查窗口最小化状态是否与记录时一致
        currentMinState := WinGetMinMax("ahk_id " winInfo.hwnd) = -1
        if (currentMinState != winInfo.wasMinimized)
            return true
    }
    
    return false
}

; 最小化工作区窗口
MinimizeWorkspaceWindows() {
    global minimizedWindows, virtualEnvEnabled, virtualEnvWindows
    
    ; 清空之前的记录
    minimizedWindows := []
    
    ; 获取光标下的窗口
    cursorHwnd := GetWindowUnderCursor()
    
    ; 如果光标在任务栏上，不执行最小化操作
    if (cursorHwnd = -1) {
        ToolTip("光标在任务栏上，不执行最小化操作")
        SetTimer () => ToolTip(), -2000
        return
    }
    
    ; 如果光标下没有有效窗口，尝试获取当前活动窗口
    if (!cursorHwnd || !WinExist("ahk_id " cursorHwnd)) {
        try {
            cursorHwnd := WinGetID("A")
        } catch {
            cursorHwnd := 0
        }
    }
    
    ; 如果虚拟环境已启用且有窗口，则最小化非虚拟环境窗口
    if (virtualEnvEnabled && virtualEnvWindows.Length > 0) {
        MinimizeNonVirtualEnvWindows(cursorHwnd)
    } 
    ; 否则，最小化除了光标下窗口之外的所有窗口
    else if (cursorHwnd && WinExist("ahk_id " cursorHwnd)) {
        MinimizeCursorWindowOrOthers(cursorHwnd)
    } 
    ; 如果没有有效窗口，提示用户
    else {
        ToolTip("无法获取有效窗口，清理操作取消")
        SetTimer () => ToolTip(), -2000
    }
}

; 最小化非虚拟环境窗口
MinimizeNonVirtualEnvWindows(excludeHwnd := 0) {
    global minimizedWindows, virtualEnvWindows, blacklistProcessNames, blacklistClasses, pinnedWindows
    
    ; 获取所有任务栏窗口
    winList := WinGetList(,, "Program Manager")
    minimizedCount := 0
    
    ; 创建虚拟环境窗口句柄的哈希表，用于快速查找
    virtualEnvHwnds := Map()
    for _, win in virtualEnvWindows {
        virtualEnvHwnds[win.hwnd] := true
    }
    
    ; 创建置顶窗口句柄的哈希表，用于快速查找
    pinnedHwndsMap := Map()
    for _, pinnedHwnd in pinnedWindows {
        pinnedHwndsMap[pinnedHwnd] := true
    }
    
    ; 遍历所有窗口
    for hwnd in winList {
        ; 跳过不是任务栏窗口的窗口
        if (!IsTaskbarWindow(hwnd))
            continue
            
        ; 跳过排除的窗口
        if (hwnd = excludeHwnd)
            continue
            
        ; 检查窗口是否在黑名单中
        className := WinGetClass("ahk_id " hwnd)
        
        ; 获取窗口进程信息
        pid := WinGetPID("ahk_id " hwnd)
        processPath := ProcessGetPath(pid)
        SplitPath(processPath, &processName)
        
        ; 检查进程名是否在黑名单中
        if (HasVal(blacklistProcessNames, processName))
            continue
            
        ; 检查类名是否在黑名单中
        isBlacklistedClass := false
        for blackClass in blacklistClasses {
            if (InStr(className, blackClass) || className = blackClass) {
                isBlacklistedClass := true
                break
            }
        }
        
        if (isBlacklistedClass)
            continue
            
        ; 检查窗口是否置顶 - 跳过置顶窗口
        exStyle := WinGetExStyle("ahk_id " hwnd)
        isPinned := (exStyle & 0x8) != 0  ; WS_EX_TOPMOST
        
        ; 如果窗口在置顶列表中或者有置顶标志，跳过
        if (isPinned || pinnedHwndsMap.Has(hwnd))
            continue
            
        ; 如果窗口不在虚拟环境中，则最小化它
        if (!virtualEnvHwnds.Has(hwnd)) {
            ; 记录窗口当前状态
            wasMinimized := WinGetMinMax("ahk_id " hwnd) = -1
            
            ; 如果窗口未最小化，则最小化它
            if (!wasMinimized) {
                ; 获取窗口信息用于显示
                title := WinGetTitle("ahk_id " hwnd)
                simpleName := RegExReplace(processName, "\.exe$", "")
                
                ; 最小化窗口
                WinMinimize("ahk_id " hwnd)
                
                ; 记录被最小化的窗口信息
                minimizedWindows.Push({
                    hwnd: hwnd,
                    title: title,
                    processName: processName,
                    simpleName: simpleName,
                    wasMinimized: wasMinimized
                })
                
                minimizedCount++
            }
        }
    }
    
    ; 显示操作结果
    if (minimizedCount > 0) {
        ToolTip("已最小化 " minimizedCount " 个非虚拟环境窗口`n再次按Ctrl+Win+Z恢复")
    } else {
        ToolTip("没有需要最小化的非虚拟环境窗口")
    }
    
    SetTimer () => ToolTip(), -2000
}

; 最小化光标下的窗口或所有其他窗口
MinimizeCursorWindowOrOthers(cursorHwnd) {
    global minimizedWindows, blacklistProcessNames, blacklistClasses, pinnedWindows
    
    ; 获取光标下窗口的信息
    title := WinGetTitle("ahk_id " cursorHwnd)
    pid := WinGetPID("ahk_id " cursorHwnd)
    processPath := ProcessGetPath(pid)
    SplitPath(processPath, &processName)
    simpleName := RegExReplace(processName, "\.exe$", "")
    
    ; 获取所有任务栏窗口
    winList := WinGetList(,, "Program Manager")
    minimizedCount := 0
    
    ; 创建置顶窗口句柄的哈希表，用于快速查找
    pinnedHwndsMap := Map()
    for _, pinnedHwnd in pinnedWindows {
        pinnedHwndsMap[pinnedHwnd] := true
    }
    
    ; 遍历所有窗口，最小化除了光标下窗口之外的所有任务栏窗口
    for hwnd in winList {
        ; 跳过不是任务栏窗口的窗口
        if (!IsTaskbarWindow(hwnd))
            continue
            
        ; 跳过光标下的窗口
        if (hwnd = cursorHwnd)
            continue
            
        ; 获取窗口信息
        winTitle := WinGetTitle("ahk_id " hwnd)
        winPid := WinGetPID("ahk_id " hwnd)
        winProcessPath := ProcessGetPath(winPid)
        SplitPath(winProcessPath, &winProcessName)
        winSimpleName := RegExReplace(winProcessName, "\.exe$", "")
        
        ; 检查窗口是否在黑名单中
        className := WinGetClass("ahk_id " hwnd)
        
        ; 检查进程名是否在黑名单中
        if (HasVal(blacklistProcessNames, winProcessName))
            continue
            
        ; 检查类名是否在黑名单中
        isBlacklistedClass := false
        for blackClass in blacklistClasses {
            if (InStr(className, blackClass) || className = blackClass) {
                isBlacklistedClass := true
                break
            }
        }
        
        if (isBlacklistedClass)
            continue
            
        ; 检查窗口是否置顶 - 跳过置顶窗口
        exStyle := WinGetExStyle("ahk_id " hwnd)
        isPinned := (exStyle & 0x8) != 0  ; WS_EX_TOPMOST
        
        ; 如果窗口在置顶列表中或者有置顶标志，跳过
        if (isPinned || pinnedHwndsMap.Has(hwnd))
            continue
            
        ; 记录窗口当前状态
        wasMinimized := WinGetMinMax("ahk_id " hwnd) = -1
        
        ; 如果窗口未最小化，则最小化它
        if (!wasMinimized) {
            ; 最小化窗口
            WinMinimize("ahk_id " hwnd)
            
            ; 记录被最小化的窗口信息
            minimizedWindows.Push({
                hwnd: hwnd,
                title: winTitle,
                processName: winProcessName,
                simpleName: winSimpleName,
                wasMinimized: wasMinimized
            })
            
            minimizedCount++
        }
    }
    
    ; 显示操作结果
    if (minimizedCount > 0) {
        ShowTooltip("已最小化 " minimizedCount " 个其他窗口`n保留窗口: " simpleName "`n再次按Ctrl+Win+Z恢复")
    } else {
        ShowTooltip("没有其他需要最小化的窗口")
    }
}

; 恢复最小化的窗口
RestoreMinimizedWindows() {
    global minimizedWindows
    
    ; 如果没有最小化的窗口记录，提示用户
    if (minimizedWindows.Length = 0) {
        ShowTooltip("没有需要恢复的窗口")
        return
    }
    
    restoredCount := 0
    invalidCount := 0
    restoredNames := []
    
    ; 恢复所有记录的窗口
    for i, winInfo in minimizedWindows {
        try {
            ; 检查窗口是否存在
            if (WinExist("ahk_id " winInfo.hwnd)) {
                ; 恢复窗口
                WinRestore("ahk_id " winInfo.hwnd)
                
                ; 尝试激活窗口
                WinActivate("ahk_id " winInfo.hwnd)
                
                ; 收集已恢复窗口的简称
                if (winInfo.HasOwnProp("simpleName") && winInfo.simpleName != "")
                    restoredNames.Push(winInfo.simpleName)
                else if (winInfo.HasOwnProp("processName"))
                    restoredNames.Push(RegExReplace(winInfo.processName, "\.exe$", ""))
                    
                restoredCount++
            } else {
                invalidCount++
            }
        } catch as e {
            ShowTooltip("恢复窗口时出错: " e.Message)
            invalidCount++
        }
    }
    
    ; 在所有窗口恢复后，尝试激活第一个恢复的窗口
    if (restoredCount > 0) {
        try {
            ; 获取第一个有效的恢复窗口
            for _, winInfo in minimizedWindows {
                if (WinExist("ahk_id " winInfo.hwnd)) {
                    WinActivate("ahk_id " winInfo.hwnd)
                    break
                }
            }
        } catch {
            ; 忽略可能的错误
        }
    }
    
    ; 清空记录
    minimizedWindows := []
    
    ; 显示操作结果
    if (restoredCount > 0) {
        resultText := "已恢复 " restoredCount " 个窗口"
        
        ; 显示前3个恢复的窗口名称
        if (restoredNames.Length > 0) {
            resultText .= "`n恢复的窗口: "
            maxNamesToShow := Min(restoredNames.Length, 3)
            Loop maxNamesToShow {
                resultText .= restoredNames[A_Index]
                if (A_Index < maxNamesToShow)
                    resultText .= ", "
            }
            
            if (restoredNames.Length > maxNamesToShow)
                resultText .= " 等..."
        }
        
        if (invalidCount > 0)
            resultText .= "`n" invalidCount " 个窗口已不存在"
        
        ToolTip(resultText)
    } else if (invalidCount > 0) {
        ToolTip("所有记录的窗口已不存在")
    }
    
    SetTimer () => ToolTip(), -2000
}

; =====================================================================
; 基本功能和辅助函数
; =====================================================================

; 切换调试提示的显示状态
ToggleDebugTooltips() {
    global showDebugTooltips
    showDebugTooltips := !showDebugTooltips
    ToolTip("调试信息显示: " (showDebugTooltips ? "开启" : "关闭"))
    SetTimer () => ToolTip(), -tipDuration
}

; 显示调试信息提示
ShowDebugTooltip(text, duration := 3000) {
    global showDebugTooltips, debugTipDuration
    
    if !showDebugTooltips
        return
        
    ToolTip(text)
    SetTimer () => ToolTip(), -duration
}

; 检查值是否在数组中
HasVal(arr, val) {
    for v in arr {
        if v = val
            return true
    }
    return false
}

; 字符串哈希算法，用于区分同一程序的不同窗口
StrHash(str) {
    hash := 0
    if (StrLen(str) == 0)
        return hash
        
    for i, char in StrSplit(str) {
        hash := ((hash << 5) - hash) + Ord(char)
        hash := hash & hash  ; 转换为32位整数
    }
    
    return hash
}

; 添加CapsLock+兼容性处理，用于调试可能的按键冲突问题
ShowKeyPressDebug(keyName, source := "AHK") {
    global showDebugTooltips
    
    if !showDebugTooltips
        return
        
    ToolTip("按键检测：" keyName " (来源: " source ")")
    SetTimer () => ToolTip(), -1000
}

; 定义常用程序的优先级映射
GetProcessPriority(processName) {
    return 900 + Mod(StrHash(processName), 100)  ; 使用哈希确保同一进程名总是获得相同的优先级
}

; 判断窗口是否为任务栏窗口
IsTaskbarWindow(hwnd) {
    global showDebugTooltips, blacklistClasses, blacklistProcessNames
    
    ; 隐藏或不可见窗口不算
    if !WinExist("ahk_id " hwnd) || !DllCall("IsWindowVisible", "Ptr", hwnd)
        return false
    
    ; 获取窗口标题
    title := WinGetTitle("ahk_id " hwnd)
    
    ; 空标题的窗口不显示在任务栏，排除
    if title = ""
        return false
    
    ; 获取窗口类名称
    className := WinGetClass("ahk_id " hwnd)
    
    ; 获取进程ID和名称
    pid := WinGetPID("ahk_id " hwnd)
    processPath := ProcessGetPath(pid)
    SplitPath(processPath, &processName)
    
    ; 排除在自定义黑名单中的进程
    if HasVal(blacklistProcessNames, processName) {
        if (showDebugTooltips) {
            ToolTip("排除黑名单进程: " processName)
            SetTimer () => ToolTip(), -1000
        }
        return false
    }
        
    ; 排除在自定义黑名单中的窗口类名
    for blackClass in blacklistClasses {
        if (InStr(className, blackClass) || className = blackClass) {
            if (showDebugTooltips) {
                ToolTip("排除黑名单类名: " className)
                SetTimer () => ToolTip(), -1000
            }
            return false
        }
    }
    
    ; 获取窗口样式
    style := WinGetStyle("ahk_id " hwnd)
    exStyle := WinGetExStyle("ahk_id " hwnd)
    
    ; 排除特定的窗口类型
    excludeClasses := [
        "Shell_TrayWnd", "Shell_SecondaryTrayWnd", "Progman", "WorkerW", 
        "XamlExplorerHostIslandWindow", "Windows.UI.Core.CoreWindow", 
        "ApplicationFrameWindow", "DV2ControlHost", "MsgrIMEWindowClass",
        "tooltips_class32", "ToolTip", "Popup", "PopupMenu", "NotifyIconOverflowWindow",
        "#32768", "DesktopWindowXamlSource", "ForegroundStaging", "Windows.UI.Composition.DesktopWindowContentBridge",
        "TaskListThumbnailWnd", "TaskSwitcherWnd", "TaskSwitcherOverlayWnd", "MultitaskingViewFrame"
    ]
    
    for _, excludeClass in excludeClasses {
        if className = excludeClass {
            if (showDebugTooltips) {
                ToolTip("排除系统类名: " className)
                SetTimer () => ToolTip(), -1000
            }
            return false
        }
    }
    
    ; 排除工具窗口（通常不在任务栏显示）
    if (exStyle & 0x80)  ; WS_EX_TOOLWINDOW
        return false
    
    ; 排除没有任务栏图标的窗口
    if !(exStyle & 0x40000) ; !WS_EX_APPWINDOW
    {
        ; 检查窗口是否有拥有者
        ownerHwnd := DllCall("GetWindow", "Ptr", hwnd, "UInt", 4, "Ptr") ; GW_OWNER = 4
        
        ; 有拥有者的窗口通常不显示在任务栏
        if ownerHwnd
            return false
    }
    
    ; 处理特殊的Explorer.exe窗口
    if (processName = "explorer.exe") {
        ; 保留文件资源管理器窗口 (CabinetWClass)
        if (className = "CabinetWClass")
            return true
            
        ; 排除桌面和其他系统UI窗口
        if (className != "CabinetWClass")
            return false
    }
    
    ; 排除 Windows 11 任务视图按钮
    if InStr(title, "Task View") && processName = "explorer.exe"
        return false
        
    ; 排除按钮、托盘窗口等特殊UI元素
    if (InStr(className, "Button") || InStr(className, "Tray") || InStr(className, "Notification"))
        return false
    
    ; 调试信息
    if (showDebugTooltips) {
        ToolTip("有效窗口: " title "`n进程: " processName "`n类名: " className)
        SetTimer () => ToolTip(), -2000
    }
    
    return true
}

; =====================================================================
; 窗口切换功能
; =====================================================================

; 任务栏窗口切换主函数
SwitchTaskbarWindow(direction) {
    ; 自动检测并切换到正确的窗口切换函数
    
    ; 如果虚拟环境中有窗口，优先使用虚拟环境切换
    if (virtualEnvEnabled && virtualEnvWindows.Length > 0)
        SwitchVirtualEnvWindow(direction)
    else
        SwitchNormalTaskbarWindow(direction)
}

; 普通任务栏窗口切换函数
SwitchNormalTaskbarWindow(direction) {
    global useTaskbarOrder, showDebugTooltips, longTipDuration, debugTipDuration, tipDuration
    
    ; 获取所有可见且在任务栏上的窗口
    taskbarWindows := []
    
    ; 使用 WinGetList 获取所有可见窗口
    winList := WinGetList(,, "Program Manager")
    
    ; 调试输出原始窗口数量
    if showDebugTooltips {
        ToolTip("原始窗口数量: " winList.Length)
        SetTimer () => ToolTip(), -debugTipDuration
    }
    
    ; 筛选出应该显示在任务栏上的窗口
    for hwnd in winList {
        if IsTaskbarWindow(hwnd) {
            ; 获取进程名称和路径
            pid := WinGetPID("ahk_id " hwnd)
            processPath := ProcessGetPath(pid)
            SplitPath(processPath, &processName)
            
            title := WinGetTitle("ahk_id " hwnd)
            
            ; 获取窗口优先级
            priority := GetProcessPriority(processName)
            
            ; 添加窗口信息到数组
            taskbarWindows.Push({ 
                hwnd: hwnd, 
                pid: pid,
                processName: processName,
                title: title,
                priority: priority
            })
        }
    }
    
    ; 调试输出找到的任务栏窗口
    if showDebugTooltips {
        ToolTip("找到的任务栏窗口数量: " taskbarWindows.Length)
        SetTimer () => ToolTip(), -debugTipDuration
    }
    
    ; 如果没有找到任何窗口，退出
    if (taskbarWindows.Length = 0) {
        if showDebugTooltips {
            ToolTip("没有找到可切换的窗口")
            SetTimer () => ToolTip(), -tipDuration
        }
        return
    }
    
    ; 按PID对数组进行排序，同一PID的窗口按标题排序
    taskbarWindows := SortByPID(taskbarWindows)
    
    ; 调试输出排序后的窗口列表
    if showDebugTooltips {
        debugText := "排序后的窗口列表 (按PID和标题):`n"
        for i, win in taskbarWindows {
            debugText .= i ": " win.title " (PID:" win.pid ", 句柄:" win.hwnd ")`n"
            if (i > 8)  ; 显示前8个窗口
                break
        }
        ToolTip(debugText)
        SetTimer () => ToolTip(), -3000
    }
    
    ; 获取当前活动窗口
    try {
        activeHwnd := WinGetID("A")
    } catch {
        ; 如果无法获取活动窗口，可能用户在桌面上
        ; 此时直接激活第一个窗口
        if (taskbarWindows.Length > 0) {
            nextIndex := direction > 0 ? 1 : taskbarWindows.Length
            nextHwnd := taskbarWindows[nextIndex].hwnd
            
            WinActivate("ahk_id " nextHwnd)
            
            ; 只在调试模式下显示切换信息
            if showDebugTooltips {
                ToolTip("从桌面切换到: " taskbarWindows[nextIndex].title " (句柄:" nextHwnd ")")
                SetTimer () => ToolTip(), -debugTipDuration
            }
            
            ; 检查新激活的窗口是否是最大化的
            Sleep(30)
            if (WinGetMinMax("ahk_id " nextHwnd) = 1) {
                MinimizeOtherMaximizedWindows(nextHwnd, taskbarWindows)
            }
        }
        return
    }
    
    ; 查找当前活动窗口在列表中的位置
    currentIndex := 0
    
    for i, win in taskbarWindows {
        if win.hwnd = activeHwnd {
            currentIndex := i
            break
        }
    }
    
    ; 调试输出当前窗口信息
    if showDebugTooltips && currentIndex > 0 {
        ToolTip("当前窗口: " taskbarWindows[currentIndex].title " (索引:" currentIndex ", 句柄:" activeHwnd ")")
        SetTimer () => ToolTip(), -debugTipDuration
    }
    
    ; 计算下一个窗口的索引
    if currentIndex = 0 {
        ; 如果当前窗口不在列表中，从列表的开头或结尾开始
        nextIndex := (direction > 0) ? 1 : taskbarWindows.Length
    } else {
        ; 否则，根据方向移动到下一个窗口
        nextIndex := currentIndex + direction
        
        ; 处理索引越界
        if nextIndex > taskbarWindows.Length
            nextIndex := 1
        else if nextIndex < 1
            nextIndex := taskbarWindows.Length
    }
    
    ; 如果找不到下一个窗口，退出
    if !taskbarWindows.Has(nextIndex) {
        if showDebugTooltips {
            ToolTip("无法找到下一个窗口")
            SetTimer () => ToolTip(), -tipDuration
        }
        return
    }
    
    ; 获取下一个窗口的句柄
    nextHwnd := taskbarWindows[nextIndex].hwnd
    
    ; 如果下一个窗口就是当前窗口，不进行操作
    if (nextHwnd = activeHwnd) {
        if showDebugTooltips {
            ToolTip("下一个窗口就是当前窗口，跳过激活")
            SetTimer () => ToolTip(), -tipDuration
        }
        return
    }
    
    ; 调试输出切换信息
    if showDebugTooltips {
        fromTitle := currentIndex > 0 ? taskbarWindows[currentIndex].title : "未知窗口"
        toTitle := taskbarWindows[nextIndex].title
        fromHwnd := currentIndex > 0 ? taskbarWindows[currentIndex].hwnd : 0
        toHwnd := taskbarWindows[nextIndex].hwnd
        
        ; 显示更详细的切换信息
        switchInfo := "从窗口 " currentIndex " 切换到 " nextIndex "`n"
        switchInfo .= "从: " fromTitle " (句柄:" fromHwnd ")`n"
        switchInfo .= "到: " toTitle " (句柄:" toHwnd ")`n"
        switchInfo .= "方向: " (direction > 0 ? "正向" : "反向")
        
        ToolTip(switchInfo)
        SetTimer () => ToolTip(), -debugTipDuration
    }
    
    ; 激活下一个窗口
    WinActivate("ahk_id " nextHwnd)
    
    ; 等待窗口激活
    Sleep(30)
    
    ; 检查新激活的窗口是否是最大化的
    if (WinGetMinMax("ahk_id " nextHwnd) = 1) {
        ; 如果是最大化的，则最小化其他所有最大化窗口
        MinimizeOtherMaximizedWindows(nextHwnd, taskbarWindows)
    }
}

; 最小化除了指定窗口外的所有最大化窗口
MinimizeOtherMaximizedWindows(activeHwnd, taskbarWindows) {
    global showDebugTooltips
    
    minimizedCount := 0
    
    for _, win in taskbarWindows {
        ; 跳过当前活动窗口
        if (win.hwnd = activeHwnd)
            continue
            
        ; 检查窗口是否最大化
        if (WinGetMinMax("ahk_id " win.hwnd) = 1) {
            ; 最小化该窗口
            WinMinimize("ahk_id " win.hwnd)
            minimizedCount++
        }
    }
    
    if (showDebugTooltips && minimizedCount > 0) {
        ToolTip("已最小化 " minimizedCount " 个其他最大化窗口")
        SetTimer () => ToolTip(), -debugTipDuration
    }
}

; =====================================================================
; 虚拟环境功能
; =====================================================================

; 清理虚拟环境中无效的窗口
CleanupVirtualEnvWindows() {
    global virtualEnvWindows, virtualEnvEnabled, showDebugTooltips
    
    if (virtualEnvWindows.Length = 0)
        return
        
    oldCount := virtualEnvWindows.Length
    validWindows := []
    for win in virtualEnvWindows {
        if WinExist("ahk_id " win.hwnd) {
            validWindows.Push(win)
        }
    }
    
    newCount := validWindows.Length
    virtualEnvWindows := validWindows
    
    ; 如果清理后没有窗口，禁用虚拟环境
    if (newCount = 0) {
        virtualEnvEnabled := false
    }
    
    ; 如果有窗口被移除，显示提示
    if (oldCount > newCount && showDebugTooltips) {
        ToolTip("已自动清理 " (oldCount - newCount) " 个已关闭的虚拟环境窗口`n剩余 " newCount " 个窗口")
        SetTimer () => ToolTip(), -2000
    }
}

; 添加窗口到虚拟环境
AddToVirtualEnv(hwnd) {
    global virtualEnvEnabled, virtualEnvWindows
    
    try {
        ; 验证窗口是否有效
        if !WinExist("ahk_id " hwnd) {
            ToolTip("无效的窗口句柄")
            SetTimer () => ToolTip(), -2000
            return
        }
        
        ; 获取窗口信息
        title := WinGetTitle("ahk_id " hwnd)
        
        ; 检查窗口是否有效
        if !IsTaskbarWindow(hwnd) {
            ToolTip("当前窗口不是任务栏窗口，无法添加到虚拟环境")
            SetTimer () => ToolTip(), -2000
            return
        }
        
        ; 获取窗口进程信息
        pid := WinGetPID("ahk_id " hwnd)
        processPath := ProcessGetPath(pid)
        SplitPath(processPath, &processName)
        
        ; 简化进程名称，移除.exe扩展名
        simpleName := RegExReplace(processName, "\.exe$", "")
        
        ; 检查窗口是否已在虚拟环境中
        for win in virtualEnvWindows {
            if win.hwnd = hwnd {
                ToolTip("窗口已在虚拟环境中: " simpleName)
                SetTimer () => ToolTip(), -2000
                return
            }
        }
        
        ; 添加窗口到虚拟环境
        virtualEnvWindows.Push({
            hwnd: hwnd,
            title: title,
            processName: processName,
            simpleName: simpleName,
            pid: pid
        })
        
        ; 启用虚拟环境
        virtualEnvEnabled := true
        
        ; 显示添加成功消息，使用简化的进程名
        windowCount := virtualEnvWindows.Length
        ToolTip("已添加窗口到虚拟环境: " simpleName "`n虚拟环境中现有 " windowCount " 个窗口")
        SetTimer () => ToolTip(), -2000
        
    } catch as e {
        ToolTip("添加窗口到虚拟环境失败: " e.Message)
        SetTimer () => ToolTip(), -2000
    }
}

; 从虚拟环境中移除窗口
RemoveFromVirtualEnv(hwnd) {
    global virtualEnvEnabled, virtualEnvWindows
    
    if !virtualEnvEnabled || virtualEnvWindows.Length = 0 {
        ToolTip("虚拟环境未启用或为空")
        SetTimer () => ToolTip(), -2000
        return
    }
    
    try {
        ; 获取窗口标题和简化名称
        title := WinGetTitle("ahk_id " hwnd)
        simpleName := ""
        
        ; 查找窗口在虚拟环境中的位置
        windowIndex := 0
        for i, win in virtualEnvWindows {
            if win.hwnd = hwnd {
                windowIndex := i
                ; 使用try-catch来检查属性是否存在
                try {
                    simpleName := win.simpleName
                } catch {
                    ; 如果访问simpleName属性出错，使用进程名的简化版本
                    simpleName := RegExReplace(win.processName, "\.exe$", "")
                }
                break
            }
        }
        
        ; 如果找到窗口，则从虚拟环境中移除
        if windowIndex > 0 {
            ; 从虚拟环境中移除窗口
            virtualEnvWindows.RemoveAt(windowIndex)
            
            ; 显示通知，使用简化的进程名
            windowCount := virtualEnvWindows.Length
            ToolTip("已从虚拟环境中移除窗口: " simpleName "`n虚拟环境中还有 " windowCount " 个窗口")
            
            ; 如果移除后虚拟环境为空，则自动退出虚拟环境
            if windowCount = 0 {
                virtualEnvEnabled := false
                ToolTip("虚拟环境现已为空，已自动退出虚拟环境模式")
            }
            
            SetTimer () => ToolTip(), -2000
        } else {
            ToolTip("当前窗口不在虚拟环境中")
            SetTimer () => ToolTip(), -2000
        }
        
    } catch as e {
        ToolTip("移除窗口失败: " e.Message)
        SetTimer () => ToolTip(), -2000
    }
}

; 清除虚拟环境
ClearVirtualEnv() {
    global virtualEnvEnabled, virtualEnvWindows
    
    if virtualEnvEnabled || virtualEnvWindows.Length > 0 {
        virtualEnvEnabled := false
        windowCount := virtualEnvWindows.Length
        virtualEnvWindows := []
        
        ToolTip("已清除虚拟环境，移除了 " windowCount " 个窗口")
        SetTimer () => ToolTip(), -2000
    } else {
        ToolTip("虚拟环境已是空的")
        SetTimer () => ToolTip(), -2000
    }
}

; 智能虚拟环境切换函数 - 如果窗口在虚拟环境中则移除，否则添加
SmartVirtualEnvToggle() {
    ShowDebugTooltip("SmartVirtualEnvToggle触发")
    global virtualEnvEnabled, virtualEnvWindows
    
    ; 先清理虚拟环境中已关闭的窗口
    CleanupVirtualEnvWindows()
    
    try {
        ; 获取光标下的窗口，而不是当前活动窗口
        targetHwnd := GetWindowUnderCursor()
        
        ; 确保窗口有效
        if (!targetHwnd || !WinExist("ahk_id " targetHwnd)) {
            ToolTip("光标下无有效窗口，无法管理虚拟环境")
            SetTimer () => ToolTip(), -2000
            return
        }
        
        targetTitle := WinGetTitle("ahk_id " targetHwnd)
        
        ; 获取窗口进程信息
        pid := WinGetPID("ahk_id " targetHwnd)
        processPath := ProcessGetPath(pid)
        SplitPath(processPath, &processName)
        
        ; 简化进程名称，移除.exe扩展名
        simpleName := RegExReplace(processName, "\.exe$", "")
        
        ; 检查窗口是否有效
        if !IsTaskbarWindow(targetHwnd) {
            ToolTip("光标下的窗口不是任务栏窗口，无法管理虚拟环境")
            SetTimer () => ToolTip(), -2000
            return
        }
        
        ; 检查窗口是否已在虚拟环境中
        windowIndex := 0
        for i, win in virtualEnvWindows {
            if win.hwnd = targetHwnd {
                windowIndex := i
                break
            }
        }
        
        ; 如果窗口在虚拟环境中，则移除它
        if (windowIndex > 0) {
            RemoveFromVirtualEnv(targetHwnd)
        } 
        ; 否则添加到虚拟环境
        else {
            AddToVirtualEnv(targetHwnd)
        }
        
    } catch as e {
        ToolTip("操作窗口失败: " e.Message)
        SetTimer () => ToolTip(), -2000
    }
}

; 在虚拟环境窗口间切换
SwitchVirtualEnvWindow(direction) {
    global virtualEnvWindows, virtualEnvEnabled, showDebugTooltips
    
    ; 清理无效窗口（可能已经关闭的窗口）
    validWindows := []
    for win in virtualEnvWindows {
        if WinExist("ahk_id " win.hwnd) {
            ; 确保每个窗口对象都有simpleName属性
            try {
                simpleName := win.simpleName
            } catch {
                ; 如果无法直接访问simpleName，创建并添加
                win.simpleName := RegExReplace(win.processName, "\.exe$", "")
            }
            validWindows.Push(win)
        }
    }
    virtualEnvWindows := validWindows
    
    ; 如果没有有效窗口，则退出虚拟环境
    if virtualEnvWindows.Length = 0 {
        ClearVirtualEnv()
        return
    }
    
    ; 调试输出虚拟环境窗口列表
    if showDebugTooltips {
        debugText := "虚拟环境窗口列表 (共" virtualEnvWindows.Length "个):`n"
        for i, win in virtualEnvWindows {
            ; 使用简化的进程名
            try {
                simpleName := win.simpleName
            } catch {
                simpleName := RegExReplace(win.processName, "\.exe$", "")
            }
            debugText .= i ": " simpleName "`n"
            if (i > 5)  ; 只显示前5个
                break
        }
        ToolTip(debugText)
        SetTimer () => ToolTip(), -3000
    }
    
    ; 如果虚拟环境中只有一个窗口，不执行切换，而是激活该窗口并提示用户
    if virtualEnvWindows.Length = 1 {
        singleHwnd := virtualEnvWindows[1].hwnd
        try {
            simpleName := virtualEnvWindows[1].simpleName
        } catch {
            simpleName := RegExReplace(virtualEnvWindows[1].processName, "\.exe$", "")
        }
        
        ; 检查当前活动窗口是否就是虚拟环境中的那个窗口
        try {
            activeHwnd := WinGetID("A")
            if (activeHwnd = singleHwnd) {
                ; 如果当前已经在虚拟环境的窗口中，提示用户
                if showDebugTooltips {
                    ToolTip("虚拟环境中只有一个窗口: " simpleName)
                    SetTimer () => ToolTip(), -2000
                }
                return
            }
        } catch {
            ; 如果无法获取活动窗口，直接激活虚拟环境中的窗口
        }
        
        ; 激活虚拟环境中的唯一窗口
        WinActivate("ahk_id " singleHwnd)
        if showDebugTooltips {
            ToolTip("已激活虚拟环境中的唯一窗口: " simpleName)
            SetTimer () => ToolTip(), -2000
        }
        
        ; 检查新激活的窗口是否是最大化的
        Sleep(30)
        if (WinGetMinMax("ahk_id " singleHwnd) = 1) {
            ; 获取所有任务栏窗口用于最小化其他最大化窗口
            winList := WinGetList(,, "Program Manager")
            taskbarWindows := []
            for hwnd in winList {
                if IsTaskbarWindow(hwnd) {
                    taskbarWindows.Push({ hwnd: hwnd })
                }
            }
            MinimizeOtherMaximizedWindows(singleHwnd, taskbarWindows)
        }
        
        return
    }
    
    ; 获取当前活动窗口
    try {
        activeHwnd := WinGetID("A")
    } catch {
        ; 如果无法获取活动窗口，激活第一个窗口
        if virtualEnvWindows.Length > 0 {
            nextIndex := direction > 0 ? 1 : virtualEnvWindows.Length
            nextHwnd := virtualEnvWindows[nextIndex].hwnd
            WinActivate("ahk_id " nextHwnd)
            
            ; 检查新激活的窗口是否是最大化的
            Sleep(30)
            if (WinGetMinMax("ahk_id " nextHwnd) = 1) {
                ; 获取所有任务栏窗口用于最小化其他最大化窗口
                winList := WinGetList(,, "Program Manager")
                taskbarWindows := []
                for hwnd in winList {
                    if IsTaskbarWindow(hwnd) {
                        taskbarWindows.Push({ hwnd: hwnd })
                    }
                }
                MinimizeOtherMaximizedWindows(nextHwnd, taskbarWindows)
            }
        }
        return
    }
    
    ; 查找当前活动窗口在虚拟环境中的位置
    currentIndex := 0
    for i, win in virtualEnvWindows {
        if win.hwnd = activeHwnd {
            currentIndex := i
            break
        }
    }
    
    ; 如果当前活动窗口不在虚拟环境中，激活第一个虚拟环境窗口
    if currentIndex = 0 {
        nextIndex := (direction > 0) ? 1 : virtualEnvWindows.Length
    } else {
        nextIndex := currentIndex + direction
        
        ; 处理索引越界
        if nextIndex > virtualEnvWindows.Length
            nextIndex := 1
        else if nextIndex < 1
            nextIndex := virtualEnvWindows.Length
    }
    
    ; 获取下一个窗口的句柄
    if !virtualEnvWindows.Has(nextIndex)
        return
        
    nextHwnd := virtualEnvWindows[nextIndex].hwnd
    
    ; 如果下一个窗口就是当前窗口，不进行操作
    if (nextHwnd = activeHwnd) {
        if showDebugTooltips {
            ToolTip("下一个窗口就是当前窗口，跳过激活")
            SetTimer () => ToolTip(), -2000
        }
        return
    }
    
    ; 调试输出切换信息
    if showDebugTooltips {
        ToolTip("从窗口 " currentIndex " 切换到窗口 " nextIndex)
        SetTimer () => ToolTip(), -2000
    }
    
    ; 激活下一个窗口
    WinActivate("ahk_id " nextHwnd)
    
    ; 等待窗口激活
    Sleep(30)
    
    ; 检查新激活的窗口是否是最大化的
    if (WinGetMinMax("ahk_id " nextHwnd) = 1) {
        ; 获取所有任务栏窗口用于最小化其他最大化窗口
        winList := WinGetList(,, "Program Manager")
        taskbarWindows := []
        for hwnd in winList {
            if IsTaskbarWindow(hwnd) {
                taskbarWindows.Push({ hwnd: hwnd })
            }
        }
        MinimizeOtherMaximizedWindows(nextHwnd, taskbarWindows)
    }
}

; =====================================================================
; CapsLock+ 兼容模块 - 文本编辑增强与鼠标控制
; =====================================================================

; 全局设置
global mouseSpeedValue := 3  ; 鼠标临时速度值(1-20)
global mouseOriginalSpeed := 10  ; 存储原始鼠标速度

; 初始化CapsLock+兼容模块
InitCapsLockPlusModule() {
    ; 获取当前系统鼠标速度作为默认值
    try {
        RegRead(&mouseOriginalSpeed, "HKCU\Control Panel\Mouse", "MouseSensitivity")
    } catch {
        mouseOriginalSpeed := 10  ; 默认值
    }
    
    ; 移除启动提示
    ; ToolTip("CapsLock+兼容模块已加载")
    ; SetTimer () => ToolTip(), -2000
}

; =====================================================================
; 鼠标速度临时调整功能
; =====================================================================

AdjustMouseSpeed() {
    global mouseSpeedValue, mouseOriginalSpeed
    
    ; 保存当前鼠标速度
    try {
        RegRead(&mouseOriginalSpeed, "HKCU\Control Panel\Mouse", "MouseSensitivity")
    } catch {
        mouseOriginalSpeed := 10
    }
    
    ; 设置鼠标速度为临时值
    try {
        ; 确保mouseSpeedValue是整数
        speedToSet := Integer(mouseSpeedValue)
        
        ; 使用RegWrite写入注册表
        RegWrite(speedToSet, "REG_SZ", "HKCU\Control Panel\Mouse", "MouseSensitivity")
        
        ; 通知系统更新鼠标设置
        ; SPI_SETMOUSESPEED = 0x71
        ; 最后的0表示不永久保存更改
        DllCall("SystemParametersInfo", "UInt", 0x71, "UInt", 0, "UInt", speedToSet, "UInt", 0)
        
        ; 移除提示，使操作静默
        ; ToolTip("已临时设置鼠标速度: " speedToSet)
        ; SetTimer () => ToolTip(), -1000
    } catch as e {
        ; 仅在出错时显示提示
        ShowTooltip("设置鼠标速度失败: " e.Message)
    }
    
    ; 等待LAlt键释放
    KeyWait("LAlt")
    
    ; 恢复原始鼠标速度
    try {
        ; 确保mouseOriginalSpeed是整数
        originalSpeed := Integer(mouseOriginalSpeed)
        
        ; 使用RegWrite写入注册表
        RegWrite(originalSpeed, "REG_SZ", "HKCU\Control Panel\Mouse", "MouseSensitivity")
        
        ; 通知系统更新鼠标设置
        DllCall("SystemParametersInfo", "UInt", 0x71, "UInt", 0, "UInt", originalSpeed, "UInt", 0)
        
        ; 移除提示，使操作静默
        ; ToolTip("已恢复鼠标速度: " originalSpeed)
        ; SetTimer () => ToolTip(), -1000
    } catch as e {
        ; 仅在出错时显示提示
        ShowTooltip("恢复鼠标速度失败: " e.Message)
    }
}

; =====================================================================
; 文本编辑增强功能 - 光标移动和选择
; =====================================================================

; 光标移动函数
MoveLeft(repeat := 1) {
    Loop repeat
        Send("{Left}")
}

MoveRight(repeat := 1) {
    Loop repeat
        Send("{Right}")
}

MoveUp(repeat := 1) {
    Loop repeat
        Send("{Up}")
}

MoveDown(repeat := 1) {
    Loop repeat
        Send("{Down}")
}

MoveWordLeft(repeat := 1) {
    Loop repeat
        Send("^{Left}")
}

MoveWordRight(repeat := 1) {
    Loop repeat
        Send("^{Right}")
}

MoveHome() {
    Send("{Home}")
}

MoveEnd() {
    Send("{End}")
}

MoveToPageBeginning() {
    Send("^{Home}")
}

MoveToPageEnd() {
    Send("^{End}")
}

; 文本选择函数
SelectLeft(repeat := 1) {
    Loop repeat
        Send("+{Left}")
}

SelectRight(repeat := 1) {
    Loop repeat
        Send("+{Right}")
}

SelectUp(repeat := 1) {
    Loop repeat
        Send("+{Up}")
}

SelectDown(repeat := 1) {
    Loop repeat
        Send("+{Down}")
}

SelectWordLeft(repeat := 1) {
    Loop repeat
        Send("^+{Left}")
}

SelectWordRight(repeat := 1) {
    Loop repeat
        Send("^+{Right}")
}

SelectHome() {
    Send("+{Home}")
}

SelectEnd() {
    Send("+{End}")
}

SelectToPageBeginning() {
    Send("^+{Home}")
}

SelectToPageEnd() {
    Send("^+{End}")
}

SelectCurrentWord() {
    ; 双击选中当前单词
    Send("^{left}+^{right}")
}

SelectCurrentLine() {
    ; 选择当前行: Home, 然后Shift+End
    Send("{Home}+{End}")
}

; 空格选择的全局变量
global lastSpaceTime := 0  ; 上次空格按下时间
global pendingOperation := false  ; 是否有待执行的操作
global doubleClickThreshold := 500  ; 双击阈值(ms)
global hasExecutedSingleClick := false  ; 是否已执行过单击操作

; 异步选中单词的处理函数
SelectWordTimer()
{
    global pendingOperation, hasExecutedSingleClick
    
    ; 如果操作已被取消（发生了双击），直接返回
    if (!pendingOperation)
        return
    
    ; 阻断输入，防止误触
    BlockInput("On")
    
    ; 检查是否已有选中内容
    originalClip := ClipboardAll()
    A_Clipboard := ""
    Send("^c")
    hasSelection := ClipWait(0.05, 0)
    selectedText := A_Clipboard
    A_Clipboard := originalClip
    originalClip := ""
    
    ; 如果已有选中内容，不执行选中单词
    if (hasSelection && selectedText != "") {
        ; 已有选中内容，不做任何操作
        pendingOperation := false
        hasExecutedSingleClick := true
        
        ; 恢复输入
        BlockInput("Off")
        return
    }
    
    ; 如果没有选中内容，执行选中单词
    if (pendingOperation) {
        SelectCurrentWord()
        pendingOperation := false
        hasExecutedSingleClick := true
    }
    
    ; 恢复输入
    BlockInput("Off")
}

; 文本编辑函数
DeleteLeft(repeat := 1) {
    Loop repeat
        Send("{Backspace}")
}

DeleteRight(repeat := 1) {
    Loop repeat
        Send("{Delete}")
}

DeleteWord() {
    Send("^+{Left}{Delete}")
}

ForwardDeleteWord() {
    Send("^+{Right}{Delete}")
}

DeleteLine() {
    Send("{Home}+{End}{Delete}")
}

DeleteToLineBeginning() {
    Send("+{Home}{Delete}")
}

DeleteToLineEnd() {
    Send("+{End}{Delete}")
}

DeleteToPageBeginning() {
    Send("^+{Home}{Delete}")
}

DeleteToPageEnd() {
    Send("^+{End}{Delete}")
}

DeleteAll() {
    Send("^a{Delete}")
}

EnterWherever() {
    Send("{End}{Enter}")
}

IndexWherever(){
    Send("{Home}{Enter}{Up}")
}

; =====================================================================
; 自定义跳转功能 - 状态机版本，使用InputHook
; =====================================================================

; 全局变量
global jumpMode := ""           ; 存储当前跳转模式
global jumpActive := false      ; 是否激活跳转模式
global jumpBuffer := ""         ; 存储输入的数字
global jumpPosition := {x: 0, y: 0}  ; 存储位置
global g_inputHook := {}        ; 存储输入钩子对象
global isWordJump := false      ; 标识是否是单词级跳转

GetCaretPosition() {
    static left := 0, top := 0, right := 0, bottom := 0
    
    ; 检查是否有光标
    if GetCaretPosEx(&left, &top, &right, &bottom, true) {
        ; 获取屏幕尺寸用于验证
        MonitorGetWorkArea(, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
        
        ; 检查是否为屏幕边缘坐标（可能是伪光标位置）
        isScreenEdge := (Abs(left - monitorRight) < 5 || Abs(bottom - monitorBottom) < 5 || (left < 5 && top < 5))
        
        ; 如果不是屏幕边缘，则认为是有效的光标位置
        if (!isScreenEdge) {
            return {x: left, y: bottom}
        }
    }
    
    ; 获取失败，返回false
    return false
}

GetCaretPosEx(&left?, &top?, &right?, &bottom?, useHook := false) {
    if getCaretPosFromGui(&hwnd := 0)
        return true
    try
        className := WinGetClass(hwnd)
    catch
        className := ""
    if className ~= "^(?:Windows|Microsoft)\.UI\..+"
        funcs := [getCaretPosFromUIA, getCaretPosFromHook, getCaretPosFromMSAA]
    else if className ~= "^HwndWrapper\[PowerShell_ISE\.exe;;[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\]"
        funcs := [getCaretPosFromHook, getCaretPosFromWpfCaret]
    else
        funcs := [getCaretPosFromMSAA, getCaretPosFromUIA, getCaretPosFromHook]
    for fn in funcs {
        if fn == getCaretPosFromHook && !useHook
            continue
        if fn()
            return true
    }
    return false

    getCaretPosFromGui(&hwnd) {
        x64 := A_PtrSize == 8
        guiThreadInfo := Buffer(x64 ? 72 : 48)
        NumPut("uint", guiThreadInfo.Size, guiThreadInfo)
        if DllCall("GetGUIThreadInfo", "uint", 0, "ptr", guiThreadInfo) {
            if hwnd := NumGet(guiThreadInfo, x64 ? 48 : 28, "ptr") {
                getRect(guiThreadInfo.Ptr + (x64 ? 56 : 32), &left, &top, &right, &bottom)
                scaleRect(getWindowScale(hwnd), &left, &top, &right, &bottom)
                clientToScreenRect(hwnd, &left, &top, &right, &bottom)
                return true
            }
            hwnd := NumGet(guiThreadInfo, x64 ? 16 : 12, "ptr")
        }
        return false
    }

    getCaretPosFromMSAA() {
        if !hOleacc := DllCall("LoadLibraryW", "str", "oleacc.dll", "ptr")
            return false
        hOleacc := { Ptr: hOleacc, __Delete: (_) => DllCall("FreeLibrary", "ptr", _) }
        static IID_IAccessible := guidFromString("{618736e0-3c3d-11cf-810c-00aa00389b71}")
        if !DllCall("oleacc\AccessibleObjectFromWindow", "ptr", hwnd, "uint", 0xfffffff8, "ptr", IID_IAccessible, "ptr*", accCaret := ComValue(13, 0), "int") {
            if A_PtrSize == 8 {
                varChild := Buffer(24, 0)
                NumPut("ushort", 3, varChild)
                hr := ComCall(22, accCaret, "int*", &x := 0, "int*", &y := 0, "int*", &w := 0, "int*", &h := 0, "ptr", varChild, "int")
            }
            else {
                hr := ComCall(22, accCaret, "int*", &x := 0, "int*", &y := 0, "int*", &w := 0, "int*", &h := 0, "int64", 3, "int64", 0, "int")
            }
            if !hr {
                pt := x | y << 32
                DllCall("ScreenToClient", "ptr", hwnd, "int64*", &pt)
                left := pt & 0xffffffff
                top := pt >> 32
                right := left + w
                bottom := top + h
                scaleRect(getWindowScale(hwnd), &left, &top, &right, &bottom)
                clientToScreenRect(hwnd, &left, &top, &right, &bottom)
                return true
            }
        }
        return false
    }

    getCaretPosFromUIA() {
        try {
            uia := ComObject("{E22AD333-B25F-460C-83D0-0581107395C9}", "{30CBE57D-D9D0-452A-AB13-7AC5AC4825EE}")
            ComCall(20, uia, "ptr*", cacheRequest := ComValue(13, 0)) ; uia->CreateCacheRequest(&cacheRequest);
            if !cacheRequest.Ptr
                return false
            ComCall(4, cacheRequest, "ptr", 10014) ; cacheRequest->AddPattern(UIA_TextPatternId);
            ComCall(4, cacheRequest, "ptr", 10024) ; cacheRequest->AddPattern(UIA_TextPattern2Id);

            ComCall(12, uia, "ptr", cacheRequest, "ptr*", focusedEle := ComValue(13, 0)) ; uia->GetFocusedElementBuildCache(cacheRequest, &focusedEle);
            if !focusedEle.Ptr
                return false

            static IID_IUIAutomationTextPattern2 := guidFromString("{506a921a-fcc9-409f-b23b-37eb74106872}")
            range := ComValue(13, 0)
            ComCall(15, focusedEle, "int", 10024, "ptr", IID_IUIAutomationTextPattern2, "ptr*", textPattern := ComValue(13, 0)) ; focusedEle->GetCachedPatternAs(UIA_TextPattern2Id, IID_PPV_ARGS(&textPattern));
            if textPattern.Ptr {
                ComCall(10, textPattern, "int*", &isActive := 0, "ptr*", range) ; textPattern->GetCaretRange(&isActive, &range);
                if range.Ptr
                    goto getRangeInfo
            }
            ; If no caret range, get selection range.
            static IID_IUIAutomationTextPattern := guidFromString("{32eba289-3583-42c9-9c59-3b6d9a1e9b6a}")
            ComCall(15, focusedEle, "int", 10014, "ptr", IID_IUIAutomationTextPattern, "ptr*", textPattern) ; focusedEle->GetCachedPatternAs(UIA_TextPatternId, IID_PPV_ARGS(&textPattern));
            if textPattern.Ptr {
                ComCall(5, textPattern, "ptr*", ranges := ComValue(13, 0)) ; textPattern->GetSelection(&ranges);
                if ranges.Ptr {
                    ; Retrieve the last selection range.
                    ComCall(3, ranges, "int*", &len := 0) ; ranges->get_Length(&len);
                    if len > 0 {
                        ComCall(4, ranges, "int", len - 1, "ptr*", range) ; ranges->GetElement(len - 1, &range);
                        if range.Ptr {
                            ; Collapse the range.
                            ComCall(15, range, "int", 0, "ptr", range, "int", 1) ; range->MoveEndpointByRange(TextPatternRangeEndpoint_Start, range, TextPatternRangeEndpoint_End);
                            goto getRangeInfo
                        }
                    }
                }
            }
            return false
getRangeInfo:
            psa := 0
            ; This is a degenerate text range, we have to expand it.
            ComCall(6, range, "int", 0) ; range->ExpandToEnclosingUnit(TextUnit_Character);
            ComCall(10, range, "ptr*", &psa) ; range->GetBoundingRectangles(&psa);
            if psa {
                rects := ComValue(0x2005, psa, 1) ; SafeArray<double>
                if rects.MaxIndex() >= 3 {
                    rects[2] := 0
                    goto end
                }
            }
            ; ExpandToEnclosingUnit by character may be invalid in some control if the range is at the end of the document.
            ; Assume that the range is at the end of the document and not in an empty line, try to expand it by line.
            ComCall(6, range, "int", 3) ; range->ExpandToEnclosingUnit(TextUnit_Line)
            ComCall(10, range, "ptr*", &psa) ; range->GetBoundingRectangles(&psa);
            if psa {
                rects := ComValue(0x2005, psa, 1) ; SafeArray<double>
                if rects.MaxIndex() >= 3 {
                    ; Here rects is {x, y, w, h}, we take the end endpoint as the caret position.
                    rects[0] := rects[0] + rects[2]
                    rects[2] := 0
                    goto end
                }
            }
            return false
end:
            left := Round(rects[0])
            top := Round(rects[1])
            right := left + Round(rects[2])
            bottom := top + Round(rects[3])
            return true
        }
        return false
    }

    getCaretPosFromWpfCaret() {
        try {
            uia := ComObject("{E22AD333-B25F-460C-83D0-0581107395C9}", "{30CBE57D-D9D0-452A-AB13-7AC5AC4825EE}")
            ComCall(8, uia, "ptr*", focusedEle := ComValue(13, 0)) ; uia->GetFocusedElement(&focusedEle);
            if !focusedEle.Ptr
                return false

            ComCall(20, uia, "ptr*", cacheRequest := ComValue(13, 0)) ; uia->CreateCacheRequest(&cacheRequest);
            if !cacheRequest.Ptr
                return false

            ComCall(17, uia, "ptr*", rawViewCondition := ComValue(13, 0)) ; uia->get_RawViewCondition(&rawViewCondition);
            if !rawViewCondition.Ptr
                return false

            ComCall(9, cacheRequest, "ptr", rawViewCondition) ; cacheRequest->put_TreeFilter(rawViewCondition);
            ComCall(3, cacheRequest, "int", 30001) ; cacheRequest->AddProperty(UIA_BoundingRectanglePropertyId);

            var := Buffer(24, 0)
            ref := ComValue(0x400C, var.Ptr)
            ref[] := ComValue(8, "WpfCaret")
            ComCall(23, uia, "int", 30012, "ptr", var, "ptr*", condition := ComValue(13, 0)) ; uia->CreatePropertyCondition(UIA_ClassNamePropertyId, CComVariant(L"WpfCaret"), &classNameCondition);
            if !condition.Ptr
                return false

            ComCall(7, focusedEle, "int", 4, "ptr", condition, "ptr", cacheRequest, "ptr*", wpfCaret := ComValue(13, 0)) ; focusedEle->FindFirstBuildCache(TreeScope_Descendants, condition, cacheRequest, &wpfCaret);
            if !wpfCaret.Ptr
                return false

            ComCall(75, wpfCaret, "ptr", rect := Buffer(16)) ; wpfCaret->get_CachedBoundingRectangle(&rect);
            getRect(rect, &left, &top, &right, &bottom)
            return true
        }
        return false
    }

    getCaretPosFromHook() {
        static WM_GET_CARET_POS := DllCall("RegisterWindowMessageW", "str", "WM_GET_CARET_POS", "uint")
        if !tid := DllCall("GetWindowThreadProcessId", "ptr", hwnd, "ptr*", &pid := 0, "uint")
            return false
        ; Update caret position
        try {
            SendMessage(0x010f, 0, 0, hwnd) ; WM_IME_COMPOSITION
        }
        ; PROCESS_CREATE_THREAD | PROCESS_QUERY_INFORMATION | PROCESS_VM_OPERATION | PROCESS_VM_WRITE | PROCESS_VM_READ
        if !hProcess := DllCall("OpenProcess", "uint", 1082, "int", false, "uint", pid, "ptr")
            return false
        hProcess := { Ptr: hProcess, __Delete: (_) => DllCall("CloseHandle", "ptr", _) }

        isX64 := isX64Process(hProcess)
        if isX64 && A_PtrSize == 4
            return false
        if !moduleBaseMap := getModulesBases(hProcess, ["kernel32.dll", "user32.dll", "combase.dll"])
            return false
        if isX64 {
            static shellcode64 := compile(true)
            shellcode := shellcode64
        }
        else {
            static shellcode32 := compile(false)
            shellcode := shellcode32
        }
        if !mem := DllCall("VirtualAllocEx", "ptr", hProcess, "ptr", 0, "ptr", shellcode.Size, "uint", 0x1000, "uint", 0x40, "ptr")
            return false
        mem := { Ptr: mem, __Delete: (_) => DllCall("VirtualFreeEx", "ptr", hProcess, "ptr", _, "uptr", 0, "uint", 0x8000) }
        link(isX64, shellcode, mem.Ptr, moduleBaseMap["user32.dll"], moduleBaseMap["combase.dll"], hwnd, tid, WM_GET_CARET_POS, &pThreadProc, &pRect)

        if !DllCall("WriteProcessMemory", "ptr", hProcess, "ptr", mem, "ptr", shellcode, "uptr", shellcode.Size, "ptr", 0)
            return false
        DllCall("FlushInstructionCache", "ptr", hProcess, "ptr", mem, "uptr", shellcode.Size)

        if !hThread := DllCall("CreateRemoteThread", "ptr", hProcess, "ptr", 0, "uptr", 0, "ptr", pThreadProc, "ptr", mem, "uint", 0, "uint*", &remoteTid := 0, "ptr")
            return false
        hThread := { Ptr: hThread, __Delete: (_) => DllCall("CloseHandle", "ptr", _) }

        if msgWaitForSingleObject(hThread)
            return false
        if !DllCall("GetExitCodeThread", "ptr", hThread, "uint*", exitCode := 0) || exitCode !== 0
            return false

        rect := Buffer(16)
        if !DllCall("ReadProcessMemory", "ptr", hProcess, "ptr", pRect, "ptr", rect, "uptr", rect.Size, "uptr*", &bytesRead := 0) || bytesRead !== rect.Size
            return false
        getRect(rect, &left, &top, &right, &bottom)
        scaleRect(getWindowScale(hwnd), &left, &top, &right, &bottom)
        return true

        static isX64Process(hProcess) {
            DllCall("IsWow64Process", "ptr", hProcess, "int*", &isWow64 := 0)
            if isWow64
                return false
            if A_PtrSize == 8
                return true
            DllCall("IsWow64Process", "ptr", DllCall("GetCurrentProcess", "ptr"), "int*", &isWow64)
            return isWow64
        }

        static getModulesBases(hProcess, modules) {
            hModules := Buffer(A_PtrSize * 350)
            if !DllCall("K32EnumProcessModulesEx", "ptr", hProcess, "ptr", hModules, "uint", hModules.Size, "uint*", &needed := 0, "uint", 3)
                return
            moduleBaseMap := Map()
            moduleBaseMap.CaseSense := false
            for v in modules
                moduleBaseMap[v] := 0
            cnt := modules.Length
            loop Min(350, needed) {
                hModule := NumGet(hModules, A_PtrSize * (A_Index - 1), "ptr")
                VarSetStrCapacity(&name, 12)
                if DllCall("K32GetModuleBaseNameW", "ptr", hProcess, "ptr", hModule, "str", &name, "uint", 13) {
                    if moduleBaseMap.Has(name) {
                        moduleInfo := Buffer(24)
                        if !DllCall("K32GetModuleInformation", "ptr", hProcess, "ptr", hModule, "ptr", moduleInfo, "uint", moduleInfo.Size)
                            return
                        if !base := NumGet(moduleInfo, "ptr")
                            return
                        moduleBaseMap[name] := base
                        cnt--
                    }
                }
            } until cnt == 0
            if cnt == 0
                return moduleBaseMap
        }

        static compile(x64) {
            if x64
                shellcodeBase64 := "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABrnppSh2UjT6uenH1oPjxQAeiAqiEg0hGT4ABgsGe4blNldFdpbmRvd3NIb29rRXhXAAAAVW5ob29rV2luZG93c0hvb2tFeABDYWxsTmV4dEhvb2tFeAAAAAAAAFNlbmRNZXNzYWdlVGltZW91dFcAQ29DcmVhdGVJbnN0YW5jZQAAAAAAAAAASIlcJAhIiXQkEFdIg+wgSYvYSIvyi/mFyXgjSIXbdB6LBQb///9BOUAQdRJIjQ3d/v//6JgBAACJBfL+//9Iiw3L/v//SI0VdP///+jnAgAASIXAdRBIi1wkMEiLdCQ4SIPEIF/DTIvLTIvGi9czyUiLXCQwSIt0JDhIg8QgX0j/4MzMzMzMzDPAw8zMzMzMQFNWSIPsSIvySIvZSIXJdQy4VwAHgEiDxEheW8NIi0kISI1UJGBIiVQkKEG4/////0iNVCQwSIl8JEAz/0iJVCQgiXwkYIvWSIsBRI1PAf9QKIXAeHJIOXwkMHRrOXwkYHRlSItLCEiNVCR4SIl8JHhIiwH/UEiL+IXAeDJIi0wkeEiFyXQoSIsBSI1UJHBMi0QkMEyNSxBIiVQkIIvW/1AgSItMJHiL+EiLAf9QEEiLTCQwSIsB/1AQi8dIi3wkQEiDxEheW8NIi3wkQLgBAAAASIPESF5bw8zMzMzMzMxIhcl0VEiF0nRPTYXAdEpIiwJIhcB1HUi4wAAAAAAAAEZIOUIIdCxJxwAAAAAAuAJAAIDDSbkD6ICqISDSEUk7wXXkSLiT4ABgsGe4bkg5Qgh11EmJCDPAw7hXAAeAw8xAU0iD7EBIi9lIjZHYAAAASItJCOhPAQAASIXAdQu4AQAAAEiDxEBbwzPJx0QkWAEAAABIjVQkaEiJTCRoSIlUJCBMjUt4M9JIiUwkYEiJTCQwiUwkUEiNS2hEjUIX/9CFwA+I7wAAAEiLTCRoSIXJD4ThAAAASIsBSI1UJFD/UBiFwA+IhQAAAEiLTCRoSI1UJGBIiwH/UDiFwHhxSItMJGBIhcl0bEiLAUiNVCQw/1AwhcB4WEiLTCQwSIXJdGZIjUNISIlLMEiJQyhMjUMoSI0Vyf7//0G5AwAAAEiJEEiNBdH9//9IiUNQSI1UJFhIiUNYSI0Fxf3//0iJQ2BIiwFIiVQkIItUJFD/UBhIi0wkYEiLVCQwSIXSdA5IiwJIi8r/UBBIi0wkYEiFyXQGSIsB/1AQSItMJGhIhcl0BkiLAf9QEItEJFj32BvAg+AESIPEQFvDuAQAAABIg8RAW8PMzMzMzMxIiVwkCEiJbCQQSIl0JBhIiXwkIEyL2kyL0UiFyXRwSIXSdGtIY0E8g7wIjAAAAAB0XYuMCIgAAACFyXRSRYtMCiBJjQQKi3AkTQPKi2gcSQPyi3gYSQPqD7YaRTPA/89BixFJA9I6GnUZD7bLSYvDSSvThMl0Lw+2SAFI/8A6DAJ08EH/wEmDwQREO8d20TPASItcJAhIi2wkEEiLdCQYSIt8JCDDSWPAD7cMRotEjQBJA8Lr28zMSIlcJAhIiWwkEEiJdCQYSIl8JCBBVkiD7EBIixlIjZGIAAAASIv5SIvL6Bn///9IjZfEAAAASIvLSIvw6Af///9IjZecAAAASIvLSIvo6PX+//9Mi/BIhfZ0ZUiF7XRgSIXAdFtEi08YSI0VoPv//0UzwEGNSAT/1kiL8EiFwHUFjUYC6z+LVxwzwEiLTxBFM8lIiUQkMEUzwMdEJCjIAAAAiUQkIP/VSIvOSIvYQf/WSIXbdQWNQwPrCotHIOsFuAEAAABIi1wkUEiLbCRYSIt0JGBIi3wkaEiDxEBBXsM="
            else
                shellcodeBase64 := "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGuemlKHZSNPq56cfWg+PFAB6ICqISDSEZPgAGCwZ7huU2V0V2luZG93c0hvb2tFeFcAAABVbmhvb2tXaW5kb3dzSG9va0V4AENhbGxOZXh0SG9va0V4AAAAAAAAU2VuZE1lc3NhZ2VUaW1lb3V0VwBDb0NyZWF0ZUluc3RhbmNlAAAAAFZX6MkCAACDfCQMAIvwi3wkFHwYhf90FItPCDtOEHUMVuhqAQAAg8QEiUYUjYaIAAAAUP826J4CAACDxAiFwHUFX17CDABX/3QkFP90JBRqAP/QX17CDAAzwMIEAMzMzIPsFFaLdCQchfZ1DLhXAAeAXoPEFMIIAItOBI1UJARSjVQkEMdEJAgAAAAAUosBagFq//90JDBR/1AUhcB4bIN8JAwAdGWDfCQEAHRei04EjVQkHFfHRCQgAAAAAFKLAVH/UCSL+IX/eC2LVCQghdJ0JYsCi0gQjUQkDFCNRghQ/3QkGP90JDBS/9GL+ItEJCBQiwj/UQiLRCQQUIsI/1EIi8dfXoPEFMIIALgBAAAAXoPEFMIIAMyLTCQIVot0JAiF9nRfhcl0W4tUJBCF0nRTiwELQQR1IYF5CMAAAAB1CYF5DAAAAEZ0MscCAAAAALgCQACAXsIMAIE5A+iAqnXpgXkEISDSEXXggXkIk+AAYHXXgXkMsGe4bnXOiTIzwF7CDAC4VwAHgF7CDADMzMyD7BBWi3QkGI2GsAAAAFD/dgToMQEAAIvIg8QIhcl1CI1BAV6DxBDDjUQkBMdEJAQAAAAAUI1GUMdEJBwAAAAAUGoXagCNRkDHRCQYAAAAAFDHRCQgAAAAAMdEJCQBAAAA/9GFwA+IywAAAItMJASFyQ+EvwAAAIsBjVQkDFdSUf9QDIXAeHCLTCQIjVQkHFJRiwH/UByFwHhdi0wkHIXJdFmLAY1UJAxSUf9QGIXAeEaLfCQMhf90UI1OMIl+HLjcAQAAiU4YA8aNVhiJAYvGBRwBAACNTCQUUYlGNIlGOLgkAQAAagMDxlL/dCQciUY8iwdX/1AMi0wkHItUJAyF0nQKiwJS/1AIi0wkHF+FyXQGiwFR/1AIi0wkBIXJdAaLAVH/UAiLRCQQ99heG8CD4ASDxBDDuAQAAABeg8QQw7gAAAAAw8zMg+wIU1VWV4t8JByF/w+EgQAAAItcJCCF23R5i0c8g3w4fAB0b4tEOHiFwHRni0w4JDP2i1Q4IAPPi2w4GAPXiUwkEItMOBwDz4lUJByJTCQUTYorixSyA9c6KnUTis2LwyvThMl0FIpIAUA6DAJ080Y79Xcfi1QkHOvZi0QkEItMJBQPtwRwiwSBA8dfXl1bg8QIw19eXTPAW4PECMPMzFNVVleLfCQUizeNR2BQVuhM////iUQkHI2HnAAAAFBW6Dv///+L2I1HdFBW6C////+LTCQsg8QYi+iFyXRshdt0aIXtdGSLxwWUAwAAiXgBuMQAAAD/dwwDx2oAUGoE/9GJRCQUhcB1DF9eXbgCAAAAW8IEAGoAaMgAAABqAGoAagD/dxD/dwj/0/90JBSL8P/VhfZ1Cl+NRgNeXVvCBACLRxRfXl1bwgQAX15duAEAAABbwgQA"
            len := StrLen(shellcodeBase64)
            shellcode := Buffer(len * 0.75)
            if !DllCall("crypt32\CryptStringToBinary", "str", shellcodeBase64, "uint", len, "uint", 1, "ptr", shellcode, "uint*", shellcode.Size, "ptr", 0, "ptr", 0)
                return
            return shellcode
        }

        static link(x64, shellcode, shellcodeBase, user32Base, combaseBase, hwnd, tid, msg, &pThreadProc, &pRect) {
            if x64 {
                NumPut("uint64", user32Base, shellcode, 0)
                NumPut("uint64", combaseBase, shellcode, 8)
                NumPut("uint64", hwnd, shellcode, 16)
                NumPut("uint", tid, shellcode, 24)
                NumPut("uint", msg, shellcode, 28)
                pThreadProc := shellcodeBase + 0x4e0
                pRect := shellcodeBase + 56
            }
            else {
                NumPut("uint", user32Base, shellcode, 0)
                NumPut("uint", combaseBase, shellcode, 4)
                NumPut("uint", hwnd, shellcode, 8)
                NumPut("uint", tid, shellcode, 12)
                NumPut("uint", msg, shellcode, 16)
                pThreadProc := shellcodeBase + 0x43c
                pRect := shellcodeBase + 32
            }
        }

        static msgWaitForSingleObject(handle) {
            while 1 == res := DllCall("MsgWaitForMultipleObjects", "uint", 1, "ptr*", handle, "int", false, "uint", -1, "uint", 7423) { ; QS_ALLINPUT := 7423
                msg := Buffer(A_PtrSize == 8 ? 48 : 28)
                while DllCall("PeekMessageW", "ptr", msg, "ptr", 0, "uint", 0, "uint", 0, "uint", 1) { ; PM_REMOVE := 1
                    DllCall("TranslateMessage", "ptr", msg)
                    DllCall("DispatchMessageW", "ptr", msg)
                }
            }
            return res
        }
    }

    static guidFromString(str) {
        DllCall("ole32\CLSIDFromString", "str", str, "ptr", buf := Buffer(16), "hresult")
        return buf
    }

    static getRect(buf, &left, &top, &right, &bottom) {
        left := NumGet(buf, 0, "int")
        top := NumGet(buf, 4, "int")
        right := NumGet(buf, 8, "int")
        bottom := NumGet(buf, 12, "int")
    }

    static getWindowScale(hwnd) {
        if winDpi := DllCall("GetDpiForWindow", "ptr", hwnd, "uint")
            return A_ScreenDPI / winDpi
        return 1
    }

    static scaleRect(scale, &left, &top, &right, &bottom) {
        left := Round(left * scale)
        top := Round(top * scale)
        right := Round(right * scale)
        bottom := Round(bottom * scale)
    }

    static clientToScreenRect(hwnd, &left, &top, &right, &bottom) {
        w := right - left
        h := bottom - top
        pt := left | top << 32
        DllCall("ClientToScreen", "ptr", hwnd, "int64*", &pt)
        left := pt & 0xffffffff
        top := pt >> 32
        right := left + w
        bottom := top + h
    }
}

; 激活跳转模式
ActivateJumpMode(mode) {
    global jumpMode, jumpActive, jumpBuffer, jumpPosition, g_inputHook
    
    ; 确保所有坐标模式一致为"Screen"，避免窗口切换导致的坐标偏移
    CoordMode "Mouse", "Screen"
    CoordMode "Caret", "Screen" 
    CoordMode "ToolTip", "Screen"
    
    ; 检查是否是相同的模式，如果是则取消跳转
    if (jumpActive && jumpMode == mode) {
        DeactivateJumpMode()
        return
    }
    
    ; 如果已经激活了不同的模式，则需要先停止当前的输入钩子
    if (jumpActive) {
        ; 停止当前输入钩子
        StopInputHook()
        
        ; 更新模式和清空缓冲区
        jumpMode := mode
        jumpBuffer := ""
    } else {
        ; 如果没有激活任何模式，进行常规的初始化
        jumpMode := mode
        jumpActive := true
        jumpBuffer := ""
    }
    
    ; 获取光标位置
    caretPos := GetCaretPosition()
    
    ; 如果无法获取有效的光标位置，取消跳转
    if (!caretPos) {
        DeactivateJumpMode()
        return
    }
    
    ; 保存光标位置用于显示ToolTip
    jumpPosition.x := caretPos.x
    jumpPosition.y := caretPos.y
    
    ; 显示初始提示
    UpdateJumpTooltip()
    
    ; 不管是首次激活还是切换模式，都重新创建并启动输入钩子
    StartInputHook()
}

; 更新跳转提示
UpdateJumpTooltip() {
    global jumpMode, jumpBuffer, jumpPosition
    
    ; 根据模式设置提示文本
    modeText := ""
    modeDefaultDir := "" ; 默认方向提示
    
    if (jumpMode = "up") {
        modeText := "移动"
        modeDefaultDir := "↑"
    }
    else if (jumpMode = "down") {
        modeText := "移动"
        modeDefaultDir := "↓"
    }
    else if (jumpMode = "left") {
        modeText := "移动"
        modeDefaultDir := "←"
    }
    else if (jumpMode = "right") {
        modeText := "移动"
        modeDefaultDir := "→"
    }
    else if (jumpMode = "selectUp") {
        modeText := "选择"
        modeDefaultDir := "↑"
    }
    else if (jumpMode = "selectDown") {
        modeText := "选择"
        modeDefaultDir := "↓"
    }
    else if (jumpMode = "selectLeft") {
        modeText := "选择"
        modeDefaultDir := "←"
    }
    else if (jumpMode = "selectRight") {
        modeText := "选择"
        modeDefaultDir := "→"
    }
    else if (jumpMode = "deleteLeft") {
        modeText := "删除"
        modeDefaultDir := "←"
    }
    else if (jumpMode = "deleteRight") {
        modeText := "删除"
        modeDefaultDir := "→"
    }
    
    ; 显示当前输入
    inputText := jumpBuffer ? jumpBuffer : ""
    
    ; 判断是否有负号
    dirSymbol := modeDefaultDir
    if (inputText && SubStr(inputText, 1, 1) == "-") {
        ; 如果有负号，显示方向已改变
        Switch modeDefaultDir {
            case "↑": dirSymbol := "↓"
            case "↓": dirSymbol := "↑"
            case "←": dirSymbol := "→"
            case "→": dirSymbol := "←"
        }
    }
    
    ; 组合提示文本（添加操作提示）
    tooltipText := ""

    ; 如果是水平方向的跳转或删除且启用了单词级跳转，添加"(按单词)"前缀
    if (isWordJump && (jumpMode == "left" || jumpMode == "right" || 
        jumpMode == "selectLeft" || jumpMode == "selectRight" ||
        jumpMode == "deleteLeft" || jumpMode == "deleteRight")) {
        tooltipText .= "(按单词) "
    }
    
    ; 添加原有的模式文本和方向
    tooltipText .= modeText . " " . dirSymbol . " " . inputText . " (Esc退出)"
    
    ; 显示提示在光标位置右上角
    CoordMode("ToolTip", "Screen")
    ToolTip(tooltipText, jumpPosition.x + 15, jumpPosition.y - 30)
}

; 取消跳转模式
DeactivateJumpMode() {
    global jumpActive, g_inputHook, jumpBuffer, isWordJump
    
    ; 停止输入钩子
    StopInputHook()
    
    ; 清除提示
    ToolTip()
    
    ; 重置状态和缓冲区
    jumpActive := false
    jumpBuffer := ""
    isWordJump := false  ; 重置单词跳转标志
}

; 执行跳转操作
ExecuteJump() {
    global jumpMode, jumpBuffer
    
    ; 如果输入为空，不执行任何操作
    if (jumpBuffer == "") {
        DeactivateJumpMode()
        return
    }
    
    ; 转换为数字
    jumpCount := 0
    Try {
        ; 检查输入的第一个字符是否是负号
        isNegative := false
        if (SubStr(jumpBuffer, 1, 1) == "-") {
            isNegative := true
            jumpBuffer := SubStr(jumpBuffer, 2)
        }
        
        ; 转换数字部分
        if (jumpBuffer == "")
            jumpCount := 1 ; 如果只输入了负号，假设是1
        else
            jumpCount := Integer(jumpBuffer)
        
        ; 根据模式设置正负
        Switch jumpMode {
            case "up", "left", "selectUp", "selectLeft":
                ; 这些模式默认为负数（向上/向左），负号会反转为正数
                jumpCount := isNegative ? jumpCount : -jumpCount
            case "down", "right", "selectDown", "selectRight":
                ; 这些模式默认为正数（向下/向右），负号会反转为负数
                jumpCount := isNegative ? -jumpCount : jumpCount
        }
    } Catch {
        ; 转换失败，取消操作
        DeactivateJumpMode()
        return
    }
    
    ; 如果输入为0，不执行操作
    if (jumpCount == 0) {
        DeactivateJumpMode()
        return
    }
    
    ; 执行对应操作
    absCount := Abs(jumpCount)

    Switch jumpMode {
        case "up", "down":
            if (jumpCount > 0) {
                MoveDown(absCount)  ; 正数向下
            } else {
                MoveUp(absCount)    ; 负数向上
            }
        case "left", "right":
            if (isWordJump) {
                ; 单词级跳转
                if (jumpCount > 0) {
                    MoveWordRight(absCount) ; 正数向右
                } else {
                    MoveWordLeft(absCount)  ; 负数向左
                }
            } else {
                ; 普通字符级跳转
                if (jumpCount > 0) {
                    MoveRight(absCount) ; 正数向右
                } else {
                    MoveLeft(absCount)  ; 负数向左
                }
            }
        case "selectUp", "selectDown":
            if (jumpCount > 0) {
                SelectDown(absCount) ; 正数向下选择
            } else {
                SelectUp(absCount)   ; 负数向上选择
            }
        case "selectLeft", "selectRight":
            if (isWordJump) {
                ; 单词级选择
                if (jumpCount > 0) {
                    SelectWordRight(absCount) ; 正数向右选择
                } else {
                    SelectWordLeft(absCount)  ; 负数向左选择
                }
            } else {
                ; 普通字符级选择
                if (jumpCount > 0) {
                    SelectRight(absCount) ; 正数向右选择
                } else {
                    SelectLeft(absCount)  ; 负数向左选择
                }
            }
        case "deleteLeft", "deleteRight":
            if (isWordJump) {
                ; 单词级删除
                if (jumpMode = "deleteLeft") {
                    Loop absCount {
                        DeleteWord()  ; 向左删除单词
                    }
                } else {
                    Loop absCount {
                        ForwardDeleteWord() ; 向右删除单词
                    }
                }
            } else {
                ; 字符级删除
                if (jumpMode = "deleteLeft") {
                    DeleteLeft(absCount)  ; 向左删除指定数量字符
                } else {
                    DeleteRight(absCount) ; 向右删除指定数量字符
                }
            }
    }
    
    ; 完成操作后退出跳转模式
    DeactivateJumpMode()
}

; 启动输入钩子
StartInputHook() {
    global g_inputHook, jumpBuffer
    
    ; 创建新的InputHook对象
    g_inputHook := InputHook("L0", "{Enter}")
    ; 设置输入结束键
    g_inputHook.EndKeys := "{Enter}"
    ; 设置可接受的按键（只允许数字和负号）
    g_inputHook.KeyOpt("{1}{2}{3}{4}{5}{6}{7}{8}{9}{0}{-}{Numpad1}{Numpad2}{Numpad3}{Numpad4}{Numpad5}{Numpad6}{Numpad7}{Numpad8}{Numpad9}{Numpad0}{NumpadSub}", "N")
    
    ; 设置字符输入回调
    g_inputHook.OnChar := OnChar
    ; 设置结束回调
    g_inputHook.OnEnd := OnInputEnd
    
    ; 启动输入钩子
    try {
        g_inputHook.Start()
    } catch as e {
        ToolTip("InputHook启动失败: " e.Message)
        SetTimer () => ToolTip(), -2000
    }
}

; 停止输入钩子
StopInputHook() {
    global g_inputHook
    
    ; 如果输入钩子有效，停止它
    if (IsObject(g_inputHook) && g_inputHook.HasProp("InProgress") && g_inputHook.InProgress)
        g_inputHook.Stop()
}

; 键按下回调函数
OnKeyDown(ih, vk, sc) {
    global jumpBuffer
    
    try {
        ; 获取按键名称
        keyName := GetKeyName(Format("vk{:x}sc{:x}", vk, sc))
        
        ; 处理退格键 - 删除最后一个字符
        if (keyName == "Backspace") {
            if (jumpBuffer != "") {
                jumpBuffer := SubStr(jumpBuffer, 1, StrLen(jumpBuffer) - 1)
                UpdateJumpTooltip()
            }
            return 1  ; 返回1表示拦截此按键，不让它传递到应用程序
        }
        
        ; 处理回车键 - 执行跳转
        if (keyName == "Enter") {
            ExecuteJump()
            return 1  ; 返回1表示拦截此按键
        }
        
        ; 处理Esc键 - 取消操作
        if (keyName == "Escape") {
            DeactivateJumpMode()
            return 1  ; 返回1表示拦截此按键
        }
    } catch as e {
        ; 捕获错误但不做任何处理，避免中断跳转功能
    }
    
    ; 默认情况下也拦截所有按键
    return 1
}

; 字符输入回调函数
OnChar(ih, char) {
    global jumpBuffer, isWordJump
    
    try {
        ; 处理0作为单词级跳转触发器
        if (char == "0" && jumpBuffer == "") {
            ; 只有在缓冲区为空时，0才表示单词级跳转
            isWordJump := true
            UpdateJumpTooltip()
            return 1
        }
        
        ; 只接受数字和负号
        if (char >= "0" && char <= "9" || char == "-") {
            ; 添加到缓冲区
            jumpBuffer .= char
            UpdateJumpTooltip()
        }
    } catch as e {
        ; 捕获错误但不做任何处理
    }
    
    ; 拦截所有字符输入
    return 1
}

; 键释放回调函数
OnKeyUp(ih, vk, sc) {
    ; 这里可以添加键释放时的操作
}

; 输入结束回调函数
OnInputEnd(ih) {
    global jumpActive
    
    if (jumpActive) {
        try {
            ; 根据结束原因处理
            if (ih.EndReason == "EndKey") {
                ; 如果是因为按下了Enter，执行跳转
                if (ih.EndKey == "Enter")
                    ExecuteJump()
                ; 如果是因为按下了Escape，取消操作
                else if (ih.EndKey == "Escape")
                    DeactivateJumpMode()
            } else {
                ; 其他原因结束，取消操作
                DeactivateJumpMode()
            }
        } catch as e {
            ; 出错也取消操作
            DeactivateJumpMode()
        }
    }
}
; 使用#HotIf指令为跳转模式定义热键
#HotIf jumpActive

; Escape键, Delete键或空格键 - 退出跳转模式
Escape::DeactivateJumpMode()
Delete::DeactivateJumpMode()
CapsLock::DeactivateJumpMode()
~LButton::DeactivateJumpMode()
~RButton::DeactivateJumpMode()
~MButton::DeactivateJumpMode()

; Backspace键 - 删除最后一个字符，如果为空则退出
Backspace:: {
    global jumpBuffer, isWordJump
    
    if (jumpBuffer != "") {
        ; 如果缓冲区不为空，删除最后一个字符
        jumpBuffer := SubStr(jumpBuffer, 1, StrLen(jumpBuffer) - 1)
        UpdateJumpTooltip()
    } else if (isWordJump) {
        ; 如果缓冲区为空但处于单词模式，退回到字符级跳转模式
        isWordJump := false
        UpdateJumpTooltip()
    } else {
        ; 缓冲区为空且不在单词模式，退出跳转模式
        DeactivateJumpMode()
    }
}

; Enter键 - 执行跳转
Enter::ExecuteJump()

#HotIf

#HotIf GetKeyState("CapsLock", "P") && !GetKeyState("Alt", "P") && jumpActive
s::Send("")
f::Send("")
e::Send("")
d::Send("")
j::Send("")
l::Send("")
k::Send("")
i::Send("")
,::Send("")
.::Send("")
w::Send("")
r::Send("")
t::Send("")
y::Send("")
u::Send("")
o::Send("")
p::Send("")
a::Send("")
g::Send("")
h::Send("")
z::Send("")
x::Send("")
c::Send("")
v::Send("")
b::Send("")
n::Send("")
m::Send("")
/::Send("")
`;::Send("")
'::Send("")
q::Send("")
[::Send("")
]::Send("")
\::Send("")
#HotIf

; =====================================================================
; 鼠标控制功能
; =====================================================================

; 鼠标移动相关全局变量
global mousePrecisionMode := false     ; 是否处于精确移动模式
global mousePrecisionFactor := 0.2     ; 精确模式下的移动距离缩放因子(0.1-1.0)
global mouseKeysPressed := Map()       ; 存储当前按下的方向键
global mouseMovementActive := false    ; 指示是否有鼠标移动正在进行
global mouseMovementTimer := 0         ; 鼠标移动定时器ID

; 鼠标移动函数
MouseMoveRelative(x, y) {
    static mouseSpeed := 1
    global mousePrecisionMode, mousePrecisionFactor
    
    ; 使鼠标速度随按键按下时间增加
    mouseSpeed += 0.3  ; 降低加速度，原来是0.5
    
    ; 限制最大速度
    if (mouseSpeed > 6)  ; 降低最大速度，原来是10
        mouseSpeed := 6
    
    ; 移动距离根据速度调整
    moveX := x * mouseSpeed
    moveY := y * mouseSpeed
    
    ; 如果处于精确移动模式，则缩小移动距离
    if (mousePrecisionMode) {
        moveX := moveX * mousePrecisionFactor
        moveY := moveY * mousePrecisionFactor
    }
    
    ; 执行相对移动
    MouseMove(moveX, moveY, 0, "R")
}

; 重置鼠标加速度
ResetMouseSpeed() {
    static mouseSpeed := 1
    mouseSpeed := 1
}

; 计算当前按下的方向键组合的移动向量
; 注意：这个函数支持同时按下多个方向键，实现斜向移动
; 例如：同时按下Up和Right键会产生右上方向的斜向移动
CalculateMovementVector() {
    global mouseKeysPressed
    
    ; 初始化移动向量
    moveX := 0
    moveY := 0
    
    ; 根据按下的键计算移动向量
    if (mouseKeysPressed.Has("Up") && mouseKeysPressed["Up"])
        moveY -= 7  ; 向上移动
    if (mouseKeysPressed.Has("Down") && mouseKeysPressed["Down"])
        moveY += 7  ; 向下移动
    if (mouseKeysPressed.Has("Left") && mouseKeysPressed["Left"])
        moveX -= 7  ; 向左移动
    if (mouseKeysPressed.Has("Right") && mouseKeysPressed["Right"])
        moveX += 7  ; 向右移动
    
    ; 当同时按下两个相邻方向键时，会自动产生斜向移动
    ; 例如：
    ; - 同时按下Up和Right：moveX=7, moveY=-7，鼠标向右上方移动
    ; - 同时按下Down和Left：moveX=-7, moveY=7，鼠标向左下方移动
    
    return {x: moveX, y: moveY}
}

; 开始连续鼠标移动
StartContinuousMouseMovement() {
    global mouseMovementActive, mouseMovementTimer
    
    ; 如果已经有移动在进行，不需要再次启动
    if (mouseMovementActive)
        return
    
    ; 标记移动为活动状态
    mouseMovementActive := true
    
    ; 创建定时器，每16ms（约60fps）执行一次移动
    mouseMovementTimer := SetTimer(ContinuousMouseMove, 16)
}

; 停止连续鼠标移动
StopContinuousMouseMovement() {
    global mouseMovementActive, mouseMovementTimer, mouseKeysPressed
    
    ; 如果没有移动在进行，不需要操作
    if (!mouseMovementActive)
        return
    
    ; 停止定时器
    if (mouseMovementTimer)
        SetTimer(mouseMovementTimer, 0)
    
    ; 重置状态
    mouseMovementActive := false
    mouseMovementTimer := 0
    
    ; 重置鼠标加速度
    ResetMouseSpeed()
}

; 连续鼠标移动函数 - 由定时器调用
ContinuousMouseMove() {
    global mouseKeysPressed
    
    ; 计算当前移动向量
    movement := CalculateMovementVector()
    
    ; 如果没有方向键被按下，停止移动
    if (movement.x == 0 && movement.y == 0) {
        StopContinuousMouseMovement()
        return
    }
    
    ; 执行移动
    MouseMoveRelative(movement.x, movement.y)
}

; 鼠标按键按下函数
MouseLeftDown() {
    Click("Left Down")
}

MouseLeftUp() {
    Click("Left Up")
}

MouseRightDown() {
    Click("Right Down")
}

MouseRightUp() {
    Click("Right Up")
}

MouseWheelUp(repeat := 1) {
    Loop repeat
        Click("WheelUp")
}

MouseWheelDown(repeat := 1) {
    Loop repeat
        Click("WheelDown")
}

; =====================================================================
; 热键映射 - 文本编辑操作
; =====================================================================

; 光标移动键映射
#HotIf GetKeyState("CapsLock", "P")
a::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        MoveToPageBeginning()
    } else {
        MoveWordLeft()
    }
}

s::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        ; 需要声明要使用的全局变量
        global jumpActive, jumpPosition, jumpBuffer, jumpMode
        
        if(jumpActive) {
            ; 保存当前位置
            prevPosition := jumpPosition
            
            ; 关闭当前模式但保持ToolTip显示
            CoordMode("ToolTip", "Screen")
            ; ToolTip("移动 ←  (Esc退出)", prevPosition.x + 15, prevPosition.y - 30)
            
            ; 关闭输入钩子但不清除ToolTip
            StopInputHook()
            jumpActive := false
            jumpBuffer := ""
            
            ; 激活新模式
            SetTimer () => ActivateJumpMode("left"), -50
        } else {
            ActivateJumpMode("left")
        }
    } else {
        MoveLeft()
    }
}

d::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        ; 需要声明要使用的全局变量
        global jumpActive, jumpPosition, jumpBuffer, jumpMode
        
        if(jumpActive) {
            ; 保存当前位置
            prevPosition := jumpPosition
            
            ; 关闭当前模式但保持ToolTip显示
            CoordMode("ToolTip", "Screen")
            ; ToolTip("移动 ↓  (Esc退出)", prevPosition.x + 15, prevPosition.y - 30)
            
            ; 关闭输入钩子但不清除ToolTip
            StopInputHook()
            jumpActive := false
            jumpBuffer := ""
            
            ; 激活新模式
            SetTimer () => ActivateJumpMode("down"), -50
        } else {
            ActivateJumpMode("down")
        }
    } else {
        MoveDown()
    }
}

e::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        ; 需要声明要使用的全局变量
        global jumpActive, jumpPosition, jumpBuffer, jumpMode
        
        if(jumpActive) {
            ; 保存当前位置
            prevPosition := jumpPosition
            
            ; 关闭当前模式但保持ToolTip显示
            CoordMode("ToolTip", "Screen")
            ; ToolTip("移动 ↑  (Esc退出)", prevPosition.x + 15, prevPosition.y - 30)
            
            ; 关闭输入钩子但不清除ToolTip
            StopInputHook()
            jumpActive := false
            jumpBuffer := ""

            ; 激活新模式
            SetTimer () => ActivateJumpMode("up"), -50
        } else {
            ActivateJumpMode("up")
        }
    } else {
        MoveUp()
    }
}

f::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        ; 需要声明要使用的全局变量
        global jumpActive, jumpPosition, jumpBuffer, jumpMode
        
        if(jumpActive) {
            ; 保存当前位置
            prevPosition := jumpPosition
            
            ; 关闭当前模式但保持ToolTip显示
            CoordMode("ToolTip", "Screen")
            ; ToolTip("移动 →  (Esc退出)", prevPosition.x + 15, prevPosition.y - 30)
            
            ; 关闭输入钩子但不清除ToolTip
            StopInputHook()
            jumpActive := false
            jumpBuffer := ""
            
            ; 激活新模式
            SetTimer () => ActivateJumpMode("right"), -50
        } else {
            ActivateJumpMode("right")
        }
    } else {
        MoveRight()
    }
}

g::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        MoveToPageEnd()
    } else {
        MoveWordRight()
    }
}

w::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    MoveHome()
}

r::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    MoveEnd()
}
#HotIf

; 文本选择键映射
#HotIf GetKeyState("CapsLock", "P")
h::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        SelectToPageBeginning()
    } else {
        SelectWordLeft()
    }
}

j::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        ; 需要声明要使用的全局变量
        global jumpActive, jumpPosition, jumpBuffer, jumpMode
        
        if(jumpActive) {
            ; 保存当前位置
            prevPosition := jumpPosition
            
            ; 关闭当前模式但保持ToolTip显示
            CoordMode("ToolTip", "Screen")
            ; ToolTip("选择 ←  (Esc退出)", prevPosition.x + 15, prevPosition.y - 30)
            
            ; 关闭输入钩子但不清除ToolTip
            StopInputHook()
            jumpActive := false
            jumpBuffer := ""
            
            ; 激活新模式
            SetTimer () => ActivateJumpMode("selectLeft"), -50
        } else {
            ActivateJumpMode("selectLeft")
        }
    } else {
        SelectLeft()
    }
}

k::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        ; 需要声明要使用的全局变量
        global jumpActive, jumpPosition, jumpBuffer, jumpMode
        
        if(jumpActive) {
            ; 保存当前位置
            prevPosition := jumpPosition
            
            ; 关闭当前模式但保持ToolTip显示
            CoordMode("ToolTip", "Screen")
            ; ToolTip("选择 ↓  (Esc退出)", prevPosition.x + 15, prevPosition.y - 30)
            
            ; 关闭输入钩子但不清除ToolTip
            StopInputHook()
            jumpActive := false
            jumpBuffer := ""
            
            ; 激活新模式
            SetTimer () => ActivateJumpMode("selectDown"), -50
        } else {
            ActivateJumpMode("selectDown")
        }
    } else {
        SelectDown()
    }
}

i::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        ; 需要声明要使用的全局变量
        global jumpActive, jumpPosition, jumpBuffer, jumpMode
        
        if(jumpActive) {
            ; 保存当前位置
            prevPosition := jumpPosition
            
            ; 关闭当前模式但保持ToolTip显示
            CoordMode("ToolTip", "Screen")
            ; ToolTip("选择 ↑  (Esc退出)", prevPosition.x + 15, prevPosition.y - 30)
            
            ; 关闭输入钩子但不清除ToolTip
            StopInputHook()
            jumpActive := false
            jumpBuffer := ""
            
            ; 激活新模式
            SetTimer () => ActivateJumpMode("selectUp"), -50
        } else {
            ActivateJumpMode("selectUp")
        }
    } else {
        SelectUp()
    }
}

l::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        ; 需要声明要使用的全局变量
        global jumpActive, jumpPosition, jumpBuffer, jumpMode
        
        if(jumpActive) {
            ; 保存当前位置
            prevPosition := jumpPosition
            
            ; 关闭当前模式但保持ToolTip显示
            CoordMode("ToolTip", "Screen")
            ; ToolTip("选择 →  (Esc退出)", prevPosition.x + 15, prevPosition.y - 30)
            
            ; 关闭输入钩子但不清除ToolTip
            StopInputHook()
            jumpActive := false
            jumpBuffer := ""
            
            ; 激活新模式
            SetTimer () => ActivateJumpMode("selectRight"), -50
        } else {
            ActivateJumpMode("selectRight")
        }
    } else {
        SelectRight()
    }
}

`;::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        SelectToPageEnd()
    } else {
        SelectWordRight()
    }
}

u::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    SelectHome()
}

o::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    SelectEnd()
}

space::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global lastSpaceTime
    global pendingOperation
    global hasExecutedSingleClick
    
    ; 获取当前时间
    currentTime := A_TickCount
    
    ; 检查是否为双击（阈值内再次按下）
    if (currentTime - lastSpaceTime < doubleClickThreshold) {
        ; 双击 - 终止可能正在进行的单击操作
        pendingOperation := false
        
        ; 阻断输入，防止误触
        BlockInput("On")
        
        ; 检查当前是否有选中内容
        originalClip := ClipboardAll()
        A_Clipboard := ""
        Send("^c")
        hasSelection := ClipWait(0.05, 0)
        selectedText := A_Clipboard
        A_Clipboard := originalClip
        originalClip := ""
        
        ; 检查选中内容是否包含换行符
        if (hasSelection && selectedText != "" && (InStr(selectedText, "`n") || InStr(selectedText, "`r"))) {
            ; 选中内容包含换行符，不执行操作
            ToolTip("已选中多行内容")
            SetTimer () => ToolTip(), -800
        } else {
            ; 直接选中当前行
            SelectCurrentLine()
        }
        
        ; 恢复输入
        BlockInput("Off")
        
        ; 重置状态
        lastSpaceTime := 0
        hasExecutedSingleClick := false
    } else {
        ; 延迟执行单击操作，给双击判断留出空间
        pendingOperation := true
        hasExecutedSingleClick := false
        
        ; 启动选中单词的定时器，稍微延迟执行
        SetTimer SelectWordTimer, -60  ; 60ms后执行，给双击判断留出更多空间
        
        ; 更新时间戳
        lastSpaceTime := currentTime
    }
}

; 删除操作键映射
,::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        ; 需要声明要使用的全局变量
        global jumpActive, jumpPosition, jumpBuffer, jumpMode
        
        if(jumpActive) {
            ; 保存当前位置
            prevPosition := jumpPosition
            
            ; 关闭当前模式但保持ToolTip显示
            CoordMode("ToolTip", "Screen")
            ; ToolTip("删除 ←  (Esc退出)", prevPosition.x + 15, prevPosition.y - 30)
            
            ; 关闭输入钩子但不清除ToolTip
            StopInputHook()
            jumpActive := false
            jumpBuffer := ""
            
            ; 激活新模式
            SetTimer () => ActivateJumpMode("deleteLeft"), -50
        } else {
            ActivateJumpMode("deleteLeft")
        }
    } else {
        DeleteLeft()
    }
}

.::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        ; 需要声明要使用的全局变量
        global jumpActive, jumpPosition, jumpBuffer, jumpMode
        
        if(jumpActive) {
            ; 保存当前位置
            prevPosition := jumpPosition
            
            ; 关闭当前模式但保持ToolTip显示
            CoordMode("ToolTip", "Screen")
            ; ToolTip("删除 →  (Esc退出)", prevPosition.x + 15, prevPosition.y - 30)
            
            ; 关闭输入钩子但不清除ToolTip
            StopInputHook()
            jumpActive := false
            jumpBuffer := ""
            
            ; 激活新模式
            SetTimer () => ActivateJumpMode("deleteRight"), -50
        } else {
            ActivateJumpMode("deleteRight")
        }
    } else {
        DeleteRight()
    }
}

BackSpace::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    DeleteLine()
}

m::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        DeleteToPageBeginning()
    } else {
        DeleteToLineBeginning()
    }
}

/::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    if (GetKeyState("Alt", "P")) {
        DeleteToPageEnd()
    } else {
        DeleteToLineEnd()
    }
}

; 特殊操作键映射
Enter::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    EnterWherever()
}

RShift::
{
   ; 标记为按下了其他键
   global otherKeyPressed := true
   
   ; 在当前行上方插入空行
   IndexWherever()
   
}

z::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    Send("^z")
}

y::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true

    Send("^y")
}

; 剪切板操作的全局变量
global ClipboardSaved_Independent := ""
x::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global ClipboardSaved_Independent
    
    ; 保存原始剪切板
    originalClip := ClipboardAll()
    A_Clipboard := ""  ; 清空剪切板
    
    ; 执行剪切操作
    Send("^x")
    if (ClipWait(0.2, 0)) {  ; 等待剪切操作完成，返回1表示成功
        ; 保存到独立剪切板
        ClipboardSaved_Independent := A_Clipboard
    } else {
        ShowTooltip("剪切失败")
    }
    
    ; 恢复原始剪切板
    A_Clipboard := originalClip
    originalClip := ""  ; 释放变量
}

c::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global ClipboardSaved_Independent
    
    ; 保存原始剪切板
    originalClip := ClipboardAll()
    A_Clipboard := ""  ; 清空剪切板
    
    ; 执行复制操作
    Send("^c")
    if (ClipWait(0.2, 0)) {  ; 等待复制操作完成，返回1表示成功
        ; 保存到独立剪切板
        ClipboardSaved_Independent := A_Clipboard
    } else {
        ShowTooltip("复制失败")
    }
    
    ; 恢复原始剪切板
    A_Clipboard := originalClip
    originalClip := ""  ; 释放变量
}

v::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global ClipboardSaved_Independent
    
    ; 检查独立剪切板是否有内容
    if (ClipboardSaved_Independent = "") {
        Send("^v")
        return
    }
    
    ; 保存原始剪切板
    originalClip := ClipboardAll()
    
    ; 设置剪切板为独立剪切板内容
    A_Clipboard := ""  ; 先清空
    ClipWait(0.1)      ; 等待清空完成
    A_Clipboard := ClipboardSaved_Independent  ; 设置新内容
    ClipWait(0.1)      ; 等待设置完成
    
    Send("^v")
    
    ; 恢复原始剪切板
    A_Clipboard := originalClip
    originalClip := ""  ; 释放变量
}

b::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    Send("^b")
}  

t::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true

    ; 检查是否有选中的内容
    local IsSelected := false
    
    if(GetKeyState("LButton", "P")){
        ShowTooltip("请先松开左键")
        return
    }

    ; 保存剪贴板
    ClipSaved := ClipboardAll()
    A_Clipboard := ""
    Send("^c")
    ClipWait(0.2, 0)

    ; 如果剪贴板不为空，说明有选中内容
    if (A_Clipboard != "") {
        IsSelected := true
    } else {
        IsSelected := false
    }

    if (!IsSelected) {
        MiniTranslate()
    } else {
        SetTimer(SelectTranslate, -100)
    }

    ; 恢复剪贴板
    A_Clipboard := ClipSaved
    ClipSaved := ""
}
#HotIf

MiniTranslate(){
    Send("^m")
}

SelectTranslate(){
    Send("{F7}")
}

; 全局变量和状态
global isSeekingSymbol := false
global interruptCheckTimer := 0  ; 将定时器变量设为全局
global initialCapsLockReleased := false  ; 追踪CapsLock是否已经释放过

; 全局变量和状态
global isSeekingSymbol := false
global interruptCheckTimer := 0
global initialCapsLockReleased := false  ; 追踪CapsLock是否已经释放过

; 主查找热键 - 只在CapsLock按下且未处于查找状态时触发
#HotIf GetKeyState("CapsLock", "P") && !isSeekingSymbol
p::
{
    ; 标记为按下了其他键并启动搜索状态
    global isSeekingSymbol := true
    global otherKeyPressed := true
    global interruptCheckTimer
    global initialCapsLockReleased := false  ; 初始设为false，等待释放
    
    ; 搜索超时计时
    searchStartTime := A_TickCount
    
    ; 保存当前剪贴板内容
    savedClipboard := ClipboardAll()
    A_Clipboard := ""
    
    ; 设置初始检查定时器 - 每20ms检查状态
    interruptCheckTimer := SetTimer(CheckForInterrupt, 20)
    
    ; 定义成对符号映射
    SymbolPairs := Map(
        "(", ")", "[", "]", "{", "}", "<", ">",
        "「", "」", "『", "』", "【", "】", "《", "》", "〈", "〉",
        "（", "）", "［", "］", "｛", "｝", "〔", "〕", "〖", "〗",
        "〘", "〙", "〚", "〛", "“", "”", "‘", "’", "‹", "›", "«", "»"
    )
    
    ReverseSymbolPairs := Map(
        ")", "(", "]", "[", "}", "{", ">", "<",
        "」", "「", "』", "『", "】", "【", "》", "《", "〉", "〈",
        "）", "（", "］", "［", "｝", "｛", "〕", "〔", "〗", "〖",
        "〙", "〘", "〛", "〚", "”", "“", "’", "‘", "›", "‹", "»", "«"
    )
    
    ; 初始化变量
    targetChar := ""
    currentChar := ""
    tempChar := ""
    tempLine_1 := ""
    tempLine_2 := ""
    tempLine_3 := ""  ; 添加第三行记录
    counter := 0
    cursortargetpos := 0

    ; 添加光标位置跟踪变量
    caretPos_1 := {x: 0, y: 0}
    caretPos_2 := {x: 0, y: 0}
    caretPos_3 := {x: 0, y: 0}
    
    ; 获取当前光标位置右侧的字符 - 使用SetTimer代替Sleep
    Send("+{Right}")
    Send("^c")
    startTime := A_TickCount
    
    ; 等待剪贴板内容 - 使用轮询替代Sleep
    while (A_Clipboard = "" && A_TickCount - startTime < 500) {
        if (!isSeekingSymbol) {
            ; 已取消
            if (interruptCheckTimer)
                SetTimer(interruptCheckTimer, 0)
            A_Clipboard := savedClipboard
            SetTimer(ReleaseCapsLock, -50)
            ToolTip("符号跳转已取消")
            SetTimer(() => ToolTip(), -1000)
            return
        }
        Sleep(10)
    }
    
    ; 超时检查
    if (A_Clipboard = "") {
        if (interruptCheckTimer)
            SetTimer(interruptCheckTimer, 0)
        A_Clipboard := savedClipboard
        ToolTip("获取字符超时")
        SetTimer(() => ToolTip(), -1500)
        isSeekingSymbol := false
        return
    }
    
    currentChar := A_Clipboard
    A_Clipboard := ""
    
    ; 检查字符是否是配对符号
    if (!SymbolPairs.Has(currentChar) && !ReverseSymbolPairs.Has(currentChar)) {
        if (interruptCheckTimer)
            SetTimer(interruptCheckTimer, 0)
        ToolTip("光标右侧非配对符号")
        SetTimer(() => ToolTip(), -1500)
        A_Clipboard := savedClipboard
        Send("{Left}")
        isSeekingSymbol := false
        return
    }
    
    ; 向前搜索
    if (SymbolPairs.Has(currentChar)) {
        targetChar := SymbolPairs[currentChar]
        counter := 1
        
        ; 移动到字符右侧
        Send("{Right}")
        Sleep(30)  ; 极短的延迟
        
        ; 开始搜索循环
        loop {
            ; 超时检查
            if (A_TickCount - searchStartTime > 10000) {  ; 10秒超时
                if (interruptCheckTimer)
                    SetTimer(interruptCheckTimer, 0)
                A_Clipboard := savedClipboard
                ToolTip("搜索超时 - 未找到匹配符号")
                SetTimer(() => ToolTip(), -1500)
                isSeekingSymbol := false
                return
            }
            
            ; 取消检查
            if (!isSeekingSymbol) {
                if (interruptCheckTimer)
                    SetTimer(interruptCheckTimer, 0)
                A_Clipboard := savedClipboard
                SetTimer(ReleaseCapsLock, -50)
                ToolTip("符号跳转已取消")
                SetTimer(() => ToolTip(), -1000)
                return
            }
            
            ; 获取当前光标位置
            currentCaretPos := GetCaretPosition()
            
            ; 获取当前光标位置到行尾
            A_Clipboard := ""
            Send("+{End}+{Right}")
            Send("^c")
            
            ; 等待剪贴板内容
            startTime := A_TickCount
            while (A_Clipboard = "" && A_TickCount - startTime < 500) {
                if (!isSeekingSymbol) {
                    if (interruptCheckTimer)
                        SetTimer(interruptCheckTimer, 0)
                    A_Clipboard := savedClipboard
                    SetTimer(ReleaseCapsLock, -50)
                    ToolTip("符号跳转已取消")
                    SetTimer(() => ToolTip(), -1000)
                    return
                }
                Sleep(10)
            }
            
            ; 超时检查
            if (A_Clipboard = "") {
                if (interruptCheckTimer)
                    SetTimer(interruptCheckTimer, 0)
                A_Clipboard := savedClipboard
                ToolTip("获取行内容超时")
                SetTimer(() => ToolTip(), -1500)
                isSeekingSymbol := false
                return
            }
            
            Send("{Left}")
            Sleep(30)
            
            ; 更新三行历史和光标位置历史
            tempLine_3 := tempLine_2
            tempLine_2 := tempLine_1
            tempLine_1 := A_Clipboard
            
            caretPos_3 := caretPos_2
            caretPos_2 := caretPos_1
            caretPos_1 := currentCaretPos
            
            ; 改进的边界检测 - 双重检查内容和光标位置
            if (tempLine_1 = tempLine_3 && tempLine_1 != "" && 
                caretPos_1 && caretPos_3 &&  ; 确保光标位置有效
                caretPos_1.x = caretPos_3.x && caretPos_1.y = caretPos_3.y) {
                
                if (interruptCheckTimer)
                    SetTimer(interruptCheckTimer, 0)
                A_Clipboard := savedClipboard
                ToolTip("到达文档边界 - 未找到匹配符号")
                SetTimer(() => ToolTip(), -1500)
                isSeekingSymbol := false
                return
            }
            
            ; 在当前行中搜索
            loop parse, tempLine_1 {
                tempChar := A_LoopField
                
                ; 取消检查
                if (!isSeekingSymbol) {
                    if (interruptCheckTimer)
                        SetTimer(interruptCheckTimer, 0)
                    A_Clipboard := savedClipboard
                    SetTimer(ReleaseCapsLock, -50)
                    ToolTip("符号跳转已取消")
                    SetTimer(() => ToolTip(), -1000)
                    return
                }
                
                if (tempChar = targetChar) {
                    counter--
                    if (counter = 0) {
                        if (interruptCheckTimer)
                            SetTimer(interruptCheckTimer, 0)
                        cursortargetpos := A_Index - 1
                        A_Clipboard := savedClipboard
                        Send("{Right " . cursortargetpos . "}")
                        ; ToolTip("找到匹配符号")
                        ; SetTimer(() => ToolTip(), -1000)
                        isSeekingSymbol := false
                        SetTimer(ReleaseCapsLock, -50)
                        return
                    }
                } else if (tempChar = currentChar) {
                    counter++
                }
            }
            
            ; 移动到下一行
            Send("{End}{Right}")
            Sleep(30)
        }
    }
    ; 向后搜索
    else if (ReverseSymbolPairs.Has(currentChar)) {
        targetChar := ReverseSymbolPairs[currentChar]
        counter := 1
        
        ; 开始搜索循环
        loop {
            ; 超时检查
            if (A_TickCount - searchStartTime > 10000) {  ; 10秒超时
                if (interruptCheckTimer)
                    SetTimer(interruptCheckTimer, 0)
                A_Clipboard := savedClipboard
                ToolTip("搜索超时 - 未找到匹配符号")
                SetTimer(() => ToolTip(), -1500)
                isSeekingSymbol := false
                return
            }
            
            ; 取消检查
            if (!isSeekingSymbol) {
                if (interruptCheckTimer)
                    SetTimer(interruptCheckTimer, 0)
                A_Clipboard := savedClipboard
                SetTimer(ReleaseCapsLock, -50)
                ToolTip("符号跳转已取消")
                SetTimer(() => ToolTip(), -1000)
                return
            }
            
            ; 获取当前光标位置
            currentCaretPos := GetCaretPosition()
            
            ; 获取当前光标位置到行首
            A_Clipboard := ""
            Send("+{Home}+{Left}")
            Send("^c")
            
            ; 等待剪贴板内容
            startTime := A_TickCount
            while (A_Clipboard = "" && A_TickCount - startTime < 500) {
                if (!isSeekingSymbol) {
                    if (interruptCheckTimer)
                        SetTimer(interruptCheckTimer, 0)
                    A_Clipboard := savedClipboard
                    SetTimer(ReleaseCapsLock, -50)
                    ToolTip("符号跳转已取消")
                    SetTimer(() => ToolTip(), -1000)
                    return
                }
                Sleep(10)
            }
            
            ; 超时检查
            if (A_Clipboard = "") {
                if (interruptCheckTimer)
                    SetTimer(interruptCheckTimer, 0)
                A_Clipboard := savedClipboard
                ToolTip("获取行内容超时")
                SetTimer(() => ToolTip(), -1500)
                isSeekingSymbol := false
                return
            }
            
            Send("{Right}")
            Sleep(30)
            
            ; 更新三行历史和光标位置历史
            tempLine_3 := tempLine_2
            tempLine_2 := tempLine_1
            tempLine_1 := A_Clipboard
            
            caretPos_3 := caretPos_2
            caretPos_2 := caretPos_1
            caretPos_1 := currentCaretPos
            
            ; 改进的边界检测 - 双重检查内容和光标位置
            if (tempLine_1 = tempLine_3 && tempLine_1 != "" && 
                caretPos_1 && caretPos_3 &&  ; 确保光标位置有效
                caretPos_1.x = caretPos_3.x && caretPos_1.y = caretPos_3.y) {
                
                if (interruptCheckTimer)
                    SetTimer(interruptCheckTimer, 0)
                A_Clipboard := savedClipboard
                ToolTip("到达文档边界 - 未找到匹配符号")
                SetTimer(() => ToolTip(), -1500)
                isSeekingSymbol := false
                return
            }
            
            ; 在当前行中从右向左搜索
            lineLength := StrLen(tempLine_1)
            loop lineLength {
                position := lineLength - A_Index + 1
                tempChar := SubStr(tempLine_1, position, 1)
                
                ; 取消检查
                if (!isSeekingSymbol) {
                    if (interruptCheckTimer)
                        SetTimer(interruptCheckTimer, 0)
                    A_Clipboard := savedClipboard
                    SetTimer(ReleaseCapsLock, -50)
                    ToolTip("符号跳转已取消")
                    SetTimer(() => ToolTip(), -1000)
                    return
                }
                
                if (tempChar = targetChar) {
                    counter--
                    if (counter = 0) {
                        if (interruptCheckTimer)
                            SetTimer(interruptCheckTimer, 0)
                        cursortargetpos := A_Index - 1
                        A_Clipboard := savedClipboard
                        Send("{Left " . cursortargetpos . "}")
                        ; ToolTip("找到匹配符号")
                        ; SetTimer(() => ToolTip(), -1000)
                        isSeekingSymbol := false
                        SetTimer(ReleaseCapsLock, -50)
                        return
                    }
                } else if (tempChar = currentChar) {
                    counter++
                }
            }
            
            ; 移动到上一行
            Send("{Up}{End}")
            Sleep(30)
        }
    }
    
    ; 理论上不应该到这里
    if (interruptCheckTimer)
        SetTimer(interruptCheckTimer, 0)
    A_Clipboard := savedClipboard
    isSeekingSymbol := false
    ToolTip("搜索结束但未找到匹配符号")
    SetTimer(() => ToolTip(), -1500)
}
#HotIf

; 释放CapsLock的函数
ReleaseCapsLock() {
    SendInput "{CapsLock Up}"
}

; 检查是否需要中断的函数
CheckForInterrupt() {
    global isSeekingSymbol, initialCapsLockReleased
    
    ; 首先检查Escape - 这个随时可以中断
    if (GetKeyState("Escape", "P")) {
        isSeekingSymbol := false
        return
    }
    
    ; 然后检查CapsLock - 使用两阶段逻辑
    capsLockCurrentState := GetKeyState("CapsLock", "P")
    
    ; 阶段1: 检测CapsLock释放
    if (!initialCapsLockReleased && !capsLockCurrentState) {
        initialCapsLockReleased := true  ; 标记CapsLock已释放
        return
    }
    
    ; 阶段2: 检测在释放后的再次按下
    if (initialCapsLockReleased && capsLockCurrentState) {
        isSeekingSymbol := false  ; 触发中断
        return
    }
}

; 用于取消的热键
#HotIf isSeekingSymbol
Escape::
{
    global isSeekingSymbol := false
    return
}

CapsLock::
{
    global isSeekingSymbol := false
    return
}
#HotIf

; 放大镜功能(调用win11原生放大镜)
#HotIf GetKeyState("CapsLock", "P")
Tab::
{
   ; 标记为按下了其他键
    global otherKeyPressed := true
    
    ; 检查放大镜是否已经开启
    if (WinExist("ahk_class MagUIClass")) {
        ; 放大镜已开启，关闭它
        WinClose("ahk_class MagUIClass")
        ShowTooltip("放大镜已关闭")
    } else {
        ; 放大镜未开启，打开它
        Send("#=")
        ShowTooltip("放大镜已开启")
        
        ; 等待放大镜窗口出现并最小化它
        SetTimer(MinimizeMagnifierWindow, -1000)
    }
}
#HotIf

; 最小化放大镜窗口
MinimizeMagnifierWindow() {
    if (WinExist("ahk_class MagUIClass")) {
        WinMinimize("ahk_class MagUIClass")
        ShowTooltip("放大镜控制窗口已最小化")
    }
}

; 检查进程是否存在
ProcessExist(processName) {
    return ProcessExist := ProcessWaitClose(processName, 0)
}

; 暂时空置其他按键, 未来按需添加, 也可以通过这些空置键来取消映射esc避免误触
#HotIf GetKeyState("CapsLock", "P")

-::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    Send("")
}

=::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    Send("")
}

[::
    {
        ; 标记为按下了其他键
        global otherKeyPressed := true
        
        Send("")
}
    
]::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    Send("")
}

\::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    Send("")
}

'::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    Send("")
}
#HotIf

; =====================================================================
; 鼠标控制热键
; =====================================================================

; 鼠标移动 - 按下时记录方向并开始移动
CapsLock & Up::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global mousePrecisionMode := GetKeyState("LAlt", "P")
    global mouseKeysPressed
    
    ; 记录按下的方向键
    mouseKeysPressed["Up"] := true
    
    ; 开始连续移动
    StartContinuousMouseMovement()
}

; 释放方向键时更新状态
CapsLock & Up Up::
{
    global mouseKeysPressed
    mouseKeysPressed["Up"] := false
}

CapsLock & Down::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global mousePrecisionMode := GetKeyState("LAlt", "P")
    global mouseKeysPressed
    
    ; 记录按下的方向键
    mouseKeysPressed["Down"] := true
    
    ; 开始连续移动
    StartContinuousMouseMovement()
}

CapsLock & Down Up::
{
    global mouseKeysPressed
    mouseKeysPressed["Down"] := false
}

CapsLock & Left::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global mousePrecisionMode := GetKeyState("LAlt", "P")
    global mouseKeysPressed
    
    ; 记录按下的方向键
    mouseKeysPressed["Left"] := true
    
    ; 开始连续移动
    StartContinuousMouseMovement()
}

CapsLock & Left Up::
{
    global mouseKeysPressed
    mouseKeysPressed["Left"] := false
}

CapsLock & Right::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global mousePrecisionMode := GetKeyState("LAlt", "P")
    global mouseKeysPressed
    
    ; 记录按下的方向键
    mouseKeysPressed["Right"] := true
    
    ; 开始连续移动
    StartContinuousMouseMovement()
}

CapsLock & Right Up::
{
    global mouseKeysPressed
    mouseKeysPressed["Right"] := false
}

; 精确移动模式切换热键
!CapsLock::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global mousePrecisionMode := !mousePrecisionMode
}

; 鼠标点击
CapsLock & RAlt::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    MouseLeftDown()
}

CapsLock & RAlt Up::
{
    MouseLeftUp()
}

CapsLock & RCtrl::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true

    MouseRightDown()
}

CapsLock & RWin::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true

    MouseRightDown()
}

CapsLock & RCtrl Up::
{
    MouseRightUp()
}

CapsLock & RWin Up::
{
    MouseRightUp()
}

; 鼠标滚轮
CapsLock & PgUp::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    MouseWheelUp(3)
}

CapsLock & PgDn::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    MouseWheelDown(3)
}

; 鼠标速度调整
CapsLock & LAlt::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    AdjustMouseSpeed()
}

; 初始化CapsLock+模块
InitCapsLockPlusModule() 

; =====================================================================
; 窗口置顶功能
; =====================================================================

; 全局变量跟踪置顶窗口
global pinnedWindows := []  ; 存储置顶窗口的句柄
global lastFullscreenWarningTime := 0  ; 上次全屏窗口警告时间
global lastFullscreenWarningHwnd := 0  ; 上次警告的全屏窗口句柄
global fullscreenWarningTimeout := 1000  ; 全屏窗口警告超时时间(毫秒)

; 检查窗口是否处于全屏状态
IsWindowFullScreen(hwnd) {
    ; 检查窗口是否最大化
    isMaximized := WinGetMinMax("ahk_id " hwnd) = 1
    
    ; 获取窗口位置和大小
    WinGetPos(&winX, &winY, &winWidth, &winHeight, "ahk_id " hwnd)
    
    ; 获取窗口所在的显示器信息
    monitorIndex := DllCall("MonitorFromWindow", "Ptr", hwnd, "UInt", 0x2)  ; MONITOR_DEFAULTTONEAREST = 0x2
    
    ; 获取显示器信息
    numput("UInt", 40, MONITORINFO := Buffer(40))  ; sizeof(MONITORINFO) = 40
    if (DllCall("GetMonitorInfo", "Ptr", monitorIndex, "Ptr", MONITORINFO)) {
        ; 提取显示器工作区域信息
        monitorLeft := NumGet(MONITORINFO, 20, "Int")    ; rcWork.left
        monitorTop := NumGet(MONITORINFO, 24, "Int")     ; rcWork.top
        monitorRight := NumGet(MONITORINFO, 28, "Int")   ; rcWork.right
        monitorBottom := NumGet(MONITORINFO, 32, "Int")  ; rcWork.bottom
        
        ; 提取显示器完整区域信息（包括任务栏）
        monitorFullLeft := NumGet(MONITORINFO, 4, "Int")    ; rcMonitor.left
        monitorFullTop := NumGet(MONITORINFO, 8, "Int")     ; rcMonitor.top
        monitorFullRight := NumGet(MONITORINFO, 12, "Int")   ; rcMonitor.right
        monitorFullBottom := NumGet(MONITORINFO, 16, "Int")  ; rcMonitor.bottom
        
        monitorWidth := monitorRight - monitorLeft
        monitorHeight := monitorBottom - monitorTop
        
        monitorFullWidth := monitorFullRight - monitorFullLeft
        monitorFullHeight := monitorFullBottom - monitorFullTop
        
        ; 检查窗口是否覆盖整个显示器区域（包括任务栏）
        ; 允许有1像素的误差
        isFullScreenWithTaskbar := (Abs(winX - monitorFullLeft) <= 1) && 
                        (Abs(winY - monitorFullTop) <= 1) && 
                        (Abs(winWidth - monitorFullWidth) <= 1) && 
                        (Abs(winHeight - monitorFullHeight) <= 1)
        
        ; 检查窗口是否覆盖整个工作区域（不包括任务栏）
        isFullScreenWorkArea := (Abs(winX - monitorLeft) <= 1) && 
                        (Abs(winY - monitorTop) <= 1) && 
                        (Abs(winWidth - monitorWidth) <= 1) && 
                        (Abs(winHeight - monitorHeight) <= 1)
        
        ; 如果窗口是最大化的，或者覆盖了整个显示器区域或工作区域，则认为是全屏
        return isMaximized || isFullScreenWithTaskbar || isFullScreenWorkArea
    }
    
    ; 如果无法获取显示器信息，使用备用方法
    ; 检查窗口是否最大化或覆盖整个主显示器
    screenWidth := A_ScreenWidth
    screenHeight := A_ScreenHeight
    
    ; 允许有1像素的误差
    isFullScreenBySize := (Abs(winX) <= 1) && 
                          (Abs(winY) <= 1) && 
                          (Abs(winWidth - screenWidth) <= 1) && 
                          (Abs(winHeight - screenHeight) <= 1)
    
    return isMaximized || isFullScreenBySize
}

; 置顶/取消置顶当前窗口
ToggleWindowPinned() {
    global lastFullscreenWarningTime, lastFullscreenWarningHwnd, fullscreenWarningTimeout
    
    try {
        ; 获取光标下的窗口，而不是当前活动窗口
        hwnd := GetWindowUnderCursor()
        
        ; 确保窗口有效
        if !WinExist("ahk_id " hwnd) {
            ToolTip("光标下无有效窗口")
            SetTimer () => ToolTip(), -2000
            return
        }
        if !IsTaskbarWindow(hwnd) {
            ToolTip("当前窗口不是任务栏窗口，无法置顶")
            SetTimer () => ToolTip(), -2000
            return
        }
        
        title := WinGetTitle("ahk_id " hwnd)
        
        ; 获取窗口进程信息
        pid := WinGetPID("ahk_id " hwnd)
        processPath := ProcessGetPath(pid)
        SplitPath(processPath, &processName)
        
        ; 简化进程名称，移除.exe扩展名
        simpleName := RegExReplace(processName, "\.exe$", "")
        
        ; 获取窗口类名
        className := WinGetClass("ahk_id " hwnd)
        
        ; 特殊处理桌面窗口 - 不允许置顶真正的桌面
        ; 只检查真正的桌面窗口类（Progman或WorkerW），而不是打开桌面路径的Explorer窗口
        if (className = "Progman" || className = "WorkerW") {
            ToolTip("桌面窗口不能置顶")
            SetTimer () => ToolTip(), -2000
            return
        }
        
        ; 获取窗口当前置顶状态
        exStyle := WinGetExStyle("ahk_id " hwnd)
        isPinned := (exStyle & 0x8) != 0  ; WS_EX_TOPMOST
        
        ; 检查窗口是否处于全屏状态
        isFullScreen := IsWindowFullScreen(hwnd)
        
        ; 如果窗口已经置顶，允许取消置顶，无论是否全屏
        if (isPinned) {
            ; 取消置顶
            WinSetAlwaysOnTop(0, "ahk_id " hwnd)
            
            ; 从置顶窗口列表中移除
            for i, pinHwnd in pinnedWindows {
                if (pinHwnd = hwnd) {
                    pinnedWindows.RemoveAt(i)
                    break
                }
            }
            
            ToolTip("已取消置顶: " simpleName)
        } else {
            ; 检查窗口是否处于全屏状态
            if (isFullScreen) {
                ; 获取当前时间
                currentTime := A_TickCount
                
                ; 检查是否是在警告超时时间内对同一窗口的二次操作
                if (hwnd = lastFullscreenWarningHwnd && 
                    (currentTime - lastFullscreenWarningTime) < fullscreenWarningTimeout) {
                    ; 是二次操作，强制置顶
                    WinSetAlwaysOnTop(1, "ahk_id " hwnd)
                    
                    ; 添加到置顶窗口列表
                    if !HasVal(pinnedWindows, hwnd)
                        pinnedWindows.Push(hwnd)
                    
                    ToolTip("已强制置顶全屏窗口: " simpleName)
                    
                    ; 重置警告状态
                    lastFullscreenWarningTime := 0
                    lastFullscreenWarningHwnd := 0
                } else {
                    ; 首次操作，显示警告
                    ToolTip("全屏窗口不建议置顶: " simpleName "`n(再次操作将强制置顶)")
                    
                    ; 记录警告时间和窗口
                    lastFullscreenWarningTime := currentTime
                    lastFullscreenWarningHwnd := hwnd
                }
                
                SetTimer () => ToolTip(), -fullscreenWarningTimeout
                return
            }
            
            ; 设置为置顶
            WinSetAlwaysOnTop(1, "ahk_id " hwnd)
            
            ; 添加到置顶窗口列表
            if !HasVal(pinnedWindows, hwnd)
                pinnedWindows.Push(hwnd)
            
            ToolTip("已置顶窗口: " simpleName)
        }
        
        SetTimer () => ToolTip(), -2000
        
    } catch as e {
        ToolTip("置顶窗口操作失败: " e.Message)
        SetTimer () => ToolTip(), -2000
    }
}

; 对光标下的文件或文件夹执行重命名操作
RenameFileUnderCursor() {
    ; 发送F2键来触发重命名操作
    Send("{F2}")
}

; 获取光标下的窗口句柄
GetWindowUnderCursor() {
    try {
        ; 获取当前鼠标位置
        MouseGetPos(&xpos, &ypos, &hwnd)
        
        ; 确保获取到有效的窗口句柄
        if (!hwnd || !WinExist("ahk_id " hwnd)) {
            ; 如果无法获取有效窗口，返回当前活动窗口
            activeHwnd := WinGetID("A")
            return activeHwnd
        }
        
        ; 获取窗口类名和进程名
        class := WinGetClass("ahk_id " hwnd)
        pid := WinGetPID("ahk_id " hwnd)
        processPath := ProcessGetPath(pid)
        SplitPath(processPath, &processName)
        
        ; 检查是否是任务栏 - 如果是任务栏，返回特殊值
        if (class = "Shell_TrayWnd" || class = "Shell_SecondaryTrayWnd") {
            if (showDebugTooltips) {
                ToolTip("光标在任务栏上，不执行最小化操作")
                SetTimer () => ToolTip(), -2000
            }
            return -1  ; 返回-1表示任务栏
        }
        
        ; 只有真正的桌面窗口类才会被排除（Progman或WorkerW）
        ; 而打开桌面路径的Explorer窗口（CabinetWClass）应该被允许置顶
        if (class = "Progman" || class = "WorkerW") {
            if (showDebugTooltips) {
                title := WinGetTitle("ahk_id " hwnd)
                ToolTip("检测到真正的桌面窗口: " title "`n类: " class)
                SetTimer () => ToolTip(), -2000
            }
            
            ; 将其视为桌面窗口，返回0表示无效窗口
            return 0
        }
        
        ; 调试信息
        if (showDebugTooltips) {
            title := WinGetTitle("ahk_id " hwnd)
            class := WinGetClass("ahk_id " hwnd)
            ToolTip("光标下窗口: " title "`n类: " class "`n句柄: " hwnd)
            SetTimer () => ToolTip(), -2000
        }
        
        return hwnd
    } catch as e {
        ; 出错时返回当前活动窗口
        try {
            activeHwnd := WinGetID("A")
            return activeHwnd
        } catch {
            ; 如果连活动窗口都无法获取，返回0
            return 0
        }
    }
}

; 按PID对数组进行排序，同一PID的窗口按标题排序
SortByPID(arr) {
    global showDebugTooltips
    
    n := arr.Length
    Loop n - 1 {
        i := A_Index
        Loop n - i {
            j := A_Index
            
            ; 首先按PID排序
            if (arr[j].pid > arr[j+1].pid) {
                ; 如果PID不同，按PID排序
                temp := arr[j]
                arr[j] := arr[j+1]
                arr[j+1] := temp
            } 
            else if (arr[j].pid = arr[j+1].pid) {
                ; 如果PID相同且启用了多实例处理，则按窗口标题排序
                try {
                    if (arr[j].title > arr[j+1].title) {
                        temp := arr[j]
                        arr[j] := arr[j+1]
                        arr[j+1] := temp
                    }
                } catch {
                    ; 如果标题比较失败，则按窗口句柄排序
                    if (arr[j].hwnd > arr[j+1].hwnd) {
                        temp := arr[j]
                        arr[j] := arr[j+1]
                        arr[j+1] := temp
                    }
                }
            }
        }
    }
    
    ; 调试输出排序后的窗口列表
    if (showDebugTooltips) {
        debugText := "排序后的窗口列表 (按PID和标题):`n"
        for i, win in arr {
            debugText .= i ": " win.title " (PID:" win.pid ", 句柄:" win.hwnd ")`n"
            if (i > 8)  ; 显示前8个窗口
                break
        }
        ToolTip(debugText)
        SetTimer () => ToolTip(), -3000
    }
    
    return arr
}

;=====================================================================
; 速记功能
;=====================================================================

; 获取桌面路径
GetDesktopPath() {
    ; 尝试从环境变量获取桌面路径
    try {
        desktopPath := EnvGet("USERPROFILE") . "\Desktop\"
        if (FileExist(desktopPath))
            return desktopPath
    } catch {
        ; 如果获取失败，使用默认路径
    }
    
    ; 备用方法：使用特殊文件夹常量获取桌面路径
    try {
        desktopPath := A_Desktop . "\"
        if (FileExist(desktopPath))
            return desktopPath
    } catch {
        ; 如果获取失败，使用默认路径
    }
    
    ; 如果上述方法都失败，返回默认路径
    return "C:\Users\" . A_UserName . "\Desktop\"
}

; 速记功能配置
global noteConfig := {
    defaultDir: GetDesktopPath() . "速记\默认\"  ; 默认保存目录
}

; 预定义目标文件 - 从INI文件读取
global noteTargets := Map()

; 从INI文件加载速记目标
LoadNoteTargetsFromINI() {
    global noteTargets
    
    ; 清空现有映射
    noteTargets := Map()
    
    ; 循环读取配置项
    Loop 10 {
        i := A_Index
        keyName := "note" . i . "1"
        pathName := "note" . i . "2"
        
        ; 读取关键字和路径
        keyword := IniRead("CapsLock++.ini", "A_Section", keyName, "")
        path := IniRead("CapsLock++.ini", "A_Section", pathName, "")
        
        ; 如果两者都不为空，则添加到映射中
        if (keyword != "" && path != "") {
            ; 处理路径 - 如果是相对路径，添加桌面路径前缀
            if (!RegExMatch(path, "^[A-Za-z]:\\")) {
                path := GetDesktopPath() . path
            }
            
            ; 添加到映射
            noteTargets[keyword] := path
        }
    }
    
    ; 如果映射为空，添加默认项
    if (noteTargets.Count = 0) {
        noteTargets := Map(
            "论文", GetDesktopPath() . "速记\论文灵感.txt",
            "日记", GetDesktopPath() . "速记\日记.txt",
            "工作", GetDesktopPath() . "速记\工作.txt",
            "想法", GetDesktopPath() . "速记\想法.txt"
        )
    }
}

; 初始化时加载速记目标
LoadNoteTargetsFromINI()

; 初始化速记目录
EnsureNoteDirectories()

#HotIf GetKeyState("CapsLock", "P")
n::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    ; 显示速记窗口
    ShowQuickNote()
}

; 显示速记窗口
ShowQuickNote() {
    static isNoteGuiOpen := false
    
    ; 防止重复打开
    if (isNoteGuiOpen)
        return
        
    isNoteGuiOpen := true
    
    ; 创建GUI
    noteGui := Gui("+AlwaysOnTop +Resize", "速记")
    
    ; 设置自定义图标（如果存在）
    iconPath := A_ScriptDir . "\Icon\QuickNote.ico"
    if (FileExist(iconPath)) {
        ; 直接使用图标文件设置窗口图标
        hIcon := DllCall("LoadImage", "Ptr", 0, "Str", iconPath, "UInt", 1, "Int", 0, "Int", 0, "UInt", 0x10, "Ptr")
        if (hIcon) {
            DllCall("SendMessage", "Ptr", noteGui.Hwnd, "UInt", 0x80, "Ptr", 1, "Ptr", hIcon)  ; WM_SETICON, ICON_BIG
            DllCall("SendMessage", "Ptr", noteGui.Hwnd, "UInt", 0x80, "Ptr", 0, "Ptr", hIcon)  ; WM_SETICON, ICON_SMALL
        }
    }
    
    ; 添加多行编辑框（添加滚动条），预先填入"## "
    noteEdit := noteGui.Add("Edit", "vNoteContent r10 w500 WantTab +VScroll", "## ")
    
    ; 添加状态栏显示帮助信息
    statusBar := noteGui.Add("StatusBar", "", "提示: 输入标题或删除## | 最后一行使用==目标==指定保存位置 | Ctrl+S保存")
    
    ; 添加按钮
    buttonBar := noteGui.Add("Text", "w500 h30 Section", "")
    saveBtn := noteGui.Add("Button", "xp y+5 w80 h25 Default", "保存")
    cancelBtn := noteGui.Add("Button", "x+10 yp w80 h25", "取消")
    
    ; 设置按钮事件
    saveBtn.OnEvent("Click", SaveNoteHandler)
    cancelBtn.OnEvent("Click", CloseNoteGui)
    
    ; 设置窗口关闭事件
    noteGui.OnEvent("Close", CloseNoteGui)
    noteGui.OnEvent("Escape", CloseNoteGui)
    
    ; 关闭窗口处理函数
    CloseNoteGui(*) {
        isNoteGuiOpen := false
        noteGui.Destroy()
    }
    
    ; GUI大小调整处理函数
    GuiResize(thisGui, minMax, width, height) {
        if (minMax = -1)  ; 窗口被最小化
            return
            
        ; 调整编辑框大小
        noteEdit.Move(,, width, height - 70)
        
        ; 调整按钮位置
        buttonBar.Move(,, width)
        saveBtn.Move(, height - 60)
        cancelBtn.Move(width - 90, height - 60)
    }
    
    ; 设置大小调整事件
    noteGui.OnEvent("Size", GuiResize)
    
    ; 显示GUI并聚焦到编辑框
    noteGui.Show()
    WinWaitActive("速记")
    ControlFocus("Edit1", "速记")
    
    ; 将光标定位到"## "之后
    SendInput("{End}")
    
    ; 注册Ctrl+S快捷键
    HotIfWinActive("速记")
    Hotkey("^s", (*) => SaveNoteHandler())
    HotIf()
    
    ; 保存笔记处理函数
    SaveNoteHandler(*) {
        ; 获取内容
        content := noteEdit.Value
        
        ; 解析内容
        lines := StrSplit(content, "`n")
        title := ""
        targetName := ""
        targetFile := ""
        
        ; 检查是否为空内容
        if (content = "") {
            MsgBox("笔记内容为空，未保存", "速记", "Icon!")
            return
        }
        
        ; 检查第一行是否为有效标题 (以"## "开头且后面有内容)
        if (lines.Length > 0 && RegExMatch(lines[1], "^##\s+(.+)$", &match)) {
            titleText := Trim(match[1])
            ; 确保标题不为空
            if (titleText != "")
                title := titleText
        }
        
        ; 检查最后一行是否指定目标文件
        if (lines.Length > 0) {
            lastLine := lines[lines.Length]
            if (RegExMatch(lastLine, "^==\s*(.+?)\s*==$", &match)) {
                targetName := Trim(match[1])
                ; 移除最后一行
                lines.Pop()
                
                ; 检查是否为预设目标
                if (noteTargets.Has(targetName)) {
                    targetFile := noteTargets[targetName]
                } else {
                    ; 检查默认文件夹中是否存在同名文件
                    potentialFile := noteConfig.defaultDir . targetName . ".txt"
                    if (FileExist(potentialFile)) {
                        targetFile := potentialFile
                    }
                }
            }
        }
        
        ; 如果第一行只是"## "或"##"（没有实际标题内容），则移除它
        if (lines.Length > 0 && RegExMatch(lines[1], "^##\s*$")) {
            lines.RemoveAt(1)
        }
        
        ; 根据解析结果保存内容
        if (targetFile && FileExist(targetFile)) {
            ; 保存到指定文件
            SaveToSpecificFile(targetFile, lines, title, targetName)
        } else if (targetName && noteTargets.Has(targetName)) {
            ; 保存到预定义文件
            SaveToTargetFile(targetName, lines, title)
        } else {
            ; 创建新文件或追加到同名文件
            SaveToNewFile(lines, title)
        }
        
        ; 关闭窗口
        isNoteGuiOpen := false
        noteGui.Destroy()
    }
}

; 创建新文件或追加到同名文件
SaveToNewFile(lines, title) {
    try {
        ; 确定文件名（清理标题中的非法字符）
        if (title) {
            ; 使用清理函数处理标题中的非法字符
            cleanTitle := CleanFileNameFromTitle(title)
            fileName := cleanTitle . ".txt"
        } else {
            fileName := FormatTime(, "yyyyMMdd_HHmmss") . ".txt"
        }
        
        filePath := noteConfig.defaultDir . fileName
        
        ; 准备内容
        content := ""
        timeStamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
        
        ; 检查文件是否已存在
        fileExists := FileExist(filePath)
        
        ; 如果文件已存在且有标题，则追加内容
        if (fileExists && title) {
            ; 添加分隔空行
            content := "`n`n`n"  ; 添加三个空行作为分隔（与上面内容空两行）
            
            ; 检查是否有新标题（第一行是否以##开头且后面有内容）
            hasNewTitle := false
            newTitle := ""
            if (lines.Length > 0 && RegExMatch(lines[1], "^##\s+(.+)$", &match)) {
                titleText := Trim(match[1])
                if (titleText != "") {
                    hasNewTitle := true
                    newTitle := lines[1]
                }
            }
            
            ; 先添加标题（如果有）
            if (hasNewTitle) {
                content .= newTitle . "`n"
            }
            
            ; 添加时间戳作为普通标记（不是标题）
            content .= "[" . timeStamp . "]`n`n"
            
            ; 添加正文
            for i, line in lines {
                ; 如果第一行是标题且我们已经提取了它，则跳过
                if (i = 1 && hasNewTitle) {
                    continue
                }
                content .= line . "`n"
            }
            
            ; 追加到文件
            FileAppend(content, filePath)
            
            ; 显示成功消息
            ToolTip("已追加到「" . filePath . "」")
        } else {
            ; 如果文件不存在或没有标题，创建新文件
            
            ; 如果有标题，添加标题和时间戳
            if (title) {
                ; 检查第一行是否为标题行
                hasTitle := false
                if (lines.Length > 0 && RegExMatch(lines[1], "^##\s+(.+)$", &match)) {
                    titleText := Trim(match[1])
                    if (titleText != "") {
                        hasTitle := true
                    }
                }
                
                if (hasTitle) {
                    ; 第一行是有效标题，保留它并添加时间戳
                    for i, line in lines {
                        if (i = 1) {
                            content .= line . "`n"
                            content .= "[" . timeStamp . "]`n`n"  ; 在标题后添加时间戳
                        } else {
                            content .= line . "`n"
                        }
                    }
                } else {
                    ; 第一行不是有效标题，添加标题和时间戳
                    content .= "## " . title . "`n"
                    content .= "[" . timeStamp . "]`n`n"
                    
                    ; 添加所有行
                    for i, line in lines {
                        content .= line . "`n"
                    }
                }
            } else {
                ; 无标题，只添加时间戳（不作为标题）
                content .= "[" . timeStamp . "]`n`n"  ; 使用方括号格式的时间戳
                
                ; 添加所有行，包括可能的"## "
                for i, line in lines {
                    content .= line . "`n"
                }
            }
            
            ; 写入文件
            FileAppend(content, filePath)
            
            ; 显示成功消息
            ToolTip("已保存到「" . filePath . "」")
        }
        
        SetTimer () => ToolTip(), -2000
    } catch as e {
        MsgBox("保存失败: " . e.Message, "速记", "Icon!")
    }
}

; 保存到目标文件
SaveToTargetFile(targetName, lines, title) {
    try {
        filePath := noteTargets[targetName]
        
        ; 准备要追加的内容
        content := "`n`n`n"  ; 添加三个空行作为分隔（与上面内容空两行）
        timeStamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
        
        ; 添加标题和时间戳
        if (title) {
            content .= "## " . title . "`n"  ; 使用Markdown格式的标题
            content .= "[" . timeStamp . "]`n`n"  ; 在标题后添加时间戳
        } else {
            ; 无标题，只添加时间戳（不作为标题）
            content .= "[" . timeStamp . "]`n`n"  ; 使用方括号格式的时间戳
        }
        
        ; 添加正文
        for i, line in lines {
            ; 如果第一行是标题行且已提取标题，则跳过
            if (i = 1 && title && RegExMatch(line, "^##\s+")) {
                continue
            }
            content .= line . "`n"
        }
        
        ; 追加到文件
        FileAppend(content, filePath)
        
        ; 显示成功消息
        ToolTip("已保存到「" . targetName . "」")
        SetTimer () => ToolTip(), -2000
    } catch as e {
        MsgBox("保存失败: " . e.Message, "速记", "Icon!")
    }
}

; 重复字符串函数
StrRepeat(str, count) {
    result := ""
    Loop count {
        result .= str
    }
    return result
}

; 确保速记目录结构存在
EnsureNoteDirectories() {
    global noteConfig, noteTargets
    
    ; 确保主速记目录存在
    noteDir := GetDesktopPath() . "速记\"
    if (!FileExist(noteDir)) {
        try {
            DirCreate(noteDir)
        } catch as e {
            ToolTip("创建速记主目录失败: " . e.Message)
            SetTimer () => ToolTip(), -3000
            return false
        }
    }
    
    ; 确保默认目录存在
    if (!FileExist(noteConfig.defaultDir)) {
        try {
            DirCreate(noteConfig.defaultDir)
        } catch as e {
            ToolTip("创建速记默认目录失败: " . e.Message)
            SetTimer () => ToolTip(), -3000
            return false
        }
    }
    
    ; 确保目标文件所在的目录存在
    for targetName, filePath in noteTargets {
        ; 获取文件所在目录
        SplitPath(filePath, , &fileDir)
        
        ; 如果目录不存在，创建它
        if (fileDir && !FileExist(fileDir)) {
            try {
                DirCreate(fileDir)
            } catch as e {
                ToolTip("创建目标文件目录失败: " . e.Message)
                SetTimer () => ToolTip(), -3000
            }
        }
    }
    
    return true
}

; 清理文件名函数 - 将标题中的非法字符替换为安全字符
CleanFileNameFromTitle(title) {
    ; 替换Windows文件系统不允许的字符: \ / : * ? " < > |
    ; 使用一个更友好的替换方式，将特殊字符替换为对应的中文描述
    title := RegExReplace(title, "\\", "「反斜杠」")
    title := RegExReplace(title, "/", "「斜杠」")
    title := RegExReplace(title, ":", "「冒号」")
    title := RegExReplace(title, "\*", "「星号」")
    title := RegExReplace(title, "\?", "「问号」")
    title := RegExReplace(title, "`"", "「引号」")
    title := RegExReplace(title, "<", "「小于」")
    title := RegExReplace(title, ">", "「大于」")
    title := RegExReplace(title, "\|", "「竖线」")
    return title
}

; 保存到指定文件
SaveToSpecificFile(filePath, lines, title, displayName) {
    try {
        ; 准备要追加的内容
        content := "`n`n`n"  ; 添加三个空行作为分隔（与上面内容空两行）
        timeStamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
        
        ; 添加标题和时间戳
        if (title) {
            content .= "## " . title . "`n"  ; 使用Markdown格式的标题
            content .= "[" . timeStamp . "]`n`n"  ; 在标题后添加时间戳
        } else {
            ; 无标题，只添加时间戳（不作为标题）
            content .= "[" . timeStamp . "]`n`n"  ; 使用方括号格式的时间戳
        }
        
        ; 添加正文
        for i, line in lines {
            ; 如果第一行是标题行且已提取标题，则跳过
            if (i = 1 && title && RegExMatch(line, "^##\s+")) {
                continue
            }
            content .= line . "`n"
        }
        
        ; 追加到文件
        FileAppend(content, filePath)
        
        ; 显示成功消息
        if (displayName) {
            ToolTip("已保存到「" . displayName . "」")
        } else {
            SplitPath(filePath, &fileName)
            ToolTip("已保存到「" . fileName . "」")
        }
        SetTimer () => ToolTip(), -2000
    } catch as e {
        MsgBox("保存失败: " . e.Message, "速记", "Icon!")
    }
}

#HotIf GetKeyState("CapsLock", "P")
q::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    
    ; 保存当前剪贴板内容
    ClipSaved := ClipboardAll()
    A_Clipboard := ""
    
    ; 复制选中内容
    Send("^c")
    ClipWait(0.5)
    SearchText := A_Clipboard
    
    ; 如果有选中文本，进行搜索或打开网址
    if (SearchText != "") {
        ; 检查是否是磁盘路径（例如 C:\Users 或 D:\ 或 \\server\share）
        if (RegExMatch(SearchText, "i)^([a-z]:\\|\\\\[^\\]+\\[^\\]+)") || RegExMatch(SearchText, "i)^[a-z]:$")) {
            ; 尝试打开路径
            try {
                Run "explorer.exe " SearchText
            } catch Error as e {
                ; 如果打开失败，使用简化的Bing搜索
                Run "https://cn.bing.com/search?q=" . UrlEncode(SearchText) . "&FORM=QBLH"
            }
        }
        ; 检查是否是网址（以http://、https://、www.开头）
        else if (RegExMatch(SearchText, "i)^(https?://|www\.)")) {
            ; 如果不是以http://或https://开头，添加https://
            if (!RegExMatch(SearchText, "i)^https?://")) {
                SearchText := "https://" . SearchText
            }
            ; 直接打开网址
            Run SearchText
        } else {
            ; 使用简化的Bing搜索URL
            Run "https://cn.bing.com/search?q=" . UrlEncode(SearchText) . "&FORM=QBLH"
        }
    }
    
    ; 恢复剪贴板
    A_Clipboard := ClipSaved
    ClipSaved := ""
}

; 改进的URL编码函数，专门针对中文和特殊字符
UrlEncode(str) {
    ; 使用更简单的方法处理URL编码
    encoded := ""
    VarSetStrCapacity(&encoded, StrLen(str) * 3)  ; 预分配足够的空间
    
    for i, char in StrSplit(str) {
        if (char ~= "[a-zA-Z0-9_.~-]") {  ; 这些字符不需要编码
            encoded .= char
        } else if (char = " ") {  ; 空格转换为加号
            encoded .= "+"
        } else {  ; 其他字符进行URL编码
            code := Ord(char)
            if (code < 128) {  ; ASCII字符
                encoded .= "%" . Format("{:02X}", code)
            } else {  ; 非ASCII字符（如中文）
                ; 使用UTF-8编码
                bytes := []
                if (code < 0x800) {
                    bytes.Push(0xC0 | (code >> 6))
                    bytes.Push(0x80 | (code & 0x3F))
                } else if (code < 0x10000) {
                    bytes.Push(0xE0 | (code >> 12))
                    bytes.Push(0x80 | ((code >> 6) & 0x3F))
                    bytes.Push(0x80 | (code & 0x3F))
                } else {
                    bytes.Push(0xF0 | (code >> 18))
                    bytes.Push(0x80 | ((code >> 12) & 0x3F))
                    bytes.Push(0x80 | ((code >> 6) & 0x3F))
                    bytes.Push(0x80 | (code & 0x3F))
                }
                
                for _, b in bytes {
                    encoded .= "%" . Format("{:02X}", b)
                }
            }
        }
    }
    
    return encoded
}
#HotIf

; =====================================================================
; 脚本重启功能
; =====================================================================

CapsLock & r Up::
{
    ; 检查Alt键是否按下
    if (GetKeyState("Alt", "P")) {
        ; 标记为按下了其他键
        global otherKeyPressed := true
        
        ; 显示重启提示
        ToolTip("正在重启脚本...", , , 1)
        
        ; 延迟一下，让用户看到提示
        Sleep(500)
        
        ; 清除提示
        ToolTip(, , , 1)
        
        ; 重启脚本
        Reload()
    } else {
        ; 如果Alt没有按下，执行原来的行尾移动功能
        global otherKeyPressed := true
        MoveEnd()
    }
}

;=====================================================================
; 快捷菜单系统 - 核心功能
;=====================================================================

; 全局变量 - 跟踪当前菜单实例
global currentMenuGui := 0  ; 当前显示的菜单GUI句柄
global currentMenuGroup := 0  ; 当前显示的菜单组索引
global checkActiveTimerId := 0  ; 检查窗口活动状态的定时器ID
global forceKeepMenu := false  ; 强制保持菜单打开的标志（用于截图等操作）

; 全局设置
global MenuSettings := {
    DarkMode: IniRead("CapsLock++.ini", "MenuGroupsColourMode", "DarkMode", "true") = "true",           ; true=深色模式, false=浅色模式
    FontName: "Segoe UI",     ; 字体名称
    FontSize: 10,             ; 字体大小
    
    ; 深色模式颜色
    DarkColors: {
        Background: "0x202020",       ; 背景色
        Text: "0xFFFFFF",             ; 文字颜色
        Border: "0x505050",           ; 边框颜色
        TitleBg: "0x303030",          ; 标题背景色
        TitleText: "0xFFFFFF"         ; 标题文字颜色
    },
    
    ; 浅色模式颜色
    LightColors: {
        Background: "0xF0F0F0",       ; 背景色
        Text: "0x000000",             ; 文字颜色
        Border: "0xD0D0D0",           ; 边框颜色
        TitleBg: "0xFFFFFF",          ; 标题背景色
        TitleText: "0x000000"         ; 标题文字颜色
    }
}

; 检查菜单窗口是否活动的函数 - 修复失去焦点功能
CheckActiveWindow(*) {
    global currentMenuGui, forceKeepMenu
    
    ; 如果设置了强制保持菜单打开的标志，则不关闭菜单
    if (forceKeepMenu)
        return
    
    ; 检查菜单是否存在
    if (!currentMenuGui || !WinExist("ahk_id " currentMenuGui))
        return
    
    ; 尝试获取当前活动窗口，如果失败则关闭菜单
    try {
        activeWin := WinGetID("A")
        
        ; 如果当前活动窗口不是菜单窗口，立即关闭菜单
        if (activeWin != currentMenuGui) {
            ; 关闭菜单，不再检查截图工具
            CloseMenu()
        }
    } catch Error as e {
        ; 如果获取活动窗口时出错，关闭菜单
        CloseMenu()
    }
}

; 读取INI文件的值，并使用UTF-8编码
ReadIniValueUTF8(filePath, section, key, defaultValue := "") {
    try {
        ; 以 UTF-8 编码读取整个文件
        fileContent := FileRead(filePath, "UTF-8")
        
        ; 使用正则表达式匹配指定节和键
        sectionPattern := "\[" section "\]\s*(?:\r?\n|\r)([^\[]*)"
        if RegExMatch(fileContent, sectionPattern, &sectionMatch) {
            sectionContent := sectionMatch[1]
            
            keyPattern := "(?m)^\s*" key "\s*=\s*(.*?)(?:\r?\n|\r|$)"
            if RegExMatch(sectionContent, keyPattern, &keyMatch) {
                return Trim(keyMatch[1])
            }
        }
        return defaultValue
    } catch {
        return defaultValue
    }
}

; 定义菜单组
; 每个组包含:
;   - name: 组名称(显示在标题)
;   - items: 菜单项目列表
; 每个项目包含:
;   - name: 项目名称
;   - icon: 图标(Emoji、FontAwesome字符或ICO文件路径)
;   - iconType: 图标类型("emoji"或"file")
;   - action: 执行的操作(函数引用)

; 从 INI 读取菜单组数量
global menuGroupNum := ReadIniValueUTF8("CapsLock++.ini", "MenuGroupNum", "num", "0")
menuGroupNum := Integer(menuGroupNum)

; 先初始化数组
global enableGroup := []
global groupName := []
global groupCount := []

; 读取每个组的启用状态
Loop 10 {
    i := A_Index
    ; 使用Push方法而不是索引赋值，这样数组会自动扩展
    enableGroup.Push(ReadIniValueUTF8("CapsLock++.ini", "MenuGroupsEnable", "enableGroup" i, "false") = "true")
}

; 读取每个组的名称
Loop 10 {
    i := A_Index
    groupName.Push(ReadIniValueUTF8("CapsLock++.ini", "MenuGroupName", "name" i, "组 " i))
}

; 读取每个组的项目数量
Loop 10 {
    i := A_Index
    count := ReadIniValueUTF8("CapsLock++.ini", "MenuGroupCount", "count" i, "0")
    groupCount.Push(Integer(count))
}

; 菜单项目配置
global MenuGroups := []

; 构建菜单组
Loop 10 {
    groupIndex := A_Index
    
    ; 如果该组启用
    if (enableGroup[groupIndex]) {
        ; 创建菜单项数组
        menuItems := []
        
        ; 为该组读取所有菜单项
        Loop groupCount[groupIndex] {
            itemIndex := A_Index
            sectionName := "MenuGroups" groupIndex "Items"
            
            ; 读取项目属性
            itemName := ReadIniValueUTF8("CapsLock++.ini", sectionName, "name" itemIndex, "")
            itemIcon := ReadIniValueUTF8("CapsLock++.ini", sectionName, "icon" itemIndex, "")
            itemIconType := ReadIniValueUTF8("CapsLock++.ini", sectionName, "icontype" itemIndex, "")
            itemActionStr := ReadIniValueUTF8("CapsLock++.ini", sectionName, "action" itemIndex, "")
            
            ; 只有当有名称和图标时才添加项目
            if (itemName != "" && itemIcon != "") {
                ; 提取action函数
                itemAction := GetActionFromString(itemActionStr)
                
                ; 添加到菜单项数组
                menuItems.Push({
                    name: itemName,
                    icon: itemIcon,
                    iconType: itemIconType,
                    action: itemAction
                })
            }
        }
        
        ; 只有当有菜单项时才添加该组
        if (menuItems.Length > 0) {
            MenuGroups.Push({
                name: groupName[groupIndex],
                items: menuItems
            })
        }
    }
}

; 获取函数引用的辅助函数
GetActionFromString(actionStr) {
    ; 先去除首尾可能存在的空格
    actionStr := Trim(actionStr)

    switch actionStr {
        ; ... (其他 case 保持不变) ...
        case "SetPowerPlan(`"节电`")": return (*) => SetPowerPlan("节电")
        case "SetPowerPlan(`"平衡`")": return (*) => SetPowerPlan("平衡")
        case "SetPowerPlan(`"性能`")": return (*) => SetPowerPlan("性能")
        case "ManageProcessWithCtrlCheck(`"启用`")": return (*) => ManageProcessWithCtrlCheck("启用")
        case "ManageProcessWithCtrlCheck(`"终止`")": return (*) => ManageProcessWithCtrlCheck("终止")
        case "GitHubAccelerate()": return (*) => GitHubAccelerate()
        case "SendInput(`"#d`")": return (*) => SendInput("#d")
        case "WebsiteLogin()": return (*) => WebsiteLogin()

        ; ActivateOrRun 函数处理
        default:
            ; 匹配 ActivateOrRun("param1", param2_content)
            pattern := 'i)^ActivateOrRun\(\s*"(.*?)"\s*,\s*(.*?)\s*\)$' ; 捕获逗号后的所有内容

            if RegExMatch(actionStr, pattern, &match) {
                param1 := match[1] ; 第一个参数 (引号内的内容)
                param2Raw := Trim(match[2]) ; 第二个参数的原始内容

                param2 := "" ; 初始化最终的第二个参数值

                ; --- 解析第二个参数 ---
                if (param2Raw = "A_MyDocuments") {
                    param2 := A_MyDocuments
                } else if (param2Raw = "A_UserProfile") {
                    param2 := EnvGet("USERPROFILE") ; 使用 EnvGet 函数获取用户配置路径
                } else if (SubStr(param2Raw, 1, StrLen("A_UserProfile")) = "A_UserProfile") {
                    ; 处理 A_UserProfile 开头的情况
                    pathPart := Trim(SubStr(param2Raw, StrLen("A_UserProfile") + 1))
                    ; --- 使用 SubStr 移除首尾引号 ---
                    if (SubStr(pathPart, 1, 1) = '"')
                        pathPart := SubStr(pathPart, 2)
                    if (SubStr(pathPart, -0) = '"') ; 检查最后一个字符
                        pathPart := SubStr(pathPart, 1, -1)
                    ; --- End SubStr ---
                    param2 := EnvGet("USERPROFILE") . pathPart ; 使用 EnvGet 拼接
                } else if (SubStr(param2Raw, 1, StrLen("A_MyDocuments")) = "A_MyDocuments") {
                    ; 处理 A_MyDocuments 开头的情况
                    pathPart := Trim(SubStr(param2Raw, StrLen("A_MyDocuments") + 1))
                     ; --- 使用 SubStr 移除首尾引号 ---
                    if (SubStr(pathPart, 1, 1) = '"')
                        pathPart := SubStr(pathPart, 2)
                    if (SubStr(pathPart, -0) = '"') ; 检查最后一个字符
                        pathPart := SubStr(pathPart, 1, -1)
                    ; --- End SubStr ---
                    param2 := A_MyDocuments . pathPart ; 正确拼接
                } else {
                     ; 认为是普通字符串参数，移除首尾引号
                     pathPart := param2Raw ; 临时变量
                     ; --- 使用 SubStr 移除首尾引号 ---
                     if (SubStr(pathPart, 1, 1) = '"')
                        pathPart := SubStr(pathPart, 2)
                     if (SubStr(pathPart, -0) = '"') ; 检查最后一个字符
                        pathPart := SubStr(pathPart, 1, -1)
                     ; --- End SubStr ---
                     param2 := pathPart
                }

                ; 返回绑定了正确参数的函数
                return (*) => ActivateOrRun(param1, param2)

            } else {
                 ; 添加调试信息，看看哪些 actionStr 没有匹配成功
                 ToolTip("GetActionFromString 无法解析: " actionStr)
                 SetTimer(() => ToolTip(), -3000)
                ; 默认空函数
                return (*) => {}
            }
    }
}

; 菜单配置重载函数
ReloadMenuGroups() {
    ; 清空当前配置
    global menuGroupNum, enableGroup, groupName, groupCount, MenuGroups
    
    ; 重新读取菜单组数量
    menuGroupNum := ReadIniValueUTF8("CapsLock++.ini", "MenuGroupNum", "num", "0")
    menuGroupNum := Integer(menuGroupNum)
    
    ; 重置数组
    enableGroup := []
    groupName := []
    groupCount := []
    
    ; 重新读取每个组的启用状态
    Loop 10 {
        i := A_Index
        enableGroup.Push(ReadIniValueUTF8("CapsLock++.ini", "MenuGroupsEnable", "enableGroup" i, "false") = "true")
    }
    
    ; 重新读取每个组的名称
    Loop 10 {
        i := A_Index
        groupName.Push(ReadIniValueUTF8("CapsLock++.ini", "MenuGroupName", "name" i, "组 " i))
    }
    
    ; 重新读取每个组的项目数量
    Loop 10 {
        i := A_Index
        count := ReadIniValueUTF8("CapsLock++.ini", "MenuGroupCount", "count" i, "0")
        groupCount.Push(Integer(count))
    }
    
    ; 清空菜单项目配置
    MenuGroups := []
    
    ; 重新构建菜单组
    Loop 10 {
        groupIndex := A_Index
        
        ; 如果该组启用
        if (enableGroup[groupIndex]) {
            ; 创建菜单项数组
            menuItems := []
            
            ; 为该组读取所有菜单项
            Loop groupCount[groupIndex] {
                itemIndex := A_Index
                sectionName := "MenuGroups" groupIndex "Items"
                
                ; 读取项目属性
                itemName := ReadIniValueUTF8("CapsLock++.ini", sectionName, "name" itemIndex, "")
                itemIcon := ReadIniValueUTF8("CapsLock++.ini", sectionName, "icon" itemIndex, "")
                itemIconType := ReadIniValueUTF8("CapsLock++.ini", sectionName, "icontype" itemIndex, "")
                itemActionStr := ReadIniValueUTF8("CapsLock++.ini", sectionName, "action" itemIndex, "")
                
                ; 只有当有名称和图标时才添加项目
                if (itemName != "" && itemIcon != "") {
                    ; 提取action函数
                    itemAction := GetActionFromString(itemActionStr)
                    
                    ; 添加到菜单项数组
                    menuItems.Push({
                        name: itemName,
                        icon: itemIcon,
                        iconType: itemIconType,
                        action: itemAction
                    })
                }
            }
            
            ; 只有当有菜单项时才添加该组
            if (menuItems.Length > 0) {
                MenuGroups.Push({
                    name: groupName[groupIndex],
                    items: menuItems
                })
            }
        }
    }
    
    ; 显示重载提示
    ToolTip("已重新加载菜单配置, 共加载了 " MenuGroups.Length " 个组")
    SetTimer () => ToolTip(), -2000
}

; 强制清除已有菜单
ClearCapsLockAhkWindows() {
    try {
        ; 获取所有窗口
        winList := WinGetList()
        clearedCount := 0
        
        ; 遍历所有窗口
        for hwnd in winList {
            ; 获取窗口标题
            title := WinGetTitle("ahk_id " hwnd)
            
            ; 检查标题是否包含CapsLock++.ahk
            if (InStr(title, "CapsLock++.ahk")) {
                ; 获取窗口样式
                style := WinGetStyle("ahk_id " hwnd)
                
                ; 检查窗口样式是否匹配0x940A0000
                if (style = 0x940A0000) {
                    ; 关闭匹配的窗口
                    WinClose("ahk_id " hwnd)
                    clearedCount++
                }
            }
        }
    } catch as e {
    }
}

; 显示菜单
ShowMenu(groupIndex) {
    global currentMenuGui, currentMenuGroup
    
    ; 先销毁已有菜单
    ClearCapsLockAhkWindows()
    if (currentMenuGui && WinExist("ahk_id " currentMenuGui)) {
        try {
            WinClose("ahk_id " currentMenuGui)
            Sleep(50)  ; 等待窗口完全关闭
        } catch {
            ; 忽略可能的错误
        }
        currentMenuGui := 0
        currentMenuGroup := 0
    }
    
    ; 检查组索引是否有效
    if (groupIndex < 1 || groupIndex > MenuGroups.Length || !MenuGroups[groupIndex].HasOwnProp("name")) {
        return  ; 无效的组索引，不显示菜单
    }
    
    ; 获取当前组
    currentGroup := MenuGroups[groupIndex]
    
    ; 检查是否有任何项目
    if (!currentGroup.HasOwnProp("items") || currentGroup.items.Length = 0) {
        return  ; 没有项目，不显示菜单
    }
    
    ; 创建菜单GUI
    CreateMenuGUI(currentGroup, groupIndex)
    
    ; 保存当前菜单组索引
    currentMenuGroup := groupIndex
}

GetCharWidthMap() {
    charWidthMap := Map()
    charWidthMap[" "] := 4
    charWidthMap["!"] := 4
    charWidthMap["`""] := 6
    charWidthMap["#"] := 9
    charWidthMap["$"] := 8
    charWidthMap["%"] := 12
    charWidthMap["&"] := 0
    charWidthMap["'"] := 4
    charWidthMap["("] := 5
    charWidthMap[")"] := 5
    charWidthMap["*"] := 6
    charWidthMap["+"] := 10
    charWidthMap[","] := 4
    charWidthMap["-"] := 6
    charWidthMap["."] := 4
    charWidthMap["/"] := 6
    charWidthMap["0"] := 8
    charWidthMap["1"] := 8
    charWidthMap["2"] := 8
    charWidthMap["3"] := 8
    charWidthMap["4"] := 8
    charWidthMap["5"] := 8
    charWidthMap["6"] := 8
    charWidthMap["7"] := 8
    charWidthMap["8"] := 8
    charWidthMap["9"] := 8
    charWidthMap[":"] := 4
    charWidthMap[";"] := 4
    charWidthMap["<"] := 10
    charWidthMap["="] := 10
    charWidthMap[">"] := 10
    charWidthMap["?"] := 7
    charWidthMap["@"] := 14
    charWidthMap["A"] := 10
    charWidthMap["B"] := 9
    charWidthMap["C"] := 9
    charWidthMap["D"] := 11
    charWidthMap["E"] := 8
    charWidthMap["F"] := 7
    charWidthMap["G"] := 10
    charWidthMap["H"] := 11
    charWidthMap["I"] := 4
    charWidthMap["J"] := 6
    charWidthMap["K"] := 9
    charWidthMap["L"] := 7
    charWidthMap["M"] := 13
    charWidthMap["N"] := 11
    charWidthMap["O"] := 11
    charWidthMap["P"] := 9
    charWidthMap["Q"] := 12
    charWidthMap["R"] := 9
    charWidthMap["S"] := 8
    charWidthMap["T"] := 8
    charWidthMap["U"] := 10
    charWidthMap["V"] := 9
    charWidthMap["W"] := 14
    charWidthMap["X"] := 9
    charWidthMap["Y"] := 9
    charWidthMap["Z"] := 9
    charWidthMap["["] := 5
    charWidthMap["\\"] := 6
    charWidthMap["]"] := 5
    charWidthMap["^"] := 10
    charWidthMap["_"] := 7
    charWidthMap["a"] := 8
    charWidthMap["b"] := 9
    charWidthMap["c"] := 7
    charWidthMap["d"] := 9
    charWidthMap["e"] := 8
    charWidthMap["f"] := 5
    charWidthMap["g"] := 9
    charWidthMap["h"] := 9
    charWidthMap["i"] := 4
    charWidthMap["j"] := 4
    charWidthMap["k"] := 8
    charWidthMap["l"] := 4
    charWidthMap["m"] := 13
    charWidthMap["n"] := 9
    charWidthMap["o"] := 9
    charWidthMap["p"] := 9
    charWidthMap["q"] := 9
    charWidthMap["r"] := 6
    charWidthMap["s"] := 7
    charWidthMap["t"] := 5
    charWidthMap["u"] := 9
    charWidthMap["v"] := 7
    charWidthMap["w"] := 11
    charWidthMap["x"] := 7
    charWidthMap["y"] := 8
    charWidthMap["z"] := 7
    charWidthMap["{"] := 5
    charWidthMap["|"] := 4
    charWidthMap["}"] := 5
    charWidthMap["中"] := 14
    charWidthMap["文"] := 14
    charWidthMap["测"] := 14
    charWidthMap["试"] := 14
    return charWidthMap
}

; 使用字符宽度映射表计算文本的像素宽度
CalculateTextWidth(text, charWidthMap) {
    width := 0
    defaultCharWidth := 7  ; 默认宽度，用于未包含在映射表中的字符
    spaceWidth := charWidthMap.Has(" ") ? charWidthMap[" "] : 4 ; 获取空格宽度
    
    for i, char in StrSplit(text) {
        ; Emoji 字符或特殊字符可能不在map里，给个估计宽度
        if (char = " " && spaceWidth > 0) ; 使用测量的空格宽度
             width += spaceWidth
        else if (charWidthMap.Has(char))
            width += charWidthMap[char]
        else if (Ord(char) > 127) ; 粗略判断为宽字符（包括Emoji）
            width += charWidthMap.Has("中") ? charWidthMap["中"] : 14 ; 使用中文宽度或默认值
        else
            width += defaultCharWidth ; 其他ASCII字符使用默认值
    }
    
    return width
}

; 创建菜单GUI
CreateMenuGUI(menuGroup, groupIndex) {
    global currentMenuGui, checkActiveTimerId
    
    ; 获取当前颜色方案
    colors := MenuSettings.DarkMode ? MenuSettings.DarkColors : MenuSettings.LightColors

    ; 创建GUI
    menuGui := Gui("-Caption +AlwaysOnTop +ToolWindow +Owner")
    menuGui.BackColor := colors.Background
    
    ; 保存GUI句柄到全局变量
    currentMenuGui := menuGui.Hwnd
    
    ; 添加标题（组名称）
    menuGui.SetFont("s12 bold c" colors.TitleText, MenuSettings.FontName)
    menuGui.Add("Text", "x10 y10 w280 Center", menuGroup.name)
    
    ; 添加菜单项目
    if (menuGroup.HasOwnProp("items") && menuGroup.items.Length > 0) {
        menuGui.SetFont("s10 c" colors.Text, MenuSettings.FontName)
        
        y := 50
        leftMargin := 35   ; 按钮的左侧边界位置
        
        ; 获取字符宽度映射表
        charWidthMap := GetCharWidthMap()
        spaceWidth := charWidthMap.Has(" ") ? charWidthMap[" "] : 4 ; 获取空格宽度，默认4
        if (spaceWidth <= 0) spaceWidth := 4 ; 防止除零错误
            
        ; 定义固定像素值
        iconOffsetPixels := 0     ; 图标从按钮左边缘的像素偏移量
        iconTextGapPixels := 15    ; 图标和文本之间的固定像素间距
        
        ; 计算最长文本的宽度（以像素为单位）
        maxTextWidthPixels := 0
        for i, item in menuGroup.items {
            if (i > 12)  ; 最多显示12个项目
                continue
            
            textWidth := CalculateTextWidth(item.name, charWidthMap)
            if (textWidth > maxTextWidthPixels)
                maxTextWidthPixels := textWidth
        }
        
        for i, item in menuGroup.items {
            if (i > 12)  ; 最多显示12个项目
                break
            
            ; 创建一个容器面板，用于放置数字和按钮
            panel := menuGui.Add("Text", "x10 y" y " w280 h40 -Background")
            
            ; 添加数字序号（在按钮左侧）
            numColor := "0x000000"
            if (i <= 9) {
                numText := menuGui.Add("Text", "x15 y" y+8 " w20 h24 BackgroundTrans", i)
                numText.SetFont("s14 bold c" numColor, MenuSettings.FontName)
            } else if (i = 10) {
                numText := menuGui.Add("Text", "x15 y" y+8 " w20 h24 BackgroundTrans", "0")
                numText.SetFont("s14 bold c" numColor, MenuSettings.FontName)
            }
            
            ; --- 精确计算空格 ---
            
            ; 1. 计算前缀空格，用于定位图标
            prefixSpaceCount := Ceil(iconOffsetPixels / spaceWidth)
            prefixSpaces := ""
            Loop prefixSpaceCount {
                prefixSpaces .= " "
            }
            
            ; 2. 计算图标和文本之间的间隙空格
            gapSpaceCount := Ceil(iconTextGapPixels / spaceWidth)
            gapSpaces := ""
            Loop gapSpaceCount {
                gapSpaces .= " "
            }
            
            ; 3. 计算后缀空格，使所有文本右边缘对齐
            ;    总宽度 = 图标偏移 + 图标宽度(估算) + 间隙 + 最大文本宽度
            ;    这里简化：总宽度 = 图标偏移 + 间隙 + 最大文本宽度 (忽略图标宽度对后缀的影响，因为图标宽度变化不大)
            totalContentWidthPixels := iconOffsetPixels + iconTextGapPixels + maxTextWidthPixels
            currentTextWidth := CalculateTextWidth(item.name, charWidthMap)
            currentTotalWidth := iconOffsetPixels + iconTextGapPixels + currentTextWidth
            textGapPixels := totalContentWidthPixels - currentTotalWidth
            suffixSpaceCount := Ceil(textGapPixels / spaceWidth)
            
            ; 创建后缀空格
            suffixSpaces := ""
            Loop suffixSpaceCount {
                suffixSpaces .= " "
            }
            
            ; --- 添加按钮 ---
            if (item.HasOwnProp("iconType") && item.iconType = "file") {
                ; 使用ICO文件作为图标
                
                ; 组合文本：前缀空格 + 间隙空格 + 文本 + 后缀空格
                fullBtnText := prefixSpaces gapSpaces item.name suffixSpaces
                
                btn := menuGui.Add("Button", "x" leftMargin " y" y " w255 h40 -TabStop", fullBtnText)
                try {
                    ; 尝试加载图标 (图标会显示在第一个非空格字符前，即gapSpaces之前)
                    btn.SetFont("s10 c" colors.Text, MenuSettings.FontName)
                    AddButtonIcon(btn, item.icon)
                } catch Error as e {
                    ; 加载失败，只显示文本
                    btn.Text := fullBtnText
                }
            } else {
                ; 使用Emoji作为图标
                ; 组合文本：前缀空格 + Emoji + 间隙空格 + 文本 + 后缀空格
                fullBtnText := prefixSpaces item.icon gapSpaces item.name suffixSpaces
                btn := menuGui.Add("Button", "x" leftMargin " y" y " w255 h40 -TabStop", fullBtnText)
            }
            
            btn.OnEvent("Click", item.action)
            
            y += 50
        }
    }
    
    ; 添加关闭按钮
    y += 10
    closeBtn := menuGui.Add("Button", "x10 y" y " w280 h30", "关闭")
    closeBtn.OnEvent("Click", (*) => CloseMenu())
    
    ; 计算GUI高度
    guiHeight := y + 40
    
    ; 定义菜单宽度
    menuWidth := 300
    
    ; ===== 其余部分保持不变 =====
    
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mouseX, &mouseY)
    MonitorGetWorkArea(1, &WLeft, &WTop, &WRight, &WBottom)
    menuCenterX := menuWidth
    menuCenterY := guiHeight
    guiX := mouseX - menuCenterX
    guiY := mouseY - menuCenterY
    
    if (showDebugTooltips) {
        ToolTip("鼠标位置: X=" mouseX " Y=" mouseY "`n"
          . "菜单尺寸: W=" menuWidth " H=" guiHeight "`n"
          . "菜单中心点: X=" (guiX + menuCenterX) " Y=" (guiY + menuCenterY))
        SetTimer () => ToolTip(), -3000
    }
    
    if (guiX + 2 * menuWidth > WRight)
        guiX := WRight - 2 * menuWidth
    if (guiX < WLeft)
        guiX := WLeft
    if (guiY + 2 * guiHeight > WBottom)
        guiY := WBottom - 2 * guiHeight
    if (guiY < WTop)
        guiY := WTop
    
    menuGui.Show("x" guiX " y" guiY " w" menuWidth " h" guiHeight)
    menuGui.OnEvent("Escape", (*) => CloseMenu())
    checkActiveTimerId := SetTimer(CheckActiveWindow, 50)
    menuGui.OnEvent("ContextMenu", (*) => CloseMenu())
}

; 关闭菜单
CloseMenu() {
    global currentMenuGui, currentMenuGroup, checkActiveTimerId, forceKeepMenu
    
    ; 强制清除已有菜单
    ClearCapsLockAhkWindows()

    ; 重置强制保持菜单打开的标志
    forceKeepMenu := false
    
    if (currentMenuGui && WinExist("ahk_id " currentMenuGui)) {
        WinClose("ahk_id " currentMenuGui)
    }
    
    currentMenuGui := 0
    currentMenuGroup := 0
    
    ; 停止定时器
    if (checkActiveTimerId) {
        SetTimer(checkActiveTimerId, 0)
        checkActiveTimerId := 0
    }
}

; 执行菜单项操作
ExecuteMenuItem(groupIndex, itemIndex) {
    ; 检查组索引是否有效
    if (groupIndex < 1 || groupIndex > MenuGroups.Length || !MenuGroups[groupIndex].HasOwnProp("name")) {
        return  ; 无效的组索引
    }
    
    ; 获取组
    menuGroup := MenuGroups[groupIndex]
    
    ; 检查项目索引是否有效
    if (!menuGroup.HasOwnProp("items") || itemIndex < 1 || itemIndex > menuGroup.items.Length) {
        return  ; 无效的项目索引
    }
    
    ; 执行操作
    menuGroup.items[itemIndex].action()
}

;=====================================================================
; 数字键绑定 - 显示对应的菜单或执行操作
;=====================================================================

#HotIf GetKeyState("CapsLock", "P")
1::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global currentMenuGroup
    
    ; 显示第1组菜单（无论当前是否已有菜单）
    ShowMenu(1)
}

2::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global currentMenuGroup
    
    ; 显示第2组菜单（无论当前是否已有菜单）
    ShowMenu(2)
}

3::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global currentMenuGroup
    
    ; 显示第3组菜单（无论当前是否已有菜单）
    ShowMenu(3)
}

4::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global currentMenuGroup
    
    ; 显示第4组菜单（无论当前是否已有菜单）
    ShowMenu(4)
}

5::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global currentMenuGroup
    
    ; 显示第5组菜单（无论当前是否已有菜单）
    ShowMenu(5)
}

6::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global currentMenuGroup
    
    ; 显示第6组菜单（无论当前是否已有菜单）
    ShowMenu(6)
}

7::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global currentMenuGroup
    
    ; 显示第7组菜单（无论当前是否已有菜单）
    ShowMenu(7)
}

8::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global currentMenuGroup
    
    ; 显示第8组菜单（无论当前是否已有菜单）
    ShowMenu(8)
}

9::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global currentMenuGroup
    
    ; 显示第9组菜单（无论当前是否已有菜单）
    ShowMenu(9)
}

0::
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    global currentMenuGroup
    
    ; 显示第10组菜单（无论当前是否已有菜单）
    ShowMenu(10)
}

; 添加单独的数字键处理（不带CapsLock）
#HotIf WinActive("ahk_id " currentMenuGui)
1::
{
    global currentMenuGroup
    
        ; 添加错误处理和有效性检查
        try {
            ; 确保currentMenuGroup是有效值
            if (currentMenuGroup > 0 && currentMenuGroup <= MenuGroups.Length) {
                ; 如果当前菜单组有足够的项目，执行第1个操作
                if (MenuGroups[currentMenuGroup].HasOwnProp("items") && MenuGroups[currentMenuGroup].items.Length >= 1) {
                    ExecuteMenuItem(currentMenuGroup, 1)
                }
            }
        } catch Error as e {
        ; 忽略错误，菜单可能尚未完全加载
    }
}

2::
{
    global currentMenuGroup
    
    try {
        ; 确保currentMenuGroup是有效值
        if (currentMenuGroup > 0 && currentMenuGroup <= MenuGroups.Length) {
            ; 如果当前菜单组有足够的项目，执行第2个操作
            if (MenuGroups[currentMenuGroup].HasOwnProp("items") && MenuGroups[currentMenuGroup].items.Length >= 2) {
                ExecuteMenuItem(currentMenuGroup, 2)
            }
        }
    } catch Error as e {
        ; 忽略错误，菜单可能尚未完全加载
    }
}

3::
{
    global currentMenuGroup
    try {
        ; 确保currentMenuGroup是有效值
        if (currentMenuGroup > 0 && currentMenuGroup <= MenuGroups.Length) {
            ; 如果当前菜单组有足够的项目，执行第3个操作
            if (MenuGroups[currentMenuGroup].HasOwnProp("items") && MenuGroups[currentMenuGroup].items.Length >= 3) {
                ExecuteMenuItem(currentMenuGroup, 3)
            }
        }
    } catch Error as e {
        ; 忽略错误，菜单可能尚未完全加载
    }
}

4::
{
    global currentMenuGroup
    
    try {
        ; 确保currentMenuGroup是有效值
        if (currentMenuGroup > 0 && currentMenuGroup <= MenuGroups.Length) {
            ; 如果当前菜单组有足够的项目，执行第4个操作
            if (MenuGroups[currentMenuGroup].HasOwnProp("items") && MenuGroups[currentMenuGroup].items.Length >= 4) {
                ExecuteMenuItem(currentMenuGroup, 4)
            }
        }
    } catch Error as e {
        ; 忽略错误，菜单可能尚未完全加载
    }
}

5::
{
    global currentMenuGroup
    
    try {
        ; 确保currentMenuGroup是有效值
        if (currentMenuGroup > 0 && currentMenuGroup <= MenuGroups.Length) {
            ; 如果当前菜单组有足够的项目，执行第5个操作
            if (MenuGroups[currentMenuGroup].HasOwnProp("items") && MenuGroups[currentMenuGroup].items.Length >= 5) {
                ExecuteMenuItem(currentMenuGroup, 5)
            }
        }
    } catch Error as e {
        ; 忽略错误，菜单可能尚未完全加载
    }
}

6::
{
    global currentMenuGroup
    
    try {
        ; 确保currentMenuGroup是有效值
        if (currentMenuGroup > 0 && currentMenuGroup <= MenuGroups.Length) {
            ; 如果当前菜单组有足够的项目，执行第6个操作
            if (MenuGroups[currentMenuGroup].HasOwnProp("items") && MenuGroups[currentMenuGroup].items.Length >= 6) {
                ExecuteMenuItem(currentMenuGroup, 6)
            }
        }
    } catch Error as e {
        ; 忽略错误，菜单可能尚未完全加载
    }
}

7::
{
    global currentMenuGroup
    
    try {
        ; 确保currentMenuGroup是有效值
        if (currentMenuGroup > 0 && currentMenuGroup <= MenuGroups.Length) {
            ; 如果当前菜单组有足够的项目，执行第7个操作
            if (MenuGroups[currentMenuGroup].HasOwnProp("items") && MenuGroups[currentMenuGroup].items.Length >= 7) {
                ExecuteMenuItem(currentMenuGroup, 7)
            }
        }
    } catch Error as e {
        ; 忽略错误，菜单可能尚未完全加载
    }
}

8::
{
    global currentMenuGroup
    
    try {
        ; 确保currentMenuGroup是有效值
        if (currentMenuGroup > 0 && currentMenuGroup <= MenuGroups.Length) {
            ; 如果当前菜单组有足够的项目，执行第8个操作
            if (MenuGroups[currentMenuGroup].HasOwnProp("items") && MenuGroups[currentMenuGroup].items.Length >= 8) {
                ExecuteMenuItem(currentMenuGroup, 8)
            }
        }
    } catch Error as e {
        ; 忽略错误，菜单可能尚未完全加载
    }
}

9::
{
    global currentMenuGroup
    
    try {
        ; 确保currentMenuGroup是有效值
        if (currentMenuGroup > 0 && currentMenuGroup <= MenuGroups.Length) {
            ; 如果当前菜单组有足够的项目，执行第9个操作
            if (MenuGroups[currentMenuGroup].HasOwnProp("items") && MenuGroups[currentMenuGroup].items.Length >= 9) {
                ExecuteMenuItem(currentMenuGroup, 9)
            }
        }
    } catch Error as e {
        ; 忽略错误，菜单可能尚未完全加载
    }
}

0::
{
    global currentMenuGroup
    
    try {
        ; 确保currentMenuGroup是有效值
        if (currentMenuGroup > 0 && currentMenuGroup <= MenuGroups.Length) {
            ; 如果当前菜单组有足够的项目，执行第10个操作
            if (MenuGroups[currentMenuGroup].HasOwnProp("items") && MenuGroups[currentMenuGroup].items.Length >= 10) {
                ExecuteMenuItem(currentMenuGroup, 10)
            }
        }
    } catch Error as e {
        ; 忽略错误，菜单可能尚未完全加载
    }
}

; ESC键关闭菜单
Escape::CloseMenu()
#HotIf

;=====================================================================
; 自定义函数示例 - 可以在菜单中调用
;=====================================================================

; 检查窗口是否存在并激活，如果不存在则运行程序
; 参数:
;   - processName: 要检查的进程名称、窗口类名或URL
;   - runCommand: 如果窗口不存在，要运行的命令
ActivateOrRun(processName, runCommand) {
    ; 关闭菜单
    CloseMenu()
    
    ; 尝试查找窗口
    windowFound := false
    
    ; 检查是否是URL
    if (InStr(processName, "http://") || InStr(processName, "https://")) {
        ; 处理URL的情况
        url := processName
        
        ; 提取域名作为匹配依据
        domainStart := InStr(url, "://") + 3
        domainEnd := InStr(url, "/", false, domainStart)
        if (domainEnd = 0) {
            domainEnd := StrLen(url) + 1
        }
        domain := SubStr(url, domainStart, domainEnd - domainStart)
        
        ; 尝试查找包含该域名的浏览器窗口
        try {
            ; 获取所有浏览器窗口
            browserWindows := WinGetList("ahk_exe msedge.exe") ; 可以添加其他浏览器如chrome.exe
            
            ; 遍历所有浏览器窗口
            for hwnd in browserWindows {
                try {
                    ; 获取窗口标题
                    winTitle := WinGetTitle("ahk_id " hwnd)
                    
                    ; 检查窗口标题是否包含域名
                    if (InStr(winTitle, domain)) {
                        ; 找到匹配的窗口，激活它
                        WinActivate("ahk_id " hwnd)
                        windowFound := true
                        break
                    }
                } catch Error as e {
                    ; 忽略错误，继续检查下一个窗口
                }
            }
            
            ; 如果没找到匹配的窗口，尝试Chrome浏览器
            if (!windowFound) {
                browserWindows := WinGetList("ahk_exe chrome.exe")
                for hwnd in browserWindows {
                    try {
                        winTitle := WinGetTitle("ahk_id " hwnd)
                        if (InStr(winTitle, domain)) {
                            WinActivate("ahk_id " hwnd)
                            windowFound := true
                            break
                        }
                    } catch Error as e {
                        ; 忽略错误
                    }
                }
            }
            
            ; 如果没找到匹配的窗口，尝试Firefox浏览器
            if (!windowFound) {
                browserWindows := WinGetList("ahk_exe firefox.exe")
                for hwnd in browserWindows {
                    try {
                        winTitle := WinGetTitle("ahk_id " hwnd)
                        if (InStr(winTitle, domain)) {
                            WinActivate("ahk_id " hwnd)
                            windowFound := true
                            break
                        }
                    } catch Error as e {
                        ; 忽略错误
                    }
                }
            }
        } catch Error as e {
            ; 忽略错误
        }
    } else if (processName = "explorer.exe") {
        ; 检查是否是资源管理器窗口，需要特殊处理
        ; 尝试查找特定路径的资源管理器窗口
        try {
            ; 获取所有资源管理器窗口
            explorerWindows := WinGetList("ahk_class CabinetWClass")
            
            ; 遍历所有资源管理器窗口
            for hwnd in explorerWindows {
                ; 获取窗口的完整路径
                try {
                    ; 获取窗口标题
                    winTitle := WinGetTitle("ahk_id " hwnd)
                    
                    ; 从runCommand中提取目标路径
                    targetPath := StrReplace(runCommand, "explorer.exe ")
                    
                    ; 获取目标文件夹名称
                    SplitPath(targetPath, &targetFolderName)
                    if (targetFolderName = "") {
                        ; 如果是根目录，获取驱动器名称
                        targetFolderName := targetPath
                    }
                    
                    ; 检查窗口标题是否包含目标文件夹名称
                    if (InStr(winTitle, targetFolderName)) {
                        ; 找到匹配的窗口，激活它
                        WinActivate("ahk_id " hwnd)
                        windowFound := true
                        break
                    }
                } catch Error as e {
                    ; 忽略错误，继续检查下一个窗口
                }
            }
        } catch Error as e {
            ; 忽略错误
        }
    } else {
        ; 对于其他类型的窗口，直接按进程名查找
        try {
            if (WinExist("ahk_exe " processName)) {
                WinActivate("ahk_exe " processName)
                windowFound := true
            }
        } catch Error as e {
            ; 忽略错误
        }
    }
    
    ; 如果没有找到窗口，运行指定的命令
    if (!windowFound) {
        try {
            Run('"' runCommand '"') 
        } catch Error as e {
            ShowTooltip("无法启动程序: `nAction: <" runCommand ">`nError: " e.Message)
        }
    }
}

; 调整系统音量
AdjustVolume(amount) {
    Send("{Volume_Up " amount "}")
}

; 打开特定文件夹
OpenFolder(path) {
    Run("explorer.exe " path)
}

; 启动应用程序
LaunchApp(exePath) {
    Run(exePath)
}

; 打开网站
OpenWebsite(url) {
    Run(url)
}

; 执行系统命令
ExecuteCommand(cmd) {
    Run(A_ComSpec " /c " cmd)
}

;=====================================================================
; 集成到CapsLock++的注意事项
;=====================================================================
; 1. 复制全部代码到CapsLock++.ahk文件中
; 2. 确保otherKeyPressed变量在CapsLock++中已定义
; 3. 如果需要调用CapsLock++中的函数(如CleanupWorkspaceHotkey)，
;    确保这些函数在MenuGroups定义之前已经定义
; 4. 数字键的绑定部分可以直接替换CapsLock++中对应的部分
; 5. 根据需要修改MenuGroups中的项目定义
;=====================================================================

;=====================================================================
; 图标路径说明
;=====================================================================
; 图标路径说明:
; 1. 相对路径: "icons/document.ico" 表示脚本所在目录下的icons文件夹中的document.ico文件
; 2. 绝对路径: "C:\Windows\System32\shell32.dll" 可以直接使用系统图标
;
; Windows系统图标位置:
; 1. 系统图标库: C:\Windows\System32\shell32.dll (包含大量系统图标)
; 2. 文档图标: C:\Windows\System32\shell32.dll,3 (第4个图标)
; 3. 下载图标: C:\Windows\System32\shell32.dll,147 (第148个图标)
; 4. 桌面图标: C:\Windows\System32\shell32.dll,34 (第35个图标)
; 5. 图片图标: C:\Windows\System32\shell32.dll,132 (第133个图标)
;
; 使用系统图标的方法:
; 1. 创建icons文件夹在脚本目录下
; 2. 使用资源提取工具(如Resource Hacker)从shell32.dll提取需要的图标
; 3. 或者直接使用格式: "C:\Windows\System32\shell32.dll,图标索引"
;=====================================================================

; 添加按钮图标的辅助函数
AddButtonIcon(ButtonCtrl, IconFile, IconNumber := 1, IconSize := 30) {
    ; 检查是否是DLL文件格式（包含逗号分隔的图标索引）
    if (InStr(IconFile, ",")) {
        ; 分割文件路径和图标索引
        parts := StrSplit(IconFile, ",")
        if (parts.Length >= 2) {
            dllPath := Trim(parts[1])
            iconIndex := Trim(parts[2])
            
            ; 加载DLL中的图标
            hIcon := LoadPicture(dllPath, "w" IconSize " h" IconSize " Icon" iconIndex, &imgType)
        } else {
            ; 如果格式不正确，尝试直接加载
            hIcon := LoadPicture(IconFile, "w" IconSize " h" IconSize, &imgType)
        }
    } else {
        ; 普通图标文件
        hIcon := LoadPicture(IconFile, "w" IconSize " h" IconSize, &imgType)
    }
    
    ; 发送消息设置图标
    SendMessage(0xF7, IconNumber, hIcon, ButtonCtrl)  ; BM_SETIMAGE
    
    return hIcon
}

;=====================================================================
; 电源计划和进程管理功能
;=====================================================================

; 设置电源计划
SetPowerPlan(planType) {
    ; 关闭菜单
    CloseMenu()

    ; 定义电源计划GUID
    bitsumHighestPerfGuid := "a6a4e4e3-ae4f-40c2-bfc7-7b9feb4f55f8" ; process lasso的最佳性能
    highPerfGuid := "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"          ; 高性能
    balancedGuid := "381b4222-f694-41f0-9685-ff5bb260df2e"          ; 平衡
    powerSaverGuid := "a1841308-3541-4fab-bc81-f71556f20b4a"        ; 节电
    
    ; 根据计划类型设置对应的电源计划
    switch planType {
        case "节电":
            ; 节电模式 - 完全按照原始批处理文件的功能实现
            ; 设置刷新率
            RunWait("C:\Windows\System32\QRes.exe /r:48", , "Hide")
            ; 设置电源计划
            RunWait('"' A_ComSpec '" /c powercfg -setactive ' powerSaverGuid, , "Hide")
            ; 服务管理优化
            RunWait("sc config `"SysMain`" start= disabled", , "Hide")
            RunWait("sc config `"DiagTrack`" start= disabled", , "Hide")
            RunWait("sc config `"WSearch`" start= disabled", , "Hide")
            ; 网络优化
            RunWait("netsh wlan set autoconfig enabled=no interface=`"Wi-Fi`"", , "Hide")
            RunWait("netsh int tcp set global autotuninglevel=disabled", , "Hide")
            RunWait("netsh interface tcp set global rss=disabled", , "Hide")
            ShowTooltip("已切换到节电模式")
        
        case "平衡":
            ; 平衡模式 - 完全按照原始批处理文件的功能实现
            ; 设置刷新率
            RunWait("C:\Windows\System32\QRes.exe /r:120", , "Hide")
            ; 设置电源计划
            RunWait('"' A_ComSpec '" /c powercfg -setactive ' balancedGuid, , "Hide")
            ; 设置休眠文件类型
            RunWait("powercfg /h /type full", , "Hide")
            ; 服务管理
            RunWait("sc config `"SysMain`" start= demand", , "Hide")
            RunWait("sc config `"DiagTrack`" start= auto", , "Hide")
            RunWait("sc config `"DPS`" start= auto", , "Hide")
            ; 网络优化
            RunWait("netsh wlan set autoconfig enabled=yes interface=`"Wi-Fi`"", , "Hide")
            RunWait("netsh int tcp set global autotuninglevel=normal", , "Hide")
            RunWait("netsh int tcp set global dca=enabled", , "Hide")
            ShowTooltip("已切换到平衡模式")
        
        case "性能":
            ; 性能模式 - 完全按照原始批处理文件的功能实现
            ; 设置刷新率
            RunWait("C:\Windows\System32\QRes.exe /r:120", , "Hide")
            ; 尝试设置卓越性能
            exitCode := RunWait('"' A_ComSpec '" /c powercfg -setactive ' bitsumHighestPerfGuid, , "Hide")
            
            tooltipMsg := ""
            if (exitCode != 0) {
                ; 卓越性能设置失败，尝试设置高性能
                exitCodeHigh := RunWait('"' A_ComSpec '" /c powercfg -setactive ' highPerfGuid, , "Hide")
                if (exitCodeHigh = 0) {
                    tooltipMsg := "已切换到高性能模式 (Process lasso不可用)"
                } else {
                    tooltipMsg := "无法切换到卓越或高性能模式"
                }
            } else {
                tooltipMsg := "已切换到卓越性能模式"
            }
            
            ; 设置休眠文件类型
            RunWait("powercfg /h /type full", , "Hide")
            ; 服务管理
            RunWait("sc config `"SysMain`" start= demand", , "Hide")
            RunWait("sc config `"DiagTrack`" start= auto", , "Hide")
            RunWait("sc config `"DPS`" start= auto", , "Hide")
            ; 网络优化
            RunWait("netsh wlan set autoconfig enabled=yes interface=`"Wi-Fi`"", , "Hide")
            RunWait("netsh int tcp set global autotuninglevel=normal", , "Hide")
            RunWait("netsh int tcp set global dca=enabled", , "Hide")
            ShowTooltip("已切换到性能模式")
    }
}

; 进程管理功能
ManageProcess(action) {
    ; 关闭菜单
    CloseMenu()
    
    switch action {
        case "启用":
            ; 启用进程 - 从INI读取列表
            ShowTooltip("正在启用指定进程...")
            i := 1
            Loop {
                processToStart := IniRead("CapsLock++.ini", "ProcessesToStart", "Item" i, "")
                if (processToStart = "")  ; 没有更多条目
                    break
                Run(processToStart)
                i++
            }
            ShowTooltip("已尝试启用指定配置进程")
            
        case "终止":
            ; 终止进程 - 从INI读取列表
            ShowTooltip("正在终止指定进程...")
            i := 1
            terminatedCount := 0
            Loop {
                processToTerminate := IniRead("CapsLock++.ini", "ProcessesToTerminate", "Item" i, "")
                if (processToTerminate = "")  ; 没有更多条目
                    break
                try {
                    ProcessClose(processToTerminate)
                    terminatedCount++
                } catch Error as e {
                    ; 忽略终止错误，可能是进程不存在
                }
                i++
            }
            ShowTooltip("已尝试终止指定进程 (" terminatedCount " 个成功)")
    }
}

; GitHub加速功能
; 全局变量，记录当前尝试次数
global githubAccelerateAttempts := 0

GitHubAccelerate() {
    ; 关闭菜单
    CloseMenu()
    
    ; 需要管理员权限检查
    if (!A_IsAdmin) {
        MsgBox("此功能需要管理员权限才能修改hosts文件。`n请以管理员身份运行脚本！", "GitHub加速", "Icon!")
        return
    }
    
    ; 重置尝试次数（如果是通过菜单点击的）
    global githubAccelerateAttempts
    if (githubAccelerateAttempts > 3) {
        githubAccelerateAttempts := 0
    }
    
    hostsFile := "C:\Windows\System32\drivers\etc\hosts"
    hosts := ""
    
    try {
        ; 检查hosts文件大小
        fileObj := FileOpen(hostsFile, "r")
        if (fileObj) {
            fileSize := fileObj.Length
            fileObj.Close()
            
            ; 如果hosts文件超过1MB，提示用户并询问是否清理
            if (fileSize > 1048576) { ; 1MB = 1048576 bytes
                fileSizeMB := Round(fileSize / 1048576, 2)
                result := MsgBox("检测到hosts文件异常大 (" . fileSizeMB . "MB)。`n这可能导致系统性能下降。`n`n是否尝试清理和保留必要内容？", "GitHub加速", "YesNo Icon!")
                if (result = "Yes") {
                    ; 备份原hosts文件
                    backupFile := hostsFile . ".bak_large_" . FormatTime(, "yyyyMMdd_HHmmss")
                    FileCopy(hostsFile, backupFile, true)
                    ShowTooltip("正在清理hosts文件...", 2000)
                    
                    ; 读取现有hosts内容并重建
                    currentHosts := FileRead(hostsFile)
                    
                    ; 尝试提取原始系统hosts部分（保留前30行）
                    originalSystemHosts := ""
                    lines := StrSplit(currentHosts, "`n", "`r")
                    lineCount := Min(30, lines.Length)
                    
                    Loop lineCount {
                        line := lines[A_Index]
                        ; 如果遇到GitHub相关行，停止提取
                        if (InStr(line, "github") || InStr(line, "GitHub") || InStr(line, "GITHUB")) {
                            break
                        }
                        ; 保留系统hosts和常见的注释
                        if (InStr(line, "localhost") || InStr(line, "127.0.0.1") || InStr(line, "::1") ||
                            SubStr(Trim(line), 1, 1) = "#" || Trim(line) = "") {
                            originalSystemHosts .= line . "`n"
                        }
                    }
                    
                    ; 如果未能提取到足够的系统hosts内容，使用默认最小系统hosts
                    if (originalSystemHosts = "") {
                        originalSystemHosts := "# 本地主机名映射`n127.0.0.1 localhost`n::1 localhost`n`n"
                    }
                    
                    ; 完全重建hosts文件，而不是尝试修改现有文件
                    ; 这样可以避免任何潜在的残留问题
                    newHosts := ""
                    
                    ; 查找"End of section"或类似边界标记的位置
                    endOfSectionPos := InStr(currentHosts, "# End of section")
                    
                    ; 如果找到了边界标记
                    if (endOfSectionPos > 0) {
                        ; 找到边界标记后的第一个换行符
                        endOfLinePos := InStr(currentHosts, "`n", false, endOfSectionPos + 1)
                        
                        ; 如果找到了换行符，则取到该换行符的位置（包含该换行符）
                        ; 否则只保留到边界标记所在行
                        if (endOfLinePos > 0) {
                            systemPart := SubStr(currentHosts, 1, endOfLinePos)
                        } else {
                            systemPart := SubStr(currentHosts, 1, endOfSectionPos + 16) . "`n"  ; "# End of section"长度为16
                        }
                        
                        ; 保留系统部分
                        newHosts := systemPart
                    } else {
                        ; 如果没有找到边界标记，则尝试提取必要的系统条目
                        systemEntries := []
                        
                        ; 从当前hosts文件中提取必要条目
                        Loop Parse, currentHosts, "`n", "`r"
                        {
                            line := Trim(A_LoopField)
                            
                            ; 跳过空行和GitHub相关行
                            if (line = "" || RegExMatch(line, "i)(github|githubusercontent|githubassets|vscode\.dev)")) {
                                continue
                            }
                            
                            ; 只保留localhost相关或不包含特殊字符的行
                            if (InStr(line, "localhost") || 
                                RegExMatch(line, "^127\.0\.0\.1") ||
                                RegExMatch(line, "^::1") ||
                                !RegExMatch(line, "[^\x00-\x7F]"))  ; 不包含非ASCII字符
                            {
                                ; 仅当不是GitHub相关的才添加
                                if (!RegExMatch(line, "i)github")) {
                                    systemEntries.Push(line)
                                }
                            }
                        }
                        
                        ; 如果没有找到必要的localhost条目，添加默认值
                        hasLocalhost := false
                        for _, entry in systemEntries {
                            if (InStr(entry, "localhost")) {
                                hasLocalhost := true
                                break
                            }
                        }
                        
                        if (!hasLocalhost) {
                            systemEntries.InsertAt(1, "127.0.0.1 localhost")
                            systemEntries.InsertAt(2, "::1 localhost")
                        }
                        
                        ; 构建系统部分
                        for _, entry in systemEntries {
                            newHosts .= entry . "`n"
                        }
                        
                        ; 添加一个边界标记，方便将来识别
                        newHosts .= "# End of system entries`n"
                    }
                    
                    ; 确保在系统条目和GitHub条目之间有一个空行
                    if (SubStr(newHosts, -1) != "`n") {
                        newHosts .= "`n"
                    }
                    newHosts .= "`n"
                    
                    ; 然后添加GitHub加速条目
                    newHosts .= "# GitHub加速 - 更新时间：" . FormatTime(, "yyyy-MM-dd HH:mm:ss") . "`n" . hosts
                    
                    ; 确保新的hosts内容非空且合理
                    if (StrLen(newHosts) < 50) {
                        ShowTooltip("生成的新hosts内容异常，操作取消", 5000)
                        
                        ; 记录错误到日志
                        try {
                            logFile := A_ScriptDir . "\GitHubAccelerate_error.log"
                            FileAppend("错误时间：" . FormatTime(, "yyyy-MM-dd HH:mm:ss") . 
                                      "`n错误信息：生成的新hosts内容异常" .
                                      "`n新hosts长度：" . StrLen(newHosts) .
                                      "`n`n", logFile)
                        } catch {
                            ; 忽略日志写入错误
                        }
                        
                        return
                    }
                    
                    ; 写入新的hosts文件
                    FileDelete(hostsFile)
                    FileAppend(newHosts, hostsFile, "UTF-8-RAW")
                    
                    ; 日志记录成功更新（调试用）
                    try {
                        logFile := A_ScriptDir . "\GitHubAccelerate_debug.log"
                        FileAppend("成功更新hosts文件`n" .
                                   "更新时间：" . FormatTime(, "yyyy-MM-dd HH:mm:ss") . 
                                   "`n新hosts长度：" . StrLen(newHosts) .
                                   "`n`n", logFile)
                    } catch {
                        ; 忽略日志写入错误
                    }

                    ; 刷新DNS缓存
                    ShowTooltip("正在刷新DNS缓存...", 2000)
                    RunWait("ipconfig /flushdns", , "Hide")
                    
                    ; 验证GitHub连接并显示结果
                    ShowTooltip("正在验证GitHub连接...", 2000)
                    if (VerifyGitHubConnection()) {
                        ShowTooltip("GitHub加速成功！`n使用: 在线源`n备份文件: " . backupFile, 5000)
                        
                        ; 重置尝试次数
                        githubAccelerateAttempts := 0
                    } else {
                        ShowTooltip("GitHub加速已应用(使用在线源)，但连接验证失败`n" .
                                  "可尝试再次点击应用或更换网络", 5000)
                    }
                }
            }
        }
        
        ; 备份当前hosts文件
        backupFile := hostsFile . ".bak_" . FormatTime(, "yyyyMMdd_HHmmss")
        FileCopy(hostsFile, backupFile, true)
        
        ; 读取现有hosts内容
        currentHosts := FileRead(hostsFile)
        
        ; 从在线源获取hosts
        ShowTooltip("正在从在线源获取最新GitHub hosts...", 3000)
        hosts := GetOnlineHostsContent()
        
        ; 检查hosts是否为空 (获取失败)
        if (hosts = "") {
            githubAccelerateAttempts += 1
            
            if (githubAccelerateAttempts <= 3) {
                ; 最多尝试3次
                ShowTooltip("在线源获取失败，正在第" . githubAccelerateAttempts . "次重试...", 3000)
                Sleep(1000)  ; 等待1秒后重试
                hosts := GetOnlineHostsContent()
                
                ; 如果重试还是失败
                if (hosts = "") {
                    ShowTooltip("重试获取在线hosts失败，将再次执行GitHub加速功能", 3000)
                    Sleep(2000)
                    GitHubAccelerate()  ; 递归调用自身
                    return
                }
            } else {
                ; 超过重试次数，显示错误并返回
                ShowTooltip("多次尝试获取在线hosts均已失败，请检查网络连接或稍后再试", 5000)
                githubAccelerateAttempts := 0  ; 重置计数器
                return
            }
        }
        
        ; 日志记录获取的hosts内容长度（调试用）
        try {
            logFile := A_ScriptDir . "\GitHubAccelerate_debug.log"
            FileAppend("获取的hosts内容长度：" . StrLen(hosts) . " 字符`n" .
                       "获取时间：" . FormatTime(, "yyyy-MM-dd HH:mm:ss") . 
                       "`n`n", logFile)
        } catch {
            ; 忽略日志写入错误
        }
        
        ; 确保hosts内容不为空且格式正确
        if (StrLen(hosts) < 10) {  ; 设定一个最小长度阈值来判断hosts是否有效
            ShowTooltip("获取的hosts内容无效或过短，操作取消", 5000)
            return
        }
        
        ; 完全重建hosts文件，而不是尝试修改现有文件
        ; 这样可以避免任何潜在的残留问题
        newHosts := ""
        
        ; 提取关键系统条目
        essentialEntries := []
        
        ; 定义必须保留的系统条目模式
        essentialPatterns := [
            "^127\.0\.0\.1\s+localhost",           ; 基本localhost映射
            "^::1\s+localhost",                     ; IPv6 localhost映射
            "^127\.0\.0\.1\s+kubernetes",           ; Kubernetes相关
            "^192\.168\.[0-9]+\.[0-9]+\s+host",     ; Docker相关
            "\.docker\."                            ; 其他Docker相关条目
        ]
        
        ; 从当前hosts文件中提取必要条目
        Loop Parse, currentHosts, "`n", "`r"
        {
            line := Trim(A_LoopField)
            
            ; 跳过空行
            if (line = "") {
                continue
            }
            
            ; 跳过任何包含编码问题或GitHub相关的行
            if (RegExMatch(line, "i)(github|githubusercontent|鏁|閸|鍔|鐠|githubassets|vscode\.dev)")) {
                continue
            }
            
            ; 检查是否是必须保留的系统条目
            isEssential := false
            for _, pattern in essentialPatterns {
                if (RegExMatch(line, pattern)) {
                    isEssential := true
                    break
                }
            }
            
            ; 保留必要的系统条目和非GitHub相关的条目
            if (isEssential || 
                (RegExMatch(line, "^127\.0\.0\.1\s+") && !InStr(line, "github")) || 
                (RegExMatch(line, "^#") && !InStr(line, "github") && !InStr(line, "GitHub"))) {
                
                ; 检查条目是否已存在
                alreadyExists := false
                for _, entry in essentialEntries {
                    if (entry = line) {
                        alreadyExists := true
                        break
                    }
                }
                
                ; 只添加不重复的条目
                if (!alreadyExists) {
                    essentialEntries.Push(line)
                }
            }
        }
        
        ; 如果没有找到必要的localhost条目，添加默认值
        hasLocalhost := false
        for _, entry in essentialEntries {
            if (InStr(entry, "localhost")) {
                hasLocalhost := true
                break
            }
        }
        
        if (!hasLocalhost) {
            essentialEntries.InsertAt(1, "127.0.0.1 localhost")
            essentialEntries.InsertAt(2, "::1 localhost")
        }
        
        ; 构建新的hosts文件
        ; 首先添加系统条目
        for _, entry in essentialEntries {
            newHosts .= entry . "`n"
        }
        
        ; 确保在系统条目和GitHub条目之间有一个空行
        newHosts .= "`n"
        
        ; 然后添加GitHub加速条目
        newHosts .= "# GitHub加速 - 更新时间：" . FormatTime(, "yyyy-MM-dd HH:mm:ss") . "`n" . hosts
        
        ; 确保新的hosts内容非空且合理
        if (StrLen(newHosts) < 50) {
            ShowTooltip("生成的新hosts内容异常，操作取消", 5000)
            
            ; 记录错误到日志
            try {
                logFile := A_ScriptDir . "\GitHubAccelerate_error.log"
                FileAppend("错误时间：" . FormatTime(, "yyyy-MM-dd HH:mm:ss") . 
                          "`n错误信息：生成的新hosts内容异常" .
                          "`n新hosts长度：" . StrLen(newHosts) .
                          "`n`n", logFile)
            } catch {
                ; 忽略日志写入错误
            }
            
            return
        }
        
        ; 写入新的hosts文件
        FileDelete(hostsFile)
        FileAppend(newHosts, hostsFile, "UTF-8-RAW")
        
        ; 日志记录成功更新（调试用）
        try {
            logFile := A_ScriptDir . "\GitHubAccelerate_debug.log"
            FileAppend("成功更新hosts文件`n" .
                       "更新时间：" . FormatTime(, "yyyy-MM-dd HH:mm:ss") . 
                       "`n新hosts长度：" . StrLen(newHosts) .
                       "`n`n", logFile)
        } catch {
            ; 忽略日志写入错误
        }

        ; 刷新DNS缓存
        ShowTooltip("正在刷新DNS缓存...", 2000)
        RunWait("ipconfig /flushdns", , "Hide")
        
        ; 验证GitHub连接并显示结果
        ShowTooltip("正在验证GitHub连接...", 2000)
        if (VerifyGitHubConnection()) {
            ShowTooltip("GitHub加速成功！`n使用: 在线源`n备份文件: " . backupFile, 5000)
            
            ; 重置尝试次数
            githubAccelerateAttempts := 0
        } else {
            ShowTooltip("GitHub加速已应用(使用在线源)，但连接验证失败`n" .
                      "可尝试再次点击应用或更换网络", 5000)
        }
    } catch as e {
        ; 显示错误消息
        errMsg := "更新hosts失败: " . e.Message
        ShowTooltip(errMsg, 5000)
        
        ; 记录错误到日志文件
        try {
            logFile := A_ScriptDir . "\GitHubAccelerate_error.log"
            FileAppend("错误时间：" . FormatTime(, "yyyy-MM-dd HH:mm:ss") . 
                      "`n错误信息：" . e.Message . 
                      "`n`n", logFile)
        } catch {
            ; 忽略日志写入错误
        }
    }
}

; 从在线源获取hosts内容
GetOnlineHostsContent() {
    ; 调用现有的DownloadGitHubHosts函数
    onlineHosts := DownloadGitHubHosts()
    
    if (onlineHosts && StrLen(onlineHosts) > 10) {  ; 确保内容有效
        ; 返回处理后的hosts内容（不使用多行字符串格式以避免可能的格式问题）
        return onlineHosts
    } else {
        return ""
    }
}

; 验证GitHub连接是否成功
VerifyGitHubConnection() {
    try {
        ShowTooltip("正在多种方式验证GitHub连接...", 2000)
        success := false
        
        ; 方法1: 使用ping方式检测，更加可靠
        try {
            pingResult := ""
            RunWait("ping -n 2 -w 1000 github.com", , "Hide", &pingResult)
            
            ; 如果ping能通，则视为连接成功
            if (InStr(pingResult, "TTL=") || InStr(pingResult, "bytes=")) {
                ShowTooltip("Ping验证成功", 1000)
                return true
            }
        } catch {
            ; 忽略ping失败的错误
        }
        
        ; 方法2: 使用curl检测HTTP响应
        try {
            curlResult := ""
            RunWait("curl -s --connect-timeout 3 -I github.com", , "Hide", &curlResult)
            
            if (InStr(curlResult, "HTTP/") && (InStr(curlResult, "200 OK") || InStr(curlResult, "301") || InStr(curlResult, "302"))) {
                ShowTooltip("Curl验证成功", 1000)
                return true
            }
        } catch {
            ; 忽略curl失败的错误
        }
        
        ; 方法3: 使用nslookup检测DNS解析
        try {
            dnsResult := ""
            RunWait("nslookup github.com", , "Hide", &dnsResult)
            
            if (InStr(dnsResult, "Address") && !InStr(dnsResult, "can't find")) {
                ShowTooltip("DNS验证成功", 1000)
                return true
            }
        } catch {
            ; 忽略DNS解析失败的错误
        }
        
        ; 方法4: 使用WinHttp直接请求
        try {
            whr := ComObject("WinHttp.WinHttpRequest.5.1")
            whr.Open("HEAD", "https://github.com", true)
            whr.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
            whr.Send()
            whr.WaitForResponse(5)  ; 等待最多5秒
            
            if (whr.Status >= 200 && whr.Status < 400) {
                ShowTooltip("HTTP请求验证成功", 1000)
                return true
            }
        } catch {
            ; 忽略HTTP请求失败的错误
        }
        
        ; 如果以上方法都失败，则认为连接失败
        return false
    } catch {
        ; 出现未知异常，返回false
        return false
    }
}

; 从TinsFox获取GitHub hosts内容
DownloadGitHubHosts() {
    try {
        ; 使用内置的下载功能获取hosts内容
        url := "https://github-hosts.tinsfox.com/hosts"
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("GET", url, true)
        whr.Send()
        whr.WaitForResponse()
        
        if (whr.Status = 200) {
            content := whr.ResponseText
            
            ; 提取GitHub相关条目，包括注释
            github_content := ""
            dataTimestamp := ""
            inGitHubSection := false
            lines := StrSplit(content, "`n", "`r")
            
            for _, line in lines {
                ; 检查是否包含时间戳信息
                if (InStr(line, "数据更新时间") || InStr(line, "更新时间")) {
                    dataTimestamp := RegExReplace(line, ".*?数据更新时间：\s*(.*?)$", "$1")
                    continue  ; 跳过时间戳行，我们会自己添加
                }
                
                ; 去除行首尾空白
                trimmedLine := Trim(line)
                
                ; 跳过空行
                if (trimmedLine = "") {
                    continue
                }
                
                ; 识别注释行
                isComment := SubStr(trimmedLine, 1, 1) = "#"
                
                ; 检查是否为GitHub相关条目（IP地址+域名或注释）
                isGithubRelated := isComment && InStr(trimmedLine, "github") 
                    || (!isComment && (InStr(trimmedLine, "github") || InStr(trimmedLine, "githubusercontent")))
                
                ; 如果是注释且包含GitHub字样，可能是区块开始
                if (isComment && InStr(trimmedLine, "github")) {
                    inGitHubSection := true
                }
                ; 如果是注释且不包含GitHub字样，可能是区块结束
                else if (isComment && !InStr(trimmedLine, "github")) {
                    inGitHubSection := false
                }
                
                ; 保存GitHub相关条目和注释
                if (isGithubRelated || (inGitHubSection && !isComment)) {
                    github_content .= line . "`n"
                }
            }
            
            ; 添加数据源和时间戳信息
            if (dataTimestamp) {
                github_content := "# 数据来源：github-hosts.tinsfox.com 更新时间：" . dataTimestamp . "`n" . github_content
            } else {
                github_content := "# 数据来源：github-hosts.tinsfox.com`n" . github_content
            }
            
            return github_content
        } else {
            ; 请求失败
            ShowTooltip("获取在线hosts失败，状态码：" . whr.Status, 3000)
            return ""
        }
    } catch as e {
        ; 出现异常
        ShowTooltip("获取在线hosts异常：" . e.Message, 3000)
        return ""
    }
}

;===============================================================
; 网站管理功能
;===============================================================

WebsiteLogin() {
    ; 检测Ctrl键是否被按下
    if (GetKeyState("Ctrl", "P")) {
        ; 如果按下Ctrl键，显示网站选择GUI
        ShowWebsiteSelectionGUI()
    } else {
        ; 否则打开默认网站
        OpenDefaultWebsite()
    }
}

; 打开默认网站
OpenDefaultWebsite() {
    ; 从INI文件读取默认网站（使用UTF-8编码）
    site := ReadIniValueUTF8("CapsLock++.ini", "CommonWebsites", "default_site", "")
    
    ; 如果未找到默认网站，尝试使用Credentials中的site作为备选
    if (site = "") {
        site := ReadIniValueUTF8("CapsLock++.ini", "Credentials", "site", "")
    }
    
    ; 如果仍未找到网站，提示错误
    if (site = "") {
        ShowTooltip("错误: 无法从INI文件读取网站配置")
        return
    }
    
    ; 读取浏览器偏好设置
    browser := ReadIniValueUTF8("CapsLock++.ini", "CommonWebsites", "browser", "edge")
    
    ; 打开网站
    OpenOneWebsite(site, browser)
}

; 打开指定网站
OpenOneWebsite(url, browser := "edge") {
    ; 确保URL格式正确
    if (!InStr(url, "://")) {
        url := "https://" . url
    }
    
    ; 根据浏览器偏好打开
    switch browser {
        case "edge", "msedge":
            Run("msedge.exe " . url)
        case "chrome":
            Run("chrome.exe " . url)
        case "firefox":
            Run("firefox.exe " . url)
        default:
            Run(url)  ; 使用系统默认浏览器
    }
    
    ; ShowTooltip("正在打开: " . url, 2000)
}

; 显示网站选择GUI
ShowWebsiteSelectionGUI() {
    ; 关闭主菜单
    try CloseMenu()
    
    ; 定义网站列表
    websiteList := []
    
    ; 从INI文件读取网站列表
    i := 1
    Loop {
        siteKey := "site" . i
        urlKey := "url" . i
        
        siteName := ReadIniValueUTF8("CapsLock++.ini", "CommonWebsites", siteKey, "")
        ; 如果读不到名称，则认为列表结束
        if (siteName = "")
            break
            
        siteUrl := ReadIniValueUTF8("CapsLock++.ini", "CommonWebsites", urlKey, "")
        if (siteUrl = "")
            continue
            
        ; 确保URL格式正确
        if (!InStr(siteUrl, "://")) {
            siteUrl := "https://" . siteUrl
        }
        
        ; 将读取到的网站信息添加到列表
        websiteList.Push({name: siteName, url: siteUrl, checked: false})
        
        i++
    }
    
    ; 如果列表为空，提示用户并返回
    if (websiteList.Length = 0) {
        ShowTooltip("错误: 无法从INI文件读取网站列表，请先配置[CommonWebsites]部分")
        return
    }

    ; 读取浏览器偏好设置
    browser := ReadIniValueUTF8("CapsLock++.ini", "CommonWebsites", "browser", "edge")

    ; 创建GUI
    websiteGui := Gui("+AlwaysOnTop +ToolWindow")
    websiteGui.Title := "选择要打开的网站"
    
    ; 添加ListView控件
    listView := websiteGui.Add("ListView", "x10 y10 w500 h300 Checked", ["网站名称", "网址"])
    
    ; 填充ListView
    for _, site in websiteList {
        row := listView.Add(site.checked ? "Check" : "", site.name, site.url)
    }
    
    ; 自动调整列宽
    listView.ModifyCol(1, 150)
    listView.ModifyCol(2, "Auto")
    
    ; 添加全选和反选按钮
    btnSelectAll := websiteGui.Add("Button", "x10 y320 w90 h30", "全选")
    btnSelectAll.OnEvent("Click", (*) => SelectAllItems(listView, true))
    
    btnSelectNone := websiteGui.Add("Button", "x110 y320 w90 h30", "全不选")
    btnSelectNone.OnEvent("Click", (*) => SelectAllItems(listView, false))
    
    btnInvert := websiteGui.Add("Button", "x210 y320 w90 h30", "反选")
    btnInvert.OnEvent("Click", (*) => InvertSelection(listView))
    
    ; 添加操作按钮
    btnOk := websiteGui.Add("Button", "x310 y320 w200 h30", "打开选中网站")
    ; 传递 websiteList 和 browser 给处理函数
    btnOk.OnEvent("Click", (*) => WebsiteSelection(listView, websiteList, browser, websiteGui))
    
    ; 添加返回和取消按钮
    btnBack := websiteGui.Add("Button", "x10 y360 w240 h30", "← 返回")
    btnBack.OnEvent("Click", BackToMainMenu)
    
    btnCancel := websiteGui.Add("Button", "x260 y360 w250 h30", "取消")
    btnCancel.OnEvent("Click", (*) => websiteGui.Destroy())
    
    ; 设置双击事件 - 直接操作单个网站
    listView.OnEvent("DoubleClick", (*) => WebsiteSingleItem(listView, websiteList, browser, websiteGui))
    
    ; 添加ESC键退出处理
    websiteGui.OnEvent("Escape", (*) => websiteGui.Destroy())
    websiteGui.OnEvent("Close", (*) => websiteGui.Destroy())
    
    ; 显示GUI并聚焦到ListView
    websiteGui.Show("w520 h400")
    listView.Focus()
    
    ; 默认全不选
    SelectAllItems(listView, false)
}

; 处理网站选择
WebsiteSelection(listView, websiteList, browser, gui) {
    ; 获取所有选中的行
    selectedItems := []
    row := 0
    
    Loop {
        row := listView.GetNext(row, "Checked")
        if (!row)
            break
        selectedItems.Push(websiteList[row])
    }
    
    ; 销毁GUI
    gui.Destroy()
    
    ; 处理选中的网站
    if (selectedItems.Length > 0) {
        ; 打开选中的网站
        for _, item in selectedItems {
            OpenOneWebsite(item.url, browser)
            Sleep(200) ; 添加延迟以避免同时打开多个网站时的问题
        }
        ; ShowTooltip("已打开 " . selectedItems.Length . " 个网站")
    } else {
        ShowTooltip("未选择任何网站")
    }
}

; 处理双击单个网站
WebsiteSingleItem(listView, websiteList, browser, gui) {
    ; 获取当前选中的行
    row := listView.GetNext(0, "Focused")
    if (row) {
        site := websiteList[row]
        
        ; 销毁GUI
        gui.Destroy()
        
        ; 打开网站
        OpenOneWebsite(site.url, browser)
        ; ShowTooltip("已打开: " . site.name, 2000)
    }
}

;===============================================================
; 进程管理功能
;===============================================================

ManageProcessWithCtrlCheck(action) {
    ; 检测Ctrl键是否被按下
    if (GetKeyState("Ctrl", "P")) {
        ; 如果按下Ctrl键，显示进程选择GUI
        ShowProcessSelectionGUI(action)
    } else {
        ; 否则执行原始功能
        ManageProcess(action)
    }
}

ShowProcessSelectionGUI(action) {
    ; 关闭主菜单
    CloseMenu()
    
    ; 定义进程列表
    processList := []
    
    ; 根据动作类型确定要读取的INI Section
    sectionName := action = "启用" ? "GUIProcessesToStart" : "GUIProcessesToTerminate"
    
    ; 从INI文件读取进程列表
    i := 1
    Loop {
        nameKey := "Item" i "_Name"
        pathKey := "Item" i "_Path"
        checkedKey := "Item" i "_Checked"
        
        procName := IniRead("CapsLock++.ini", sectionName, nameKey, "")
        ; 如果读不到名称，则认为列表结束
        if (procName = "")
            break
            
        procPath := IniRead("CapsLock++.ini", sectionName, pathKey, "")
        procChecked := IniRead("CapsLock++.ini", sectionName, checkedKey, "true") = "true"
        
        ; 将读取到的进程信息添加到列表
        processList.Push({name: procName, path: procPath, checked: procChecked})
        
        i++
    }
    
    ; 如果列表为空，提示用户并返回
    if (processList.Length = 0) {
        ShowTooltip("错误: 无法从INI文件读取 " . sectionName . " 列表")
        return
    }

    ; 创建GUI
    processGui := Gui("+AlwaysOnTop +ToolWindow")
    processGui.Title := action = "启用" ? "选择要启用的进程" : "选择要终止的进程"
    
    ; 添加ListView控件
    listView := processGui.Add("ListView", "x10 y10 w400 h300 Checked", ["进程名称", "路径"])
    
    ; 填充ListView
    for _, proc in processList { ; 使用下划线忽略索引
        row := listView.Add(proc.checked ? "Check" : "", proc.name, proc.path)
    }
    
    ; 自动调整列宽
    listView.ModifyCol(1, 150)
    listView.ModifyCol(2, "Auto")
    
    ; 添加全选和反选按钮
    btnSelectAll := processGui.Add("Button", "x10 y320 w90 h30", "全选")
    btnSelectAll.OnEvent("Click", (*) => SelectAllItems(listView, true))
    
    btnSelectNone := processGui.Add("Button", "x110 y320 w90 h30", "全不选")
    btnSelectNone.OnEvent("Click", (*) => SelectAllItems(listView, false))
    
    btnInvert := processGui.Add("Button", "x210 y320 w90 h30", "反选")
    btnInvert.OnEvent("Click", (*) => InvertSelection(listView))
    
    ; 添加操作按钮 (不设置为Default，避免获取焦点)
    btnOk := processGui.Add("Button", "x310 y320 w100 h30", action)
    ; 传递 processList 给处理函数
    btnOk.OnEvent("Click", (*) => ProcessSelection(action, listView, processList, processGui)) 
    
    ; 添加返回和取消按钮
    btnBack := processGui.Add("Button", "x10 y360 w190 h30", "← 返回")
    btnBack.OnEvent("Click", BackToMainMenu)
    
    btnCancel := processGui.Add("Button", "x210 y360 w200 h30", "取消")
    btnCancel.OnEvent("Click", (*) => processGui.Destroy())
    
    ; 设置双击事件 - 直接操作单个进程
    ; 传递 processList 给处理函数
    listView.OnEvent("DoubleClick", (*) => ProcessSingleItem(action, listView, processList, processGui)) 
    
    ; 添加ESC键退出处理
    processGui.OnEvent("Escape", (*) => processGui.Destroy())
    ; 设置键盘钩子，以便在窗口有焦点时捕获ESC键
    processGui.OnEvent("Close", (*) => processGui.Destroy())
    
    ; 显示GUI并聚焦到ListView
    processGui.Show("w420 h400")
    listView.Focus()  ; 焦点设置在ListView上，而不是按钮上
}

; 返回主菜单
BackToMainMenu(*) {
    ; 关闭所有当前GUI窗口
    processGui := WinGetID("A")
    if (processGui)
        WinClose("ahk_id " processGui)
    
    ; 重新显示第1组菜单
    ShowMenu(1)
}

; 全选或全不选
SelectAllItems(listView, check) {
    totalItems := listView.GetCount()
    
    Loop totalItems {
        if (check) {
            listView.Modify(A_Index, "Check")
        } else {
            listView.Modify(A_Index, "-Check")
        }
    }
}

; 反选功能
InvertSelection(listView) {
    totalItems := listView.GetCount()
    
    Loop totalItems {
        ; 获取当前选中状态
        isChecked := listView.GetNext(A_Index - 1, "Checked") = A_Index
        
        ; 反转选中状态
        if (isChecked) {
            listView.Modify(A_Index, "-Check")
        } else {
            listView.Modify(A_Index, "Check")
        }
    }
}

; 处理进程选择
ProcessSelection(action, listView, processList, gui) {
    ; 获取所有选中的行
    selectedItems := []
    row := 0
    
    Loop {
        row := listView.GetNext(row, "Checked")
        if (!row)
            break
        selectedItems.Push(processList[row])
    }
    
    ; 销毁GUI
    gui.Destroy()
    
    ; 处理选中的进程
    if (selectedItems.Length > 0) {
        if (action = "启用") {
            ; 启用选中的进程
            for _, item in selectedItems {
                try {
                    Run(item.path)
                    Sleep(200) ; 添加延迟以避免同时启动多个进程时的问题
                } catch {
                    ; 忽略启动错误
                }
            }
            ShowTooltip("已启用 " . selectedItems.Length . " 个进程")
        } else {
            ; 终止选中的进程
            for _, item in selectedItems {
                try {
                    ProcessClose(item.path)
                } catch {
                    ; 忽略终止错误
                }
            }
            ShowTooltip("已终止 " . selectedItems.Length . " 个进程")
        }
    }
}

; 处理双击单个进程
ProcessSingleItem(action, listView, processList, gui) {
    ; 获取当前选中的行
    row := listView.GetNext(0, "Focused")
    if (row) {
        proc := processList[row]
        
        ; 销毁GUI
        gui.Destroy()
        
        ; 处理单个进程
        if (action = "启用") {
            try {
                Run(proc.path)
                ShowTooltip("已启用: " . proc.name)
            } catch {
                ShowTooltip("无法启用: " . proc.name)
            }
        } else {
            try {
                ProcessClose(proc.path)
                ShowTooltip("已终止: " . proc.name)
            } catch {
                ShowTooltip("无法终止: " . proc.name)
            }
        }
    }
}

;===============================================================
; 悄悄话功能
;===============================================================

; 窗口透明器 - 将窗口设置为完全透明与鼠标穿透，并在主屏幕上显示窗口内容

class WindowTransparency {
    ; 存储变量
    static targetWindowHwnd := 0       ; 目标窗口句柄
    static originalStyle := 0          ; 原始窗口样式
    static isTransparent := false      ; 是否已设置为透明
    static previewGui := 0             ; 预览窗口
    static isPreviewActive := false    ; 预览窗口是否活动
    static captureArea := {x: 0, y: 0, w: 400, h: 300}  ; 截取区域
    
    ; 窗口样式常量
    static WS_EX_LAYERED := 0x80000
    static WS_EX_TRANSPARENT := 0x20
    static LWA_ALPHA := 0x2
    static GWL_EXSTYLE := -20

    ; 切换窗口透明度
    static ToggleTransparency() {
        if (WindowTransparency.isTransparent) {
            ; 已经设置为透明，恢复原始样式
            WindowTransparency.RestoreWindow()
        } else {
            ; 获取鼠标下的窗口并设置为透明
            WindowTransparency.MakeTransparent()
        }
    }

    ; 将鼠标下的窗口设置为透明
    static MakeTransparent() {
        ; 获取鼠标下的窗口
        MouseGetPos(, , &mouseWin)
        if (!mouseWin) {
            ToolTip("无法获取鼠标下的窗口")
            SetTimer () => ToolTip(), -2000
            return
        }

        ; 检查是否是微信窗口
        processName := WinGetProcessName("ahk_id " mouseWin)
        if (processName != "WeChat.exe") {
            ToolTip("该功能仅支持微信窗口")
            SetTimer () => ToolTip(), -2000
            return
        }

        ; 保存窗口信息
        WindowTransparency.targetWindowHwnd := mouseWin
        
        ; 获取窗口当前样式
        style := WinGetStyle("ahk_id " mouseWin)
        WindowTransparency.originalStyle := style
        
        ; 设置窗口为透明和鼠标穿透
        ; 1. 添加分层窗口样式
        DllCall("SetWindowLong", "Ptr", mouseWin, "Int", WindowTransparency.GWL_EXSTYLE, 
                "Ptr", style | WindowTransparency.WS_EX_LAYERED | WindowTransparency.WS_EX_TRANSPARENT)
        
        ; 2. 设置窗口透明度为0（完全透明）
        DllCall("SetLayeredWindowAttributes", "Ptr", mouseWin, "UInt", 0, "UChar", 0, "UInt", WindowTransparency.LWA_ALPHA)
        
        ; 更新状态
        WindowTransparency.isTransparent := true
        
        ; 创建预览窗口并开始截图
        WindowTransparency.StartPreview()
        
        ; 获取主窗口位置和大小
        WinGetPos(&mainX, &mainY, &mainW, &mainH, "ahk_id " mouseWin)
        
        ; 查找并关闭阴影窗口
        if (shadowHwnd := WinExist("ahk_class popupshadow")) {
            ; 获取阴影窗口位置和大小
            WinGetPos(&shadowX, &shadowY, &shadowW, &shadowH, "ahk_id " shadowHwnd)
            
            ; 检查位置关系
            if ((shadowX == mainX - 20) 
                && (shadowY == mainY - 20)
                && (shadowW == mainW + 40)
                && (shadowH == mainH + 40)) {
                ; 关闭阴影窗口
                WinClose("ahk_id " shadowHwnd)
                ToolTip("已关闭阴影窗口")
                SetTimer () => ToolTip(), -2000
            }
        }
        
        ToolTip("窗口已设置为透明，屏幕预览已启动")
        SetTimer () => ToolTip(), -2000
    }
    
    ; 恢复窗口原始样式
    static RestoreWindow() {
        ; 停止预览
        WindowTransparency.StopPreview()
        
        if (!WindowTransparency.targetWindowHwnd || !WinExist("ahk_id " WindowTransparency.targetWindowHwnd)) {
            ToolTip("无法恢复窗口样式：目标窗口不存在")
            SetTimer () => ToolTip(), -2000
            return
        }
        
        ; 获取保存的样式
        style := WindowTransparency.originalStyle
        if (style == "") {
            ToolTip("无法恢复窗口样式：未保存原始样式")
            SetTimer () => ToolTip(), -2000
            return
        }
        
        ; 恢复窗口样式
        ; 1. 移除分层窗口样式
        DllCall("SetWindowLong", "Ptr", WindowTransparency.targetWindowHwnd, "Int", WindowTransparency.GWL_EXSTYLE, "Ptr", style)
        
        ; 2. 移除透明度
        DllCall("SetLayeredWindowAttributes", "Ptr", WindowTransparency.targetWindowHwnd, "UInt", 0, "UChar", 255, "UInt", WindowTransparency.LWA_ALPHA)
        
        ; 更新状态
        WindowTransparency.isTransparent := false
        WindowTransparency.targetWindowHwnd := 0
        WindowTransparency.originalStyle := 0
        
        ToolTip("窗口样式已恢复")
        SetTimer () => ToolTip(), -2000
    }
    
    ; 创建预览窗口并开始截图
    static StartPreview() {
        ; 创建预览窗口 - 使用最简单的方式
        WindowTransparency.previewGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
        WindowTransparency.previewGui.MarginX := 0
        WindowTransparency.previewGui.MarginY := 0
        
        ; 任务栏高度固定
        taskbarHeight := 93
        previewH :=taskbarHeight
        guiY := A_ScreenHeight - previewH
        
        ; 根据原始截图区域的宽高比例计算预览宽度
        aspectRatio := WindowTransparency.captureArea.w / WindowTransparency.captureArea.h
        previewW := Round(previewH * aspectRatio)
        
        ; 添加图片控件，绑定点击事件
        previewImage := WindowTransparency.previewGui.Add("Picture", "w" previewW " h" previewH " vPreviewImage")
        previewImage.OnEvent("Click", WindowTransparency.OnPreviewClick)
        
        ; 计算右对齐位置（与屏幕右边界内切）
        windowX := A_ScreenWidth - 2 * previewW
        
        ; 显示预览窗口
        WindowTransparency.previewGui.Show("x" windowX " y" guiY " w" previewW " h" previewH " NoActivate")
        
        ; 开始定时更新
        SetTimer UpdatePreviewFunction, 50

        ; 设置预览窗口状态为活动
        WindowTransparency.isPreviewActive := true
    }
    
    ; 停止预览
    static StopPreview() {
        ; 设置预览窗口状态为非活动
        WindowTransparency.isPreviewActive := false

        ; 停止定时器
        SetTimer UpdatePreviewFunction, 0
        
        ; 关闭预览窗口
        if (WindowTransparency.previewGui) {
            WindowTransparency.previewGui.Destroy()
            WindowTransparency.previewGui := 0
        }
        
        ; 清除静态变量中存储的位图句柄
        if (WindowTransparency.UpdatePreview.HasOwnProp("previousBitmap") && WindowTransparency.UpdatePreview.previousBitmap) {
            DllCall("DeleteObject", "Ptr", WindowTransparency.UpdatePreview.previousBitmap)
            WindowTransparency.UpdatePreview.previousBitmap := 0
        }
    }
    
    ; 更新预览图像
    static UpdatePreview() {
        static previousBitmap := 0  ; 保存前一个位图的句柄

        try {
            ; 检查是否还存在目标窗口
            if (!WindowTransparency.targetWindowHwnd || !WinExist("ahk_id " WindowTransparency.targetWindowHwnd)) {
                WindowTransparency.StopPreview()
                return
            }
            
            ; 获取目标窗口句柄和截取区域
            hwnd := WindowTransparency.targetWindowHwnd
            area := WindowTransparency.captureArea
            
            ; 创建屏幕DC
            screenDC := DllCall("GetDC", "Ptr", 0, "Ptr")
            
            ; 创建兼容DC
            memDC := DllCall("CreateCompatibleDC", "Ptr", screenDC, "Ptr")
            
            ; 获取窗口DC
            windowDC := DllCall("GetDC", "Ptr", hwnd, "Ptr")
            
            ; 创建位图
            hBitmap := DllCall("CreateCompatibleBitmap", "Ptr", screenDC, "Int", area.w, "Int", area.h, "Ptr")
            
            ; 选择位图到内存DC
            oldBitmap := DllCall("SelectObject", "Ptr", memDC, "Ptr", hBitmap, "Ptr")
            
            ; 从窗口DC复制到内存DC (BitBlt可以从屏幕外区域获取内容)
            DllCall("BitBlt", "Ptr", memDC, "Int", 0, "Int", 0, "Int", area.w, "Int", area.h, 
                    "Ptr", windowDC, "Int", area.x, "Int", area.y, "UInt", 0x00CC0020) ; SRCCOPY
            
            ; 清理
            DllCall("SelectObject", "Ptr", memDC, "Ptr", oldBitmap)
            DllCall("ReleaseDC", "Ptr", hwnd, "Ptr", windowDC)
            DllCall("DeleteDC", "Ptr", memDC)
            DllCall("ReleaseDC", "Ptr", 0, "Ptr", screenDC)
            
            ; 将位图设置到预览窗口
            if (WindowTransparency.previewGui) {
                previewImage := WindowTransparency.previewGui["PreviewImage"]
                if (previewImage) {
                    ; 删除之前的位图，防止内存泄漏
                    if (previousBitmap) {
                        DllCall("DeleteObject", "Ptr", previousBitmap)
                    }
                    
                    ; 设置新位图
                    previewImage.Value := "HBITMAP:*" hBitmap
                    
                    ; 保存当前位图句柄以便下次释放
                    previousBitmap := hBitmap
                } else {
                    ; 如果没有图片控件，直接删除位图
                    DllCall("DeleteObject", "Ptr", hBitmap)
                }
            } else {
                ; 如果没有预览窗口，直接删除位图
                DllCall("DeleteObject", "Ptr", hBitmap)
            }
        } catch as e {
            OutputDebug("预览更新错误: " e.Message)
            
            ; 即使发生错误，也要尝试删除位图
            if (IsSet(hBitmap) && hBitmap) {
                DllCall("DeleteObject", "Ptr", hBitmap)
            }
        }
    }
    
    ; 激活目标窗口
    static ActivateTarget() {
        if (WindowTransparency.targetWindowHwnd) {
            WinActivate("ahk_id " WindowTransparency.targetWindowHwnd)
        }
    }
    
    ; 调整截取区域
    static AdjustCaptureArea(deltaY := 0) {
        area := WindowTransparency.captureArea
        
        ; 获取目标窗口的大小
        if (WindowTransparency.targetWindowHwnd) {
            WinGetPos(, , &targetW, &targetH, "ahk_id " WindowTransparency.targetWindowHwnd)
        } else {
            targetW := A_ScreenWidth
            targetH := A_ScreenHeight
        }
        
        ; 只调整Y位置
        area.y += deltaY
        
        ; 确保数值合理且不超出窗口边界
        area.x := 0  ; X始终为0
        area.y := Max(0, Min(area.y, targetH - area.h // 2))  ; 确保截取区域不超出窗口底部
        area.w := targetW  ; 宽度始终等于窗口宽度
        
        WindowTransparency.captureArea := area
        
        ; 更新预览窗口大小
        if (WindowTransparency.previewGui) {
            ; 任务栏高度固定
            taskbarHeight := 93
            previewH := taskbarHeight
            
            ; 保持宽高比
            aspectRatio := area.w / area.h
            previewW := Round(previewH * aspectRatio)
            
            ; 获取当前位置，保持Y坐标不变
            WinGetPos(, &guiY, , , WindowTransparency.previewGui.Hwnd)
            
            ; 计算右对齐位置（与屏幕右边界内切）
            windowX := A_ScreenWidth - 2 * previewW
            
            ; 更新窗口大小和位置
            WindowTransparency.previewGui.Show("x" windowX " y" guiY " w" previewW " h" previewH " NoActivate")
            
            ; 更新图片控件大小
            previewImage := WindowTransparency.previewGui["PreviewImage"]
            if (previewImage)
                previewImage.Move(0, 0, previewW, previewH)
        }
        
        ToolTip("截取区域: X=" area.x ", Y=" area.y ", 宽=" area.w ", 高=" area.h)
        SetTimer () => ToolTip(), -1000
    }

    ; 处理预览窗口点击事件
    static OnPreviewClick(thisCtrl, *) {
        ; 确保目标窗口存在
        if (!WindowTransparency.targetWindowHwnd || !WinExist("ahk_id " WindowTransparency.targetWindowHwnd)) {
            return
        }
        
        ; 激活窗口
        WinActivate("ahk_id " WindowTransparency.targetWindowHwnd)
    }
}

; 定时器回调函数 - 更新预览
UpdatePreviewFunction() {
    WindowTransparency.UpdatePreview()
}

#HotIf GetKeyState("CapsLock", "P")
; 定义热键 - 按下 Capslock+p 切换窗口位置和预览
`:: 
{
    ; 标记为按下了其他键
    global otherKeyPressed := true
    WindowTransparency.ToggleTransparency()
}
#HotIf

#HotIf WindowTransparency.isPreviewActive && WinActive("ahk_id " WindowTransparency.targetWindowHwnd)
+WheelUp:: WindowTransparency.AdjustCaptureArea(-20)  ; 上移
+WheelDown:: WindowTransparency.AdjustCaptureArea(20)   ; 下移
#HotIf

;==================================================
; 窗口裁切工具 - 使窗口除了特定区域外都变为透明，同时禁用窗口阴影
; 使用 CapsLock+Y 裁切光标下的窗口，通过鼠标拖拽选择保留区域，并自动禁用窗口阴影
; 使用 CapsLock+Y+Space 恢复窗口原始状态
;==================================================
; 常量定义
WS_EX_LAYERED := 0x80000
WS_EX_TRANSPARENT := 0x20
WS_EX_TOOLWINDOW := 0x80
WS_EX_DLGMODALFRAME := 0x1
WS_EX_TOPMOST := 0x8       ; 置顶窗口样式
LWA_ALPHA := 0x2
LWA_COLORKEY := 0x1

; 保存原始窗口状态
global windowStates := Map()
global closedShadows := Map()  ; 用于存储关闭的阴影窗口信息

; 鼠标选择相关变量
global isSelecting := false
global targetWindow := 0
global selectionBox := 0
global startX := 0, startY := 0
global endX := 0, endY := 0
global winOffsetX := 0, winOffsetY := 0
global trackingTimer := 0  ; 添加全局变量声明
global lastTrackX := 0, lastTrackY := 0  ; 添加全局变量声明

global isDragging := false
global dragHwnd := 0
global startX := 0, startY := 0
global winX := 0, winY := 0
global offsetX := 0, offsetY := 0  ; 鼠标与窗口左上角的偏移量

; 热键：裁切窗口
^+x::StartClipWindow()

; 热键：恢复窗口 - 如果正在选择区域，则取消选择
^+z::
{
    if (isSelecting) {
        CancelSelection()
        ShowTooltip("已取消选择")
    } else {
        RestoreWindow()
    }
}

#HotIf IsActiveWindowClipped()
; 控制裁切区域移动的热键
^WheelDown:: MoveClipRegion("down")
^WheelUp:: MoveClipRegion("up")
^+WheelDown:: MoveClipRegion("right")
^+WheelUp:: MoveClipRegion("left")
; 添加新的热键 - 拖动裁切过的窗口
^+LButton:: StartDragClippedWindow()
^+LButton Up:: StopDragClippedWindow()
; 添加新的热键 - 调整裁剪区域大小
^!WheelUp:: ResizeClipRegion("heightIncrease")    ; Ctrl+Alt+滚轮上 - 增加高度
^!WheelDown:: ResizeClipRegion("heightDecrease")  ; Ctrl+Alt+滚轮下 - 减少高度
^+!WheelUp:: ResizeClipRegion("widthIncrease")    ; Ctrl+Shift+Alt+滚轮上 - 增加宽度
^+!WheelDown:: ResizeClipRegion("widthDecrease")  ; Ctrl+Shift+Alt+滚轮下 - 减少宽度
#HotIf

; 开始裁切窗口过程
StartClipWindow() {
    global isSelecting, targetWindow, winOffsetX, winOffsetY
    
    ; 清除所有可能存在的ToolTip，避免干扰
    ClearAllToolTips()
    
    ; 获取鼠标下的窗口句柄
    MouseGetPos(&mouseX, &mouseY, &hWnd)
    
    if (!hWnd) {
        ;ShowToolTip("未检测到窗口")
        return
    }
    
    ; 获取窗口信息进行黑名单检查
    winTitle := WinGetTitle("ahk_id " hWnd)
    className := WinGetClass("ahk_id " hWnd)
    processName := WinGetProcessName("ahk_id " hWnd)
    processPath := WinGetProcessPath("ahk_id " hWnd)
    
    ; 硬编码排除特定窗口
    ; 排除桌面/Program Manager
    if (className == "Progman" && winTitle == "Program Manager" && processName == "explorer.exe") {
        ShowToolTip("无法裁剪桌面窗口")
        return
    }
    
    ; 排除任务栏
    if (className == "Shell_TrayWnd" && processName == "explorer.exe") {
        ShowToolTip("无法裁剪任务栏")
        return
    }
    
    ; 排除文件资源管理器
    if (className == "CabinetWClass" && processName == "explorer.exe") {
        ShowToolTip("无法裁剪资源管理器")
        return
    }
    
    ; 排除QQ悬浮栏
    if ((processName == "QQ.exe" || processName == "QQScLauncher.exe" || processName == "QQProtect.exe") && InStr(winTitle, "QQ")) {
        ShowToolTip("无法裁剪QQ悬浮栏")
        return
    }
    
    ; 排除VueMinder日历窗口
    if (InStr(className, "WindowsForms10.Window.8.app.0.1a0e24_r10_ad1") && 
        (processName == "VueMinder.exe" || InStr(processPath, "VueMinder.exe"))) {
        ShowToolTip("无法裁剪VueMinder窗口")
        return
    }
    
    ; 排除QuinkNote的SwitchPlug插件
    if (InStr(className, "HwndWrapper[SwitchPlug.exe") || 
        InStr(processPath, "QuinkNote\plugins\switchPlug\bin\SwitchPlug.exe")) {
        ShowToolTip("无法裁剪SwitchPlug窗口")
        return
    }

    ; 排除fences的folder portal
    if (className == "ExplorerBrowserOwner" && processName == "explorer.exe") {
        ShowToolTip("无法裁剪folder portal")
        return
    }

    if (className == "TaskManagerWindow" && processName == "Taskmgr.exe"){
        ShowToolTip("无法裁剪任务管理器")
        return
    }
    
    ; 检查是否已经在选择中
    if (isSelecting) {
        ;ShowToolTip("已经在选择区域中，请完成当前操作")
        return
    }
    
    ; 检查窗口是否已被处理，如果是则先恢复
    if (windowStates.Has(hWnd)) {
        RestoreWindow(hWnd)
    }

    ; 排除OneCommander
    if (InStr(className, "HwndWrapper[OneCommander.exe") || processName == "OneCommander.exe") {
        ShowToolTip("无法裁剪OneCommander窗口")
        return
    }
    
    ; 获取窗口位置和大小
    WinGetPos(&winX, &winY, &winW, &winH, "ahk_id " hWnd)
    
    ; 保存窗口原始状态 - 包括更多属性以支持阴影禁用
    SaveWindowState(hWnd)
    
    ; 确保窗口有分层样式
    if (!(WinGetExStyle(hWnd) & WS_EX_LAYERED)) {
        WinSetExStyle("+0x80000", "ahk_id " hWnd)
    }
    
    ; 记录目标窗口和偏移量
    targetWindow := hWnd
    winOffsetX := winX
    winOffsetY := winY
    
    ; 设置状态为选择中
    isSelecting := true
    
    ; 设置等待裁剪的光标样式
    SetWaitCursor()
    
    ; 创建临时热键来监听鼠标事件 (只在选择模式下有效)
    Hotkey("*LButton", HandleLeftButtonDown, "On")
    Hotkey("Escape", CancelSelection, "On")
}

; 创建选择框
CreateSelectionBox() {
    global selectionBox
    
    ; 如果已存在选择框，先销毁
    if (selectionBox && WinExist("ahk_id " . selectionBox)) {
        Gui(selectionBox . ":Destroy")
    }
    
    ; 直接创建一个新的窗口而不保存对象
    selectionBox := Gui("+AlwaysOnTop -Caption +ToolWindow +LastFound +E0x20")
    
    ; 设置窗口颜色和透明度 - 使用黑色以保持一致性
    selectionBox.BackColor := "000000"
    WinSetTransparent(100, selectionBox)
    
    ; 先显示在屏幕左上角以确保窗口创建成功 - 使用NA参数避免激活窗口
    selectionBox.Show("x0 y0 w5 h5 NA")
    
    ; 设置置顶和透明点击穿透属性
    WinSetExStyle("+0x8 +0x20", "ahk_id " . selectionBox.Hwnd) ; WS_EX_TOPMOST | WS_EX_TRANSPARENT
    
    return selectionBox.Hwnd
}

; 更新选择框位置
UpdateSelectionBox(x, y, w, h) {
    global selectionBox
    
    ; 立即捕获有效性判断结果，避免在判断和使用之间状态变化
    validSelection := selectionBox && WinExist("ahk_id " . selectionBox)
    if (!validSelection) {
        return false
    }
    
    ; 确保最小尺寸
    w := Max(w, 5)
    h := Max(h, 5)
    
    ; 使用更安全的错误处理方式更新窗口
    success := false
    try {
        ; 直接使用修改后的静态调用，避免引用问题
        hwnd := selectionBox + 0  ; 确保是数字
        if (hwnd) {
            WinMove(x, y, w, h, "ahk_id " . hwnd)
            success := true
        }
    }
    
    return success
}

; 初始化鼠标追踪系统
InitMouseTracking() {
    global trackingTimer
    
    ; 停止任何现有的追踪计时器
    if (trackingTimer) {
        SetTimer(trackingTimer, 0)
    }
    
    ; 设置高频率计时器来追踪鼠标 - 修复This引用
    trackingTimer := TrackMouseMovement
    SetTimer(trackingTimer, 16)  ; 约60FPS的刷新率
    
    ; 也同时设置消息钩子作为备份机制
    OnMessage(0x0200, MouseMove)  ; WM_MOUSEMOVE
}

; 停止鼠标追踪
StopMouseTracking() {
    global trackingTimer
    
    ; 停止追踪计时器
    if (trackingTimer) {
        SetTimer(trackingTimer, 0)
        trackingTimer := 0
    }
    
    ; 移除消息钩子
    OnMessage(0x0200, MouseMove, 0)
}

; 计时器调用的鼠标追踪函数
TrackMouseMovement(*) {
    global isSelecting, targetWindow, selectionBox, startX, startY, endX, endY, lastTrackX, lastTrackY
    
    if (!isSelecting)
        return
    
    ; 获取鼠标位置
    CoordMode("Mouse", "Screen")
    MouseGetPos(&screenX, &screenY)
    
    ; 如果位置与上次相同，不需要更新
    if (screenX = lastTrackX && screenY = lastTrackY)
        return
    
    ; 更新最后追踪位置
    lastTrackX := screenX
    lastTrackY := screenY
    
    ; 获取目标窗口位置
    WinGetPos(&winX, &winY, &winW, &winH, "ahk_id " . targetWindow)
    
    ; 更新结束坐标（相对于窗口）
    endX := screenX - winX
    endY := screenY - winY
    
    ; 计算选择框位置（确保正确的上左到右下顺序）
    left := Min(startX, endX)
    top := Min(startY, endY)
    width := Abs(endX - startX)
    height := Abs(endY - startY)
    
    ; 转换回屏幕坐标
    screenLeft := winX + left
    screenTop := winY + top
    
    ; 更新选择框
    if (!UpdateSelectionBox(screenLeft, screenTop, width, height)) {
        ; 如果更新失败，尝试重新创建选择框
        selectionBox := CreateSelectionBox()
        UpdateSelectionBox(screenLeft, screenTop, width, height)
    }
}

; 处理鼠标左键按下
HandleLeftButtonDown(*) {
    global isSelecting, targetWindow, startX, startY, selectionBox, lastTrackX, lastTrackY
    
    if (!isSelecting)
        return
    
    ; 关闭左键监听，以免冲突
    Hotkey("*LButton", "Off")
    
    ; 获取鼠标相对于屏幕的坐标
    CoordMode("Mouse", "Screen")
    MouseGetPos(&screenX, &screenY)
    
    ; 记录起始位置用于追踪比较
    lastTrackX := screenX
    lastTrackY := screenY
    
    ; 获取目标窗口的位置
    WinGetPos(&winX, &winY, &winW, &winH, "ahk_id " . targetWindow)
    
    ; 计算鼠标相对于窗口的位置
    startX := screenX - winX
    startY := screenY - winY
    
    ; 设置全局十字光标
    SetCursorCross()
    
    ; 创建选择框
    selectionBox := CreateSelectionBox()
    
    ; 立即更新选择框到起始位置
    UpdateSelectionBox(screenX, screenY, 5, 5)
    
    ; 初始化持续追踪系统
    InitMouseTracking()
    
    ; 设置左键释放监听
    Hotkey("*LButton up", HandleLeftButtonUp, "On")
}

; 处理鼠标左键释放
HandleLeftButtonUp(*) {
    global isSelecting, targetWindow, startX, startY, endX, endY
    
    if (!isSelecting)
        return
    
    ; 停止鼠标追踪
    StopMouseTracking()
    
    ; 获取最终鼠标位置
    CoordMode("Mouse", "Screen")
    MouseGetPos(&screenX, &screenY)
    
    ; 转换为相对于窗口的坐标
    WinGetPos(&winX, &winY, , , "ahk_id " targetWindow)
    endX := screenX - winX
    endY := screenY - winY
    
    ; 确保坐标正确（左上到右下）
    NormalizeCoordinates()
    
    ; 计算宽高
    width := endX - startX
    height := endY - startY
    
    ; 检查选择区域是否过小
    if (width < 10 || height < 10) {
        CancelSelection()
        ShowToolTip("选择区域太小，已取消")
        return
    }
    
    ; 验证目标窗口状态在Map中存在
    if (!windowStates.Has(targetWindow)) {
        ; 窗口状态不存在，可能是因为在选择过程中窗口被恢复了
        ; 重新保存窗口状态
        SaveWindowState(targetWindow)
    }
    
    ; 先处理分层样式和阴影（除了微信）
    winExe := windowStates[targetWindow].winExe
    if (winExe != "WeChat.exe") {
        ; 确保窗口有分层样式
        if (!(WinGetExStyle(targetWindow) & WS_EX_LAYERED)) {
            WinSetExStyle("+0x80000", "ahk_id " targetWindow)
        }
        
        ; 禁用阴影 - 在应用裁剪前
        DisableShadowForWindow(targetWindow)
    }
    
    ; 然后应用裁切
    ApplyClipToWindow(targetWindow, startX, startY, width, height)
    
    ; 对于微信，在裁剪后处理阴影
    if (winExe = "WeChat.exe") {
        DisableShadowForWindow(targetWindow)
    }
    
    ; 清理选择相关资源
    CleanupSelection()
    
    ; 恢复正常鼠标指针
    RestoreDefaultCursor()
}

; 设置鼠标十字形样式 - 使用系统光标API实现持久替换
SetCursorCross() {
    ; 加载系统十字光标
    hCross := DllCall("LoadCursor", "Ptr", 0, "Ptr", 32515, "Ptr")  ; IDC_CROSS = 32515
    
    ; 替换所有系统光标为十字光标
    ; 注意：这会替换系统的所有光标类型，但在我们的场景中这是可接受的
    ; 0 = OCR_NORMAL (标准箭头)
    DllCall("SetSystemCursor", "Ptr", DllCall("CopyIcon", "Ptr", hCross, "Ptr"), "UInt", 32512)  ; OCR_NORMAL
    ; 1 = OCR_IBEAM (I-形文本光标)
    DllCall("SetSystemCursor", "Ptr", DllCall("CopyIcon", "Ptr", hCross, "Ptr"), "UInt", 32513)  ; OCR_IBEAM
}

; 恢复默认鼠标样式
RestoreDefaultCursor() {
    ; 恢复所有系统光标为默认值
    DllCall("SystemParametersInfo", "UInt", 0x0057, "UInt", 0, "Ptr", 0, "UInt", 0)  ; SPI_SETCURSORS = 0x0057
}

; 确保坐标顺序正确（左上角到右下角）
NormalizeCoordinates() {
    global startX, startY, endX, endY
    
    ; 如果结束点在起始点的左边，交换X坐标
    if (endX < startX) {
        temp := startX
        startX := endX
        endX := temp
    }
    
    ; 如果结束点在起始点的上边，交换Y坐标
    if (endY < startY) {
        temp := startY
        startY := endY
        endY := temp
    }
}

; 保存窗口状态 - 增强版，支持更多属性
SaveWindowState(hWnd) {
    ; 获取当前窗口样式
    exStyle := WinGetExStyle(hWnd)
    
    ; 获取窗口进程名
    winExe := WinGetProcessName(hWnd)
    
    ; 检查是否已经是分层窗口
    isLayered := (exStyle & WS_EX_LAYERED) != 0
    
    ; 如果是分层窗口，获取当前透明度
    alpha := 255
    if (isLayered) {
        try {
            transparency := WinGetTransparent("ahk_id " hWnd)
            if (transparency != "") {
                alpha := transparency
            }
        }
    }
    
    ; 检查窗口是否为最大化或全屏
    isMaximized := WinGetMinMax("ahk_id " hWnd) = 1
    
    ; 检查窗口是否可能是全屏(通过检测窗口位置和屏幕尺寸)
    MonitorGetWorkArea(, &monLeft, &monTop, &monRight, &monBottom)
    WinGetPos(&winX, &winY, &winW, &winH, "ahk_id " hWnd)
    isFullscreen := (winX <= monLeft) && (winY <= monTop) && 
                   (winX + winW >= monRight) && (winY + winH >= monBottom)
    
    ; 保存状态 - 包括更多信息以支持阴影禁用
    windowStates[hWnd] := {
        exStyle: exStyle, 
        isLayered: isLayered, 
        alpha: alpha,
        region: "",          ; 原始区域（通常为空）
        winExe: winExe,      ; 进程名，用于特殊处理不同应用
        shadowDisabled: false, ; 标记阴影是否已被禁用
        isMaximized: isMaximized, ; 窗口是否最大化
        isFullscreen: isFullscreen ; 窗口是否可能是全屏
    }
}

; 应用裁切到窗口
ApplyClipToWindow(hWnd, x, y, width, height) {
    try {
        ; 创建矩形区域
        hRgn := DllCall("CreateRectRgn", "Int", x, "Int", y, 
                        "Int", x + width, "Int", y + height, "Ptr")
        
        ; 应用区域到窗口
        if (hRgn) {
            ; 获取窗口当前属性
            winExe := windowStates[hWnd].winExe
            
            ; 应用区域
            DllCall("SetWindowRgn", "Ptr", hWnd, "Ptr", hRgn, "Int", true)
            
            ; 保存使用的区域信息以便恢复
            windowStates[hWnd].region := {x: x, y: y, w: width, h: height}
            
            ; 对于非微信窗口，设置透明度为254
            if (winExe != "WeChat.exe" && windowStates[hWnd].isLayered) {
                DllCall("SetLayeredWindowAttributes", "Ptr", hWnd, "UInt", 0, "UChar", 254, "UInt", LWA_ALPHA)
            }
        } else {
            ShowToolTip("创建区域失败")
        }
    } catch as e {
        ShowToolTip("裁切窗口失败: " e.Message)
    }
}

; 禁用窗口阴影
DisableShadowForWindow(hWnd) {
    ; 如果窗口状态不存在，返回
    if (!windowStates.Has(hWnd)) {
        return
    }
    
    ; 确保winExe属性存在
    if (!windowStates[hWnd].HasOwnProp("winExe")) {
        ; 如果没有winExe属性，获取并设置它
        windowStates[hWnd].winExe := WinGetProcessName(hWnd)
    }
    
    ; 获取窗口进程名
    winExe := windowStates[hWnd].winExe
    winTitle := WinGetTitle(hWnd)
    
    ; 对于微信使用特殊处理
    if (winExe = "WeChat.exe") {
        DisableWeChatShadow(hWnd)
    } else {
        DisableNormalShadow(hWnd)
    }
    
    ; 标记阴影已被禁用
    windowStates[hWnd].shadowDisabled := true
}

; 针对微信的阴影禁用方法
DisableWeChatShadow(hWnd) {
    ; 查找微信阴影窗口
    shadowHwnd := WinExist("ahk_class popupshadow")
    if (!shadowHwnd) {
        return  ; 如果没找到阴影窗口就不处理
    }
    
    ; 获取主窗口位置和大小
    WinGetPos(&mainX, &mainY, &mainW, &mainH, "ahk_id " hWnd)
    
    ; 获取阴影窗口位置和大小
    WinGetPos(&shadowX, &shadowY, &shadowW, &shadowH, "ahk_id " shadowHwnd)
    
    ; 检查位置关系
    isMatch := (shadowX <= mainX - 10) 
            && (shadowY <= mainY - 10)
            && (shadowW >= mainW + 20)
            && (shadowH >= mainH + 20)
            
    if (isMatch) {
        ; 保存阴影窗口信息，用于恢复
        closedShadows[hWnd] := {
            shadowHwnd: shadowHwnd, 
            shadowX: shadowX, 
            shadowY: shadowY, 
            shadowW: shadowW, 
            shadowH: shadowH
        }
        
        ; 关闭阴影窗口
        WinClose("ahk_id " shadowHwnd)
    }
}

; 针对普通窗口的阴影禁用方法
DisableNormalShadow(hWnd) {
    ; 尝试多种DWM属性来禁用阴影
    ; DWMWA_NCRENDERING_POLICY = 2
    value := Buffer(4, 0)
    NumPut("Int", 1, value)
    DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hWnd, "Int", 2, "Ptr", value, "Int", 4)
    
    ; DWMWA_EXCLUDED_FROM_PEEK = 12
    NumPut("Int", 1, value)
    DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hWnd, "Int", 12, "Ptr", value, "Int", 4)
    
    ; 检查特定应用的特殊处理
    winExe := windowStates[hWnd].winExe
    
    ; 对于特定应用添加额外样式
    if (winExe = "Typora.exe") {
        ; 对于Typora，添加额外的样式
        newStyle := WinGetExStyle(hWnd) | WS_EX_DLGMODALFRAME
        DllCall("SetWindowLong", "Ptr", hWnd, "Int", -20, "Int", newStyle)
    }
    
    ; 设置透明度为254（几乎完全不透明）- 这有助于去除阴影
    if (windowStates[hWnd].isLayered) {
        DllCall("SetLayeredWindowAttributes", "Ptr", hWnd, "UInt", 0, "UChar", 254, "UInt", LWA_ALPHA)
    }
    
    ; 强制重绘窗口
    DllCall("SetWindowPos", "Ptr", hWnd, "Ptr", 0, "Int", 0, "Int", 0, 
           "Int", 0, "Int", 0, "UInt", 0x0027) ; SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_FRAMECHANGED
}

; 取消选择
CancelSelection(*) {
    global isSelecting
    
    if (!isSelecting)
        return
    
    CleanupSelection()
    ShowToolTip("已取消选择")
}

; 清理选择相关资源
CleanupSelection() {
    global isSelecting, selectionBox, trackingTimer
    
    ; 停止鼠标追踪
    StopMouseTracking()
    
    ; 关闭临时热键
    Hotkey("*LButton", "Off")
    Hotkey("*LButton up", HandleLeftButtonUp, "Off")  ; 修改为包含回调函数名称
    Hotkey("Escape", "Off")
    
    ; 销毁选择框 - 修复销毁方法
    if (selectionBox && WinExist("ahk_id " . selectionBox)) {
        try {
            WinClose("ahk_id " . selectionBox)
        } catch {
            ; 忽略错误
        }
        selectionBox := 0
    }
    
    ; 清除工具提示
    ToolTip("")
    
    ; 恢复默认鼠标
    RestoreDefaultCursor()
    
    ; 重置状态
    isSelecting := false
}

; 恢复窗口函数 - 增强版，同时恢复阴影
RestoreWindow(hWnd := 0) {
    global windowStates, closedShadows
    
    ; 如果未提供窗口句柄，获取鼠标下的窗口
    if (!hWnd) {
        MouseGetPos(, , &hWnd)
        
        if (!hWnd) {
            ShowToolTip("未检测到窗口")
            return
        }
    }
    
    winTitle := WinGetTitle(hWnd)
    
    ; 检查是否有保存的窗口状态
    if (!windowStates.Has(hWnd)) {
        ShowToolTip("未找到窗口原始状态: " winTitle)
        return
    }
    
    ; 获取保存的状态
    state := windowStates[hWnd]
    
    ; 清除窗口区域设置（恢复完整窗口）
    DllCall("SetWindowRgn", "Ptr", hWnd, "Ptr", 0, "Int", true)
    
    ; 如果阴影被禁用了，恢复阴影
    if (state.shadowDisabled) {
        RestoreShadowForWindow(hWnd)
    }
    
    ; 如果原始窗口不是分层窗口，则移除分层样式
    if (!state.isLayered) {
        WinSetExStyle("-0x80000", "ahk_id " hWnd)
    } else if (state.alpha != 255) {
        ; 恢复原来的透明度
        WinSetTransparent(state.alpha, "ahk_id " hWnd)
    }
    
    ; 从Map中移除
    windowStates.Delete(hWnd)
    
    ;ShowToolTip("已恢复窗口: " winTitle)
}

; 恢复窗口阴影
RestoreShadowForWindow(hWnd) {
    ; 检查是否是微信窗口且有关闭的阴影窗口
    if (closedShadows.Has(hWnd)) {
        ; 从映射中删除条目，微信重新激活后会自动创建阴影
        closedShadows.Delete(hWnd)
        return
    }
    
    ; 普通窗口阴影恢复
    ; 还原DWM渲染策略
    value := Buffer(4, 0)
    NumPut("Int", 0, value) ; DWMNCRP_USEWINDOWSTYLE = 0
    DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hWnd, "Int", 2, "Ptr", value, "Int", 4)
    
    ; 还原DWMWA_EXCLUDED_FROM_PEEK
    NumPut("Int", 0, value)
    DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hWnd, "Int", 12, "Ptr", value, "Int", 4)
    
    ; 检查是否需要恢复特定的窗口样式
    if (windowStates[hWnd].HasOwnProp("exStyle")) {
        ; 还原窗口扩展样式
        DllCall("SetWindowLong", "Ptr", hWnd, "Int", -20, "Int", windowStates[hWnd].exStyle)
    }
    
    ; 强制重绘窗口
    DllCall("SetWindowPos", "Ptr", hWnd, "Ptr", 0, "Int", 0, "Int", 0, 
           "Int", 0, "Int", 0, "UInt", 0x0027) ; SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_FRAMECHANGED
}

; 设置等待裁剪的光标样式 - 使用手型光标表示可以进行选择
SetWaitCursor() {
    ; 加载手型光标
    hHand := DllCall("LoadCursor", "Ptr", 0, "Ptr", 32649, "Ptr")  ; IDC_HAND = 32649
    
    ; 替换标准箭头光标为手型光标
    DllCall("SetSystemCursor", "Ptr", DllCall("CopyIcon", "Ptr", hHand, "Ptr"), "UInt", 32512)  ; OCR_NORMAL
}

; 清除所有ToolTip
ClearAllToolTips() {
    ; AutoHotkey支持最多20个ToolTip
    Loop 20 {
        ToolTip("", , , A_Index)
    }
}

; 检查当前活动窗口是否被裁剪的函数
IsActiveWindowClipped() {
    try {
        ; 获取当前活动窗口的句柄
        activeHwnd := WinGetID("A")
        
        ; 检查窗口是否在裁剪窗口列表中且确实有区域信息
        if (windowStates.Has(activeHwnd) && windowStates[activeHwnd].HasOwnProp("region")) {
            ; 确认region不为空并且有有效的尺寸
            region := windowStates[activeHwnd].region
            if (IsObject(region) && region.HasOwnProp("w") && region.HasOwnProp("h") && 
                region.w > 0 && region.h > 0) {
                return true
            }
        }
        
        return false
    } catch {
        ; 如果获取窗口信息失败（比如没有激活窗口），返回 false
        return false
    }
}

; 显示当前裁剪窗口信息的函数
ShowClipInfo() {
    try {
        activeHwnd := WinGetID("A")
        
        if (windowStates.Has(activeHwnd) && windowStates[activeHwnd].HasOwnProp("region")) {
            region := windowStates[activeHwnd].region
            winTitle := WinGetTitle(activeHwnd)
            
            infoText := "窗口: " winTitle "`n"
            infoText .= "裁剪区域: x=" region.x ", y=" region.y ", w=" region.w ", h=" region.h
            
            ShowToolTip(infoText, 3000)
        }
    } catch {
        ; 如果获取窗口信息失败，静默忽略
        return
    }
}

; 裁切区域移动函数 - 添加边界限制
MoveClipRegion(direction, step := 50) {
    try {
        activeHwnd := WinGetID("A")
        
        ; 确保窗口是已裁剪的
        if (!windowStates.Has(activeHwnd) || !windowStates[activeHwnd].HasOwnProp("region")) {
            return
        }
        
        ; 获取当前裁切区域
        region := windowStates[activeHwnd].region
        x := region.x
        y := region.y
        w := region.w
        h := region.h
        
        ; 获取窗口尺寸用于边界检查
        WinGetPos(, , &winWidth, &winHeight, "ahk_id " activeHwnd)
        
        ; 根据方向移动裁切区域，同时检查边界
        switch direction {
            case "up":
                y := Max(0, y - step)
            case "down":
                ; 确保裁切区域的底部不超出窗口
                y := Min(y + step, winHeight - h)
            case "left":
                x := Max(0, x - step)
            case "right":
                ; 确保裁切区域的右侧不超出窗口
                x := Min(x + step, winWidth - w)
        }
        
        ; 应用新的裁切区域
        ApplyClipToWindow(activeHwnd, x, y, w, h)
        
        /*
        ; 显示更新后的裁切区域信息
        infoText := "窗口: " WinGetTitle(activeHwnd) "`n"
        infoText .= "裁剪区域: x=" x ", y=" y ", w=" w ", h=" h "`n"
        infoText .= "窗口尺寸: " winWidth "x" winHeight
        ShowToolTip(infoText, 1000)
        */
    } catch {
        ; 如果获取窗口信息失败，静默忽略
        return
    }
}

; 调整裁剪区域大小
ResizeClipRegion(action, step := 20) {
    try {
        activeHwnd := WinGetID("A")
        
        ; 确保窗口是已裁剪的
        if (!windowStates.Has(activeHwnd) || !windowStates[activeHwnd].HasOwnProp("region")) {
            return
        }
        
        ; 获取当前裁切区域
        region := windowStates[activeHwnd].region
        x := region.x
        y := region.y
        w := region.w
        h := region.h
        
        ; 获取窗口尺寸用于边界检查
        WinGetPos(, , &winWidth, &winHeight, "ahk_id " activeHwnd)
        
        ; 根据操作调整裁剪区域大小
        switch action {
            case "widthIncrease":
                ; 确保不超出窗口右边界
                w := Min(w + step, winWidth - x)
            case "widthDecrease":
                ; 确保宽度不小于最小值
                w := Max(w - step, 20)
            case "heightIncrease":
                ; 确保不超出窗口底部边界
                h := Min(h + step, winHeight - y)
            case "heightDecrease":
                ; 确保高度不小于最小值
                h := Max(h - step, 20)
        }
        
        ; 应用新的裁切区域
        ApplyClipToWindow(activeHwnd, x, y, w, h)
        
        ; 显示调整后的区域信息
        infoText := "裁剪区域: " w "×" h
        ShowTooltip(infoText, 1000)
    } catch {
        ; 如果获取窗口信息失败，静默忽略
        return
    }
}

; 开始拖动裁切窗口 - 改进版
StartDragClippedWindow() {
    global isDragging, dragHwnd, startX, startY, offsetX, offsetY
    
    try {
        ; 获取当前活动窗口
        activeHwnd := WinGetID("A")
        
        ; 确保是被裁切的窗口
        if (!IsActiveWindowClipped())
            return
        
        ; 获取鼠标初始位置
        CoordMode("Mouse", "Screen")
        MouseGetPos(&startX, &startY)
        
        ; 获取窗口初始位置
        WinGetPos(&winX, &winY, , , "ahk_id " activeHwnd)
        
        ; 计算鼠标相对于窗口左上角的偏移量
        offsetX := startX - winX
        offsetY := startY - winY
        
        ; 设置状态
        isDragging := true
        dragHwnd := activeHwnd
        
        ; 使用SetTimer来实现窗口拖动
        SetTimer(DragWindowTimer, 6)
    } catch {
        ; 如果获取窗口信息失败（比如没有激活窗口），静默忽略
        return
    }
}

; 停止拖动裁切窗口
StopDragClippedWindow() {
    global isDragging, dragHwnd
    
    ; 停止计时器
    SetTimer(DragWindowTimer, 0)
    
    ; 重置状态
    isDragging := false
    dragHwnd := 0
    
    ; 清除提示
    ;ToolTip("")
}

; 拖动窗口的计时器函数 - 改进版
DragWindowTimer() {
    global isDragging, dragHwnd, offsetX, offsetY
    
    ; 如果没有在拖动，或者窗口不存在，停止计时器
    if (!isDragging || !WinExist("ahk_id " dragHwnd)) {
        SetTimer(DragWindowTimer, 0)
        isDragging := false
        return
    }
    
    ; 获取当前鼠标位置
    CoordMode("Mouse", "Screen")
    MouseGetPos(&currentX, &currentY)
    
    ; 计算新窗口位置 - 考虑鼠标在窗口内的偏移量
    newX := currentX - offsetX
    newY := currentY - offsetY
    
    ; 移动窗口到新位置
    WinMove(newX, newY, , , "ahk_id " dragHwnd)
}