#Requires AutoHotkey v2.0

; CapsLockNative.ahk - 恢复CapsLock原生功能并提供切换回CapsLock++的快捷键

; 显示启动提示
ToolTip("已临时切换回CapsLock原生模式")
SetTimer () => ToolTip(), -2000

; 确保CapsLock可以正常工作
SetCapsLockState(-1)  ; 尝试释放CapsLock状态控制

; 定义CapsLock+Esc快捷键用于启动CapsLock++
#HotIf GetKeyState("CapsLock", "P")
Escape::
{
    ; 检查文件是否存在并决定使用哪个扩展名
    capsLockPlusPlus := A_ScriptDir "\CapsLock++.ahk"
    if (!FileExist(capsLockPlusPlus))
        capsLockPlusPlus := A_ScriptDir "\CapsLock++.exe"
    
    ; 启动CapsLock++脚本
    Try {
        Run(capsLockPlusPlus)
    } Catch as err {
        MsgBox("无法启动CapsLock++: " err.Message)
        Return
    }
    
    ; 显示提示
    ToolTip("正在启动CapsLock++...")
    SetTimer () => ToolTip(), -2000
    
    ; 退出当前脚本
    ExitApp
}
#HotIf

; 添加托盘图标和菜单
TraySetIcon(A_ScriptDir "\Icon\CapsLock.ico")  ; 使用自定义CapsLock图标
A_IconTip := "CapsLock 原生模式"

; 创建托盘菜单
trayMenu := A_TrayMenu
trayMenu.Delete()  ; 清除默认菜单项
trayMenu.Add("切换到CapsLock++", SwitchToCapsLockPlusPlus)
trayMenu.Add()  ; 添加分隔线
trayMenu.Add("退出", (*) => ExitApp())

; 切换到CapsLock++的函数
SwitchToCapsLockPlusPlus(*) {
    ; 检查文件是否存在并决定使用哪个扩展名
    capsLockPlusPlus := A_ScriptDir "\CapsLock++.ahk"
    if (!FileExist(capsLockPlusPlus))
        capsLockPlusPlus := A_ScriptDir "\CapsLock++.exe"
    
    Try {
        Run(capsLockPlusPlus)
        ExitApp  ; 切换后退出当前脚本
    } Catch as err {
        MsgBox("无法启动CapsLock++: " err.Message)
    }
}