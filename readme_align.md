# CapsLock++ 功能说明

## 0. 前言

本项目最初受[Capslock+](https://capslox.com/capslock-plus/)启发, 并借鉴了vim的很多键位.

主要自用于win11 24H2, 其他系统未测试, 应该也不会测试(除非作者换电脑).

项目本身用AutoHotkey v2编写, 配置助手使用python编写. 其中获取光标位置使用了[Tebayaki](https://github.com/Tebayaki)的[GetCratePosEx](https://github.com/Tebayaki/AutoHotkeyScripts/blob/main/lib/GetCaretPosEx), 翻译功能使用的是映射欧路词典的查词翻译快捷键, 没有欧路词典的话翻译功能(Caps+T)将无法使用, 作者买了欧路词典的终身会员与AI会员, 免费版是否可用未测试. 宏录制相关功能使用了开源项目tinytask, 屏幕刷新率相关功能(电源管理中)使用了Qres.

需要注意的是ahk的右键钩子与WGesture2(WGesture1未测试)冲突, 作者的做法是禁用WGesture2的右键相关功能, 只使用其移到屏幕上下边缘时的音量与亮度调整功能.

同时本readme只在typora的最新版开启mathjax的physics支持时可以完全渲染, vscode(cursor)中用MPE插件的mathjax或katex都有报错. 没有typora的话可以查看[README.pdf](./README.pdf)或[README.html](./README.html)文件.

## 1. 基本功能

### 1.1 CapsLock重映射

   $\begin{alignat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}\ 单击(\le0.3\mathrm{s}):&\ & 发送\mathrm{Esc}键\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}\ 长按(\geq0.3\mathrm{s}):&\ & 犹豫操作, 不触发Esc
\end{alignat}$

### 1.2 CapsLock状态管理

   $\begin{alignat}{2}
&\unicode{0x2022}\ \ 自动维持\ \mathrm{CapsLock}\hspace{2pt} &&关闭状态\ \\
&\unicode{0x2022}\ \ \mathrm{Ctrl}+\mathrm{CapsLock}:&&手动切换\ \mathrm{CapsLock}\ 状态
\end{alignat}$

### 1.3 控制与调试

   $\begin{alignat}{2}
&\unicode{0x2022}\ \ \mathrm{Alt}+\mathrm{CapsLock}+\mathrm{R}:&\  手动重启脚本\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Esc}:&\  临时禁用脚本\\
&\unicode{0x2022}\ \ \mathrm{Ctrl}+\mathrm{Alt}+\mathrm{I}:&\  显示调试信息
\end{alignat}$

## 2. 实用小功能

### 2.1 文件重命名

   $\unicode{0x2022}\ \ \mathrm{Capslock}+左键: 重命名文件, 无需选中再按快捷键或右键$

### 2.2 放大镜

   $\begin{alignat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Tab}&:&\ 放大镜\\
&\unicode{0x2022}\ \ \mathrm{Ctrl}+\mathrm{Alt}+滚轮&:&\ 放大缩小
\end{alignat}$

### 2.3 快速搜索

   $\begin{alignat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Q}:\ 搜索选中的文本内容\\
&\qquad \unicode{0x2218}\ \ 如果选中的是网址, 则直接打开该网址\\
&\qquad \unicode{0x2218}\ \ 如果选中的是磁盘路径(绝对路径), 则打开该路径\\
&\qquad \unicode{0x2218}\ \ 如果选中的是普通文本, 则使用\ \mathrm{bing}\ 搜索该文本\\
&\qquad \unicode{0x2218}\ \ 如果没有选中任何内容, 则不执行操作
\end{alignat}$

### 2.4 查词翻译

   $\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{T}
\left\{
\begin{alignedat}{1}
&打开查词面板(如果未选中任何内容)\\
&翻译当前选中内容  
\end{alignedat}
\right.  
\\$

### 2.5 DPI临时调整

$\begin{alignedat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Alt}& 时鼠标移动调整为临时\mathrm{DPI}\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Alt}+\hspace{-8pt}&滚轮上下: 调整临时\mathrm{DPI}
\end{alignedat}$

### 2.6 悄悄话功能

   $\begin{alignat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\Large{\textasciitilde} &:&\  微信窗口透明化与预览\\
&\qquad \unicode{0x2218}\ \ \rlap{将当前鼠标所在的微信窗口设置为完全透明且鼠标穿透}\\
&\qquad \unicode{0x2218}\ \ \rlap{在屏幕底部显示该窗口内容的预览}\\
&\qquad \unicode{0x2218}\ \ \rlap{再次按下该快捷键恢复窗口原样}\\
&\unicode{0x2022}\ \ \mathrm{Shift}+滚轮上&:&\  微信预览窗口内容上移\\
&\unicode{0x2022}\ \ \mathrm{Shift}+滚轮下&:&\  微信预览窗口内容下移
\end{alignat}$

## 3. 文本编辑增强

### 2.1 光标移动

   $\begin{alignedat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{A}&:&\  向左移动一个单词\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{G}&:&\  向右移动一个单词\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{S}&:&\  向左移动一个字符\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{F}&:&\  向右移动一个单词\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{E}&:&\  向上移动一行\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{D}&:&\  向下移动一行\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{W}&:&\  移动到行首\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{R}&:&\  移动到行尾\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\rlap{\mathrm{Alt}+\mathrm{A}: 移动到文件开头}\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\rlap{\mathrm{Alt}+\mathrm{G}: 移动到文件末尾}
\end{alignedat}$

### 2.2 文本选择

   $\begin{alignedat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{H}&:&\  向左选择一个单词\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\ ; &:&\  向右选择一个单词\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{J}&:&\  向左选择一个字符\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{L}&:&\  向右选择一个单词\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{I}&:&\  向上选择一行\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{K}&:&\  向下选择一行\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{U}&:&\  选择到行首\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{O}&:&\  选择到行尾\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\rlap{\mathrm{空格}: 选择当前单词}\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\rlap{\mathrm{双击空格}: 选择当前单词}\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\rlap{\mathrm{Alt}+\mathrm{H}\hspace{0,8pt}: 选择到文件开头}\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\rlap{\mathrm{Alt}+\ \mathrm{;}\ \hspace{0.6pt}: 选择到文件末尾}\\
\end{alignedat}$

### 2.3 删除操作

   $\begin{alignedat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{<}&:&\  向左删除一个字符\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{>}&:&\  向右删除一个字符\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{M}&:&\  删除到行首\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\hspace{2.2pt} \mathrm{?}&:&\  删除到行尾\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\rlap{\mathrm{Alt}+\mathrm{M}:删除到文件开头} \\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\rlap{\mathrm{Alt}+\hspace{2.2pt}\mathrm{?}\hspace{2.3pt}:删除到文件末尾} \\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\rlap{\mathrm{Backspace}:删除整行}
\end{alignedat}$

### 2.4 自定义跳转

$输入跳转字符数, \mathrm{enter}跳转, \mathrm{CapsLock/Esc/Delete/左右键}\ 退出,$$
可以使用\mathrm{BackSpace}删除最后一位, 清空后再次\mathrm{BackSpace}退出$

$输入0开启单词模式(针对水平方向\mathrm{SFJL,.}有效), 开启后再按\mathrm{BackSpace}可退回字符模式$

   $\begin{alignedat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Alt}+\mathrm{S}&:&\ 向左跳转指定字符数\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Alt}+\mathrm{F}&:&\ 向右跳转指定字符数\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Alt}+\mathrm{E}&:&\ 向上跳转指定行数\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Alt}+\mathrm{D}&:&\ 向下跳转指定行数\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Alt}+\mathrm{J}&:&\ 向左选择指定字符数\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Alt}+\mathrm{L}&:&\ 向右选择指定字符数\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Alt}+\mathrm{I}&:&\ 向上选择指定行数\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Alt}+\mathrm{K}&:&\ 向下选择指定行数\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Alt}+\mathrm{<}&:&\ 向左删除指定字符数\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Alt}+\mathrm{>}&:&\ 向右删除指定字符数  
\end{alignedat}$

### 2.5 特殊操作

   $\begin{alignedat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Z}&:&\ 撤销\\
&\begin{alignedat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{X}\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{C}\\
\end{alignedat}
&&\hspace{-5.5pt}\left.\begin{alignedat}{1}
&:剪切\\
&:复制
\end{alignedat}\right\}独立剪切板
\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{V}&:&\ 粘贴\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Y}&:&\ 重做\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\rlap{\mathrm{Enter}\hspace{3.5pt} :\ 在当前行末尾插入换行}\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\rlap{\mathrm{RShift} :\ 在当前行上方插入空行 }
\end{alignedat}$

### 2.6 符号定位功能

   $\begin{alignedat}{1}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+ \mathrm{P} :定位到对应的配对符号位置(再次长按\mathrm{CapsLock}或\mathrm{Esc}取消)&&\\
&\qquad \unicode{0x2218}\ \ 英文标点: ()、[]、\{\}、\hspace{-5pt}<>\\
&\qquad \unicode{0x2218}\ \ 中文标点:\hspace{-5pt} 「」\hspace{-5pt}、\hspace{-6.5pt}『』\hspace{-5pt}、\hspace{-4.4pt}【】\hspace{-5pt}、\hspace{-2.2pt}《》\hspace{-5pt}、〈〉、\hspace{-5pt}\mathrm{（）}\hspace{-5pt}、\hspace{-5pt}\mathrm{［］}、\hspace{-5pt}\mathrm{｛｝}\hspace{-5pt}、\hspace{-5pt}〔〕\hspace{-5pt}、\hspace{-5pt}〖〗\hspace{-5pt}、\hspace{-5pt}〘〙\hspace{-5pt}、\hspace{-5pt}〚〛\hspace{-5pt}、\unicode{0x201C}\unicode{0x201D}、\unicode{0x2018}\unicode{0x2019}、‹›、«»
\end{alignedat}$

## 3. 窗口管理增强

### 3.1 窗口置顶

   $\unicode{0x2022}\ \ \mathrm{CapsLock}+右键 : 置顶/取消置顶光标所在窗口$

### 3.2 窗口切换

   $\begin{alignedat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+滚轮向下(或侧键1): 按\ \mathrm{PID}\ 正序切换非黑名单窗口\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+滚轮向上(或侧键2): 按\ \mathrm{PID}\ 逆序切换非黑名单窗口\\
&\left.
\begin{alignedat}{1}
&\unicode{0x2022}\ \ \mathrm{Alt}+\mathrm{Esc}:\qquad\hspace{13pt} 按\ \mathrm{PID}\ 正序切换非黑名单窗口 \\
&\unicode{0x2022}\ \ \mathrm{Alt}+\mathrm{Shift}+\mathrm{Esc}:按\ \mathrm{PID}\ 正序切换非黑名单窗口
\end{alignedat}
\right\} \tiny \begin{alignedat}{1}&增强原生快捷键,\\&可以添加黑名单\end{alignedat}
\end{alignedat}$

### 3.3 虚拟环境

   $\begin{alignedat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+中键:添加/删除光标下窗口到虚拟环境&&\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+Alt+中键:清空虚拟环境&&\\
&\qquad \unicode{0x2218}\ \ 使用虚拟环境后, 窗口切换功能只切换虚拟环境内窗口\\
&\qquad \phantom{\unicode{0x2218}}\ \ 顺序为添加时的顺序
\end{alignedat}$

### 3.4 工作区清理

   $\begin{alignedat}{1}&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Win}+\mathrm{Z}:最小化当前窗口之外或虚拟环境之外的窗口\\&\phantom{\unicode{0x2022}\ \ }若清理之后窗口布局未改变, 再次按下恢复\end{alignedat}$

### 3.5 窗口裁切工具

   $\begin{alignat}{3}
&\unicode{0x2022}\ \ \mathrm{Ctrl}+\mathrm{Shift}+\mathrm{X}&:&\  开始裁切当前窗口\\
&\unicode{0x2022}\ \ \mathrm{Ctrl}+\mathrm{Shift}+\mathrm{Z}&:&\  恢复窗口原始状态\\
&\unicode{0x2022}\ \ \mathrm{Ctrl}+滚轮上/下&:&\  上下移动裁切区域\\
&\unicode{0x2022}\ \ \mathrm{Ctrl}+\mathrm{Shift}+滚轮&:&\  左右移动裁切区域\\
&\unicode{0x2022}\ \ \mathrm{Ctrl}+\mathrm{Shift}+左键&:&\  拖动裁切过的窗口\\
&\unicode{0x2022}\ \ \mathrm{Ctrl}+\mathrm{Alt}+滚轮&:&\  增加/减少裁切区域高度\\
&\unicode{0x2022}\ \ \mathrm{Ctrl}+\mathrm{Shift}+\mathrm{Alt}\rlap{+滚轮:\  增加/减少裁切区域宽度}&&
\end{alignat}$

### 3.7 标签页切换

   $\begin{alignedat}{3}
&\unicode{0x2022}\ \ 右键+滚轮向下: 正序切换标签页\\
&\unicode{0x2022}\ \ 右键+滚轮向上: 逆序切换标签页\\
\end{alignedat}$

## 4. 鼠标模拟功能

### 4.1 鼠标移动

   $\begin{alignat}{1}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+方向键:移动鼠标指针\\
&\qquad
\left.
\begin{alignedat}{}
\unicode{0x2218}\ \ 上:向上移动\\
\unicode{0x2218}\ \ 下:向下移动\\
\unicode{0x2218}\ \ 左:向左移动\\
\unicode{0x2218}\ \ 右:向右移动
\end{alignedat}
\right\}支持左上, 右上, 左下, 右下\\[1ex]
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{Alt}+方向键: 精确移动鼠标指针
\end{alignat}$

### 4.2 鼠标点击

   $\begin{alignedat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{RAlt}&:&左键点击\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\underset{或RCtrl}{\mathrm{RWin}}&:&右键点击\\
&\qquad \unicode{0x2218}\ \ \rlap{支持模拟按住, 拖动和释放}
\end{alignedat}$

### 4.3 鼠标滚轮

   $\begin{alignedat}{3}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{PgUp}&:&\ 向上滚动\\
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{PgDn}&:&\ 向下滚动
\end{alignedat}$

## 5. 速记功能

### 5.1 基本操作

   $\begin{alignedat}{1}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\mathrm{N}:\ \text{打开速记窗口} \\
&\qquad \unicode{0x2218}\ \ \text{窗口中可以快速记录笔记、想法或任何文本内容} \\
&\qquad \unicode{0x2218}\ \ \text{保存时自动添加时间戳记录创建时间} \\
&\qquad \unicode{0x2218}\ \ \text{如果不需要标题, 可以直接删除预填的"\#\# "或直接换行}
\end{alignedat}$

### 5.2 保存选项

   $\begin{alignedat}{1}
&\unicode{0x2022}\ \ \text{默认保存:} \\
&\qquad \unicode{0x2218}\ \ \text{如果笔记有标题 (以"\#\# "开头的第一行), 则以标题为文件名} \\
&\qquad \unicode{0x2218}\ \ \text{如果没有标题, 则以当前时间戳为文件名} \\
&\qquad \unicode{0x2218}\ \ \text{保存位置: 桌面的"速记\\默认"文件夹} \\[1ex] % 增加垂直间距
&\unicode{0x2022}\ \ \text{指定目标保存:} \\
&\qquad \unicode{0x2218}\ \ \text{在笔记最后一行使用 "\texttt{==目标==}" 格式指定保存位置} \\
&\qquad \unicode{0x2218}\ \ \text{支持的预设目标:} \\
&\qquad\qquad \bullet\ \text{论文: 保存到 "速记\\论文灵感.txt"} \\
&\qquad\qquad \bullet\ \text{日记: 保存到 "速记\\日记.txt"} \\
&\qquad\qquad \bullet\ \text{工作: 保存到 "速记\\工作.txt"} \\
&\qquad\qquad \bullet\ \text{想法: 保存到 "速记\\想法.txt"} \\
&\qquad \unicode{0x2218}\ \ \text{也可以指定默认文件夹中已存在的文件名:} \\
&\qquad\qquad \bullet\ \text{例如 "\texttt{==我的项目==}" 会查找并追加到 "速记\\默认\\我的项目.txt"} \\
&\qquad\qquad \bullet\ \text{如果指定的文件不存在且不是预设目标, 则忽略该指定}
\end{alignedat}$

### 5.3 快捷键

   $\begin{alignedat}{1}
&\unicode{0x2022}\ \ \mathrm{Ctrl}+\mathrm{S} : \text{保存笔记} \\
&\unicode{0x2022}\ \ \mathrm{Esc/CapsLock}:\ \text{取消并关闭窗口}
\end{alignedat}$

### 5.4 特殊功能

   $\begin{alignedat}{1}
&\unicode{0x2022}\ \ \text{自动追加:} \\
&\qquad \unicode{0x2218}\ \ \text{如果保存到已存在的同名文件, 会自动追加内容} \\
&\qquad \unicode{0x2218}\ \ \text{在追加时会添加分隔空行和时间戳} \\
&\qquad \unicode{0x2218}\ \ \text{时间戳格式为 "[yyyy-MM-dd HH:mm:ss]", 不会被作为标题处理} \\[1ex]
&\unicode{0x2022}\ \ \text{环境适应:} \\
&\qquad \unicode{0x2218}\ \ \text{使用环境变量自动适配不同用户的桌面路径} \\
&\qquad \unicode{0x2218}\ \ \text{首次使用时自动创建所需的目录结构} \\[1ex]
&\unicode{0x2022}\ \ \text{内容处理:} \\
&\qquad \unicode{0x2218}\ \ \text{如果第一行只有 "\#\# " 或被删除, 该行会被自动移除 (无标题)} \\
&\qquad \unicode{0x2218}\ \ \text{所有时间戳都以 "[时间戳]" 格式显示, 不会被作为标题}
\end{alignedat}$

## 6. 快捷菜单系统

### 6.1 基本操作

   $\begin{alignedat}{1}
&\unicode{0x2022}\ \ \mathrm{CapsLock}+\text{数字键}(1-0): \text{显示对应的快捷菜单组} \\
&\qquad \unicode{0x2218}\ \ 1: \text{电源/进程管理} \\
&\qquad \unicode{0x2218}\ \ 2: \text{文档/网站} \\
&\qquad \unicode{0x2218}\ \ 3: \text{文本/开发工具} \\
&\qquad \unicode{0x2218}\ \ 4: \text{Office应用} \\
&\qquad \unicode{0x2218}\ \ 5: \text{Adobe应用} \\
&\qquad \unicode{0x2218}\ \ 6: \text{网盘应用} \\
&\qquad \unicode{0x2218}\ \ 7: \text{数学/物理} \\
&\qquad \unicode{0x2218}\ \ 8-0: \text{预留菜单组（可自定义）}
\end{alignedat}$

### 6.2 菜单使用

   $\begin{alignedat}{1}
&\unicode{0x2022}\ \ \text{菜单显示后, 可以通过以下方式操作:} \\
&\qquad \unicode{0x2218}\ \ \text{点击对应按钮执行操作} \\
&\qquad \unicode{0x2218}\ \ \text{按数字键}(1-9,0) \text{执行对应序号的操作} \\
&\qquad \unicode{0x2218}\ \ \text{按Esc键或点击"关闭"按钮关闭菜单} \\
&\qquad \unicode{0x2218}\ \ \text{点击菜单外区域自动关闭菜单}\\
&\qquad \unicode{0x2218}\ \ \text{对于进程管理的两项,可以Ctrl+左键单击进入对话框精细操作}  
\end{alignedat}$

### 6.3 特殊功能

   $\begin{alignedat}{1}
&\unicode{0x2022}\ \ \text{电源计划管理:} \\
&\qquad \unicode{0x2218}\ \ \text{节电模式: 降低刷新率(需将\texttt{Qres.exe}放到\texttt{System32}路径中)、禁用不必要服务、优化网络设置} \\
&\qquad \unicode{0x2218}\ \ \text{平衡模式: 恢复标准设置} \\
&\qquad \unicode{0x2218}\ \ \text{性能模式: 启用高性能设置} \\[1ex]
&\unicode{0x2022}\ \ \text{进程管理:} \\
&\qquad \unicode{0x2218}\ \ \text{启用进程: 启动预设的常用程序} \\
&\qquad \unicode{0x2218}\ \ \text{终止进程: 关闭预设的后台程序} \\[1ex]
&\unicode{0x2022}\ \ \text{网站速启:} \\
&\qquad \unicode{0x2218}\ \ \text{快速打开常用网站, 根据配置的默认站点和浏览器}\\
&\qquad \unicode{0x2218}\ \ \text{支持多个预设网站, 按Ctrl+左键点击可进入选择界面} \\
&\qquad \unicode{0x2218}\ \ \text{可配置偏好的浏览器打开指定网站}
\end{alignedat}$

## 7. 自定义指南

   $\begin{alignedat}{1}
&\qquad \unicode{0x2218}\ \ \text{所有配置均在脚本目录下的 \texttt{CapsLock++.ini} 文件中修改 } \\
   &\qquad \unicode{0x2218}\ \ \text{修改并保存INI文件后,通常需要重启脚本 (\texttt{Alt+CapsLock+R}) 才能生效 }
\end{alignedat}$

### 7.1 黑名单配置 (虚拟环境/窗口切换)

   $\begin{alignedat}{1}
&\text{用于排除某些不希望被窗口切换或虚拟环境管理的程序或窗口 } \\
&\unicode{0x2022}\ \ \texttt{[blacklist\_virtual\_env]} \\
&\qquad \unicode{0x2218}\ \ \text{基于进程可执行文件路径排除 } \\
&\qquad \unicode{0x2218}\ \ \text{格式: \texttt{blackX = "进程路径"} (X为数字, 路径用引号包裹)} \\
&\unicode{0x2022}\ \ \texttt{[blacklist\_classes\_virtual\_env]} \\
&\qquad \unicode{0x2218}\ \ \text{基于窗口类名排除 } \\
&\qquad \unicode{0x2218}\ \ \text{格式: \texttt{blackclassesX = "窗口类名"} (X为数字, 类名用引号包裹)}
\end{alignedat}$

### 7.2 速记功能配置

   $\begin{alignedat}{1}
&\unicode{0x2022}\ \ \texttt{[noteTargets]} \\
&\qquad \unicode{0x2218}\ \ \text{定义速记窗口的预设保存目标 } \\
&\qquad \unicode{0x2218}\ \ \text{格式:} \\
&\qquad\qquad \bullet\ \texttt{noteX1 = "关键字"} \\
&\qquad\qquad \bullet\ \texttt{noteX2 = "目标文件路径"} \\
&\qquad \unicode{0x2218}\ \ \text{"关键字" 用于在速记窗口最后一行 \texttt{==关键字==} 中引用 } \\
&\qquad \unicode{0x2218}\ \ \text{"目标文件路径" 可以是绝对路径, 或相对于桌面的路径 (如 \texttt{速记\\日记.txt}) }
\end{alignedat}$

### 7.3 快捷菜单系统配置

   $\begin{alignedat}{1}
&\unicode{0x2022}\ \ \texttt{[MenuGroupsColourMode]} \\
&\qquad \unicode{0x2218}\ \ \texttt{DarkMode = true} \text{ (深色模式) 或 } \texttt{false} \text{ (浅色模式)} \\[1ex]
&\unicode{0x2022}\ \ \texttt{[MenuGroupsEnable]} \\
&\qquad \unicode{0x2218}\ \ \text{控制 10 个菜单组是否启用 } \\
&\qquad \unicode{0x2218}\ \ \text{格式: \texttt{enableGroupX = true/false} (X为1-10)} \\[1ex]
&\unicode{0x2022}\ \ \texttt{[MenuGroupNum]} \\
&\qquad \unicode{0x2218}\ \ \text{指定实际启用的菜单组数量 (影响切换和显示) } \\
&\qquad \unicode{0x2218}\ \ \text{格式: \texttt{num = 数量}} \\[1ex]
&\unicode{0x2022}\ \ \texttt{[MenuGroupCount]} \\
&\qquad \unicode{0x2218}\ \ \text{指定每个菜单组包含的项目数量 } \\
&\qquad \unicode{0x2218}\ \ \text{格式: \texttt{countX = 数量} (X为1-10)} \\[1ex]
&\unicode{0x2022}\ \ \texttt{[MenuGroupName]} \\
&\qquad \unicode{0x2218}\ \ \text{定义每个菜单组的标题名称 } \\
&\qquad \unicode{0x2218}\ \ \text{格式: \texttt{nameX = "组名称"} (X为1-10)} \\[1ex]
&\unicode{0x2022}\ \ \texttt{[MenuGroupsXItems]} \text{ (X为1-10, 代表对应菜单组)} \\
&\qquad \unicode{0x2218}\ \ \text{定义每个菜单组内的具体项目 每个项目包含 4 行:} \\
&\qquad\qquad \bullet\ \texttt{nameY = "项目显示名称"} \\
&\qquad\qquad \bullet\ \texttt{iconY = "图标"} \text{ (Emoji字符 或 文件路径)} \\
&\qquad\qquad \bullet\ \texttt{icontypeY = "icon类型"} \text{ (\texttt{emoji} 或 \texttt{file})} \\
&\qquad\qquad \bullet\ \texttt{actionY = "执行动作"} \\
&\qquad \unicode{0x2218}\ \ \text{图标路径:} \\
&\qquad\qquad \bullet\ \text{相对路径: 基于脚本目录, 如 \texttt{Icon\\Typora.ico}} \\
&\qquad\qquad \bullet\ \text{绝对路径: 如 \texttt{C:\\Windows\\System32\\shell32.dll,3}} \\
&\qquad \unicode{0x2218}\ \ \text{执行动作 (\texttt{actionY}):} \\
&\qquad\qquad \bullet\ \text{电源计划: \texttt{SetPowerPlan("模式名")}} (模式名: 节电, 平衡, 性能) \\
&\qquad\qquad \bullet\ \text{进程管理: \texttt{ManageProcessWithCtrlCheck("动作")}} (动作: 启用, 终止) \\
&\qquad\qquad \bullet\ \text{GitHub加速: \texttt{GitHubAccelerate()}} \\
&\qquad\qquad \bullet\ \text{网站登录: \texttt{WebsiteLogin()}} (\textit{需配合下方凭据配置}) \\
&\qquad\qquad \bullet\ \text{模拟按键: \texttt{SendInput("按键代码")}} (如 \texttt{SendInput("\#d")}) \\
&\qquad\qquad \bullet\ \text{激活或运行程序/打开文件夹: \texttt{ActivateOrRun("窗口标识", "运行命令")}} \\
&\qquad\qquad\qquad \triangleright\ \texttt{"窗口标识"} \text{可以是进程名 (\texttt{Typora.exe}), 类名等 } \\
&\qquad\qquad\qquad \triangleright\ \texttt{"运行命令"} \text{可以是程序完整路径, 或特殊路径变量:} \\
&\qquad\qquad\qquad\qquad \unicode{0x25CF}\ \ \texttt{A\_MyDocuments} \text{ (我的文档)} \\
&\qquad\qquad\qquad\qquad \unicode{0x25CF}\ \ \texttt{A\_UserProfile} \text{ (用户文件夹)} \\
&\qquad\qquad\qquad\qquad \unicode{0x25CF}\ \ \text{例如: \texttt{ActivateOrRun("explorer.exe", A\_UserProfile "\\Downloads")}}
\end{alignedat}$

### 7.4 进程管理配置

   $\begin{alignedat}{1}
&\unicode{0x2022}\ \ \texttt{[ProcessesToStart]} \text{ 和 } \texttt{[ProcessesToTerminate]} \\
&\qquad \unicode{0x2218}\ \ \text{定义菜单中"启用/终止进程"直接点击时操作的进程列表 } \\
&\qquad \unicode{0x2218}\ \ \text{格式: \texttt{itemX = "进程名或启动路径"} (X为数字)} \\[1ex]
&\unicode{0x2022}\ \ \texttt{[GUIProcessesToStart]} \text{ 和 } \texttt{[GUIProcessesToTerminate]} \\
&\qquad \unicode{0x2218}\ \ \text{定义菜单中"启用/终止进程"按住\texttt{Ctrl}点击时, 弹出的选择窗口中的进程列表 } \\
&\qquad \unicode{0x2218}\ \ \text{格式 (每个进程3行):} \\
&\qquad\qquad \bullet\ \texttt{ItemX\_Name=进程友好名称} \\
&\qquad\qquad \bullet\ \texttt{ItemX\_Path=要启动/终止的进程名或路径} \\
&\qquad\qquad \bullet\ \texttt{ItemX\_Checked=true/false} \text{ (默认是否勾选)}
\end{alignedat}$

### 7.5 网站速启配置

   $\begin{alignedat}{1}
&\unicode{0x2022}\ \ \texttt{[CommonWebsites]} \\
&\qquad \unicode{0x2218}\ \ \text{配合菜单动作 \texttt{WebsiteLogin()} 使用 } \\
&\qquad \unicode{0x2218}\ \ \text{格式:} \\
&\qquad\qquad \bullet\ \texttt{default\_site = 默认网站URL} \\
&\qquad\qquad \bullet\ \texttt{browser = 浏览器选项} \text{ (edge, chrome, firefox, default)} \\
&\qquad\qquad \bullet\ \texttt{site1 = 网站名称1} \\
&\qquad\qquad \bullet\ \texttt{url1 = 网站URL1} \\
&\qquad \unicode{0x2218}\ \ \text{可配置多个网站, 通过 site2/url2, site3/url3 等继续添加} \\
&\qquad \unicode{0x2218}\ \ \text{按Ctrl+左键点击菜单项时可进入选择对话框}
\end{alignedat}$

### 7.6 配置助手使用说明

   $\begin{alignedat}{1}
&\text{为了简化 CapsLock++.ini 文件的编辑过程, 提供了图形化的配置助手工具} \\
&\text{(\texttt{配置助手.py} 或编译后的 \texttt{配置助手.exe}) } \\
&\text{运行此工具可以直观地修改 \texttt{CapsLock++.ini} 中的大部分配置项 } \\[1ex]
&\unicode{0x2022}\ \ \textbf{主界面按钮:} \\
&\qquad \unicode{0x2218}\ \ \texttt{重新加载:} \text{ 从 \texttt{CapsLock++.ini} 重新载入所有配置到助手界面 } \\
&\qquad \unicode{0x2218}\ \ \texttt{保存配置:} \text{ 将助手界面中的所有修改保存回 \texttt{CapsLock++.ini}, 并尝试重启主脚本} \\
&\qquad\qquad \bullet\ \text{注意: 保存后通常需要主脚本重启才能完全生效} \\[1ex]
&\unicode{0x2022}\ \ \textbf{选项卡说明:} \\
&\qquad \unicode{0x2218}\ \ \texttt{黑名单窗口:} \\
&\qquad\qquad \bullet\ \text{管理进程黑名单和窗口类名黑名单 } \\
&\qquad\qquad \bullet\ \text{可通过 "选择窗口" 按钮捕获目标窗口的进程名和类名 } \\
&\qquad \unicode{0x2218}\ \ \texttt{速记路径:} \\
&\qquad\qquad \bullet\ \text{配置速记功能的关键词和对应的文件保存路径 } \\
&\qquad\qquad \bullet\ \text{支持添加, 编辑, 删除和上下移动条目 } \\
&\qquad \unicode{0x2218}\ \ \texttt{菜单配置:} \\
&\qquad\qquad \bullet\ \text{管理 10 个快捷菜单组和其中的菜单项 } \\
&\qquad\qquad \bullet\ \text{左侧面板: 编辑组名称, 启用/禁用组, 上下移动组顺序 } \\
&\qquad\qquad \bullet\ \text{右侧面板: 添加, 编辑, 删除菜单项, 上下移动项顺序 } \\
&\qquad\qquad \bullet\ \text{可设置菜单项的名称, 图标 (Emoji或文件), 图标类型, 执行动作 } \\
&\qquad\qquad \bullet\ \text{可在此处切换菜单的 深色/浅色 模式 } \\
&\qquad \unicode{0x2218}\ \ \texttt{进程管理:} \\
&\qquad\qquad \bullet\ \text{管理菜单动作 \texttt{ManageProcessWithCtrlCheck()} 的相关配置 } \\
&\qquad\qquad \bullet\ \text{包含四个子选项卡: "直接启用", "直接终止", "Ctrl+启用", "Ctrl+终止" } \\
&\qquad\qquad \bullet\ \text{可添加, 删除, 移动进程条目 } \\
&\qquad\qquad \bullet\ \text{在 "Ctrl+" 模式下可设置默认是否勾选 } \\
&\qquad \unicode{0x2218}\ \ \texttt{网站配置:} \\
&\qquad\qquad \bullet\ \text{配置菜单动作 \texttt{WebsiteLogin()} 的相关设置 } \\
&\qquad\qquad \bullet\ \text{设置默认网站URL和偏好的浏览器 } \\
&\qquad\qquad \bullet\ \text{管理网站列表 (添加, 删除, 移动) }
\end{alignedat}$
