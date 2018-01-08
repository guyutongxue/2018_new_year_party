# 2018_new_year_party

### 2018年班级联欢晚会PPT后台系统

## 总述

考虑到班级联欢会的自由度，吸取去年的沉痛教训，我决定将此次班级联欢会的PPT制作为可以随时后台更改的模式。这样可以灵活控制时间、结构，添加内容，并实现一些意想不到的内容。

初步主要方法是：

- 一台终端上传更改数据至网络；
- PPT从网络获取数据；
- 把数据显示在PPT中。

随后经过详细的理论和实际测试，完善为如下方案：

- 采用越狱的iPhone作为后台终端
- iPhone上安装Git，添加SSH-RSA证书，使用Shell脚本自动提交、上传更改至GitHub
- GitHub上搭建GitHub Pages，可以通过HTTP(S)获取数据
- 使用MinGW内的wget(wget64.exe)来获取GitHub数据
- PPT中使用VBA宏：
	- 用宏调用cmd批处理文件，执行 `wget64.exe` 获取数据到文件
	- 用宏和Windows API从文件读取字符串
	- 用宏读取字符串并添加到PPT中

下面进行逐步说明。

## iPhone 后台终端搭建

### 配置软件和证书

- 安装Git。直接从Cydia安装即可。用 `dpkg` 也无所谓。
- 配置SSH-RSA证书。

``` Shell
$ ssh-keygen -t rsa -C "guyutongxue@163.com"
```

### 本地数据结构

智障地，出于可拓展性考虑，我使用了如下的数据结构：

- next
	- .git
	- info.txt
- _codename1
	- info.txt
- _codename2
	- info.txt
- ...

每一个文件夹代表一张幻灯片，用codename来区分。next文件夹初始化为一个Git，用来保存实际演示时下一张幻灯片，并随时准备提交和上传。

### info.txt

info.txt存储下一张幻灯片的内容。格式为：

```
text
_title
_text1
_text2
...
```
或者

```
slide
_n
```
第一行是 `text` 的话，显示一张基本格式的幻灯片。接下来一行 `_title` 为该张幻灯片标题，在接下来数行直至文件结束为该张幻灯片正文内容。

第一行是 `slide` 的话，跳转到任意幻灯片。接下来一行 `_n` 表示跳转到第几张幻灯片。

### 脚本

我写了三个脚本来半自动化进行修改、上传等操作。

#### create.sh

用于快速新建幻灯片内容。

```Shell
echo "name: "
read name
mkdir $name
cd $name
touch info.txt
echo "finish"
```

#### switch.sh

用于切换下一张幻灯片。

```Shell
echo "name:"
read name
cp $name/info.txt next
echo "finish."
```

#### save.sh

用于提交并上传下一张幻灯片到GitHub。

```Shell
cd next
git add .
git commit -m "Modify"
git push
echo "finish."
```

## PPT VBA 宏

### 获取GitHub Pages文件到本地

使用VBA调用cmd：

```VB
Sub update()
    Shell ("cmd.exe /c " + Application.ActivePresentation.Path + "\update.cmd")
End Sub
```
update.cmd：运行wget64.exe获取数据

```BAT
del %~dp0\info.txt
%~dp0\wget64.exe -O %~dp0\info.txt https://guyutongxue.github.io/2018_new_year_party/info.txt
```

### 从文本文件中获取字符串数据

由于微软的尿性，UTF-8字符需要转换一下才能使用。直接从网上抄来代码然后改成64位系统的。

```VB
'UTF-8文件读取.API函数声明
Option Explicit
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001

'UTF-8文件读取.读文件至变量
Private Function GetFile(FileName As String) As String
    Dim i As Integer, BB() As Byte
    If Dir(FileName) = "" Then Exit Function
    i = FreeFile
    ReDim BB(FileLen(FileName) - 1)
    Open FileName For Binary As #i
    Get #i, , BB
    Close #i
    GetFile = BB
End Function

'UTF-8文件读取.把UTF-8字符转化成ANSI字符
Public Function UTF8_Decode(FileName As String) As String
    Dim sUTF8 As String
    Dim lngUtf8Size As Long
    Dim strBuffer As String
    Dim lngBufferSize As Long
    Dim lngResult As Long
    Dim bytUtf8() As Byte
    Dim n As Long
    sUTF8 = GetFile(FileName)
    If LenB(sUTF8) = 0 Then Exit Function
    On Error GoTo EndFunction
    bytUtf8 = sUTF8
    lngUtf8Size = UBound(bytUtf8) + 1
    lngBufferSize = lngUtf8Size * 2
    strBuffer = String$(lngBufferSize, vbNullChar)
    lngResult = MultiByteToWideChar(CP_UTF8, 0, bytUtf8(0), _
        lngUtf8Size, StrPtr(strBuffer), lngBufferSize)
    If lngResult Then
        UTF8_Decode = Left(strBuffer, lngResult)
    End If
EndFunction:
End Function
```

### 获取到字符串后应用在PPT中

```VB
'sleep for a while...
Sub Sleep(second As Integer)
    Dim tt
tt = Timer
Do Until Timer - tt > 1
DoEvents
Loop
End Sub

Sub switch()
    Call update
    Call Sleep(2)
    Dim text As String
    Dim strArr As Variant
    text = UTF8_Decode(Application.ActivePresentation.Path + "\info.txt")
    strArr = Split(text, Chr(10))
    If strArr(0) = "text" Then
        Application.ActivePresentation.Slides(2).Shapes("captain").TextFrame.TextRange.text = strArr(1)
        Application.ActivePresentation.Slides(2).Shapes("text").TextFrame.TextRange.text = ""
        Dim i As Integer
        For i = 2 To UBound(strArr) - 1
            Application.ActivePresentation.Slides(2).Shapes("text").TextFrame.TextRange.text = Application.ActivePresentation.Slides(2).Shapes("text").TextFrame.TextRange.text + strArr(i) + Chr(10)
        Next
        Application.ActivePresentation.Slides(2).Shapes("text").TextFrame.TextRange.text = Application.ActivePresentation.Slides(2).Shapes("text").TextFrame.TextRange.text + strArr(UBound(strArr))
        Application.SlideShowWindows(1).View.GotoSlide (2)
    ElseIf strArr(0) = "slide" Then
        Application.SlideShowWindows(1).View.GotoSlide (Val(strArr(1)))
    End If
End Sub
```
注：其中 `Slides(2)` 这一页有两个元素：

- 标题占位符 captain
- 内容占位符 text

执行switch宏，就可以实现下一页幻灯片的播放了。

（未完待续）