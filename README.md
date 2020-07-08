
# 簡明Excel VBA
Last update date：06/02/2020 18:57

> `VBA` 縮寫於 *Visual Basic for Applications*。

<!-- TOC -->

## 目錄

- [x] [0x00 如何創建一個宏](#createAMacro) (English Version)
- [x] [0x01 語法說明](#explanation) (done)
    - [1.1 數據和數據類型](#1.1)
    - [1.2 常量和變量](#1.2)
    - [1.3 數組](#1.3)
    - [1.4 運算符](#1.4)
    - [1.5 語句結構](#1.5)
        - [1.5.1 選擇語句](#1.5.1)
        - [1.5.2 循環語句](#1.5.2)
        - [1.5.3 GoTo語句](#1.5.3)
    - [1.6 過程(Sub)和函數(Function)](#1.6)
      - [1.6.1 Sub 過程](#1.6.1)
      - [1.6.2 Function 函數](#1.6.2)
      - [1.6.3 VBA的參數傳遞](#1.6.3)
      - [1.6.4 ByRef vs ByVal](#1.6.4)
    - [1.7 正則表達式(Regular Expression)](#1.7)
    - [1.8 註釋（Comments code）](#1.8)
    - [1.9 補充](#1.9)
    - [1.10 示例](#1.10)
- [x] [0x02 VBA界面介紹](#layout) (done)
    - [2.1 整體界面說明](#2.1)
    - [2.2 工程資源管理器（Project Explore）說明](#2.2)
    - [2.3 設置VBA Macro Project 密碼保護](#2.3)
    - [2.4 常用快捷欄及窗口設置](#2.4)
- [x] [0x03 對象操作說明](#object-option) (done)
    - [3.1 對象簡述](#3.1)
    - [3.2 Application對象](#3.2)
- [x] [0x04 字符串 String 相關常用操作](#string-option) (done)
    - [4.1 Trim](#4.1)
    - [4.2 Instr 和 InStrRev (類似indexOf函數)](#4.2)
    - [4.3 Mid (類似substring函數)](#4.3)
    - [4.4 Left 和 Right](#4.4)
    - [4.5 Replace](#4.5)
    - [4.6 StrReverse 倒轉函數](#4.6)
    - [4.7 其他字符串函數](#4.7)
- [ ] [0x05 Excel 相關常用操作](#excel-option) (doing)
    - [5.1 Excel 基礎操作](#5.1)
    - [5.2 打開Excel兩種方式](#5.2)
    - [5.3 操作Excel工作表（Worksheet）](#5.3)
    - [5.4 Excel AutoFilter / Excel 自動篩選操作](#5.4)
    - [5.5 清理Excel數據相關操作](#5.5)
- [x] [0x06 文件 相關常用操作](#0x06) (done)
    - [6.1 判斷文件，文件夾等是否存在](#6.1)
    - [6.2 文件相關操作](#6.2)
    - [6.3 文件夾相關操作](#6.3)
    - [6.4 其他操作（獲取文件名等）](#6.4)
- [ ] [0x07 日期和時間 相關函數](#0x07) (done)
    - [7.1 Date, Time, Now 函數](#7.1)
    - [7.2 日期函數：Year, Month, Day](#7.2)
    - [7.3 CDate 和 DateValue 函數](#7.3)
    - [7.4 IsDate 函數](#7.4)
- [x] [0x10 VBA 轉換函數一覽](#0x10) (done) (*English Version*)
- [x] [0x90 VBA Best Practices（VB代碼規範/開發規約）](#0x90) (English Version)
- [ ] [0x08 Trouble shooting](#0x08) (doing)
    - [91.1 消除Excel保存時警告（Privacy Warning:this document contains macros...）](#19.1)
    - [91.2 清除Excel數據透視表中過濾器緩存（舊項目）](#91.2)
    - [91.3 解決辦法：The macros in this project are disabled. Please refer to ...](#91.3)
    - [91.4 解決辦法：添加一個宏文件(第三方插件)到快速訪問欄](troubleshootings/Macro2QuickToolBar.md)
    - [91.5 解決辦法：如何修改編輯一個.xlam文件/解決保存修改後的.xlam文件再次內容消失問題](troubleshootings/EditXlamFile.md)
    - [91.6 解決辦法：使用SaveAs方法保存.xlsx後，再次打開提示: 文件損壞,後綴名錯誤（格式錯誤）](troubleshootings/SaveAsIssue.md)
    - [91.7 解決辦法：Excel每次保存時都彈出警告：“此文檔中包含宏、Activex控件、XML擴展包信息”（office 2007/2010/365+）](#91.7)
    - [91.8 解決辦法：使用.xlam宏文件執行VBA程序時，操作excel無任何反應](#91.8)
- [x] [0x92 VBA示例代碼](#0x09) (done)
- [ ] [0x93 Excel-VBA 快捷鍵](#0x10) (doing)
- [x] [0x94 Excel-VBA Debug調試](Debug.md) (done)
- [x] [0xFF 學習資源列表](#docslist) (done)

<!-- /TOC -->


<a name="createAMacro"></a>
## 0x00 如何創建一個宏
*Ref：* [如何創建一個宏](CreateAMacro.md) (English Version)

<a name="explanation"></a>
## 0x01 語法說明

都知道學會了英語語法，再加上大量的詞彙基礎，就算基本掌握了英語了。
類似的要使用vba，也要入鄉隨俗，瞭解他的構成，簡單的說vba包含`數據類型`、
`變量`/`常量`、`對象`和常用的`語句結構`。

不過呢在量和複雜度上遠低於英語，不用那麼痛苦的記單詞了，所以vba其實很簡單的。
熟悉了規則之後剩下就是查官方函數啦，查Excel提供的可操作對象啦。

順帶一提的是，函數其實也很容易理解，方便使用。拿到一個函數，例如`Sum`，
只要知道它是求多個數的和就夠了，剩下的就是用了。例如`Sum(1000,9)`結果就是`1009`了。
函數的一大好處就是隱藏具體實現細節，提供簡潔的使用方法。

<a name="1.1"></a>
### 1.1 數據和數據類型

Excel裏的每一個單元格都是一個`數據`，無論是數字、字母或標點都是數據。
對數據排排隊，吃果果，對不同的數據扔到不同的籃子裏歸類，籃子就是`數據類型`了。

在Excel-vba中，`數據類型`只有`數值`、`文本`、`日期`、`邏輯`或`錯誤`五種類型。
前四種最爲常用。具體描述參見下表：


| 類型 | 類型名稱 | 範圍 | 佔用空間|聲明符號 | 備註|
|--------|-------|-----|--------|-----|----|
| **邏輯型**|
| 布爾 | Boolean|邏輯值True或False|2|
|**數值型**|
|字節| Byte | 0~255的整數|1|
|整數| Integer| -32768~32767|2|%|
|長整數|Long|-2147483648~2147483647|4|&|
|單精度浮點|Single||4|!|
|雙精度浮點|Double||4|#|
|貨幣|Currency||8|@|
|小數|Decimal||14|
|**日期型**|
|日期|Date|日期範圍:100/1/1~9999/12/31|8|
|**文本型**|
|變長字符串|String|0~20億||$|
|定長字符串|String|1~65400||
|**其他**|
|變體型|Variant(數值)|保存任意數值，也可以存儲Error,Empty,Nothing,Null等特殊數值|
|對象|Object|引用對象|4|

表1.1 VBA數據類型

補充一點是，數組就像一筐水果，裏面可以存不止一個數據。
他不是一個具體的數據類型，叫數據結構更合適些。

<a name="1.2"></a>
### 1.2 常量和變量

定義後不能被改變的量，就是`常量`；相反的`變量`就能修改具體值。

在vba裏，使用一個 變量/常量 要先聲明。

`常量`聲明方法如下：</br>
` Const 常量名稱 As 數據類型 = 存儲在常量中的數據`
例如：
```vba
Const PI As Single = 3.14 ' 定義一個浮點常量爲PI，值爲3.14
```

`變量`聲明方法如下：</br>
```vba
Dim 變量名 As 數據類型
```
變量名，必須**字母**或**漢字**開頭，**不能** 包含空格、句號、感嘆號等。

數據類型，對應上面 ↑　表1.1裏的那些

更多的聲明方法，跟`Dim`聲明的區別是作用範圍不同：
```vba
Private v1 As Integer   ' v1爲私有整形變量
Public v2 As String     ' v2爲共有字符串變量
Static v3 As Integer    ' v3爲靜態變量，程序結束後值不變

' 變量聲明之後，就可以賦值和使用了
v1 = 1009
v2 = "1009"
v3 = 1009

' 使用類型聲明符，可以達到跟上面同樣的效果
public v2$  ' 與 Public v2 As String 效果一樣

' 聲明變量時，不指定具體的類型就變成了Variant類型，根據需要轉換數據類型
Dim v4
```

<a name="1.3"></a>
### 1.3 數組

使用數組和對象時，也要聲明，這裏說下數組的聲明：
```vba
' 確定範圍的數組，可以存儲b - a + 1個數，a、b爲整數
Dim 數組名稱(a To b) As 數據類型

Dim arr(1 TO 100) As Integer ' 表示arr可以存儲100個整數
arr(100) '表示arr中第100個數據

' 不指定a，直接聲明時，默認a爲0
Dim arr2(100) As Integer ' 表示arr可以存儲101個整數,從0數
arr2(100) '表示arr2中第101個數據

' 多維數組
Dim arr3(1 To 3,1 To 3,1 To 3) As Integer ' 定義了一個三維數組，可以存儲3*3*3=27個整數

' 動態數組，不確定數組大小時使用
Dim arr4() As Integer   ' 定義arr4爲整形動態數組
ReDim arr4(1 To v1)     ' 設定arr4的大小，不能重新設定arr4的類型

```

除了用`Dim`做常規的數組的聲明，還有下面這些聲明數組的方式:
```vba
' 使用Array函數將已知的數據常量放到數組裏
Dim arr As Variant        ' 定義arr爲變體類型
arr = Array(1, 1, 2, 3, 5, 8, 13, 21) ' 將整數存儲到arr中,索引默認從0開始

' 使用Split函數分隔字符串創建數組
Dim arr2 As Variant
arr2 = Split("hello, world", ", ") ' 按,分隔字符串 hello,world 並賦值給arr2

' 使用Excel單元格區域創建數組
' 這種方式創建的數組，索引默認從1開始
Dim arr3 As Variant
arr3 = Range("A1:C3").Value   ' 將A1:C3中的數組存儲到arr3中
Range("A4:C6").Value= arr3    ' 將arr3中的數據寫入到A4:C6中的區域
```


**數組常用的函數**

|函數|函數說明|參數說明|示例|
|----|----|----|----|
|`UBound(Array arr, [Integer i])`|數組最大的索引值|`arr`：數組；`i`：整形，數組維數|
|`LBound(Array arr, [Integer i])`|數組最小的索引值|同上|
|`Join(Array arr, [String s])`|合併字符串|`arr`：數組；`s`：合併的分隔符|
|`Split(String str, [String s])`|分割字符串|`str`：待分割的字符串；`s`：分割字符串的分隔符|

函數說明

UBound(Array arr,[Integer i]);</br>
UBound爲函數名</br>
arr和i 爲UBound的的參數，用中括號括起來的表示i爲非必填參數</br>
arr和i 之前的Array，Integer表示對應參數的數據類型</br>

> 補充
> [VBA 內置函數列表](https://msdn.microsoft.com/zh-cn/library/office/jj692811.aspx)


<a name="1.4"></a>
### 1.4 運算符

運算符的作用是對數據進行操作，像加減乘除等。這塊不再具體說明，列一下vba中常用的運算符。

|運算符|作用|示例|
|----|----|----|
|**算術運算符**|
|+|求兩個數的和|
|-|求兩個數的差|
|*|求兩個數的乘積|
|/|求兩個數的商|
|`\`|求兩個數相除後所得商的整數|
|^|求一個數的某次方|
|Mod|求兩個數相除後所得的餘數| 10 Mod 9=1|
|**比較運算符**|
|=|比較兩個數據是否相等|相等返回 True;否則返回False|
|<>|不相等|
|<|小於|
|>|大於|
|<=|不大於|
|>=|不小於|
|Is|比較連個對象的引用關係|
|Like|比較兩個字符串是否匹配| String1 Like String2|
|**文本運算符**|
|+|連接兩個字符串|
|&|連接兩個字符串|
|**邏輯運算符**|
|And|邏輯與|
|Or|邏輯或|
|Not|邏輯非|
|Xor|邏輯抑或|`表達式1 Xor 表達式2`兩個表達式返回的值不相等時爲True|
|Eqv|邏輯等價|`表達式1 Eqv 表達式2`兩個表達式返回的值相等時爲True|
|Imp|邏輯蘊含|

```vba
' Like是個比較有用的運算符，常用來做匹配或模糊匹配。
' 在模糊匹配的時候，有一些通配符能方便模糊匹配規則的書寫
"這是一個demo1" Like "*demo1" = True    ' * 號表示匹配任意多個字符
"這是一個demo2" Like "????demo2" = True ' ? 號表示匹配任意單個字符
"這是一個demo3" Like "*demo#" = True    ' # 號表示匹配任意數字
```


#### 三目運算符

正常在VBA中沒有類似java的 `expression ? true : false` 寫法，但是可以使用 `IFF` 來代替：
```vba
x = IIF(expression, A, B）
x = IIF(條件, 如果成立A賦值給X, 如果不成立B賦值給X）
```

作用也等同於如下：
```
If ... Then
Else
End If
```


<a name="1.5"></a>
### 1.5 語句結構

程序通常都是順序依次執行的。語句結構用來控制程序執行的步驟，
一般有**選擇**語句、**循環** 語句。

<a name="1.5.1"></a>
#### 1.5.1 選擇語句

選擇語句用來判斷程序執行那一部分代碼

語法：If ... Then ... End If</br>
If選擇可以嵌套使用</br>

常用的三種形式：

1. 普通模式
```vba
If 10 > 3 Then
    操作1  ' 執行這一步
End If

' 增加Else和Else If邏輯
If 1 > 2 Then
    操作1
ElseIf 1 = 2 Then
    操作2
Else
    操作3  ' 執行這一步
End If
```

2. 嵌套If語句
```vba
If 10 > 3 Then
    If 1 > 2 Then
        操作1
    Else
        操作2  ' 執行這一步
    End If
Else
    操作3
End If
```

3. Select ... Case ... 多選一，類似於java中的 Switch ... Case ... 語句
```vba
Dim Length As Integer
Length = 10
Select Length
    Case Is >= 8
        操作1  ' 執行這一步
    Case Is > 20
        操作2
    Case Else
        操作3
End Select
```
sample code:

```vba
Private Sub switch_demo_Click()
    Dim MyVar As Integer
    MyVar = 1

    Select Case MyVar
        Case 1
            Debug.Print "The Number is the Least Composite Number"
        Case 2
            Debug.Print "The Number is the only Even Prime Number"
        Case 3
            Debug.Print "The Number is the Least Odd Prime Number"
        Case Else
            Debug.Print "Unknown Number"
    End Select
End Sub
```

<a name="1.5.2"></a>
#### 1.5.2 循環語句

循環語句用來讓程序重複執行某段代碼

1. 普通For ... Next循環</br>
語法：For 循環變量 = 初始值 To 終值 Step 步長</br>
注：在VBA循環中可以使用`Exit`關鍵字來跳出循環，類似於Java中的break，
在for循環中語法爲：`Exit For`，在do while循環中爲：`Exit Do`，也可以利用`GoTo`語句
跳出本次循環，詳見：[1.5.3 GoTo語句](#1.5.3)</br>
```vba
Dim i As Integer
For i = 1 To 10 Step 2 ' 設定i從1到10，每次增加2，總共執行5次
    操作1   ' 可以通過設定 Exit For 退出循環
Next i
```

2. For Each ... 循環</br>
語法：For Each 變量 In 集合或數組
```vba
Dim arr
Dim i As Integer
arr = Array(1, 2, 3, 4, 5)
For Each i In arr ' 定義變量i，遍歷arr數組
    操作1
Next i
```

3. Do ... While循環</br>
語法：</br>
- 前置循環條件：</br>
![Alt text](doc/source/images/dowhileloopsyntax.png)

- 後置循環條件：</br>
![Alt text](doc/source/images/dowhileloopsyntax_suffix.png)

Sample code:
```vba
Dim i As Integer
i = 1
Do While i < 5  ' 循環5次
    i = i + 1
Loop

' ===============================================
' 將判斷條件後置的Do...While
Dim i As Integer
i = 1
Do
    i = i + 1
Loop While i < 5 ' 循環4次
```

4. Do Until 直到...循環</br>
語法：</br>
Do Until 表達式    表達式爲真時跳出循環
```vba
Dim i As Integer
i = 5
Do Until i < 1  
    i = i - 1
Loop

' ===============================================
' 後置的Do Until
Dim i As Integer
i = 5
Do
    i = i - 1
Loop Until i < 1  
```

<a name="1.5.3"></a>
#### 1.5.3 GoTo語句

**GoTo**
無條件地分支直接跳轉到過程中指定的行。

**注：** GoTo語句大多用於錯誤處理時，但會影響程序結構，增加閱讀和代碼調試難度，
除非必要時，應儘量避免使用GoTo語句。

```vba
Sub TestGoTo

    Dim lngSum As Long, i As Integer
    i = 1

JUMPX:
    i = i + 1
    If i <= 100 Then GoTo JUMPX
    Debug.Print "1到100的自然數之和是：" & lngSum

End Sub
```

**CONTINUE**

循環中實現continue操作，類似java語言的continue直接跳出本次循環
```vba
Sub continueTest()
    Dim i

    For i = 0 To 5
        If i = 1 Then
            '// 跳轉到CONTINUE部分
            GoTo CONTINUE
        ElseIf i = 3 Then
            '// 跳轉到CONTINUE部分
            GoTo CONTINUE
        End If

        '//沒有GoTo語句的時候打印counter: i
        Debug.Print i

CONTINUE:   '// countinue跳轉塊，可以寫邏輯，如果沒有邏輯就直接進行下次循環
    Next

End Sub
```

`選擇`和`循環`提供了多種實現同一目的的語句結構，他們都能實現同樣的作用，
差別一般是初始條件。還有書寫的複雜度。正確的選擇要使用的語句結構，
代碼邏輯上會更清楚，方便人的閱讀。

**簡寫**

在操作對象的屬性時常常要先把對象調用路徑都寫出來，用`with`可以簡化這一操作
```vba
' 簡化前
WorkSheets("表1").Range("A1").Font.Name="仿宋"
WorkSheets("表1").Range("A1").Font.Size=12
WorkSheets("表1").Range("A1").Font.ColorIndex=3

' 使用`With`
With WorkSheets("表1").Range("A1").Font
    .Name = "仿宋"
    .Size = 12
    .ColorIndex =3
End With
```

<a name="1.6"></a>
### 1.6 過程(Sub)和函數(Function)

概述Sub和Function的區別：   

**Sub** 和 **Function** 是VBA提供的兩種封裝體。
* 利用宏錄製得到的就是`Sub`。
* `Sub` 定義時無需定義返回值類型，而 `Function` 一般需要用 “As 數據類型” 定義函數返回值類型。
* `Sub` 中沒有對過程名賦值的語句，而 `Function` 中有對函數名賦值的語句，一般在函數最後返回值，格式如下：
```vba
Set functionName = xxxxxx
```
* 調用過程：調用 Sub 過程與 Function 過程不同。調用 Sub 過程的是一個獨立的語句，而調用函數過程只是表達式的一部分。另外，自定義函數並不允許修改工作表和單元格格式 (A UDF will only return a value it won't allow you to change the properties of a cell/sheet/workbook. )。但是，與 Function 一樣，Sub 也可以修改傳遞給它們的任何變量的值。
* 調用 Sub 過程有三種方法：   [參見1.6.1](#1.6.1)   

~~以下語句都調用了名爲 ProcExcel 的 Sub 過程。~~

  ~~Call  ProcExcel (FirstArgument, SecondArgument) '使用Call關鍵字調用~~   
  ~~ProcExcel  FirstArgument, SecondArgument        '直接調用~~   
  ~~Application.Run "ProcExcel" FirstArgument, SecondArgument~~   


~~**注意** ：當使用 Call 語法時，**參數必須在括號內**。若省略 Call 關鍵字，則也必須省略參數兩邊的括號。~~


<a name="1.6.1"></a>
#### 1.6.1 Sub 過程
```vba
[Private|Public] [Static] Sub 過程名([參數列表 [As 數據類型]])
    [語句塊]
End Sub
' [Private|Public]定義過程的作用範圍
' [Static]定義過程是否爲靜態
' [參數列表]定義需要傳入的參數
```

調用`Sub`的方法有三種，使用 `Call`、<u>直接調用</u>和使用 `Application.Run`:

舉個例子：
![Alt text](/doc/source/images/1505555701907.png)

**注意** ：當使用 Call 語法時，**參數必須在括號內**。若省略 Call 關鍵字，則也必須省略參數兩邊的括號。

<a name="1.6.2"></a>
#### 1.6.2 Function 函數

vba內部提供了大量的函數，也可以通過`Function`來定義函數，實現個性化的需求。
```vba
[Public|private] [Static] Function 函數名([參數列表 [As 數據類型]]) [As 數據類型]
    [語句塊]
    [函數名=過程結果]
End Function
```
使用函數完成上面的例子：
![Alt text](/doc/source/images/1505556598033.png)


<a name="1.6.3"></a>
#### 1.6.3 VBA的參數傳遞

參數傳遞的方式有兩種，引用和傳值。
傳值，只是將數據的內容給到函數，不會對數據本身進行修改。
引用，將數據本身傳給函數，在函數內部對數據的修改將同樣的影響到數據本身的內容。

參數定義時，使用`ByVal`關鍵字定義傳值，子過程中對參數的修改不會影響到原有變量的內容。
默認情況下，過程是按引用方式傳遞參數的。在這個過程中對參數的修改會影響到原有的變量。
也可以使用`ByRef`關鍵字顯示的聲明按引用傳參。
```vba
Sub St1(ByVal n As Integer, ByRef range)
    ...Other code
End SUb
```

<a name="1.6.4"></a>
#### 1.6.4 ByRef vs ByVal

舉個簡單栗子來解釋值傳和引用傳遞的區別：   
可以參照[Create A Macro](CreateAMacro.md) 在工作表上放置一個command button，並添加以下代碼：

```
Dim x As Integer
x = 10

MsgBox Triple(x)
MsgBox x
```

在上述代碼中調用了`Triple`函數，按照如下步驟添加一個`Triple`函數模塊：

1. 打開 [Visual Basic Editor](CreateAMacro.md#visual-basic-editor)，點擊菜單欄中的 <U>I</U>nsert ，選擇插入一個 <U>M</U>odule.

2. 添加如下代碼：

```
Function Triple(ByRef x As Integer) As Integer

x = x * 3
Triple = x

End Function
```

當點擊 command button 的時候顯示如下結果：

![Alt text](/doc/source/images/ByRefandByVal/byref-result.png)

![Alt text](/doc/source/images/ByRefandByVal/byref-result.png)

3. 使用 `ByVal`替換`ByRef`:

```
Function Triple(ByVal x As Integer) As Integer

x = x * 3
Triple = x

End Function
```
當點擊 command button 的時候顯示如下結果爲：

![Alt text](/doc/source/images/ByRefandByVal/byref-result.png)

![Alt text](/doc/source/images/ByRefandByVal/byval-result-2.png)

**說明：** 當通過引用(ByRef)傳遞參數時，我們引用的是原始值。函數中`x`的值(原始值)發生了變化。因此，第二個MsgBox顯示的值爲30。當通過值傳遞(ByVal)參數時，我們是在向函數傳遞一個副本。原始值沒有改變。因此，第二個MsgBox顯示的值爲10(原始值)。

**總結：**
**ByRef** 傳遞一個指向變量的指針，因此任何更改都會在使用該變量的任何地方反映出來（改變一處，其他所有使用該變量的地方均會改變）。   
**ByVal** 將變量的副本傳遞給函數，因此對該變量的任何更改都不會影響其原始值。當使用ByVal傳遞一個對象，你傳遞的是一個指針的拷貝而不是原始的指針(**注意:** 不是對象的拷貝)


**注意：**

1. 數組變量（Array）總是通過ByRef傳遞（只適用於實際聲明爲 *Array* 的變量，不適用於`Variants`聲明的數組變量）。
2. VBA在不具體指定傳值方式的時候，默認爲`ByRef`方式傳值。

```
Function Triple(x As Integer) As Integer '當不聲明指定具體值傳遞還是引用傳遞的時候，VBA默認爲 ByRef 方式傳值

'Or

Function Triple(ByRef x As Integer) As Integer

```

<a name="1.7"></a>
### 1.7 正則表達式(Regular Expression)
在VBA中使用正則表達式，因爲正則表達式不是vba自有的對象，
故此要用它就必須採用兩種方式引用它：一種是前期綁定，另外一種是後期綁定。

前期綁定：就是手工勾選工具/引用中的Microsoft VBScript Regular Expressions 5.5；
然後在代碼中定義對象：`Dim regExp As New RegExp`；</br>
後期綁定：使用CreateObject方法定義對象：`CreateObject("vbscript.regexp")`

RegExp對象的屬性：
   - Global – 設置或返回一個Boolean值，該值指明在整個搜索字符串時模式是全部匹配還是隻匹配第一個。如果搜索應用於整個字符串，Global 屬性的值應該爲 True，否則其值爲 False。默認的設置爲True。
   - Multiline – 返回正則表達式是否具有標誌, 缺省值爲False。如果指定的搜索字符串分佈在多行，這個屬性是要設置爲True的。
   - IgnoreCase – 設置或返回一個Boolean值，指明模式搜索是否區分大小寫。如果搜索是區分大小寫的，則IgnoreCase 屬性應該爲False；否則應該設爲True。缺省值爲True。
   - Pattern – 設置或返回被搜索的正則表達式模式。被搜索的正則字符串表達式。它包含各種正則表達式字符。

RegExp對象的方法：
- Execute – 對指定的字符串執行正則表達式搜索。需要傳入要在其上執行正則表達式的文本字符串。正則表達式搜索的設計模式是通過RegExp對象的Pattern來設置的。Execute方法返回一個Matches集合，其中包含了在string中找到的每一個匹配的Match對象。如果未找到匹配，Execute將返回空的Matches集合。
- Replace – 替換在正則表達式查找中找到的文本。
- Test – 對指定的字符串執行一個正則表達式搜索，並返回一個Boolean值指示是否找到匹配的模式。Global屬性對Test方法沒有影響。如果找到了匹配的模式，Test方法返回True；否則返回False。
- MatchCollection對象與Match對象
匹配到的所有對象放在MatchCollection集合中，這個集合對象只有兩個只讀屬性：
- Count：匹配到的對象的數目
- Item：集合的又一通用方法，需要傳入Index值獲取指定的元素。
一般，可以使用ForEach語句枚舉集合中的對象。集合中對象的類型是Match。
- Match對象有以下幾個只讀的屬性：
    - FirstIndex – 匹配字符串在整個字符串中的位置，值從0開始。
    - Length – 匹配字符串的長度。
    - Value – 匹配的字符串。
    - SubMatches – 集合，匹配字符串中每個分組的值。作爲集合類型，有Count和Item兩個屬性。

Sample Code（前期綁定）：
```vba
Private Function IsStringDate(ByVal strDate As String)
    Dim strDatePattern
    ' 前期綁定
    Dim regEx As New RegExp, matches

    Dim str MatchContent As String

    strDatePattern = "^(([0-9])|([0-2][0-9])|([3][0-1]))\-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\-\d{4}$"

    With regEx
        .Global = True      ' 搜索字符串中的全部字符，如果爲假，則找到匹配的字符就停止搜索！
        .MultiLine = False  ' 是否指定多行搜索
        .IgnoreCase = True  ' 指定大小寫敏感（True）
        .Pattern = strDatePattern   ' 所匹配的正則
    End With

    If regEx.Test(strDate) Then     ' 如果與正則相匹配
        Set matches = regEx.Execute(strDate)
        MatchContent = matches(0).Value
    Else
        MatchContent = "Not Matched"
    End If

    IsStringDate = regEx.Test(strDate)

End Function
```

Sample Code（後期綁定）：
```vba
Function ExtractNumber(str As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("vbscript.regexp")  ' 後期綁定
    With regEx
        .Global = True       ' 搜索字符串中的全部字符，如果爲假，則找到匹配的字符就停止搜索！
        .Pattern = "\D"      ' 非數字字符的正則表達式
        ExtractNumber = .Replace(str, "")        ' 把非數字字符替換成空字符串
    End With
    Set regEx = Nothing      ' 清除內存中的對象變量的地址，即釋放內存。
End Function
```

<a name="1.8"></a>
### 1.8 註釋（Comments code）
> 個人覺得代碼註釋起着非常重要的作用。 --  *bluetata* 11/28/2018 18:40

註釋語句是用來說明程序中某些語句的功能和作用；VBA 中有兩種方法標識爲註釋語句。</br>
單引號 `'` 舉例：`' 定義全局變量`；可以位於別的語句之尾，也可單獨一行。</br>
`Rem` 舉例：`Rem 定義全局變量`；只能單獨一行

以下列舉出了不同級別的註釋代碼，也可以[點擊這裏](SampleCode.bas)查看 VBA Sample Code。

#### 1. 源碼概要註釋/Source version Comments Code</br>
在每個source文件的最開頭
```vba
'--------------------------------------
' Creation date : 03/05/2017  (cn)
' Last update   : 11/28/2018  (cn)
' Author(s)     : Sekito.Lv
' Contributor(s):
' Tested on Excel 2016
'--------------------------------------
```

#### 2. 區塊註釋/Use Title Blocks Comments code for Each Macro</br>
在每個Function或者Sub上下，根據個人風格，可以在緊貼在函數上面一行處，
也可以在函數名的下面一行處。
```vba
'=======================================================
' Program:   DoMemoData
' Desc:      Writes memo data to the memo sheet
' Called by: PrintControl
' Call:      DoMemoData wbkReport, oStopRow
' Arguments: wbkReport--Name of the report workbook
'            oStopRow--Number of the last row to process
' Comments: (1) RunReport initializes the m_oMemoRowNum
'               variable
'           (2) wksMemo doesn't need to be static. And
'               it's over-defined. Fix this at some
'               point.
' Changes----------------------------------------------
' Date        Programmer    Change
' 11/26/2018  Sekito.Lv     Written
' 11/28/2018  Sekito.Lv     Re-set memo object. This is
'                           needed at times in Excel 8
'                           when the report workbook must
'                           close then re-open.
'=======================================================
Sub DoMemoData(wbkReport As Workbook, oStopRow As Long)
```

#### 3. 行內註釋/Use In-Line Comments
```vba
' If this routine was called by the batch routine...
If g_bCalledByBatch Then

    'Get the reference of the changing date cell
    sDateRef = GetNameVal("ChgDateCell", 0, g_nReference)

    ' If the date name is empty, return null sDateFormula
    If sDateRef = g_sNull Then
        sDateFormula = g_sNull

    ' Else, get the beginning formula in the date cell
    Else
        sDateFormula = m_wbkReport.Worksheets(1). _
        Evaluate(sDateRef).Formula
    End If
Else
```

#### 4. 函數列表註釋/List of Function Comments</br>
一般緊挨着源碼概要註釋下面，與其空一行到兩行
```vba
'-------------------------------------
' List of functions :
' - 1  - PublicHolidayFr
' - 2  - WorkingDay
' - 3  - WorkableDay
' - 4  - NextWorkingDay
' - 5  - NextWorkableDay
' - 6  - PrevWorkingDay
'-------------------------------------
```

<a name="1.9"></a>
### 1.9 補充

- 在vba中使用 `'`進行代碼註釋
- 在很長的語句中使用`_`來分割成多行
- 在有很多嵌套判斷中，代碼的可讀性會變得很差，一般講需要返回的內容及時返回，減少嵌套
- `Sub`中默認按引用傳遞參數，所以注意使用，一般不要對外面的變量進行修改，將封裝保留在內部


- `Dim`和`Set`的關係及區分

很明顯的是 vba中使用Dim設定變量類型，Set將對象引用賦值給變量

```vba
' 將Range對象賦值給變量rg
Dim rg As Range         ' 聲明rg爲Range對象
Set rg = Range("A1")    ' 設定rg爲Range("A1")的引用，之後操作rg和操作Range("A1")一樣了

' 如果不使用Set，下面的代碼將報錯
Dim rg As Range
rg = Range("A1")   ' 這段代碼將報錯

' 在非顯示聲明rg的前提下，下面的代碼將會得到不一樣的結果
rg = Range("A1")       ' rg將會是Range("A1")的內容，rg的類型將會是一種基本類型，Integer/String等
Set rg = Range("A1")   ' 這種情況下，rg將會是Range對象
```

- VBA中變量用Dim定義和不用Dim定義而直接使用有何區別？

用Dim語句聲明變量就是定義該變量應存儲的數據類型；
如果不指定數據類型或對象類型，也就是不用Dim定義，且在模塊中沒有 `Deftype` 語句，
則該變量按缺省設置是 `Variant` 類型。

- VBA中用Set賦值和不用Set賦值有什麼區別？

給普通變量賦值使用`Let`，Let 可以**省略**。</br>
給對象變量賦值使用`Set`，Set **不能** 省略。

```vba
Sub AssignString()
    Dim strA As String
    Dim strB As String

    strA = "hello"      ' 本句也可寫成 LET strA = "hello"
    Set strB = "hello"  ' 錯誤寫法/Compile error
EndSub
```

<a name="1.10"></a>
### 1.10 示例

舉個排序的例子，要對`A1:A20`的單元格區域進行排序，區域內的內容爲1-100的隨機整數，
規則是大於50的倒序排列，小於50的正序排列。將結果顯示在`B1:B20`的區域裏。

在這個例子中，首先定義一個`Sub`過程來隨機生成`A1:A20`區域的內容。
代碼如下:

![Alt text](/doc/source/images/demo1.1.gif)

```vba
' 創建隨機整數，並賦值
Sub createRandom(times As Integer)
    Dim num As Integer
    Dim arr() As Integer
    ReDim arr(times)

    For num = 1 To times
        Randomize (1) ' 初始化隨機數
        arr(num) = Rnd(1) * 10000 \ 100 ' Rnd隨機數函數生成0~1的浮點數
        ' 上面使用了運算符進行取整，也可以根據需求使用vba內部的取整函數達到同樣的效果
        ' arr(num) = Int(Rnd(1) * 100)
        ' arr(num) = Round(Rnd(1) * 100)
        Range("A" & num) = arr(num)
    Next num
End Sub

' 自定義排序
Function defSort(rgs) As Variant
    Dim arr() As Integer
    Dim total As Integer
    Dim rg
    Dim st As Integer  ' 數組開始標記
    Dim ed As Integer  ' 數組結束標記

    Debug.Print "rgs類型:"; TypeName(rgs)
    total = UBound(rgs)
    ReDim arr(total)
    st = 1
    ed = total

    ' 對數組分區
    For Each rg In rgs
        If rg > 50 Then
            arr(ed) = rg
            ed = ed - 1
        Else
            arr(st) = rg
            st = st + 1
        End If
    Next rg

    Dim i As Integer
    Dim j As Integer
    Dim tmp As Integer

    ' 冒泡排序
    For i = 1 To total
        For j = i To total
            If arr(i) > 50 And arr(j) > 50 Then '大於50的倒序排列
                If arr(i) < arr(j) Then
                    tmp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = tmp

                    Debug.Print "大於50的"; i; j; tmp ' 程序運行過程中在立即窗口顯示執行內容，用於調試程序
                End If
            Else If arr(i) <= 50 And arr(j) <= 50 Then ' 小於50的正序排列
                If arr(i) > arr(j) Then
                    tmp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = tmp

                    Debug.Print "不大於50的"; i; j; tmp
                End If
            Else
                Exit For
            End If
        Next j
    Next i
    defSort = arr
End Function


' 程序入口
Sub main()
    Const SORT_NUM = 20
    Dim rgs
    Dim arr

    createRandom SORT_NUM ' 初始化待排序區域

    rgs = range("A1:A" & SORT_NUM)
    arr = defSort(rgs)

    ' 循環賦值
    For i = 1 To SORT_NUM
        range("B" & i) = arr(i)
    Next i
End Sub
```


<a name="layout"></a>
## 0x02 VBA界面介紹

<a name="2.1"></a>
### 2.1 整體界面說明

（點擊圖片查看大圖）   
![Alt text](/doc/source/images/1505749555407.png)

<a name="2.2"></a>
### 2.2 工程資源管理器（Project Explore）說明

顯示快捷鍵：`Ctrl + R`，也可以點擊菜單欄 View -> <u>P</u>roject Explore 顯示。
在一個VBA項目中，實際可以在5個代碼模塊中書寫VBA代碼，如下圖所示：

![Alt text](/doc/source/images/vba_code_modules.png)

1. Code Modules – Code Modules是我們存儲宏的最常見的地方。
模塊位於工作簿中的 `Modules` 文件夾中。

2. Sheet Modules – 工作簿中的每個工作表在Microsoft Excel Objects文件夾中
都有一個工作表對象。雙擊sheet對象就會打開它的代碼模塊，我們可以在其中添加事件過程(宏)。
這些宏在用戶執行表單中的特定操作時運行。比如如下code：
如果在該sheet中的選擇位置發生改變，就會*自動執行* `Worksheet_SelectionChange` 方法，
選擇所選單元格的整個行和列。

```VBA
Private Sub Worksheet_SelectionChange(ByVal Target As Range) ' Worksheet_SelectionChange
    Application.EnableEvents = False

    With Target
        Union(.EntireRow, .EntireColumn).Select
        .Activate
    End With

    Application.EnableEvents = True
End Sub
```

3. ThisWorkbook Module – 每個工作簿都包含一個 `ThisWorkbook` 對象，
其總是位於和工作表對象相同的文件夾(Microsoft Excel Objects)內的最底部。
我們可以在這個工作簿中運行基於事件的宏。

4. Userforms – 做過VB項目的人對這個應該不會陌生。在這個模塊下我們可以創建Windows窗體，
進行圖形化交互。在這個模塊寫的code大部分都是和win窗體相關的代碼。

5. Class Modules – 在`Class Modules`文件夾中，允許我們編寫宏來創建對象、屬性和方法。
當我們想要創建對象庫中不存在的自定義對象或集合時，可以使用該類模塊。

**總結**：`Modules`、 `ThisWorkbook`、 `Sheet` 三者區別：

`Modules` 是相似功能和子程序的集合，通常根據功能進行分組。

`ThisWorkbook` 是Workbook對象的私有模塊。
例如，Workbook_Open()，Workbook_Close() 例程駐留在此模塊中。
（[工作簿對象參考](https://docs.microsoft.com/zh-cn/office/vba/api/excel.workbook)）

`Sheet1`，`Sheet2` 是單個工作表的私有模塊。在它們中，您將會放入該表的特定功能。
例如：`Worksheet_Activate` ， `Worksheet_Deactivate` ， `Workbook_SheetChange`
是提供給的默認事件，這樣你就可以在各自的私有工作表模塊中處理它們。
（[工作表對象參考](https://msdn.microsoft.com/en-us/library/office/ff847327.aspx)）

在模塊裏使用Cells、range等時表示的是當前激活的工作表；而在sheet裏面寫的話，
爲當前工作表裏的cells，如果你在sheet1代碼裏要引用其他工作表的話，不能這樣。

```vba
sheet2.select
cells(1, 1) = 1
```

因爲你的代碼在sheet1下，cells就一定是sheet1的
另外，在sheet下面可以使用Me，表示自身
如sheet1.visible = False，可以簡化爲: Me.visible = False

如果一個Funtion是在`Modules`裏定義的，那麼就可以在任意的Worksheet裏調用，
但如果只是在Worksheet裏定義的Funtion，其他的Worksheet是調用不了的。
也就是說，模塊（Modules）是公共的地方。


<a name="2.3"></a>
### 2.3 設置VBA Macro Project 密碼保護

#### 2.3.1 利用密碼保護工作表或者sheet

在VBA編輯界面依次點擊：<u>T</u>ools → VBAProject Prop<u>e</u>rties…

![Alt text](/doc/source/images/password_protect_setting_1.png)


在彈出界面選擇 `Projection`，勾選 `Lock project for viewing`後，輸入密碼，如下圖所示：

![Alt text](/doc/source/images/password_protect_setting_2.png)


#### 2.3.2 Macro執行時密碼保護

如果想要使用密碼控制Macro是否可以運行，可以參考如下代碼：
```
Dim password As Variant
password = Application.InputBox("Enter Password", "Password Protected")

Select Case password
    Case Is = False
        ' do nothing
    Case Is = "P@ssw0rd"  ' 驗證密碼
        Range("A1").Value = "This is secret code"   ' 執行密碼保護的代碼塊。
    Case Else
        MsgBox "Incorrect Password"
End Select
```


<a name="2.4"></a>
### 2.4 常用快捷欄及窗口設置
默認情況下某些常用的窗口VBA界面是不顯示的，比如立即窗口，編輯操作捷欄（批量註釋取消等）

#### 2.4.1 顯示編輯欄
鼠標右鍵點擊空白的快捷欄位置，勾選 `Edit` 選項會顯示出如下快捷欄

![Alt text](/doc/source/images/toolbars_edit_setting.png)

#### 2.4.2 顯示立即窗口(Immediate window)
Immediate window（立即窗口）：類似其他IDE的console控制檯。</br>
顯示快捷鍵：`Ctrl + G`，也可以點擊菜單欄 View -> <u>I</u>mmediate window 顯示。</br>
當在調試debug的時候，可以使用`Debug.Print "xxxlog"`的時候可以在該窗口直接顯示打印結果。

<a name="object-option"></a>
## 0x03 對象操作說明
Excel中的每個單元格，工作簿都是可以操作的對象；可以對對象進行復制、粘貼、刪除等，
也可操作對象的各種屬性，來控制其展示和行爲。

在Excel中，對象有不同的層級關係:

![Alt text](/doc/source/images/1505548045994.png)

實際上Excel中可操作的對象遠不止這些，具體的可以參考
[Excel 對象模型](https://msdn.microsoft.com/zh-cn/library/office/ff194068.aspx)

類似於數組，將各種類型的對象封裝到一塊可以組成集合。
一個集合中調用對象的例子：
![Alt text](/doc/source/images/1505548422147.png)


<a name="3.1"></a>
### 3.1 對象簡述

對象一般包含下面三種特性：

- 屬性

屬性表示對象的特徵，一般爲名詞。例如`Workbook.ActiveSheet`表示工作簿當前
處於激活狀態的工作表對象。

- 方法

方法表示對象可用的操作或可執行的動作。例如`Workbook.Activate`表示
激活工作簿的第一個工作表。

- 事件

事件表示對象可以被觸發的行爲，一般觸發後會執行對應的代碼。
例如`Workbook.Activate`表示工作簿中的工作表被激活了，然後執行對應的方法。

下面的代碼就是在`Workbook`被打開時，將工作簿最大化的例子。

```vba
Private Sub Workbook_Open()
    Application.WindowState = xlMaximized
End Sub
```

VBA中有很多對象，常用的對象如下:

|對象|對象說明| 文檔地址|
|----|----|----|
|Application|代表Excel應用程序|[文檔](https://msdn.microsoft.com/zh-cn/library/ff194565.aspx)|
|Workbook|代表Excel的工作簿|[文檔](https://msdn.microsoft.com/zh-cn/library/ff835568.aspx)|
|Worksheet|代表Excel的工作表|[文檔](https://msdn.microsoft.com/zh-cn/library/ff194464.aspx)|
|Range|代表Excel的單元格，可以是單個單元格或單元格區域|[文檔](https://msdn.microsoft.com/zh-cn/library/office/ff838238.aspx)|


<a name="3.2"></a>
### 3.2 Application對象
    參照Application對象[官方文檔](https://docs.microsoft.com/zh-CN/office/vba/api/Excel.Application(object))
### 3.3 Range對象
![Alt text](/doc/source/images/1505548886377.png)

![Alt text](/doc/source/images/1505549069568.png)

<a name="string-option"></a>
## 0x04 字符串String相關常用操作


<a name="4.1"></a>
### 4.1 Trim
`Trim`函數刪除給定輸入字符串的前導空格和尾隨空格。</br>
語法：Trim(String)

<a name="4.2"></a>
### 4.2 Instr 和 InStrRev
`InStr`函數返回一個字符串第一次出現在一個字符串，從左到右搜索。返回搜索到的字符索引位置。</br>
`InStrRev`函數與`InStr`功能相同，從**右**到左搜索。返回搜索到的字符索引位置。

語法：InStr([start, ]string1, string2[, compare])
參數：
   * Start   - 一個可選參數。指定搜索的起始位置。搜索從第一個位置開始，從左到右。
   * String1 - 必需的參數。要搜索的字符串。
   * String2 - 必需的參數。要在String1中搜索的字符串。
   * Compare - 一個可選參數。指定要使用的字符串比較。它可以採取以下提到的值：
       - 0 = vbBinaryCompare - 執行二進制比較(默認)
       - 1 = vbTextCompare - 執行文本比較

```vba
Private Sub Constant_demo_Click()
    Dim Var As Variant
    Var = "Microsoft VBScript"
    Debug.Print InStr(1, Var, "s")        ' 6
    Debug.Print InStr(7, Var, "s")        ' 0
    Debug.Print InStr(1, Var, "f", 1)     ' 8
    Debug.Print InStr(1, Var, "t", 0)     ' 9
    Debug.Print InStr(1, Var, "i")        ' 2
    Debug.Print InStr(7, Var, "i")        ' 16
    Debug.Print InStr(Var, "VB")          ' 11
End Sub
```

<a name="4.3"></a>
### 4.3 Mid
`Mid`函數返回給定輸入字符串中指定數量的字符。</br>
語法：Mid(String, start[, Length])</br>
參數：
   - String - 必需的參數。輸入從中返回指定數量的字符的字符串。
   - Start - 必需的參數。一個整數，它指定了字符串的起始位置。
   - Length - 必需的參數。一個整數，指定要返回的字符數。

```vba
    Private Sub Constant_demo_Click()
        Dim var as Variant
        var = "Microsoft VBScript"
        Debug.Print Mid(var, 2)       ' icrosoft VBScript
        Debug.Print Mid(var, 2, 5)    ' icros
        Debug.Print Mid(var, 5, 7)    ' osoft V
    End Sub
```

<a name="4.4"></a>
### 4.4 Left 和 Right
`Left` 和 `Right` 截取字符串，從左或者從右開始。</br>
語法：Left(String, Length)</br>
參數：
   - String - 必需的參數。 輸入從左側返回指定數量的字符的字符串。
   - Length - 必需的參數。 一個整數，指定要返回的字符數。
```vba
Private Sub Constant_demo_Click()
    Dim var as Variant

    var = "Microsoft VBScript"
    Debug.Print Left(var,2)     ' Mi

    var = "MS VBSCRIPT"
    Debug.Print Left(var,5)     ' MS VB

    var = "microsoft"
    Debug.Print Left(var,9)     ' microsoft
End Sub
```

<a name="4.5"></a>
### 4.5 Replace 函數
`Replace` 函數 將一個字符串替換另一個字符串，可指定的次數。</br>
語法：Replace(string, findString, replaceWith[, start[, count[, compare]]])</br>
參數：
   - String - 必需的參數。需要被搜索的字符串。
   - findString - 必需的參數。將被替換的字符串部分。
   - replaceWith - 必需的參數。用於替換的子字符串。
   - start - 可選的參數。規定開始位置。默認是 1。
   - count - 規定指定替換的次數。默認是 -1，表示進行所有可能的替換。
   - compare - 可選的參數。規定所使用的字符串比較類型。
       - 0 = vbBinaryCompare - 執行二進制比較(默認)
       - 1 = vbTextCompare - 執行文本比較

示例：</br>
```vba
dim txt
txt="This is a beautiful day!"
Debug.Print Replace(txt, "beautiful", "horrible")   ' This is a horrible day!
```

<a name="4.6"></a>
### 4.6 StrReverse 倒轉函數
語法：StrReverse(string) </br>
示例：</br>
```vba
Private Sub StrReverse_Demo()
    Debug.Print StrReverse("VBSCRIPT"))             ' TPIRCSBV
    Debug.Print StrReverse("My First VBScript"))    ' tpircSBV tsriF yM
    Debug.Print StrReverse("123.45"))               ' 54.321
End Sub
```

<a name="4.7"></a>
### 4.7 其他字符串函數
- `&` 字符串連接操作，在VBA中連個字符串連接使用`&`進行連接
- `Ltrim(string)` 去掉 string 左端空白
- `Rtrim(string)` 去掉 string 右端空白
- `Len(string)` 計算 string 長度
- `Lcase(string)` 和 `Ucase(string)` 轉換爲小寫和大寫


<a name="excel-option"></a>
## 0x05 Excel 相關常用操作

<a name="5.1"></a>
### 5.1 Excel 基礎操作

1. Range相關
Range 屬性的一些 A1 樣式引用
```vba
Range("A1")             ' 單元格 A1
Range("A1:B5")          ' 從單元格 A1 到單元格 B5 的區域
Range("C5:D9, G9:H16")  ' 多塊選定區域
' 選中不關聯的單元格，cells(2, 3)返回結果爲：B3
Union(Range("A1:A10"), Range("K10"), Range("A1:" & cells(2, 3).Address)).Select
Range("A:A")            ' A 列
Range("1:1")            ' 第一行
Range("A:C")            ' 從 A 列到 C 列的區域
Range("1:5")            ' 從第一行到第五行的區域
Range("1:1, 3:3, 8:8")  ' 第 1、3 和 8 行
Range("A:A, C:C, F:F")  ' A 、C 和 F 列
```

2. 行列相關
行和列的引用
```vba
Rows(1)         ' 第一行
Rows            ' 工作表上所有的行
Columns(1)      ' 第一列
Columns("A")    ' 第一列
Columns         ' 工作表上所有的列
Union(Rows(1), Rows(3), Rows(5))  ' 引用第1, 3, 5行
```
3. 循環Selction區域的每一個單元格Cell
```vba
For Each rngDataCell In RngDataSelection
    If Not rngDataCell.HasFormula And Not (Trim(rngDataCell.Value)  = "") Then
        ...
    End If
Next rngDataCell
```

4. 選擇當前工作表中的單元格
```vba
ActiveSheet.Cells(5, 4).Select
或：ActiveSheet.Range("D5").Select
```

5. 選擇同一工作簿中其它工作表上的單元格
```vba
Application.Goto (ActiveWorkbook.Sheets("Sheet2").Range("E6"))
' 也可以先激活該工作表，然後再選擇：
Sheets("Sheet2").Activate
ActiveSheet.Cells(6, 5).Select
```

6. 選擇與當前單元格相關的單元格/偏離當前單元格(Offset)</br>
語法：Offset(D, R) 以當前爲基礎原點，向下D，且向右D移動，如果負數即爲向反方向移動
即向上和向左移動。</br>
例如，要選擇距當前單元格下面5行左側4列的單元格
```vba
ActiveCell.Offset(5, -4).Select
```

7. 選擇一個指定的區域並擴展區域的大小
```vba
' 要選擇當前工作表中名爲“Database”區域，然後將該區域向下擴展5行，可以使用下面的代碼：
Range("Database").Select
Selection.Resize(Selection.Rows.Count + 5, Selection.Columns.Count).Select
```

8. 選擇一個指定的區域，再偏離，然後擴展區域的大小
```vba
' 選擇名爲“Database”區域下方4行右側3列的一個區域，然後擴展2行和1列，可以使用下面的代碼：
Range("Database").Select
Selection.Offset(4, 3).Resize(Selection.Rows.Count + 2, Selection.Columns.Count + 1).Select
```

9. 同時選擇兩個或多個指定區域</br>
**注意**：所選區域必須在同一工作表(sheet)中。
```vba
Set rngUnionSelection = Application.Union(Range("Sheet1!A1:B2"), Range("Sheet1!C3:D4"))
```

10. 選擇兩個或多個指定區域的交叉區域
**注意**：所選區域必須在同一工作表(sheet)中。
```vba
' 要選擇名爲“Test1”和“Test2”的兩個區域的交叉區域
Application.Intersect(Range("Test1"), Range("Test2")).Select
```

11. 利用End函數的相關操作

End(xldown)：從被選中的單元格向下尋找，如果被選中單元格爲空，則一直向下走到
第一個非空單元格；如果被選中單元格爲非空，則向下走到最後一個非空單元格。</br>
`End`函數的4個方向參數：xlUp, xlDown, xlToLeft, xlToRight。

```vba
' 選擇連續數據列中的最後一個單元格
ActiveSheet.Range("a1").End(xlDown).Select
' 選擇連續數據列底部的空單元格
ActiveSheet.Range("a1").End(xlDown).Offset(1, 0).Select
' 獲取連續數據最後一行的行號
Selection.end(xldown).Row
' 想選擇連續數據最後面的空白行
Rows(Selection.End(xldown).Row + 1).Select
' 選擇某列中連續數據單元格區域
ActiveSheet.Range("A1", ActiveSheet.Range("a1").End(xlDown)).Select
ActiveSheet.Range("A1:" & ActiveSheet.Range("a1").End(xlDown).Address).Select
' 選擇某列中非連續數據單元格區域
ActiveSheet.Range("A1", ActiveSheet.Range("a65536").End(xlUp)).Select
ActiveSheet.Range("A1:" & ActiveSheet.Range("a65536").End(xlUp).Address).Select
```

補充： 對於上述代碼中非連續數據，也可以利用UsedRange.Rows.Count獲取所有數據的條/行數。
```vba
Dim lngCountData As Long
lngCountData = ActiveSheet.UsedRange.Rows.Count
```


<a name="5.2"></a>
### 5.2 打開Excel兩種方式

- 利用 `GetObject` 方法打開Excel文檔
```vba
    Sub GetWorkbook()
        Dim wbWorkFile As Workbook
        Set wbWorkFile = GetObject("D:\test.xlsx")
        ' wbWorkFile.Windows(1).Visible = True ' 這種方法打開的文件是隱藏的，如果需要顯示，則設置Visible值爲ture
        wbWorkFile.Close False
        Set wbWorkFile = Nothing
    End Sub
```

- 利用 `Open` 方法打開Excel文檔
```vba
Sub OpenWorkbook()
    Dim wbWorkFile As Workbook
    Set wbWorkFile = Workbooks.Open("D:\test.xlsx")
    wbWorkFile.Windows(1).Visible = False
    wbWorkFile.Close False
    Set wbWorkFile = Nothing
End Sub
```

延伸其擴展方法：
- GetObject封裝方法，可以作爲共通Function

```vba
Sub GetWorkbook()
    Dim objExcel                As Object       ' 用於存放Microsoft Excel 引用的變量。
    Dim blnExcelWasNotRunning   As Boolean      ' 用於最後釋放的標記。

    ' 測試 Microsoft Excel 的副本是否在運行。
    On Error Resume Next                        ' 延遲錯誤捕獲。
    ' 不帶第一個參數調用 Getobject 函數將返回對該應用程序的實例的引用。如果該應用程序不在運行，則會產生錯誤。
    Set objExcel = Getobject(, "Excel.Application")
    If Err.Number <> 0 Then blnExcelWasNotRunning = True
    Err.Clear                                   ' 如果發生錯誤則要清除 Err 對象。

    Set objExcel = Getobject("C:\excel.xlsx")   ' 將對象變量設爲對要看的文件的引用。

    ' 設置其 Application 屬性，顯示 Microsoft Excel。然後使用 objExcel 對象引用的 Windows 集合顯示包含該文件的實際窗口。
    objExcel.Application.Visible = True
    objExcel.Parent.Windows(1).Visible = True
    ' 在此處對文件進行操作。
    ' ...
    ' 如果在啓動時，Microsoft Excel 的這份副本不在運行中，則使用 Application 屬性的 Quit 方法來關閉它。
    ' 注意，當試圖退出 Microsoft Excel 時，標題欄會閃爍，並顯示一條消息詢問是否保存所加載的文件。
    If blnExcelWasNotRunning = True Then
        objExcel.Application.Quit
    End IF

    Set objExcel = Nothing   ' 釋放對該應用程序

End Sub
```

- OpenWorkbook封裝方法，可以作爲共通Function

```vba
Function OpenWorkbook(ByVal strWorkbookFilePath As String)
    Dim wb As Workbook
    Dim fileName As String
    fileName = Dir(strWorkbookFilePath)

    On Error Resume Next
    Set wb = Workbooks(fileName)
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Workbooks.Open(strWorkbookFilePath)
    End If

    Set OpenWorkbook = wb

End Function
```

<a name="5.3"></a>
### 5.3 操作Excel工作表（Worksheet）

#### 5.3.1 移動工作表

移動工作表是指將工作表移到工作簿中的其他位置。
在VBA中，可以使用WorkSheet.Move方法來移動工作表。

語法：表達式.Move(Before, After)
其中，在Move方法中，主要包含兩個參數，其功能如下：

Before 在其之前放置移動工作表的工作表。如果指定了After，則不能指定Before。
After 在其之後放置移動工作表的工作表。如果指定了Before，則不能指定After。
例如：移動 "工資表" 至Sheet3工作表之後，可以輸入以下代碼：

```vba
Sub 移動工作表()
    Sheets("工資表").Select
    Sheets("工資表").Move After:=Sheets(3)
End Sub
```

另外，如果既不指定Before也不指定After，Microsoft Excel將新建一個工作簿，
其中包含所移動的工作表。例如，輸入以下代碼，即可新建一個工作簿，
且該工作表中包含有 "工資表" 工作表。

```vba
Sub A()
    Sheets("工資表").Move
End Sub
```

#### 5.3.2 複製工作表

複製工作表是指將工作表進行備份，以便於用戶對備份文件進行操作時，不會損壞原有文件。
在VBA中，使用Sheets.Copy方法可以將工作表複製到工作簿的另一位置。
語法：

```vba
表達式.Copy(Before, After)
```

其中，在Copy方法中，包含的兩個參數與在Move方法中的參數相似，其參數功能如下：
Before 將要在其之前放置所複製工作表的工作表。如果指定了After，則不能指定Before。
After 將要在其之後放置所複製工作表的工作表。如果指定了Before，則不能指定After。
例如：複製 "工資表" 表格至Sheet3工作表之後，可以輸入以下代碼：

```vba
Sub 複製工作表()
    Sheets("工資表").Select
    Sheets("工資表").Copy After:=Sheets(3)
End Sub
```

另外，用戶還可以在不同的工作簿之間進行復制。
例如：將當前工作簿中的“工資表”工作表複製到打開的Book1工作表中，可以輸入以下代碼：

```vba
Sub 複製工作表至Book1中()
    Sheets("工資表").Copy After:=Workbooks("Book1").Sheets(1)
End Sub
```

<a name="5.4"></a>
### 5.4 Excel AutoFilter / Excel 自動篩選操作

#### 5.4.1 顯示所有數據記錄
```vba
Sub ShowAllRecords()
    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
End Sub
```

#### 5.4.2 開關Excel自動篩選

先判斷是否有自動篩選，如果沒有爲A1添加一個自動篩選
```vba
Sub TurnAutoFilterOn()
    'check for filter, turn on if none exists
    If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("A1").AutoFilter
    End If
End Sub
```

清除自動篩選
```vba
Sub TurnFilterOff()
    'removes AutoFilter if one exists
    Worksheets("Data").AutoFilterMode = False
End Sub
```

#### 5.4.3 隱藏過濾箭頭

隱藏所有的箭頭
```vba
Sub HideALLArrows()
    'hides all arrows in heading row
    'the Filter remains ON
    Dim c As Range
    Dim i As Integer
    Dim rng As Range
    Set rng = ActiveSheet.AutoFilter.Range.Rows(1)
    i = 1
    Application.ScreenUpdating = False

    For Each c In rng.Cells
        c.AutoFilter Field:=i, _
            Visibledropdown:=False
        i = i + 1
    Next

    Application.ScreenUpdating = True
End Sub
```

只保留一個箭頭，其他的過濾箭頭全隱藏

![Alt text](doc/source/images/autofiltermacros01.png)

```vba
Sub HideArrowsExceptOne()
'hides all arrows except
' in specified field number
Dim c As Range
Dim rng As Range
Dim i As Long
Dim iShow As Long
Set rng = ActiveSheet.AutoFilter.Range.Rows(1)
i = 1
iShow = 2 'leave this field's arrow visible
Application.ScreenUpdating = False

For Each c In rng.Cells
    If i = iShow Then
        c.AutoFilter Field:=i, _
        Visibledropdown:=True
    Else
        c.AutoFilter Field:=i, _
        Visibledropdown:=False
    End If
    i = i + 1
Next

Application.ScreenUpdating = True
End Sub
```

隱藏部分箭頭

![Alt text](doc/source/images/autofiltermacros02.png)

```vba
Sub HideArrowsSpecificFields()
    'hides arrows in specified fields
    Dim c As Range
    Dim i As Integer
    Dim rng As Range
    Set rng = ActiveSheet.AutoFilter.Range.Rows(1)
    i = 1
    Application.ScreenUpdating = False

    For Each c In rng.Cells
        Select Case i
            Case 1, 3, 4
            c.AutoFilter Field:=i, _
                Visibledropdown:=False
        Case Else
            c.AutoFilter Field:=i, _
                Visibledropdown:=True
        End Select
        i = i + 1
    Next

    Application.ScreenUpdating = True
End Sub
```

#### 5.4.4 複製所有的過濾後的數據

```vba
Sub CopyFilter()
    'by Tom Ogilvy
    Dim rng As Range
    Dim rng2 As Range

    With ActiveSheet.AutoFilter.Range
        On Error Resume Next
            Set rng2 = .Offset(1, 0).Resize(.Rows.Count - 1, 1) _
            .SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
    End With
    If rng2 Is Nothing Then
        MsgBox "No data to copy"
    Else
        Worksheets("Sheet2").Cells.Clear
        Set rng = ActiveSheet.AutoFilter.Range
        rng.Offset(1, 0).Resize(rng.Rows.Count - 1).Copy _
        Destination:=Worksheets("Sheet2").Range("A1")
    End If

    ActiveSheet.ShowAllData
End Sub
```

#### 5.4.5 檢查是否有自動篩選：

可以打開立即窗口，即類似於控制檯的 Immediate Window，快捷鍵：`Ctrl+G` ,查看如下code的
iARM的打印值。

![Alt text](doc/source/images/autofiltermacros03.png)

```vba
Sub CountSheetAutoFilters()
    Dim iARM As Long
    'counts all worksheet autofilters
    'even if all arrows are hidden
    If ActiveSheet.AutoFilterMode = True Then iARM = 1
    Debug.Print "AutoFilterMode: " & iARM
End Sub  
```


<a name="5.5"></a>
### 5.5 清理Excel數據相關操作

#### 5.5.1 清理單元格或Range中的內容

如若清空某個選中的單元格中的數據，使用的API爲：`ClearContents`。   
示例：
```vba
Range("A1").Select
Selection.ClearContents
```

#### 5.5.1 清理/刪除Excel中第一個標題行以外的所有行

同樣使用ClearContents方法，主要是確定如何選中除第一行以外的表格。   
示例代碼如下：
```vba
Sub ClearContentExceptFirst()
    Rows("2:" & Rows.Count).ClearContents
End Sub
```

<a name="0x06"></a>
## 0x06 文件，文件夾等 相關常用操作

以下文件，文件夾等相關方法可自行封裝成共通(common function)以便項目中使用。

<a name="6.1"></a>
### 6.1 判斷文件，文件夾等是否存在
1. 文件是否存在（File exists）：
```vba
Sub FileExists()
    Dim fso as Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists("D:\test.txt") = True Then
        MsgBox "The file is exists."
    Else
        MsgBox "The file isn't exists."
    End If
End Sub
```

2. 文件夾是否存在（Folder exists）：
```vba
Sub FolderExists()
    Dim fso as Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists("D:\testFolder") = True Then
        MsgBox "The folder is exists."
    Else
        MsgBox "The folder isn't exists."
    End If
End Sub
```

3. 硬盤是否存在（Drive exists）：
```vba
Sub DriveExists()
    Dim fso as Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.DriveExists("D:\") = True Then
        MsgBox "The drive is exists."
    Else
        MsgBox "The drive isn't exists."
    End If
End Sub
```

<a name="6.2"></a>
### 6.2 文件相關操作
1. 文件複製（File copy）：
```vba
Sub CopyFile()
    Dim fso as Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile "c:\Makro.txt", "c:\Macros\"
End Sub
```

2. 文件移動（File move）：
```vba
Sub MoveFile()
    Dim fso as Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFile "c:\*.txt", "c:\Documents and Settings\"
End Sub
```

3. 文件刪除（File delete）：
```vba
    Sub DeleteFile()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFile "c:\Documents and Settings\Macros\Makro.txt"
End Sub
```

<a name="6.3"></a>
### 6.3 文件夾相關操作

1. 創建文件夾（Folder create）：
```vba
Sub CreateFolder()
    Dim fso as Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder "c:\Documents and Settings\NewFolder"
End Sub
```

2. 複製文件夾（Folder copy）：
```vba
Sub CopyFolder()
    Dim fso as Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFolder "C:\Documents and Settings\NewFolder", "C:\"
End Sub
```

3. 移動文件夾（Folder move）：
```vba
Sub MoveFolder()
    Dim fso as Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFolder "C:\Documents and Settings\NewFolder", "C:\"
End Sub
```

4. 刪除文件件（Folder delete）：
```vba
Sub DeleteFolder()
    Dim fso as Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder "C:\Documents and Settings\NewFolder"
End Sub
```

<a name="6.4"></a>
### 6.4 其他操作（獲取文件名等）

1. 獲取文件全名，帶有後綴（Get file name）
```vba
Sub GetFileName()
    Dim fso as Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    MsgBox fso.GetFileName("c:\Documents and Settings\Makro.txt")   ' Makro.txt
End Sub
```

2. 獲取文件名，無後綴（Get base name）
```vba
Sub GetBaseName()
    Dim fso as Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    MsgBox fso.GetBaseName("c:\Documents and Settings\Makro.txt")   ' Makro
End Sub
```

3. 獲取文件後綴格式（Get extension name）
```vba
Sub GetExtensionName()
    Dim fso as Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    MsgBox fso.GetExtensionName("c:\Documents and Settings\Makro.txt")  ' txt
End Sub
```

4. 獲取盤符名（Get drive name）
```vba
Sub GetDriveName()
    Dim fso as Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    MsgBox fso.GetDriveName("c:\Documents and Settings\Makro.txt")  ' c:
End Sub
```


<a name="0x07"></a>
## 0x07 日期和時間 相關函數


<a name="7.1"></a>
### 7.1 Date, Time, Now 函數

`Date` 函數返回當前的系統日期。   
`Time` 函數返回當前的系統時間。   
`Now`  函數返回當前的系統日期和時間。   

**注意：** 如果同時讀取 Date、Time 以及 Now，那麼 Now = Date + Time，但是實際上，我們不可能同時調用這三個函數，因爲執行完一個函數之後，才能執行另一個函數，所以如果您在程序中必需同時取得當時的日期和時間，必需調用 Now，再利用 DateVale 及 TimeValue 分別取出日期和時間。


```
Private Sub CommandButton2_Click()    
    MsgBox Now    ' 20/05/2020 16:28:04
    MsgBox Date   ' 20/05/2020
    MsgBox Time   ' 16:28:39
End Sub
```

<a name="7.2"></a>
### 7.2 日期函數：Year, Month, Day

`Year`, `Month`, `Day` 函數分別返回 **數字格式** 的年，月，日。

```
Dim exampleDate As Date

exampleDate = DateValue("May 19, 2020")

MsgBox Year(exampleDate)  ' 2020
MsgBox Month(exampleDate) ' 5
MsgBox Day(exampleDate)   ' 19
```


<a name="7.3"></a>
### 7.3 CDate 和 DateValue 函數
VBA中的CDate和DateValue的區別(Difference between CDate and DateValue in VBA)

1. `CDate` 函數可把一個合法的日期和時間 *表達式* 轉換爲 `Date` 類型，並返回結果。

**提示：** 請使用 [IsDate](#7.4) 函數來判斷 date 是否可被轉換爲日期或時間。

※　舉例參照如下小節

2. DateValue(date) 函數 返回一Date類型數據

*date* 參數通常是一個字符串表達式, 表示從100年1月1日到9999年12月31日之間的日期。 但是，*date* 還可是任何表示該範圍內的日期、時間或日期和時間的表達式。

**備註（Reference From [MSDN](https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/datevalue-function)）：** 如果 *date* 是一個僅包含由有效日期分隔符分隔的數字的字符串, 則DateValue將根據您爲系統指定的短日期格式識別月、日和年的順序。 DateValue 還能清楚地識別包含月名稱（長名稱或簡寫形式）的日期。 例如，除了識別 12/30/1991 和 12/30/91 之外，DateValue 還識別 December 30, 1991 和 Dec 30, 1991。
如果省略 date 的年部分，則 DateValue 將使用計算機系統日期中的當前年。
如果 date 參數包含時間信息，則 DateValue 將不會返回它。 但是，如果 date 包含的時間信息無效（如“89:98”），則將出錯。


CDate與DateValue的舉例：   
```
' 舉例1：
d1 = "April 22, 2001"
d2 = "6:15:45 PM"
d3 = "2014-07-24 15:43:06"

If IsDate(d1) And IsDate(d2) And IsDate(d3) Then
    MsgBox CDate(d1)      ' 22/04/2020
    MsgBox CDate(d2)      ' 18:15:45
    MsgBox CDate(d3)      ' 24/07/2014 15:43:06

    MsgBox DateValue(d1)  ' 22/04/2020
    MsgBox DateValue(d2)  ' 00:00:00
    MsgBox DateValue(d3)  ' 24/07/2014
End If
```

```
' 舉例2：
MsgBox CDate(43972)             ' 5/21/2020
MsgBox CDate("12/25/2020")      ' 1/15/2014

MsgBox DateValue("12/25/2020")  ' 1/15/2014
MsgBox DateValue(43972)         ' Throws a Type mismatch error(Run-time error 13)
```


**總結：** CDate和DateValue的區別：   
* 從上述 *舉例1* 代碼結果總結出，`DateValue` 只返回一個Date類型結果，而`CDate`返回結果將保留日期和時間，當參數爲一個時間類型的時候，`DateValue` 只能返回一個 `00:00:00`的結果。

* 從上述 *舉例2* 代碼結果總結出，另外，`DateValue` (或`TimeValue`)只接受 `String` 類型的參數，而 `CDate` 也可以處理 **數字** 。可使用`CDate(Int(num))`和`CDate(num - Int(num))`函數。


<a name="7.4"></a>
### 7.4 IsDate 函數
`IsDate` 函數返回一個布爾值，用於判斷一個表達式是否可被轉換爲日期。如果表達式是日期，或可被轉換爲日期，則返回 True 。否則，返回 False 。

**註釋：** `IsDate` 函數使用本地設置來檢測字符串是否可以轉換爲日期。在 Windows 中, 有效日期的範圍是公元100年1月1日至公元9999年12月31日;各操作系統的範圍各不相同。

示例：   
```
Dim MyVar, MyCheck
MyVar = "04/28/2014"        ' Assign valid date value.
MyCheck = IsDate(MyVar)     ' Returns True.

MyVar = "April 28, 2014"    ' Assign valid date value.
MyCheck = IsDate(MyVar)     ' Returns True.

MyVar = "13/32/2014"        ' Assign invalid date value.
MyCheck = IsDate(MyVar)     ' Returns False.

MyVar = "04.28.14"          ' Assign valid time value.
MyCheck = IsDate(MyVar)     ' Returns True.

MyVar = "04.28.2014"        ' Assign invalid time value.
MyCheck = IsDate(MyVar)     ' Returns False.

MyVar = "Hello World!"      ' Assign invalid time value.
MyCheck = IsDate(MyVar)     ' Returns False.
```


<a name="7.5"></a>
### 7.5 DateAdd 函數

`DateAdd` 函數可返回已添加指定時間間隔的日期。

語法：DateAdd(*interval, number, date*)

| 參數 | 描述 |
|--------|-------|
|interval|必需的。需要增加的時間間隔。</br>可採用下面的值：</br>yyyy - 年   </br>q - 季度          </br>m - 月          </br>y - 當年的第幾天          </br>d - 日          </br>w - 當週的第幾天          </br>ww - 周          </br>h - 小時          </br>n - 分鐘          </br>s - 秒|
|number|必需的。需要添加的時間間隔的數目。可對未來的日期使用正值，對過去的日期使用負值。|
|date|必需的。代表被添加的時間間隔的日期的變量或文字。|

示例：

```
DateAdd("m",1,"31-Jan-01")    ' 2/28/2001
DateAdd("m",1,"31-Jan-00")    ' 2/29/2000
DateAdd("m",-1,"31-Jan-01")   ' 12/31/2000
```

在此情況下，DateAdd 返回 2001 年 2 月 28 日，而不是 2001 年 2 月 31 日。 如果 date 爲 2000 年 1 月 31 日，則它將返回 2000 年 2 月 29 日，因爲 2000 年是閏年。

如果計算的日期位於年份數 100 前（即，你減去的年份數大於 date 中的年份，則出現錯誤。

如果 *number* 不是 Long 值，則在計算之前將其 **四捨五入** 到最接近的整數。

**另外：** 在使用“w”時間間隔（包括一週的所有天，從星期日到星期六）向日期添加天數時，DateAdd 函數會向日期添加您指定的總天數，而不是像您預期的那樣僅向日期添加工作日（從星期一到星期五）數。


<a name="0x10"></a>
## 0x10 VBA 轉換函數一覽

* [Type conversion functions](type-conversion-functions.md)
- [10.1 CBool](type-conversion-functions.md#CBool-function-example)
- [10.2 CByte](type-conversion-functions.md#CByte-function-example)
- [10.3 CCur](type-conversion-functions.md#CCur-function-example)
- [10.4 CDate](type-conversion-functions.md#CDate-function-example)
- [10.5 CDbl](type-conversion-functions.md#CDbl-function-example)
- [10.6 CDec](type-conversion-functions.md#CDec-function-example)
- [10.7 CInt](type-conversion-functions.md#CInt-function-example)
- [10.8 CLng](type-conversion-functions.md#CLng-function-example)
- [10.9 CLngLng](type-conversion-functions.md) (Valid on 64-bit platforms only.)
- [10.10 CLngPtr](type-conversion-functions.md)
- [10.11 CSng](type-conversion-functions.md#CSng-function-example)
- [10.12 CStr](type-conversion-functions.md#CStr-function-example)
- [10.13 CVar](type-conversion-functions.md#CVar-function-example)


<a name="0x90"></a>
## 0x90 VBA Best Practices
1. Always have Option Explicit at the top of your code modules to
enforce variable declaration.
2. Never write procedures and functions that are longer than a full screen
as these are hard to understand. Procedures should fit on one screen -
ie be 40-50 lines long maximum.- ie be 40-50 lines long maximum.
3. Always prefix your variables so you can quickly identify their datatype.
4. Never use the Variant datatype unless absolutely necessary.</br>
**注**：儘量不要使用`Variant`，要顯示的聲明具體的數據類型。Variant是VBA中的一種特殊類型，
所有沒有聲明的數據類型的變量都默認是Variant型。但Variant型所佔的存儲空間遠大於其他的
數據類型，所以除非必要，否則應該避免申明變量爲Variant型。
5. Always use the keyword "**Call**" to call your procedures.
6. Always put your arguments in parentheses.
7. Never use Global variables unless absolutely necessary.
Pass parameters ByVal (ByRef is the default) - only use ByRef where
you intend to modify the parameter and pass the change back to the caller.
8. Always use tabs to indent your code to bring structure, never use spaces.
9. Add "value added" comments which explain why, do not add trivial comments.
10. Always add an Error Handler to every procedure and function.
11. Use the line continuation character to make your code more readable and
to reduce the amount of scrolling.
12. Never use the Option Base or Option Compare statements.

**More reference:** [VBA Code Guidelines/Best-practices](CodingStandards.md)

<a name="0x08"></a>
## 0x91 Trouble shooting


<a name="91.1"></a>
### 91.1 調試經驗 Excel點擊保存時總是彈出隱私信息警告（Privacy Warning:this document contains macros...）的解決方法

警告信息：
> Privacy Warning:this document contains macros, ActiveX controls, XML expansion pack information or web components. these may include personal information that cannot be removed by the document Inspector.

菜單依次點擊：</br>
File → Options → Trust Center → Trust Center Settings → Privacy Options
取消勾選(Uncheck) "Remove personal information from file properties on save" 選項

![Alt text](doc/source/images/trouble_shooting_01.png)


<a name="91.2"></a>
### 91.2 清除Excel數據透視表中過濾器緩存（舊項目）

如下圖所示，根據數據範圍創建數據透視表時，從源範圍中刪除數據後，即使刷新數據透視表，舊項目仍將存在於數據透視表的下拉菜單中。 如果要從數據透視表的下拉菜單中刪除所有舊項目，可參照如下兩種方法：


**1. 通過更改選項來清除數據透視表中的過濾器緩存（舊項目）**

Step1: 右鍵單擊數據透視表中的任何單元格，然後單擊 數據透視表選項 從上下文菜單。 看截圖：   

![Alt text](doc/source/images/doc-clear-filter-cache-1.png)

Step2: 在裏面 數據透視表選項 對話框中，單擊 **數據** 標籤，選擇 沒有 來自 **每個字段要保留的項目數量** 下拉列表，然後單擊 OK 按鈕。   

![Alt text](doc/source/images/doc-clear-filter-cache-2.png)

Step3: 右鍵單擊“數據透視表”單元格，然後單擊 **刷新** 從右鍵菜單。 看截圖：   

![Alt text](doc/source/images/doc-clear-filter-cache-3.png)

然後你可以看到舊的項目從數據透視表的下拉菜單中刪除，如下圖所示。   

![Alt text](doc/source/images/doc-clear-filter-cache-4.png)


**2. 使用VBA代碼清除所有數據透視表中的過濾器緩存（舊項目）**

在 **項目** 窗格打開 **ThisWorkbook（Code）** 窗口，然後將下面的VBA代碼複製並粘貼到窗口中。
```VBA
Private Sub Workbook_Open()
    Dim xPt As PivotTable
    Dim xWs As Worksheet
    Dim xPc As PivotCache
    Application.ScreenUpdating = False
    For Each xWs In ActiveWorkbook.Worksheets
        For Each xPt In xWs.PivotTables
            xPt.PivotCache.MissingItemsLimit = xlMissingItemsNone
        Next xPt
    Next xWs
    For Each xPc In ActiveWorkbook.PivotCaches
        On Error Resume Next
        xPc.Refresh
    Next xPc
    Application.ScreenUpdating = True
End Sub
```
![Alt text](doc/source/images/doc-clear-filter-cache-5.png)

按 F5 鍵來運行代碼，然後從活動工作簿中的所有數據透視表的下拉菜單中立即刪除舊項目。


<a name="8.3"></a>
### 91.3 解決辦法：The macros in this project are disabled.  Please refer to the online help or documentation of the host application to determine how to enable macros.

錯誤現象： Excel2016（365）運行macro宏時，彈出標題警告↓   

![Alt text](doc/source/images/trouble_shooting_03_01.png)   

解決辦法：   
step1：先確認Excel的設置是否正確
依次點擊 File >> Options >> Trust Center >> Trust Center Settings >> Macro Settings 按照如下圖所示設置：   

![Alt text](doc/source/images/trouble_shooting_03_02.png)   

step2：確認自己機器的安全級別   
一般如果按照step1設置後再次運行macro宏依然彈出警告，另一種情況就是你自己的機器（遠程PC/Server/VDI等）自身的安全級別過高造成的。   
打開瀏覽器依次點擊 Tools >> Internet options >> Security >> Customer level...  將安全級別從 **高(High)** 改成 **中(Medium)**。   

![Alt text](doc/source/images/trouble_shooting_03_03.png)   


<a name="91.7"></a>
### 91.7 Excel每次保存時都彈出警告：”此文檔中包含宏、Activex控件、XML擴展包信息“（office 2007/2010/365+）

**1.** office 2003版本：

 依次點擊：“工具” → “選項" → "在安全" 選項卡中 **勾選** ”保存時從文件屬性中刪除個人信息”。

![Alt text](doc/source/images/trouble-shootings/remove_activeX_warning_1.png)


**2.** office 2007/2010/365+版本：

單擊“Office按鈕（或文件菜單） → Excel選項（或選項） → 信任中心”，單擊“信任中心設置”按鈕，選擇“個人信息選項”（隱私選項Privacy Options），在“文檔特定設置”下 **取消** 勾選 “保存時從文件屬性中刪除個人信息” 後確定。

**注意：** 該選項僅對當前工作簿有效。另外，新建工作簿時該選項爲灰色不可用，只有用“文檔檢查器”檢查了文檔並刪除了個人信息後該選項纔可用。

![Alt text](doc/source/images/trouble-shootings/remove_activeX_warning_2.png)


<a name="91.8"></a>
### 91.8 解決辦法：使用.xlam宏文件執行VBA程序時，操作excel無任何反應

今天小夥伴遇到一個問題，自己做了一個xlam的工具，實現更改當前excel的所有sheet名稱，但是當點擊工具button的時候程序無反應。   
代碼如下：

```
Sub demo45()
    Dim ws As Worksheet          ' 把ws 定義爲一個工作表對象
    For Each ws In Worksheets    ' 用for each 遍歷對象集合
    ws.Name = "Test_" & ws.Name  ' 改名
    Next
End Sub
```

原因及解決辦法：   
首先看代碼並沒有什麼問題，調查的時候在循環語句中添加了ThisWorkbook指定，
但執行macro的時候發現sheet名字依然沒有改變。在無意間看 工程資源管理器（Project Explore）
的時候，發現VBAProject(XXX.xlam)中的默認Sheet1的名字被改變，這也就說明了，不是VBA程序
沒有起作用，而是程序在執行的時候默認操作了xlam工作簿。後修改程序中的默認制定工作簿語句
爲：`ActiveWorkbook` marco可以正常執行操作，問題得以解決。


<a name="0x92"></a>
## 0x92 VBA示例代碼
VBA示例代碼查看：[點擊這裏](SampleCode.bas)。


<a name="docslist"></a>
## 0xFF VBA學習資源列表
- [MSDN 函數 (Visual Basic for Applications)](https://docs.microsoft.com/zh-cn/office/vba/language/reference/functions-visual-basic-for-applications)
- [Excel-vba coding規約/開發規範](https://github.com/Youchien/development-specification/blob/master/doc/source/Excel-vba%20Language%20Specification.md)
- [Excel VBA 參考,官方文檔,適用2013及以上](https://msdn.microsoft.com/zh-cn/library/ee861528.aspx)
- [Excel宏教程 (宏的介紹與基本使用)](http://blog.csdn.net/lyhdream/article/details/9060801)
- [Excel2010中的VBA入門,官方文檔](https://docs.microsoft.com/zh-cn/previous-versions/office/ee814737(v=office.14))
- [Excel VBA的一些書籍資源,百度網盤](https://pan.baidu.com/s/1ktVmW63s8utBpAdcGnJfJA)  （提取碼: `j92n`）
- [Excel 函數速查手冊](https://support.office.com/zh-cn/article/Excel-%E5%87%BD%E6%95%B0%EF%BC%88%E6%8C%89%E7%B1%BB%E5%88%AB%E5%88%97%E5%87%BA%EF%BC%89-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb?ui=zh-CN&rs=zh-CN&ad=CN)
- [VBA的一些使用心得](http://www.cnblogs.com/techyc/p/3355054.html)
- [VBA函數參考](https://msdn.microsoft.com/zh-cn/library/office/jj692811.aspx)
- [VBA入門參考，英文](http://analystcave.com/vba-cheat-sheet/)


<a name="license"></a>
## 開源許可
本Repository除特殊註明外，均採用 Creative Commons [BY-NC-ND 4.0](LICENSE)（自由轉載-保持署名-非商用-禁止演繹）協議發佈。


## 鳴謝列表
### Code Contributors
| <img src="https://avatars0.githubusercontent.com/u/25427352" alt="bluetata" width="100px" height="100px"/> |<img src="https://avatars2.githubusercontent.com/u/46813661" alt="chromeheart" width="100px" height="100px"/> |<img src="https://avatars3.githubusercontent.com/u/3829140" alt="BobBJSun" width="100px" height="100px"/> | | | | |
| :----: |:----: |:----: |:----: |:----: |:----: |:----: |
| [bluetata](https://github.com/bluetata) |[chromeheart](https://github.com/chromeheart) |[BobBJSun](https://github.com/BobBJSun) | | | | |
