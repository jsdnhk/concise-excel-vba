###  0x94 Excel-VBA Debug調試相關操作

在工作窗口，上方菜單欄中，有一個專門的額菜單：Debug 菜單，裏面有debug相關操作。

除此之外你也需要一些輔助窗口來幫助你更好的進行調試，

**1. Immediate window（立即窗口）：**

類似其他IDE的console控制檯。</br>
顯示快捷鍵：`Ctrl + G`，也可以點擊菜單欄 View -> <u>I</u>mmediate window 顯示。</br>
當在調試debug的時候，可以使用`Debug.Print "xxxlog"`的時候可以在該窗口直接顯示打印結果。

**2. Watches window（監視窗口）：**

右鍵點擊所要監視的變量，點擊 `Add Watch…` 點擊OK 會出現監視窗口，方便我們監視變量值。


**3. Locals window（本地窗口）：**

點擊菜單來View →　Locals Window，會顯示本地窗口，顯示所有參數變量。

view：
![Alt text](./doc/source/images/debug/debug.jpg)  
