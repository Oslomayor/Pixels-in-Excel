# Pixels-in-Excel
Show a Photo in Microsoft Excel pixel-by-pixel  
在 Excel 中显示照片

## 代码文件

包含两个文件

* photo_excel.m

  主程序

* num2col_recur.m

  坐标转换函数。Excel中的列坐标以A-Z的字母表示，主程序调用该函数，将循环产生的数字列转化为字母列

## 效果

通过像素点遍历的方法，设置Excel不同单元格的颜色，从而在Excel中而拼成图片
(正常运行需要保证Excel已激活，否则MATLAB程序会中断，必须手动调试绕过office软件的激活提示)

![](https://github.com/Oslomayor/Markdown-Imglib/blob/master/Imgs/photo_excel.png?raw=true)

