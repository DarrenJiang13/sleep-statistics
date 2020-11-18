# sleep-statistics
给医学院同学做的批量处理rtf文件脚本，最终结果导出成excel。  
处理逻辑为将RTF格式转化为docx，然后用python读取对应位置的数据并存储在excel中。

请依照下列简单步骤来运行脚本文件"script.py"（默认已安装anaconda）：
1. 拖动待处理的文件夹到"script.py"同一目录下，
2. 按住“shift”键，单击鼠标打开Powershell
3. 输入`python script.py`敲击回车运行
4. 在当前目录下查看结果文件"result.xlsx"

注意:
1. 所有的结果都保存在"result.xlsx"，表格中每个sheet都以文件夹名字命名
2. 对那些格式错误的文件, 你可以在 对应文件夹下的"文件待修改格式.txt"中查看. 这些文件的结果不会被记录在"result.xlsx"中
3. 如果要改变年龄，请在No.2代码块中更改年龄计算代码。
4. 要将数据类型更改为float或int，请使用“ int()”或“ float()”强制转换
5. 要删除字符串中的空格，请使用".replace('','')"
6.  要增加从另一个数据计算得到的新数据
    - 添加一个标题到sheet_head
    - 参考 No.2 代码块中的示例
    ```
	  sheet.write(patient_count, sheet_head.index('MAD'), (meantime_hypo*times_hypo+times_apnea*meantime_apnea)/(times_hypo+times_apnea))
    ```
7. 要更正格式错误的文件，您可以更改错误文件中的数据,再跑一边脚本，而无需手动键入数据。
8. 如果stage数目为4，请解除代码162行到169行的注释，注释掉170行到173行
