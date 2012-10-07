Attribute VB_Name = "Variable"
Option Base 1
Public itx As ListItem
Public i As Integer
Public k As Integer
Public Cnn As ADODB.Connection
Public rs As ADODB.Recordset
Public MyData As String '= "database\OPC.mdb"
Type Report_Data
    TagName As String       '结构体中标签名
    TagDIS As String             '结构体中量程上限
    LL As Double            '结构体中量程下限
    N As Integer            '结构体中有效采集累加的计数值
    Value As Double         '对应的采集值
    
End Type



'存储1个小时累计量的 结构体数组
'Public TagValue() As Report_Data
