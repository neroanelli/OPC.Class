Attribute VB_Name = "Search"
Option Base 1

Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long  '判断数组为空
Public Declare Function timeGetTime Lib "winmm.dll" () As Long             '获取开机至今过去多少时间
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long     '时间分辨率
Public Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

'读取数据库

Public Function ReadData(TableName As String, TagName As String) As Report_Data()
    Dim TagValue() As Report_Data
    Dim i As Integer
    Dim rs_Num As Integer
    MyData = "database\OPC.mdb"
    Set Cnn = New ADODB.Connection
'建立与数据库的连接
    With Cnn
        .Provider = "microsoft.jet.oledb.4.0"
        .Open MyData
    End With
'查询数据表
    Set rs = New ADODB.Recordset
    rs.Open TableName, Cnn, 1, 1
    rs.MoveFirst
    rs_Num = rs.RecordCount
   ' rs_Numb = rs.Fields.Count
  ReDim TagValue(rs_Num) As Report_Data
' ReDim ReadData(rs_Num) As Report_Data
     i = 1

 Do While Not rs.EOF
 TagValue(i).TagName = rs(TagName)
 TagValue(i).N = rs(0)
' TagValue(i).TagDIS = rs("TagDIS")
                i = i + 1
                rs.MoveNext
Loop

 ReadData = TagValue
'弹出提示信息
    'MsgBox "XXXXXXXXXX", vbInformation + vbOKOnly
    '关闭数据集和与数据库的连接，并释放变量
    rs.Close
    Cnn.Close
    Set rs = Nothing
    Set Cnn = Nothing

End Function
'存储数据
Public Function SaveData(TableNames As String, UpNum As Integer, DownNum As Integer, LastNum As Integer, VarName() As Variant, Now_Hour As Integer)
 MyData = "database\OPC1.mdb"
 Set Cnn = New ADODB.Connection
    With Cnn
        .Provider = "microsoft.jet.oledb.4.0"
        .Open MyData
    End With
    Set rs = New ADODB.Recordset
    rs.Open TableNames, Cnn, 1, 3
    rs.MoveLast
    rs.AddNew
    rs("日期") = Now
    If SafeArrayGetDim(VarName) = 0 Then
    For i = UpNum To DownNum + 2 'fmMain.LvListView.ListItems.Count
        rs(i - LastNum - 2) = fmMain.LvListView.ListItems(i).SubItems(3)
    Next i
    
    Else
        If Now_Hour = 16 Then '中班记录插入算法
        rs(2) = fmMain.LvListView.ListItems(1).SubItems(3) - VarName(2, 0) '产量
        rs(3) = fmMain.LvListView.ListItems(2).SubItems(3) - VarName(3, 0) '运行时间
           If rs(3) > 0 Then
             For i = UpNum To DownNum 'fmMain.LvListView.ListItems.Count
                rs(i - LastNum) = (fmMain.LvListView.ListItems(i + 2).SubItems(3) * fmMain.LvListView.ListItems(2).SubItems(3) - VarName(i - LastNum, 0) * VarName(3, 0)) / rs(3)  '参数均值
             Next i
            Else
                For i = UpNum To DownNum
                rs(i - LastNum) = 0
                Next i
             
           End If
         Else
         If Now_Hour = 0 Then '夜班记录插入算法
         rs(2) = fmMain.LvListView.ListItems(1).SubItems(3) - VarName(2, 1) - VarName(2, 0) '产量
         rs(3) = fmMain.LvListView.ListItems(2).SubItems(3) - VarName(3, 1) - VarName(3, 0) '运行时间
         If rs(3) > 0 Then
            For i = UpNum To DownNum 'fmMain.LvListView.ListItems.Count
                 rs(i - LastNum) = (fmMain.LvListView.ListItems(i + 2).SubItems(3) * fmMain.LvListView.ListItems(2).SubItems(3) - VarName(i - LastNum, 0) * VarName(3, 0) - VarName(i - LastNum, 1) * VarName(3, 1)) / rs(3)
            Next i
            Else
                 For i = UpNum To DownNum
                 rs(i - LastNum) = 0
                 Next i
            End If
         End If
        End If
    End If
    rs.Update
    
   
    rs.Close
    Cnn.Close
    Set rs = Nothing
    Set Cnn = Nothing

End Function
Public Function CheckData(TableName As String, k As Integer) As Variant
 MyData = "database\OPC1.mdb"
 Dim tt() As Variant
         Set Cnn = New ADODB.Connection
         With Cnn
         .Provider = "microsoft.jet.oledb.4.0"
         .Open MyData
          End With
          Set rs = New ADODB.Recordset
          rs.Open TableName, Cnn, 1, 1
          If Not rs.EOF Then
          rs.MoveLast
       End If
          If IsNull(rs(4)) Or IsNull(rs("日期")) Or k = 7 Then
         CheckData = tt
          Else
          If 9 < Hour(Now - rs("日期")) Then
         CheckData = tt
         Else
         If k = 16 Then
         CheckData = rs.GetRows(1)
         Else
         If k = 0 Then
         rs.MovePrevious
         CheckData = rs.GetRows(2)
         End If
         End If
          End If
          End If
           rs.Close
  Cnn.Close
    Set rs = Nothing
   Set Cnn = Nothing



















End Function
