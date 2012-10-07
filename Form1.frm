VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fmMain 
   Caption         =   "OPC Application By Mister.T"
   ClientHeight    =   9000
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   14535
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "Message"
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   14055
      Begin VB.Label Label3 
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   13215
      End
   End
   Begin VB.CommandButton Cmd_Refresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Read"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   9
      Top             =   4560
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   8625
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19976
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2012-4-11"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LvListView 
      Height          =   5895
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10398
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "导入变量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.CheckBox DataChgChk 
      Caption         =   "使用订阅数据采取"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12120
      TabIndex        =   3
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.Timer tmUpdate 
      Enabled         =   0   'False
      Left            =   13440
      Top             =   4200
   End
   Begin VB.CommandButton btnAddItem 
      Caption         =   "读取"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton btnQuit 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   1
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "连接"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   0
      Top             =   3000
      Width           =   975
   End
   Begin VB.Frame OPC 
      Caption         =   "OPC数据采集"
      Height          =   7335
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   14055
      Begin VB.TextBox Text2 
         Height          =   855
         Left            =   7080
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "Form1.frx":0000
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Text            =   "\\127.0.0.1"
         Top             =   360
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         Left            =   1560
         TabIndex        =   12
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "Sever Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Server Node:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Menu fileMenu 
      Caption         =   "文件&F"
      Begin VB.Menu MenuOpen 
         Caption         =   "打开&O"
      End
   End
   Begin VB.Menu reporMenu 
      Caption         =   "报表&R"
      Begin VB.Menu MreportMenu 
         Caption         =   "月报表"
      End
      Begin VB.Menu DreportMenu 
         Caption         =   "班报表"
      End
   End
   Begin VB.Menu aboutMenu 
      Caption         =   "关于&A"
   End
End
Attribute VB_Name = "fmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Option Explicit

''\ OPC对象的声明
'Dim WithEvents objserver As OPCServer
'Dim objGroups As OPCGroups
'Dim WithEvents objtestgrp As OPCGroup '事件的对应
'Dim objItems As OPCItems
'Dim LServerHandles() As Long
    Dim x As Integer
    Dim y As Integer
    Dim mon As Integer
    Dim d As Integer
    Dim h As Integer
    Dim m As Integer
    Dim s As Integer
    Dim Num_all1 As Long '变量总数
   Public UserSession  As Boolean
   
        Private WithEvents opcserver As OPCClass
Attribute opcserver.VB_VarHelpID = -1
    
'Dim lTransID_Rd As Long
'Dim lCancelID_Rd As Long
'Dim lTransID_Wt As Long
'Dim lCancelID_Wt As Long

'Sub Connect(strProgID As String, Optional strNode As String)
'
'    If objserver Is Nothing Then
'        ' 建立一个OPC服务器对象
'        Set objserver = New OPCServer
'    End If
'
'    If objserver.ServerState = OPCDisconnected Then
'        ' 连接OPC服务器
'        objserver.Connect strProgID, strNode
'    End If
'
'    If objGroups Is Nothing Then
'        ' 建立一个OPC组集合
'        Set objGroups = objserver.OPCGroups
'    End If
'
'    If objtestgrp Is Nothing Then
'        ' 添加一个OPC组
'        Set objtestgrp = objGroups.Add("TestGrp")
'    End If
'
'End Sub

'Sub Disconnect()
'    Dim lErrors() As Long
'
'    If Not objItems Is Nothing Then
'        If objItems.Count > 0 Then
'            ' 清除OPC项
'            objItems.Remove 113, LServerHandles, lErrors
'        End If
'        Set objItems = Nothing
'    End If
'
'    If Not objtestgrp Is Nothing Then
'        ' 清除OPC组
'        objGroups.Remove "TestGrp"
'        Set objtestgrp = Nothing
'    End If
'
'    If Not objGroups Is Nothing Then
'        Set objGroups = Nothing
'    End If
'
'    If Not objserver Is Nothing Then
'        If objserver.ServerState <> OPCDisconnected Then
'            ' 断开OPC服务器.
'            objserver.Disconnect
'        End If
'
'        Set objserver = Nothing
'    End If
'
'End Sub

'Sub AddItem()
'    Dim strItemIDs(113) As String
'    Dim lClientHandles(113) As Long
'    Dim lErrors() As Long
'    Dim i As Integer
'    Dim ExcelApp As Excel.Application
'    Dim ExcelBook As Excel.Workbook
'    Dim ExcelSheet As Excel.Worksheet
'    Dim Strfilename As String
'
'    If objtestgrp Is Nothing Then
'        Exit Sub
'    End If
'
'    If Not objItems Is Nothing Then
'        If objItems.Count > 0 Then
'            Exit Sub
'        End If
'    End If
'
'    ' 设置组活动状态
'    If DataChgChk.Value = vbChecked Then
'
'        objtestgrp.IsActive = True
'    Else
'        objtestgrp.IsActive = False
'    End If
'    ' 启动组非同期通知
'    objtestgrp.IsSubscribed = True
'
'    ' 建立OPC项集合
'    Set objItems = objtestgrp.OPCItems
'    '查询数据库，得到TAG名
'     Dim TagName As String, MyTable As String
'     TagName = "TAGNAME"
'     MyTable = "TAG"
'     Dim testg As Variant
'
'    ' testg = ReadData(MyTable, Tagname)
'   For i = 1 To LvListView.ListItems.Count
'
'   lClientHandles(i) = i
'   strItemIDs(i) = LvListView.ListItems(i).SubItems(1)
'   Next i
'
''     I = 1
'' Do While Not rs.EOF
' ' strItemIDs(I) = rs("TAGNAME")
'  '              lClientHandles(I) = I
'   '             I = I + 1
'    '            rs.MoveNext
''Loop
'
'
'
''弹出提示信息
'   ' MsgBox "XXXXXXXXXX", vbInformation + vbOKOnly
'    '关闭数据集和与数据库的连接，并释放变量
'
'
'    StatusBar1.Panels(1).Text = " 数据加载成功"
'
'    ' 生成从TAG1到TAG8的项标识符
'   ' For I = 1 To 5
'  '      strItemIDs(I) = "Simulation Items.Integer.Int_0" & I
' '       lClientHandles(I) = I
' '   Next
' '   For I = 6 To 8
' '            strItemIDs(I) = "Simulation Items.Real.Real_0" & I - 5
''        lClientHandles(I) = I
''    Next
'    ' 添加OPC项
'    Call objItems.AddItems(113, strItemIDs, _
'        lClientHandles, LServerHandles, lErrors)
'
'End Sub

'Sub AsyncRead()
'    Dim lErrors() As Long
'   StatusBar1.Panels(1).Text = "Data Reading..."
'    If objtestgrp Is Nothing Then
'        Exit Sub
'    End If
'
'    If objtestgrp.OPCItems.Count > 0 Then
'        ' 非同期读取
'        lTransID_Rd = lTransID_Rd + 1
'        objtestgrp.AsyncRead 113, LServerHandles, _
'            lErrors, lTransID_Rd, lCancelID_Rd
'    End If
'
'End Sub

Sub AsyncWrite(nIndex As Integer, ByRef vtItemValues() As Variant, _
    ByRef lErrors() As Long)
Dim lHandle(1) As Long
    
    If objtestgrp Is Nothing Then
        Exit Sub
    End If
    
    If objtestgrp.OPCItems.Count > 0 Then
        lHandle(1) = LServerHandles(nIndex)
        
        ' 非同期写入
        lTransID_Wt = lTransID_Wt + 1
        objtestgrp.AsyncWrite 1, lHandle(), vtItemValues, _
                lErrors, lTransID_Wt, lCancelID_Wt
    End If

End Sub

Private Sub aboutMenu_Click()
fmAbout.Show
End Sub

Private Sub Cmd_Refresh_Click()
   On Error Resume Next
    Combo1.Clear
    Dim i As Integer
    Dim servername As Variant
    Dim opcs As OPCClass

Set opcs = New OPCClass
    servername = opcs.GetOPCServers(Text1.Text)
Set opcs = Nothing
   For i = LBound(servername) To UBound(servername)
       Combo1.AddItem servername(i), i - 1
   Next i
End Sub

Private Sub Command1_Click()   '导入变量

    If frmLogin.LoginSucceeded = False Then
    
    frmLogin.Show
    Else
    Dim i As Integer
    Dim ExcelApp As Excel.Application
    Dim ExcelBook As Excel.Workbook
    Dim ExcelSheet As Excel.Worksheet
    Dim Strfilename As String
    Dim xxx(8) As Variant
    Dim yyy(8) As Variant

    Strfilename = App.Path & "\database\TAG.xls"
    Set ExcelApp = New Excel.Application
    Set ExcelBook = ExcelApp.Workbooks.Open(Strfilename)
    
    Set ExcelSheet = ExcelBook.Sheets(1)
    ExcelApp.Visible = False
    
    With Worksheets("Sheet1")
            For i = 1 To 8
            ' 从工作表中得到TAG1到TAG8的项标识符
                xxx(i) = .Cells(i, 1).Text
              '  yyy(i) = .Cells(i + 1, 2).Text
            Next i
  '       .Range("A2:i65").ClearContents
        End With
    ExcelApp.Quit
    Set ExcelBook = Nothing
    Set ExcelSheet = Nothing
    Set ExcelApp = Nothing
    

           
           '连接ACCESS

           
    Dim MyData As String, MyTable As String
    Dim Cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim x As Integer
    Dim myArray As Variant
    
   MyData = "Database\OPC.mdb"
MyTable = "TAG"
'建立与数据库的连接
    Set Cnn = New ADODB.Connection
    With Cnn
        .Provider = "microsoft.jet.oledb.4.0"
        .Open MyData
End With
'查询数据表
    Set rs = New ADODB.Recordset
    rs.Open MyTable, Cnn, 1, 3
    '添加记录
With rs
       
        For x = 1 To 8
         .AddNew      '添加各个字段的数据
            rs("TAGName") = xxx(x)
          '  rs("TAGDIS") = yyy(x)
         .Update
        Next x
             '更新数据表
End With
'弹出提示信息
       '关闭数据集和与数据库的连接，并释放变量
rs.Close
    Cnn.Close
    Set rs = Nothing
    Set Cnn = Nothing
     MsgBox "已经成功将变量导入数据库中！", vbInformation + vbOKOnly

'Dim mycat As New adox.Catalog  '定义ADOX的catalog对象变量
'Dim mytbl As New Table  '定义table对象变量
'
'
'
''设置数据库名称（包括完整路径）
'
''设置要创建的数据表名称
'MyTable = "TAGVALUE"
''建立与数据库的连接
'mycat.ActiveConnection = "provider=microsoft.jet.oledb.4.0;" & "data source=" & MyData
''删除数据库中已经存在的数据表
''mycat.Tables.Delete mytable
'
''创建数据表，并添加字段
'With mytbl
'    .Name = MyTable
'    .Columns.Append "学号", adVarWChar, 10
'    For x = 1 To 10
'        .Columns.Append yyy(x), adDouble
'    Next x
''    .Columns.Append "姓名", adVarWChar, 6
''    .Columns.Append "性别", adVarWChar, 1
''    .Columns.Append "班级", adVarWChar, 10
''    .Columns.Append "数学", adDouble
' '   .Columns.Append "语文", adSingle
''    .Columns.Append "物理", adSingle
''    .Columns.Append "化学", adSingle
''    .Columns.Append "英语", adSingle
''    .Columns.Append "DATE", adDBTimeStamp
'End With
''将创建的数据表添加到ADOX的tables集合中
'mycat.Tables.Append mytbl
''释放变量
'Set mycat = Nothing
'Set mytbl = Nothing
''弹出信息
'MsgBox "数据表<" & MyTable & ">创建成功！", vbInformation, "创建数据表"
End If

End Sub

Private Sub Command2_Click()
opcserver.AsyncRead
End Sub



Private Sub DataChgChk_Click()

    If DataChgChk.Value = vbChecked Then
        tmUpdate.Enabled = False
'        If Not objtestgrp Is Nothing Then
'            objtestgrp.IsActive = True
'        End If
    Else
        tmUpdate.Enabled = True
'        If Not objtestgrp Is Nothing Then
'            objtestgrp.IsActive = False
'        End If
    End If

End Sub


Private Sub Form_Load()
    ' 初始化全局变量
    On Error Resume Next
    
    DataChgChk.Value = vbUnchecked
    'tmUpdate.Enabled = True
    tmUpdate.Interval = 1000
    Dim i As Integer
    Dim servername As Variant
    Dim opcs As OPCClass

Set opcs = New OPCClass
    servername = opcs.GetOPCServers(Text1.Text)
Set opcs = Nothing
   For i = LBound(servername) To UBound(servername)
       Combo1.AddItem servername(i), i - 1
   Next i
    LvListView.ColumnHeaders.Add 1, , "序号", LvListView.Width / 4
    LvListView.ColumnHeaders.Add 2, , "标签变量", LvListView.Width / 4
    LvListView.ColumnHeaders.Add 3, , "标签名称", LvListView.Width / 4
    LvListView.ColumnHeaders.Add 4, , "标签值", LvListView.Width / 4
    
     Dim TagName As String, MyTable As String
     TagName = "TAGNAME"
     MyTable = "TAG"
     Dim testg() As Report_Data
   '  ReDim testg(113) As Report_Data'''''''''''''''''''''''''''''''''''''''///////////////////////////////////////////////////?Necessary？
     
     testg = ReadData(MyTable, TagName)
     Num_all1 = CInt(UBound(testg))
   For i = LBound(testg) To UBound(testg)
 'For i = 1 To 8
   
Set itx = LvListView.ListItems.Add(, , i)
itx.SubItems(1) = testg(i).TagName
itx.SubItems(2) = testg(i).N
   Next i
 
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' 调用Disconnect子程序
   Set opcserver = New OPCClass
    opcserver.Disconnect
    Set opcserver = Nothing
End Sub

Private Sub btnConnect_Click()
    '调用Connect子程序
    Set opcserver = New OPCClass
        If opcserver.Connect(Combo1.Text, Text1.Text) Then
      '  MsgBox "Connect OK!"
   'Call Connect("DSxPOpcSimulator.TSxOpcSimulator.1")
  ' Call Connect("DSxPOpcSimulator.TSxOpcSimulator.1", "127.0.0.1")
     '  Call Connect("Freelance2000OPCServer.49.1", "172.16.1.48")
    btnAddItem.Enabled = True
    End If
    
End Sub

Private Sub btnAddItem_Click()
    ' 调用AddItem子程序
'    Call AddItem
Dim tag As String
'Dim Savetime  As Double
tag = LvListView.ListItems(1).SubItems(1)
For i = 2 To LvListView.ListItems.Count
    tag = tag & "," & LvListView.ListItems(i).SubItems(1)
   Next i
   opcserver.AddItem tag
   Command2.Enabled = True
Label1:
opcserver.AsyncRead
Dim Savetime As Double
timeBeginPeriod 1
Savetime = timeGetTime
While timeGetTime < Savetime + 1000
'     If ss = True Then
'     timeEndPeriod 1
'     opcserver.Disconnect
'     Set opcserver = Nothing
'        Exit Sub
'    End If
DoEvents
Wend
GoTo Label1

   
End Sub

Private Sub btnQuit_Click()
    ' 卸载窗体
    Unload fmMain

End Sub









'Private Sub tmUpdate_Timer()
'    ' 非同期读取
'    Set opcserver = New OPCClass
'opcserver.Connect Combo1.Text, Text1.Text
'    Call opcserver.AsyncRead(Num_all1)
'
'End Sub


'Private Sub objtestgrp_AsyncReadComplete
Private Sub opcserver_AsyncReadComplete( _
    ByVal TransactionID As Long, ByVal NumItems As Long, _
    ClientHandles() As Long, ItemValues() As Variant, _
    Qualities() As Long, TimeStamps() As Date, Errors() As Long)
    Dim strBuf As String
    Dim nWidth As Integer
    Dim nHeight As Integer
    Dim nDrawHeight As Integer
    Dim sglScale As Single
    Dim i As Integer
    Dim Index As Integer
    Dim temp As String
 temp = "abcdefghijk"
 Text2.Text = ""
        For i = 1 To UBound(ItemValues)
         
       '   Text2.Text = Right(temp, Len(temp) - 1) & Left(temp, 1)
         Text2.Text = ItemValues(i) & vbCrLf
          Set itx = LvListView.ListItems(i)
         itx.SubItems(3) = ItemValues(i)
         Next i

 StatusBar1.Panels(1).Text = "Read Complete......"
 
 
   ReDim LastData(CInt(UBound(ItemValues))) As Variant
    
    StatusBar1.Panels(3).Text = Time
    x = x + 1
    s = Second(Time)
    m = Minute(Time)
    h = Hour(Time)

        
   If (s = 0 And m = 30 And h = 7) Or (h = 0 And m = 0 And s = 0) Or (h = 16 And m = 30 And s = 0) Then
   
    Dim MyData As String
    Dim MyTable(4) As String
    Dim xx As Integer
    Dim xxx(50) As Variant
    Dim myArray As Variant
    Dim j As Integer
    Dim DB_ep() As Variant
    Dim ModNub As Long
    xx = 8
    MyData = "database\OPC1.mdb"
    MyTable(1) = "立磨系统"
    MyTable(2) = "窑系统"
    MyTable(3) = "水泥A磨"
    MyTable(4) = "水泥B磨"
    DB_ep = CheckData(MyTable(1), h)
Call SaveData(MyTable(1), 1, 19, -3, DB_ep, h)
    DB_ep = CheckData(MyTable(2), h)
Call SaveData(MyTable(2), 22, 57, 18, DB_ep, h)
    DB_ep = CheckData(MyTable(3), h)
Call SaveData(MyTable(3), 60, 84, 56, DB_ep, h)
    DB_ep = CheckData(MyTable(4), h)
Call SaveData(MyTable(4), 87, 111, 83, DB_ep, h)
'Call SaveData(MyTable, 1, 21, -1,h)
'Call SaveData("tagvalue1", 22, 59, 20,h)
'Call SaveData("", 60, 86, 58,h)
'Call SaveData("", 87, 113, 85,h)
'Call SaveData("tagvalue1", 9, 16, 8, DB_ep, 16)
Else
End If
End Sub

Private Sub objtestgrp_AsyncWriteComplete( _
    ByVal TransactionID As Long, ByVal NumItems As Long, _
    ClientHandles() As Long, Errors() As Long)

End Sub

Private Sub objTestGrp_DataChange( _
    ByVal TransactionID As Long, ByVal NumItems As Long, _
    ClientHandles() As Long, ItemValues() As Variant, _
    Qualities() As Long, TimeStamps() As Date)
    Dim strBuf As String
    Dim nWidth As Integer
    Dim nHeight As Integer
    Dim nDrawHeight As Integer
    Dim sglScale As Single
    Dim i As Integer
    Dim Index As Integer
    ReDim LastData(CInt(UBound(ItemValues))) As Variant
    s = Second(Time)
    m = Minute(Time)
    h = Hour(Time)
    If h = 7 Then
    Else
    
    End If
    
        For i = 1 To UBound(ItemValues)

          Set itx = LvListView.ListItems(i)
         itx.SubItems(3) = ItemValues(i)
         Next i

        

End Sub



Private Sub MreportMenu_Click()
fmMrep.Show
End Sub
