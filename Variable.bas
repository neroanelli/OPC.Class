Attribute VB_Name = "Variable"
Option Base 1
Public itx As ListItem
Public i As Integer
Public k As Integer
Public Cnn As ADODB.Connection
Public rs As ADODB.Recordset
Public MyData As String '= "database\OPC.mdb"
Type Report_Data
    TagName As String       '�ṹ���б�ǩ��
    TagDIS As String             '�ṹ������������
    LL As Double            '�ṹ������������
    N As Integer            '�ṹ������Ч�ɼ��ۼӵļ���ֵ
    Value As Double         '��Ӧ�Ĳɼ�ֵ
    
End Type



'�洢1��Сʱ�ۼ����� �ṹ������
'Public TagValue() As Report_Data
