<%
'���������������������������������������������������������������� 
'����˵����DataBaseClass����ʵ�����ݿ����ӵ��࣬�����������ݿ������ַ����ӿ�
'����ģ�飺�ޣ�һ�㶼�Ǳ�����ģ�����
'���÷�����1�����ʹ��ԭ�����ݿ����ӣ����ø������ݿ������ַ���ConnStr
'             �������Ϊ��Set DBC=New DataBaseClass
'                         DBC.ConnStr="���������ַ���"
'          2������ʹ�ã�Set Conn=DBC.OpenConnection()�õ�һ�����Ӷ���
'���������������������������������������������������������������� 
Const IsSqlDataBase=0
Dim StrSqlDate
Class DataBaseClass
'���������������������������������������������������������������� 
'������� 
Private IConnStr 
'���������������������������������������������������������������� 
' ConnStr����
Public Property Let ConnStr(Val)
    IConnStr = Val
End Property
'���������������������������������������������������������������� 
' ConnStr���� 
Public Property Get ConnStr()
    ConnStr = IConnStr
End Property
'���������������������������������������������������������������� 
' ���ʼ�� 
Private Sub Class_initialize()
	ConnStr = "DBQ=" + Server.MapPath(DataBaseConnectStr) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
	StrSqlDate = "Date()"
End Sub 
'���������������������������������������������������������������� 
' ��ע�� 
Private Sub Class_Terminate() 
	ConnStr = Null 
End Sub 
'���������������������������������������������������������������� 
' ����һ������ 
Public Function OpenConnection() 
	Dim TempConn
	On Error Resume Next
	Set TempConn = Server.CreateObject(G_FS_CONN)
	TempConn.Open ConnStr 
	Set OpenConnection = TempConn 
	Set TempConn = Nothing 
	if Err.Number <> 0 then
		Response.Write("<font size=""2"">[���ݿ����Ӵ���]<br>����ϵͳ��������>>վ�㳣������,����/inc/const.asp�ļ�!</font>")  
        Response.End
	end if
End Function 
End Class
%>