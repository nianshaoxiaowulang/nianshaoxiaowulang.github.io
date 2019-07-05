<%
'──────────────────────────────── 
'功能说明：DataBaseClass类是实现数据库连接的类，里面留有数据库连接字符串接口
'包括模块：无，一般都是被其他模块包括
'调用方法：1、如果使用原有数据库连接，则不用更改数据库连接字符串ConnStr
'             具体操作为：Set DBC=New DataBaseClass
'                         DBC.ConnStr="其他连接字符串"
'          2、方法使用：Set Conn=DBC.OpenConnection()得到一个连接对象
'──────────────────────────────── 
Const IsSqlDataBase=0
Dim StrSqlDate
Class DataBaseClass
'──────────────────────────────── 
'定义变量 
Private IConnStr 
'──────────────────────────────── 
' ConnStr属性
Public Property Let ConnStr(Val)
    IConnStr = Val
End Property
'──────────────────────────────── 
' ConnStr属性 
Public Property Get ConnStr()
    ConnStr = IConnStr
End Property
'──────────────────────────────── 
' 类初始化 
Private Sub Class_initialize()
	ConnStr = "DBQ=" + Server.MapPath(DataBaseConnectStr) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
	StrSqlDate = "Date()"
End Sub 
'──────────────────────────────── 
' 类注销 
Private Sub Class_Terminate() 
	ConnStr = Null 
End Sub 
'──────────────────────────────── 
' 建立一个连接 
Public Function OpenConnection() 
	Dim TempConn
	On Error Resume Next
	Set TempConn = Server.CreateObject(G_FS_CONN)
	TempConn.Open ConnStr 
	Set OpenConnection = TempConn 
	Set TempConn = Nothing 
	if Err.Number <> 0 then
		Response.Write("<font size=""2"">[数据库连接错误]<br>请检查系统参数设置>>站点常量设置,或者/inc/const.asp文件!</font>")  
        Response.End
	end if
End Function 
End Class
%>