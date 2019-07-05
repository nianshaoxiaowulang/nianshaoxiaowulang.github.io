<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/NoSqlHack.asp" -->
<%  dim DBC,conn
    set DBC=new databaseclass
	set conn=DBC.openconnection()
	set DBC=nothing

Dim VoteID,VoteResultObj,Title,Totalcount,VoteResultSql,TotalNum,TDWidth
	VoteID=request("VoteID")
Set VoteResultObj = Conn.execute("select Name from FS_Vote where VoteID='"&VoteID&"'")
	Title = VoteResultObj("Name")
	VoteResultObj.close
Set VoteResultObj = Nothing
Set VoteResultObj = Server.CreateObject(G_FS_RS)
	VoteResultSql = "select * from FS_VoteOption where VoteID='"&VoteID&"'"
	VoteResultObj.Open VoteResultSql,Conn,1,1
	Totalcount = 0
	do while not VoteResultObj.eof
		Totalcount = Totalcount+cint(VoteResultObj("ClickNum"))
		VoteResultObj.movenext
	loop
	TotalNum = Totalcount
	if Totalcount=0 then
	   Totalcount=1
	end if
%>

<head>
<title>投票结果</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF" scroll="auto">
<div align="center">
  <center>
    <table width="100%" height="96" border="0" cellpadding="0" cellspacing="5" bordercolorlight="#000000" bordercolordark="#FFFFFF">
      <tr> 
        <td width="596" height="23" colspan="3"><div align="center"><b><%=title%><font color=#FF0000>(总投票<b><font color=red>数:</font><font color=red><%=TotalNum%></font></b>)</font></b></div></td>
      </tr>
      <%
	  VoteResultObj.movefirst
	  Do While Not VoteResultObj.eof
	  %>
      <tr> 
        <td width="164" height="26"><div align="left">&nbsp;&nbsp;&nbsp;<%=VoteResultObj("OptionName")%></div></td>
        <%TDWidth=(Cint(VoteResultObj("ClickNum"))*300)/Totalcount%>
        <td width="300" height="26" valign="middle"> 
          <table border="0" width="<%=TDWidth%>" height="20" bgcolor="<%=VoteResultObj("OptionColor")%>" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr> 
              <td width="100%" style="border-style: solid; border-width: 1" valign="middle"></td>
            </tr>
          </table>
        </td>
        <td width="129" height="26"> 
          <font color=green><%=VoteResultObj("ClickNum")%></font> 票(<%=(cint(VoteResultObj("ClickNum"))*10000/totalcount)/100%>%)</td>
      </tr>
      <%VoteResultObj.MoveNext
	  Loop%>
      <tr> 
        <td width="596" height="21" colspan="3"></td>
      </tr>
      <tr> 
        <td height="21" colspan="3"> 
          <div align="center">
            <input type="button" name="Submit" value="关闭窗口" onClick="window.close();">
            </div></td>
      </tr>
    </table>
  </center>
</div>
<%
 Set Conn = Nothing
 %>