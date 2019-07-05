<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030705") then Call ReturnError1()
Dim TypeSql,RsTypeObj,LableSql,RsLableObj
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>标签插入</title>
</head>
<style>
.LableSelectItem {
	background-color:highlight;
	cursor: hand;
	color: white;
	text-decoration: underline;
}
.LableItem {
	cursor: hand;
}
</style>
<link href="../../CSS/FS_css.css" rel="stylesheet">
<body ondragstart="return false;" onselectstart="return false;" topmargin="0" leftmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<%
TypeSql = "Select * from FS_NewsClass where ParentID='0' and delFlag=0 and IsOutClass=0"
Set RsTypeObj = Conn.Execute(TypeSql)
if Not RsTypeObj.Eof then
	do while Not RsTypeObj.Eof
%>
  <tr ParentID="<% = RsTypeObj("ParentID") %>">
	<td><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../Images/Folder/folderclosed.gif"></td>
          <td><span onDblClick="SubmitLable(this);" Extend="False" class="LableItem" TypeID="<% = RsTypeObj("ClassID") %>" onClick="SelectFolder(this)"><% = RsTypeObj("ClassCName") %></span></td>
        </tr>
      </table>
    </td>
</tr>
<%
		Response.Write(GetChildTypeList(RsTypeObj("ClassID"),""," style=""display:none;"" "))
		RsTypeObj.MoveNext
	loop
end if
%>
</table>
</body>
</html>
<%
Set Conn = Nothing
Function GetChildTypeList(TypeID,CompatStr,ShowStr)
	Dim ChildTypeListRs,ChildTypeListStr,TempStr
	Set ChildTypeListRs = Conn.Execute("Select * from FS_NewsClass where ParentID='" & TypeID & "' and IsOutClass=0" )
	TempStr = CompatStr & "&nbsp;&nbsp;&nbsp;&nbsp;"
	do while Not ChildTypeListRs.Eof
	  	GetChildTypeList = GetChildTypeList & "<tr TypeFlag=""Class"" ParentID=""" & ChildTypeListRs("ParentID") & """ " & ShowStr & ">" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & "<td>" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & Chr(13) & Chr(10) & "<tr>"  & Chr(13) & Chr(10) & "<td>" & TempStr & "<img src=""../Images/Folder/folderclosed.gif"" width=""18"" height=""18""></td>"
		GetChildTypeList = GetChildTypeList & "<td><span onDblClick=""SubmitLable(this);"" class=""LableItem"" TypeID=""" & ChildTypeListRs("ClassID") & """ Extend=""False"" onClick=""SelectFolder(this)"">" & ChildTypeListRs("ClassCName") & "</span></td>" & Chr(13) & Chr(10) & "</tr>" & Chr(13) & Chr(10) & "</table>"
		GetChildTypeList = GetChildTypeList & "</td>" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & "</tr>" & Chr(13) & Chr(10)
		GetChildTypeList = GetChildTypeList & GetChildTypeList(ChildTypeListRs("ClassID"),TempStr,ShowStr)
		ChildTypeListRs.MoveNext
	loop
	ChildTypeListRs.Close
	Set ChildTypeListRs = Nothing
End Function
%>
<script language="JavaScript">
function SelectLable(Obj)
{
	for (var i=0;i<document.all.length;i++)
	{
		if (document.all(i).className=='LableSelectItem') document.all(i).className='LableItem';
	}
	Obj.className='LableSelectItem';
}
function SelectFolder(Obj)
{
	var CurrObj;
	for (var i=0;i<document.all.length;i++)
	{
		if (document.all(i).className=='LableSelectItem') document.all(i).className='LableItem';
	}
	Obj.className='LableSelectItem';
	if (Obj.Extend=='True')
	{
		ShowOrDisplay(Obj,'none',true);
		Obj.Extend='False';
	}
	else
	{
		ShowOrDisplay(Obj,'',false);
		Obj.Extend='True';
	}
}
function ShowOrDisplay(Obj,Flag,Tag)
{
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if (CurrObj.ParentID==Obj.TypeID)
		{
			CurrObj.style.display=Flag;
			if (Tag) 
			if (CurrObj.TypeFlag=='Class') ShowOrDisplay(CurrObj.children(0).children(0).children(0).children(0).children(1).children(0),Flag,Tag);
		}
	}
}
function SubmitLable(Obj)
{
	var LableName=Obj.TypeID+'***'+Obj.innerText;
	window.returnValue=LableName;
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>