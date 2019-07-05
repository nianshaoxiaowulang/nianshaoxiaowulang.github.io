<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
Dim DBC,Conn,LableTypeObj,ID,NameType,TypeDes,TypeType,BigTypeID
Set DBC=new DataBaseClass
Set Conn=DBC.OpenConnection
Set DBC=Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not ((JudgePopedomTF(Session("Name"),"P030801")) OR (JudgePopedomTF(Session("Name"),"P030803"))) then Call ReturnError1()
ID=request("ID")
BigTypeID=request("BigTypeID")
if BigTypeID <> "" then
	if BigTypeID < 0 OR BigTypeID = "" then
		BigTypeID = "0"
	else
		BigTypeID = BigTypeID
	end if
end if
set LableTypeObj = Server.CreateObject(G_FS_RS)
if ID<>"" then
	LableTypeObj.open "select * from FS_LableType where ID="&ID,conn,3,3
	NameType = LableTypeObj("TypeName")
	TypeDes = LableTypeObj("Description")
	TypeType = Cstr(LableTypeObj("ParentID"))
	LableTypeObj.close
else
	NameType = ""
	TypeDes = ""
	TypeType = BigTypeID
end if
if Request.Form("Action") = "Submit" then
	if Request.Form("TypeName")="" then
		Response.Write("<script>alert(""请填写类名称"");location.href=""LableTypeAddNew.asp?BigTypeID="&request("BigTypeID")&""";</script>")
		Response.End
	end if
	on error resume next
	if ID<>"" then
		LableTypeObj.open "select * from FS_LableType where ID="&ID,conn,3,3
		dim TempTF
		TempTF = SelectObj(CInt(Request.form("TypeType")),ID)
		if TempTF then
			Response.Write("<script>alert(""错误：\n父类型不能移到子类型中"");location.href=""LableTypeAddNew.asp?BigTypeID="&request("BigTypeID")&"&ID="&request("ID")&""";</script>")
			Response.End
		end if
	else
		LableTypeObj.open "select * from FS_LableType",conn,3,3
		LableTypeObj.addnew
	end if
	LableTypeObj("TypeName")=NoCSSHackAdmin(request.Form("TypeName"),"类型名称")
	LableTypeObj("Description")=request.Form("Description")
	if Request.form("TypeType") <> "" then
		LableTypeObj("ParentID") = CInt(Request.form("TypeType"))
	else
		LableTypeObj("ParentID") = 0
	end if
	LableTypeObj.Update
	if err.number=0 then 
		Response.Redirect("Templet_LableList.asp?TypeID=" & LableTypeObj("ParentID"))
	else
		if ID<>"" then
			AlertUser "修改失败"
		else
			AlertUser "添加失败"
		end if
	end if
end if
Sub AlertUser(ErrorStr)
	Set Conn = Nothing
	%>
	<script language="javascript">
		alert ('<% = ErrorStr %>')
	</script>
	<%
End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>标签类型</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body scroll=no topmargin="2" leftmargin="2">
<form name=TypeForm method=post action="" >
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="document.TypeForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp;
              <input type=hidden name="Action" value="Submit"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="100" height="30"> <div align="center">类型名称</div></td>
      <td height="30"> <input style="width:100%;" name="TypeName" size="30" value="<% = NameType %>"></td>
    </tr>
    <tr> 
      <td height="30"> 
        <div align="center">父类型</div></td>
      <td height="30"> <select style="width:100%;" name="TypeType">
          <option <% if BigTypeID = "0" then Response.Write("Selected") %> value="0">根类型</option>
          <%
		Dim TempTypeObj
		Set TempTypeObj = Conn.Execute("Select * from FS_LableType")
		do while Not TempTypeObj.Eof 
			if TempTypeObj("TypeName") <> NameType then
				%>
          <option <% if TypeType = Cstr(TempTypeObj("ID")) then Response.Write("Selected") %> value="<% = TempTypeObj("ID") %>">
          <% = TempTypeObj("TypeName") %>
          </option>
          <%
			end if
		TempTypeObj.MoveNext
		Loop
		%>
        </select></td>
    </tr>
    <tr> 
      <td><div align="center">类型描述</div></td>
      <td><textarea name="Description" style="width:100%;" rows="8"><% = TypeDes %></textarea></td>
    </tr>
</table>
</form>
</body>
</html>
<%
Function SelectObj(SID,ID)
	Dim TempObj,Str
	Str = "Select * From FS_LableType Where ParentID= "& ID
	Set TempObj = conn.Execute(Str)
	do while not TempObj.eof
		if SID = TempObj("ID") then
			SelectObj = true
			Exit do
		end if
		SelectObj = SelectObj(SID,TempObj("ID"))
		if SelectObj = true then Exit do
	TempObj.movenext
	Loop
	TempObj.Close
	Set TempObj = Nothing
End Function
Set conn = Nothing
%>