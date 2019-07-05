<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../Refresh/Function.asp" -->
<%
Dim DBC,Conn
Set DBC=new DataBaseClass
Set Conn=DBC.OpenConnection
Set DBC=Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not ((JudgePopedomTF(Session("Name"),"P030802")) OR (JudgePopedomTF(Session("Name"),"P030803"))) then Call ReturnError1()
Dim LableID,RsLableObj,SQLStr,LableName,LableDescription,LableContent,Operation,LableType,BigTypeID
LableID = Request("ID")
BigTypeID = Request("BigTypeID")
if BigTypeID <> "" then
	if BigTypeID < 0 OR BigTypeID = "" then
		BigTypeID = "0"
	else
		BigTypeID = BigTypeID
	end if
end if
Operation = Request.Form("Operation")
if LableID <> "" then
	if Operation = "Modify" then
			LableName = Request.Form("LableName")
			LableDescription = Request.Form("Description")
			LableContent = Request.Form("LableContent")
			LableType = Request.Form("LableType")
	else
		Set RsLableObj = CreateObject(G_FS_RS)
		SQLStr = "Select * From FS_Lable where ID="&LableID&""
		RsLableObj.Open SQLStr,Conn,1,3
		if Not RsLableObj.Eof then
			LableName = RsLableObj("LableName")
			LableDescription = RsLableObj("Description")
			LableContent = RsLableObj("LableContent")
			LableType = CStr(RsLableObj("Type"))
		else
			LableName = ""
			LableDescription = ""
			LableContent = ""
			LableType = ""
		end if
		Set RslableObj = Nothing
	end if
else
	LableName = Request.Form("LableName")
	LableDescription = Request.Form("Description")
	LableContent = Request.Form("LableContent")
	if Request.Form("LableType") = "" then
		LableType = BigTypeID
	else
		LableType = Request.Form("LableType")
	end if
end if
LableContent = Replace(Replace(Replace(LableContent,"""","%22"),"'","%27"),WebDomain,"")
if Operation = "Modify" then SaveLable LableID
Sub SaveLable(EditLableID)
	Dim RsTemp,EditSql,RsCheckObj,CheckSql
	if Replace(Replace(Request.form("LableName"),"{FS_",""),"}","") = "" then
		AlertUser "请填写标签名称"
		Exit Sub
	else
		if EditLableID = "" then
			CheckSql = "Select * from FS_Lable where LableName='{FS_" & Request.form("LableName") & "}'"
		else
			CheckSql = "Select * from FS_Lable where LableName='{FS_" & Request.form("LableName") & "{FS_' and ID<>" & EditLableID
		end if

		Set RsCheckObj = Conn.Execute(CheckSql)
		if Not RsCheckObj.Eof then
			AlertUser "标签名已经存在"
			Set RsCheckObj = Nothing
			Exit Sub
		end if
		Set RsCheckObj = Nothing
	end if
	On Error Resume Next
	Set RsTemp = Server.CreateObject(G_FS_RS)
	if EditLableID = "" then
		EditSql = "Select * from FS_Lable where 1=0"
		RsTemp.Open EditSql,Conn,3,3
		RsTemp.AddNew
	else
		EditSql = "Select * from FS_Lable where ID=" & LableID
		RsTemp.Open EditSql,Conn,3,3
		if RsTemp.Eof then AlertAndCloseWindow "修改的标签不存在"
	end if
	RsTemp("LableName") = "{FS_" & Request.form("LableName") & "}"
	RsTemp("LableContent") = Request.form("LableContent")
	RsTemp("Description") = Request.form("Description")
	if Request.form("LableType") <> "" then
		RsTemp("Type") = CInt(Request.form("LableType"))
	else
		RsTemp("Type") = 0
	end if
	RsTemp.UpDate
	RsTemp.Close
	Set RsTemp = Nothing
	if err.Number=0 then
		if EditLableID = "" then
			PromptUser
		else
			Response.Redirect("Templet_LableList.asp?TypeID=" & BigTypeID)
		end if
	else
		AlertUser("操作失败！")
	end if
End Sub
Sub PromptUser()
	%>
	<script language="javascript">
		if (confirm('还要添加标签吗？')) location='LableAddNew.asp?BigTypeID=<% = BigTypeID %>';
		else location='Templet_LableList.asp?TypeID=<% = BigTypeID %>';
	</script>
	<%
End Sub
Sub AlertUser(Str)
	%>
	<script language="javascript">
		alert('<% = Str %>');
	</script>
	<%
End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建标签</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body scroll=yes bgcolor="#FFFFFF" topmargin="2" leftmargin="2"  oncontextmenu="return false;">
<form name=LableForm method=post action="" >
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="AddLableHead();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp;<input type="hidden" name="ID" value="<% = LableID %>"> <input type=hidden name=operation value=Modify>
              <input type="hidden" name="LableContent" value="<% = LableContent %>"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td  height="30"> 
        <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60" height="19"> 
              <div align="left">标签名称</div></td>
            <td><input value="<% = Replace(Replace(LableName,"{FS_",""),"}","") %>" name="LableName" style="width:100%;"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="30"> 
        <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60" height="16"> 
              <div align="left">标签描述</div></td>
            <td><textarea name="Description" rows="3" style="width:100%;"><% = LableDescription %></textarea></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="30"> 
        <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60" height="16"><div align="left">标签分类</div></td>
            <td><select name="LableType">
                <option <% if BigTypeID = "0" then Response.Write("Selected") %> value="0">根类型</option>
                <%
				Dim LableTypeObj
				Set LableTypeObj = Conn.Execute("Select * from FS_LableType")
				do while Not LableTypeObj.Eof
				%>
                <option <% if LableType = CStr(LableTypeObj("ID")) then Response.Write("Selected") %> value="<% = LableTypeObj("ID") %>">
                <% = LableTypeObj("TypeName") %>
                </option>
                <%
					LableTypeObj.MoveNext
				Loop
				Set LableTypeObj = Nothing
				%>
              </select></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td align="center"><iframe id="Editer" src="../../Editer/LableEditer.asp" scrolling="no" width="100%" height="94%" frameborder="0"></iframe></td>
    </tr>
</table>
</form>
</body>
</html>
<%
Set Conn = Nothing
%>
<script language="JavaScript">
var DocumentReadyTF=false;
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	SetHTML();
	DocumentReadyTF=true;
}
function AddLableHead()
{
	if (CheckAdminForm())
	{
		SetCode();
		//document.LableForm.LableName.value='{FS_'+document.LableForm.LableName.value+'}';
		document.LableForm.submit();
	}
}
function CheckAdminForm()
{
	var ErrorCode='';
	if (frames["Editer"].CurrMode!='EDIT') {alert('其他模式下无法保存，请切换到编辑模式');return;}
	if (document.LableForm.LableName.value=='') ErrorCode=ErrorCode+'标签名称不能为空！\n';
	if (ErrorCode!='') 
	{
		alert(ErrorCode);
		return false
	}
	else return true;
}
function SetHTML()
{

	frames["Editer"].EditArea.document.body.innerHTML=unescape(document.all.LableContent.value);
	frames["Editer"].ShowTableBorders();
}
function SetCode()
{
	document.all.LableContent.value=frames["Editer"].EditArea.document.body.innerHTML;
	frames["Editer"].ShowTableBorders();
}
</script>
