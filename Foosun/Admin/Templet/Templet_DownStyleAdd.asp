<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not ((JudgePopedomTF(Session("Name"),"P031001")) OR (JudgePopedomTF(Session("Name"),"P031002"))) then Call ReturnError()
Dim StyleID,RsStyleObj,SQLStr,StyleName,StyleContent,Operation
StyleID = Request("ID")
If (Not IsNumeric(StyleID)) and StyleID<>"" then response.end
If instr(Request.form("StyleName"),";")<>0 or instr(Request.form("StyleName"),"'")<>0 then response.end
Operation = Request.Form("Operation")
if StyleID <> "" then
	if Operation = "Modify" then
		StyleName = NoCSSHackAdmin(Request.Form("StyleName"),"��ʽ����")
		StyleContent = Request.Form("StyleContent")
	else
		Set RsStyleObj = CreateObject("ADODB.RecordSet")
		SQLStr = "Select * From FS_DownListStyle where ID=" & StyleID & ""
		RsStyleObj.Open SQLStr,Conn,1,3
		if Not RsStyleObj.Eof then
			StyleName = RsStyleObj("Name")
			StyleContent = RsStyleObj("Content")
		else
			StyleName = ""
			StyleContent = ""
		end if
		Set RsStyleObj = Nothing
	end if
else
	StyleName = NoCSSHackAdmin(Request.Form("StyleName"),"��ʽ����")
	StyleContent = Request.Form("StyleContent")
end if
StyleContent = Replace(Replace(StyleContent,"""","%22"),"'","%27")
if Operation = "Modify" then SaveStyle StyleID
Sub SaveStyle(EditStyleID)
	Dim RsTemp,EditSql,RsCheckObj,CheckSql
	if Request.form("StyleName") = "" then
		Alert "����д��ʽ����"
		Exit Sub
	else
		if EditStyleID = "" then
			CheckSql = "Select * from FS_DownListStyle where Name='" & Request.form("StyleName") & "'"
		else
			CheckSql = "Select * from FS_DownListStyle where Name='" & Request.form("StyleName") & "' and ID<>" & EditStyleID
		end if
		Set RsCheckObj = Conn.Execute(CheckSql)
		if Not RsCheckObj.Eof then
			Alert "��ǩ���Ѿ�����"
			Set RsCheckObj = Nothing
			Exit Sub
		end if
		Set RsCheckObj = Nothing
	end if
	'On Error Resume Next
	Set RsTemp = Server.CreateObject("ADODB.recordset")
	if EditStyleID = "" then
		EditSql = "Select * from FS_DownListStyle where 1=0"
		RsTemp.Open EditSql,Conn,3,3
		RsTemp.AddNew
	else
		EditSql = "Select * from FS_DownListStyle where ID=" & StyleID
		RsTemp.Open EditSql,Conn,3,3
		if RsTemp.Eof then Alert "�޸ĵı�ǩ������"
	end if
	RsTemp("Name") = Request.Form("StyleName")
	RsTemp("Content") = Request.form("StyleContent")
	RsTemp.UpDate
	RsTemp.Close
	Set RsTemp = Nothing
	if err.Number=0 then
		Response.Redirect("Templet_DownStyleList.asp")
	else
		if StyleID <> "" then
			Alert "�޸�ʧ��"
		else
			Alert "���ʧ��"
		end if
	end if
End Sub
Sub Alert(ErrorStr)
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
<title>��Ӻ��޸������б���ʽ</title>
</head>
<script language="javascript" event="onerror(msg, url, line)" for="window">return true;</script>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<body topmargin="0" leftmargin="0">
<table width="100%" height="100%" border="0" cellpadding="1" cellspacing="1">
  <form name=StyleForm method=post action="" >
    <tr> 
      <td colspan="5" height="32" valign="top"> 
        <table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
		  <tr bgcolor="#EEEEEE"> 
			<td height="26" colspan="5" valign="middle">
			  <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
				<tr>
				  <td width=35 align="center" alt="����" onClick="AddLableHead();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
				  <td width=2 class="Gray">|</td>
				  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
				  <td>&nbsp;</td>
				</tr>
			  </table>
			  </td>
		  </tr>
		</table>
	  </td>
    </tr>
	<tr> 
      <td  height="30" id="StyleNameArea"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60"> <div align="center">��ʽ����</div></td>
            <td><input value="<% = StyleName %>" name="StyleName" style="width:100%;">
			<input type="hidden" name="ID" value="<% = StyleID %>"> 
        <input type="hidden" name="operation" value="Modify">
              <input type="hidden" name="StyleContent" value="<% = StyleContent %>">
            </td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td  height="30" id="StyleNameArea"> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="60"> <div align="center">�����ֶ�</div></td>
            <td> <select style="width:82%;" name="FieldList">
                <option value="" selected>ѡ�������ֶ�</option>
                <option value="{DownLoad_Name}">��������</option>
                <option value="{DownLoad_Version}">�汾</option>
                <option value="{DownLoad_Types}">��������</option>
                <option value="{DownLoad_ClickNum}">���ش���</option>
<!--                <option value="{DownLoad_Property}">��������</option>-->
                <option value="{DownLoad_Language}">����</option>
                <option value="{DownLoad_Accredit}">��Ȩ</option>
                <option value="{DownLoad_FileSize}">�ļ���С</option>
                <option value="{DownLoad_Appraise}">����</option>
                <option value="{DownLoad_SystemType}">ϵͳƽ̨</option>
                <option value="{DownLoad_EMail}">��ϵ��EMAIL</option>
                <option value="{DownLoad_ProviderUrl}">�ṩ��Url��ַ</option>
                <option value="{DownLoad_Provider}">������</option>
                <option value="{DownLoad_Pic}">��ʾͼƬ</option>
                <option value="{DownLoad_Description}">���</option>
                <option value="{DownLoad_PassWord}">��ѹ����</option>
                <option value="{DownLoad_AddTime}">���ʱ��</option>
                <option value="{DownLoad_EditTime}">�޸�ʱ��</option>
              </select> <input name="Submitfff" type="button" id="Submitfff" onClick="InsertField();" value="�����ֶ�" style="color=#FF0000"> 
            </td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td colspan="2" align="center"><iframe id="Editer" src="../../Editer/DownStyleEditer.asp" scrolling="no" width="100%" height="100%" frameborder="0"></iframe></td>
    </tr>
  </form>
</table>
</body>
</html>
<%
Set Conn = Nothing
%>
<script language="JavaScript">
var BodyHeight=document.body.clientHeight;
var EditAreaHeight=BodyHeight;
setTimeout('SetHTML();',300);
function SetHTML()
{
	if (frames["Editer"].EditArea)
	{
		frames["Editer"].EditArea.document.body.innerHTML=unescape(document.all.StyleContent.value);
		frames["Editer"].ShowTableBorders();
	}
	else
	{
		setTimeout('SetHTML();',300);
	}
}
function AddLableHead()
{
	if (CheckAdminForm()) 
	{
		
		document.StyleForm.submit();
	}
}
function CheckAdminForm()
{
	var ErrorCode='';
	document.StyleForm.StyleContent.value=frames["Editer"].EditArea.document.body.innerHTML;
	if (frames["Editer"].CurrMode!='EDIT') {alert('����ģʽ���޷����棬���л����༭ģʽ');return;}
	if (document.StyleForm.StyleName.value=='') ErrorCode=ErrorCode+'��ʽ���Ʋ���Ϊ�գ�\n';
	if (document.StyleForm.StyleContent.value=='') ErrorCode=ErrorCode+'��ʽ���ݲ���Ϊ�գ�\n';
	if (ErrorCode!='') 
	{
		alert(ErrorCode);
		return false
	}
	else return true;
}
function InsertField()
{
	var ReturnValue=document.all.FieldList.value;
	frames["Editer"].EditArea.focus();
	if (ReturnValue!='') frames["Editer"].InsertHTMLStr(ReturnValue);
}
</script>
