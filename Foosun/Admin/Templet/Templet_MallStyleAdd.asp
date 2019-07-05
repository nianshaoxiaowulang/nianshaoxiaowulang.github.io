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
if Not ((JudgePopedomTF(Session("Name"),"P031201")) OR (JudgePopedomTF(Session("Name"),"P031202"))) then Call ReturnError()
Dim StyleID,RsStyleObj,SQLStr,StyleName,StyleContent,Operation
StyleID = Request("ID")
If (Not IsNumeric(StyleID)) and StyleID<>"" then response.end
If instr(Request.form("StyleName"),";")<>0 or instr(Request.form("StyleName"),"'")<>0 then response.end
Operation = Request.Form("Operation")
if StyleID <> "" then
	if Operation = "Modify" then
		StyleName = NoCSSHackAdmin(Request.Form("StyleName"),"样式名称")
		StyleContent = Request.Form("StyleContent")
	else
		Set RsStyleObj = CreateObject("ADODB.RecordSet")
		SQLStr = "Select * From FS_MallListStyle where ID=" & StyleID & ""
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
	StyleName = NoCSSHackAdmin(Request.Form("StyleName"),"样式名称")
	StyleContent = Request.Form("StyleContent")
end if
StyleContent = Replace(Replace(StyleContent,"""","%22"),"'","%27")
if Operation = "Modify" then SaveStyle StyleID
Sub SaveStyle(EditStyleID)
	Dim RsTemp,EditSql,RsCheckObj,CheckSql
	if Request.form("StyleName") = "" then
		Alert "请填写样式名称"
		Exit Sub
	else
		if EditStyleID = "" then
			CheckSql = "Select * from FS_MallListStyle where Name='" & Request.form("StyleName") & "'"
		else
			CheckSql = "Select * from FS_MallListStyle where Name='" & Request.form("StyleName") & "' and ID<>" & EditStyleID
		end if
		Set RsCheckObj = Conn.Execute(CheckSql)
		if Not RsCheckObj.Eof then
			Alert "标签名已经存在"
			Set RsCheckObj = Nothing
			Exit Sub
		end if
		Set RsCheckObj = Nothing
	end if
	'On Error Resume Next
	Set RsTemp = Server.CreateObject("ADODB.recordset")
	if EditStyleID = "" then
		EditSql = "Select * from FS_MallListStyle where 1=0"
		RsTemp.Open EditSql,Conn,3,3
		RsTemp.AddNew
	else
		EditSql = "Select * from FS_MallListStyle where ID=" & StyleID
		RsTemp.Open EditSql,Conn,3,3
		if RsTemp.Eof then Alert "修改的标签不存在"
	end if
	RsTemp("Name") = Request.Form("StyleName")
	RsTemp("Content") = Request.form("StyleContent")
	RsTemp.UpDate
	RsTemp.Close
	Set RsTemp = Nothing
	if err.Number=0 then
		Response.Redirect("Templet_MallStyleList.asp")
	else
		if StyleID <> "" then
			Alert "修改失败"
		else
			Alert "添加失败"
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
<title>添加和修改列表样式</title>
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
				  <td width=35 align="center" alt="保存" onClick="AddLableHead();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
				  <td width=2 class="Gray">|</td>
				  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
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
            <td width="60"> <div align="center">样式名称</div></td>
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
            <td width="60"> <div align="center">插入字段</div></td>
            <td> <select style="width:82%;" name="FieldList">
                <option value="" selected>选择插入的字段</option>
                <option value="{Mall_Product_Name}">╅ 商品名称</option>
                <option value="{Mall_Products_MemberPrice}">╅ 现在价格</option>
                <option value="{Mall_Products_OldPrice}">╅ 原始价格</option>
                <option value="{Mall_Products_MakeCompany}">╅ 制造厂家</option>
                <option value="{Mall_Products_Address}">╅ 产地</option>
                <option value="{Mall_Products_Stockpile}">╅ 库存状况</option>
                <option value="{Mall_Product_Type}">╅ 产品型号</option>
                <option value="{Mall_Products_MaintainRule}">╅ 保修条款</option>
                <option value="{Mall_Products_serial}">╅ 产品编号</option>
                <option value="{Mall_Products_weight}">╅ 商品重量</option>
                <option value="{Mall_Products_technic}">╅ 技术资料</option>
                <option value="{Mall_Products_description}">╅ 产品描述</option>
                <option value="{Mall_Products_AddTime}">╅ 产品添加时间</option>
                <option value="{Mall_Products_MakeTime}">╅ 产品制造时间</option>
                <option value="{Mall_Products_standard}">╅ 产品详细规格</option>
                <option value="{Mall_Products_Package}">╅ 包装清单</option>
                <option value="{Mall_Products_Keyword}">╅ 产品关键字</option>
                <option value="{Mall_Products_Picture}">╅ 产品图片_大图</option>
                <option value="{Mall_Products_sPicture}">╅ 产品图片_小图</option>
                <option value="{Mall_Favorite}">╅ 添加到收藏夹</option>
                <option value="{Mall_BuyOrder}">╅ 添加到购物车</option>
                <option value="{Mall_ClickNum}">╅ 点击率</option>
				<option value="{Mall_isSpecial}">╅ 是否特价</option>
				<option value="{Mall_SpecielName}">╅ 所属专区</option>
				<option value="{Mall_ComContent}">╅ 评论内容</option>
				<option value="{Mall_ComLinke}">╅ 评论连接</option>
            	<option value="{Mall_Special}">╅ 特别说明</option>
              </select> <input name="Submitfff" type="button" id="Submitfff" onClick="InsertField();" value="插入字段" style="color=#FF0000"> 
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
	if (frames["Editer"].CurrMode!='EDIT') {alert('其他模式下无法保存，请切换到编辑模式');return;}
	if (document.StyleForm.StyleName.value=='') ErrorCode=ErrorCode+'样式名称不能为空！\n';
	if (document.StyleForm.StyleContent.value=='') ErrorCode=ErrorCode+'样式内容不能为空！\n';
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
