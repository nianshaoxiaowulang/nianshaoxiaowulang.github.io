<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../Inc/Cls_Info.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
Dim DBC,Conn,sRootDir
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
if SysRootDir<>"" then sRootDir="/"+SysRootDir else sRootDir=""
'==============================================================================
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System(FoosunCMS V3.1.0930)
'���¸��£�2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'��ҵע����ϵ��028-85098980-601,��Ŀ������028-85098980-606��609,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��394226379,159410,125114015
'����֧��QQ��315485710,66252421 
'��Ŀ����QQ��415637671��655071
'���򿪷����Ĵ���Ѷ�Ƽ���չ���޹�˾(Foosun Inc.)
'Email:service@Foosun.cn
'MSN��skoolls@hotmail.com
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.cn  ��ʾվ�㣺test.cooin.com 
'��վͨϵ��(���ܿ��ٽ�վϵ��)��www.ewebs.cn
'==============================================================================
'��Ѱ汾���ڳ�����ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'��Ѷ��˾�����˳���ķ���׷��Ȩ��
'�������2�ο��������뾭����Ѷ��˾������������׷����������
'==============================================================================
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),""&Request("ClassID")&"") then Call ReturnError1()
if Not JudgePopedomTF(Session("Name"),"P010200") then Call ReturnError1()
Dim ClassID,Sql,RsClassEditObj,ClassEName,ClassCName,ClassTemp,Contribution,AddTime,ParentID,ParentIDStr,HaveContTF,BrowPop,IsOutClass
Dim SaveFilePath,FileExtName,ShowTF,NewsTemp,DownLoadTemp,ProductTemp,DoMain,DoMainDisabledStr,SelectPathBtnDisabledStr,TempParentID,FileTime,Orders,RedirectList
Dim CheckRootClassNumber  '������Ŀ��������ֹ��ѭ��
Dim StrInClass,StrOutClass,StrClassLink
CheckRootClassNumber = 30
SelectPathBtnDisabledStr = ""
ClassID = Request("ClassID")
Dim RsMenuConfigObj,HaveValueTF
Set RsMenuConfigObj = Conn.execute("Select IsShop From FS_Config")
if RsMenuConfigObj("IsShop") = 1 then
	HaveValueTF = True
Else
	HaveValueTF = False
End if
Set RsMenuConfigObj = Nothing
if ClassID <> "" then
	Sql = "Select * from FS_NewsClass where ClassID='" & ClassID & "' and DelFlag=0"
	Set RsClassEditObj = Conn.Execute(Sql)
	if RsClassEditObj.Eof then
		Set RsClassEditObj = Nothing
		Set Conn = Nothing
		Alert "��Ŀ�Ѿ���ɾ��"
	else
		Dim RsTempObj,SelectPath
		ClassEName = RsClassEditObj("ClassEName")
		ClassCName = RsClassEditObj("ClassCName")
		ClassTemp = RsClassEditObj("ClassTemp")
		NewsTemp = RsClassEditObj("NewsTemp")
		DownLoadTemp = RsClassEditObj("DownLoadTemp")
		ProductTemp =RsClassEditObj("ProductTemp")
		Contribution = RsClassEditObj("Contribution")
		ShowTF = RsClassEditObj("ShowTF")
		AddTime = RsClassEditObj("AddTime")
		ParentID = RsClassEditObj("ParentID")
		SaveFilePath = RsClassEditObj("SaveFilePath")
		FileExtName = RsClassEditObj("FileExtName")
		BrowPop = RsClassEditObj("BrowPop")
		DoMain = RsClassEditObj("DoMain")
		FileTime = RsClassEditObj("FileTime")
		Orders = RsClassEditObj("Orders")
		IsOutClass=RsClassEditObj("IsOutClass")
		StrClassLink=RsClassEditObj("ClassLink")
		RedirectList = RsClassEditObj("RedirectList")
		If IsOutClass="1" then 
			StrInClass="style=""display:none;"""
		Else
			StrOutClass="style=""display:none;"""
		End If
		if ParentID <> "0" then
			Set RsTempObj = Conn.Execute("Select ClassCName,ParentID,DoMain,DelFlag,SaveFilePath,RedirectList from FS_NewsClass where ClassID='" & ParentID & "'")
			if RsTempObj.Eof then
				Set RsTempObj = Nothing
				Set RsClassEditObj = Nothing
				Alert "����Ŀ������"
				Response.End
			else
				if RsTempObj("DelFlag") = 1 then
					Set RsTempObj = Nothing
					Set RsClassEditObj = Nothing
					Alert "����Ŀ�ڻ���վ"
					Response.End
				else
					Dim CheckRootClassIndex
					CheckRootClassIndex = 1
					ParentIDStr = RsTempObj("ClassCName")
					TempParentID = RsTempObj("ParentID")
					do while Not (RsTempObj("ParentID") = "0")
						CheckRootClassIndex = CheckRootClassIndex + 1
						RsTempObj.Close
						Set RsTempObj = Nothing
						Set RsTempObj = Conn.Execute("Select ClassCName,ParentID,Domain,SaveFilePath,RedirectList from FS_NewsClass where ClassID='" & TempParentID & "'")
						if RsTempObj.Eof then
							Set RsTempObj = Nothing
							Set RsClassEditObj = Nothing
							Alert "����Ŀ������"
							Response.End
						end if
						TempParentID = RsTempObj("ParentID")
						if CheckRootClassIndex > CheckRootClassNumber then TempParentID = "0" '��ֹ��ѭ��
					Loop
					DoMain = RsTempObj("DoMain")
					if (Not IsNull(DoMain)) And (DoMain <> "") then
						SelectPath = RsTempObj("SaveFilePath")
						SelectPathBtnDisabledStr = " disabled"
					else
						SelectPath =sRootDir & "/" & ClassDir
					end if
				end if
			end if
			Set RsTempObj = Nothing
		else
			ParentIDStr = "����Ŀ"
			SelectPath =sRootDir & "/" & ClassDir
		end if
		if Contribution = 1 then
			Set RsTempObj = Conn.Execute("Select ContID from FS_Contribution where ClassID='" & ClassID & "'")
			if Not RsTempobj.Eof then
				HaveContTF = True
			else
				HaveContTF = False
			end if
			Set RsTempObj = Nothing
		else
			HaveContTF = False
		end if
	end if
else
	Alert "�������ݴ���"
end if
if ParentID <> "0" then
	DoMainDisabledStr = " disabled"
end if

Dim Result
Result = Request.Form("Result")
if Result = "Submit" then
	Dim CClass,ReturnCheckInfo,ReturnValueArray
	Set CClass = New InfoClass
	CClass.TForm = Request.Form
	ReturnCheckInfo = CClass.AddAndModifyClass()
	Set CClass = Nothing
	ReturnValueArray = Split(ReturnCheckInfo,"||")
	if ReturnValueArray(0) = "Success" then
		%>
		<script>
			top.GetNavFoldersObject().location='../Menu_Folders.asp?Action=ContentTree&OpenClassIDList=<% = ReturnValueArray(1) %>';		
		</script>
		<%
	else
		%>
			<script>alert('<% = ReturnCheckInfo %>');history.back();</script>
		<%
	end if
	response.end
ElseIf Result="OutClass" then
	Dim RsEditClass
	Set RsEditClass=Server.CreateObject(G_FS_RS)
	RsEditClass.open "Select ClassCName,ClassLink,ShowTf,Orders,ParentID from FS_NewsClass where ClassID='"&ClassID &"'",Conn,3,3
	if Request.Form("OutOrders") <> "" then
		if IsNumeric(Request.Form("OutOrders")) then RsEditClass("Orders") = Request.Form("OutOrders")
	end if
	if Request.Form("ShowTF") = "1" then
		RsEditClass("ShowTF") = 1
	else
		RsEditClass("ShowTF") = 0
	end If
	RsEditClass("ClassCName")=NoCSSHackAdmin(request.Form("ClassName"),"��Ŀ����")
	RsEditClass("ClassLink")=NoCSSHackAdmin(Request.Form("ClassLink"),"���ӵ�ַ")
	RsEditClass("ParentID")=0
	RsEditClass.UpDate
	RsEditClass.Close
	Set RsEditClass=Nothing
	If err=0 then 
		%>
		<script>
			top.GetNavFoldersObject().location='../Menu_Folders.asp?Action=ContentTree&OpenClassIDList=<% = ClassID %>';		
		</script>
		<%
	else
		%>
		<script>alert('<% = err.description %>');history.back();</script>
		<%
	end if
	response.end
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ŀ�޸�</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body scroll=no topmargin="2" leftmargin="2">
<div <%=StrInClass%>>
<form action="" method="post" name="ClassForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
              <td width="35" align="center" alt="����" onClick="InOK();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
			  <td width=2 class="Gray">|</td>
			  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
            <td>&nbsp; <input name="Result" type="hidden" id="Result" value="Submit"> 
              <input name="ClassID" value="<% = ClassID %>" type="hidden" id="ClassID2"> 
              <input name="ParentID" value="<% = ParentID %>" type="hidden" id="ParentID2"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
    <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#E3E3E3">
      <tr bgcolor="#FFFFFF"> 
        <td width="100" height="26"> 
          <div align="center">��������</div></td>
        <td> 
          <div align="center"> 
            <input value="<% = ClassCName %>" name="ClassCName" type="text" id="ClassCName2" style="width:100%;">
          </div></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">Ӣ������</div></td>
        <td> 
          <div align="center"> 
            <input value="<% = ClassEName %>" readonly name="ClassEName" type="text" id="ClassEName" style="width:100%;">
          </div></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">����Ŀ</div></td>
        <td> 
          <div align="center"> 
            <input readonly value="<% = ParentIDStr %>" style="width:100%;" type="text" name="textfield3">
          </div></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">��������</div></td>
        <td nowrap> 
          <input <% = DoMainDisabledStr %> type="text" name="DoMain" style="width:100%;" value="<% = DoMain %>"></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">���Ȩ��</div></td>
        <td nowrap> 
          <select name="BrowPop" id="select3" style="width:100%" onChange="CheckFileExtName(this);">
            <option value="" <%if BrowPop = 0 then Response.Write("selected")%>> 
            </option>
            <%
		Dim BrowPopObj
		set BrowPopObj = Conn.Execute("Select Name,PopLevel from FS_MemGroup order by PopLevel asc")
		while not BrowPopObj.eof
		%>
            <option value="<%=BrowPopObj("PopLevel")%>" <%if BrowPop <> "" And IsNull(BrowPop) = False then if BrowPop = Cint(BrowPopObj("PopLevel")) then Response.Write("selected") end if end if%>><%=BrowPopObj("Name")%></option>
            <%
			BrowPopObj.Movenext
		Wend
		BrowPopObj.Close
		Set BrowPopObj = Nothing
		%>
          </select></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">��չ��</div></td>
        <td nowrap> 
          <select name="FileExtName" style="width:100%;">
            <option value="htm" <%if FileExtName = "htm" then Response.Write("selected")%>>htm</option>
            <option value="html" <%if FileExtName = "html" then Response.Write("selected")%>>html</option>
            <option value="shtml" <%if FileExtName = "shtml" then Response.Write("selected")%>>shtml</option>
            <option value="asp" <%if FileExtName = "asp" then Response.Write("selected")%>>asp</option>
          </select></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">��Ŀģ��</div></td>
        <td nowrap> 
          <div align="left"> 
            <input readonly value="<% = ClassTemp %>" type="text" style="width:78%;" name="ClassTemp">
            <input type="button" name="Submit" value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<%=sRootDir %>/<% = TempletDir %>',400,300,window,document.ClassForm.ClassTemp);document.ClassForm.ClassTemp.focus();">
          </div></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">��������ģ��</div></td>
        <td nowrap> 
          <input readonly value="<% = NewsTemp %>" type="text" style="width:78%;" name="NewsTemp"> 
          <input type="button" name="Submit2" value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<%=sRootDir %>/<% = TempletDir %>',400,300,window,document.ClassForm.NewsTemp);document.ClassForm.NewsTemp.focus();"></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">��������ģ��</div></td>
        <td nowrap> 
          <input type="text" style="width:78%;" name="DownLoadTemp" value="<% = DownLoadTemp %>" readonly> 
          <input type="button" name="Submit2" value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<%=sRootDir %>/<% = TempletDir %>',400,300,window,document.ClassForm.DownLoadTemp);document.ClassForm.DownLoadTemp.focus();"></td>
      </tr>
      <tr bgcolor="#FFFFFF" <%if HaveValueTF = False then response.Write("style=""display:none""")%>> 
        <td height="26"> 
          <div align="center">������Ʒģ��</div></td>
        <td nowrap> 
          <input type="text" style="width:78%;" name="ProductTemp" value="<% = ProductTemp %>" readonly> 
          <input type="button" name="Submit2" value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<%=sRootDir %>/<% = TempletDir %>',400,300,window,document.ClassForm.ProductTemp);document.ClassForm.ProductTemp.focus();"></td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">����·��</div></td>
        <td nowrap> 
          <input readonly type="text" style="width:78%;" value="<% = RsClassEditObj("SaveFilePath") %>" name="SaveFilePath"> 
          <input type="button" name="Submit4" value="ѡ��·��" onClick="OpenWindowAndSetValue('../../FunPages/SelectPathFrame.asp?CurrPath=<%=sRootDir %>/<% = ClassDir %>',400,300,window,document.ClassForm.SaveFilePath);document.ClassForm.SaveFilePath.focus();"> 
        </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">�������</div></td>
        <td nowrap> 
          <input readonly value="<% = AddTime %>" type="text" style="width:78%;" name="AddTime"> 
          <input type="button" name="Submit3" value="ѡ������" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,120,window,document.ClassForm.AddTime);document.ClassForm.AddTime.focus();"> 
        </td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">���Ӳ���</div></td>
        <td nowrap> 
          <input onClick="CheckHaveContTF(this);" name="Contribution" haveconttf="<% = HaveContTF %>" type="checkbox" id="Contribution2" value="1" <% if Contribution = 1 then Response.Write("checked") %>>
          ����Ͷ�� &nbsp;&nbsp;&nbsp;&nbsp; <input name="ShowTF" type="checkbox" id="ShowTF2" value="1" <% if ShowTF = 1 then Response.Write("checked") %>>
          ǰ̨��ʾ&nbsp;&nbsp;&nbsp;&nbsp; ��Ŀ���� 
          <input name="Orders" type="text" size="5" maxlength="4" value="<% = Orders %>">
          ���Ź鵵ʱ�䣨���֣� 
          <input name="FileTime" type="text" value="<% if FileTime = "" then Response.Write("100") else Response.Write(FileTime)%>" size="5" maxlength="3">
          �� ��
</td>
      </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">Ĭ��ת��</div></td>
        <td nowrap> 
          <select name="RedirectList" id="RedirectList" style="width:100%;">
            <option value="1" <% if RedirectList = 1 then  Response.Write("Selected")%>>�����б�</option>
            <option value="2" <% if RedirectList = 2 then  Response.Write("Selected")%>>�����б�</option>
            <%if HaveValueTF = True then%>
            <option value="3" <% if RedirectList = 3 then  Response.Write("Selected")%>>��Ʒ�б�</option>
            <%End if%>
          </select></td>
      </tr>
    </table>
</form>
</div>
<div dwcopytype="CopyTableRow" id="OutClass" <%=StrOutClass%>>
<form action="" method="post" name="OutClassForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		      <td width=35 align="center" alt="����" onClick="OutOK();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
    <table width="100%"  border="0" cellpadding="3" cellspacing="1" bgcolor="#E3E3E3">
      <tr bgcolor="#FFFFFF"> 
        <td width="100" height="26"> 
          <div align="center">��Ŀ����</div></td>
        <td> 
          <input name="ClassName" type="text" id="Name" style="width:100%" value="<% = ClassCName %>"></td>
    </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">���ӵ�ַ</div></td>
        <td> 
          <input name="ClassLink" type="text" id="Link" style="width:100%" value="<%=StrClassLink%>"></td>
    </tr>
      <tr bgcolor="#FFFFFF"> 
        <td height="26"> 
          <div align="center">���Ӳ���</div></td>
        <td> 
          <input name="ClassID" value="<% = ClassID %>" type="hidden" id="ClassID2"> 
        <input name="ShowTF" type="checkbox" id="ShowTF2" value="1" <% if ShowTF = 1 then Response.Write("checked") %>>
        ǰ̨��ʾ &nbsp;&nbsp;&nbsp;&nbsp;��Ŀ���� <input name="Result" type="hidden" id="Result" value="OutClass"> 
        <input name="OutOrders" type="text" size="5" maxlength="4" value="<%=Orders%>">
      </td>
    </tr>
</table>
</form>
</div>
</body>
</html>
<%
Set Conn = Nothing
Set RsClassEditObj = Nothing
Set Conn = Nothing
Function Alert(InfoStr)
%>
<script language="JavaScript">
alert('<% = InfoStr %>');
history.back();
window.close();
</script>
<%
End Function
%>
<script language="JavaScript">
function CheckHaveContTF(Obj)
{
	if (Obj.HaveContTF=='True')
	{
		alert('����Ŀ����Ͷ�壬�����޸Ĵ�����');
		Obj.checked=true;
	}
}
</script>
<script language="JavaScript">
CheckFileExtName(document.all.BrowPop);
function InOK()
{
	if (CheckEnglishStr(document.ClassForm.ClassEName,'Ӣ������')==true)
	{
		document.ClassForm.submit();
	}
}
function OutOK()
{
	if (document.OutClassForm.ClassName.value=='') {alert('��д��Ŀ��������');document.OutClassForm.ClassName.focus();return;}
	if (document.OutClassForm.ClassLink.value=='') {alert('��д��Ŀ���ӵ�ַ');document.OutClassForm.ClassLink.focus();return;}
	document.OutClassForm.submit();
}
function CheckFileExtName(Obj)
{
	if (Obj.value!='')
	{
		for (var i=0;i<document.all.FileExtName.length;i++)
		{
			if (document.all.FileExtName.options(i).value=='asp') document.all.FileExtName.options(i).selected=true;
		}
		document.all.FileExtName.disabled=true;
	}
	else
	{
		document.all.FileExtName.disabled=false;
	}
}
</script>