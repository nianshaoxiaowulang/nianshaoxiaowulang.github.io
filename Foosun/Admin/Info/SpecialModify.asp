<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P020200") then Call ReturnError1()
Dim SpecialID,SpModObj
If Request("SpecialID")="" then
   Response.Write("<script>alert(""�������ݴ���"");window.close();</script>")
   Response.End
 else
	 SpecialID = Cstr(Request("SpecialID"))
	 Set SpModObj = Conn.Execute("Select * From FS_Special where SpecialID='"&SpecialID&"'")
	 if SpModObj.eof then
	   Response.Write("<script>alert(""�������ݴ���"");window.close();</script>")
	   Response.End
	 end if
 end if
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Ƶ��/ר���޸�</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2">
<form action="" name="SpecialFrom" method="post">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="document.SpecialFrom.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp; <input name="action" type="hidden" id="action" value="mod"> 
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#E0E0E0">
    <tr bgcolor="#FFFFFF"> 
      <td width="100"> 
        <div align="right">��������</div></td>
      <td> 
        <input name="CName" type="text" id="CName" style="width:100%" value="<%=SpModObj("CName")%>">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="right">Ӣ������</div></td>
      <td> 
        <input name="EName" type="text" id="EName" style="width:100%" disabled value="<%=SpModObj("EName")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="right">����ͼƬ</div></td>
      <td> 
        <input name="NaviPic" type="text" id="NaviPic" size="68" value="<%=SpModObj("NaviPic")%>"> 
        <input type="button" name="Submit" value="ѡ��ͼƬ" onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,290,window,document.SpecialFrom.NaviPic);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="right">Ƶ��/ר��ģ��</div></td>
      <td> 
        <input name="Templet" type="text" readonly id="Templet" size="68" value="<%=SpModObj("Templet")%>"> 
        <input type="button" name="Submit2" value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<%=sRootDir %>/<% = TempletDir %>',400,300,window,document.SpecialFrom.Templet);document.SpecialFrom.Templet.focus();"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="right">����·��</div></td>
      <td> 
        <input name="SaveFilePath" readonly type="text" size="68" value="<%=SpModObj("SaveFilePath")%>"> 
        <input type="button" name="Submit5" value="ѡ��·��" onClick="OpenWindowAndSetValue('../../FunPages/SelectPathFrame.asp?CurrPath=<%=sRootDir %>/<% = ClassDir %>',400,300,window,document.SpecialFrom.SaveFilePath);document.SpecialFrom.SaveFilePath.focus();"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="right">����ʱ��</div></td>
      <td> 
        <input name="AddTime" type="text" id="AddTime" value="<%=SpModObj("AddTime")%>" size="68"> 
        <input name="sdaf" type="button" id="sdaf" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,120,window,document.SpecialFrom.AddTime);document.SpecialFrom.AddTime.focus();" value="ѡ������"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="right">��ҳ˵��</div></td>
      <td> 
        <textarea name="IndexNaviWord" rows="6" id="IndexNaviWord" style="width:100%;" ><%=SpModObj("IndexNaviWord")%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" style="display:none;"> 
      <td> 
        <div align="right">��Ŀ˵��</div></td>
      <td> 
        <input name="ClassNaviWord" type="text" id="ClassNaviWord" style="width:100%" value="<%=SpModObj("ClassNaviWord")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="right">����ͼƬ</div></td>
      <td> 
        <input name="MorePic" type="text" id="MorePic" style="width:100%" value="<%=SpModObj("MorePic")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="right">��չ��</div></td>
      <td> 
        <select name="FileExtName" style="width:100%;">
          <option value="htm" <%if SpModObj("FileExtName")="htm" then Response.Write("selected")%>>htm</option>
          <option value="html" <%if SpModObj("FileExtName")="html" then Response.Write("selected")%>>html</option>
          <option value="shtm" <%if SpModObj("FileExtName")="shtm" then Response.Write("selected")%>>shtm</option>
          <option value="shtml" <%if SpModObj("FileExtName")="shtml" then Response.Write("selected")%>>shtml</option>
          <option value="asp" <%if SpModObj("FileExtName")="asp" then Response.Write("selected")%>>asp</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="right">������ʾ</div></td>
      <td> 
        <input name="ShowNaviTF" type="checkbox" id="ShowNaviTF2" value="1" <%if SpModObj("ShowNaviTF")="1" then Response.Write("checked") end if%>></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
  if Request("action")="mod" then
     Dim SpecialObj,SpecialSql,SpCName,SpEName,TempObj,SpAddDate
	 if Request.Form("CName") <> "" then
	 	SpCName = Replace(Replace(Request.Form("CName"),"""",""),"'","")
		if Len(SpCName)>=25 then
			Response.Write("<script>alert(""ר���������Ʋ��ܳ���25���ַ�"");</script>")
			Response.End
		end if
	 else
	    Response.Write("<script>alert(""������ר����������"");</script>")
		Response.End
	 end if
	 if Request.Form("Templet")="" then
	    Response.Write("<script>alert(""��ѡ��ר��ģ��"");</script>")
		Response.End
	 end if
	 if Request.Form("SaveFilePath")="" or isnull(Request.Form("SaveFilePath")) then
	    Response.Write("<script>alert(""��ѡ���ļ�����·��"");</script>")
		Response.End
	 End if
	 if Request.Form("FileExtName")="" or isnull(Request.Form("FileExtName")) then
	    Response.Write("<script>alert(""��ѡ���ļ���չ��"");</script>")
		Response.End
	 End If
     if isdate(Request.Form("AddTime")) then
		 SpAddDate = Formatdatetime(Request.Form("AddTime"))
	 else
	    Response.Write("<script>alert(""ר�����ʱ�����ʹ���"");</script>")
		Response.End
	 end if
	  Set SpecialObj=server.createobject(G_FS_RS)
	  SpecialSql="select * from FS_Special where SpecialID='"&SpecialID&"'"
	  SpecialObj.open SpecialSql,Conn,1,3
	  SpecialObj("CName") = SpCName
	  if Request.Form("NaviPic")<>"" then
		  SpecialObj("NaviPic") = Request.Form("NaviPic")
	  end if
	  if Request.Form("IndexNaviWord")<>"" then
		  SpecialObj("IndexNaviWord") = Request.Form("IndexNaviWord")
	  end if
	  if Request.Form("ClassNaviWord")<>"" then
		  SpecialObj("ClassNaviWord") = Request.Form("ClassNaviWord")
	  end if
	  if Request.Form("MorePic")<>"" then
		  SpecialObj("MorePic") = Request.Form("MorePic")
	  end if
	  SpecialObj("Templet") = Request.Form("Templet")
	  if Request.Form("ShowNaviTF") = "1" then
		  SpecialObj("ShowNaviTF") = "1"
	  else
		  SpecialObj("ShowNaviTF") = "0"
	  end if 
	  SpecialObj("SaveFilePath") = Request.Form("SaveFilePath")
	  SpecialObj("FileExtName") = Request.Form("FileExtName")
	  SpecialObj("AddTime") = SpAddDate
	  SpecialObj.update
	  SpecialObj.Close
	  Set SpecialObj = Nothing
		%>
		<script>
			top.GetNavFoldersObject().location='../Menu_Folders.asp?Action=Special';		
		</script>
		<%
  end if
%>