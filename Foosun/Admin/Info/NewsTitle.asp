<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Refresh/Function.asp" -->
<!--#include file="../Refresh/RefreshFunction.asp" -->
<!--#include file="../Refresh/SelectFunction.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
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
Dim RsMenuConfigObj,sHaveValueTF
Set RsMenuConfigObj = Conn.execute("Select IsShop From FS_Config")
if RsMenuConfigObj("IsShop") = 1 then
	sHaveValueTF = True
Else
	sHaveValueTF = False
End if
if Not JudgePopedomTF(Session("Name"),""&Request("ClassID")&"") then Call ReturnError1()
if Request("NewsID") <> "" then
	if Not JudgePopedomTF(Session("Name"),"P010502") then Call ReturnError1()
else
	if Not JudgePopedomTF(Session("Name"),"P010501") then Call ReturnError1()
end if
Dim TempClassID,OldClassObj,OldClassEName
Dim INewsID,ITitle,ISubTitle,IClassID,TitleBoldstr,TitleUstr,ITitleColor,ITodayNewsTF,IAddDate,IRecTF
Dim IAuditTF,IShowReviewTF,IReviewTF,ISBSNews,IMarqueeNews
Dim IProclaimNews,ILinkTF,IFilterNews,IPicNewsTF,IHeadNewsPath,INaviWords,IPicPath
Dim EditContentTF,Action
EditContentTF = False
Action = Request("Action")
IClassID = Request.Form("ClassID")
IClassID = Replace(Replace(Replace(Replace(Replace(IClassID,"'",""),"and",""),"select",""),"or",""),"union","")
If IClassID="" then IClassID=Request("ClassID")
INewsID = Request("NewsID")
INewsID = Replace(Replace(Replace(Replace(Replace(INewsID,"'",""),"and",""),"select",""),"or",""),"union","")
if INewsID = "" then
	EditContentTF = False
else
	EditContentTF = True
end if
If IClassID <> "" then
	TempClassID = Cstr(IClassID)
	Set OldClassObj = Conn.Execute("Select ClassID,ClassEName,ClassCName from FS_NewsClass where ClassID='" & TempClassID & "'")
	if Not OldClassObj.Eof then
		OldClassEName = OldClassObj("ClassCName")
	end if
	OldClassObj.Close
	Set OldClassObj = Nothing
else
	Response.Write("<script>alert(""�������ݴ���"");history.back();</script>")
	Response.End
End If
Dim RsSelectObj,HaveValueTF
if Action = "Submit" then
	HaveValueTF = False
else
	if INewsID <> "" Then
		Set RsSelectObj = Conn.Execute("Select * from FS_News where NewsID='" & INewsID & "'")
		if Not RsSelectObj.Eof then
			ITitle = RsSelectObj("Title")
			ISubTitle = RsSelectObj("SubTitle")
			ITitleColor = Left(RsSelectObj("Titlestyle"),7)
			TitleBoldstr = Mid(RsSelectObj("Titlestyle"),8,1)
			TitleUstr = Right(RsSelectObj("Titlestyle"),1)
			IPicNewsTF = RsSelectObj("PicNewsTF")
			ITodayNewsTF = RsSelectObj("TodayNewsTF")
			IAddDate = RsSelectObj("AddDate")
			IRecTF = RsSelectObj("RecTF")
			IAuditTF = RsSelectObj("AuditTF")
			IShowReviewTF = RsSelectObj("ShowReviewTF")
			IReviewTF = RsSelectObj("ReviewTF")
			ISBSNews = RsSelectObj("SBSNews")
			IMarqueeNews = RsSelectObj("MarqueeNews")
			IProclaimNews = RsSelectObj("ProclaimNews")
			ILinkTF = RsSelectObj("LinkTF")
			IFilterNews = RsSelectObj("FilterNews")
			INaviWords = RsSelectObj("NaviWords")
			IHeadNewsPath = RsSelectObj("HeadNewsPath")
			IPicPath = RsSelectObj("PicPath")
			HaveValueTF = True
		else
			HaveValueTF = False
		end if
		Set RsSelectObj = Nothing
	else
		HaveValueTF = False
	end if
end if
if HaveValueTF = False then
	ITitle = NoCSSHackAdmin(Request("Title"),"���ű���")
	ISubTitle = Request("SubTitle")
	ITitleColor = Request("TitleColor")
	TitleBoldstr = Request("TitleBold")
	TitleUstr = Request("Titles")
	IPicNewsTF = Request("PicNewsTF")
	ITodayNewsTF = Request("TodayNewsTF")
	IAddDate = Request("AddDate")
	if IAddDate = "" then IAddDate = Now()
	IRecTF = Request("RecTF")
	IAuditTF = Request("AuditTF")
	IShowReviewTF = Request("ShowReviewTF")
	IReviewTF = Request("ReviewTF")
	ISBSNews = Request("SBSNews")
	IMarqueeNews = Request("MarqueeNews")
	IProclaimNews = Request("ProclaimNews")
	ILinkTF = Request("LinkTF")
	IFilterNews = Request("FilterNews")
	INaviWords = Request("NaviWords")
	IHeadNewsPath = Request("HeadNewsPath")
	IPicPath = Request("PicPath")
end if
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������</title>
</head>
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body  topmargin="2" leftmargin="2">
<form action="" name="NewsForm" method="post">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="document.NewsForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="�����������" onClick="location='NewsWords.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="���ͼƬ����" onClick="location='NewsPic.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ͼƬ</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="�������" onClick="location='DownLoad.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <%If sHaveValueTF = True then%>
		  <td width=35 align="center" alt="�����Ʒ" onClick="location='../mall/mall_addproducts.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">��Ʒ</td>
		  <td width=2 class="Gray">|</td>
		  <%End if%>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp; <input name="action" type="hidden" id="action" value="Submit"><input type="hidden" name="ClassID" value="<% = IClassID %>"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
    <tr> 
      <td width="100" height="30"> 
        <div align="center">���ű���</div></td>
      <td> 
        <input name="Title" type="text" id="Title" style="width:60%;" value="<% = ITitle %>"> 
        <select name="TitleColor" id="select2">
			<option <% if ITitleColor = "#UUUUUU" then Response.Write("Selected")%> value="#UUUUUU" selected>������ɫ</option>
			<option <% if ITitleColor = "#ff0000" then Response.Write("Selected")%> value="#ff0000" style="background-color:#ff0000;color: #ffffff">#ff0000</option>
			<option <% if ITitleColor = "#000000" then Response.Write("Selected")%> value="#000000" style="background-color:#000000;color: #ffffff">#000000</option>
			<option <% if ITitleColor = "#FFFFFF" then Response.Write("Selected")%> value="#FFFFFF" style="background-color:#ffffff;color: #000000">#FFFFFF</option>
			<option <% if ITitleColor = "#000099" then Response.Write("Selected")%> value="#000099" style="background-color:#000099;color: #ffffff">#000099</option>
			<option <% if ITitleColor = "#660066" then Response.Write("Selected")%> value="#660066" style="background-color:#660066;color: #ffffff">#660066</option>
			<option <% if ITitleColor = "#FF6600" then Response.Write("Selected")%> value="#FF6600" style="background-color:#FF6600;color: #ffffff">#FF6600</option>
			<option <% if ITitleColor = "#666666" then Response.Write("Selected")%> value="#666666" style="background-color:#666666;color: #ffffff">#666666</option>
			<option <% if ITitleColor = "#009900" then Response.Write("Selected")%> value="#009900" style="background-color:#009900;color: #ffffff">#009900</option>
			<option <% if ITitleColor = "#0066CC" then Response.Write("Selected")%> value="#0066CC" style="background-color:#0066CC;color: #ffffff">#0066CC</option>
			<option <% if ITitleColor = "#990000" then Response.Write("Selected")%> value="#990000" style="background-color:#990000;color: #ffffff">#990000</option>
			<option <% if ITitleColor = "#CC9900" then Response.Write("Selected")%> value="#CC9900" style="background-color:#CC9900;color: #ffffff">#CC9900</option>
			<option <% if ITitleColor = "#CCCCCC" then Response.Write("Selected")%> value="#CCCCCC" style="background-color:#CCCCCC;color: #000000">#CCCCCC</option>
			<option <% if ITitleColor = "#99FF00" then Response.Write("Selected")%> value="#99FF00" style="background-color:#99FF00;color: #000000">#99FF00</option>
			<option <% if ITitleColor = "#0000FF" then Response.Write("Selected")%> value="#0000FF" style="background-color:#0000FF;color: #FFFFFF">#0000FF</option>
			<option <% if ITitleColor = "#9966CCU" then Response.Write("Selected")%> value="#9966CC" style="background-color:#9966CC;color: #FFFFFF">#9966CC</option>
        </select>
        <input name="TitleBold" <% if TitleBoldstr = "1" then Response.Write("Checked") %> type="checkbox" id="TitleBold2" value="1">
        �Ӵ� 
        <input name="Titles" <% if TitleUstr = "1" then Response.Write("Checked") %> type="checkbox" id="Titles" value="1">
        б�� </td>
    </tr>
    <tr> 
      <td height="30"> 
        <div align="center">������Ŀ</div></td>
      <td> 
        <input type="text" style="width:74%;" name="ClassCNameShow" readonly value="<% = OldClassEName %>"> 
              &nbsp; <input type="button" name="Submit" value="ѡ����Ŀ" onClick="SelectClass();"></td>
    </tr>
    <tr> 
      <td> 
        <div align="center">��������</div></td>
      <td> 
        <textarea name="NaviWords" rows="6" id="NaviWords" style="width:100%"><% = INaviWords %></textarea></td>
    </tr>
    <tr> 
      <td height="30"> 
        <div align="center">���ӵ�ַ</div></td>
      <td> 
        <input name="HeadNewsPath" type="text" id="HeadNewsPath3" style="width:100%" value="<%if IHeadNewsPath = "" or IHeadNewsPath="http://" then Response.Write("http://") else Response.Write(IHeadNewsPath) end if%>"></td>
    </tr>
    <tr> 
      <td height="30"> 
        <div align="center">ͼƬ��ַ</div></td>
      <td> 
        <input name="PicPath" type="text" readonly id="PicPath" style="width:74% " value="<% = IPicPath %>"> 
        &nbsp; <input type="button" name="PicChooseButton" value="ѡ��ͼƬ" disabled onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,290,window,document.NewsForm.PicPath);"></td>
    </tr>
    <tr> 
      <td height="30"> 
        <div align="center">�������</div></td>
      <td> 
        <input name="AddDate" readonly type="text" id="AddDate" style="width:74% " value="<% if IAddDate = "" then Response.Write(now()) else Response.Write(IAddDate) end if%>"> 
        &nbsp; <input type="button" name="Submit42" value="ѡ������" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,110,window,document.NewsForm.AddDate);document.NewsForm.AddDate.focus();"></td>
    </tr>
    <tr> 
      <td height="26"> 
        <div align="center">��ѡ����</div></td>
      <td> 
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="10%" height="30">
<div align="center"> 
                <input <% if IPicNewsTF = "1" then Response.Write("Checked")%> name="PicNewsTF" type="checkbox" id="PicNewsTF2" value="1" onClick="ChoosePic();">
                ͼƬ����</div></td>
            <td width="10%"><div align="center"> 
                <input name="MarqueeNews" type="checkbox" id="MarqueeNews" value="1" <%if IMarqueeNews = "1" then Response.Write("checked") %>>
                ��������</div></td>
            <td width="14%"><div align="center"> 
                <input name="ReviewTF" type="checkbox" id="ReviewTF" value="1" onClick="ChooseRiview();" <%if IReviewTF = "1" then Response.Write("checked") %>>
                �������� </div></td>
            <td width="11%"><div align="center"> 
                <input name="ShowReviewTF" type="checkbox" id="ShowReviewTF" value="1" disabled <%if IShowReviewTF = "1" then Response.Write("checked") %>>
                ��ʾ���� </div></td>
            <td width="9%"><div align="center"> 
                <input name="ProclaimNews" type="checkbox" id="ProclaimNews" value="1" <%if IProclaimNews = "1" then Response.Write("checked")%>>
                �������� </div></td>
          </tr>
          <tr> 
            <td height="30">
<div align="center"> 
                <input name="RecTF" type="checkbox" id="RecTF4" value="1" <%if IRecTF = "1" then Response.Write("checked") %>>
                �Ƽ�����</div></td>
            <td><div align="center"> 
                <input name="AuditTF" type="checkbox" id="AuditTF3" value="1" checked <%if IAuditTF = "1" then Response.Write("checked")%>>
                ͨ����� </div></td>
            <td><div align="center"> 
                <input name="SBSNews" type="checkbox" id="SBSNews3" value="1" <%if ISBSNews = "1" then Response.Write("checked")%>>
                ��������</div></td>
            <td><div align="center"> 
                <input name="TodayNewsTF" type="checkbox" id="TodayNewsTF" value="1" <%if ITodayNewsTF = "1" then Response.Write("checked")%>>
                ����ͷ��</div></td>
            <td><div align="center"> 
                <input name="FilterNews" type="checkbox" disabled id="FilterNews3" value="1" <%if IFilterNews = "1" then Response.Write("checked")%>>
                �õ�����</div></td>
          </tr>
        </table></td>
    </tr>
</table>
</form>
</body>
</html>
<script language="javascript">
function ChooseRiview()
{
	if (document.NewsForm.ReviewTF.checked==true) document.NewsForm.ShowReviewTF.disabled=false;
	else document.NewsForm.ShowReviewTF.disabled=true;
}
	
function ChoosePic()
{
	if (document.NewsForm.PicNewsTF.checked==true)
	{
		document.NewsForm.PicChooseButton.disabled=false;
		document.NewsForm.NaviWords.disabled=false;
		document.NewsForm.FilterNews.disabled=false;
	}
	else
	{
		document.NewsForm.PicChooseButton.disabled=true;
		document.NewsForm.NaviWords.disabled=true;
		document.NewsForm.FilterNews.disabled=true;
	}
}
	
function SubmitFun()
{
	document.NewsForm.submit();
}
function SelectClass()
{
	var ReturnValue='',TempArray=new Array();
	ReturnValue = OpenWindow('../../FunPages/SelectClassFrame.asp',400,300,window);
	if (ReturnValue.indexOf('***')!=-1)
	{
		TempArray = ReturnValue.split('***');
		document.all.ClassID.value=TempArray[0]
		document.all.ClassCNameShow.value=TempArray[1]
	}
}
ChooseRiview();
ChoosePic();
</script>
<%
if Action = "Submit" then
	Dim INewsAddObj,INewsAddSql,NewsFileNames,RsNewsConfigObj
	if ITitle <> "" then
		ITitle = Replace(Replace(ITitle,"""",""),"'","")
	else
		Response.Write("<script>alert('���������ű���');</script>")
		Response.End
	end if
	if IClassID <> "" then
		IClassID = Replace(Replace(IClassID,"""",""),"'","")
	else
		Response.Write("<script>alert('��Ŀ�������ݴ���');</script>")
		Response.End
	end if
	if IsDate(IAddDate) then
		IAddDate = Formatdatetime(IAddDate)
	else
		Response.Write("<script>alert('�������ʱ�����ʹ���,����������');</script>")
		Response.End
	end if
	if LCase(IHeadNewsPath) = "http://" then
		Response.Write("<script>alert('�������ӵ�ַ����Ϊ��');</script>")
		Response.End
	end if
	Set RsNewsConfigObj = Conn.Execute("Select NewsFileName,AutoClass,AutoIndex from FS_Config")
	if INewsID <> "" then
		Set INewsAddObj = Server.CreateObject(G_FS_RS)
		INewsAddSql = "select * from FS_News where NewsID='" & INewsID & "'"
		INewsAddObj.open INewsAddSql,Conn,3,3
	else
		INewsID = GetRandomID18()
		Set INewsAddObj = Server.CreateObject(G_FS_RS)
		INewsAddSql = "select * from FS_News where 1=0"
		INewsAddObj.open INewsAddSql,Conn,3,3
		INewsAddObj.AddNew
		INewsAddObj("NewsID") = INewsID    '����ID
		NewsFileNames = NewsFileName(RsNewsConfigObj("NewsFileName"),IClassID,INewsID)
		INewsAddObj("FileName") = NewsFileNames   '�����ļ��� 
	end if
	INewsAddObj("Title") =  ITitle
	If ISubTitle <> "" then
		INewsAddObj("SubTitle") = Replace(Replace(ISubTitle,"""",""),"'","")
	end if
	If TitleBoldstr <> "" then
		TitleBoldstr = "1"		
	else
		TitleBoldstr = "0"		
	end if
	If TitleUstr <> "" then
		TitleUstr = "1"		
	else
		TitleUstr = "0"		
	end if
	INewsAddObj("Titlestyle") =  ITitleColor & TitleBoldstr & TitleUstr
	INewsAddObj("ClassID") =  IClassID
	INewsAddObj("HeadNewsTF") = 1
	if ITodayNewsTF <> "" then
		INewsAddObj("TodayNewsTF") = 1
	else
		INewsAddObj("TodayNewsTF") = 0
	end if
	INewsAddObj("FileExtName") = "html"     '�����ļ���չ��
	INewsAddObj("Path") =  "/" & year(now())&"-"&month(now())&"/"&day(now())             '����·�� 
	INewsAddObj("AddDate") =  IAddDate
	if IRecTF = "1" then
		INewsAddObj("RecTF") = 1
	else
		INewsAddObj("RecTF") = 0
	end if
	if IAuditTF = "1" then
		INewsAddObj("AuditTF") = 1
	else
		INewsAddObj("AuditTF") = 0
	end if
	INewsAddObj("DelTF") =  "0"
	if IShowReviewTF = "1" then
		INewsAddObj("ShowReviewTF") = 1
	else
		INewsAddObj("ShowReviewTF") = 0
	end if
	if IReviewTF = "1" then
		INewsAddObj("ReviewTF") = 1
	else
		INewsAddObj("ReviewTF") = 0
	end if
	if ISBSNews = "1" then
		INewsAddObj("SBSNews") = 1
	else
		INewsAddObj("SBSNews") = 0
	end if
	if IMarqueeNews = "1" then
		INewsAddObj("MarqueeNews") = 1
	else
		INewsAddObj("MarqueeNews") = 0
	end if
	if IProclaimNews = "1" then
		INewsAddObj("ProclaimNews") = 1
	else
		INewsAddObj("ProclaimNews") = 0
	end if
	if ILinkTF = "1" then
		INewsAddObj("LinkTF") = 1
	else
		INewsAddObj("LinkTF") = 0
	end if
	if IFilterNews = "1" then
		INewsAddObj("FilterNews") = 1
	else
		INewsAddObj("FilterNews") = 0
	end if
	if IPicNewsTF = "1" then
		INewsAddObj("PicNewsTF") = 1
	else
		INewsAddObj("PicNewsTF") = 0
	end if
	INewsAddObj("HeadNewsPath") =  IHeadNewsPath
	INewsAddObj("NaviWords") =  GotTopic(INaviWords,100)
	INewsAddObj("PicPath") =  IPicPath
	INewsAddObj.Update
	INewsAddObj.Close
	Set INewsAddObj = Nothing
	if EditContentTF = True then
		Response.Redirect("NewsList.asp?ClassID=" & IClassID)
	else
		If RsNewsConfigObj("AutoClass")="1" and RsNewsConfigObj("AutoIndex")="1" then
			Response.Write("<script>if (confirm(""����������ӳɹ�,�Ƿ����ɴ���Ŀ����ҳ?"")==true) {window.location='../refresh/refreshauto.asp?ClassID=" & IClassID & "&AutoClass="&RsNewsConfigObj("AutoClass")&"&AutoIndex="&RsNewsConfigObj("AutoIndex")&"';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		ElseIf RsNewsConfigObj("AutoClass")="1" then
			Response.Write("<script>if (confirm(""����������ӳɹ�,�Ƿ����ɴ���Ŀ?"")==true) {window.location='../refresh/refreshauto.asp?ClassID=" & IClassID & "&AutoClass="&RsNewsConfigObj("AutoClass")&"&AutoIndex="&RsNewsConfigObj("AutoIndex")&"';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		ElseIf RsNewsConfigObj("AutoIndex")="1" then
			Response.Write("<script>if (confirm(""����������ӳɹ�,�Ƿ�������ҳ?"")==true) {window.location='../refresh/refreshauto.asp?ClassID=" & IClassID & "&AutoClass="&RsNewsConfigObj("AutoClass")&"&AutoIndex="&RsNewsConfigObj("AutoIndex")&"';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		Else
			Response.Write("<script>if (confirm(""����������ӳɹ�,�Ƿ�������?"")==false) {window.location='NewsList.asp?ClassID=" & IClassID & "';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		End If
	end if
	Set RsNewsConfigObj = Nothing
	Response.End
end if
%>