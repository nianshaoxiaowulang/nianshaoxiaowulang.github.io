<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/ThumbnailFunction.asp" -->
<!--#include file="../Refresh/Function.asp" -->
<!--#include file="../Refresh/RefreshFunction.asp" -->
<!--#include file="../Refresh/SelectFunction.asp" -->
<%
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
response.buffer=true
Dim DBC,Conn,sRootDir
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
if SysRootDir<>"" then sRootDir="/"+SysRootDir else sRootDir=""
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
if Not JudgePopedomTF(Session("Name"),"" & Request("ClassID") & "") then Call ReturnError1()
if Request("NewsID") <> "" then
	if Not JudgePopedomTF(Session("Name"),"P010502") then Call ReturnError1()
else
	if Not JudgePopedomTF(Session("Name"),"P010501") then Call ReturnError1()
end if
Dim TempClassID,OldClassObj,OldClassEName,DummyPath_Riker,NewsExtFileName
Dim Action
Dim ITodayNewsTF,IBrowPop,IAddDate,IKeyWords,ITxtSource,IAuthor,IFilterNews,IAuditTF,ITitleSHowReview
Dim IEditer,IClickNum,ISpecialID,IPicPath,IShowReviewTF,IReviewTF,ISBSNews,IMarqueeNews,IProclaimNews,ILinkTF,IClassBuildNewsTemp
Dim IFocusNewsTF,IClassicalNewsTF,INewsTemplet,INaviWords,ITitleColor,ISavePic,IFileName,IFileExtName,IPath,IRecTF
Dim EditContentTF
Dim RsSelectObj,HaveValueTF
EditContentTF = False
Action = Request("Action")
IClassID = Request.Form("ClassID")
if IClassID="" then IClassID=Request("ClassID")
INewsID = Request("NewsID")
if INewsID = "" then
	EditContentTF = False
else
	EditContentTF = True
end if
If IClassID <> "" then
	TempClassID = Cstr(IClassID)
	TempClassID = Replace(Replace(Replace(Replace(Replace(TempClassID,"'",""),"and",""),"select",""),"or",""),"union","")
	Set OldClassObj = Conn.Execute("Select ClassID,ClassEName,NewsTemp,ClassCName,FileExtName from FS_NewsClass where ClassID='" & TempClassID & "'")
	if Not OldClassObj.Eof then
		NewsExtFileName=OldClassObj("FileExtName")
		OldClassEName = OldClassObj("ClassCName")
		IClassBuildNewsTemp = OldClassObj("NewsTemp")
	end if
	OldClassObj.Close
	Set OldClassObj = Nothing
else
	Response.Write("<script>alert(""�������ݴ���"");history.back();</script>")
	Response.End
End If
If SysRootDir<>"" then
	DummyPath_Riker = "/" & SysRootDir
Else
	DummyPath_Riker = ""
End If
if Action = "Submit" then
	HaveValueTF = False
else
	if INewsID <> "" Then
		INewsID = Replace(Replace(Replace(Replace(Replace(INewsID,"'",""),"and",""),"select",""),"or",""),"union","")
		Set RsSelectObj = Conn.Execute("Select * from FS_News where NewsID='" & INewsID & "'")
		if Not RsSelectObj.Eof then
			ITitle = RsSelectObj("Title")
			ISubTitle = RsSelectObj("SubTitle")
			ITitleColor = Left(RsSelectObj("Titlestyle"),7)
			TitleBoldstr = Mid(RsSelectObj("Titlestyle"),8,1)
			TitleUstr = Right(RsSelectObj("Titlestyle"),1)
			IContent = RsSelectObj("Content")
			ITodayNewsTF = RsSelectObj("TodayNewsTF")
			IManuTF = RsSelectObj("ManuTF")
			IFileName = RsSelectObj("FileName")
			IBrowPop = RsSelectObj("BrowPop")
			IFileExtName = RsSelectObj("FileExtName")
			IPath = RsSelectObj("Path")
			IAddDate = RsSelectObj("AddDate")
			IKeyWords = RsSelectObj("KeyWords")
			ITxtSource = RsSelectObj("TxtSource")
			IAuthor = RsSelectObj("Author")
			IEditer = RsSelectObj("Editer")
			IClickNum = RsSelectObj("ClickNum")
			IRecTF = RsSelectObj("RecTF")
			ISpecialID = RsSelectObj("SpecialID")
			IAuditTF = RsSelectObj("AuditTF")
			IShowReviewTF = RsSelectObj("ShowReviewTF")
			IReviewTF = RsSelectObj("ReviewTF")
			ISBSNews = RsSelectObj("SBSNews")
			IMarqueeNews = RsSelectObj("MarqueeNews")
			IProclaimNews = RsSelectObj("ProclaimNews")
			ILinkTF = RsSelectObj("LinkTF")
			IFilterNews = RsSelectObj("FilterNews")
			INewsTemplet = RsSelectObj("NewsTemplet")
			IPicPath = RsSelectObj("PicPath")
			IFocusNewsTF = RsSelectObj("FocusNewsTF")
			IClassicalNewsTF = RsSelectObj("ClassicalNewsTF")
			ITitleSHowReview=RsSelectObj("TitleSHowReview")
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
	ITitle = NoCSSHackAdmin(Request("Title"),"����")
	ISubTitle = Request("SubTitle")
	ITitleColor = Request("TitleColor")
	TitleBoldstr = Request("TitleBold")
	TitleUstr = Request("Titles")
	ISavePic = Request("SavePic")
	Dim TempForVar
	For TempForVar = 1 To Request.Form("Content").Count
		IContent = IContent & Request.Form("Content")(TempForVar)
	Next
	ITodayNewsTF = Request("TodayNewsTF")
	IManuTF = Request("ManuTF")
	IFileName = Request("FileName")
	IBrowPop = Request("BrowPop")
	IFileExtName = Request("FileExtName")
	IPath = Request("Path")
	IAddDate = Request("AddDate")
	if IAddDate = "" then IAddDate = Now()
	IKeyWords = Request("KeyWords")
	ITxtSource = Request("TxtSource")
	IAuthor = Request("Author")
	IEditer = Request("Editer")
	IClickNum = Request("ClickNum")
	IRecTF = Request("RecTF")
	ISpecialID = Request("SpecialID")
	IAuditTF = Request("AuditTF")
	IShowReviewTF = Request("ShowReviewTF")
	IReviewTF = Request("ReviewTF")
	ISBSNews = Request("SBSNews")
	IMarqueeNews = Request("MarqueeNews")
	IProclaimNews = Request("ProclaimNews")
	ILinkTF = Request("LinkTF")
	IFilterNews = Request("FilterNews")
	INewsTemplet = Request("NewsTemplet")
	IPicPath = Request("PicPath")
	IFocusNewsTF = Request("FocusNewsTF")
	IClassicalNewsTF = Request("ClassicalNewsTF")
	ITitleSHowReview=Request("TitleSHowReview")
	INaviWords=Request("NaviWords")
end if
if IsNull(IContent) then
	IContent = ""
else
	IContent = Replace(Replace(IContent,"""","%22"),"'","%27")
end if
if INewsTemplet = "" OR INewsTemplet = Null then
	INewsTemplet = IClassBuildNewsTemp
end if
if IFileExtName = "" OR IFileExtName = Null then
	IFileExtName = NewsExtFileName
end if

if Request.Form("Action")="Submit" then
	Dim INewsAddObj,INewsAddSql
	Dim NewsFileNames,RsNewsConfigObj
	if ITitle <> "" then
		ITitle = Replace(Replace(ITitle,"""",""),"'","")
	else
		Response.Write("<script>alert('���������ű���');history.back();</script>")
		Response.End
	end if
	if IClassID <> "" then
		IClassID = Replace(Replace(IClassID,"""",""),"'","")
	else
		Response.Write("<script>alert('��Ŀ�������ݴ���');history.back();</script>")
		Response.End
	end if
	if INewsTemplet <> "" then
		INewsTemplet = Replace(Replace(INewsTemplet,"""",""),"'","")
	else
		Response.Write("<script>alert('��ѡ������ģ���ļ�');history.back();</script>")
		Response.End
	end if 
	if Isnumeric(IClickNum) then
		IClickNum = Clng(IClickNum)
	else
		Response.Write("<script>alert('���ų�ʼ�����������Ϊ��������');history.back();</script>")
		Response.End
	end if 
	if IsDate(IAddDate) then
		IAddDate = Formatdatetime(IAddDate)
	else
		Response.Write("<script>alert('�������ʱ�����ʹ���,����������');history.back();</script>")
		Response.End
	end if
	if IPicPath = "" and request("ToWords")="" then
		Response.Write("<script>alert('������ͼƬ��ַ');history.back();</script>")
		Response.End
	end if
	if IContent = "" or IsNull(IContent) then
		Response.Write("<script>alert('��������������');history.back();</script>")
		Response.End
	end if
	Set RsNewsConfigObj = Conn.Execute("Select DoMain,NewsFileName,AutoClass,AutoIndex,ThumbnailComponent from FS_Config")
	if INewsID <> "" Then
		INewsID = Replace(Replace(Replace(Replace(Replace(INewsID,"'",""),"and",""),"select",""),"or",""),"union","")
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
		INewsAddObj("Path") =  "/" & year(now())&"-"&month(now())&"/"&day(now())             '����·�� 
	end if
	Dim INewsID,ITitle,ISubTitle,TitleBoldstr,TitleUstr,IClassID,IContent,IManuTF
	INewsAddObj("Title") = ITitle
	'If ISubTitle <> "" then
		INewsAddObj("SubTitle") = Replace(Replace(ISubTitle,"""",""),"'","")
	'end if
	If request("ToWords")<>"" then 
		INewsAddObj("PicNewsTF") =  0
	else
		INewsAddObj("PicNewsTF") =  1
	End If
	If TitleBoldstr <> "" then
		TitleBoldstr = "1"		
	else
		TitleBoldstr="0"		
	end if
	If TitleUstr <> "" then
		TitleUstr="1"		
	else
		TitleUstr="0"		
	end if
	INewsAddObj("Titlestyle") =  ITitleColor & TitleBoldstr & TitleUstr
	INewsAddObj("ClassID") =  IClassID
	INewsAddObj("HeadNewsTF") = "0"
	Dim Content_Loop_Var,Save_Content
	For Content_Loop_Var = 1 To Request.Form("Content").Count
		Save_Content = Save_Content & Request.Form("Content")(Content_Loop_Var)
	Next
	'===========================
	'�Զ���ҳ
	If instr(Save_Content,"[NoPage]") then
		Save_Content=replace(replace(Save_Content,"[Page]",""),"[NoPage]","")
	Else
		Save_Content=AutoSplitPages(Save_Content)
	End If
	'============================
	If ISavePic = "1" then
		CreateDateDir(Server.MapPath(DummyPath_Riker&"/"&UpFiles & "/" & BeyondPicDir))
		Save_Content = ReplaceRemoteUrl(Save_Content,"/" & UpFiles & "/" & BeyondPicDir&"/"&year(Now())&"-"&month(now())&"/"&day(Now()),RsNewsConfigObj("DoMain"),DummyPath_Riker)
	End If
	'============================
	'��������ͼ
	Dim AutoRefreshSmallPic,PicFileName,CreateSmallPicOK
	CreateSmallPicOK=False
	AutoRefreshSmallPic=RsNewsConfigObj("ThumbnailComponent")
	If AutoRefreshSmallPic="1" then
		PicFileName=mid(IPicPath,InstrRev(IPicPath,"/")+1)
		If left(IPicPath,4)="http" then
			SaveRemoteFile sRootDir&"/"&UpFiles&"/"&BeyondPicDir&"/sPic_"&PicFileName,IPicPath
			CreateSmallPicOK=CreateThumbnailEx(sRootDir&"/"&UpFiles&"/"&BeyondPicDir&"/sPic_"&PicFileName,sRootDir&"/"&UpFiles&"/"&BeyondPicDir&"/sPic_"&PicFileName)
			If CreateSmallPicOK=true then  IPicPath="/"&UpFiles&"/"&BeyondPicDir&"/sPic_"&PicFileName
		ElseIf Instr(IPicPath,"/sPic_")=0 then
			CreateSmallPicOK=CreateThumbnailEx(sRootDir&IPicPath,sRootDir&left(IPicPath,InstrRev(IPicPath,"/"))&"sPic_"&PicFileName)
			If CreateSmallPicOK=true then  IPicPath=left(IPicPath,InstrRev(IPicPath,"/"))&"sPic_"&PicFileName
		End If
	End If
	'============================
	INewsAddObj("PicPath") = IPicPath
	INewsAddObj("Content") = replace(Save_Content,WebDomain,"")   '�������� ��δ�ж�
	if IManuTF <> "" then
		INewsAddObj("ManuTF") = 1 
	else
		INewsAddObj("ManuTF") = 0
	end if
	If ITodayNewsTF <> "" then
		INewsAddObj("TodayNewsTF") = 1
	Else
		INewsAddObj("TodayNewsTF") = 0
	End If
	if IBrowPop <> "" then
		INewsAddObj("FileExtName") =  "asp"     '�����ļ���չ��
	else
		INewsAddObj("FileExtName") = IFileExtName     '�����ļ���չ��
	end if 
	INewsAddObj("AddDate") =  IAddDate
'=======================================================
'������Դ���ؼ��֡����ߡ����α༭
	if Request("ChkKeyword") = "SaveKeyWords" then 
		call SaveOption(Request("KeywordText"),1)
	End If
	if Request("ChkSource") = "SaveSource" then 
		call SaveOption(Request("TxtSourceText"),2)
	End If
	if Request("ChkAuthor") = "SaveAuthor" then 
		call SaveOption(Request("AuthorText"),3)
	End If
	if Request("ChkEditer") = "SaveEditer" then 
		call SaveOption(Request("EditerText"),4)
	End If
'==================================================
		INewsAddObj("KeyWords") = Replace(Replace(Request("KeywordText"),"""",""),"'","") 
	'end if
	'if Request("TxtSourceText") <> "" then
		INewsAddObj("TxtSource") = Replace(Replace(Request("TxtSourceText"),"""",""),"'","") 
	'end if
	'if Request("AuthorText") <> "" then
		INewsAddObj("Author") = Replace(Replace(Request("AuthorText"),"""",""),"'","")
	'end if
	'if Request("EditerText") <> "" then
		INewsAddObj("Editer") = Replace(Replace(Request("EditerText"),"""",""),"'","")
	'end if
	if IRecTF <> "" then
		INewsAddObj("RecTF") =  1
	else
		INewsAddObj("RecTF") =  0
	end if
	'if ISpecialID <> "" then
		INewsAddObj("SpecialID") = Replace(Replace(ISpecialID,"""",""),"'","")
	'end if
	if IAuditTF = "1" then
		INewsAddObj("AuditTF") =  1
	else
		INewsAddObj("AuditTF") = 0
	end if
	INewsAddObj("DelTF") = 0
	if IBrowPop <> "" then
		INewsAddObj("BrowPop") =  Replace(Replace(IBrowPop,"""",""),"'","")
	else
		INewsAddObj("BrowPop") =  0
	end if
	if IShowReviewTF<> "" then
		INewsAddObj("ShowReviewTF") = 1
	else
		INewsAddObj("ShowReviewTF") = 0
	end if
	if IReviewTF<> "" then
		INewsAddObj("ReviewTF") = 1
	else
		INewsAddObj("ReviewTF") = 0
	end if
	if ISBSNews <> "" then
		INewsAddObj("SBSNews") = 1
	else
		INewsAddObj("SBSNews") = 0
	end if
	If ITitleShowReview<>"" then 
		INewsAddObj("TitleShowReview")=1
	Else
		INewsAddObj("TitleShowReview")=0
	End If
	if IMarqueeNews <> "" then
		INewsAddObj("MarqueeNews") = 1
	else
		INewsAddObj("MarqueeNews") = 0
	end if
	if IProclaimNews <> "" then
		INewsAddObj("ProclaimNews") = 1
	else
		INewsAddObj("ProclaimNews") = 0
	end if
	if ILinkTF <> "" then
		INewsAddObj("LinkTF") = 1
	else
		INewsAddObj("LinkTF") = 0
	end if
	if IFilterNews <> "" then
		INewsAddObj("FilterNews") = 1
	Else
		INewsAddObj("FilterNews") = 0
	End if
	If IFocusNewsTF <> "" then
		INewsAddObj("FocusNewsTF") = 1
	Else
		INewsAddObj("FocusNewsTF") = 0
	End If
	If IClassicalNewsTF <> "" then
		INewsAddObj("ClassicalNewsTF") = 1
	Else
		INewsAddObj("ClassicalNewsTF") = 0
	End If
	INewsAddObj("NewsTemplet") =  INewsTemplet  
	INewsAddObj("NaviWords") = INaviWords
	if IClickNum <> "" then
		INewsAddObj("ClickNum") = IClickNum
	else
		INewsAddObj("ClickNum") = 0
	end if
	INewsAddObj.Update
	INewsAddObj.Close
	Set INewsAddObj = Nothing
	if IAuditTF = "1" then
		Dim CreatePageObj
		Set CreatePageObj = Conn.Execute("Select * from FS_News where NewsID='"&INewsID&"'")
		If Not CreatePageObj.eof then
			RefreshNews CreatePageObj
		Else
			Response.Write("<script>if (confirm(""ͼƬ������ӳɹ�,��δ�ܳɹ����������ļ�,�Ƿ�������?"")==false) {window.location='NewsList.asp?ClassID=" & IClassID & "';} else {window.location=""?ClassID="&IClassID&""";}</script>")
			Set RsNewsConfigObj = Nothing
			Response.End
		End If	
		CreatePageObj.Close
		Set CreatePageObj = Nothing  
	end if
	if EditContentTF = True then
		Response.Redirect("NewsList.asp?ClassID=" & IClassID)
	else
		If RsNewsConfigObj("AutoClass")="1" and RsNewsConfigObj("AutoIndex")="1" then
			Response.Write("<script>if (confirm(""ͼƬ������ӳɹ�,�Ƿ����ɴ���Ŀ����ҳ?"")==true) {window.location='../refresh/refreshauto.asp?ClassID=" & IClassID & "&AutoClass="&RsNewsConfigObj("AutoClass")&"&AutoIndex="&RsNewsConfigObj("AutoIndex")&"';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		ElseIf RsNewsConfigObj("AutoClass")="1" then
			Response.Write("<script>if (confirm(""ͼƬ������ӳɹ�,�Ƿ����ɴ���Ŀ?"")==true) {window.location='../refresh/refreshauto.asp?ClassID=" & IClassID & "&AutoClass="&RsNewsConfigObj("AutoClass")&"&AutoIndex="&RsNewsConfigObj("AutoIndex")&"';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		ElseIf RsNewsConfigObj("AutoIndex")="1" then
			Response.Write("<script>if (confirm(""ͼƬ������ӳɹ�,�Ƿ�������ҳ?"")==true) {window.location='../refresh/refreshauto.asp?ClassID=" & IClassID & "&AutoClass="&RsNewsConfigObj("AutoClass")&"&AutoIndex="&RsNewsConfigObj("AutoIndex")&"';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		Else
			Response.Write("<script>if (confirm(""ͼƬ������ӳɹ�,�Ƿ�������?"")==false) {window.location='NewsList.asp?ClassID=" & IClassID & "';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		End If
	end if
	Set RsNewsConfigObj = Nothing
	Response.End
end if

%>
<html>
<head>
<script language="JavaScript" type="text/JavaScript">
<!--
function insertPicAddress() { 
	if (document.NewsForm.ToWords.checked==true)
	{
		PicPathAddress.style.display='none';
		document.NewsForm.PicPath.value='';
	}
	else
		PicPathAddress.style.display='';
}
//-->
</script>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������</title>
</head>
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body topmargin="2" leftmargin="2">
<form action="" name="NewsForm" method="post">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="����" onClick="SubmitFun();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="�����������" onClick="location='NewsWords.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="��ӱ�������" onClick="location='NewsTitle.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="�������" onClick="location='DownLoad.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <%If sHaveValueTF = True then%>
		  <td width=35 align="center" alt="�����Ʒ" onClick="location='../mall/mall_addproducts.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">��Ʒ</td>
		  <td width=2 class="Gray">|</td>
		  <%End if%>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
          <td>&nbsp;<input type="hidden" name="Content" value="<% = IContent %>">
              <input type="hidden" name="Action" value="Submit"><input type="hidden" name="ClassID" value="<% = IClassID %>"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="30"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="60" height="30"> <div align="center">����</div></td>
            <td height="30"> <input style="width:60%;" type="text" name="Title" value="<% = ITitle %>">
              <input type="checkbox" name="TitleShowReview" value="1" title="����Ŀ���ű�������������2��"<%If ITitleShowReview="1" then response.write("Checked") End If%>>
              ��ʾ����&nbsp;&nbsp;&nbsp;&nbsp;
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
              <input type="checkbox" <% if TitleBoldstr = "1" then Response.Write("Checked") %> name="TitleBold" value="1">
              ���� 
              <input type="checkbox" <% if TitleUstr = "1" then Response.Write("Checked") %> name="Titles" value="1">
              б�� </td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">������</div></td>
            <td height="30"> <input style="width:70%;" type="text" name="SubTitle" value="<% = ISubTitle %>">
              &nbsp;&nbsp;&nbsp;&nbsp;
              <input name="ToWords" type="checkbox"  value="ToWords" onClick="insertPicAddress()">
              תΪ�������� </td>
          </tr>
		  <tr style="display:none;"><td><input name="IsPicNews" type="checkbox"  value="IsPicNews" checked=true></td></tr>
          <tr id="PicPathAddress"> 
            <td height="30"> 
              <div align="center">ͼƬ��ַ</div></td>
            <td> <input name="PicPath" type="text" id="PicPath2" style="width:74%" value="<% = IPicPath %>" > 
              &nbsp; <input type="button" name="Submit4" value="ѡ��ͼƬ" onClick="var TempReturnValue=OpenWindow('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',500,290,window);if (TempReturnValue!='') document.NewsForm.PicPath.value=TempReturnValue;"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td> <table width="100%" border="0" cellpadding="0" cellspacing="0" height="20">
          <tr> 
            <td width="60" height="26" align="center" bgcolor="#EFEFEF" class="LableSelected" id="ContentFolder" onClick="ChangeFolder(this);">��������</td>
            <td width="5" align="center" class="ToolBarButtonLine" style="cursor:default;">&nbsp;</td>
            <td onClick="ChangeFolder(this);" id="AttributeFolder" width="60" align="center" class="LableDefault">��������</td>
            <td class="ToolBarButtonLine" style="cursor:default;">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr style="display:none;" id="AttributeArea"> 
      <td height="30"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="ButtonListLeft">
          <tr> 
            <td width="86" height="30"> <div align="center">������Ŀ</div></td>
            <td colspan="3"> <input type="text" style="width:74%;" name="ClassCNameShow" readonly value="<% = OldClassEName %>"> 
              &nbsp; <input type="button" name="Submit" value="ѡ����Ŀ" onClick="SelectClass();"></td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">����ģ��</div></td>
            <td colspan="3"> <input name="NewsTemplet" type="text" id="NewsTemplet" readonly style="width:74%;" value="<% = INewsTemplet %>"> 
              &nbsp; <input type="button" name="Submit2" value="ѡ��ģ��" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<%=sRootDir %>/<% = TempletDir %>',400,300,window,document.NewsForm.NewsTemplet);document.NewsForm.NewsTemplet.focus();"> 
            </td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">�������</div></td>
            <td colspan="3"> <input name="AddDate" readonly type="text" id="AddDate" style="width:74%;" value="<% = IAddDate %>"> 
              &nbsp; <input type="button" name="Submit43" value="ѡ������" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,110,window,document.NewsForm.AddDate);document.NewsForm.AddDate.focus();"> 
            </td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">����ר��</div></td>
            <td colspan="3"> <%
			if Not IsNull(ISpecialID) And (ISpecialID <> "") then
				Dim RsSpecialObj,TempISpecialID,SpecialCNameText
				SpecialCNameText = ""
				TempISpecialID = Replace(Replace(Replace(Replace(Replace(ISpecialID,"'",""),"and",""),"select",""),"or",""),"union","")
				TempISpecialID = Replace(TempISpecialID,",","','")
				Set RsSpecialObj = Conn.Execute("Select * from FS_Special where SpecialID in ('" & TempISpecialID & "')")
				do while Not RsSpecialObj.Eof
					if SpecialCNameText = "" then
						SpecialCNameText = RsSpecialObj("CName")
					else
						SpecialCNameText = SpecialCNameText & "," & RsSpecialObj("CName")
					end if
					RsSpecialObj.MoveNext
				Loop
				Set RsSpecialObj = Nothing
			end if
			%> <input name="SpecialIDText" type="text" style="width:62%" readonly value="<% = SpecialCNameText %>"> 
              <input type="hidden" name="SpecialID" value="<% = ISpecialID %>"> 
              <select name="select" style="width:25%" onChange=ChooseSpecial(this.options[this.selectedIndex].value)>
                <option value="" selected> </option>
                <option value="<%=""&"***"&"Clean"%>" style="color:red">���</option>
                <%
				Dim SpecialIDObj
				set SpecialIDObj = Conn.Execute("Select CName,SpecialID from FS_Special order by ID desc")
				while not SpecialIDObj.eof
					%>
                <option value="<%=SpecialIDObj("CName")&"***"&SpecialIDObj("SpecialID")%>"><%=SpecialIDObj("CName")%></option>
                <%
					SpecialIDObj.Movenext
				Wend
				SpecialIDObj.Close
				Set SpecialIDObj = Nothing
				%>
              </select></td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">������Դ</div></td>
            <td colspan="3"> <input name="TxtSourceText" type="text" style="width:62%" value="<% = ITxtSource %>"> 
              <input type="hidden" name="TxtSource" value="<% = ITxtSource %>"> 
              <select name="select" style="width:25%" onChange="Dosusite(this.options[this.selectedIndex].value);">
                <option value="" selected> </option>
                <option value="<%=""&"***"&"Clean"%>" style="color:red">���</option>
                <option value="��վԭ��">��վԭ��</option>
                <option value="����">����</option>
                <%
		Dim TxtSourceObj
		set TxtSourceObj = Conn.Execute("select * from FS_Routine where Type=2 order by ID desc")
		While not TxtSourceObj.eof
		%>
                <option value="<%=TxtSourceObj("Name")&"***"&TxtSourceObj("Url")%>"><%=TxtSourceObj("Name")%></option>
                <%
		TxtSourceObj.Movenext
		Wend
		TxtSourceObj.Close
		Set TxtSourceObj = Nothing
		%>
              </select> <input name="ChkSource" type="checkbox" id="ChkSource" value="SaveSource">
              ���� 
              <div align="center"></div></td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">�� �� ��</div></td>
            <td colspan="3"> <input name="KeywordText" type="text" style="width:62%" value="<% = IKeyWords %>"> 
              <input type="hidden" name="KeyWords" value="<% = IKeyWords %>"> 
              <select name="select" style="width:25%" onChange=Dokesite(this.options[this.selectedIndex].value)>
                <option value="" selected> </option>
                <option value="Clean" style="color:red">���</option>
                <%
		Dim KeyWordsObj
		set KeyWordsObj = Conn.Execute("select * from FS_Routine where Type=1 order by ID desc")
		while not KeyWordsObj.eof
		%>
                <option value="<%=KeyWordsObj("Name")%>"><%=KeyWordsObj("Name")%></option>
                <%
		KeyWordsObj.Movenext
		Wend
		KeyWordsObj.Close
		Set KeyWordsObj = Nothing
		%>
              </select> <input name="ChkKeyWord" type="checkbox" id="ChkKeyWord" value="SaveKeyWords">
              ����</td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">��������</div></td>
            <td colspan="3"> <input name="AuthorText" type="text" style="width:62%" value="<% = IAuthor %>"> 
              <input type="hidden" name="Author" value="<% = IAuthor %>"> <select name="select" id="select8" style="width:25%" onChange="Doauthsite(this.options[this.selectedIndex].value);">
                <option value="" selected> </option>
                <option value="<%=""&"***"&"Clean"%>" style="color:red">���</option>
                <option value="����">����</option>
                <option value="��վ">��վ</option>
                <option value="δ֪">δ֪</option>
                <%
		Dim AuthorObj
		set AuthorObj = Conn.Execute("select * from FS_Routine where Type=3 order by ID desc")
		while not AuthorObj.eof
		%>
                <option value="<%=AuthorObj("Name")&"***"&AuthorObj("Url")%>"><%=AuthorObj("Name")%></option>
                <%
		AuthorObj.Movenext
		Wend
		AuthorObj.Close
		Set AuthorObj=nothing
		%>
              </select> <input name="ChkAuthor" type="checkbox" id="ChkAuthor" value="SaveAuthor">
              ���� </td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">���α༭</div></td>
            <td colspan="3"> <input name="EditerText" type="text" style="width:62%" value="<% = IEditer %>"> 
              <input type="hidden" name="Editer" value="<% = IEditer %>"> <select name="select" style="width:25%"  onChange="Editsite(this.options[this.selectedIndex].value);">
                <option value="" selected> </option>
                <option value="<%=""&"***"&"Clean"%>" style="color:red">���</option>
                <%
		Dim EditerObj
		Set EditerObj = Conn.Execute("Select * from FS_Routine where Type=4 order by ID desc")
		while not EditerObj.eof
		%>
                <option value="<%=EditerObj("Name")&"***"&EditerObj("Url")%>"><%=EditerObj("Name")%></option>
                <%
		EditerObj.Movenext
		Wend
		EditerObj.Close
		Set EditerObj = Nothing

		%>
              </select> <input name="ChkEditer" type="checkbox" id="ChkEditer" value="SaveEditer">
              ����</td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">���Ȩ��</div></td>
            <td colspan="3"> <select name="BrowPop" id="select7" style="width:100%" onChange="ChooseExeName();">
                <option value="" <%if IBrowPop = "" then Response.Write("selected")%>> 
                </option>
                <%
		Dim BrowPopObj
		set BrowPopObj = Conn.Execute("Select Name,PopLevel from FS_MemGroup order by PopLevel asc")
		while not BrowPopObj.eof
		%>
                <option value="<%=BrowPopObj("PopLevel")%>" <%if IBrowPop <> "" and IsNull(IBrowPop)=false then if Cint(IBrowPop) = Cint(BrowPopObj("PopLevel")) then Response.Write("selected") end if end if%>><%=BrowPopObj("Name")%></option>
                <%
		BrowPopObj.Movenext
		Wend
		BrowPopObj.Close
		Set BrowPopObj = Nothing
		%>
              </select></td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;��</div></td>
            <td width="458"> <input name="ClickNum" style="width:100%;" type="text" id="ClickNum3" size="10" value="<% if IClickNum = "" OR IsNull(IClickNum) then Response.Write("1") else Response.Write(IClickNum) %>"></td>
            <td width="108"><div align="center">��&nbsp;չ&nbsp;��</div></td>
            <td width="335"><select name="FileExtName" id="FileExtName" style="width:100%">
                <option <% if IFileExtName = "html" then Response.Write("Selected") %> value="html">html</option>
                <option <% if IFileExtName = "htm" then Response.Write("Selected") %> value="htm">htm</option>
                <option <% if IFileExtName = "shtm" then Response.Write("Selected") %> value="shtm">shtm</option>
                <option <% if IFileExtName = "shtml" then Response.Write("Selected") %> value="shtml">shtml</option>
                <option <% if IFileExtName = "asp" then Response.Write("Selected") %> value="asp">asp</option>
              </select></td>
          </tr>
          <tr> 
            <td> <div align="center">��������</div></td>
            <td height="30" colspan="3"><textarea name="NaviWords" rows="3" id="textarea2" style="width:100%"><% = INaviWords %></textarea></td>
          </tr>
          <tr> 
            <td rowspan="2"><div align="center">��ѡ����</div></td>
            <td height="30" colspan="3"> <input name="MarqueeNews" type="checkbox" id="MarqueeNews" value="1" <%if IMarqueeNews = 1 then Response.Write("checked") end if%>>
              ���� 
              <input name="ProclaimNews" type="checkbox" id="ProclaimNews2" value="1" <%if IProclaimNews = 1 then Response.Write("checked") end if%>>
              ���� 
              <input name="ClassicalNewsTF" type="checkbox" id="ClassicalNewsTF" value="1" <%if IClassicalNewsTF = 1 then Response.Write("checked") end if%>>
              ���� 
              <input name="TodayNewsTF" type="checkbox" id="TodayNewsTF2" value="1" <%if ITodayNewsTF = 1 then Response.Write("checked") end if%>>
              ͷ�� 
              <input name="ReviewTF" type="checkbox" id="ReviewTF2" value="1" onClick="ChooseRiview();" <%if IReviewTF = 1 then Response.Write("checked") end if%>>
              �������� 
              <input name="ShowReviewTF" type="checkbox" id="ShowReviewTF2" value="1" disabled <%if IShowReviewTF = 1 then Response.Write("checked") end if%>>
              ��ʾ���� 
              <input name="LinkTF" type="checkbox" id="LinkTF2" value="1" <%if ILinkTF = 1 then Response.Write("checked") end if%>>
              �ڲ����� </td>
          </tr>
          <tr> 
            <td height="30" colspan="3"><input name="SavePic" type="checkbox" id="SavePic2" value="1" <% if ISavePic = 1 then Response.Write("checked") end if %>>
              ��ͼ 
              <input name="FilterNews" type="checkbox" id="FilterNews2" value="1" <%if IFilterNews = 1 then Response.Write("checked") end if%>>
              �õ� 
              <input name="RecTF" type="checkbox" id="RecTF" value="1" <%if IRecTF = 1 then Response.Write("checked") end if%>>
              �Ƽ� 
              <input name="AuditTF" type="checkbox" id="AuditTF2" value="1" checked <%if IAuditTF = 1 then Response.Write("checked") end if%>>
              ��� 
              <input name="FocusNewsTF" type="checkbox" id="FocusNewsTF2" value="1" <%if IFocusNewsTF = 1 then Response.Write("checked") end if%>>
              ���� 
              <input name="SBSNews" type="checkbox" id="SBSNews" value="1" <%if ISBSNews = 1 then Response.Write("checked") end if%>>
              ���� 
              <input name="ManuTF" type="checkbox" id="ManuTF2" value="1" <%if IManuTF = 1 then Response.Write("checked") end if%>>
              Ͷ�� </td>
          </tr>
        </table></td>
    </tr>
    <tr id="ContentArea"> 
      <td colspan="3"><iframe id='NewsContent' src='../../Editer/NewsEditer.asp' frameborder=0 scrolling=no width='100%' height='410'></iframe></td> 
    </tr>
</table>
</form>
</body>
</html>
<script language="javascript">
function ChangeFolder(el)
{
	if (el.className=='LableSelected') return;
	var OperObj=null;
	var FolderIDArray=new Array('ContentFolder','AttributeFolder');
	var EditAreaIDArray=new Array('ContentArea','AttributeArea');
	el.className='LableSelected';
	el.bgColor='#EFEFEF';
	for (var i=0;i<FolderIDArray.length;i++)
	{
		OperObj=document.getElementById(FolderIDArray[i]);
		AreaObj=document.getElementById(EditAreaIDArray[i]);
		if (OperObj!=null)
		{
			if (OperObj.id!=el.id)
			{
				OperObj.className='LableDefault';
				OperObj.bgColor='#FFFFFF';
				if (AreaObj!=null) AreaObj.style.display='none';			
			}
			else if (AreaObj!=null) AreaObj.style.display='';
		}
	}
}
 function ChooseRiview()
   {
      if (document.NewsForm.ReviewTF.checked==true)
	      {
		    document.NewsForm.ShowReviewTF.disabled=false;
		   }
      else
	      {
	        document.NewsForm.ShowReviewTF.disabled=true;
		   }
	}
	
function ChooseExeName()
{
  if (document.NewsForm.BrowPop.value!='') document.NewsForm.FileExtName.disabled=true;
  else document.NewsForm.FileExtName.disabled=false;
 }

function SubmitFun()
{
	if (frames["NewsContent"].CurrMode!='EDIT') {alert('����ģʽ���޷����棬���л������ģʽ');return;}
	frames["NewsContent"].SaveCurrPage();
	var TempContentArray=frames["NewsContent"].NewsContentArray;
	document.NewsForm.Content.value='';
	for (var i=0;i<TempContentArray.length;i++)
	{
		if (TempContentArray[i]!='')
		{
			if (document.NewsForm.Content.value=='') document.NewsForm.Content.value=TempContentArray[i];
			else document.NewsForm.Content.value=document.NewsForm.Content.value+'[Page]'+TempContentArray[i];
		} 
	}
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
ChooseExeName();
</script>
