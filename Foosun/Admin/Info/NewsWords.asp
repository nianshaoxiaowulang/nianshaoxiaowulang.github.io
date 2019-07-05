<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Refresh/Function.asp" -->
<!--#include file="../Refresh/RefreshFunction.asp" -->
<!--#include file="../Refresh/SelectFunction.asp" -->
<%
response.buffer=true
Dim DBC,Conn,sRootDir
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
if SysRootDir<>"" then sRootDir="/"+SysRootDir else sRootDir=""
'response.write sRootDir
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
Dim TempClassID,OldClassObj,OldClassEName,DummyPath_Riker,NewsExtFileName
Dim Action
Dim INewsID,ITitle,ISubTitle,ITitleColor,TitleBoldstr,TitleUstr,IClassID,IHeadNewsTF,IPicNewsTF,ISavePic,IClassBuildNewsTemp
Dim IContent,ITodayNewsTF,IManuTF,IFileName,IBrowPop,IFileExtName,IPath,IAddDate,IKeyWords,ITxtSource,IAuthor,ITitleSHowReview
Dim IEditer,IClickNum,IRecTF,ISpecialID,IAuditTF,IDelTF,IShowReviewTF,IReviewTF,ISBSNews,IMarqueeNews,IProclaimNews,ILinkTF,IFilterNews,INewsTemplet,iPicPath,IPicNews
Dim EditContentTF
EditContentTF = False
Action = Request("Action")
IClassID = Request.Form("ClassID")
if IClassID="" then IClassID=Request("ClassID")
INewsID = Request("NewsID")
INewsID = Replace(Replace(Replace(Replace(Replace(INewsID,"'",""),"and",""),"select",""),"or",""),"union","")
IClassID = Replace(Replace(Replace(Replace(Replace(IClassID,"'",""),"and",""),"select",""),"or",""),"union","")
if INewsID = "" then
	EditContentTF = False
else
	EditContentTF = True
end if
If IClassID <> "" then
	TempClassID = Cstr(IClassID)
	Set OldClassObj = Conn.Execute("select ClassID,ClassEName,ClassCName,NewsTemp,FileExtName from FS_NewsClass where ClassID='" & TempClassID & "'")
	if Not OldClassObj.Eof then
		NewsExtFileName=OldClassObj("FileExtName")
		OldClassEName = OldClassObj("ClassCName")
		IClassBuildNewsTemp = OldClassObj("NewsTemp")
	end if
	OldClassObj.Close
	Set OldClassObj = Nothing
else
	Response.Write("<script>alert(""参数传递错误"");history.back();</script>")
	Response.End
End If
If SysRootDir<>"" then
	DummyPath_Riker = "/" & SysRootDir
Else
	DummyPath_Riker = ""
End If
Dim RsSelectObj,HaveValueTF
if Action = "Submit" then
	HaveValueTF = False
else
	if INewsID <> "" then
		Set RsSelectObj = Conn.Execute("Select * from FS_News where NewsID='" & INewsID & "'")
		if Not RsSelectObj.Eof then
			ITitle = RsSelectObj("Title")
			ISubTitle = RsSelectObj("SubTitle")
			ITitleColor = Left(RsSelectObj("Titlestyle"),7)
			TitleBoldstr = Mid(RsSelectObj("Titlestyle"),8,1)
			TitleUstr = Right(RsSelectObj("Titlestyle"),1)
			IHeadNewsTF = RsSelectObj("HeadNewsTF")
			IPicNewsTF = RsSelectObj("PicNewsTF")
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
			IDelTF = RsSelectObj("DelTF")
			IShowReviewTF = RsSelectObj("ShowReviewTF")
			IReviewTF = RsSelectObj("ReviewTF")
			ISBSNews = RsSelectObj("SBSNews")
			IMarqueeNews = RsSelectObj("MarqueeNews")
			IProclaimNews = RsSelectObj("ProclaimNews")
			ILinkTF = RsSelectObj("LinkTF")
			IFilterNews = RsSelectObj("FilterNews")
			INewsTemplet = RsSelectObj("NewsTemplet")
			ITitleSHowReview= RsSelectObj("TitleSHowReview")
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
	ITitle = NoCSSHackAdmin(Request("Title"),"标题")
	ISubTitle = Request("SubTitle")
	ITitleColor = Request("TitleColor")
	TitleBoldstr = Request("TitleBold")
	TitleUstr = Request("Titles")
	IHeadNewsTF = Request("HeadNewsTF")
	IPicNewsTF = Request("PicNewsTF")
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
	IDelTF = Request("DelTF")
	IShowReviewTF = Request("ShowReviewTF")
	IReviewTF = Request("ReviewTF")
	ISBSNews = Request("SBSNews")
	IMarqueeNews = Request("MarqueeNews")
	IProclaimNews = Request("ProclaimNews")
	ILinkTF = Request("LinkTF")
	IFilterNews = Request("FilterNews")
	INewsTemplet = Request("NewsTemplet")
	IPicNews=Request("IsPicNews")
	IPicPath=Request("PicPath")
	ITitleSHowReview=Request("TitleSHowReview")
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


if Action = "Submit" then
	Dim INewsAddObj,INewsAddSql,NewsFileNames,RsNewsConfigObj
	If IPicNews="IsPicNews" then 
		If IPicPath="" then 
			response.write ("<script>alert(""请选择图片地址"");history.back();</script>")
			Response.End
		End If
	End If
	if ITitle <> "" then
		ITitle = Replace(Replace(ITitle,"""",""),"'","")
	else
		Response.Write("<script>alert(""请输入新闻标题"");history.back();</script>")
		Response.End
	end if
	if IClassID <> "" then
		IClassID = Replace(Replace(IClassID,"""",""),"'","")
	else
		Response.Write("<script>alert(""栏目参数传递错误"");history.back();</script>")
		Response.End
	end if
	if INewsTemplet<>"" then
		INewsTemplet = Replace(Replace(INewsTemplet,"""",""),"'","")
	else
		Response.Write("<script>alert(""请选择新闻模板文件"");history.back();</script>")
		Response.End
	end if 
	if Isnumeric(IClickNum) then
		IClickNum = Clng(IClickNum)
	else
		Response.Write("<script>alert(""新闻初始点击次数必须为数字类型"");history.back();</script>")
		Response.End
	end if 
	if IsDate(IAddDate) then
		IAddDate = Formatdatetime(IAddDate)
	else
		Response.Write("<script>alert(""新闻添加时间类型错误,请重新输入"");history.back();</script>")
		Response.End
	end if
	if IContent = "" OR IsNull(IContent) then
		Response.Write("<script>alert(""新闻内容不能为空"");history.back();</script>")
		Response.End
	end if
	Set RsNewsConfigObj = Conn.Execute("Select DoMain,NewsFileName,AutoClass,AutoIndex from FS_Config")
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
		INewsAddObj("NewsID") = INewsID    '新闻ID
		NewsFileNames = NewsFileName(RsNewsConfigObj("NewsFileName"),IClassID,INewsID)
		INewsAddObj("FileName") = NewsFileNames  '新闻文件名
		INewsAddObj("Path") =  "/" & year(now())&"-"&month(now())&"/"&day(now())             '新闻路径 
	end if
	If IPicNews<>"" then 
		INewsAddObj("PicNewsTF")=1
		INewsAddObj("PicPath")=iPicPath
	else
		INewsAddObj("PicNewsTF") = 0
	end if
	
	INewsAddObj("Title") =  ITitle
	'If ISubTitle <> "" then
		INewsAddObj("SubTitle") = ISubTitle
	'end if
	If TitleBoldstr <> "" then
		TitleBoldstr="1"		
	else
		TitleBoldstr="0"		
	end if
	If TitleUstr <> "" then
		TitleUstr = "1"		
	else
		TitleUstr = "0"		
	end if
	INewsAddObj("Titlestyle") =  ITitleColor & TitleBoldstr & TitleUstr
	INewsAddObj("ClassID") =  IClassID
	INewsAddObj("HeadNewsTF") = 0
	Dim Content_Loop_Var,Save_Content
	For Content_Loop_Var = 1 To Request.Form("Content").Count
		Save_Content = Save_Content & Request.Form("Content")(Content_Loop_Var)
	Next
	If instr(Save_Content,"[NoPage]") then
		Save_Content=replace(replace(Save_Content,"[Page]",""),"[NoPage]","")
	Else
		Save_Content=AutoSplitPages(Save_Content)
	End If
	If ISavePic = "1" then
		CreateDateDir(Server.MapPath(DummyPath_Riker&"/"&UpFiles & "/" & BeyondPicDir))
		Save_Content = ReplaceRemoteUrl(Save_Content,"/" & UpFiles & "/" & BeyondPicDir&"/"&year(Now())&"-"&month(now())&"/"&day(Now()),RsNewsConfigObj("DoMain"),DummyPath_Riker)
	End If
	INewsAddObj("Content") = replace(Save_Content,WebDomain,"")
	If ITitleShowReview<>"" then 
		INewsAddObj("TitleShowReview")=1
	Else
		INewsAddObj("TitleShowReview")=0
	End If
	If ITodayNewsTF <> "" then
		INewsAddObj("TodayNewsTF") = 1
	Else
		INewsAddObj("TodayNewsTF") = 0
	End If
	if IManuTF = "1" then
		INewsAddObj("ManuTF") =  1
	else
		INewsAddObj("ManuTF") = 0
	end if 
	if IBrowPop <> "" then
		INewsAddObj("FileExtName") =  "asp"     '新闻文件扩展名
	else
		INewsAddObj("FileExtName") =  IFileExtName     '新闻文件扩展名
	end if 
	INewsAddObj("AddDate") =  IAddDate
'=======================================================
'保存来源、关键字、作者、责任编辑
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
	INewsAddObj("TxtSource") = Replace(Replace(Request("TxtSourceText"),"""",""),"'","") 
	INewsAddObj("Author") = Replace(Replace(Request("AuthorText"),"""",""),"'","")
	INewsAddObj("Editer") = Replace(Replace(Request("EditerText"),"""",""),"'","")
	INewsAddObj("ClickNum") =  IClickNum
	if IRecTF = "1" then
		INewsAddObj("RecTF") =  1
	else
		INewsAddObj("RecTF") =  0
	end if
	INewsAddObj("SpecialID") = Replace(Replace(ISpecialID,"""",""),"'","")
	if IAuditTF = "1" then
		INewsAddObj("AuditTF") = 1
	else
		INewsAddObj("AuditTF") = 0
	end if
	INewsAddObj("DelTF") = 0
	if IBrowPop <> "" then
		INewsAddObj("BrowPop") =  Replace(Replace(IBrowPop,"""",""),"'","")
	else
		INewsAddObj("BrowPop") =  0
	end if
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
	INewsAddObj("FilterNews") = 0
	INewsAddObj("NewsTemplet") =  INewsTemplet
	INewsAddObj.Update
	INewsAddObj.Close
	Set INewsAddObj = Nothing
	if IAuditTF = "1" then
		Dim CreatePageObj
		Set CreatePageObj = Conn.Execute("Select * from FS_News where NewsID='" & INewsID & "'")
		If Not CreatePageObj.eof then
			RefreshNews CreatePageObj
		Else
			CreatePageObj.Close
			Set CreatePageObj = Nothing
			RsNewsConfigObj.Close
			Set RsNewsConfigObj = Nothing
			if EditContentTF = False then
				Response.Write("<script>if (confirm(""文字新闻添加成功,但未能成功生成新闻文件"")==false) {window.location='NewsList.asp?ClassID="&IClassID&"';} else {window.location=""?ClassID=" & IClassID & """;}</script>")
			else
				Response.Redirect("NewsList.asp?ClassID=" & IClassID & "")
			end if
			Response.End
		End If
		CreatePageObj.Close
		Set CreatePageObj = Nothing
	end if 
	if EditContentTF = False then
		If RsNewsConfigObj("AutoClass")="1" and RsNewsConfigObj("AutoIndex")="1" then
			Response.Write("<script>if (confirm(""文字新闻添加成功,是否生成此栏目和首页?"")==true) {window.location='../refresh/refreshauto.asp?ClassID=" & IClassID & "&AutoClass="&RsNewsConfigObj("AutoClass")&"&AutoIndex="&RsNewsConfigObj("AutoIndex")&"';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		ElseIf RsNewsConfigObj("AutoClass")="1" then
			Response.Write("<script>if (confirm(""文字新闻添加成功,是否生成此栏目?"")==true) {window.location='../refresh/refreshauto.asp?ClassID=" & IClassID & "&AutoClass="&RsNewsConfigObj("AutoClass")&"&AutoIndex="&RsNewsConfigObj("AutoIndex")&"';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		ElseIf RsNewsConfigObj("AutoIndex")="1" then
			Response.Write("<script>if (confirm(""文字新闻添加成功,是否生成首页?"")==true) {window.location='../refresh/refreshauto.asp?ClassID=" & IClassID & "&AutoClass="&RsNewsConfigObj("AutoClass")&"&AutoIndex="&RsNewsConfigObj("AutoIndex")&"';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		Else
			Response.Write("<script>if (confirm(""文字新闻添加成功,是否继续添加?"")==false) {window.location='NewsList.asp?ClassID=" & IClassID & "';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		End If
	else
		Response.Redirect("NewsList.asp?ClassID=" & IClassID & "")
	end if
	RsNewsConfigObj.Close
	Set RsNewsConfigObj = Nothing
	Response.End
end if
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新闻添加</title>
</head>
<script language="JavaScript" type="text/JavaScript">
<!--
function insertPicAddress() { 
	if (document.NewsForm.IsPicNews.checked==true)
		NewsPicAddress.style.display=''
	else
		NewsPicAddress.style.display='none'
}
//-->
</script>
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body topmargin="2" leftmargin="2">
<form action="" name="NewsForm" method="post">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="SubmitFun();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="添加标题新闻" onClick="location='NewsTitle.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">标题</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="添加图片新闻" onClick="location='NewsPic.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">图片</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="添加下载" onClick="location='DownLoad.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">下载</td>
		  <td width=2 class="Gray">|</td>
		  <%If sHaveValueTF = True then%>
		  <td width=35 align="center" alt="添加商品" onClick="location='../mall/mall_addproducts.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">商品</td>
		  <td width=2 class="Gray">|</td>
		  <%End if%>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp;<input type="hidden" name="Content" value="<% = IContent %>">
              <input type="hidden" name="Action" value="Submit">
              <input type="hidden" name="ClassID" value="<% = IClassID %>"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr>
      <td>
	    <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="60" height="30"> 
              <div align="center">标题</div></td>
            <td height="30"> 
              <input style="width:60%;" type="text" name="Title" value="<% = ITitle %>">
              <input type="checkbox" name="TitleShowReview" value="1" title="在栏目新闻标题后面加上评论2字"<%If ITitleShowReview="1" then response.write("Checked") End If%>>
              显示评论&nbsp;&nbsp;&nbsp;&nbsp; 
              <select name="TitleColor" id="select">
                <option <% if ITitleColor = "#UUUUUU" then Response.Write("Selected")%> value="#UUUUUU" selected>字体颜色</option>
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
              粗体 
              <input type="checkbox" <% if TitleUstr = "1" then Response.Write("Checked") %> name="Titles" value="1">
              斜体 </td>
          </tr>
          <tr> 
            <td height="30"> 
              <div align="center">副标题</div></td>
            <td height="30"> 
              <input style="width:70%;" type="text" name="SubTitle" value="<% = ISubTitle %>">
              &nbsp;&nbsp;&nbsp;&nbsp;
              <input name="IsPicNews" type="checkbox"  id="IsPicNews2" value="IsPicNews" onClick="insertPicAddress()">
              转为图片新闻 </td>
          </tr>
		  <tr id="NewsPicAddress" style="display:none;"> 
            <td height="30"> 
              <div align="center">图片地址</div></td>
            <td height="30"> 
              <input type="text" name="PicPath" style="width:74%" value=<% = iPicPath %>>&nbsp;<input type="button" name="Submit4" value="选择图片" onClick="var TempReturnValue=OpenWindow('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',500,290,window);if (TempReturnValue!='') document.NewsForm.PicPath.value=TempReturnValue;"></td>
          </tr>
        </table></td>
	</tr>
    <tr> 
      <td colspan="2"><table width="100%" border="0" cellpadding="0" cellspacing="0" height="20">
          <tr> 
            <td width="60" height="26" align="center" bgcolor="#EFEFEF" class="LableSelected" id="ContentFolder" onClick="ChangeFolder(this);">新闻内容</td>
            <td width="5" height="26" align="center" class="ToolBarButtonLine" style="cursor:default;">&nbsp;</td>
            <td width="60" height="26" align="center" class="LableDefault" id="AttributeFolder" onClick="ChangeFolder(this);">新闻属性</td>
            <td height="26" class="ToolBarButtonLine" style="cursor:default;">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr id="AttributeArea" style="display:none;"> 
      <td colspan="2"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="ButtonListLeft">
          <tr> 
            <td width="60" height="30"> 
              <div align="center">所属栏目</div></td>
            <td> <input type="text" style="width:74%;" name="ClassCNameShow" readonly value="<% = OldClassEName %>"> 
              &nbsp; <input type="button" name="Submit" value="选择栏目" onClick="SelectClass();"></td>
          </tr>
          <tr> 
            <td height="30"> 
              <div align="center">新闻模板</div></td>
            <td> 
              <input name="NewsTemplet" type="text" id="NewsTemplet" value="<% = INewsTemplet %>" readonly style="width:74%;" > 
              &nbsp; <input type="button" name="Submit" value="选择模板" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<%=sRootDir %>/<% = TempletDir %>',400,300,window,document.NewsForm.NewsTemplet);document.NewsForm.NewsTemplet.focus();"> 
            </td>
          </tr>
          <tr> 
            <td height="30"> 
              <div align="center">添加日期</div></td>
            <td> 
              <input name="AddDate" readonly type="text" value="<% = IAddDate %>" id="AddDate" style="width:74%;"> 
              &nbsp; <input type="button" name="Submit4" value="选择日期" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,110,window,document.NewsForm.AddDate);document.NewsForm.AddDate.focus();"> 
            </td>
          </tr>
          <tr> 
            <td height="30"> 
              <div align="center">所属专题</div></td>
            <td> 
			<%
			if Not IsNull(ISpecialID) And (ISpecialID <> "") then
				Dim RsSpecialObj,TempISpecialID,SpecialCNameText
				SpecialCNameText = ""
				TempISpecialID = ISpecialID
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
			%>
              <input name="SpecialIDText" type="text" style="width:62%" value="<% = SpecialCNameText %>" readonly>  
              <input type="hidden" name="SpecialID" value="<% = ISpecialID %>">  
              <select name="SelSpecialID" style="width:25%" onChange=ChooseSpecial(this.options[this.selectedIndex].value)>
                <option value="" selected> </option> 
                <option value="<%=""&"***"&"Clean"%>" style="color:red">清空</option> 
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
            <td height="30"> 
              <div align="center">新闻来源</div></td> 
            <td> 
              <input name="TxtSourceText" type="text" style="width:62%" value="<% = ITxtSource %>">  
              <input type="hidden" name="TxtSource" value="<% = ITxtSource %>">  
              <select name="SelTxtSource" style="width:25%" onChange="Dosusite(this.options[this.selectedIndex].value);">
                <option value="" selected> </option>
                <option value="<%=""&"***"&"Clean"%>" style="color:red">清空</option>
                <option value="本站原创">本站原创</option>
                <option value="不详">不详</option>
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
              </select>
              <input name="ChkSource" type="checkbox" id="ChkSource" value="SaveSource">
              保存
<div align="center"></div></td>
          </tr>
          <tr> 
            <td height="30"> 
              <div align="center">关 键 字</div></td>
            <td> 
              <input name="KeywordText" type="text" style="width:62%" value="<% = IKeyWords %>"> 
              <input type="hidden" name="KeyWords" value="<% = IKeyWords %>"> 
              <select name="SelKeyWords" style="width:25%" onChange=Dokesite(this.options[this.selectedIndex].value)>
                <option value="" selected> </option>
                <option value="Clean" style="color:red">清空</option>
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
              </select>
              <input name="ChkKeyWord" type="checkbox" id="ChkKeyWord" value="SaveKeyWords">
              保存</td>
          </tr>
          <tr> 
            <td height="30"> 
              <div align="center">新闻作者</div></td>
            <td> 
              <input name="AuthorText" type="text" style="width:62%" value="<% = IAuthor %>"> 
              <input type="hidden" name="Author" value="<% = IAuthor %>"> 
              <select name="SelAuthor" id="select8" style="width:25%" onChange="Doauthsite(this.options[this.selectedIndex].value);">
                <option value="" selected> </option>
                <option value="<%=""&"***"&"Clean"%>" style="color:red">清空</option>
                <option value="佚名">佚名</option>
                <option value="本站">本站</option>
                <option value="未知">未知</option>
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
              </select>
              <input name="ChkAuthor" type="checkbox" id="ChkAuthor" value="SaveAuthor">
              保存 </td>
          </tr>
          <tr> 
            <td height="30"> 
              <div align="center">责任编辑</div></td>
            <td> 
              <input name="EditerText" type="text" style="width:62%" value="<% = IEditer %>"> 
              <input type="hidden" name="Editer" value="<% = IEditer %>"> 
              <select name="Editer1" style="width:25%"  onChange="Editsite(this.options[this.selectedIndex].value);">
                <option value="" selected> </option>
                <option value="<%=""&"***"&"Clean"%>" style="color:red">清空</option>
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
              </select>
              <input name="ChkEditer" type="checkbox" id="ChkEditer" value="SaveEditer">
              保存</td>
          </tr>          <tr> 
            <td height="30"> 
              <div align="center">浏览权限</div></td>
            <td> 
              <select name="BrowPop" id="select7" style="width:100%" onChange="ChooseExeName();">
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
            <td height="30"> 
              <div align="center">点&nbsp;&nbsp;&nbsp;&nbsp;击</div></td>
            <td>
<input name="ClickNum" style="width:100%;" type="text" id="ClickNum3" size="10" value="<% if IClickNum = "" OR IsNull(IClickNum) then Response.Write("1") else Response.Write(IClickNum) %>"></td>
          </tr>
          <tr> 
            <td height="30"> 
              <div align="center">扩&nbsp;展&nbsp;名</div></td>
            <td>
<select name="FileExtName" id="select10" style="width:100%">
                <option <% if IFileExtName = "htm" then Response.Write("Selected") %> value="htm">htm</option>
                <option <% if IFileExtName = "html" then Response.Write("Selected") %> value="html">html</option>
                <option <% if IFileExtName = "shtm" then Response.Write("Selected") %> value="shtm">shtm</option>
                <option <% if IFileExtName = "shtml" then Response.Write("Selected") %> value="shtml">shtml</option>
                <option <% if IFileExtName = "asp" then Response.Write("Selected") %> value="asp">asp</option>
              </select></td>
          </tr>
          <tr> 
            <td height="30"> 
              <div align="center">可选属性</div></td>
            <td><input name="SBSNews" type="checkbox" id="SBSNews" value="1" <%if ISBSNews = "1" then Response.Write("checked") end if%>>
              并排 　　 
              <input name="MarqueeNews" type="checkbox" id="MarqueeNews" value="1" <%if IMarqueeNews = "1" then Response.Write("checked") end if%>>
              滚动 　　　 
              <input name="ProclaimNews" type="checkbox" id="ProclaimNews" value="1" <%if IProclaimNews = "1" then Response.Write("checked") end if%>>
              公告　　　　 
              <input name="RecTF" type="checkbox" id="RecTF" value="1" <%if IRecTF = "1" then Response.Write("checked") end if%>>
              推荐　　　 
              <input name="AuditTF" type="checkbox" id="AuditTF" value="1" checked <%if IAuditTF = "1" then Response.Write("checked") end if%>>
              审核 　　　 
              <input name="ReviewTF" type="checkbox" id="ReviewTF" value="1" onClick="ChooseRiview();" <%if IReviewTF = "1" then Response.Write("checked") end if%>>
              允许评论 
              <input name="ShowReviewTF" type="checkbox" id="ShowReviewTF" value="1" disabled <%if IShowReviewTF = "1" then Response.Write("checked") end if%>>
              显示评论 　<br>
              <br>
              <input name="LinkTF" type="checkbox" id="LinkTF" value="1" <%if ILinkTF = "1" then Response.Write("checked") end if%>>
              内部链接 &nbsp;
              <input name="ManuTF" type="checkbox" id="ManuTF" value="1" <%if Request("ManuTF")="1" then Response.Write("checked") end if%>>
              投稿 　　 &nbsp;&nbsp;
              <input type="checkbox" name="SavePic" value="1" <%if ISavePic = "1" then Response.Write("checked") end if%>>
              远程存图 　 &nbsp;
              <input name="TodayNewsTF" type="checkbox" id="TodayNewsTF" value="1" <%if ITodayNewsTF = "1" then Response.Write("checked") end if%>>
              今日头条 </td>
          </tr>
        </table></td>
    </tr>
    <tr id="ContentArea"> 
      <td colspan="2"><iframe id='NewsContent' src='../../Editer/NewsEditer.asp' frameborder=0 scrolling=no width='100%' height='440'></iframe></td>
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
	if (frames["NewsContent"].CurrMode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return;}
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