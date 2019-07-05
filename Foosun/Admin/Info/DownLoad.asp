<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Refresh/Function.asp" -->
<!--#include file="../Refresh/RefreshFunction.asp" -->
<!--#include file="../Refresh/SelectFunction.asp" -->
<%
Dim DBC,Conn,sRootDir
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
if SysRootDir<>"" then sRootDir="/"+SysRootDir else sRootDir=""
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System(FoosunCMS V3.1.0930)
'最新更新：2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'商业注册联系：028-85098980-601,项目开发：028-85098980-606、609,客户支持：608
'产品咨询QQ：394226379,159410,125114015
'技术支持QQ：315485710,66252421 
'项目开发QQ：415637671，655071
'程序开发：四川风讯科技发展有限公司(Foosun Inc.)
'Email:service@Foosun.cn
'MSN：skoolls@hotmail.com
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.cn  演示站点：test.cooin.com 
'网站通系列(智能快速建站系列)：www.ewebs.cn
'==============================================================================
'免费版本请在程序首页保留版权信息，并做上本站LOGO友情连接
'风讯公司保留此程序的法律追究权利
'如需进行2次开发，必须经过风讯公司书面允许。否则将追究法律责任
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
Set RsMenuConfigObj = Nothing
if Not JudgePopedomTF(Session("Name"),"" & Request("ClassID") & "") then Call ReturnError1()
if Not JudgePopedomTF(Session("Name"),"P010000") then Call ReturnError1()
if Request("DownLoadID") <> "" then
	if Not JudgePopedomTF(Session("Name"),"P010702") then Call ReturnError1()
else
	if Not JudgePopedomTF(Session("Name"),"P010701") then Call ReturnError1()
end if
Dim TempClassID,OldClassObj,OldClassEName,DummyPath_Riker
Dim Action
Dim IDownLoadID,IName,IClassID,IVersion,ITypes,IProperty,ILanguage,IAccredit,IFileSize,IAppraise,ISystemType
Dim IEMail,IProvider,IProviderUrl,IPic,IBrowPop,IDescription,IPassWord,IAddTime,IRecTF,IClassBuildNewsTemp
Dim IAuditTF,IFileExtName,IClickNum,INewsTemplet,IEditTime,IReviewTF,IShowReviewTF
Dim EditContentTF
Dim RsSelectObj,HaveValueTF
Dim AddressNum,AddressIDArrays,RequestNameArrays,RequestUrlArrays,RequestNumberArray,RsDownAddrObj,RsDASql,i

EditContentTF = False
Action = Request("Action")
IClassID = Request.Form("ClassID")
if IClassID="" then IClassID=Request("ClassID")
IDownLoadID = Request("DownLoadID")
if IDownLoadID = "" then
	EditContentTF = False
else
	EditContentTF = True
end if
If IClassID <> "" then
	TempClassID = Cstr(IClassID)
	Set OldClassObj = Conn.Execute("Select ClassID,ClassEName,DownLoadTemp,ClassCName from FS_NewsClass where ClassID='" & TempClassID & "'")
	if Not OldClassObj.Eof then
		OldClassEName = OldClassObj("ClassCName")
		IClassBuildNewsTemp = OldClassObj("DownLoadTemp")
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
if Action = "Submit" then
	HaveValueTF = False
else
	if IDownLoadID <> "" then
		Set RsSelectObj = Conn.Execute("Select * from FS_DownLoad where DownLoadID='" & IDownLoadID & "'")
		if Not RsSelectObj.Eof then
			IName = RsSelectObj("Name")
			IVersion = RsSelectObj("Version")
			ITypes = RsSelectObj("Types")
			IProperty = RsSelectObj("Property")
			ILanguage = RsSelectObj("Language")
			IAccredit = RsSelectObj("Accredit")
			IFileSize = RsSelectObj("FileSize")
			IAppraise = RsSelectObj("Appraise")
			ISystemType = RsSelectObj("SystemType")
			IEMail = RsSelectObj("EMail")
			IProvider = RsSelectObj("Provider")
			IProviderUrl = RsSelectObj("ProviderUrl")
			IPic = RsSelectObj("Pic")
			IBrowPop = RsSelectObj("BrowPop")
			IDescription = RsSelectObj("Description")
			IPassWord = RsSelectObj("PassWord")
			IAddTime = RsSelectObj("AddTime")
			IEditTime = RsSelectObj("EditTime")
			IRecTF = RsSelectObj("RecTF")
			IAuditTF = RsSelectObj("AuditTF")
			IFileExtName = RsSelectObj("FileExtName")
			IClickNum = RsSelectObj("ClickNum")
			INewsTemplet = RsSelectObj("NewsTemplet")
			IReviewTF = RsSelectObj("ReviewTF")
			IShowReviewTF = RsSelectObj("ShowReviewTF")
			RequestNameArrays = ""
			RequestUrlArrays = ""
			RequestNumberArray = ""
			AddressIDArrays = ""
			Set RsDownAddrObj = Server.CreateObject(G_FS_RS)
			RsDASql = "Select * from FS_DownLoadAddress where DownLoadID='" & IDownLoadID & "' order by Number asc"
			RsDownAddrObj.Open RsDASql,Conn,1,1
			AddressNum = RsDownAddrObj.RecordCount
			for i = 0 to RsDownAddrObj.RecordCount-1
					RequestNameArrays = RequestNameArrays & "," & RsDownAddrObj("AddressName")
					RequestUrlArrays = RequestUrlArrays & "," & RsDownAddrObj("Url")
					RequestNumberArray = RequestNumberArray & "," & RsDownAddrObj("Number")
					AddressIDArrays = AddressIDArrays & "," & RsDownAddrObj("ID")
				RsDownAddrObj.MoveNext
			next
			RsDownAddrObj.Close
			Set RsDownAddrObj = Nothing
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
	IName = NoCSSHackAdmin(Request("Name"),"名称")
	IVersion = Request("Version")
	ITypes = Request("Types")
	IProperty = Request("Property")
	ILanguage = Request("Language")
	IAccredit = Request("Accredit")
	IFileSize = Request("FileSize")
	IAppraise = Request("Appraise")
	ISystemType = Request("SystemType")
	IEMail = Request("EMail")
	IProvider = Request("Provider")
	IProviderUrl = Request("ProviderUrl")
	IPic = Request("Pic")
	IBrowPop = Request("BrowPop")
	Dim TempForVar
	For TempForVar = 1 To Request.Form("Description").Count
		IDescription = IDescription & Request.Form("Description")(TempForVar)
	Next
	IPassWord = Request("PassWord")
	IAddTime = Request("AddTime")
	IEditTime = Request("EditTime")
	IRecTF = Request("RecTF")
	IAuditTF = Request("AdutiTF")
	IFileExtName = Request("FileExtName")
	IClickNum = Request("ClickNum")
	INewsTemplet = Request("NewsTemplet")
	IReviewTF = Request("ReviewTF")
	IShowReviewTF = Request("ShowReviewTF")
	AddressNum = Request.Form("AddressNum")
	if AddressNum = "" then AddressNum = 1
	for i = 1 to AddressNum
		RequestNameArrays = RequestNameArrays & "," & Request.Form("AddressName" & i)
		RequestUrlArrays = RequestUrlArrays & "," & Request.Form("Url" & i)
		RequestNumberArray = RequestNumberArray & "," & Request.Form("Number" & i)
		AddressIDArrays = AddressIDArrays & "," & Request.Form("AddressID" & i)
	next
end if
if IsNull(IDescription) then
	IDescription = ""
else
	IDescription = Replace(Replace(IDescription,"""","%22"),"'","%27")
end if
if INewsTemplet = "" OR INewsTemplet = Null then
	if IClassBuildNewsTemp = Null then
		INewsTemplet = ""
	else
		INewsTemplet = IClassBuildNewsTemp
	end if
end if
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>添加下载</title>
</head>
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body topmargin="2" leftmargin="2">
<form action="" name="DownForm" method="post">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width=35 align="center" alt="保存" onClick="SubmitData();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="添加文字新闻" onClick="location='NewsWords.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">文字</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="添加标题新闻" onClick="location='NewsTitle.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">标题</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="添加图片新闻" onClick="location='NewsPic.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">图片</td>
		  <td width=2 class="Gray">|</td>
		  <%If sHaveValueTF = true then%>
		  <td width=35 align="center" alt="添加商品" onClick="location='../mall/mall_addproducts.asp?ClassID=<% = IClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">商品</td>
		  <td width=2 class="Gray">|</td>
		  <%End if%>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp; <input name="action" type="hidden" id="action" value="Submit"> 
             <input type="hidden" name="Description" value="<% = IDescription %>"><input type="hidden" name="ClassID" value="<% = IClassID %>"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="30" colspan="2"><table width="100%" height="30" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="100"><div align="center">名称</div></td>
            <td><input name="Name" type="text" id="Name" style="width:90%" value="<% = IName %>"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td colspan="2"> <table width="100%" border="0" cellpadding="0" cellspacing="0" height="20">
          <tr> 
            <td width="60" height="26" align="center" bgcolor="#EFEFEF" class="LableSelected" id="ContentFolder" onClick="ChangeFolder(this);">下载简介</td>
            <td width="5" align="center" class="ToolBarButtonLine" style="cursor:default;">&nbsp;</td>
            <td onClick="ChangeFolder(this);" id="AttributeFolder" width="60" align="center" class="LableDefault">下载属性</td>
            <td width="5" align="center" class="ToolBarButtonLine" style="cursor:default;">&nbsp;</td>
            <td onClick="ChangeFolder(this);" id="AddressFolder" width="60" align="center" class="LableDefault">下载地址</td>
			<td class="ToolBarButtonLine" style="cursor:default;">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr id="AttributeArea" style="display:none;"> 
      <td height="30" colspan="2"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="ButtonListLeft">
          <tr> 
            <td width="100" height="30"> <div align="center">所属栏目</div></td>
            <td colspan="3"> <input type="text" style="width:74%;" name="ClassCNameShow" readonly value="<% = OldClassEName %>"> 
              &nbsp; <input type="button" name="Submit" value="选择栏目" onClick="SelectClass();"></td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">页面模板</div></td>
            <td colspan="3"><input name="NewsTemplet" type="text" id="NewsTemplet2" style="width:60% " readonly value="<%If INewsTemplet = "" OR INewsTemplet = Null then Response.Write( "/" & RemoveVirtualPath(TempletDir) & "/NewsClass/DownPage.htm") else Response.Write(INewsTemplet)%>" > 
              <input type="button" name="Submit" value="选择模板" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<%=sRootDir %>/<% = TempletDir %>',400,300,window,document.DownForm.NewsTemplet);document.DownForm.NewsTemplet.focus();"></td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">显示图片</div></td>
            <td colspan="3"><input name="Pic" type="text" id="Pic2" style="width:60% " value="<% = IPic %>" > 
              <input type="button" name="Submit4" value="选择图片" onClick="var TempReturnValue=OpenWindow('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',500,290,window);if (TempReturnValue!='') document.DownForm.Pic.value=TempReturnValue;"></td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">添加日期</div></td>
            <td colspan="3"><input name="AddTime" readonly type="text" id="AddTime2" style="width:60% " value="<% if IAddTime = "" then Response.Write(now()) else Response.Write(IAddTime) end if%>"> 
              <input type="button" name="Submit42" value="选择日期" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,110,window,document.DownForm.AddTime);document.DownForm.AddTime.focus();"></td>
          </tr>
          <% if IDownLoadID <> "" then%>
          <tr> 
            <td height="30"> <div align="center">修改日期</div></td>
            <td colspan="3"><input name="EditTime" readonly type="text" id="EditTime" style="width:60% " value="<% if IAddTime = "" then Response.Write(now()) else Response.Write(IAddTime) end if%>"> 
              <input type="button" name="Submit42" value="选择日期" onClick="OpenWindowAndSetValue('../../FunPages/SelectDate.asp',280,110,window,document.DownForm.EditTime);document.DownForm.EditTime.focus();"></td>
          </tr>
          <% end if %>
          <tr> 
            
            <td><div align="center">下载类型</div></td>
            <td colspan="3"><select name="Types" id="select" style="width:60%">
                <option value="1" <%If CStr(ITypes) = "1" then Response.Write("selected")%>>图片</option>
                <option value="2" <%If CStr(ITypes) = "2" then Response.Write("selected")%>>文件</option>
                <option value="3" <%If CStr(ITypes) = "3" then Response.Write("selected")%>>程序</option>
                <option value="4" <%If CStr(ITypes) = "4" then Response.Write("selected")%>>Flash</option>
                <option value="5" <%If CStr(ITypes) = "5" then Response.Write("selected")%>>音乐</option>
                <option value="6" <%If CStr(ITypes) = "6" then Response.Write("selected")%>>影视</option>
                <option value="7" <%If CStr(ITypes) = "7" then Response.Write("selected")%>>其它</option>
              </select></td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">程序语言</div></td>
            <td><select name="Language" id="select4" style="width:90%">
                <option value="1" <%If CStr(ILanguage) = "1" then Response.Write("selected")%>>简体中文</option>
                <option value="2" <%If CStr(ILanguage) = "2" then Response.Write("selected")%>>繁体中文</option>
                <option value="3" <%If CStr(ILanguage) = "3" then Response.Write("selected")%>>英文</option>
                <option value="4" <%If CStr(ILanguage) = "4" then Response.Write("selected")%>>法文</option>
                <option value="5" <%If CStr(ILanguage) = "5" then Response.Write("selected")%>>日文</option>
                <option value="6" <%If CStr(ILanguage) = "6" then Response.Write("selected")%>>德文</option>
              </select></td>
            <td><div align="center">文件大小</div></td>
            <td><input name="FileSize" type="text" id="FileSize" style="width:90%" value="<% = IFileSize %>"></td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">系统平台</div></td>
            <td colspan="3"><input name="SystemType" type="text" id="SystemType2" style="width:63%" value="<% = ISystemType %>"> 
              <select name="SystemChoose" id="select5" style="width:32%" onChange=ChooseSystem(this.options[this.selectedIndex].value)>
                <option value="Clean" style="color:red">清空</option>
                <option <% if ISystemType = "Win95" then Response.Write("Selected") %> value="Win95">Win95</option>
                <option <% if ISystemType = "Win98" then Response.Write("Selected") %> value="Win98">Win98</option>
                <option <% if ISystemType = "WinMe" then Response.Write("Selected") %> value="WinMe">WinMe</option>
                <option <% if ISystemType = "WinNT" then Response.Write("Selected") %> value="WinNT">WinNT</option>
                <option <% if ISystemType = "Win2000" then Response.Write("Selected") %> value="Win2000">Win2000</option>
                <option <% if ISystemType = "WinXP" then Response.Write("Selected") %> value="WinXP" selected>WinXP</option>
                <option <% if ISystemType = "Win2003" then Response.Write("Selected") %> value="Win2003">Win2003</option>
                <option <% if ISystemType = "Linux" then Response.Write("Selected") %> value="Linux">Linux</option>
              </select></td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">开 发 商</div></td>
            <td><input name="Provider" type="text" id="Provider2" style="width:90%" value="<% = IProvider %>"></td>
            <td><div align="center">开发商Url</div></td>
            <td><input name="ProviderUrl" type="text" id="ProviderUrl" style="width:90%" value="<% = IProviderUrl %>"></td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">扩 展 名</div></td>
            <td width="390"><select name="FileExtName" id="select6" style="width:90%">
                <option value="htm" <%If IFileExtName = "htm" then Response.Write("selected")%>>htm</option>
                <option value="html" <%If IFileExtName = "html" then Response.Write("selected")%>>html</option>
                <option value="shtm" <%If IFileExtName = "shtm" then Response.Write("selected")%>>shtm</option>
                <option value="shtml" <%If IFileExtName = "shtml" then Response.Write("selected")%>>shtml</option>
                <option value="asp" <%If IFileExtName = "asp" then Response.Write("selected")%>>asp</option>
              </select></td>
            <td width="70"><div align="center">下载权限</div></td>
            <td width="427"><select name="BrowPop" id="select8" style="width:90%" onChange="ChooseExeName();">
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
            <td height="30"> <div align="center">解压密码</div></td>
            <td><input name="PassWord" type="text" id="PassWord" style="width:90%" value="<% = IPassWord %>"></td>
            <td><div align="center">下载次数</div></td>
            <td><input name="ClickNum" type="text" id="ClickNum2" style="width:90%" value="<%if IClickNum = "" then Response.Write("0") else Response.Write(IClickNum) end if %>"></td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">版本名称</div></td>
            <td><input name="Version" type="text" id="Version" style="width:90%" value="<% = IVersion %>"></td>
            <td><div align="center">程序授权</div></td>
            <td><select name="Accredit" id="select7" style="width:90%">
                <option value="1" <%If CStr(IAccredit) = "1" then Response.Write("selected")%>>免费</option>
                <option value="2" <%If CStr(IAccredit)="2" then Response.Write("selected")%>>共享</option>
                <option value="3" <%If CStr(IAccredit)="3" then Response.Write("selected")%>>试用</option>
                <option value="4" <%If CStr(IAccredit)="4" then Response.Write("selected")%>>演示</option>
                <option value="5" <%If CStr(IAccredit)="5" then Response.Write("selected")%>>注册</option>
                <option value="6" <%If CStr(IAccredit)="6" then Response.Write("selected")%>>破解</option>
                <option value="7" <%If CStr(IAccredit)="7" then Response.Write("selected")%>>零售</option>
                <option value="8" <%If CStr(IAccredit)="8" then Response.Write("selected")%>>其它</option>
              </select></td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">星级评价</div></td>
            <td><select name="Appraise" id="select10" style="width:90%">
                <option value="1" <%If CStr(IAppraise)="1" then Response.Write("selected")%>>★</option>
                <option value="2" <%If CStr(IAppraise)="2" then Response.Write("selected")%>>★★</option>
                <option value="3" <%If CStr(IAppraise)="3" then Response.Write("selected")%>>★★★</option>
                <option value="4" <%If CStr(IAppraise)="4" then Response.Write("selected")%>>★★★★</option>
                <option value="5" <%If CStr(IAppraise)="5" then Response.Write("selected")%>>★★★★★</option>
                <option value="6" <%If CStr(IAppraise)="6" then Response.Write("selected")%>>★★★★★★</option>
              </select></td>
            <td><div align="center">E_Mail</div></td>
            <td><input name="EMail" type="text" id="EMail2" style="width:90%" value="<% = IEMail %>"></td>
          </tr>
          <tr> 
            <td height="30"> <div align="center">推荐下载 </div></td>
            <td colspan="3"><select name="RecTF" id="select">
                <option value="1" <%If CStr(IRecTF)="1" then Response.Write("selected")%>>是</option>
                <option value="0" <%If CStr(IRecTF)="0" then Response.Write("selected")%>>否</option>
              </select>
              　　　　　可选属性：审核 
              <input name="AuditTF" type="checkbox" value="1" <% If CStr(IAuditTF)="1" then Response.Write("checked") %>> 
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 允许评论 
              <input name="ReviewTF" type="checkbox" id="ReviewTF2" value="1" onClick="ChooseRiview();" <%if IReviewTF = "1" then Response.Write("checked") end if%>> 
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 显示评论 
              <input name="ShowReviewTF" type="checkbox" id="ShowReviewTF2" value="1" <%if IShowReviewTF = "1" then Response.Write("checked") end if%> <%if IReviewTF = "0" then Response.Write("disabled") end if %>> 
            </td>
          </tr>
        </table></td>
    </tr>
    <tr id="ContentArea"> 
      <td height="30" colspan="2"><iframe id='NewsContent' src='../../Editer/DownLoadEditer.asp' frameborder=0 scrolling=no width='100%' height='470'></iframe></td>
    </tr>
    <tr id="AddressArea" style="display:none;"> 
      <td height="30" colspan="2">
		  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="ButtonListLeft">
          <tr> 
			  
            <td width="11%" height="30"></td>
			  <td width="89%">下载地址数量&nbsp;
				<input name="AddressNum" type="text" id="AddressNum" value="<%=AddressNum%>" size="8">
				<input type="button" name="Submit3" value=" 设 置 " onClick="ChooseOption();SetOptionsValue()">
			  </td>
			</tr>
			<tr> 
			  
            <td height="30" colspan="2" id="Options">&nbsp;</td>
			</tr>
		  </table>
	  </td>
    </tr>
</table>
</form>
</body>
</html>
<script language="javascript">
var AddressCount=0;
var RequestNameArray=new Array();
var RequestColorArray=new Array();
var RequestAddressIDArray=new Array();
var DocumentReadyTF=false;
var TempRequestNameArray,TempRequestColorArray,TempRequestNumberArray,TempRequestAddressIDArray;
TempRequestNameArray='<% = RequestNameArrays %>';
TempRequestColorArray='<% = RequestUrlArrays %>';
TempRequestNumberArray='<% = RequestNumberArray %>';
TempRequestAddressIDArray='<% = AddressIDArrays %>';
RequestNameArray = TempRequestNameArray.split(",");
RequestColorArray = TempRequestColorArray.split(",");
RequestNumberArray = TempRequestNumberArray.split(",");
RequestAddressIDArray = TempRequestAddressIDArray.split(",");
AddressCount=RequestNumberArray.length;
function document.onreadystatechange()
{
	if (document.readyState!="complete") return;
	if (DocumentReadyTF) return;
	DocumentReadyTF = true;
	ChooseOption();
	SetOptionsValue();
}
function ChooseRiview()
{
 	if (document.DownForm.ReviewTF.checked==true)
	{
		document.DownForm.ShowReviewTF.disabled=false;
	}
 	else
	{
		document.DownForm.ShowReviewTF.disabled=true;
	}
}
function SetOptionsValue()
{
	if ((RequestNameArray.length==0)||(RequestColorArray.length==0)||(RequestNumberArray.length==0)||(RequestAddressIDArray.length==0)) return;
	var AddressNum=document.DownForm.AddressNum.value;
	for (i=1;i<=AddressNum;i++)
	{
		if (i>=AddressCount) 
		{
			document.all('AddressName'+i).value='地址'+i;
			document.all('Url'+i).value='';
			document.all('Number'+i).value='';
			document.all('AddressID'+i).value='';
		}
		else
		{
			document.all('AddressName'+i).value=RequestNameArray[i];
			document.all('Url'+i).value=RequestColorArray[i];
			document.all('Number'+i).value=RequestNumberArray[i];
			document.all('AddressID'+i).value=RequestAddressIDArray[i];
		}
	}
}
function ChooseOption()
 {
	var AddressNum = document.DownForm.AddressNum.value;
	var i,Optionstr;
	Optionstr = '<table width="100%" border="0" cellspacing="5" cellpadding="0">';
	for (i=1;i<=AddressNum;i++)
	{
	   Optionstr = Optionstr+'<tr><td><div align="center">地址名称'+i+'&nbsp;<input type="text" size="20" name="AddressName'+i+'" value="">&nbsp;</div></td><td><div align="center">链接地址&nbsp;<input type="text" size="30" name="Url'+i+'" value="">&nbsp;</div></td><td><div align="center"><input type="button" name="Submit4" onClick="SetFilePath(document.DownForm.Url'+i+');" value="选择文件"></div></td><td><div align="center">地址排序&nbsp;<input type="text" name="Number'+i+'" value="'+i+'" size="3"><input name="AddressID'+i+'" type="hidden" value=""></div></td></tr>';
	}
	Optionstr = Optionstr+'</table>'; 
	document.all.Options.innerHTML = Optionstr;
  }
function SetFilePath(Obj)
{
	var ReturnValue=OpenWindow('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300);
	if (ReturnValue!='007007007007') Obj.value=ReturnValue;
}
function ChangeFolder(el)
{
	if (el.className=='LableSelected') return;
	var OperObj=null;
	var FolderIDArray=new Array('ContentFolder','AttributeFolder','AddressFolder');
	var EditAreaIDArray=new Array('ContentArea','AttributeArea','AddressArea');
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
function ChooseExeName()
{
  if (document.DownForm.BrowPop.value!='') document.DownForm.FileExtName.disabled=true;
  else document.DownForm.FileExtName.disabled=false;
 }
function SubmitData()
{
	document.DownForm.Description.value=frames["NewsContent"].EditArea.document.body.innerHTML;
	document.DownForm.submit();
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
ChooseExeName();
</script>
<%
if Request.Form("action") = "Submit" then
	Dim IDownObj,DownLoadSql
	Dim RsVoteObj,RsVoteSql,VoteName,RsDownAddrAddObj,RsVoteOptionSql
	If Replace(Replace(Replace(Replace(IName,"/",""),"\",""),"'","")," ","") = "" then
		Response.Write("<script>alert(""下载名称为空或是含有非法字符"");</script>")
		Response.End 
	End If
	If IClassID = "" then
		Response.Write("<script>alert(""请选择栏目"");</script>")
		Response.End 
	End If
	If IsDate(IAddTime) = false then
		Response.Write("<script>alert(""添加时间格式不正确"");</script>")
		Response.End 
	End If
	If IsNumeric(IClickNum) = false then
		Response.Write("<script>alert(""下载次数必须为数字类型"");</script>")
		Response.End 
	End If
	If INewsTemplet = "" then
		Response.Write("<script>alert(""文件模板名称不能为空"");</script>")
		Response.End 
	End If
	'判断下载地址
	For i=1 to Request.Form("AddressNum")
		if Request.Form("AddressName" & i) = "" or isnull(Request.Form("AddressName" & i & "")) then
			Response.Write("<script>alert(""地址名称" & i & "不能为空"");</script>")
			Response.End
		end if
		If Request.Form("Url"&i)="" or isnull(Request.Form("Url"&i)) then
			Response.Write("<script>alert(""链接地址" & i & "不能为空"");</script>")
			Response.End
		End If
	Next
	'判断下载地址
	Dim NewsFileNames,RsNewsConfigObj,INewsAddSql
	Set RsNewsConfigObj = Conn.Execute("Select DoMain,NewsFileName,AutoClass,AutoIndex from FS_Config")
	if IDownLoadID <> "" then
		Set IDownObj = Server.CreateObject(G_FS_RS)
		DownLoadSql = "Select * from FS_DownLoad where DownLoadID='" & IDownLoadID & "'"
		IDownObj.open DownLoadSql,Conn,3,3
	else
		IDownLoadID = GetRandomID18()
		Set IDownObj = Server.CreateObject(G_FS_RS)
		DownLoadSql = "Select * from FS_DownLoad where 1=0"
		IDownObj.open DownLoadSql,Conn,3,3
		IDownObj.AddNew
		IDownObj("DownLoadID") = IDownLoadID    '新闻ID
		NewsFileNames = NewsFileName(RsNewsConfigObj("NewsFileName"),IClassID,IDownLoadID)
		IDownObj("FileName") = NewsFileNames
	end if
	IDownObj("Name") = IName
	IDownObj("ClassID") = IClassID
	if IVersion <> "" then
		IDownObj("Version") = Replace(Replace(IVersion,"""",""),"'","")
	end if
	IDownObj("Types") = Cint(ITypes)
	IDownObj("Property") = Cint(IProperty)
	IDownObj("Language") = Cint(ILanguage)
	if IAccredit <> "" then 
		IDownObj("Accredit") = Cint(IAccredit)
	end if
	If IFileSize <> "" then
		IDownObj("FileSize") = Replace(Replace(IFileSize,"""",""),"'","")
	End If
	If Isnumeric(IAppraise) then
		IDownObj("Appraise") = Cint(IAppraise)
	End If
	If ISystemType <> "" then
		IDownObj("SystemType") = Replace(Replace(ISystemType,"""",""),"'","")
	End If
	If IEMail <> "" then
		IDownObj("EMail") = Replace(Replace(IEMail,"""",""),"'","")
	End If
	If IProvider <> "" then
		IDownObj("Provider") = IProvider
	End If
	If IProviderUrl <> "" then
		IDownObj("ProviderUrl") = IProviderUrl
	End If
	if IPic <> "" then
		IDownObj("Pic") = Replace(IPic,"'","")
	end if
	if IBrowPop <> "" then
		IDownObj("BrowPop") = Cint(IBrowPop)
	end if
	Dim Description_Loop_Var,Save_Description
	For Description_Loop_Var = 1 To Request.Form("Description").Count
		Save_Description = Save_Description & Request.Form("Description")(Description_Loop_Var)
	Next
	IDownObj("Description") = replace(Save_Description,WebDomain,"")
	IDownObj("PassWord") = IPassWord
	IDownObj("AddTime") = Formatdatetime(IAddTime)
	if IDownLoadID <> "" then
		IDownObj("EditTime") = Formatdatetime(IEditTime)
	end if
	if IRecTF = "1" then
		IDownObj("RecTF") = 1
	else
		IDownObj("RecTF") = 0
	end if
	if Request("AuditTF") = "1" then
		IDownObj("AuditTF") = 1
	else
		IDownObj("AuditTF") = 0
	end if
	If Request.Form("BrowPop") <> "" then
		IDownObj("FileExtName") = "asp"
	Else
		IDownObj("FileExtName") = IFileExtName
	End If
	if IClickNum <> "" then
		IDownObj("ClickNum") = Clng(IClickNum)
	end if
	IDownObj("NewsTemplet") = Cstr(INewsTemplet)
	if IReviewTF = "1" Then
		IDownObj("ReviewTF") = 1
	Else
		IDownObj("ReviewTF") = 0
	End if
	if IShowReviewTF = "1" Then
		IDownObj("ShowReviewTF") = 1
	Else
		IDownObj("ShowReviewTF") = 0
	End if
	IDownObj.Update
	IDownObj.Close
	Set IDownObj = Nothing
	'保存下载地址
	Set RsDownAddrAddObj = Server.CreateObject(G_FS_RS)
	For i = 1 to Request.Form("AddressNum")
		if Request.Form("AddressID" & i) <> "" then
			RsVoteOptionSql = "Select * from FS_DownLoadAddress where ID='" & Request.Form("AddressID" & i) & "'"
		else
			RsVoteOptionSql = "Select * from FS_DownLoadAddress where 1=0"
		end if
		RsDownAddrAddObj.Open RsVoteOptionSql,Conn,3,3
		if RsDownAddrAddObj.Eof then
			RsDownAddrAddObj.AddNew
			RsDownAddrAddObj("ID") = GetRandomID18()
		end if
		RsDownAddrAddObj("DownLoadID") = Cstr(IDownLoadID)
		RsDownAddrAddObj("AddressName") = Replace(Request.Form("AddressName" & i),"'","")
		RsDownAddrAddObj("Url") = Request.Form("Url" & i)
		If isnumeric(Request.Form("Number" & i)) then
			RsDownAddrAddObj("Number") = Cint(Request.Form("Number" & i))
		Else
			RsDownAddrAddObj("Number") = i
		End If
		RsDownAddrAddObj.Update
		RsDownAddrAddObj.Close
	Next
	Set RsDownAddrAddObj = Nothing
	'保存下载地址
	'生成文件
	if Request.Form("AuditTF") = "1" then
		Dim CreatePageObj
		Set CreatePageObj = Conn.Execute("Select * from FS_DownLoad where DownLoadID='" & IDownLoadID & "'")
		If Not CreatePageObj.eof then
			RefreshDownLoad CreatePageObj
		Else
		  Response.Write("<script>if (confirm(""下载添加成功,但未能成功生成新闻文件,是否继续添加?"")==false) {window.location='NewsList.asp?ClassID="&IClassID&"';} else {window.location=""?ClassID="&IClassID&""";}</script>")
		  Response.End
		End If	
		CreatePageObj.Close
		Set CreatePageObj = Nothing
	end if 
	if EditContentTF = True then
		Response.Redirect("DownloadList.asp?ClassID=" & IClassID)
	else
		If RsNewsConfigObj("AutoClass")="1" and RsNewsConfigObj("AutoIndex")="1" then
			Response.Write("<script>if (confirm(""下载添加成功,是否生成此栏目和首页?"")==true) {window.location='../refresh/refreshauto.asp?ClassID=" & IClassID & "&AutoClass="&RsNewsConfigObj("AutoClass")&"&AutoIndex="&RsNewsConfigObj("AutoIndex")&"';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		ElseIf RsNewsConfigObj("AutoClass")="1" then
			Response.Write("<script>if (confirm(""下载添加成功,是否生成此栏目?"")==true) {window.location='../refresh/refreshauto.asp?ClassID=" & IClassID & "&AutoClass="&RsNewsConfigObj("AutoClass")&"&AutoIndex="&RsNewsConfigObj("AutoIndex")&"';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		ElseIf RsNewsConfigObj("AutoIndex")="1" then
			Response.Write("<script>if (confirm(""下载添加成功,是否生成首页?"")==true) {window.location='../refresh/refreshauto.asp?ClassID=" & IClassID & "&AutoClass="&RsNewsConfigObj("AutoClass")&"&AutoIndex="&RsNewsConfigObj("AutoIndex")&"';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		Else
			Response.Write("<script>if (confirm(""下载添加成功,是否继续添加?"")==false) {window.location='NewsList.asp?ClassID=" & IClassID & "';} else {window.location=""?ClassID="&IClassID&""";} </script>")
		End If
	end if
	Set RsNewsConfigObj = Nothing
	Response.End
end if
%>