<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
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
dim conn,RsConfig,DBC,SQLStr,FSOObj1
set DBC=New DataBaseClass
set conn=DBC.OpenConnection
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P040501") then Call ReturnError1()
SQLStr="select * from FS_config"
set RsConfig=server.CreateObject(G_FS_RS)
RsConfig.open SQLStr,conn,1,3
if request.Form("operation")="Modify" then
	On  Error Resume Next
	if request.Form("NewsFileName8")="" then
		if len(request.form("NewsFileName1")&request.form("NewsFileName2")&request.form("NewsFileName3")&request.form("NewsFileName4")&request.form("NewsFileName5")&request.form("NewsFileName6")&request.form("NewsFileName7")&request.form("NewsFileName8")&request.form("NewsFileName9")&request.form("NewsFileName10")&request.form("NewsFileName11")&request.form("NewsFileName12")&request.form("NewsFileName13"))<3 then
			Response.Write("<script>alert(""错误：\n\n如果生成文件名没有选择NewsID,请选择至少3位以上的随机文件名组合,并带一个随机数"");location.href=""SysParameter.asp"";</script>")
			Response.End
		end if
	end if
	if str=IsValidEmail (request.Form("Email")) then
		Response.Write("<script>alert(""错误：\n请输入一个正确的Email"&CopyRight&""");location.href=""SysParameter.asp"";</script>")
		Response.End
	else
		RsConfig("Email")=Replace(Replace(request.form("Email"),"'",""),"""","")
	end if
	RsConfig("Domain")=Replace(Replace(request.form("Domain"),"'",""),"""","")
	RsConfig("UpFileSize")=Replace(Replace(request.form("UpFileSize"),"'",""),"""","")
	RsConfig("UpFileType")=Replace(Replace(request.form("UpFileType"),"'",""),"""","")
	RsConfig("ThumbnailComponent")=Replace(Replace(request.form("ThumbnailComponent"),"'",""),"""","")
	RsConfig("RateTF")=Replace(Replace(request.form("RateTF"),"'",""),"""","")
	RsConfig("ThumbnailWidth")=Replace(Replace(request.form("ThumbnailWidth"),"'",""),"""","")
	RsConfig("ThumbnailHeight")=Replace(Replace(request.form("ThumbnailHeight"),"'",""),"""","")
	RsConfig("ThumbnailRate")=Replace(Replace(request.form("ThumbnailRate"),"'",""),"""","")
	RsConfig("MarkComponent")=Replace(Replace(request.form("MarkComponent"),"'",""),"""","")
	RsConfig("MarkType")=Replace(Replace(request.form("MarkType"),"'",""),"""","")
	RsConfig("MarkText")=Replace(Replace(request.form("MarkText"),"'",""),"""","")
	RsConfig("MarkFontSize")=Replace(Replace(request.form("MarkFontSize"),"'",""),"""","")
	RsConfig("MarkFontColor")=Replace(Replace(request.form("MarkFontColor"),"'",""),"""","")
	RsConfig("MarkFontName")=Replace(Replace(request.form("MarkFontName"),"'",""),"""","")
	RsConfig("MarkFontBond")=Replace(Replace(request.form("MarkFontBond"),"'",""),"""","")
	RsConfig("MarkPicture")=Replace(Replace(request.form("MarkPicture"),"'",""),"""","")
	RsConfig("MarkOpacity")=Csng(Replace(Replace(request.form("MarkOpacity"),"'",""),"""",""))
	RsConfig("MarkWidth")=Replace(Replace(request.form("MarkWidth"),"'",""),"""","")
	RsConfig("MarkHeight")=Replace(Replace(request.form("MarkHeight"),"'",""),"""","")
	RsConfig("MarkTranspColor")=Replace(Replace(request.form("MarkTranspColor"),"'",""),"""","")
	RsConfig("MarkPosition")=Replace(Replace(request.form("MarkPosition"),"'",""),"""","")
	RsConfig("SiteName")=Replace(Replace(request.form("SiteName"),"'",""),"""","")
	RsConfig("AutoRefreshTF")=Replace(Replace(request.form("AutoRefreshTF"),"'",""),"""","")
	RsConfig("AutoJSTF")=Replace(Replace(request.form("AutoJSTF"),"'",""),"""","")
	RsConfig("MailServer")=Replace(Replace(request.form("MailServer"),"'",""),"""","")
	RsConfig("MailName")=Replace(Replace(request.form("MailName"),"'",""),"""","")
	RsConfig("MailPass")=Replace(Replace(request.form("MailPass"),"'",""),"""","")
	RsConfig("Copyright")=Replace(Replace(request.form("Copyright"),"'",""),"""","")
	RsConfig("SendPoint")=Clng(Replace(Replace(request.form("SendPoint"),"'",""),"""",""))
	RsConfig("UserConfer")=Request.Form("UserConfer")
	RsConfig("NewsFileName")=request.form("NewsFileName1")&request.form("NewsFileName2")&request.form("NewsFileName3")&request.form("NewsFileName4")&request.form("NewsFileName5")&request.form("NewsFileName6")&request.form("NewsFileName7")&request.form("NewsFileName8")&request.form("NewsFileName9")&request.form("NewsFileName10")&request.form("NewsFileName11")&request.form("NewsFileName12")&request.form("NewsFileName13")
	if request.form("AutoClass")<>"" then
		RsConfig("AutoClass")=1
	else
		RsConfig("AutoClass")=0
	end if
	if request.form("UseDatePath")<>"0" then
		RsConfig("UseDatePath")=1
		Application.lock
		Application("UseDatePath")="1"
		Application.unlock
	else
		RsConfig("UseDatePath")=0
		Application.lock
		Application("UseDatePath")="0"
		Application.unlock
	end if
	RsConfig("IsShop")=0
	if request.form("IsEmail")<>"0" then
		RsConfig("IsEmail")=1
	else
		RsConfig("IsEmail")=0
	end if
	RsConfig("IsChange")=0
	if request.form("HelpTF")<>"0" then
		RsConfig("HelpTF")=1
	else
		RsConfig("HelpTF")=0
	end if
	if request.form("MemberType")="1" then
		RsConfig("MemberType")=1
	elseif request.form("MemberType")="2" then
		RsConfig("MemberType")=2
	else
		RsConfig("MemberType")=0
	end if
	if request.form("AutoIndex")<>"" then
		RsConfig("AutoIndex")=1
	else
		RsConfig("AutoIndex")=0
	end if
	if request.form("MakeType")="0" then
		RsConfig("MakeType")=0
	else
		RsConfig("MakeType")=1
	end if
	if request.Form("NumberLoginPoint")<>"" then 
		RsConfig("NumberLoginPoint")=clng(request.Form("NumberLoginPoint"))
	else
		RsConfig("NumberLoginPoint")=1
	end if
	if request.Form("MaxContent")<>"" then 
		RsConfig("MaxContent")=clng(request.Form("MaxContent"))
	else
		RsConfig("MaxContent")=2000
	end if
	if request.Form("NumberContPoint")<>"" then 
		RsConfig("NumberContPoint")=clng(request.Form("NumberContPoint"))
	else
		RsConfig("NumberContPoint")=10
	end if
	if request.Form("QPoint")<>"" then 
		RsConfig("QPoint")=clng(request.Form("QPoint"))
	else
		RsConfig("QPoint")=5
	end if
	if request.form("ReviewShow") = "1" then
		RsConfig("ReviewShow") = 1
	else
		RsConfig("ReviewShow") = 0
	end if
	RsConfig("IndexExtName")=Replace(Replace(request.form("IndexExtName"),"'",""),"""","")
	RsConfig.update
	'Set FSOObj1 = Server.CreateObject(G_FS_FSO)
	'if FSOObj1.FileExists(Server.MapPath("/")&"\"& SysRootDir &"\index."&Request.Form("oldIndexExtName")&"") then
	'   FSOObj1.DeleteFile(Server.MapPath("/")&"\"& SysRootDir &"\index."&Request.Form("oldIndexExtName")&"")
	'end if
	'Set FSOObj1=nothing
	if Err.Mumber = 0 then
%>
<script language="javascript">
alert('修改成功<%=CopyRight%>');window.location='SysParameter.asp';
</script>
<%
	else
%>
<script language="javascript">
alert('有错误发生，请刷新后再试');window.location='SysParameter.asp';
</script>
<%
	Response.Redirect("SysParameter.asp")  
	end if
	Response.End

end if 

%>

<html>
<title>网站基本参数设置</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.SysParaButtonStyle {
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-right-color: #999999;
	border-bottom-color: #999999;
	border-left-color: #FFFFFF;
	background-color: #E6E6E6;
}
-->
</style>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2" scroll=yes onLoad="ShowInfo(<%=RsConfig("MarkComponent")%>);ShowThumbnailInfo(<%=RsConfig("ThumbnailComponent")%>);ShowThumbnailSetting(<%=RsConfig("RateTF")%>)"  oncontextmenu="return false;">
<form name=form method=post action="" >
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
      <td height="26" colspan="5" valign="middle" bgcolor="#EEEEEE"> 
        <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="document.form.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td width="905"> <input type=hidden name=operation value=Modify> </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="2" cellspacing="1"  bordercolor="e6e6e6" bgcolor="#E3E3E3">
    <tr bgcolor="#FFFFFF"> 
      <td colspan="5" height="1"></td>
    </tr>
    <tr valign="middle" bgcolor="#F2F2F2"> 
      <td height="23" colspan="2"align="left"><strong>系统参数</strong></td>
    </tr>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td width="184" height="23"align="right"> 您的信箱</td>
      <td width="613" height="1"> <input name="Email" type="text" value="<%=RsConfig("Email")%>" size="30"> 
      </td>
    </tr>
    <tr valign="middle" bgcolor="#F5F5F5" style="display:none"> 
      <td height="23"> <div align="right">当前版本</div></td>
      <td width="613" height="23" bgcolor="#F5F5F5"> <input name="Version" type="text"  value="<%=RsConfig("Version")%>" size="30" disabled> 
      </td>
    </tr>
    <tr valign="middle" bgcolor="#F5F5F5" > 
      <td height="23"> <div align="right">程序使用的域名</div></td>
      <td width="613" height="23" bgcolor="#F5F5F5"> <input name="DoMain" type="text" id="DoMain" value="<%=RsConfig("DoMain")%>" size="30"> 
        <font color="#FF0000">(请使用http://标识),后面不能带&quot;/&quot;符号</font></td>
    </tr>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="23"> <div align="right">站点名称</div></td>
      <td width="613" height="23" bgcolor="#F5F5F5"> <input name="SiteName" type="text" value="<%=RsConfig("SiteName")%>" size="30"></td>
    </tr>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="23"> <div align="right">上传文件的类型</div></td>
      <td width="613" height="23" bgcolor="#F5F5F5"> <input name="UpFileType" type="text"  value="<%=RsConfig("UpFileType")%>" size=30> 
        <span class="Notices"><font color="#FF0000">(请用“,”逗号隔开) </font></span></td>
    </tr>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="23"> <div align="right">限制上传文件大小</div></td>
      <td width="613" height="23" bgcolor="#F5F5F5"> <input type="text" size=10 name="UpFileSize" value="<%=RsConfig("UpFileSize")%>" >
        KB</td>
    </tr>
    <tr bgcolor="#F5F5F5" > 
      <td height="23" align="right">选择图片缩略图功能</td>
      <td width="613" bgcolor="#F5F5F5"> <select name="ThumbnailComponent" id="ThumbnailComponent" onChange="ShowThumbnailInfo(this.value)" style="width:30%">
          <option value=0 <%If RsConfig("ThumbnailComponent") = "0" Then Response.Write("selected") End if%>>关闭 
          <option value=1 <%If RsConfig("ThumbnailComponent") = "1" Then Response.Write("selected") End if%>>AspJpeg组件 
          <option value=2 <%If RsConfig("ThumbnailComponent") = "2" Then Response.Write("selected") End if%>>wsImage组件 
          <option value=3 <%If RsConfig("ThumbnailComponent") = "3" Then Response.Write("selected") End if%>>SA-ImgWriter组件 
          <option value=4 <%If RsConfig("ThumbnailComponent") = "4" Then Response.Write("selected") End if%>>CreatePreviewImage组件 
        </select> <span id="ThumbnailComponentInfo"></span> </td>
    </tr>
    <tr bgcolor="#F5F5F5" id ="ThumbnailSetting" style="display:none"> 
      <td height="23" align="right"> <input type="radio" name="RateTF" value="1" onClick="ShowThumbnailSetting(1);" <%If RsConfig("RateTF") = "1" Then Response.Write("checked") End if%>>
        按比例 
        <input type="radio" name="RateTF" value="0" onClick="ShowThumbnailSetting(0);" <%If RsConfig("RateTF") = "0" Then Response.Write("checked") End if%>>
        按大小 </td>
      <td width="613" bgcolor="#F5F5F5"> <div id ="ThumbnailSetting0" style="display:none">宽度： 
          <input type="text" name="ThumbnailWidth" size=10 value="<%=RsConfig("ThumbnailWidth")%>">
          象素高度： 
          <input type="text" name="ThumbnailHeight" size=10 value="<%=RsConfig("ThumbnailHeight")%>">
          象素</div>
        <div id ="ThumbnailSetting1" style="display:none">比例： 
          <input type="text" name="ThumbnailRate" size=10 value="<%If Left(RsConfig("ThumbnailRate"),1) = "." Then Response.Write("0"&RsConfig("ThumbnailRate")) Else Response.Write(RsConfig("ThumbnailRate")) End if%>">
          如60%请填写0.6 </div></td>
    </tr>
    <tr bgcolor="#F5F5F5" > 
      <td height="23" align="right">选择图片水印功能</td>
      <td width="613" bgcolor="#F5F5F5"> <select name="MarkComponent" id="MarkComponent" onChange="ShowInfo(this.value)" style="width:30%">
          <option value=0 <%If RsConfig("MarkComponent") = "0" Then Response.Write("selected") End if%>>关闭 
          <option value=1 <%If RsConfig("MarkComponent") = "1" Then Response.Write("selected") End if%>>AspJpeg组件 
          <option value=2 <%If RsConfig("MarkComponent") = "2" Then Response.Write("selected") End if%>>wsImage组件 
          <option value=3 <%If RsConfig("MarkComponent") = "3" Then Response.Write("selected") End if%>>SA-ImgWriter组件 
        </select> <span id="ComponentInfo"></span> </td>
    </tr>
    <tr align="left" valign="top" bgcolor="#F5F5F5" id="WaterMarkSetting" style="display:none" cellpadding="0" cellspacing="0"> 
      <td colspan=2> <table width=100% cellpadding="2" cellspacing="1"  bordercolor="e6e6e6" bgcolor="#E3E3E3">
          <tr bgcolor="#FFFFFF"> 
            <td width=193 height="23" align="right">水印类型</td>
            <td width="615"> <SELECT name="MarkType" id="MarkType">
                <OPTION value="1" <%If RsConfig("MarkType") = "1" Then Response.Write("selected") End if%>>文字效果</OPTION>
                <OPTION value="2" <%If RsConfig("MarkType") = "2" Then Response.Write("selected") End if%>>图片效果</OPTION>
              </SELECT> </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="23" align="right">水印文字信息（可为空）</td>
            <td> <INPUT TYPE="text" NAME="MarkText" size=40 value="<%=RsConfig("MarkText")%>"> 
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="23" align="right">水印字体大小</td>
            <td> <INPUT TYPE="text" NAME="MarkFontSize" size=10 value="<%=RsConfig("MarkFontSize")%>"> 
              <b>px</b> </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="23" align="right">水印字体颜色</td>
            <td> <input type="text" name="MarkFontColor" maxlength = 7 size = 7 id="MarkFontColor" value="<%=RsConfig("MarkFontColor")%>" readonly> 
              <img border=0 id="MarkFontColorShow" src="../../images/rect.gif" style="cursor:pointer;background-Color:<%=RsConfig("MarkFontColor")%>;" onClick="GetColor(this,'MarkFontColor');" title="选取颜色!"> 
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="23" align="right">水印字体名称</td>
            <td> <SELECT name="MarkFontName" id="MarkFontName">
                <option value="宋体" <%If RsConfig("MarkFontName") = "宋体" Then Response.Write("selected") End if%>>宋体</option>
                <option value="楷体_GB2312" <%If RsConfig("MarkFontName") = "楷体_GB2312" Then Response.Write("selected") End if%>>楷体</option>
                <option value="新宋体" <%If RsConfig("MarkFontName") = "新宋体" Then Response.Write("selected") End if%>>新宋体</option>
                <option value="黑体" <%If RsConfig("MarkFontName") = "黑体" Then Response.Write("selected") End if%>>黑体</option>
                <option value="隶书" <%If RsConfig("MarkFontName") = "隶书" Then Response.Write("selected") End if%>>隶书</option>
                <OPTION value="Andale Mono" <%If RsConfig("MarkFontName") = "Andale Mono" Then Response.Write("selected") End if%>>Andale 
                Mono</OPTION>
                <OPTION value="Arial" <%If RsConfig("MarkFontName") = "Arial" Then Response.Write("selected") End if%>>Arial</OPTION>
                <OPTION value="Arial Black" <%If RsConfig("MarkFontName") = "Arial Black" Then Response.Write("selected") End if%>>Arial 
                Black</OPTION>
                <OPTION value="Book Antiqua" <%If RsConfig("MarkFontName") = "Book Antiqua" Then Response.Write("selected") End if%>>Book 
                Antiqua</OPTION>
                <OPTION value="Century Gothic" <%If RsConfig("MarkFontName") = "Century Gothic" Then Response.Write("selected") End if%>>Century 
                Gothic</OPTION>
                <OPTION value="Comic Sans MS" <%If RsConfig("MarkFontName") = "Comic Sans MS" Then Response.Write("selected") End if%>>Comic 
                Sans MS</OPTION>
                <OPTION value="Courier New" <%If RsConfig("MarkFontName") = "Courier New" Then Response.Write("selected") End if%>>Courier 
                New</OPTION>
                <OPTION value="Georgia" <%If RsConfig("MarkFontName") = "Georgia" Then Response.Write("selected") End if%>>Georgia</OPTION>
                <OPTION value="Impact" <%If RsConfig("MarkFontName") = "Impact" Then Response.Write("selected") End if%>>Impact</OPTION>
                <OPTION value="Tahoma" <%If RsConfig("MarkFontName") = "Tahoma" Then Response.Write("selected") End if%>>Tahoma</OPTION>
                <OPTION value="Times New Roman" <%If RsConfig("MarkFontName") = "Times New Roman" Then Response.Write("selected") End if%>>Times 
                New Roman</OPTION>
                <OPTION value="Trebuchet MS" <%If RsConfig("MarkFontName") = "Trebuchet MS" Then Response.Write("selected") End if%>>Trebuchet 
                MS</OPTION>
                <OPTION value="Script MT Bold" <%If RsConfig("MarkFontName") = "Script MT Bold" Then Response.Write("selected") End if%>>Script 
                MT Bold</OPTION>
                <OPTION value="Stencil" <%If RsConfig("MarkFontName") = "Stencil" Then Response.Write("selected") End if%>>Stencil</OPTION>
                <OPTION value="Verdana" <%If RsConfig("MarkFontName") = "Verdana" Then Response.Write("selected") End if%>>Verdana</OPTION>
                <OPTION value="Lucida Console" <%If RsConfig("MarkFontName") = "Lucida Console" Then Response.Write("selected") End if%>>Lucida 
                Console</OPTION>
              </SELECT> </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="23" align="right">水印字体是否粗体</td>
            <td> <SELECT name="MarkFontBond" id="MarkFontBond">
                <OPTION value=0 <%If RsConfig("MarkFontBond") = "0" Then Response.Write("selected") End if%>>否</OPTION>
                <OPTION value=1 <%If RsConfig("MarkFontBond") = "1" Then Response.Write("selected") End if%>>是</OPTION>
              </SELECT> </td>
          </tr>
          <!-- 上传图片添加水印LOGO图片定义 -->
          <tr bgcolor="#FFFFFF"> 
            <td height="23" align="right">水印LOGO图片（可为空）<br> </td>
            <td> <INPUT TYPE="text" NAME="MarkPicture" size=40 value="<%=RsConfig("MarkPicture")%>">
              填写LOGO的图片相对路径 </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="23" align="right">水印LOGO图片透明度</td>
            <td> <INPUT TYPE="text" NAME="MarkOpacity" size=10 value="<%If Left(RsConfig("MarkOpacity"),1) = "." Then Response.Write("0"&RsConfig("MarkOpacity")) Else Response.Write(RsConfig("MarkOpacity")) End if%>">
              如60%请填写0.6 </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="23" align="right">水印图片去除底色</td>
            <td> <INPUT TYPE="text" NAME="MarkTranspColor" ID="MarkTranspColor" maxlength = 7 size = 7 value="<%=RsConfig("MarkTranspColor")%>"> 
              <img border=0 id="MarkTranspColorShow" src="../../images/rect.gif" style="cursor:pointer;background-Color:<%=RsConfig("MarkTranspColor")%>;" onClick="GetColor(this,'MarkTranspColor');" title="选取颜色!"> 
              保留为空则水印图片不去除底色。 </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="23" align="right">水印图片的长宽区域定义<br> </td>
            <td> 宽度
<INPUT TYPE="text" NAME="MarkWidth" size=10 value="<%=RsConfig("MarkWidth")%>">
              象素 高度： 
              <INPUT TYPE="text" NAME="MarkHeight" size=10 value="<%=RsConfig("MarkHeight")%>">
              象素 如水印图片的宽度和高度。 </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="23" align="right">水印LOGO位置坐标</td>
            <td> <SELECT NAME="MarkPosition" id="MarkPosition">
                <option value="1" <%If RsConfig("MarkPosition") = "1" Then Response.Write("selected") End if%>>左上</option>
                <option value="2" <%If RsConfig("MarkPosition") = "2" Then Response.Write("selected") End if%>>左下</option>
                <option value="3" <%If RsConfig("MarkPosition") = "3" Then Response.Write("selected") End if%>>居中</option>
                <option value="4" <%If RsConfig("MarkPosition") = "4" Then Response.Write("selected") End if%>>右上</option>
                <option value="5" <%If RsConfig("MarkPosition") = "5" Then Response.Write("selected") End if%>>右下</option>
              </SELECT> </td>
          </tr>
        </table></td>
    </tr>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="22"> <div align="right">发送邮件SMTP服务器</div></td>
      <td width="613" height="22" bgcolor="#F5F5F5"> <input name="MailServer" type="text" id="MailServer"  value="<%=RsConfig("MailServer")%>" size=30> 
      </td>
    </tr>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="21"> <div align="right">发送邮件SMTP服务器用户名</div></td>
      <td width="613" height="21" bgcolor="#F5F5F5"> <input name="MailName" type="text" id="MailName"  value="<%=RsConfig("MailName")%>" size=30></td>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="21"> <div align="right">发送邮件SMTP服务器密码</div></td>
      <td height="21"> <input name="MailPass" type="password" id="MailPass"  value="<%=RsConfig("MailPass")%>" size=30></td>
    <tr valign="middle" bgcolor="#F5F5F5">
      <td height="21"><div align="right">是否显示帮助</div></td>
      <td height="21"><input name="HelpTF" type="radio" value="0" <%if RsConfig("HelpTF")=0 then response.Write("checked") %>>
        否 
        <input type="radio" name="HelpTF" value="1"  <%if RsConfig("HelpTF")=1 then response.Write("checked") %>>
        是<font color=red>&nbsp;</font></td>
      
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="21"> <div align="right">评论是否审核后显示</div></td>
      <td height="21"> <input name="ReviewShow" type="checkbox" value="1" <%if RsConfig("ReviewShow")=1 then response.Write("checked")%>></td>
    <tr valign="middle" bgcolor="#f2f2f2"> 
      <td height="21" colspan="2"><strong>生成参数</strong></td>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="21"> <div align="right">新闻生成文件名样式</div></td>
      <td width="613" height="21" bgcolor="#F5F5F5"> <input name="NewsFileName1" type="checkbox" id="NewsFileName" value="Y" <%if InStr(1, RsConfig("NewsFileName"),"Y" ,1)<>0 then response.Write("checked") %>>
        年 
        <input name="NewsFileName2" type="checkbox" id="NewsFileName" value="M" <%if InStr(1, RsConfig("NewsFileName"),"M" ,1)<>0 then response.Write("checked") %>>
        月 
        <input name="NewsFileName3" type="checkbox" id="NewsFileName" value="D" <%if InStr(1, RsConfig("NewsFileName"),"D" ,1)<>0 then response.Write("checked") %>>
        日 
        <input name="NewsFileName4" type="checkbox" id="NewsFileName" value="H"  <%if InStr(1, RsConfig("NewsFileName"),"H" ,1)<>0 then response.Write("checked") %>>
        时 
        <input name="NewsFileName5" type="checkbox" id="NewsFileName" value="I"  <%if InStr(1, RsConfig("NewsFileName"),"I" ,1)<>0 then response.Write("checked") %>>
        分 
        <input name="NewsFileName6" type="checkbox" id="NewsFileName" value="S"  <%if InStr(1, RsConfig("NewsFileName"),"S" ,1)<>0 then response.Write("checked") %>>
        秒 
        <input name="NewsFileName7" type="checkbox" id="NewsFileName" value="A"  <%if InStr(1, RsConfig("NewsFileName"),"A" ,1)<>0 then response.Write("checked") %>>
        ClassID 
        <input name="NewsFileName8" type="checkbox" id="NewsFileName" value="N"  <%if InStr(1, RsConfig("NewsFileName"),"N" ,1)<>0 then response.Write("checked") %>>
        Newsid <br> <input name="NewsFileName9" type="checkbox" id="NewsFileName" value="Z"  <%if InStr(1, RsConfig("NewsFileName"),"Z" ,1)<>0 then response.Write("checked") %>>
        2位随机数 
        <input name="NewsFileName10" type="checkbox" id="NewsFileName" value="X" <%if InStr(1, RsConfig("NewsFileName"),"X" ,1)<>0 then response.Write("checked") %>>
        3位随机数 
        <input name="NewsFileName11" type="checkbox" id="NewsFileName" value="C" <%if InStr(1, RsConfig("NewsFileName"),"C" ,1)<>0 then response.Write("checked") %>>
        4位随机数 
        <input name="NewsFileName12" type="checkbox" id="NewsFileName" value="V"  <%if InStr(1, RsConfig("NewsFileName"),"V" ,1)<>0 then response.Write("checked") %>>
        5位随机数 
        <input name="NewsFileName13" type="checkbox" id="NewsFileName" value="U" <%if InStr(1, RsConfig("NewsFileName"),"U" ,1)<>0 then response.Write("checked") %>>
        分割&quot;_&quot; <br> <font color="#FF0000">采用任意组合,注意：如果你没有选择newsid,您必须选择至少3个（包含3个）以上的组合,随机数只能选择一个。否则默认识别“3位随机数”</font></td>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="21"> <div align="right">使用日期目录(年-月/日/文件名)</div></td>
      <td width="613" height="21" bgcolor="#F5F5F5"> <input name="UseDatePath" type="radio" value="0" <%if RsConfig("UseDatePath")=0 then response.Write("checked") %>>
        不使用 
        <input type="radio" name="UseDatePath" value="1"  <%if RsConfig("UseDatePath")=1 then response.Write("checked") %>>
        使用 </td>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="21"> <div align="right">自动生成分类</div></td>
      <td width="613" height="21" bgcolor="#F5F5F5"> <input name="AutoClass" type="checkbox" id="AutoClass" value="1"  <%if RsConfig("AutoClass")=1 then response.Write("checked") %>></td>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="21"> <div align="right">自动生成首页</div></td>
      <td width="613" height="21" bgcolor="#F5F5F5"> <input name="Autoindex" type="checkbox" id="Autoindex" value="1"  <%if RsConfig("Autoindex")=1 then response.Write("checked") %> >
        　　生成首页的扩展名： 
        <select name="IndexExtName" id="IndexExtName">
          <option value="htm" <%if RsConfig("IndexExtName")="htm" then response.Write("selected")%>>htm</option>
          <option value="html" <%if RsConfig("IndexExtName")="html" then response.Write("selected")%>>html</option>
          <option value="shtml" <%if RsConfig("IndexExtName")="shtml" then response.Write("selected")%>>shtml</option>
          <option value="shtm" <%if RsConfig("IndexExtName")="shtm" then response.Write("selected")%>>shtm</option>
          <option value="asp" <%if RsConfig("IndexExtName")="asp" then response.Write("selected")%>>asp</option>
        </select> <input name="oldIndexExtName" type="hidden" id="oldIndexExtName" value="<%=RsConfig("IndexExtName")%>"> 
      </td>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="21"> <div align="right">生成方式</div></td>
      <td width="613" height="21" bgcolor="#F5F5F5"> <input name="MakeType" type="radio" value="0" <%if RsConfig("MakeType")=0 then response.Write("checked") %>>
        Fso(File System Object) 
        <input type="radio" name="MakeType" value="1"  <%if RsConfig("MakeType")=1 then response.Write("checked") %>>
        Adodb.Stream</td>
    <tr valign="middle" bgcolor="#f2f2f2"> 
      <td height="21" colspan="2"><strong>会员参数</strong></td>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="21"> <div align="right">切换会员系统</div></td>
      <td width="613" height="21" bgcolor="#F5F5F5"> <input name="MemberType" type="radio" value="0" <%if RsConfig("MemberType")=0 then response.Write("checked")%>>
        本系统会员系统 
        <input type="radio" name="MemberType" value="1" <%if RsConfig("MemberType")=1 then response.Write("checked")%>>
        动网论坛会员系统 
        <input type="radio" name="MemberType" value="2" <%if RsConfig("MemberType")=2 then response.Write("checked")%>>
        极限论坛会员系统</td>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="21"> <div align="right">会员登陆增加点数</div></td>
      <td width="613" height="21" bgcolor="#F5F5F5"> <input name="NumberLoginPoint" type="text" id="NumberLoginPoint" value="<%=RsConfig("NumberLoginPoint")%>" size="8">
        会员投稿审核后增加： 
        <input name="NumberContPoint" type="text" id="NumberContPoint" value="<%=RsConfig("NumberContPoint")%>" size="8"></td>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="21"><div align="right">注册会员留言内容最大字符数</div></td>
      <td height="21" bgcolor="#F5F5F5"> <input name="MaxContent" type="text" id="MaxContent" value="<%=RsConfig("MaxContent")%>"> 
        <font color="#FF0000">（请填写正整数）</font> </td>
    <tr valign="middle" bgcolor="#F5F5F5">
      <td height="21"><div align="right">会员发言及回复增加点数</div></td>
      <td height="21" bgcolor="#F5F5F5">
<input name="QPoint" type="text" id="QPoint" value="<%=RsConfig("QPoint")%>"></td>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="21">　 
        <div align="right">会员注册协议<font color="#FF0000"><br>
          可以使用html语法</font></div></td>
      <td width="613" height="21" bgcolor="#F5F5F5"> <textarea name="UserConfer" cols="60" rows="6" id="UserConfer"><%=RsConfig("UserConfer")%></textarea> 
        <font color="#FF0000">&nbsp; </font></td>
    <tr valign="middle" bgcolor="#F5F5F5"> 
      <td height="21">　 
        <div align="right">版权信息<br>
          <font color="#FF0000"> 可以使用html语法</font> </div></td>
      <td width="613" height="21" bgcolor="#F5F5F5"> <textarea name="Copyright" cols="60" rows="6" id="Copyright"><%=RsConfig("Copyright")%></textarea> 
      </td>
      <iframe width="260" height="165" id="colourPalette" src="selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></iframe>
  </table>
</form>
</body>
</html>
<%
RsConfig.close
set RsConfig =nothing
set conn=nothing
set DBC=nothing
Dim ComponentName(3),i
ComponentName(0) = "Persits.Jpeg"
ComponentName(1) = "wsImage.Resize"
ComponentName(2) = "SoftArtisans.ImageGen"
ComponentName(3) = "CreatePreviewImage.cGvbox"
%>
<script language="javascript">
var ComponentNameArray = new Array();
var ComponentInfoArray = new Array();
<%
	Dim ExpiredStr
	For i = 0 to UBound(ComponentName)
%>
ComponentNameArray[ComponentNameArray.length] = "<%= ComponentName(i)%>";
<%
		If IsObjInstalled(ComponentName(i)) Then
			If IsExpired(ComponentName(i)) Then
				ExpiredStr = "，但已<font color=red>过期</font>"
			else
				ExpiredStr = ""
			End if
%>
ComponentInfoArray[ComponentInfoArray.length] = "<font color='0076AE'> √</font>支持<%=ExpiredStr%>";
<%
		Else
%>
ComponentInfoArray[ComponentInfoArray.length] = "<font color='red'>×</font>不支持"
<%
		End if
	Next
%>
function ShowInfo(ComponentID)
{
	if(ComponentID == 0)
	{
		document.all.ComponentInfo.innerHTML = "";
		document.all.WaterMarkSetting.style.display = "none";
		document.all.colourPalette.style.visibility="hidden";
	}
	else
	{
		document.all.ComponentInfo.innerHTML = ComponentNameArray[ComponentID - 1] + ComponentInfoArray[ComponentID - 1];
		document.all.WaterMarkSetting.style.display = "";
	}
}
function ShowThumbnailInfo(ThumbnailComponentID)
{
	if(ThumbnailComponentID == 0)
	{
		document.all.ThumbnailComponentInfo.innerHTML = "";
		document.all.ThumbnailSetting.style.display = "none";
	}
	else
	{
		document.all.ThumbnailComponentInfo.innerHTML = ComponentNameArray[ThumbnailComponentID - 1] + ComponentInfoArray[ThumbnailComponentID - 1];
		document.all.ThumbnailSetting.style.display = "";
	}
}
function ShowThumbnailSetting(ThumbnailSettingid)	
{
	if(ThumbnailSettingid == 0)
	{
		document.all.ThumbnailSetting1.style.display = "none";
		document.all.ThumbnailSetting0.style.display = "";
	}
	else
	{
		document.all.ThumbnailSetting1.style.display = "";
		document.all.ThumbnailSetting0.style.display = "none";
	}
}
function GetColor(img_val,input_val)
{
	var obj = document.getElementById("colourPalette");
	ColorImg = img_val;
	ColorValue = document.getElementById(input_val);
	if (obj){
	obj.style.left = getOffsetLeft(ColorImg) + "px";
	obj.style.top = (getOffsetTop(ColorImg) + ColorImg.offsetHeight) + "px";
	if (obj.style.visibility=="hidden")
	{
	obj.style.visibility="visible";
	}else {
	obj.style.visibility="hidden";
	}
	}
}
function getOffsetTop(elm) {
	var mOffsetTop = elm.offsetTop;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent){
		mOffsetTop += mOffsetParent.offsetTop;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetTop;
}
function getOffsetLeft(elm) {
	var mOffsetLeft = elm.offsetLeft;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent) {
		mOffsetLeft += mOffsetParent.offsetLeft;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetLeft;
}
function setColor(color)
{
	if (ColorValue){ColorValue.value = color;}
	if (ColorImg){ColorImg.style.backgroundColor = color;}
	document.getElementById("colourPalette").style.visibility="hidden";
}

</script>



