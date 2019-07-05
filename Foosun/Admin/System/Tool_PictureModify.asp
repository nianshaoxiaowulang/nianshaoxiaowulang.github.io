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
'if Not JudgePopedomTF(Session("Name"),"P040501") then Call ReturnError1()
%>
<html>
<title>图片修改工具</title>
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
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" scroll=yes onload="ShowInfo(1)">
<form name="form1" method="post" action="">
<table width="808" border="0" cellpadding="0" cellspacing="1"  bordercolor="e6e6e6" bgcolor="#E3E3E3">
  <tr bgcolor="#FFFFFF">
	<td height="23" align="right">选择图片处理组件：</td>
	<td> 
	<select name="MarkComponent" id="MarkComponent" onChange="ShowInfo(this.value)">
	<option value=1>AspJpeg组件
	<option value=2>wsImage组件
	<option value=3>SA-ImgWriter组件
	</select><span id="ComponentInfo"></span>
	</td>
  	</tr>
	<tr bgcolor="#FFFFFF"> 		
    <td height="23" align="right">文字信息：</td>
	<td> <input type="text" name="Text" size=40 value="">对齐方式：
		<SELECT NAME="MarkPosition" id="MarkPosition">
		<option value="1">左上</option>
		<option value="2">左下</option>
		<option value="3">居中</option>
		<option value="4">右上</option>
		<option value="5">右下</option>
		</SELECT></td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
    <td height="23" align="right">字体大小：</td>
	<td> <INPUT TYPE="text" NAME="FontSize" size=10 value="24">
        象素 </td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
		
    <td height="23" align="right">字体颜色：</td>
	<td> <input type="text" name="FontColor" maxlength = 7 size = 7 id="FontColor" value="#000000" readonly>
	<img border=0 id="MarkFontColorShow" src="../../images/rect.gif" style="cursor:pointer;background-Color:#000000;" onclick="GetColor(this,'FontColor');" title="选取颜色!">
	</td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
		
    <td height="23" align="right">字体名称：</td>
	<td> <SELECT name="FontName" id="FontName">
	<option value="宋体">宋体</option>
	<option value="楷体_GB2312">楷体</option>
	<option value="新宋体">新宋体</option>
	<option value="黑体">黑体</option>
	<option value="隶书">隶书</option>
	<OPTION value="Andale Mono">Andale Mono</OPTION> 
	<OPTION value="Arial">Arial</OPTION> 
	<OPTION value="Arial Black">Arial Black</OPTION> 
	<OPTION value="Book Antiqua">Book Antiqua</OPTION>
	<OPTION value="Century Gothic">Century Gothic</OPTION> 
	<OPTION value="Comic Sans MS">Comic Sans MS</OPTION>
	<OPTION value="Courier New">Courier New</OPTION>
	<OPTION value="Georgia">Georgia</OPTION>
	<OPTION value="Impact">Impact</OPTION>
	<OPTION value="Tahoma">Tahoma</OPTION>
	<OPTION value="Times New Roman">Times New Roman</OPTION>
	<OPTION value="Trebuchet MS">Trebuchet MS</OPTION>
	<OPTION value="Script MT Bold">Script MT Bold</OPTION>
	<OPTION value="Stencil">Stencil</OPTION>
	<OPTION value="Verdana">Verdana</OPTION>
	<OPTION value="Lucida Console">Lucida Console</OPTION>
	</SELECT>
	</td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
    <td height="23" align="right">字体是否粗体：</td>
	<td> <SELECT name="FontBond" id="FontBond">
		<OPTION value=0>否</OPTION>
		<OPTION value=1>是</OPTION>
		</SELECT>
	</td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
    <td height="23" align="right">字体是否斜体：</td>
	<td> <SELECT name="FontItalic" id="FontItalic">
		<OPTION value=0>否</OPTION>
		<OPTION value=1>是</OPTION>
		</SELECT>
	</td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
    <td height="23" align="right">
        <input name="BgType" type="radio" value="1" checked>&nbsp;&nbsp;&nbsp;
      背景色：</td>
	<td> <INPUT TYPE="text" NAME="BgColor" ID="BgColor" maxlength = 7 size = 7 value="#FFFFFF" readonly>
      <img border=0 id="MarkTranspColorShow" src="../../images/rect.gif" style="cursor:pointer;background-Color:#FFFFFF;" onclick="GetColor(this,'BgColor');" title="选取颜色!"> 
    </td>
	</tr>	<tr bgcolor="#FFFFFF"> 
		
    <td height="23" align="right"><input name="BgType" type="radio" value="2">
        背景图片：<br>
	  </td>
	<td> <INPUT TYPE="text" NAME="BgPicture" size=40 value=""><Input type="button" Value="选择图片">
		<Select Name="AlianType"><option value=1>平铺<option value=2>居中<option value=3>拉伸</Select>
    </td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
      <td height="23" align="right">长宽区域定义：<br>
	  </td>
	<td> 宽度：<INPUT TYPE="text" NAME="MarkWidth" size=10 value=""> 象素
	高度：<INPUT TYPE="text" NAME="MarkHeight" size=10 value="">
        象素</td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
      <td height="23" align="right">文件信息：<br>
	  </td>
	<td>
	文件名：<INPUT TYPE="text" NAME="FileName" size=10 value="">
	扩展名：<Select NAME="FileExtName"><option value="jpg">jpg<option value="gif">gif<option value="bmp">bmp</Select>
	保存路径：<INPUT TYPE="text" NAME="Path"  value="">
     </td>
	</tr>
  <iframe width="260" height="165" id="colourPalette" src="selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></iframe>
</table>
</form>
</body>
</html>
<%
Dim ComponentName(2),i
ComponentName(0) = "Persits.Jpeg"
ComponentName(1) = "wsImage.Resize"
ComponentName(2) = "SoftArtisans.ImageGen"
%>
<script language="javascript">
var ComponentNameArray = new Array();
var ComponentInfoArray = new Array();
<%
	For i = 0 to UBound(ComponentName)
%>
ComponentNameArray[ComponentNameArray.length] = "<%= ComponentName(i)%>";
<%
		If IsObjInstalled(ComponentName(i)) Then
%>
ComponentInfoArray[ComponentInfoArray.length] = "<font color='0076AE'> √</font>支持";
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
		document.all.colourPalette.style.visibility="hidden";
	}
	else
	{
		document.all.ComponentInfo.innerHTML = ComponentNameArray[ComponentID - 1] + ComponentInfoArray[ComponentID - 1];
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