<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
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
<title>ͼƬ�޸Ĺ���</title>
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
	<td height="23" align="right">ѡ��ͼƬ���������</td>
	<td> 
	<select name="MarkComponent" id="MarkComponent" onChange="ShowInfo(this.value)">
	<option value=1>AspJpeg���
	<option value=2>wsImage���
	<option value=3>SA-ImgWriter���
	</select><span id="ComponentInfo"></span>
	</td>
  	</tr>
	<tr bgcolor="#FFFFFF"> 		
    <td height="23" align="right">������Ϣ��</td>
	<td> <input type="text" name="Text" size=40 value="">���뷽ʽ��
		<SELECT NAME="MarkPosition" id="MarkPosition">
		<option value="1">����</option>
		<option value="2">����</option>
		<option value="3">����</option>
		<option value="4">����</option>
		<option value="5">����</option>
		</SELECT></td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
    <td height="23" align="right">�����С��</td>
	<td> <INPUT TYPE="text" NAME="FontSize" size=10 value="24">
        ���� </td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
		
    <td height="23" align="right">������ɫ��</td>
	<td> <input type="text" name="FontColor" maxlength = 7 size = 7 id="FontColor" value="#000000" readonly>
	<img border=0 id="MarkFontColorShow" src="../../images/rect.gif" style="cursor:pointer;background-Color:#000000;" onclick="GetColor(this,'FontColor');" title="ѡȡ��ɫ!">
	</td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
		
    <td height="23" align="right">�������ƣ�</td>
	<td> <SELECT name="FontName" id="FontName">
	<option value="����">����</option>
	<option value="����_GB2312">����</option>
	<option value="������">������</option>
	<option value="����">����</option>
	<option value="����">����</option>
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
    <td height="23" align="right">�����Ƿ���壺</td>
	<td> <SELECT name="FontBond" id="FontBond">
		<OPTION value=0>��</OPTION>
		<OPTION value=1>��</OPTION>
		</SELECT>
	</td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
    <td height="23" align="right">�����Ƿ�б�壺</td>
	<td> <SELECT name="FontItalic" id="FontItalic">
		<OPTION value=0>��</OPTION>
		<OPTION value=1>��</OPTION>
		</SELECT>
	</td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
    <td height="23" align="right">
        <input name="BgType" type="radio" value="1" checked>&nbsp;&nbsp;&nbsp;
      ����ɫ��</td>
	<td> <INPUT TYPE="text" NAME="BgColor" ID="BgColor" maxlength = 7 size = 7 value="#FFFFFF" readonly>
      <img border=0 id="MarkTranspColorShow" src="../../images/rect.gif" style="cursor:pointer;background-Color:#FFFFFF;" onclick="GetColor(this,'BgColor');" title="ѡȡ��ɫ!"> 
    </td>
	</tr>	<tr bgcolor="#FFFFFF"> 
		
    <td height="23" align="right"><input name="BgType" type="radio" value="2">
        ����ͼƬ��<br>
	  </td>
	<td> <INPUT TYPE="text" NAME="BgPicture" size=40 value=""><Input type="button" Value="ѡ��ͼƬ">
		<Select Name="AlianType"><option value=1>ƽ��<option value=2>����<option value=3>����</Select>
    </td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
      <td height="23" align="right">���������壺<br>
	  </td>
	<td> ��ȣ�<INPUT TYPE="text" NAME="MarkWidth" size=10 value=""> ����
	�߶ȣ�<INPUT TYPE="text" NAME="MarkHeight" size=10 value="">
        ����</td>
	</tr>
	<tr bgcolor="#FFFFFF"> 
      <td height="23" align="right">�ļ���Ϣ��<br>
	  </td>
	<td>
	�ļ�����<INPUT TYPE="text" NAME="FileName" size=10 value="">
	��չ����<Select NAME="FileExtName"><option value="jpg">jpg<option value="gif">gif<option value="bmp">bmp</Select>
	����·����<INPUT TYPE="text" NAME="Path"  value="">
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
ComponentInfoArray[ComponentInfoArray.length] = "<font color='0076AE'> ��</font>֧��";
<%
		Else
%>
ComponentInfoArray[ComponentInfoArray.length] = "<font color='red'>��</font>��֧��"
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