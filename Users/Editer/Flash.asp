<%@language=vbscript codepage=936 %>
<% Option Explicit %>
<!--#include file="../../Inc/Function.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<%
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
Dim DirectoryRoot
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
DirectoryRoot = GetConfigDoMain
Set Conn = Nothing
Dim LimitUpFileFlag
LimitUpFileFlag = Request("LimitUpFileFlag")
%>
<HTML>
<HEAD>
<TITLE>插入FLASH文件</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="PublicJS.js"></script>
<script language="JavaScript">
function OK()
{
  var str1="";
  var strurl=document.FlashForm.url.value;
  if (!IsExt(strurl,'swf'))
  {
	  alert('文件类型不对，请重新选择！');
	  return;
  }
  if (strurl==""||strurl=="http://")
  {
  	alert("请先输入FLASH文件地址，或者上传FLASH文件！");
	document.FlashForm.url.focus();
	return false;
  }
  else
  {
    str1="<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="+document.FlashForm.width.value+" height="+document.FlashForm.height.value+"><param name=movie value="+document.FlashForm.url.value+"><param name=quality value=high><embed src="+document.FlashForm.url.value+" pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width="+document.FlashForm.width.value+" height="+document.FlashForm.height.value+"></embed></object>"
    window.returnValue = str1+"$$$"+document.FlashForm.UpFileName.value;
    window.close();
  }
}

function IsExt(url,opt)
{  
	var sTemp; 
	var b=false; 
	var s=opt.toUpperCase().split("|");  
	for (var i=0;i<s.length ;i++ ) 
	{ 
		sTemp=url.substr(url.length-s[i].length-1); 
		sTemp=sTemp.toUpperCase();
		s[i]="."+s[i];
		if (s[i]==sTemp)
		{
			b=true;
			break;
		}
	}
	return b;
}

function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>
</head>
<BODY bgColor=menu topmargin=15 leftmargin=15 >
<form name="FlashForm" method="post" action="">
  <table width=100% border="0" cellpadding="0" cellspacing="2">
    <tr>
      <td> <FIELDSET align=left>
        <LEGEND align=left>FLASH动画参数</LEGEND>
        <TABLE border="0" cellpadding="0" cellspacing="3" >
          <TR>
            <TD height="17" >地址： <INPUT name="url" id=url value="http://" size=30>
            </td>
          </TR>
          <TR>
            <TD >宽度：
              <INPUT name="width" id=width ONKEYPRESS="event.returnValue=IsDigit();" value=500 size=7 maxlength="4"> 
              &nbsp;&nbsp;高度：
              <INPUT name="height" id=height ONKEYPRESS="event.returnValue=IsDigit();" value=300 size=7 maxlength="4"></TD>
          </TR>
        </TABLE>
        </fieldset></td>
      <td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  确定  " onClick="OK();"> 
        <br> <br>
        <input name="cmdCancel" type=button id="cmdCancel" onclick="window.close();" value='  取消  '> 
        <input name="UpFileName" type="hidden" id="UpFileName2" value="None"></td>
    </tr>
  </table>
</form>
</body>
</html>
