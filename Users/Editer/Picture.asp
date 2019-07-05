<%@language=vbscript codepage=936 %>
<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Function.asp" -->
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
Dim LimitUpFileFlag
LimitUpFileFlag = Request("LimitUpFileFlag")
Set Conn = Nothing
%>
<HTML>
<HEAD>
<TITLE>插入图片文件</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<script language="JavaScript" src="PublicJS.js"></script>
<script language="JavaScript">
function OK()
{
  var str1="";
  var strurl=document.PicForm.url.value;
  if (strurl==""||strurl=="http://")
  {
  	alert("请先输入图片地址，或者上传图片！");
	document.PicForm.url.focus();
	return false;
  }
  else
  {
    str1="<img src='"+document.PicForm.url.value+"' alt='"+document.PicForm.alttext.value+"' ";
    if(document.PicForm.width.value!=''&&document.PicForm.width.value!='0') str1=str1+"width='"+document.PicForm.width.value+"' ";
    if(document.PicForm.height.value!=''&&document.PicForm.height.value!='0') str1=str1+"height='"+document.PicForm.height.value+"' ";
    str1=str1+"border='"+document.PicForm.PicBorder.value+"' align='"+document.PicForm.aligntype.value+"' ";
	if(document.PicForm.vspace.value!=''&&document.PicForm.vspace.value!='0') str1=str1+"vspace='"+document.PicForm.vspace.value+"' ";
	if(document.PicForm.hspace.value!=''&&document.PicForm.hspace.value!='0') str1=str1+"hspace='"+document.PicForm.hspace.value+"' ";
	if(document.PicForm.styletype.value!='')	str1=str1+"style='"+document.PicForm.styletype.value+"'";
    str1=str1+">";
    window.returnValue=str1+"$$$"+document.PicForm.UpFileName.value;
    window.close();
  }
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
<form name="PicForm" method="post" action="">
  <table width=100% border="0" align="center" cellpadding="0" cellspacing="2">
    <tr>
      <td> <FIELDSET align=left>
        <LEGEND align=left>输入图片参数</LEGEND>
        <table border="0" align=center cellpadding="0" cellspacing="3">
          <tr> 
            <td colspan="2">图片地址： 
              <input name="url" id="url" value='http://' size=30 maxlength="200">
               
            </td>
          </tr>
          <tr> 
            <td> 说明文字： 
              <input name="alttext" id=alttext size=20 maxlength="100"> </td>
            <td>图片边框： 
              <input name="PicBorder" id="PicBorder" ONKEYPRESS="event.returnValue=IsDigit();"  value="0" size=5 maxlength="2">
              像素 </td>
          </tr>
          <tr> 
            <td> 特殊效果： 
              <select name="styletype" id=styletype>
                <option selected>不应用</option>
                <option value="filter:Alpha(Opacity=50)">半透明效果</option>
                <option value="filter:Alpha(Opacity=0, FinishOpacity=100, Style=1, StartX=0, StartY=0, FinishX=100, FinishY=140)">线型透明效果</option>
                <option value="filter:Alpha(Opacity=10, FinishOpacity=100, Style=2, StartX=30, StartY=30, FinishX=200, FinishY=200)">放射透明效果</option>
                <option value="filter:blur(add=1,direction=14,strength=15)">模糊效果</option>
                <option value="filter:blur(add=true,direction=45,strength=30)">风动模糊效果</option>
                <option value="filter:Wave(Add=0, Freq=60, LightStrength=1, Phase=0, Strength=3)">正弦波纹效果</option>
                <option value="filter:gray">黑白照片效果</option>
                <option value="filter:Chroma(Color=#FFFFFF)">白色为透明</option>
                <option value="filter:DropShadow(Color=#999999, OffX=7, OffY=4, Positive=1)">投射阴影效果</option>
                <option value="filter:Shadow(Color=#999999, Direction=45)">阴影效果</option>
                <option value="filter:Glow(Color=#ff9900, Strength=5)">发光效果</option>
                <option value="filter:flipv">垂直翻转显示</option>
                <option value="filter:fliph">左右翻转显示</option>
                <option value="filter:grays">降低彩色度</option>
                <option value="filter:xray">X光照片效果</option>
                <option value="filter:invert">底片效果</option>
              </select> </td>
            <td>图片位置： 
              <select name="aligntype" id=aligntype>
                <option selected>默认位置 
                <option value="left">居左 
                <option value="right" >居右 
                <option value="top">顶部 
                <option value="middle">中部 
                <option value="bottom">底部 
                <option value="absmiddle">绝对居中 
                <option value="absbottom">绝对底部 
                <option value="baseline">基线 
                <option value="texttop">文本顶部 </select></td>
          </tr>
          <tr> 
            <td>图片宽度： 
              <input name="width" id=width2  ONKEYPRESS="event.returnValue=IsDigit();" size=4 maxlength="4">
              像素</td>
            <td>图片高度： 
              <input name="height" id="height3" onKeyPress="event.returnValue=IsDigit();" size=4 maxlength="4">
              像素</td>
          </tr>
          <tr> 
            <td>上下间距： 
              <input name="vspace" id=vspace  ONKEYPRESS="event.returnValue=IsDigit();" value="0" size=4 maxlength="2">
              像素</td>
            <td>左右间距： 
              <input name="hspace" id=hspace onKeyPress="event.returnValue=IsDigit();"  value="0" size=4 maxlength="2">
              像素</td>
          </tr>
        </table>
        <br>
        <br>
        </fieldset></td>
      <td width=80 align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="40"> <div align="center"> 
                <input name="cmdOK" type="button" id="cmdOK3" value="  确定  " onClick="OK();">
                <input name="UpFileName" type="hidden" id="UpFileName3" value="None">
              </div></td>
          </tr>
          <tr> 
            <td height="40"> <div align="center"> 
                <input name="cmdCancel" type=button id="cmdCancel3" onClick="window.close();" value='  取消  '>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
</body>
</html>

