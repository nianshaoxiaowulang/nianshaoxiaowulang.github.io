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
%>
<!--#include file="../../Inc/Session.asp" -->
<%
Dim LimitUpFileFlag
LimitUpFileFlag = Request("LimitUpFileFlag")
Set Conn = Nothing
%>
<HTML>
<HEAD>
<TITLE>����ͼƬ�ļ�</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../../CSS/FS_css.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<script language="JavaScript">
function OK()
{
  var str1="";
  var strurl=document.PicForm.url.value;
  if (strurl==""||strurl=="http://")
  {
  	alert("��������ͼƬ��ַ�������ϴ�ͼƬ��");
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
        <LEGEND align=left>����ͼƬ����</LEGEND>
        <table border="0" align=center cellpadding="0" cellspacing="3">
          <tr> 
            <td colspan="2">ͼƬ��ַ�� 
              <input name="url" id="url" value='http://' size=30 maxlength="200">
              <input type="button" name="Button" value="ѡ��ͼƬ" onClick="var TempReturnValue=OpenWindow('../FunPages/SelectPic.asp?LimitUpFileFlag='+'<% = LimitUpFileFlag %>'+'&CurrPath=/<% = UpFiles %>',500,290,window);if (TempReturnValue!='') document.PicForm.url.value='<% = DirectoryRoot %>'+TempReturnValue;" class=Anbutc> 
            </td>
          </tr>
          <tr> 
            <td> ˵�����֣� 
              <input name="alttext" id=alttext size=20 maxlength="100"> </td>
            <td>ͼƬ�߿� 
              <input name="PicBorder" id="PicBorder" ONKEYPRESS="event.returnValue=IsDigit();"  value="0" size=5 maxlength="2">
              ���� </td>
          </tr>
          <tr> 
            <td> ����Ч���� 
              <select name="styletype" id=styletype>
                <option selected>��Ӧ��</option>
                <option value="filter:Alpha(Opacity=50)">��͸��Ч��</option>
                <option value="filter:Alpha(Opacity=0, FinishOpacity=100, Style=1, StartX=0, StartY=0, FinishX=100, FinishY=140)">����͸��Ч��</option>
                <option value="filter:Alpha(Opacity=10, FinishOpacity=100, Style=2, StartX=30, StartY=30, FinishX=200, FinishY=200)">����͸��Ч��</option>
                <option value="filter:blur(add=1,direction=14,strength=15)">ģ��Ч��</option>
                <option value="filter:blur(add=true,direction=45,strength=30)">�綯ģ��Ч��</option>
                <option value="filter:Wave(Add=0, Freq=60, LightStrength=1, Phase=0, Strength=3)">���Ҳ���Ч��</option>
                <option value="filter:gray">�ڰ���ƬЧ��</option>
                <option value="filter:Chroma(Color=#FFFFFF)">��ɫΪ͸��</option>
                <option value="filter:DropShadow(Color=#999999, OffX=7, OffY=4, Positive=1)">Ͷ����ӰЧ��</option>
                <option value="filter:Shadow(Color=#999999, Direction=45)">��ӰЧ��</option>
                <option value="filter:Glow(Color=#ff9900, Strength=5)">����Ч��</option>
                <option value="filter:flipv">��ֱ��ת��ʾ</option>
                <option value="filter:fliph">���ҷ�ת��ʾ</option>
                <option value="filter:grays">���Ͳ�ɫ��</option>
                <option value="filter:xray">X����ƬЧ��</option>
                <option value="filter:invert">��ƬЧ��</option>
              </select> </td>
            <td>ͼƬλ�ã� 
              <select name="aligntype" id=aligntype>
                <option selected>Ĭ��λ�� 
                <option value="left">���� 
                <option value="right" >���� 
                <option value="top">���� 
                <option value="middle">�в� 
                <option value="bottom">�ײ� 
                <option value="absmiddle">���Ծ��� 
                <option value="absbottom">���Եײ� 
                <option value="baseline">���� 
                <option value="texttop">�ı����� </select></td>
          </tr>
          <tr> 
            <td>ͼƬ��ȣ� 
              <input name="width" id=width2  ONKEYPRESS="event.returnValue=IsDigit();" size=4 maxlength="4">
              ����</td>
            <td>ͼƬ�߶ȣ� 
              <input name="height" id="height3" onKeyPress="event.returnValue=IsDigit();" size=4 maxlength="4">
              ����</td>
          </tr>
          <tr> 
            <td>���¼�ࣺ 
              <input name="vspace" id=vspace  ONKEYPRESS="event.returnValue=IsDigit();" value="0" size=4 maxlength="2">
              ����</td>
            <td>���Ҽ�ࣺ 
              <input name="hspace" id=hspace onKeyPress="event.returnValue=IsDigit();"  value="0" size=4 maxlength="2">
              ����</td>
          </tr>
        </table>
        <br>
        <br>
        </fieldset></td>
      <td width=80 align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="40"> <div align="center"> 
                <input name="cmdOK" type="button" id="cmdOK3" value="  ȷ��  " onClick="OK();">
                <input name="UpFileName" type="hidden" id="UpFileName3" value="None">
              </div></td>
          </tr>
          <tr> 
            <td height="40"> <div align="center"> 
                <input name="cmdCancel" type=button id="cmdCancel3" onClick="window.close();" value='  ȡ��  '>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
</body>
</html>

