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
%>
<!--#include file="../../Inc/Session.asp" -->
<%
DirectoryRoot = GetConfigDoMain
Set Conn = Nothing
Dim LimitUpFileFlag
LimitUpFileFlag = Request("LimitUpFileFlag")
%>
<HTML><HEAD><TITLE>������Ƶ�ļ�</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../../CSS/FS_css.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<script language="JavaScript">
function OK(){
  var str1="";
  var strurl=document.VideoForm.url.value;
  if (!IsExt(strurl,'avi|wmv|asf|mpg|mp3'))
  {
	  alert('�ļ����Ͳ��ԣ�������ѡ��');
	  document.VideoForm.url.focus();
	  return;
  }
  if (strurl==""||strurl=="http://")
  {
  	alert("����������Ƶ�ļ���ַ�������ϴ���Ƶ�ļ���");
	document.VideoForm.url.focus();
	return false;
  }
  else
  {
	str1=str1+"<embed src="+document.VideoForm.url.value+" width="+document.VideoForm.width.value+"height="+document.VideoForm.height.value+" autostart=true loop=true>"
	str1=str1+"</embed>"
	window.returnValue = str1+"$$$"+document.VideoForm.UpFileName.value;
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
<form name="VideoForm" method="post" action="">
  <table width=100% border="0" cellpadding="0" cellspacing="2">
    <tr>
      <td> <FIELDSET align=left>
        <LEGEND align=left>��Ƶ�ļ�����</LEGEND>
        <TABLE border="0" cellpadding="0" cellspacing="3">
          <TR>
            <TD >��ַ��
              <INPUT name="url" id=url  value="http://" size=30>
              <input type="button" name="Button" value="ѡ����Ƶ�ļ�" onClick="var TempReturnValue=OpenWindow('../FunPages/SelectPic.asp?LimitUpFileFlag='+'<% = LimitUpFileFlag %>'+'&CurrPath=/<% = UpFiles %>',500,290,window);if (TempReturnValue!='') document.VideoForm.url.value='<% = DirectoryRoot %>'+TempReturnValue;" class=Anbutc> 
            </td>
          </TR>
          <TR>
            <TD >��ȣ�
              <INPUT name="width" id=width  ONKEYPRESS="event.returnValue=IsDigit();" value=352 size=14 maxlength="4"> 
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�߶ȣ�
              <INPUT id=height ONKEYPRESS="event.returnValue=IsDigit();" value=288 size=14 maxlength="4"></TD>
          </TR>
          <TR>
            <TD align=center>֧�ָ�ʽΪ��avi��wmv��mpg��asf��mp3</TD>
          </TR>
        </TABLE>
        </fieldset></td>
      <td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  ȷ��  " onClick="OK();"> 
        <br> <br> <input name="UpFileName" type="hidden" id="UpFileName2" value="None"> 
        <input name="cmdCancel" type=button id="cmdCancel" onclick="window.close();" value='  ȡ��  '></td>
    </tr>
  </table>
</form>
</body>
</html>
