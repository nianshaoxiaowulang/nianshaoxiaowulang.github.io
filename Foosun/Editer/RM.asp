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
%>
<!--#include file="../../Inc/Session.asp" -->
<%
Dim LimitUpFileFlag
LimitUpFileFlag = Request("LimitUpFileFlag")
Set Conn = Nothing
%>
<HTML><HEAD><TITLE>����RealPlay�ļ�</TITLE>
<link rel="stylesheet" type="text/css" href="../../CSS/FS_css.css">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<script language="JavaScript">
function OK()
{
  var str1="";
  var strurl=document.RmForm.url.value;
  if (!IsExt(strurl,'rm|ra|ram|mp3|rmvb|avi'))
  {
	  alert('�ļ����Ͳ��ԣ�������ѡ��');
	  document.RmForm.url.focus();
	  return;
  }
  if (strurl==""||strurl=="http://")
  {
  	alert("��������RealPlay�ļ���ַ�������ϴ�RealPlay�ļ���");
	document.RmForm.url.focus();
	return false;
  }
  else
  {
		str1="<object classid=clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa width="+document.RmForm.width.value+" height="+document.RmForm.height.value+">"
		str1=str1+"<param name=src value="+document.RmForm.url.value+">"
		str1=str1+"<param name=console value=clip1><param name=controls value=imagewindow>"
		str1=str1+"<param name=autostart value=true>"
		str1=str1+"</object><object classid=clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa height=60 width="+document.RmForm.width.value+">"
		str1=str1+"<param name=src value="+document.RmForm.url.value+">"
		str1=str1+"<param name=controls value=controlpanel><param name=console value=clip1></object>"  
		window.returnValue = str1+"$$$"+document.RmForm.UpFileName.value;
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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY bgColor=menu topmargin=15 leftmargin=15 >
<form name="RmForm" method="post" action="">
  <table width=100% border="0" cellpadding="0" cellspacing="2">
    <tr>
      <td> <FIELDSET align=left>
        <LEGEND align=left>RealPlay�ļ�����</LEGEND>
        <TABLE border="0" cellpadding="0" cellspacing="3">
          <TR>
            <TD >��ַ�� <INPUT name="url" id=url  value="http://" size=30>
              <input type="button" name="Button" value="ѡ���ļ�" onClick="var TempReturnValue=OpenWindow('../FunPages/SelectPic.asp?LimitUpFileFlag='+'<% = LimitUpFileFlag %>'+'&CurrPath=/<% = UpFiles %>',500,290,window);if (TempReturnValue!='') document.RmForm.url.value='<% = DirectoryRoot %>'+TempReturnValue;" class=Anbutc> 
            </td>
          </TR>
          <TR>
            <TD >��ȣ�
              <INPUT name="width" id=width ONKEYPRESS="event.returnValue=IsDigit();" value=500 size=7 maxlength="4" > 
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�߶ȣ�
              <INPUT name="height" id=height ONKEYPRESS="event.returnValue=IsDigit();" value=300 size=7 maxlength="4"></TD>
          </TR>
          <TR>
            <TD align=center>֧�ָ�ʽΪ��rm��ra��ram��rmvb��mp3��avi</TD>
          </TR>
        </TABLE>
        </fieldset></td>
      <td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  ȷ��  " onClick="OK();"> 
        <br> <br> <input name="cmdCancel" type=button id="cmdCancel" onclick="window.close();" value='  ȡ��  '> 
        <input name="UpFileName" type="hidden" id="UpFileName2" value="None"></td>
    </tr>
  </table>
</form>
</body>
</html>