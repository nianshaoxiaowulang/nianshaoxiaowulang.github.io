<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
Dim ExtendEditFile
ExtendEditFile = ""
if SysRootDir = "" then
	ExtendEditFile = "/Inc/Templet_NotDelete.htm"
else
	ExtendEditFile = "/" & SysRootDir & "/Inc/Templet_NotDelete.htm"
end if
%>
<!--#include file="../../Inc/Session.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ǩ�༭��</title>
</head>
<link rel="stylesheet" href="Editer.css">
<script language="JavaScript" src="Editer.js"></script>
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<body onLoad="return LoadEditFile();">
<table width="100%" height="30" border="0" cellpadding="0" cellspacing="0" id="Toolbar">
  <tr> 
    <td><table height="30" border="0" cellpadding="0" cellspacing="0" class="ToolSet">
        <tr> 
          <td width="30" align="center"><img title="ɾ������HTML��ʶ" onClick="DelAllHtmlTag()" class="Btn" src="../Images/Editer/geshi.gif" ></td>
          <td width="30" align="center"><img title="ɾ�����ָ�ʽ" onClick="Format('removeformat')" class="Btn" src="../Images/Editer/clear.gif" ></td>
		  <td width="30" align="center"><img src="../Images/Editer/TextColor.gif" width="23" height="22" class="Btn" title="���ֱ���ɫ" onClick="TextColor()" ></td>
		  <td width="30" align="center"><img title="���ֱ���ɫ" onClick="TextBGColor()" class="Btn" src="../Images/Editer/fgbgcolor.gif" ></td>
          <td width="30" align="center"><img title="���뻻�з���" onClick="InsertBR(0)" class="Btn" src="../Images/Editer/chars.gif" ></td>
          <td width="30" align="center"><img title="����ͼƬ��֧�ָ�ʽΪ��jpg��gif��bmp��png��" onClick="InsertPicture()" class="Btn" src="../Images/Editer/img.gif" ></td>
          <td width="30" align="center"><img src="../Images/Editer/url.gif" width="23" height="22" class="Btn" title="���볬������" onClick="InsertHref('CreateLink')" ></td>
          <td width="30" align="center"><img src="../Images/Editer/nourl.gif" width="23" height="22" class="Btn" title="ȡ����������" onClick="InsertHref('unLink')" ></td>
          <td width="1" align="center"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30" align="center"><img title="�����" onClick="Format('justifyleft')" class="Btn" src="../Images/Editer/Aleft.gif" ></td>
          <td width="30" align="center"><img title="����" onClick="Format('justifycenter')" class="Btn" src="../Images/Editer/Acenter.gif" ></td>
          <td width="30" align="center"><img title="�Ҷ���" onClick="Format('justifyright')" class="Btn" src="../Images/Editer/Aright.gif" ></td>
          <td width="1" align="center"> <div align="center" class="ToolSeparator"></div></td>
          <td width="26" align="center"><img src="../Images/Editer/Inserttable.gif"  class="Btn" title="������" onClick="InsertTable()"></td>
          <td width="26" align="center"><img src="../Images/Editer/inserttable1.gif" width="23" height="22"  class="Btn" title="������" onClick="InsertRow()"></td>
          <td width="26" align="center"><img src="../Images/Editer/delinserttable1.gif" width="23" height="22"  class="Btn" title="ɾ����" onClick="DeleteRow()"></td>
          <td width="26" align="center"><img src="../Images/Editer/inserttablec.gif" width="23" height="22"  class="Btn" title="������" onClick="InsertColumn()"></td>
          <td width="26" align="center"><img src="../Images/Editer/delinserttablec.gif" width="23" height="22"  class="Btn" title="ɾ����" onClick="DeleteColumn()"></td>
          <td style="display:none;" width="26" align="center"><img src="../Images/Editer/insterttable2.gif" width="23" height="22"  class="Btn" title="���뵥Ԫ��" onClick="InsertCell()"></td>
          <td style="display:none;" width="26" align="center"><img src="../Images/Editer/delinsterttable2.gif" width="23" height="22"  class="Btn" title="ɾ����Ԫ��" onClick="DeleteCell()"></td>
          <td width="26" align="center"><img src="../Images/Editer/MargeTD.gif" width="23" height="22"  class="Btn" title="�ϲ���" onClick="MergeColumn()"></td>
          <td width="26" align="center"><img src="../Images/Editer/Hbtable.gif" width="23" height="22"  class="Btn" title="�ϲ���" onClick="MergeRow()"></td>
          <td width="23" align="center"><img src="../Images/Editer/cftable.gif" width="23" height="22"  class="Btn" title="�����" onClick="SplitRows()"></td>
          <td width="23" align="center"><img src="../Images/Editer/SplitTD.gif" width="23" height="22"  class="Btn" title="�����" onClick="SplitColumn()"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="30"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" class="ToolSet">
        <tr>
          <td id="ShowObject">&nbsp;</td>
		  <td width="30"><div align="center"><img src="../Images/Editer/tablemodify.gif" width="23" height="22"  class="Btn" title="����" onClick="ExeEditAttribute()"></div></td>
          <td width="30"><div align="center"><img src="../Images/Editer/delLable.gif" width="23" height="22"  class="Btn" title="ɾ����ǩ" onClick="DeleteHTMLTag();"></div></td>
		</tr>
      </table></td>
  </tr>
  <tr><td>
  <iframe src="<% = ExtendEditFile %>" name="EditArea" ID="EditArea" MARGINHEIGHT="1" MARGINWIDTH="1" width="100%" scrolling="yes">
</iframe></td></tr>
   <tr> 
    <td height="20" id="SetModeArea"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="60" height="20" align="center" class="ModeBarBtnOff" id=Editer_CODE onClick="setTempletMode('CODE');"><img src="../Images/Editer/CodeMode.GIF" width="50" height="15"></td>
          <td style="display:none;" width="60" height="20" align="center" class="ModeBarBtnOff" id=Editer_VIEW onClick="setTempletMode('VIEW');"><img src="../Images/Editer/PreviewMode.gif" width="50" height="15"></td>
          <td width="60" height="20" align="center" class="ModeBarBtnOn" id=Editer_EDIT onClick="setTempletMode('EDIT');"><img src="../Images/Editer/EditMode.GIF" width="50" height="15"></td>
          <td style="display:none;" width="60" height="20" align="center" class="ModeBarBtnOff" id=Editer_TEXT onClick="setTempletMode('TEXT');"><img src="../Images/Editer/TextMode.GIF" width="50" height="15"></td>
          <td height="20">&nbsp;</td>
          <td style="display:none;" width="30" height="20" align="center" onClick="AddHeight();"><img class="Btn" src="../Images/Editer/AddHeight.gif" width="23" height="22"></td>
          <td style="display:none;" width="30" height="20" align="center" onClick="MinusHeight();"><img class="Btn" src="../Images/Editer/MinusHeight.gif" width="23" height="22"></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
var DocumentReadyTF=false;
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	SetEditAreaHeight();
	SetBodyStyle();
	DocumentReadyTF=true;
}
function SetEditAreaHeight()
{
	var BodyHeight=document.body.clientHeight;
	var EditAreaHeight=BodyHeight-100//document.all.Toolbar.height;
	document.all.EditArea.height=EditAreaHeight;
}
function SetBodyStyle()
{
	//EditArea.document.body.runtimeStyle.fontSize='10pt';
}
</script>