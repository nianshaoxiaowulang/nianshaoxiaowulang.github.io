<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<%
Dim DBC,Conn,HelpConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = "DBQ=" + Server.MapPath("Foosun_help.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set HelpConn = DBC.OpenConnection()
Set DBC = Nothing

'==============================================================================
'������ƣ�FoosunHelp System Form FoosunCMS
'��ǰ�汾��Foosun Content Manager System 3.0 ϵ��
'���¸��£�2005.12
'==============================================================================
'��ҵע����ϵ��028-85098980-601,602 ����֧�֣�028-85098980-605��607,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��159410,394226379,125114015,655071
'����֧��:���г���ʹ�����⣬�����ʵ�bbs.foosun.net���ǽ���ʱ�ش���
'���򿪷�����Ѷ������ & ��Ѷ���������
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.net  ��ʾվ�㣺test.cooin.com    
'��վ����ר����www.cooin.com
'==============================================================================
'��Ѱ汾����������ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'==============================================================================
%>
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<%
Dim FuncName,FileName,PageField,HelpContent,HelpSingleContent

Dim Action,HelpID
HelpID = Request.QueryString("ID")

if isNumeric(HelpID)=false or HelpID="" Then
	if Not JudgePopedomTF(Session("Name"),"P070801") then Call ReturnError1()
	Action = "addnew"
	HelpID = 0
	FileName = session("FileName")
	FuncName = session("FuncName")
Else
	if Not JudgePopedomTF(Session("Name"),"P070802") then Call ReturnError1()
	Action = "modify"
	Dim tempRs
	Set tempRs = Server.CreateObject(G_FS_RS)
	tempRs.open "Select * From [Fs_Help] where id="&Clng(HelpID),HelpConn,1,1
	if not tempRs.eof then
		HelpID = tempRs("ID")
		FuncName = tempRs("FuncName")
		FileName = tempRs("FileName")
		PageField = tempRs("PageField")
		HelpContent = tempRs("HelpContent")
		HelpSingleContent = tempRs("HelpSingleContent")
	Else
		Response.write "<script>alert('û���ҵ�����޸�����');</script>"
		HelpID=0
		Action = "addnew"
		FileName = session("FileName")
		FuncName = session("FuncName")
	end if
	tempRs.close
	set tempRs = Nothing
end If
%>
<script src="../SysJS/PublicJS.js" language="JavaScript"></script>
<style type="text/css">
	td {line-height:20px;}
	.FS_EditorStyle {padding:0px; background: #FDFDDF ; font-size: 12px; font-family: Tahoma; font-style: oblique; line-height: normal; font-weight: bold;}
</style>
<link href="../../CSS/FS_css.css" rel="stylesheet">
<link rel="stylesheet" href="../Editer/Editer.css">

<body bgcolor="#FFFFFF" scroll=auto topmargin=2 leftmargin=2>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="3"></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="35" align="center" alt="����" onClick="if(BindSubmit())HelpForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
			<td width=2 class="Gray">|</td>
			<td width="35" align="center" alt="����" onClick="LoadSearchBox();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
			<td width=2 class="Gray">|</td>
		  	<td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
            <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="3" bgcolor="#FFFFFF"></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D9D9D9">
  <form action="SaveField.asp" method="post" name="HelpForm" onsubmit="return BindSubmit();">
    <input type="hidden" name="action" value="<%=Action%>">
    <input type="hidden" name="HelpID" value="<%=HelpID%>">
    <tr bgcolor="#FFFFFF"> 
      <td width="12%" align="center" bgcolor="#EFEFEF">ҳ�湦��</td>
      <td width="87%" bgcolor="#EFEFEF" style="padding:0px 5px; "> <input type="text" size=72 name="FuncName" value="<%=FuncName%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="center" bgcolor="#EFEFEF">ҳ���ַ</td>
      <td bgcolor="#EFEFEF" style="padding:0px 5px; "> <input type="text" size=72 name="FileName" value="<%=FileName%>">
        �� <input type="button" name="selectFileName" value="��ȡ����" title="��ȡ����ҳ���ַ�д��ڵİ�������" onClick="LoadHelpFileName(this.form);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="12%" align="center" bgcolor="#EFEFEF">ѡ��ؼ���</td>
      <td width="87%" bgcolor="#EFEFEF" style="padding:0px 5px; "> <input type="text" size="72" Name="NewPageField" value="<%=PageField%>">
        <br> <select Name="PageField" onChange="SetZone();" style="width:66%">
          <option value="">�½��ؼ���</option>
        </select> <font color=red>�������������ͬ�Ĺؼ���,���á�,���ָ�</font> <input name="HelpSingleContent" type="hidden" value="<%=server.HtmlEncode(HelpSingleContent)%>"> 
        <input name="HelpContent" type="hidden" value="<%=server.HtmlEncode(HelpContent)%>"></td>
    </tr>
    <tr> 
      <td colspan="2"><table width="100%" border="0" cellpadding="0" cellspacing="0" height="20" bgcolor="#FFFFFF">
          <tr> 
            <td width="60" height="26" align="center" bgcolor="#EFEFEF" class="LableSelected" id="SingleID" onClick="ChangeFolder(this);">�򵥽���</td>
            <td width="5" height="26" align="center" class="ToolBarButtonLine" style="cursor:default;">&nbsp;</td>
            <td width="60" height="26" align="center" class="LableDefault" id="ContentID" onClick="ChangeFolder(this);">��ϸ˵��</td>
            <td height="26" class="ToolBarButtonLine" style="cursor:default;">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#FFFFFF" id="SingleArea"> 
      <td height="380" colspan=2> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0" height="100%">
          <tr> 
            <td height="349"> 
              <iframe class="FS_EditorStyle" id='HelpSingleEditor' name="HelpSingleEditor" frameborder=0 scrolling=no width='100%' height='100%'></iframe></td>
          </tr>
          <tr> 
            <td height="17">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr bgcolor="#E8E8E8"> 
                  <td width="60" height="17" align="center" id="SingleHTML" class="ModeBarBtnOff" onClick="SetFrameMode(this,'HTML')"><img src="../Images/Editer/CodeMode.GIF" width="50" height="15"></td>
                  <td width="60" height="17" align="center" id="SingleEDIT" class="ModeBarBtnOn"  onClick="SetFrameMode(this,'EDIT')"><img src="../Images/Editer/EditMode.GIF" width="50" height="15"></td>
                  <td width="*">&nbsp;</td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#FFFFFF" id="ContentArea" style="display:none;"> 
      <td height="380" colspan=2> 
        <iframe id='HelpEditor' name="HelpEditor" src='../Editer/HelpEditer.asp' frameborder=0 scrolling=no width='100%' height='100%'></iframe></td>
    </tr>
  </form>
</table>
<iframe src="" id="FrameArea" name="FrameArea" width=0 height=0></iframe>

<SCRIPT LANGUAGE="JavaScript">
<!--
//��ʼ��������
var oSingleEditor,oHelpEditor;
if (document.all){
	oSingleEditor=frames["HelpSingleEditor"];
	oHelpEditor=frames["HelpEditor"];
}else{
	oSingleEditor=document.getElementById("HelpSingleEditor").contentWindow;
	oHelpEditor=document.getElementById("HelpEditor").contentWindow;
}

//document.HelpForm.attachEvent("onsubmit", BindSubmit);
InitSingleEditor(oSingleEditor,'HelpSingleContent');

	
//���ύ
function BindSubmit()
{
	if (oHelpEditor.CurrMode!='EDIT') {alert('����ģʽ���޷����棬���л����༭ģʽ');return false;}
	if(SingleEditorMode == "HTML") document.HelpForm.HelpSingleContent.value = oSingleEditor.document.body.innerText;
	else document.HelpForm.HelpSingleContent.value = oSingleEditor.document.body.innerHTML;
	document.HelpForm.HelpContent.value = oHelpEditor.EditArea.document.body.innerHTML;
	return true;
}
//��ʼ���򵥰���˵������
function InitSingleEditor(oIframe,HiddenID)
{
	var charset="gb2312";
	var InitValue=document.getElementById(HiddenID).value;
	if (navigator.appVersion.indexOf("MSIE 6.0",0)==-1){oIframe.document.designMode="On";}
	oIframe.document.open();
	oIframe.document.write('<html>');
	oIframe.document.write('<head>');
	oIframe.document.write('</head>');
	oIframe.document.write('<body bgcolor=\"#FFFFFF\" topmargin=4 leftmargin=4>');
	oIframe.document.write("</body>");
	oIframe.document.write("</html>");
	if (InitValue!="")
	{
		oIframe.document.body.innerHTML=InitValue;
	}
	oIframe.document.close();
	oIframe.document.body.contentEditable = "True";
	oIframe.document.charset=charset;
}

//��������Ϣ�༭���ڱ༭ģʽ
var SingleEditorMode = "EDIT";
function SetFrameMode(obj,Style)
{
	obj.className='ModeBarBtnOn';
	if(Style=="HTML"){
		document.getElementById("SingleEDIT").className = 'ModeBarBtnOff';
		if(SingleEditorMode == "EDIT") {oSingleEditor.document.body.innerText = oSingleEditor.document.body.innerHTML;SingleEditorMode='HTML'}
	}
	if(Style=="EDIT"){
		document.getElementById("SingleHTML").className = 'ModeBarBtnOff';
		if(SingleEditorMode == "HTML") {oSingleEditor.document.body.innerHTML = oSingleEditor.document.body.innerText;SingleEditorMode='EDIT'}
	}
	SingleBtnStatus = obj.id;
}

//�ı�༭����������

function ChangeFolder(el)
{
	if (el.className=='LableSelected') return;
	var OperObj=null;
	var FolderIDArray=new Array('SingleID','ContentID');
	var EditAreaIDArray=new Array('SingleArea','ContentArea');
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

//���ذ����ؼ���
function LoadHelpFileName(obj)
{
	var oFrame;
	if(obj.FileName.value==''){
		alert('���������ҳ���ļ���ַ');
		return false;
	}
	if (document.all){
		oFrame=frames['FrameArea'];
	}else{
		oFrame=document.getElementById("FrameArea").contentWindow;
	}
	//window.open('SelectFileName.asp?FileName='+obj.FileName.value+'&PageField='+obj.PageField.value);
	oFrame.location.href='SelectFileName.asp?FileName='+obj.FileName.value+'&PageField='+obj.PageField.value;
}


//��ʼ������˵������
var SingleContentArray=new Array(),ContentArray=new Array();
var oPageFieldArray=new Array();
oPageFieldArray[0]=SingleContentArray[0]=ContentArray[0]='';

//�༭���е�ֵ�л�
function SetZone()
{
	var v = HelpForm.PageField.options[HelpForm.PageField.selectedIndex];
	//���ӹؼ��ֵ�input
	if(HelpForm.NewPageField.value.indexOf(v.value)==-1)
		HelpForm.NewPageField.value = (HelpForm.NewPageField.value + ',' + v.value).replace(/^,/,'');
	for(var i=0;i<oPageFieldArray.length;i++)
	{
		if(v.value == oPageFieldArray[i])break;
	}
	if(SingleEditorMode=="HTML") oSingleEditor.document.body.innerText = SingleContentArray[i];
	else  oSingleEditor.document.body.innerHTML = SingleContentArray[i];
	oHelpEditor.EditArea.document.body.innerHTML = ContentArray[i]
}

//���ؼ���������
function LoadSearchBox()
{
	var retValue = OpenWindow('SearchBox.asp',280,120,window)
	if(retValue) window.location = retValue;
}
-->
</script>
<%
Set HelpConn = Nothing
Set Conn = Nothing
%>