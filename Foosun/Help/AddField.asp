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
'软件名称：FoosunHelp System Form FoosunCMS
'当前版本：Foosun Content Manager System 3.0 系列
'最新更新：2005.12
'==============================================================================
'商业注册联系：028-85098980-601,602 技术支持：028-85098980-605、607,客户支持：608
'产品咨询QQ：159410,394226379,125114015,655071
'技术支持:所有程序使用问题，请提问到bbs.foosun.net我们将及时回答您
'程序开发：风讯开发组 & 风讯插件开发组
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.net  演示站点：test.cooin.com    
'网站建设专区：www.cooin.com
'==============================================================================
'免费版本请在新闻首页保留版权信息，并做上本站LOGO友情连接
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
		Response.write "<script>alert('没有找到相关修改数据');</script>"
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
            <td width="35" align="center" alt="保存" onClick="if(BindSubmit())HelpForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
			<td width=2 class="Gray">|</td>
			<td width="35" align="center" alt="搜索" onClick="LoadSearchBox();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">搜索</td>
			<td width=2 class="Gray">|</td>
		  	<td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
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
      <td width="12%" align="center" bgcolor="#EFEFEF">页面功能</td>
      <td width="87%" bgcolor="#EFEFEF" style="padding:0px 5px; "> <input type="text" size=72 name="FuncName" value="<%=FuncName%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="center" bgcolor="#EFEFEF">页面地址</td>
      <td bgcolor="#EFEFEF" style="padding:0px 5px; "> <input type="text" size=72 name="FileName" value="<%=FileName%>">
        　 <input type="button" name="selectFileName" value="读取帮助" title="读取给定页面地址中存在的帮助内容" onClick="LoadHelpFileName(this.form);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="12%" align="center" bgcolor="#EFEFEF">选择关键字</td>
      <td width="87%" bgcolor="#EFEFEF" style="padding:0px 5px; "> <input type="text" size="72" Name="NewPageField" value="<%=PageField%>">
        <br> <select Name="PageField" onChange="SetZone();" style="width:66%">
          <option value="">新建关键字</option>
        </select> <font color=red>多个帮助内容相同的关键字,请用“,”分割</font> <input name="HelpSingleContent" type="hidden" value="<%=server.HtmlEncode(HelpSingleContent)%>"> 
        <input name="HelpContent" type="hidden" value="<%=server.HtmlEncode(HelpContent)%>"></td>
    </tr>
    <tr> 
      <td colspan="2"><table width="100%" border="0" cellpadding="0" cellspacing="0" height="20" bgcolor="#FFFFFF">
          <tr> 
            <td width="60" height="26" align="center" bgcolor="#EFEFEF" class="LableSelected" id="SingleID" onClick="ChangeFolder(this);">简单介绍</td>
            <td width="5" height="26" align="center" class="ToolBarButtonLine" style="cursor:default;">&nbsp;</td>
            <td width="60" height="26" align="center" class="LableDefault" id="ContentID" onClick="ChangeFolder(this);">详细说明</td>
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
//初始化帮助表单
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

	
//绑定提交
function BindSubmit()
{
	if (oHelpEditor.CurrMode!='EDIT') {alert('其他模式下无法保存，请切换到编辑模式');return false;}
	if(SingleEditorMode == "HTML") document.HelpForm.HelpSingleContent.value = oSingleEditor.document.body.innerText;
	else document.HelpForm.HelpSingleContent.value = oSingleEditor.document.body.innerHTML;
	document.HelpForm.HelpContent.value = oHelpEditor.EditArea.document.body.innerHTML;
	return true;
}
//初始化简单帮助说明数据
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

//交换简单信息编辑窗口编辑模式
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

//改变编辑器内容区域

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

//加载帮助关键字
function LoadHelpFileName(obj)
{
	var oFrame;
	if(obj.FileName.value==''){
		alert('请输入帮助页面文件地址');
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


//初始化帮助说明变量
var SingleContentArray=new Array(),ContentArray=new Array();
var oPageFieldArray=new Array();
oPageFieldArray[0]=SingleContentArray[0]=ContentArray[0]='';

//编辑器中的值切换
function SetZone()
{
	var v = HelpForm.PageField.options[HelpForm.PageField.selectedIndex];
	//增加关键字到input
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

//加载检索条件表单
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