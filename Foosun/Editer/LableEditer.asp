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
<table width="100%" height="90" border="0" cellpadding="0" cellspacing="0" id="Toolbar">
  <tr> 
    <td><table height="30" border="0" cellpadding="0" cellspacing="0" class="ToolSet">
        <tr> 
          <td width="30"><div align="center"><img onClick="InsertScript('Class');"  class="Btn" alt="��Ŀ�����б�" src="../Images/Lable/Class.gif" width="24" height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('ChildClass');"  class="Btn" alt="����Ŀ�����б�" src="../Images/Lable/ChildClass.gif" width="24" height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('NewsList');"  class="Btn" alt="�ռ��б�" src="../Images/Lable/ZNews.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('Special');"  class="Btn" alt="ר�������б�" src="../Images/Lable/Special.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('SpecialNewsList');"  class="Btn" alt="ר���ռ������б�" src="../Images/Lable/endspeacl.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('SpecialNewsindex');"  class="Btn" alt="ר�⵼��" src="../Images/Lable/spnavi.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('News');"  class="Btn" alt="���ű�ǩ" src="../Images/Lable/News.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('PageTitle');"  class="Btn" alt="ҳ�����" src="../Images/Lable/PageTItle1.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('CopyRight');"  class="Btn" alt="��Ȩ��Ϣ" src="../Images/Lable/Copyright.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('Location');"  class="Btn" alt="��ǰλ�õ���" src="../Images/Lable/Location.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('NaviNavi');"  class="Btn" alt="��վ����" src="../Images/Lable/Navi.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('ClassNavi');"  class="Btn" alt="��Ŀ����" src="../Images/Lable/Navi1.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('UseLogin');"  class="Btn" alt="�û���½" src="../Images/Lable/Login.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('InfoStat');"  class="Btn" alt="��Ϣͳ��" src="../Images/Lable/InfoSta.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('Search');"  class="Btn" alt="����" src="../Images/Lable/Search.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('AdvancedSearch');"  class="Btn" alt="�߼�����" src="../Images/Lable/AdvanceSearch.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('FriendLink');"  class="Btn" alt="��������" src="../Images/Lable/FriendLink.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('TodayNews');"  class="Btn" alt="����ͷ��" src="../Images/Lable/todaynews.gif" width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('NaviReadNews');"  class="Btn" alt="��������" src="../Images/Lable/DNews.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('LastNews');"  class="Btn" alt="��������" src="../Images/Lable/LastNews.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('HotNews');"  class="Btn" alt="�ȵ�����" src="../Images/Lable/HotNews.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('RecNews');"  class="Btn" alt="�Ƽ�����" src="../Images/Lable/RecNews.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('Marquee');"  class="Btn" alt="��������" src="../Images/Lable/Marquee.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('RelateNews');"  class="Btn" alt="�������" src="../Images/Lable/RelateNews.gif"  width="24"  height="24"></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table height="30" border="0" cellpadding="0" cellspacing="0" class="ToolSet">
        <tr> 
          <td width="30"><div align="center"><img onClick="InsertScript('SiteMap');"  class="Btn" alt="վ���ͼ" src="../Images/Lable/map.gif" width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('PicNews');"  class="Btn" alt="ͼƬ����" src="../Images/Lable/PicNews.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('Filter');"  class="Btn" alt="�õ�Ƭ����" src="../Images/Lable/FNews.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('FocusNews');"  class="Btn" alt="����ͼƬ" src="../Images/Lable/Pic_1.gif"  width="24"  height="24"></div></td>
          <td width="1"><div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('RecPic');"  class="Btn" alt="�Ƽ�ͼƬ" src="../Images/Lable/Pic_2.gif" width="24" height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('ClassicalNews');"  class="Btn" alt="���ʻع�" src="../Images/Lable/Pic_3.gif" width="24" height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('ClassicalPic');"  class="Btn" alt="����ͼƬ" src="../Images/Lable/Pic_4.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('LastClassPic');"  class="Btn" alt="�ռ�ͼƬ" src="../Images/Lable/Pic_54.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('ClassDownLoad');"  class="Btn" alt="������Ŀ" src="../Images/Lable/Down_1.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('DonwLoadList');"  class="Btn" alt="�ռ�����" src="../Images/Lable/Down_2.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('LastDownList');"  class="Btn" alt="��������" src="../Images/Lable/Down_3.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('RecDownList');"  class="Btn" alt="�Ƽ�����" src="../Images/Lable/Down_4.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('HotDownList');"  class="Btn" alt="�ȵ�����" src="../Images/Lable/Down_5.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('DownList');"  class="Btn" alt="��������" src="../Images/Lable/Down_6.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('DownInfoStat');"  class="Btn" alt="����ͳ��" src="../Images/Lable/Down_7.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('RelateSpecial');"  class="Btn" alt="���ר��" src="../Images/Lable/RelateSpecial.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
		  <td width="30"><div align="center"><img onClick="InsertScript('FileLable');"  class="Btn" alt="�鵵��ǩ" src="../Images/Lable/FileLable.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('PrePageLable');"  class="Btn" alt="��ƪ����" src="../Images/Lable/PrePage.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('NextPageLable');"  class="Btn" alt="��ƪ����" src="../Images/Lable/NextPage.gif"  width="24"  height="24"></div></td>
          <td width="1"> <div align="center" class="ToolSeparator"></div></td>
          <td width="30"><div align="center"><img onClick="InsertScript('FreeLable');"  class="Btn" alt="���ɱ�ǩ" src="../Images/Lable/FreeLable.gif"  width="24"  height="24"></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table height="30" border="0" cellpadding="0" cellspacing="0" class="ToolSet">
        <tr> 
          <td width="30" align="center"><img title="ɾ������HTML��ʶ" onClick="DelAllHtmlTag()" class="Btn" src="../Images/Editer/geshi.gif" ></td>
          <td width="30" align="center"><img title="ɾ�����ָ�ʽ" onClick="Format('removeformat')" class="Btn" src="../Images/Editer/clear.gif" ></td>
          <td width="30" align="center"><img src="../Images/Editer/TextColor.gif" width="23" height="22" class="Btn" title="������ɫ" onClick="TextColor()" ></td>
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
  <tr> 
    <td><iframe src="<% = ExtendEditFile %>" name="EditArea" ID="EditArea" MARGINHEIGHT="1" MARGINWIDTH="1" width="100%" scrolling="yes"></iframe></td>
  </tr>
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
	if (document.readyState!="complete") return;
	if (DocumentReadyTF) return;
	var i,j,s,curr;
	for (i=0; i<document.body.all.length;i++)
	{
		curr=document.body.all[i];
		if (curr.className=="Btn") InitBtn(curr);
	}
	SetEditAreaHeight();
	SetBodyStyle();
	DocumentReadyTF=true;
}
function SetEditAreaHeight()
{
	var BodyHeight=document.body.clientHeight;
	var EditAreaHeight=BodyHeight-140;
	document.all.EditArea.height=EditAreaHeight;
}
function SetBodyStyle()
{
	//EditArea.document.body.runtimeStyle.fontSize='10pt';
}
function InsertScript(Flag)
{
	var ReturnValue='';
	switch (Flag)
	{
		case 'FreeLable':
			ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableFreeLable.asp&PageTitle=�������ɱ�ǩ����',500,300,window);
			break;	
		case 'Class':
			ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableClass.asp&PageTitle=ѡ����Ŀ��ǩ����',336,260,window);
			break;
		case 'ChildClass':
			ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableChildClass.asp&PageTitle=ѡ������Ŀ��ǩ����',336,310,window);
			break;
		case 'Special':
			ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableSpecial.asp&PageTitle=ѡ��ר���ǩ����',336,260,window);
			break;
		case 'News':
			ReturnValue=OpenWindow('../FunPages/LableNews.asp',200,90,window);
			break;
		case 'SpecialNewsList':
			ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableSpecialNewslist.asp&PageTitle=ѡ��ר���ռ������б��ǩ����',350,240,window);
			break;
		case 'SpecialNewsindex':
			ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableSpecialNewsindex.asp&PageTitle=ѡ��ר�⵼�����ǩ����',336,160,window);
			break;
		case 'Location':
			var ReturnValue=OpenWindow('../FunPages/LableLocation.htm',336,105,window);
			break;
		case 'NaviNavi':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableNavi.asp&PageTitle=ѡ����վ������ǩ����',336,170,window);
			break;
		case 'ClassNavi':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableClassNavi.asp&PageTitle=ѡ����Ŀ������ǩ����',336,156,window);
			break;
		case 'UseLogin':
			ReturnValue='{%=UseLogin()%}';
			break;
		case 'AdvancedSearch':
			ReturnValue='{%=AdvancedSearch()%}';
			break;
		case 'Search':
			ReturnValue='{%=Search()%}';
			break;
		case 'NaviReadNews':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableNaviReadNews.asp&PageTitle=ѡ�񵼶����ű�ǩ����',336,240,window);
			break;
		case 'HotNews':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableHotNews.asp&PageTitle=ѡ���ȵ����ű�ǩ����',336,230,window);
			break;
		case 'RecNews':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableRecNews.asp&PageTitle=ѡ���Ƽ����ű�ǩ����',336,220,window);
			break;
		case 'LastNews':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableLastNews.asp&PageTitle=ѡ���������ű�ǩ����',336,220,window);
			break;
		case 'Marquee':
			var ReturnValue=OpenWindow('../FunPages/LableMarquee.htm',336,170,window);
			break;
		case 'RelateNews':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableRelateNews.asp&PageTitle=ѡ��������ű�ǩ����',336,176,window);
			break;
		case 'PicNews':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LablePicNews.asp&PageTitle=ͼƬ��������',336,170,window);
			break;
		case 'TodayNews':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableTodayNews.asp&PageTitle=����ͷ��',360,215,window);
			break;
		case 'Filter':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableFilter.asp&PageTitle=�õ�Ƭ��������',336,150,window);
			break;
		case 'FocusNews':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableFocusPic.asp&PageTitle=����ͼƬ',336,170,window);
			break;
		case 'RecPic':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableRecPic.asp&PageTitle=�Ƽ�ͼƬ',336,170,window);
			break;
		case 'ClassicalNews':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableClassicalNews.asp&PageTitle=���ʻع�',336,170,window);
			break;
		case 'ClassicalPic':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableClassicalPic.asp&PageTitle=����ͼƬ',336,170,window);
			break;
		case 'LastClassPic':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableLastClassPic.asp&PageTitle=�ռ�ͼƬ',336,170,window);
			break;
		case 'NewsList':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableClassNewsList.asp&PageTitle=��Ŀ�����б�����',346,240,window);
			break;
		case 'InfoStat':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableInfoStat.asp&PageTitle=��Ϣͳ�Ʊ�ǩ����',336,110,window);
			break;
		case 'SiteMap':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableSiteMap.asp&PageTitle=վ��ͳ�Ʊ�ǩ����',336,110,window);
			break;
		case 'FriendLink':
			var ReturnValue=OpenWindow('../FunPages/LableFriendLink.htm',336,145,window);
			break;
		case 'PageTitle':
			var ReturnValue=OpenWindow('../FunPages/LablePageTitle.htm',336,90,window);
			break;
		case 'CopyRight':
			var ReturnValue='{%=CopyRightStr()%}';
			break;
		case 'ClassDownLoad':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableClassDownLoad.asp&PageTitle=��Ŀ���ر�ǩ����',336,220,window);
			break;
		case 'DonwLoadList':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableDonwLoadList.asp&PageTitle=�ռ���Ŀ���ر�ǩ����',360,200,window);
			break;
		case 'LastDownList':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableLastDownList.asp&PageTitle=�������ر�ǩ����',336,200,window);
			break;
		case 'RecDownList':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableRecDownList.asp&PageTitle=�Ƽ����ر�ǩ����',336,200,window);
			break;
		case 'HotDownList':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableHotDownList.asp&PageTitle=�ȵ����ر�ǩ����',336,200,window);
			break;
		case 'DownList':
			var ReturnValue=OpenWindow('../FunPages/LableDownList.htm',200,120,window);
			break;
		case 'DownInfoStat':
			var ReturnValue=OpenWindow('../FunPages/LableDownInfoStat.asp',336,110,window);
			break;
		case 'RelateSpecial':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableRelateSpecial.asp&PageTitle=���ר���ǩ����',336,176,window);
			break;
		case 'FileLable':
			var ReturnValue=OpenWindow('../FunPages/Frame.asp?FileName=LableFile.asp&PageTitle=�鵵��ǩ����',336,220,window);
			break;
		case 'PrePageLable':
			var ReturnValue='{%=PrePageNews()%}';
			break;
		case 'NextPageLable':
			var ReturnValue='{%=NextPageNews()%}';
			break;
		default :
			return '';
	}
	if (ReturnValue!='') parent.frames["Editer"].InsertHTMLStr(ReturnValue);
}
</script>