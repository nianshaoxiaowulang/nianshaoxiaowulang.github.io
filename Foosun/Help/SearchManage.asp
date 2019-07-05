<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<%
'==============================================================================
'软件名称：FoosunHelp System Form FoosunCMS
'当前版本：Foosun Content Manager System 3.0 系列
'最新更新：2004.12
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
Dim DBC,Conn,HelpConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = "DBQ=" + Server.MapPath("Foosun_help.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set HelpConn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070800") then Call ReturnError1()
'权限判断
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>帮助文件列表</title>
</head>
<script src="../SysJS/PublicJS.js" language="JavaScript"></script>
<link href="../../CSS/FS_css.css" rel="stylesheet">
<body onselectstart="return false;" onClick="ClickFileName();" topmargin="2" leftmargin="2">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
			<td width="35" align="center" alt="新建" onClick="AddField();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新建</td>
			<td width=2 class="Gray">|</td>
			<td width="35" align="center" alt="修改" onClick="ModiField()" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">修改</td>
			<td width=2 class="Gray">|</td>
			<td width="35" align="center" alt="删除" onClick="DeleteField()" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
			<td width=2 class="Gray">|</td>
			<td width="35" align="center" alt="查看" onClick="ReadField()" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">查看</td>
			<td style="display:none;" width=2 class="Gray">|</td>
			<td style="display:none;" width="35" align="center" alt="复制" onClick="CopyField()" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">复制</td>
			<td width=2 class="Gray">|</td>
            <td width="35" align="center" alt="检索" onClick="LoadSearchBox();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">检索</td>
			<td width=2 class="Gray">|</td>
		  	<td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" height="95%" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr> 

    <td valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr bgcolor="#E7E7E7"> 
              <td width="23%" height="26" class="ButtonListleft" align="center">文件名 </td>
              <td width="18%" height="26" class="ButtonList" align="center">文件功能</td>
              <td width="*" height="26" class="ButtonList" align="center">关键字</td>
              <!--td width="40%" height="26" class="ButtonList" align="center">简单介绍</td-->
              <td width="10%" height="26" class="ButtonList" align="center">更新时间</td>
            </tr>
        <%
		Dim FuncName,FileName,PageField
		Dim strSQL,sqlCondition
		FuncName = Request.QueryString("FuncName")
		FileName = Request.QueryString("FileName")
		PageField = Request.QueryString("PageField")
		if FuncName<>"" Then sqlCondition = sqlCondition & " and FuncName like '%"& FuncName &"%'"
		if FileName<>"" Then sqlCondition = sqlCondition & " and FileName='"& FileName &"'"
		if PageField<>"" Then sqlCondition = sqlCondition & " and PageField like '%"& PageField &"%'"

		strSQL = "Select * From [FS_Help] where 1=1 "&sqlCondition&" order by ID DESC"
		Dim Page_No,Page_Total,Record_All,Page_Size
		Page_No = Cint(Request.QueryString("Page_No"))
		Page_size = 20

		Dim RsHelpObj
		set RsHelpObj = Server.CreateObject("Adodb.Recordset")
		RsHelpObj.Open strSQL,Helpconn,1,1

		Record_All = RsHelpObj.RecordCount
		RsHelpObj.PageSize = Page_Size
		Page_Total = Record_All\Page_size
		If Record_All mod Page_Size <> 0 Then Page_Total = Record_All\Page_size + 1
		If Page_No<=0 Then Page_No = 1
		If Page_No>Page_Total Then Page_No = Page_Total
		Dim isCheck,AllItemID

		IF not RsHelpObj.eof THEN
			RsHelpObj.AbsolutePage = Page_No
			While not RsHelpObj.eof And Page_Size>0
%>
        <tr> 
          <td align="left"> 
              <table border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td><img src="../Images/Info/WordNews.gif" width="20" height="20"></td>
                  <td><span class="TempletItem" HelpID="<% = RsHelpObj("ID") %>"><% = RsHelpObj("FileName") %></span></td>
                </tr>
              </table>
          </td>

          <td align="center"><% = RsHelpObj("FuncName") %></td>
          <td>　<% = RsHelpObj("PageField") %></td>
          <!--td>　<% = RsHelpObj("HelpSingleContent")%></td-->
		  <td align="center"><%=FormatDateTime(RsHelpObj("SvTime"),2)%></td>
        </tr>
		<%
				Page_Size = Page_Size - 1
				RsHelpObj.MoveNext
			Wend
		END IF
		RsHelpObj.Close
		Set RsHelpObj = Nothing
		%>
      </table></td>
  </tr>
   <tr> 
    <td height="20" class="ButtonListLeft">
		<table width="100%" height="100%" border="0" cellpadding="3" cellspacing="0">
			<tr> 
			  <td align="right"><% = PageStr %></td>
			</tr>
		</table></td>
  </tr>
</table>
</td>
</tr>
</table>
<iframe id="hideFrame" src="" width="0" height="0" style="display:none;"></iframe>
</body>
</html>
<%
Set Conn=Nothing
Function PageStr()
	PageStr = "位置:<b>"& Page_No &"</b>/<b>"& Page_Total &"</b>,&nbsp;&nbsp;&nbsp;"
	if Page_Total = 1 then
		PageStr = PageStr & "&nbsp;<img src=""../images/FirstPage.gif"" border=0 alt=首页>&nbsp;" & Chr(13) & Chr(10)
		PageStr = PageStr & "&nbsp;<img src=""../images/prePage.gif"" border=0 alt=上一页>&nbsp;" & Chr(13) & Chr(10)
		PageStr = PageStr & "&nbsp;<img src=""../images/nextPage.gif"" border=0 alt=下一页>&nbsp;" & Chr(13) & Chr(10)
		PageStr = PageStr & "&nbsp;<img src=""../images/endPage.gif"" border=0 alt=尾页>&nbsp;" & Chr(13) & Chr(10)
	else
		if cint(Page_No) <> 1 and cint(Page_No) <> Page_Total then
			PageStr = PageStr & "&nbsp;<span onclick=""ChangePageNO('1','Page_No');"" style=""cursor:hand;""><img src=""../images/FirstPage.gif"" border=0 alt=首页></span>&nbsp;" & Chr(13) & Chr(10)
			PageStr = PageStr & "&nbsp;<span onclick=""ChangePageNO('" & Page_No - 1 & "','Page_No');"" style=""cursor:hand;""><img src=../images/prePage.gif border=0 alt=上一页></span>&nbsp;" & Chr(13) & Chr(10)
			PageStr = PageStr & "&nbsp;<span onclick=""ChangePageNO('" & Page_No + 1 & "','Page_No');"" style=""cursor:hand;""><img src=../images/nextPage.gif border=0 alt=下一页></span>&nbsp;" & Chr(13) & Chr(10)
			PageStr = PageStr & "&nbsp;<span onclick=""ChangePageNO('" & Page_Total & "','Page_No');"" style=""cursor:hand;""><img src=../images/endPage.gif border=0 alt=尾页></span>&nbsp;" & Chr(13) & Chr(10)
		elseif cint(Page_No) = 1 then
			PageStr = PageStr & "&nbsp;<img src=../images/FirstPage.gif border=0 alt=首页>&nbsp;" & Chr(13) & Chr(10)
			PageStr = PageStr & "&nbsp;<img src=../images/prePage.gif border=0 alt=上一页>&nbsp;" & Chr(13) & Chr(10)
			PageStr = PageStr & "&nbsp;<span onclick=""ChangePageNO('" & Page_No + 1 & "','Page_No');"" style=""cursor:hand;""><img src=../images/nextPage.gif border=0 alt=下一页></span>&nbsp;" & Chr(13) & Chr(10)
			PageStr = PageStr & "&nbsp;<span onclick=""ChangePageNO('" & Page_Total & "','Page_No');"" style=""cursor:hand;""><img src=../images/endpage.gif border=0 alt=尾页></span>&nbsp;" & Chr(13) & Chr(10)
		else
			PageStr = PageStr & "&nbsp;<span onclick=""ChangePageNO('1','News_Page_No');"" style=""cursor:hand;""><img src=../images/FirstPage.gif border=0 alt=首页></span>&nbsp;" & Chr(13) & Chr(10)
			PageStr = PageStr & "&nbsp;<span onclick=""ChangePageNO('" & Page_No - 1 & "','Page_No');"" style=""cursor:hand;""><img src=../images/prePage.gif border=0 alt=上一页></span>&nbsp;" & Chr(13) & Chr(10)
			PageStr = PageStr & "&nbsp;<img src=../images/nextPage.gif border=0 alt=下一页></span>&nbsp;" & Chr(13) & Chr(10)
			PageStr = PageStr & "&nbsp;<img src=../images/endpage.gif border=0 alt=尾页>&nbsp;" & Chr(13) & Chr(10)
		end if
	end if
End Function

%>
<script language="JavaScript">
var DocumentReadyTF=false;
var ListObjArray = new Array();
var ContentMenuArray=new Array();
var ObjPopupMenu=window.createPopup();
document.oncontextmenu=new Function("return ShowMouseRightMenu(window.event);");
function ShowMouseRightMenu(event)
{
	ContentMenuShowEvent();
	var width=100;
	var height=0;
	var lefter=event.clientX;
	var topper=event.clientY;
	var ObjPopDocument=ObjPopupMenu.document;
	var ObjPopBody=ObjPopupMenu.document.body;
	var MenuStr='';
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (ContentMenuArray[i].ExeFunction=='seperator')
		{
			MenuStr+=FormatSeperator();
			height+=16;
		}
		else
		{
			MenuStr+=FormatMenuRow(ContentMenuArray[i].ExeFunction,ContentMenuArray[i].Description,ContentMenuArray[i].EnabledStr);
			height+=20;
		}
	}
	MenuStr="<TABLE border=0 cellpadding=0 cellspacing=0 class=Menu width=100>"+MenuStr
	MenuStr=MenuStr+"<\/TABLE>";
	ObjPopDocument.open();
	ObjPopDocument.write("<head><link href=\"../../CSS/ContentMenu.css\" type=\"text/css\" rel=\"stylesheet\"></head><body scroll=\"no\" onConTextMenu=\"event.returnValue=false;\" onselectstart=\"event.returnValue=false;\">"+MenuStr);
	ObjPopDocument.close();
	height+=4;
	if(lefter+width > document.body.clientWidth) lefter=lefter-width;
	ObjPopupMenu.show(lefter, topper, width, height, document.body);
	return false;
}
function FormatSeperator()
{
	var MenuRowStr="<tr><td height=16 valign=middle><hr><\/td><\/tr>";
	return MenuRowStr;
}
function FormatMenuRow(MenuOperation,MenuDescription,EnabledStr)
{
	var MenuRowStr="<tr "+EnabledStr+"><td align=left height=20 class=MouseOut onMouseOver=this.className='MouseOver'; onMouseOut=this.className='MouseOut'; valign=middle"
	if (EnabledStr=='') MenuRowStr+=" onclick=\""+MenuOperation+"parent.ObjPopupMenu.hide();\">&nbsp;&nbsp;&nbsp;&nbsp;";
	else MenuRowStr+=">&nbsp;&nbsp;&nbsp;&nbsp;";
	MenuRowStr=MenuRowStr+MenuDescription+"<\/td><\/tr>";
	return MenuRowStr;
}

function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function InitialContentListContentMenu()
{	
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.AddField();','新建','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.ModiField();','修改','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.DeleteField();','删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.ReadField();','查看','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.PrintThePage();','打印本页','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','刷新','');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'本页面路径属性\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','路径属性','');
	IntialListObjArray();
}
function ContentMenuShowEvent()
{
	ChangeHelpMenuStatus();
}
function ChangeHelpMenuStatus()
{
	var EventObjInArray=false,SelectContent='',DisabledContentMenuStr='';
	for (var i=0;i<ListObjArray.length;i++)
	{
		if (event.srcElement==ListObjArray[i].Obj)
		{
			if (ListObjArray[i].Selected==true) EventObjInArray=true;
			break;
		}
	}
	for (var i=0;i<ListObjArray.length;i++)
	{
		if (event.srcElement==ListObjArray[i].Obj)
		{
			ListObjArray[i].Obj.className='TempletSelectItem';
			ListObjArray[i].Selected=true;
			if (SelectContent=='') SelectContent=ListObjArray[i].Obj.NewsID;
			else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.NewsID;
		}
		else
		{
			if (!EventObjInArray)
			{
				ListObjArray[i].Obj.className='TempletItem';
				ListObjArray[i].Selected=false;
			}
			else
			{
				if (ListObjArray[i].Selected==true)
				{
					if (SelectContent=='') SelectContent=ListObjArray[i].Obj.NewsID;
					else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.NewsID;
				}
			}
		}
	}
	if (SelectContent=='') DisabledContentMenuStr=',删除,修改,查看,';
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}

function FolderFileObj(Obj,Index,Selected)
{
	this.Obj=Obj;
	this.Index=Index;
	this.Selected=Selected;
}
function IntialListObjArray()
{
	var CurrObj=null,j=1;
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if (CurrObj.HelpID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function ClickFileName()
{
	var el=event.srcElement;
	var i=0;
	if ((event.ctrlKey==true)||(event.shiftKey==true))
	{
		if (event.ctrlKey==true)
		{
			for (i=0;i<ListObjArray.length;i++)
			{
				if (el==ListObjArray[i].Obj)
				{
					if (ListObjArray[i].Selected==false)
					{
						ListObjArray[i].Obj.className='TempletSelectItem';
						ListObjArray[i].Selected=true;
					}
					else
					{
						ListObjArray[i].Obj.className='TempletItem';
						ListObjArray[i].Selected=false;
					}
				}
			}
		}
		if (event.shiftKey==true)
		{
			var MaxIndex=0,ObjInArray=false,EndIndex=0,ElIndex=-1;
			for (i=0;i<ListObjArray.length;i++)
			{
				if (ListObjArray[i].Selected==true)
				{
					if (ListObjArray[i].Index>=MaxIndex) MaxIndex=ListObjArray[i].Index;
				}
				if (el==ListObjArray[i].Obj)
				{
					ObjInArray=true;
					ElIndex=i;
					EndIndex=ListObjArray[i].Index;
				}
			}
			if (ElIndex>MaxIndex)
			{
				if (MaxIndex>0)
					for (i=MaxIndex-1;i<EndIndex;i++)
					{
						ListObjArray[i].Obj.className='TempletSelectItem';
						ListObjArray[i].Selected=true;
					}
				else
				{
					ListObjArray[ElIndex].Obj.className='TempletSelectItem';
					ListObjArray[ElIndex].Selected=true;
				}
			}
			else
			{
				if (ObjInArray)
				{
					for (i=EndIndex;i<MaxIndex-1;i++)
					{
						ListObjArray[i].Obj.className='TempletSelectItem';
						ListObjArray[i].Selected=true;
					}
					if (ElIndex>=0)
					{
						ListObjArray[ElIndex].Obj.className='TempletSelectItem';
						ListObjArray[ElIndex].Selected=true;
					}
				}
			}
		}
	}
	else
		for (i=0;i<ListObjArray.length;i++)
		{
			if (el==ListObjArray[i].Obj)
			{
				ListObjArray[i].Obj.className='TempletSelectItem';
				ListObjArray[i].Selected=true;
			}
			else
			{
				ListObjArray[i].Obj.className='TempletItem';
				ListObjArray[i].Selected=false;
			}
		}
	
}


function AddField(){location='AddField.asp';}
function ModiField(){
	var SelectedFile='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.HelpID!=null)
			{
				if (SelectedFile=='') SelectedFile=ListObjArray[i].Obj.HelpID;
				else  SelectedFile=SelectedFile+','+ListObjArray[i].Obj.HelpID;
			}
		}
	}
	if(SelectedFile.indexOf(",")!=-1 || SelectedFile ==''){alert('请选择一条记录');return false;}
	location = 'AddField.asp?ID=' + SelectedFile;
}

function DeleteField()
{
	var SelectedFile='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.HelpID!=null)
			{
				if (SelectedFile=='') SelectedFile=ListObjArray[i].Obj.HelpID;
				else  SelectedFile=SelectedFile+','+ListObjArray[i].Obj.HelpID;
			}
		}
	}
	if(SelectedFile ==''){alert('请选择一条以上的记录');return false;}
	if(confirm('确认删除选中的数据吗？')) hideFrame.location = 'DeleteField.asp?ID=' + SelectedFile;
}

function ReadField()
{
	var SelectedFile='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.HelpID!=null)
			{
				if (SelectedFile=='') SelectedFile=ListObjArray[i].Obj.HelpID;
				else  SelectedFile=SelectedFile+','+ListObjArray[i].Obj.HelpID;
			}
		}
	}
	if(SelectedFile.indexOf(",")!=-1 || SelectedFile ==''){alert('请选择一条记录');return false;}
	window.open('ReadMore.asp?ID=' + SelectedFile,'HelpWindow','width=720,height=380,top='+(screen.height-380)/2+',left='+(screen.width-720)/2+',resizable=yes,status=1,scrollbars=1');
}

//临时用的copy数据
function CopyField()
{
	var SelectedFile='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.HelpID!=null)
			{
				if (SelectedFile=='') SelectedFile=ListObjArray[i].Obj.HelpID;
				else  SelectedFile=SelectedFile+','+ListObjArray[i].Obj.HelpID;
			}
		}
	}
	if(SelectedFile ==''){alert('请选择一条以上的记录');return false;}
	if(confirm('本功能是临时使用，确认copy数据吗？')) hideFrame.location.href = 'CopyField.asp?ID=' + SelectedFile;
}

function LoadSearchBox()
{
	var retValue = OpenWindow('SearchBox.asp',280,120,window)
	if(retValue) window.location = retValue;
}


function ChangePageNO(NO,SearchStr)
{
	var LocationStr=window.location.href;
	var SearchLocation=LocationStr.lastIndexOf(SearchStr);
	if (SearchLocation!=-1)
	{
		var TempSearchLocation=LocationStr.indexOf('&',SearchLocation);
		if (TempSearchLocation!=-1)
		{
			LocationStr=LocationStr.slice(0,SearchLocation)+SearchStr+'='+NO+window.location.href.slice(TempSearchLocation);
		}
		else LocationStr=LocationStr.slice(0,SearchLocation)+SearchStr+'='+NO;
	}
	else
	{
		if (LocationStr.lastIndexOf('?')!=-1) LocationStr=LocationStr+'&'+SearchStr+'='+NO;
		else  LocationStr=LocationStr+'?'+SearchStr+'='+NO;
	}
	window.location=LocationStr;
}
</script>
<%Set HelpConn=nothing%>