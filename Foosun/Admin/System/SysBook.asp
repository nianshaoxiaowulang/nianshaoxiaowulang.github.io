<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
'==============================================================================
'软件名称：FoosunShop System Form FoosunCMS
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
Dim DBC,Conn,sRootDir
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070700") then Call ReturnError1()
Dim RsAdminConfigObj
Set RsAdminConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop,QPoint from FS_Config")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>FoosunCMS Shop 1.0.0930</title>
</head>
<style type="text/css">
<!--
.SearchBtnStyle {
	border: 1px solid #000000;
}
.menu {
	position:absolute;
	background: menu;
	border-top: 1px solid buttonhighlight;
	border-left: 1px solid buttonhighlight;
	border-bottom: 2px solid buttonshadow;
	border-right: 2px solid buttonshadow;
	padding: 2px;
	font: menu;
	cursor:default;
	font-size:9pt;
	width:90pt;
	visibility: hidden;
	z-index: 2;
	overflow: visible;
}
.menushow {
	position:absolute;
	visibility:visible;
	background:#EFEFEF;
	border-top: 1px solid #000000;
	border-left: 1px solid #000000;
	border-bottom: 1px solid #000000;
	border-right: 1px solid #000000;
	padding:0px;
	font: 9pt "menu";
	cursor:default;
	width:50pt;
}
-->
</style>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<script language="JavaScript">
var ContentMenuArray=new Array();
var ListObjArray=new Array();
var DocumentReadyTF=false;
var ClassID='<% = ClassID %>';
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.ReadBook();','查看','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditContent();",'修改','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelContent();",'删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.RefreshList();','刷新','');
	IntialListObjArray();
}

function RefreshList()
{
	location.href=location.href;
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function ContentMenuShowEvent()
{
	ChangeContentMenuStatus();
}
function ChangeContentMenuStatus()
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
			if (SelectContent=='') SelectContent=ListObjArray[i].Obj.ContentID;
			else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.ContentID;
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
					if (SelectContent=='') SelectContent=ListObjArray[i].Obj.ContentID;
					else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.ContentID;
				}
			}
		}
	}
	if (SelectContent=='') DisabledContentMenuStr=',修改,删除,查看';
	else
	{
		if (SelectContent.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',修改,'
	}
//	if ((top.MainInfo.SourceNews=='')&&(top.MainInfo.SourceProduct=='')) DisabledContentMenuStr=DisabledContentMenuStr+',粘贴,';
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
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
function AddLocationStr(LocationStr,Value,SearchStr)
{
	var SearchLocation=LocationStr.lastIndexOf(SearchStr);
	if (SearchLocation!=-1)
	{
		var TempSearchLocation=LocationStr.indexOf('&',SearchLocation);
		if (TempSearchLocation!=-1)
		{
			var TempLocationStr=LocationStr.slice(TempSearchLocation)
			LocationStr=LocationStr.slice(0,SearchLocation)+SearchStr+'='+Value+TempLocationStr;
		}
		else LocationStr=LocationStr.slice(0,SearchLocation)+SearchStr+'='+Value;
	}
	else
	{
		if (LocationStr.lastIndexOf('?')!=-1) LocationStr=LocationStr+'&'+SearchStr+'='+Value;
		else  LocationStr=LocationStr+'?'+SearchStr+'='+Value;
	}
	return LocationStr;
}
function auditcontent()
{
	var LocationStr=window.location.href;
	LocationStr = LocationStr.replace(/&Audit=IsAuditTF/g,"").replace(/&Audit=NoAuditTF/g,"");
	LocationStr=LocationStr+"&Audit=IsAuditTF"
	window.location=LocationStr;
}
function noauditcontent()
{
	var LocationStr=window.location.href;
	LocationStr = LocationStr.replace(/&Audit=IsAuditTF/g,"").replace(/&Audit=NoAuditTF/g,"");
	LocationStr=LocationStr+"&Audit=NoAuditTF"
	window.location=LocationStr;
}
function ClickBook()
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
	{
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
}
function IntialListObjArray()
{
	var CurrObj=null,j=1;
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if (CurrObj.ContentID!=null)
		{
			ListObjArray[ListObjArray.length]=new NewsOrProductObj(CurrObj,j,false);
			j++;
		}
	}
}
function NewsOrProductObj(Obj,Index,Selected)
{
	this.Obj=Obj;
	this.Index=Index;
	this.Selected=Selected;
}
function EditContent()
{
	var SelectedContent='',SelectContentObj=null;
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.ContentID!=null)
			{
				if (SelectedContent=='') SelectedContent=ListObjArray[i].Obj.ContentID;
				else  SelectedContent=SelectedContent+'***'+ListObjArray[i].Obj.ContentID;
			}
			SelectContentObj=ListObjArray[i].Obj;
		}
	}
	if (SelectedContent!='')
	{
		if (SelectedContent.indexOf('***')==-1)
		{
			if (SelectContentObj.ContentTypeStr!=null)
			{
				location='SysBookModify.asp?ID='+SelectedContent;
			}
		}
		else alert('一次只能够修改一条帖子');
	}
	else alert('请选择要修改的帖子');
}
function ReadBook()
{
	var SelectedContent='',SelectContentObj=null;
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.ContentID!=null)
			{
				if (SelectedContent=='') SelectedContent=ListObjArray[i].Obj.ContentID;
				else  SelectedContent=SelectedContent+'***'+ListObjArray[i].Obj.ContentID;
			}
			SelectContentObj=ListObjArray[i].Obj;
		}
	}
	if (SelectedContent!='')
	{
		if (SelectedContent.indexOf('***')==-1)
		{
			if (SelectContentObj.ContentTypeStr!=null)
			{
				location='SysBookRead.asp?ID='+SelectedContent;
			}
		}
		else alert('一次只能够查看一条帖子');
	}
	else alert('请选择要查看的帖子');
}
function DelContent()
{
	var SelectedContent='',SelectContentObj=null;
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.ContentID!=null)
			{
				if (SelectedContent=='') SelectedContent=ListObjArray[i].Obj.ContentID;
				else  SelectedContent=SelectedContent+','+ListObjArray[i].Obj.ContentID;
			}
			SelectContentObj=ListObjArray[i].Obj;
		}
	}
	if (SelectedContent!='')
	{
			if (SelectContentObj.ContentTypeStr!=null)
			{
				location='SysBookDelSave.asp?Action=Del&GID='+SelectedContent;
			}
	}
	else alert('请选择要删除的帖子');
}
function ShowAddMenu()
{
	var MenuObj=document.all.AddContentMenu;
	var el=event.srcElement;
	MenuObj.style.display='';
	MenuObj.style.posLeft=el.offsetLeft;
	MenuObj.style.posTop=el.offsetHeight;
	MenuObj.className="menushow";
	MenuObj.setCapture();
}
function MouseOverRightMenu() 
{   
	var el=event.srcElement;
	if (el.tagName!='TD') return;
	if (el.ExeFunction==null) return;
	if (el.style.backgroundColor=="highlight") {el.style.backgroundColor='';el.style.color='black';}
	else {el.style.backgroundColor="highlight";el.style.color='white';}
}
function ClickMenu(MenuObj)
{
	var CurrObj=null;
	var IMGObj=document.body.getElementsByTagName('IMG');
	for (var i=0;i<IMGObj.length;i++)
	{
		CurrObj=IMGObj(i);
		if (CurrObj.className=='BtnMouseOver') CurrObj.className='';
	}
	var el=event.srcElement;
	MenuObj.releaseCapture();
	MenuObj.className="menu";
	for (var i=0;i<MenuObj.children.length;i++)
	{
		var CurrObj=MenuObj.children(i);
		for (var j=0;j<CurrObj.children.length;j++)
		{
			if (CurrObj.children(j).className=='MenuShow') {CurrObj.children(j).className='Menu';}	
		}
	}
	if (el.ExeFunction!=null) eval(el.ExeFunction);
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
function SearchLyPage()
{
	SearchPage.style.display='';
}
</script>
<body topmargin="2" leftmargin="2" onclick="ClickBook(event);" onselectstart="return false;"> 
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999"> 
	<tr bgcolor="#EEEEEE"> 
		<td height="26" colspan="5" valign="middle">
			<table  height="22" border="0" cellpadding="0" cellspacing="0"> 
				<tr> 
					<td width=55 align="center" alt="所有留言" onClick="top.GetEkMainObject().location='SysBook.asp';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">所有留言</td> 
					<td width=2 class="Gray">|</td> 
					<td width=55 align="center" alt="添加留言" onClick="top.GetEkMainObject().location='SysBookWrite.asp';" onMouseMove="BtnMouseOver(this);//ShowAddMenu();" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">发表留言</td> 
					<td width=2 class="Gray">|</td> 
					<td width=65 align="center" alt="已回复留言" onClick="top.GetEkMainObject().location='SysBook.asp?Action=Q';" onMouseMove="BtnMouseOver(this);//ShowAddMenu();" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">已回复留言</td> 
					<td width=2 class="Gray">|</td> 
					<td width=65 align="center" alt="未回复留言" onClick="top.GetEkMainObject().location='SysBook.asp?Action=UnQ';" onMouseMove="BtnMouseOver(this);//ShowAddMenu();" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">未回复留言</td> 
					<td width=2 class="Gray">|</td>
					<td width=55 align="center" alt="留言搜索" onClick="SearchLyPage();" onMouseMove="BtnMouseOver(this);//ShowAddMenu();" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">留言搜索</td> 
					<td width=2 class="Gray">|</td>
		  			<td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
				</tr> 
			</table></td> 
	</tr> 
</table> 
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" height="90%"> 
	<tr> 
		<td height="2"></td> 
	</tr> 
	<tr> 
		<td valign="top"> 
				<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#CCCCCC"> 
						<tr bgcolor="#E8E8E8"> 
							<td width="36%" class="ButtonListleft"><div align="center">标题</div></td> 
							<td width="15%" class="ButtonList"><div align="center">发言人</div></td> 
							<td width="19%" class="ButtonList"><div align="center">发表时间</div></td> 
							<td width="17%" class="ButtonList"><div align="center">回复时间</div></td> 
						</tr> 
						<%
    dim RsCon,strpage,select_count,select_pagecount
	strpage=request.querystring("page")
	if len(strpage)=0 then
		strpage="1"
	end if
	Dim QS
	IF Request("Action")="Q" then
		QS = " and isQ=1"
	ElseIf Request("Action")="UnQ" then
		QS = " and isQ=0"
	Else
		QS = ""
	End if
	Dim Ks
	If Request("Keyword")<>"" then
		Ks = " and Content like '%"& Replace(Request("Keyword"),"'","") &"%' or Title like '%"& Replace(Request("Keyword"),"'","") &"%'"
	Else
		Ks = ""
	End if
	Set RsCon = Server.CreateObject (G_FS_RS)
	RsCon.Source="select * from FS_GBook where QID=0 "& QS & Ks &" Order by Orders,Qtime desc,Addtime desc"
	RsCon.Open RsCon.Source,Conn,1,1
	If RsCon.eof then
		   RsCon.close
		   set RsCon=nothing
		   Response.Write"<TR><TD colspan=""7"" bgcolor=FFFFFF>没有记录。</TD></TR>"
	Else
		Dim Product_Page_Size,Product_Page_No,Product_Page_Total,Product_Record_All,Product_For_Var
		Product_Page_Size = 20
		Product_Page_No = Request.Querystring("Product_Page_No")
		if Product_Page_No <= 0 or Product_Page_No = "" then Product_Page_No = 1
		RsCon.PageSize = Product_Page_Size
		Product_Page_Total = RsCon.PageCount
		if (Cint(Product_Page_No) > Product_Page_Total) then Product_Page_No = Product_Page_Total
		RsCon.AbsolutePage = Product_Page_No
		Product_Record_All = RsCon.RecordCount
		for Product_For_Var = 1 to RsCon.PageSize
			if RsCon.Eof then Exit For
	%> 
						<tr bgcolor="#FFFFFF"> 
							<td nowrap><%If RsCon("Orders")=1 then%> 
								<img src="../../../<%=UserDir%>/GBook/Images/ztop.gif" alt="固顶帖" width="18" height="15"> 
								<%Else%> 
								<img src="../../../<%=UserDir%>/GBook/Images/hotfolder.gif" alt="一般帖子" width="18" height="12"> 
								<%End if%> 
								<span ContentTypeStr="5" AuditTF="" class="TempletItem" ContentID="<% = RsCon("ID") %>" align="center"> 
								<%if len(RsCon("Title"))>18 then
				     Response.Write left(RsCon("Title"),18)&".."
				  Else
				     Response.Write RsCon("Title")
				  End if
				%> 
								</SPAN></td> 
							<td align="center"> <%
					  If RsCon("UserID")=0 then
							Response.Write("<font color=#990000>管理员</font>")
					  Else
					  	Dim MemberObj
						Set MemberObj = Conn.execute("Select MemName From FS_Members Where id="&Replace(Replace(RsCon("UserID"),"'",""),Chr(39),""))
						If Not MemberObj.eof then
							Response.Write("<a href=../../../"&UserDir&"/ReadUser.Asp?UserName="&MemberObj("MemName")&">"& MemberObj("MemName")&"</a>")
						Else
							Response.Write("用户已被删除")
						End if
					End If
					  %> </td> 
							<td align="center"> <% = RsCon("Addtime")%> </td> 
							<td align="center"> <font color="#FF0000"> 
								<%
						If RsCon("Qtime")=RsCon("Addtime") Or RsCon("Qtime")="" Or RsCon("isQ") =0 then
							Response.Write("")
						Else
							Response.Write RsCon("Qtime")
						End if
						%> 
								</font> </td> 
						</tr> 
						<%
		RsCon.MoveNext
	Next
	%> 
				</table> 
				<%
End if
Function ProductPageStr()
	ProductPageStr = "位置:<b>"& Product_Page_No &"</b>/<b>"& Product_Page_Total &"</b>&nbsp;&nbsp;&nbsp;"
	if Product_Page_Total = 1 then
		ProductPageStr = ProductPageStr & "&nbsp;<img src=""../../images/FirstPage.gif"" border=0 alt=首页>&nbsp;" & Chr(13) & Chr(10)
		ProductPageStr = ProductPageStr & "&nbsp;<img src=""../../images/prePage.gif"" border=0 alt=上一页>&nbsp;" & Chr(13) & Chr(10)
		ProductPageStr = ProductPageStr & "&nbsp;<img src=""../../images/nextPage.gif"" border=0 alt=下一页>&nbsp;" & Chr(13) & Chr(10)
		ProductPageStr = ProductPageStr & "&nbsp;<img src=""../../images/endPage.gif"" border=0 alt=尾页>&nbsp;" & Chr(13) & Chr(10)
	else
		if cint(Product_Page_No) <> 1 and cint(Product_Page_No) <> Product_Page_Total then
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('1','Products_Page_No');"" style=""cursor:hand;""><img src=""../../images/FirstPage.gif"" border=0 alt=首页></span>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('" & Product_Page_No - 1 & "','Product_Page_No');"" style=""cursor:hand;""><img src=../../images/prePage.gif border=0 alt=上一页></span>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('" & Product_Page_No + 1 & "','Product_Page_No');"" style=""cursor:hand;""><img src=../../images/nextPage.gif border=0 alt=下一页></span>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('" & Product_Page_Total & "','Product_Page_No');"" style=""cursor:hand;""><img src=../../images/endPage.gif border=0 alt=尾页></span>&nbsp;" & Chr(13) & Chr(10)
		elseif cint(Product_Page_No) = 1 then
			ProductPageStr = ProductPageStr & "&nbsp;<img src=../../images/FirstPage.gif border=0 alt=首页>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<img src=../../images/prePage.gif border=0 alt=上一页>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('" & Product_Page_No + 1 & "','Product_Page_No');"" style=""cursor:hand;""><img src=../../images/nextPage.gif border=0 alt=下一页></span>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('" & Product_Page_Total & "','Product_Page_No');"" style=""cursor:hand;""><img src=../../images/endpage.gif border=0 alt=尾页></span>&nbsp;" & Chr(13) & Chr(10)
		else
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('1','Product_Page_No');"" style=""cursor:hand;""><img src=../../images/FirstPage.gif border=0 alt=首页></span>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('" & Product_Page_No - 1 & "','Product_Page_No');"" style=""cursor:hand;""><img src=../../images/prePage.gif border=0 alt=上一页></span>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<img src=../../images/nextPage.gif border=0 alt=下一页></span>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<img src=../../images/endpage.gif border=0 alt=尾页>&nbsp;" & Chr(13) & Chr(10)
		end if
	end if
End Function
%> 
				<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E7E7E7" id="SearchPage" style="display:none;"> 
					<tr bgcolor="#FFFFFF"> 
						<td width="9%"><a href="SysBook.asp">留言搜索</a><a href="SysBook.asp?Action=UnQ"></a></td> 
						<form name="form1" method="post" action="SysBook.asp"> 
							<td width="91%"> <input name="Keyword" type="text" id="Keyword"> 
								<input type="submit" name="Submit2" value="搜索"> </td> 
						</form> 
					</tr> 
				</table> 
			</td> 
	</tr> 
</table> 
<table width="100%" border="0" cellpadding="5" cellspacing="0"> 
	<tr> 
		<td align="right" class="ButtonListLeft"><%=ProductPageStr%></td> 
	</tr> 
</table> 
</body>
</html>
