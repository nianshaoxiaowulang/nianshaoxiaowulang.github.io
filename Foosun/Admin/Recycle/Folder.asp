<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Inc/Cls_JS.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070100") then Call ReturnError1()
Dim RsClassObj,RsNewsObj,Sql,AlreadyDelClassIDStr,NewsType,PicStr,SunNumAsp
AlreadyDelClassIDStr = ""

Dim RecaSysRootDir
if SysRootDir = "" then
	RecaSysRootDir = ""
else
	RecaSysRootDir = "/" & SysRootDir
end if
'回收站操作
Dim OperateType,MyFile
Set MyFile=Server.CreateObject(G_FS_FSO)
OperateType = Request("OperateType")
if OperateType = "DelAll" then
	if Not JudgePopedomTF(Session("Name"),"P070104") then Call ReturnError1()
	'------------------删除栏目目录物理文件---------------------
	Dim DellNewsClassObjj,RsDellNewsObj,RssDelNewsTempClass
	Set DellNewsClassObjj = Conn.Execute("Select SaveFilePath,ClassEName from FS_NewsClass where DelFlag=1")
	Do while Not DellNewsClassObjj.eof
		If MyFile.FolderExists(Server.Mappath(RecaSysRootDir&DellNewsClassObjj("SaveFilePath")&"/"&DellNewsClassObjj("ClassEName"))) then
			MyFile.DeleteFolder(Server.Mappath(RecaSysRootDir&DellNewsClassObjj("SaveFilePath")&"/"&DellNewsClassObjj("ClassEName")))
		End if
		DellNewsClassObjj.MoveNext
	Loop
	DellNewsClassObjj.Close
	Set DellNewsClassObjj = Nothing
	'----------------删除单个新闻文件---------------------------
	Set RsDellNewsObj = Conn.Execute("Select FileName,FileExtName,ClassID from FS_News where DelTF=1 and ClassID in (Select ClassID from FS_NewsClass where DelFlag=0)")
	Do while not RsDellNewsObj.eof
	Set RssDelNewsTempClass = Conn.Execute("Select ClassEName,SaveFilePath from FS_NewsClass where ClassID='"&RsDellNewsObj("ClassID")&"'")
		If MyFile.FileExists(Server.Mappath(RecaSysRootDir&RssDelNewsTempClass("SaveFilePath")&"/"&RssDelNewsTempClass("ClassEName"))&"/"&RsDellNewsObj("FileName")&"."&RsDellNewsObj("FileExtName")) then
		   MyFile.DeleteFile(Server.Mappath(RecaSysRootDir&RssDelNewsTempClass("SaveFilePath")&"/"&RssDelNewsTempClass("ClassEName"))&"/"&RsDellNewsObj("FileName")&"."&RsDellNewsObj("FileExtName"))
		End if
		RssDelNewsTempClass.Close
		Set RssDelNewsTempClass = Nothing
		RsDellNewsObj.MoveNext
	Loop
	RsDellNewsObj.Close
	Set RsDellNewsObj = Nothing
	'------------------------------------------------------------
	Sql = "Delete from FS_News where DelTF=1"
	Conn.Execute(Sql)
	Sql = "Delete from FS_Contribution where ClassID in (Select ClassID from FS_NewsClass where DelFlag=1)"
	Conn.Execute(Sql)
	Sql = "Delete from FS_DownLoad where ClassID in (Select ClassID from FS_NewsClass where DelFlag=1)"
	Conn.Execute(Sql)
	Sql = "Delete from FS_NewsClass where DelFlag=1"
	Conn.Execute(Sql)
	Conn.Execute("Delete from FS_FreeJsFile where DelFlag=1")
	
elseif OperateType = "UnDoAll" then
	if Not JudgePopedomTF(Session("Name"),"P070105") then Call ReturnError1()
	Sql = "UpDate FS_News Set DelTF=0 where DelTF=1"
	Conn.Execute(Sql)
	Sql = "UpDate FS_NewsClass Set DelFlag=0 where DelFlag=1"
	Conn.Execute(Sql)
	'---------------重新生成相关自由JS------------------------
	Dim FunDelCreFreeJsObj,TemmpStrr,JsENameArr,Riker_ii,RikerCreTempObj,JSClassObj
	Set FunDelCreFreeJsObj = Conn.Execute("Select distinct JSName from FS_FreeJsFile where DelFlag=1")
	TemmpStrr = ""
	Do while Not FunDelCreFreeJsObj.eof
	If TemmpStrr = "" then
		TemmpStrr = FunDelCreFreeJsObj("JSName")
	Else
		TemmpStrr = TemmpStrr & "," & FunDelCreFreeJsObj("JSName")
	End If
		FunDelCreFreeJsObj.MoveNext
	Loop
	FunDelCreFreeJsObj.Close
	Set FunDelCreFreeJsObj = Nothing
	Conn.Execute("Update FS_FreeJsFile set DelFlag=0 where DelFlag=1")
	JsENameArr = Array("")
	JsENameArr = Split(TemmpStrr,",")
	Set JSClassObj = New JSClass
	JSClassObj.SysRootDir = RecaSysRootDir
	For Riker_ii = 0 to UBound(JsENameArr)
		Set RikerCreTempObj = Conn.Execute("Select Manner,EName from FS_FreeJS where EName='"&JsENameArr(Riker_ii)&"'")
		If Not RikerCreTempObj.eof then
			Select case RikerCreTempObj("Manner")
			 case "1"   JSClassObj.WCssA RikerCreTempObj("EName"),True
			 case "2"   JSClassObj.WCssB RikerCreTempObj("EName"),True
			 case "3"   JSClassObj.WCssC RikerCreTempObj("EName"),True
			 case "4"   JSClassObj.WCssD RikerCreTempObj("EName"),True
			 case "5"   JSClassObj.WCssE RikerCreTempObj("EName"),True
			 case "6"   JSClassObj.PCssA RikerCreTempObj("EName"),True
			 case "7"   JSClassObj.PCssB RikerCreTempObj("EName"),True
			 case "8"   JSClassObj.PCssC RikerCreTempObj("EName"),True
			 case "9"   JSClassObj.PCssD RikerCreTempObj("EName"),True
			 case "10"   JSClassObj.PCssE RikerCreTempObj("EName"),True
			 case "11"   JSClassObj.PCssF RikerCreTempObj("EName"),True
			 case "12"   JSClassObj.PCssG RikerCreTempObj("EName"),True
			 case "13"   JSClassObj.PCssH RikerCreTempObj("EName"),True
			 case "14"   JSClassObj.PCssI RikerCreTempObj("EName"),True
			 case "15"   JSClassObj.PCssJ RikerCreTempObj("EName"),True
			 case "16"   JSClassObj.PCssK RikerCreTempObj("EName"),True
			 case "17"   JSClassObj.PCssL RikerCreTempObj("EName"),True
		   End Select
	   End If
	   RikerCreTempObj.Close
	   Set RikerCreTempObj = Nothing
	Next
	Set JSClassObj = Nothing
	'---------------------------------------------------------
end if
Set MyFile = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>回收站</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2" onclick="ClickClassOrNews();" ondragstart="return false;" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width=35  align="center" alt="还原" onClick="Revert();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">还原</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="删除" onClick="DelClassOrNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="属性" onClick="ShowAttribute();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">属性</td>
		  <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="清空回收站" onClick="RecyleOperation('DelAll');" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">清空</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="全部还原" onClick="RecyleOperation('UnDoAll');" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">全部还原</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp; </td>
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
  <td valign="top"><table width="100%" border="0" cellpadding="2" cellspacing="0">
        <tr> 
          <td width="46%" height="26" class="ButtonListLeft"> 
            <div align="center">名称</div></td>
          <td width="15%" height="26" class="ButtonList"> 
            <div align="center">类型</div></td>
          <td width="22%" height="26" class="ButtonList"> 
            <div align="center">删除时间</div></td>
          <td width="17%" height="26" class="ButtonList"> 
            <div align="center">大小</div></td>
  </tr>
  <%
	Sql = "Select * from FS_NewsClass where DelFlag=1"
	Set RsClassObj = Conn.Execute(Sql)
	do while Not RsClassObj.Eof
		if AlreadyDelClassIDStr = "" then
			AlreadyDelClassIDStr = "'" & RsClassObj("ClassID") & "'"
		else
			AlreadyDelClassIDStr = AlreadyDelClassIDStr & "," & "'" & RsClassObj("ClassID") & "'"
		end if
%>

        <tr> 
          <td height="22">
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="../../Images/Folder/folderclosed.gif"></td>
                <td><span class="TempletItem" ContentTypeStr="Class" ContentID="<% = RsClassObj("ClassID") %>" align="center"><% = RsClassObj("ClassCName") %></span></td>
              </tr>
            </table>
            
          </td>
    <td><div align="center" class="TempletItem">
              栏目</div></td>
    <td><div align="center" class="TempletItem"><% = RsClassObj("DelTime") %></div></td>
          <td><div align="center">--</div></td>
	</tr>
  <%
		RsClassObj.MoveNext
	loop
	if AlreadyDelClassIDStr <> "" then
		Sql = "Select * from FS_News where DelTF=1 order by DelTime desc" 
	else
		Sql = "Select * from FS_News where DelTF=1 order by DelTime desc"
	end if
	Set RsNewsObj = Server.CreateObject(G_FS_RS)
	RsNewsObj.Open Sql,Conn,1,1
	SunNumAsp = RsNewsObj.RecordCount
	If Not RsNewsObj.eof then
	Dim page_size,page_no,page_total,record_all
	page_size = 20
	page_no = request.querystring("page_no")
	if page_no<=1 or page_no="" then page_no=1
	If Request.QueryString("page_no")="" then
		page_no=1
	end if
	RsNewsObj.PageSize=page_size
	page_total=RsNewsObj.PageCount
	if (cint(page_no)>page_total) then page_no=page_total
	If page_no=0 then page_no=1
	RsNewsObj.AbsolutePage=page_no
 	Dim i
	for i=1 to RsNewsObj.PageSize
		if RsNewsObj.eof then exit for		
		if RsNewsObj("HeadNewsTF")<>"1" and RsNewsObj("PicNewsTF")<>"1" then
		   NewsType = "文字新闻"
		   PicStr = "../../Images/Info/WordNews.gif"
		elseif RsNewsObj("HeadNewsTF")="1" then
		   NewsType = "标题新闻"
		   PicStr = "../../Images/Info/TitleNews.gif"
		else
		   NewsType = "图片新闻"
		   PicStr = "../../Images/Info/PicNews.gif"
		end if
%>
        <tr> 
          <td height="22"> <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="<% = PicStr %>"></td>
                <td><span class="TempletItem" ContentTypeStr="News" ContentID="<% = RsNewsObj("NewsID") %>" align="center"><% = GotTopic(RsNewsObj("Title"),40) %></span></td>
              </tr>
            </table>
		 </td>
          <td> <div align="center" class="TempletItem"> 
              <% = NewsType %>
            </div></td>
          <td> <div align="center" class="TempletItem"> 
              <% = RsNewsObj("DelTime") %>
            </div></td>
          <td><div align="center" class="TempletItem"> 
              <% = Len(RsNewsObj("Content")) %>
              字符 </div></td>
        </tr>
<%
		RsNewsObj.MoveNext
	Next
	End If
%>
</table>
</td>
</tr>
<tr> 
<td valign="middle" height="10">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="ButtonListLeft">
        <tr height="1"> 
          <td height="25"> <div align="right"> 
              <% =  "NO.<b>"& page_no &"</b>,&nbsp;&nbsp;" %>
              <% = "Totel:<b>"& page_total &"</b>,&nbsp;RecordCounts:<b>" & record_all &"</b>&nbsp;&nbsp;&nbsp;"%>
              <%
				if Page_Total=1 then
						response.Write "&nbsp;<img src=../../images/FirstPage.gif border=0 alt=首页></img>&nbsp;"
						response.Write "&nbsp;<img src=../../images/prePage.gif border=0 alt=上一页></img>&nbsp;"
						response.Write "&nbsp;<img src=../../images/nextPage.gif border=0 alt=下一页></img>&nbsp;"
						response.Write "&nbsp;<img src=../../images/endPage.gif border=0 alt=尾页></img>&nbsp;"
				else
					if cint(Page_No)<>1 and cint(Page_No)<>Page_Total then
						response.Write "&nbsp;<a href=?page_no=1&Keywords="&Request("Keywords")&"><img src=../../images/FirstPage.gif border=0 alt=首页></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&Keywords="&Request("Keywords")&"><img src=../../images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)&"&Keywords="&Request("Keywords")&"><img src=../../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../../images/endPage.gif border=0 alt=尾页></img></a>&nbsp;"
					elseif cint(Page_No)=1 then
						response.Write "&nbsp;<img src=../../images/FirstPage.gif border=0 alt=首页></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../../images/prePage.gif border=0 alt=上一页></img>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1)&"&Keywords="&Request("Keywords")&"><img src=../../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../../images/endpage.gif border=0 alt=尾页></img></a>&nbsp;"
					else
						response.Write "&nbsp;<a href=?page_no=1&Keywords="&Request("Keywords")&"><img src=../../images/FirstPage.gif border=0 alt=首页></img>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1)&"&Keywords="&Request("Keywords")&"><img src=../../images/prePage.gif border=0 alt=上一页></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../../images/nextPage.gif border=0 alt=下一页></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../../images/endpage.gif border=0 alt=尾页></img>&nbsp;"
					end if
				end if
				%>
              <select onChange="ChangePage(this.value);" style="width:50;" name="select">
                <% for i=1 to Page_Total %>
                <option <% if cint(Page_No) = i then Response.Write("selected")%> value="<% = i %>"> 
                <% = i %>
                </option>
                <% next %>
              </select>
            </div></td>
        </tr>
      </table></td>
	</tr>
</table>

</body>
</html>
<%
Set RsNewsObj = Nothing
Set RsClassObj = Nothing
Set Conn = Nothing
%>
<script language="JavaScript">
var ListObjArray=new Array();
//SelectedObj=null;
function document.onreadystatechange()
{
	IntialListObjArray();
}
function IntialListObjArray()
{
	var CurrObj=null,j=1;
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if (CurrObj.ContentID!=null)
		{
			ListObjArray[ListObjArray.length]=new NewsOrClassObj(CurrObj,j,false);
			j++;
		}
	}
}
function ClickClassOrNews()
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
function NewsOrClassObj(Obj,Index,Selected)
{
	this.Obj=Obj;
	this.Index=Index;
	this.Selected=Selected;
}
function Revert()
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
				if (SelectContentObj.ContentTypeStr=='Class')
				{
					OpenWindow('Frame.asp?PageTitle=回收站-还原&FileName=Revert.asp&OperateType=Class&ID='+SelectedContent,260,104,window);
				}
				else
				{
					OpenWindow('Frame.asp?PageTitle=回收站-还原&FileName=Revert.asp&OperateType=News&ID='+SelectedContent,260,104,window);
				}
			}
			location.href=location.href;
		}
		else alert('一次只能还原一个栏目或者一条新闻');
	}
	else alert('请选择栏目或新闻！');
}
function DelClassOrNews()
{
	var SelectedNews='',SelectedClass='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if ((ListObjArray[i].Obj.ContentID!=null)&&(ListObjArray[i].Obj.ContentTypeStr!=null))
			{
				if (ListObjArray[i].Obj.ContentTypeStr=='Class')
				{
					if (SelectedClass=='') SelectedClass=ListObjArray[i].Obj.ContentID;
					else  SelectedClass=SelectedClass+'***'+ListObjArray[i].Obj.ContentID;
				}
				else
				{
					if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.ContentID;
					else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.ContentID;
				} 
			}
		}
	}
	if ((SelectedNews!='')||(SelectedClass!=''))
	{
		OpenWindow('Frame.asp?PageTitle=回收站-删除&FileName=Del.asp&NewsID='+SelectedNews+'&ClassID='+SelectedClass,260,100,window);
		location.href=location.href;
	}
	else
	{
		alert('请选择删除内容');
	}
}
function ShowAttribute()
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
				if (SelectContentObj.ContentTypeStr=='Class')
				{
					OpenWindow('Frame.asp?FileName=ClassAttribute.asp&PageTitle=回收站-栏目属性&OperateType=Class&ID='+SelectedContent,250,270,window);
				}
				else
				{
					OpenWindow('Frame.asp?FileName=NewsAttribute.asp&PageTitle=回收站-新闻属性&OperateType=News&ID='+SelectedContent,250,270,window);
				}
			}
			location.href=location.href;
		}
		else alert('一次只能查看一个栏目或者一条新闻');
	}
	else alert('请选择栏目或新闻！');

}
function ChangePage(PageNum)
{
	window.location.href='?page_no='+PageNum;
}
function RecyleOperation(Type)
{
	var PromptInfo='';
	switch (Type)
	{
		case 'DelAll':
			PromptInfo='删除全部';
			break;
		case 'UnDoAll':
			PromptInfo='还原全部';
			break;
		default :
			alert('操作无效');
			return;
			break;
	}
	if (confirm('确定要'+PromptInfo+'吗？'))
	{
		location='?OperateType='+Type;
	}
}
</script>