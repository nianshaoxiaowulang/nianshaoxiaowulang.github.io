<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
'==============================================================================
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System v3.1 
'���¸��£�2004.12
'==============================================================================
'��ҵע����ϵ��028-85098980-601,602 ����֧�֣�028-85098980-606��607,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��159410,655071,66252421
'����֧��:���г���ʹ�����⣬�����ʵ�bbs.foosun.net���ǽ���ʱ�ش���
'���򿪷�����Ѷ������ & ��Ѷ���������
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.net  ��ʾվ�㣺test.cooin.com    
'��վ����ר����www.cooin.com
'==============================================================================
'��Ѱ汾����������ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'==============================================================================
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P060703") then Call ReturnError1()
Dim TempType,RsJSObj,JsSql,JSType,FileNum,TempNumStr,TempObj,MannerStr,JsEName,JSFlag
 JsSql = "select * from FS_FreeJS order by Type asc,ID asc"
 JSFlag = "����JS�б�"
Set RsJSObj = Server.CreateObject(G_FS_RS)
RsJSObj.Open JsSql,conn,1,1
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����JS�б�</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" leftmargin="2" onclick="SelectJS();"  ondragstart="return false;" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="���ô���" onClick="GetCode();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
	  <tr>
	  <td valign="top">
	  <table width="100%" border="0" cellpadding="2" cellspacing="0">
        <tr>
			
          <td width="21%" height="26" class="ButtonListLeft">
<div align="center">����</div></td>
			
          <td width="18%" height="26" class="ButtonList">
<div align="center">����</div></td>
			
          <td width="20%" height="26" class="ButtonList">
<div align="center">��ʽ</div></td>
			
          <td width="21%" height="26" class="ButtonList">
<div align="center">��������</div></td>
			
          <td width="20%" height="26" class="ButtonList">
<div align="center">���ʱ��</div></td> 
		  </tr>
  <%
if Not RsJSObj.Eof then
  Dim page_no,page_total,record_all,PageNums,i
	page_no=request.querystring("page_no")
	if page_no<=1 or page_no="" then page_no=1
	If Request.QueryString("page_no")="" then
		page_no=1
	end if
	RsJSObj.PageSize=20
	page_total=RsJSObj.PageCount
	if (cint(page_no)>page_total) then page_no=page_total
	RsJSObj.AbsolutePage=page_no
	record_all=RsJSObj.RecordCount
	for i=1 to RsJSObj.PageSize
		if RsJSObj.eof then exit for
		select  case RsJSObj("Type")
			case "0"  JSType = "����"
			case "1"  JSType = "ͼƬ"
		 end select
		Set TempObj = Conn.Execute("select count(ID) from FS_FreeJsFile where JSName='"&RsJSObj("EName")&"'")
		if TempObj.eof then
			FileNum = "0"
		else
			FileNum = TempObj(0)
		end if
		TempNumStr = FileNum&"/"&RsJSObj("NewsNum")
		Select case RsJSObj("Manner")
		   case "1" MannerStr = "��ʽA"
		   case "2" MannerStr = "��ʽB"
		   case "3" MannerStr = "��ʽC"
		   case "4" MannerStr = "��ʽD"
		   case "5" MannerStr = "��ʽE"
		   case "6" MannerStr = "��ʽA"
		   case "7" MannerStr = "��ʽB"
		   case "8" MannerStr = "��ʽC"
		   case "9" MannerStr = "��ʽD"
		   case "10" MannerStr = "��ʽE"
		   case "11" MannerStr = "��ʽF"
		   case "12" MannerStr = "��ʽG"
		   case "13" MannerStr = "��ʽH"
		   case "14" MannerStr = "��ʽI"
		   case "15" MannerStr = "��ʽJ"
		   case "16" MannerStr = "��ʽK"
		   case "17" MannerStr = "��ʽL"
		End Select
			%>
			  <tr> 
				
          <td> <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="../../Images/Folder/folderclosed.gif"></td>
                <td><span JsID="<%=RsJSObj("ID")%>" class="TempletItem"><%=RsJSObj("CName")%></span></td>
              </tr>
            </table>
           </td>
          <td> 
            <div align="center"><%=JSType%></div></td>
				
          <td> 
            <div align="center"><%=MannerStr%></div></td>
				
          <td> 
            <div align="center"><%=TempNumStr%></div></td>
				
          <td> 
            <div align="center"><%=RsJSObj("AddTime")%></div></td>
			  </tr>
			  <%
  
		RsJSObj.MoveNext
	next
end if
%>
</table>
</td>
<%if page_total>1 then%>
</tr>
 <tr> 
<td valign="middle"  height="10">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr height="1">
		  <td width="42%" height="18"><table width="99%" border="0" cellspacing="0" cellpadding="0">
			  <tr>
			  </tr>
		   </table> </td>
		
		<td width="51%" height="25" valign="middle"> <div align="right">
			<% =  "NO.<b>"& page_no &"</b>,&nbsp;&nbsp;" %>
			<% = "Totel:<b>"& page_total &"</b>,&nbsp;RecordCounts:<b>" & record_all &"</b>&nbsp;&nbsp;&nbsp;"%>
			<%
				if Page_Total=1 then
						response.Write "&nbsp;<img src=../images/First1.gif border=0 alt=��ҳ></img>&nbsp;"
						response.Write "&nbsp;<img src=../images/prePage.gif border=0 alt=��һҳ></img>&nbsp;"
						response.Write "&nbsp;<img src=../images/nextPage.gif border=0 alt=��һҳ></img>&nbsp;"
						response.Write "&nbsp;<img src=../images/endPage.gif border=0 alt=βҳ></img>&nbsp;"
				else
					if cint(Page_No)<>1 and cint(Page_No)<>Page_Total then
						response.Write "&nbsp;<a href=?page_no=1" &"&Keywords="&Request("Keywords")&"><img src=../images/First1.gif border=0 alt=��ҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1) &"&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1) &"&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../images/endPage.gif border=0 alt=βҳ></img></a>&nbsp;"
					elseif cint(Page_No)=1 then
						response.Write "&nbsp;<img src=../images/First1.gif border=0 alt=��ҳ></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../images/prePage.gif border=0 alt=��һҳ></img>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)+1) &"&Keywords="&Request("Keywords")&"><img src=../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="& Page_Total &"&Keywords="&Request("Keywords")&"><img src=../images/endpage.gif border=0 alt=βҳ></img></a>&nbsp;"
					else
						response.Write "&nbsp;<a href=?page_no=1" &"&Keywords="&Request("Keywords")&"><img src=../images/First1.gif border=0 alt=��ҳ></img>&nbsp;"
						response.Write "&nbsp;<a href=?page_no="&cstr(cint(Page_No)-1) &"&Keywords="&Request("Keywords")&"><img src=../images/prePage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../images/nextPage.gif border=0 alt=��һҳ></img></a>&nbsp;"
						response.Write "&nbsp;<img src=../images/endpage.gif border=0 alt=βҳ></img>&nbsp;"
					end if
				end if
				%>
		</div></td>
		<td width="7%" valign="middle"><select onChange="ChangePage(this.value);" style="width:50;" name="select">
		  <% for i=1 to Page_Total %>
		  <option <% if cint(Page_No) = i then Response.Write("selected")%> value="<% = i %>">
		  <% = i %>
		  </option>
		  <% next %>
		</select></td>
	  </tr>
	</table></td>
	</tr>
	<%end if%>
	</table>
</body>
</html>
<%
RsJSObj.close
set RsJSObj=nothing
Set Conn = Nothing
%>
<script>
var DocumentReadyTF=false;
var ListObjArray = new Array();
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	IntialListObjArray();
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Attribute();','����','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.GetCode();','���ô���','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','ˢ��','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'��ҳ��·������\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','·������','');
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
	var EventObjInArray=false,SelectAds='',DisabledContentMenuStr='';
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
			if (SelectAds=='') SelectAds=ListObjArray[i].Obj.JsID;
			else SelectAds=SelectAds+'***'+ListObjArray[i].Obj.JsID;
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
					if (SelectAds=='') SelectAds=ListObjArray[i].Obj.JsID;
					else SelectAds=SelectAds+'***'+ListObjArray[i].Obj.JsID;
				}
			}
		}
	}
	if (SelectAds=='') DisabledContentMenuStr=',����,���ô���,';
	else
	{
		if (SelectAds.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',����,���ô���,'
	}
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
		if (CurrObj.JsID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectJS()
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
function GetCode()
{
	var SelectedJS='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.JsID!=null)
			{
				if (SelectedJS=='') SelectedJS=ListObjArray[i].Obj.JsID;
				else  SelectedJS=SelectedJS+'***'+ListObjArray[i].Obj.JsID;
			}
		}
	}
	if (SelectedJS!='')
	{
		if (SelectedJS.indexOf('***')==-1) OpenWindow('Frame.asp?PageTitle=��ȡJS���ô���&FileName=UseCode.asp&JSName=Ename&JSTable=FS_FreeJS&JsID='+SelectedJS,360,140,window);
		else alert('��ѡ��һ��JS');
	}
	else alert('��ѡ��JS');
}
function Attribute()
{
	var SelectedJS='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.JsID!=null)
			{
				if (SelectedJS=='') SelectedJS=ListObjArray[i].Obj.JsID;
				else  SelectedJS=SelectedJS+'***'+ListObjArray[i].Obj.JsID;
			}
		}
	}
	if (SelectedJS!='')
	{
		if (SelectedJS.indexOf('***')==-1) OpenWindow('Frame.asp?PageTitle=JS����&FileName=JsContent.asp&ID='+SelectedJS,360,190,window);
		else alert('��ѡ��һ��JS');
	}
	else alert('��ѡ��JS');
}
function ChangePage(PageNum)
{
	window.location.href='?page_no='+PageNum;
}
</script>
