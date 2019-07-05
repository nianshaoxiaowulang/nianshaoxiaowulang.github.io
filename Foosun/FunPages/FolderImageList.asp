<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<%
Dim CurrPath,FsoObj,SubFolderObj,FolderObj,FileObj,i,FsoItem,OType
Dim ParentPath,FileExtName,AllowShowExtNameStr
Dim ShowVirtualPath,sRootDir 
if SysRootDir<>"" then sRootDir="/" & SysRootDir else sRootDir=""
Set FsoObj = Server.CreateObject(G_FS_FSO)
OType = Request("Type")
if OType <> "" then
	Dim Path,PhysicalPath
	if OType = "DelFolder" then
		Path = Request("Path") 
		if Path <> "" then
			Path = Server.MapPath(Path)
			if FsoObj.FolderExists(Path) = true then FsoObj.DeleteFolder Path
		end if
	elseif OType = "DelFile" then
		Dim DelFileName
		Path = Request("Path") 
		DelFileName = Request("FileName") 
		if (DelFileName <> "") And (Path <> "") then
			Path = Server.MapPath(Path)
			if FsoObj.FileExists(Path & "\" & DelFileName) = true then FsoObj.DeleteFile Path & "\" & DelFileName
		end if
	elseif OType = "AddFolder" then
		Path = Request("Path")
		if Path <> "" then
			Path = Server.MapPath(Path)
			if FsoObj.FolderExists(Path) = True then
				Response.Write("<script>alert('Ŀ¼�Ѿ�����');</script>")
			else
				FsoObj.CreateFolder Path
			end if
		end if
	elseif OType = "FileReName" then
		Dim NewFileName,OldFileName
		Path = Request("Path")
		if Path <> "" then
			NewFileName = Request("NewFileName")
			OldFileName = Request("OldFileName")
			if (NewFileName <> "") And (OldFileName <> "") then
				PhysicalPath = Server.MapPath(Path) & "\" & OldFileName
				if FsoObj.FileExists(PhysicalPath) = True then
					PhysicalPath = Server.MapPath(Path) & "\" & NewFileName
					if FsoObj.FileExists(PhysicalPath) = False then
						Set FileObj = FsoObj.GetFile(Server.MapPath(Path) & "\" & OldFileName)
						FileObj.Name = NewFileName
						Set FileObj = Nothing
					end if
				end if
			end if
		end if
	elseif OType = "FolderReName" then
		Dim NewPathName,OldPathName
		Path = Request("Path")
		if Path <> "" then
			NewPathName = Request("NewPathName")
			OldPathName = Request("OldPathName")
			if (NewPathName <> "") And (OldPathName <> "") then
				PhysicalPath = Server.MapPath(Path) & "\" & OldPathName
				if FsoObj.FolderExists(PhysicalPath) = True then
					PhysicalPath = Server.MapPath(Path) & "\" & NewPathName
					if FsoObj.FolderExists(PhysicalPath) = False then
						Set FileObj = FsoObj.GetFolder(Server.MapPath(Path) & "\" & OldPathName)
						FileObj.Name = NewPathName
						Set FileObj = Nothing
					end if
				end if
			end if
		end if
	end if
end if


ShowVirtualPath = Request("ShowVirtualPath")
AllowShowExtNameStr = "jpg,txt,gif"
CurrPath = Request("CurrPath")
if CurrPath = "" then
	CurrPath = "/" & SysRootDir
	ParentPath = ""
else
	ParentPath = Mid(CurrPath,1,InstrRev(CurrPath,"/")-1)
	if ParentPath = "" then
		ParentPath = "/" & SysRootDir
	end if
end if
Set FolderObj = FsoObj.GetFolder(Server.MapPath(CurrPath))
Set SubFolderObj = FolderObj.SubFolders
Set FileObj = FolderObj.Files
Function CheckFileShowTF(AllowShowExtNameStr,ExtName)
	if ExtName="" then
		CheckFileShowTF = False
	else
		if InStr(1,AllowShowExtNameStr,ExtName) = 0 then
			CheckFileShowTF = False
		else
			CheckFileShowTF = True
		end if
	end if
End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ļ���Ŀ¼�б�</title>
</head>
<link href="../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript">
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
</script>
<body topmargin="0" leftmargin="0" onClick="SelectFolder();">
<table width="99%" border="0" align="center" cellpadding="2" cellspacing="0">
  <tr> 
    <td width="30%" class="ButtonListLeft"> <div align="center">����</div></td>
    <td width="40%" class="ButtonList"> <div align="center">����</div></td>
    <td class="ButtonList"> <div align="center">��С</div></td>
  </tr>
  <%
if (CurrPath <> sRootDir & "/" & UpFiles) and (CurrPath <> sRootDir & "/" & StyleFiles) and (ParentPath <> "") then
%>
  <tr title="�ϼ�Ŀ¼<% = ParentPath %>" onClick="SelectUpFolder(this);" Path="<% = ParentPath %>" onDblClick="OpenParentFolder(this);"> 
    <td><table width="62" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><font color="#FFFFFF"><img src="../Images/Folder/folderclosed.gif"></font></td>
          <td><font color="#FFFFFF">...</font></td>
        </tr>
      </table></td>
    <td><div align="center"><font color="#FFFFFF">-</font></div></td>
    <td><div align="center"><font color="#FFFFFF">-</font></div></td>
  </tr>
  <%
end if
for each FsoItem In SubFolderObj
%>
  <tr> 
    <td width="30%"><table border="0" cellspacing="0" cellpadding="0">
        <tr title="˫���������Ŀ¼"> 
          <td><img src="../Images/Folder/folderclosed.gif"></td>
          <td> <span class="TempletItem" Path="<% = FsoItem.name %>" onDblClick="OpenFolder(this);"> 
            <% = FsoItem.name %>
            </span> </td>
        </tr>
      </table></td>
    <td><div align="center">�ļ���</div></td>
    <td><div align="center"> 
        <% = FsoItem.Size %>
      </div></td>
  </tr>
  <%
next
for each FsoItem In FileObj
	FileExtName = LCase(Mid(FsoItem.name,InstrRev(FsoItem.name,".")+1))
	if True then 'CheckFileShowTF(AllowShowExtNameStr,FileExtName) = 
%>
  <tr title="˫�����ѡ����ļ�"> 
    <td> <span class="TempletItem" File="<% = FsoItem.name %>" onDblClick="SetFile(this);" onClick="SelectFile(this);"> 
      <% = FsoItem.name %>
      </span> </td>
    <td><div align="center"> 
        <% = FsoItem.Type %>
      </div></td>
    <td><div align="center"> 
        <% = FsoItem.Size %>
        �ֽ� </div></td>
  </tr>
  <%
  	end if
next
%>
</table>
</body>
</html>
<%
Set FsoObj = Nothing
Set SubFolderObj = Nothing
Set FileObj = Nothing
%>
<script language="JavaScript">
var CurrPath='<% = CurrPath %>';
var SysRootDir='<% = SysRootDir %>';
var ShowVirtualPath='<% = ShowVirtualPath %>';
var SelectedObj=null;
var ContentMenuArray=new Array();
DocumentReadyTF=false;
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	InitialClassListContentMenu();
	DocumentReadyTF=true;
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function ContentMenuShowEvent()
{
	SelectFolder();
}
function InitialClassListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddFolderOperation();",'�½�Ŀ¼','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("if (confirm('ȷ��Ҫɾ����')==true) parent.DelFolderFile();",'ɾ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditFolder();",'������','disabled');
}
function SelectFolder()
{
	Obj=event.srcElement,DisabledContentMenuStr='';
	if (SelectedObj!=null) SelectedObj.className='TempletItem';
	if ((Obj.Path!=null)||(Obj.File!=null))
	{
		Obj.className='TempletSelectItem';
		SelectedObj=Obj;
	}
	else SelectedObj=null;
	if (SelectedObj!=null)	DisabledContentMenuStr='';
	else DisabledContentMenuStr=',ɾ��,������,';
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function SelectFile(Obj)
{
	for (var i=0;i<document.all.length;i++)
	{
		if (document.all(i).className=='TempletSelectItem') document.all(i).className='TempletItem';
	}
	Obj.className='TempletSelectItem';
	PreviewFile(Obj);
}
function OpenParentFolder(Obj)
{
	location.href='FolderImageList.asp?CurrPath='+Obj.Path;
	SearchOptionExists(parent.document.all.FolderSelectList,Obj.Path);
}

function OpenFolder(Obj)
{
	var SubmitPath='';
	if (CurrPath=='/') SubmitPath=CurrPath+Obj.Path;
	else SubmitPath=CurrPath+'/'+Obj.Path;
	location.href='FolderImageList.asp?CurrPath='+SubmitPath;
	AddFolderList(parent.document.all.FolderSelectList,SubmitPath,SubmitPath);
}

function SelectUpFolder(Obj)
{
	for (var i=0;i<document.all.length;i++)
	{
		if (document.all(i).className=='TempletSelectItem') document.all(i).className='TempletItem';
	}
	Obj.className='TempletSelectItem';
}
function PreviewFile(Obj)
{
	var Url='';
	var Path=escape();
	if (CurrPath=='/') Path=escape(CurrPath+Obj.File);
	else Path=escape(CurrPath+'/'+Obj.File);
	Url='PreviewImage.asp?FilePath='+Path;
	parent.frames["PreviewArea"].location=Url;
}
function AddFolderList(SelectObj,Lable,LableContent)
{
	var i=0,AddOption;
	if (!SearchOptionExists(SelectObj,Lable))
	{
		AddOption = document.createElement("OPTION");
		AddOption.text=Lable;
		AddOption.value=LableContent;
		SelectObj.add(AddOption);
		SelectObj.options(SelectObj.length-1).selected=true;
	}
}
function SearchOptionExists(Obj,SearchText)
{
	var i;
	for(i=0;i<Obj.length;i++)
	{
		if (Obj.options(i).text==SearchText)
		{
			Obj.options(i).selected=true;
			return true;
		}
	}
	return false;
}
function SetFile(Obj)
{
	//if (ShowVirtualPath=='')
	//{
		var PathInfo='',TempPath='';
		if (SysRootDir!='')
		{
			TempPath=CurrPath;
			PathInfo=TempPath.substr(TempPath.indexOf(SysRootDir)+SysRootDir.length);
		}
		else
		{
			PathInfo=CurrPath;
		}
	//}
	//else PathInfo=CurrPath;
	if (CurrPath=='/')	window.returnValue=PathInfo+Obj.File;
	else window.returnValue=PathInfo+'/'+Obj.File;
	window.close();
}
window.onunload=CheckReturnValue;
function CheckReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
function AddFolderOperation()
{
	var ReturnValue=prompt('�½�Ŀ¼����','');
	if ((ReturnValue!='') && (ReturnValue!=null))
		window.location.href='?Type=AddFolder&Path='+CurrPath+'/'+ReturnValue+'&CurrPath='+CurrPath;
}
function DelFolderFile()
{
	if (SelectedObj!=null)
	{
		if (SelectedObj.Path!=null) window.location.href='?Type=DelFolder&Path='+CurrPath+'/'+SelectedObj.Path+'&CurrPath='+CurrPath;
		if (SelectedObj.File!=null) window.location.href='?Type=DelFile&Path='+CurrPath+'&FileName='+SelectedObj.File+'&CurrPath='+CurrPath;
	}
	else alert('��ѡ��Ҫɾ����Ŀ¼');
}
function EditFolder()
{
	var ReturnValue='';
	if (SelectedObj!=null)
	{
		if (SelectedObj.Path!=null)
		{
			ReturnValue=prompt('�޸ĵ����ƣ�',SelectedObj.Path);
			if ((ReturnValue!='') && (ReturnValue!=null)) window.location.href='?Type=FolderReName&Path='+CurrPath+'&CurrPath='+CurrPath+'&OldPathName='+SelectedObj.Path+'&NewPathName='+ReturnValue;
		}
		if (SelectedObj.File!=null)
		{
			ReturnValue=prompt('�޸ĵ����ƣ�',SelectedObj.File);
			if ((ReturnValue!='') && (ReturnValue!=null)) window.location.href='?Type=FileReName&Path='+CurrPath+'&CurrPath='+CurrPath+'&OldFileName='+SelectedObj.File+'&NewFileName='+ReturnValue;
		}
	}
	else alert('����дҪ������Ŀ¼����');
}
</script>