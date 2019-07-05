<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030700") then Call ReturnError1()
Dim NewsTempletPath,FS,FolderObj,SubFolderObj,FileObj,FileItem,FolderItem,FileIconDic,ParentPath,NewsTempleConfigObj
Set NewsTempleConfigObj = Conn.Execute("Select DoMain from FS_Config")
NewsTempletPath = Request("Path")
if NewsTempletPath = "" then
	if SysRootDir = "" then
		NewsTempletPath = "/"
		ParentPath = ""
	else
		NewsTempletPath = "/" & SysRootDir
		ParentPath = ""
	end if
else
	ParentPath = Mid(NewsTempletPath,1,InstrRev(NewsTempletPath,"/")-1)
	if ParentPath = "" then
		ParentPath = "/" & SysRootDir
	end if
end if
Set FS = Server.CreateObject(G_FS_FSO)
Set FolderObj = FS.GetFolder(Server.MapPath(NewsTempletPath))
Set SubFolderObj = FolderObj.SubFolders
Set FileObj = FolderObj.Files
Set FileIconDic = CreateObject("Scripting.Dictionary")
FileIconDic.Add "txt","../../Images/FileIcon/txt.gif"
FileIconDic.Add "gif","../../Images/FileIcon/gif.gif"
FileIconDic.Add "exe","../../Images/FileIcon/exe.gif"
FileIconDic.Add "asp","../../Images/FileIcon/asp.gif"
FileIconDic.Add "html","../../Images/FileIcon/html.gif"
FileIconDic.Add "htm","../../Images/FileIcon/html.gif"
FileIconDic.Add "jpg","../../Images/FileIcon/jpg.gif"
FileIconDic.Add "jpeg","../../Images/FileIcon/jpg.gif"
FileIconDic.Add "pl","../../Images/FileIcon/perl.gif"
FileIconDic.Add "perl","../../Images/FileIcon/perl.gif"
FileIconDic.Add "zip","../../Images/FileIcon/zip.gif"
FileIconDic.Add "rar","../../Images/FileIcon/zip.gif"
FileIconDic.Add "gz","../../Images/FileIcon/zip.gif"
FileIconDic.Add "doc","../../Images/FileIcon/doc.gif"
FileIconDic.Add "xml","../../Images/FileIcon/xml.gif"
FileIconDic.Add "xsl","../../Images/FileIcon/xml.gif"
FileIconDic.Add "dtd","../../Images/FileIcon/xml.gif"
FileIconDic.Add "vbs","../../Images/FileIcon/vbs.gif"
FileIconDic.Add "js","../../Images/FileIcon/vbs.gif"
FileIconDic.Add "wsh","../../Images/FileIcon/vbs.gif"
FileIconDic.Add "sql","../../Images/FileIcon/script.gif"
FileIconDic.Add "bat","../../Images/FileIcon/script.gif"
FileIconDic.Add "tcl","../../Images/FileIcon/script.gif"
FileIconDic.Add "eml","../../Images/FileIcon/mail.gif"
FileIconDic.Add "swf","../../Images/FileIcon/flash.gif"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ģ���б�</title> 
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<script language="JavaScript">
var CurrPath=escape('<% = NewsTempletPath %>');
var ListObjArray = new Array();
var ContentMenuArray=new Array();
var DocumentReadyTF=false;
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	IntialListObjArray();
	InitialClassListContentMenu();
	DocumentReadyTF=true;
}
function InitialClassListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.PreEditTemplet();",'���ӱ༭','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.TxtEditTemplet();",'�ı��༭','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddFile();",'�½��ļ�','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditFolder();",'������','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelFolderOrFile();",'ɾ��','');
}
function ContentMenuShowEvent()
{
	ChangeTempletMenuStatus();
}
function ChangeTempletMenuStatus()
{
	var EventObjInArray=false,SelectFolder='',SelectFile='',DisabledContentMenuStr='';
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
			if (ListObjArray[i].Obj.Path!=null)
			{
				if (SelectFolder=='') SelectFolder=ListObjArray[i].Obj.Path;
				else SelectFolder=SelectFolder+'***'+ListObjArray[i].Obj.Path
			}
			if (ListObjArray[i].Obj.File!=null)
			{
				if (SelectFile=='') SelectFile=ListObjArray[i].Obj.File;
				else SelectFile=SelectFile+'***'+ListObjArray[i].Obj.File
			}
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
					if (ListObjArray[i].Obj.Path!=null)
					{
						if (SelectFolder=='') SelectFolder=ListObjArray[i].Obj.Path;
						else SelectFolder=SelectFolder+'***'+ListObjArray[i].Obj.Path
					}
					if (ListObjArray[i].Obj.File!=null)
					{
						if (SelectFile=='') SelectFile=ListObjArray[i].Obj.File;
						else SelectFile=SelectFile+'***'+ListObjArray[i].Obj.File
					}
				}
			}
		}
	}
	if ((SelectFolder=='')&&(SelectFile=='')) DisabledContentMenuStr=',���ӱ༭,ɾ��,Ԥ��,�ı��༭,������,';
	else
	{
		if (SelectFile=='') DisabledContentMenuStr=',���ӱ༭,Ԥ��,�ı��༭,';
		else
		{
			if (SelectFile.indexOf('***')!=-1) DisabledContentMenuStr=',���ӱ༭,Ԥ��,�ı��༭,';
			else
			{
				if (SelectFolder!='') DisabledContentMenuStr=',���ӱ༭,Ԥ��,�ı��༭,';
				else DisabledContentMenuStr='';
			}
		}
	}
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function AddFolder()
{
	OpenWindow('../../FunPages/Frame.asp?PageTitle=������Ŀ&FileName=AddFolder.asp&Path='+CurrPath,200,90,window);
}
function AddFile()
{
	OpenWindow('../../FunPages/Frame.asp?PageTitle=�����ļ�&FileName=AddFile.asp&Path='+CurrPath,200,90,window);
}
function DelFolderOrFile()
{
	var SelectedFolder='',SelectedFile='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.Path!=null)
			{
				if (SelectedFolder=='') SelectedFolder=ListObjArray[i].Obj.Path;
				else  SelectedFolder=SelectedFolder+'***'+ListObjArray[i].Obj.Path;
			}
			if (ListObjArray[i].Obj.File!=null)
			{
				if (SelectedFile=='') SelectedFile=ListObjArray[i].Obj.File;
				else  SelectedFile=SelectedFile+'***'+ListObjArray[i].Obj.File;
			}
		}
	}
	SelectedFile=escape(SelectedFile);
	SelectedFolder=escape(SelectedFolder);
	if ((SelectedFile!='')||(SelectedFolder!=''))
	OpenWindow('../../FunPages/Frame.asp?PageTitle=ɾ���ļ�&Path='+CurrPath+'&FileName=DelFolderAndFile.asp&DelFolder='+SelectedFolder+'&DelFile='+SelectedFile,200,90,window);
	else alert('û��ѡ��Ҫɾ����Ŀ¼�����ļ�');
}
function ImportTempletFile()
{
	OpenWindow('../../FunPages/Frame.asp?FileName=UpFileForm.asp&PageTitle=�ϴ��ļ�&Path='+CurrPath,350,170,window);
}
function TxtEditTemplet()
{
	var SelectedFolder='',SelectedFile='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.Path!=null)
			{
				if (SelectedFolder=='') SelectedFolder=ListObjArray[i].Obj.Path;
				else  SelectedFolder=SelectedFolder+'***'+ListObjArray[i].Obj.Path;
			}
			if (ListObjArray[i].Obj.File!=null)
			{
				if (SelectedFile=='') SelectedFile=ListObjArray[i].Obj.File;
				else  SelectedFile=SelectedFile+'***'+ListObjArray[i].Obj.File;
			}
		}
	}
	SelectedFile=escape(SelectedFile);
	if (SelectedFile!='')
	{
		if (SelectedFile.indexOf('***')==-1) location='../../Editer/TextEditer.asp?Path='+CurrPath+'&FileName='+SelectedFile;
		else alert('һ��ֻ�ܹ��༭һ���ļ�');
	}
	else alert('��ѡ��Ҫ�༭���ļ�');
}
function PreEditTemplet()
{
	var SelectedFolder='',SelectedFile='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.Path!=null)
			{
				if (SelectedFolder=='') SelectedFolder=ListObjArray[i].Obj.Path;
				else  SelectedFolder=SelectedFolder+'***'+ListObjArray[i].Obj.Path;
			}
			if (ListObjArray[i].Obj.File!=null)
			{
				if (SelectedFile=='') SelectedFile=ListObjArray[i].Obj.File;
				else  SelectedFile=SelectedFile+'***'+ListObjArray[i].Obj.File;
			}
		}
	}
	SelectedFile=escape(SelectedFile);
	if (SelectedFile!='')
	{
		if (SelectedFile.indexOf('***')==-1) location='Templet_Edit.asp?Path='+CurrPath+'&FileName='+SelectedFile;
		else alert('һ��ֻ�ܹ��༭һ���ļ�');
	}
	else alert('��ѡ��Ҫ�༭���ļ�');
}
function PreviewTemplet()
{
	var SelectedFolder='',SelectedFile='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.Path!=null)
			{
				if (SelectedFolder=='') SelectedFolder=ListObjArray[i].Obj.Path;
				else  SelectedFolder=SelectedFolder+'***'+ListObjArray[i].Obj.Path;
			}
			if (ListObjArray[i].Obj.File!=null)
			{
				if (SelectedFile=='') SelectedFile=ListObjArray[i].Obj.File;
				else  SelectedFile=SelectedFile+'***'+ListObjArray[i].Obj.File;
			}
		}
	}
	if (SelectedFile!='')
	{
		if (SelectedFile.indexOf('***')==-1) window.open(CurrPath+'/'+escape(SelectedFile));
		else alert('һ��ֻ��Ԥ����һ���ļ�');
	}
	else alert('��ѡ��ҪԤ�����ļ�');
}
function EditFolder()
{
	var ReturnValue='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.Path!=null)
			{
				ReturnValue=prompt('�޸ĵ����ƣ�',ListObjArray[i].Obj.Path);
				if ((ReturnValue!='') && (ReturnValue!=null)) window.location.href='../../FunPages/FolderFileReName.asp?Type=FolderReName&Path='+CurrPath+'&CurrPath='+CurrPath+'&OldPathName='+ListObjArray[i].Obj.Path+'&NewPathName='+ReturnValue;
			}
			if (ListObjArray[i].Obj.File!=null)
			{
				ReturnValue=prompt('�޸ĵ����ƣ�',ListObjArray[i].Obj.File);
				if ((ReturnValue!='') && (ReturnValue!=null)) window.location.href='../../FunPages/FolderFileReName.asp?Type=FileReName&Path='+CurrPath+'&CurrPath='+CurrPath+'&OldPathName='+ListObjArray[i].Obj.File+'&NewPathName='+ReturnValue;
			}
		}
	}
}
</script>
<body topmargin="2" leftmargin="2" onClick="SelectFolderOrFile();" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=55 align="center" alt="����Ŀ¼" onClick="AddFolder();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����Ŀ¼</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="�����ļ�" onClick="AddFile();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�����ļ�</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="ɾ��" onClick="DelFolderOrFile();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ɾ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����ģ��" onClick="ImportTempletFile();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td width=2 class="Gray">|</td>
          <td width=55 align="center" alt="�ı��༭" onClick="TxtEditTemplet();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�ı��༭</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="���ӱ༭" onClick="PreEditTemplet();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">���ӱ༭</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="Ԥ��" onClick="PreviewTemplet();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">Ԥ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="" border="0" cellpadding="1" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr> 
    <td height="26" class="ButtonListLeft"> <div align="center">����</div></td>
    <td height="20" class="ButtonList"> <div align="center">����</div></td>
    <td height="20" class="ButtonList"> <div align="center">��С</div></td>
    <td height="20" class="ButtonList"> <div align="center">����ʱ��</div></td>
    <td height="20" class="ButtonList"> <div align="center">����޸�ʱ��</div></td>
  </tr>
  <% if ParentPath <> "/" & SysRootDir and ParentPath <> "/" then %>
  <tr> 
    <td height=""><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../../Images/arrow.gif" width="18" height="18"></td>
          <td><span class="TempletItem" title="�ϼ�Ŀ¼<% = ParentPath %>" onDblClick="OpenParentFolder(this);" Path="<% = ParentPath %>">�ϼ�Ŀ¼</span></td>
        </tr>
      </table></td>
    <td height="20">&nbsp;</td>
    <td height="20">&nbsp;</td>
    <td height="20">&nbsp;</td>
    <td height="20">&nbsp;</td>
  </tr>
  <%
end if
for Each FileItem In SubFolderObj
%>
  <tr> 
    <td height=""> 
      <table border="0" cellspacing="0" cellpadding="0">
        <tr title="˫���������Ŀ¼"> 
          <td><img src="../../Images/Folder/folderclosed.gif"></td>
          <td><span class="TempletItem" Path="<% = FileItem.name %>" onDblClick="OpenFolder(this);"> 
            <% = FileItem.name %>
            </span> </td>
        </tr>
      </table></td>
    <td> 
      <div align="center">�ļ���</div></td>
    <td> 
      <div align="center"> 
        <% = FileItem.Size %>
      </div></td>
    <td> 
      <div align="center"> 
        <% = FileItem.DateCreated %>
      </div></td>
    <td> 
      <div align="center"> 
        <% = FileItem.DateLastModified %>
      </div></td>
  </tr>
  <%
Next
for Each FileItem In FileObj
	Dim FileIcon,FileExtName
	FileExtName = Mid(CStr(FileItem.Name),Instr(CStr(FileItem.Name),".")+1)
	'/////////////////////////////lzpֻ��ʾĿ¼��ģ���ļ��������ļ�����ʾ
	if lcase(FileExtName)="html" or lcase(FileExtName)="htm" or lcase(FileExtName)="sty"then 
	'///////////////////////////////
		FileIcon = FileIconDic.Item(LCase(FileExtName))
		if FileIcon = "" then
			FileIcon = "../../Images/FileIcon/unknown.gif"
		end if
%>
	  <tr style="background:white;cursor:default;"> 
		<td> 
		  <table border="0" cellspacing="0" cellpadding="0">
			<tr title="˫���������Ŀ¼"> 
			  
          <td><img src="<% = FileIcon %>"></td>
			  <td><span File="<% = FileItem.Name %>"><% = FileItem.Name %></span></td>
			</tr>
		  </table></td>
		<td> 
		  <div align="center"> 
			<% = FileItem.Type %>
		  </div></td>
		<td> 
		  <div align="center"> 
			<% = FileItem.Size %>
			�ֽ�</div></td>
		<td> 
		  <div align="center"> 
			<% = FileItem.DateCreated %>
		  </div></td>
		<td> 
		  <div align="center"> 
			<% = FileItem.DateLastModified %>
		  </div></td>
	  </tr>
	  <%
  else
  end if
next
%>
</table>
</body>
</html>
<%
Set Conn = Nothing
Set FS = Nothing
Set FolderObj = Nothing
Set FileObj = Nothing
Set SubFolderObj = Nothing
Set FileIconDic = Nothing
%>
<script>
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
		if ((CurrObj.Path!=null)||(CurrObj.File!=null))
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectFolderOrFile()
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
function OpenFolder(Obj)
{
	var SubmitPath='';
	if (CurrPath=='/') SubmitPath=escape(CurrPath+Obj.Path);
	else SubmitPath=escape(CurrPath+'/'+Obj.Path);
	location.href='NewsTemplet_List.asp?Path='+SubmitPath;
}
function OpenParentFolder(Obj)
{
	location.href='NewsTemplet_List.asp?Path='+Obj.Path;
}
</script>