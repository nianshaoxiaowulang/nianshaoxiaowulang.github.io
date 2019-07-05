<% option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../Refresh/Function.asp" -->
<%
Dim DBC,Conn,ClassParentID
Set DBC=new databaseclass
Set Conn=DBC.openconnection()
Set DBC=nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<!--#include file="../Inc/Cls_RefreshJs.asp" -->
<%
if Not ((JudgePopedomTF(Session("Name"),"P060601")) OR (JudgePopedomTF(Session("Name"),"P060501"))) then Call ReturnError1()
Dim Types
Types = Request("Types")
Dim TeempSysRootDir
If SysRootDir = "" then
	TeempSysRootDir = ""
Else
	TeempSysRootDir = SysRootDir & "/"
End If

Dim TempClassListStr
	TempClassListStr = ClassList
Function ClassList()
	Dim Rs
	Set Rs = Conn.Execute("select ClassID,ClassCName from FS_newsclass where ParentID = '0' and DelFlag=0 order by AddTime desc")
	do while Not Rs.Eof
		ClassList = ClassList & "<option value="&Rs("ClassID")&"" & ">" & Rs("ClassCName") & chr(10) & chr(13)
		ClassList = ClassList & ChildClassList(Rs("ClassID"),"")
		Rs.MoveNext	
	loop
	Rs.Close
	Set Rs = Nothing
End Function
Function ChildClassList(ClassID,Temp)
	Dim TempRs,TempStr
	Set TempRs = Conn.Execute("Select ClassID,ClassCName,ChildNum from FS_NewsClass where ParentID = '" & ClassID & "' and DelFlag=0 order by AddTime desc ")
	TempStr = Temp & " - "
	do while Not TempRs.Eof
		if TempRs("ChildNum") = 0 then
			ChildClassList = ChildClassList & "<option value="&TempRs("ClassID")&"" & ">" & TempStr & TempRs("ClassCName") & "</option>"& chr(10) & chr(13)
		else
			ChildClassList = ChildClassList & "<option value="&TempRs("ClassID")&"" & ">" & TempStr & TempRs("ClassCName") & "</option>"& chr(10) & chr(13)
		end if
		ChildClassList = ChildClassList & ChildClassList(TempRs("ClassID"),TempStr)
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>栏目JS添加</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2">
<form action="" method="post" name="ClassJSForm">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="document.ClassJSForm.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
          <td>&nbsp; <input name="action" type="hidden" id="action" value="add"> 
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#E6E6E6">
    <tr bgcolor="#FFFFFF"> 
      <td width="15%" height="26">&nbsp;&nbsp;&nbsp;&nbsp;中文名称</td>
      <td width="35%"> 
        <input name="FileCName" type="text" id="FileCName" style="width:90%" value="<%=Request("FileCName")%>"></td>
      <td width="15%">&nbsp;&nbsp;&nbsp;&nbsp;文件名称</td>
      <td width="35%"> 
        <input name="FileName" type="text" id="FileName" style="width:90%" value="<%=Request("FileName")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;栏目名称</td>
      <td> 
        <select name="ClassID" style="width:90%" <%If Types = "System" then Response.Write("disabled")%>>
          <% =TempClassListStr %>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;新闻类型</td>
      <td> 
        <select name="NewsType" style="width:90%" onChange="ChooseNewsType(this.options[this.selectedIndex].value);">
          <option value="RecNews" <%if Request("NewsType") = "RecNews" then Response.Write("selected")%>>推荐新闻</option>
          <option value="MarqueeNews" <%if Request("NewsType") = "MarqueeNews" then Response.Write("selected")%>>滚动新闻</option>
          <option value="SBSNews" <%if Request("NewsType") = "SBSNews" then Response.Write("selected")%>>并排新闻</option>
          <option value="PicNews" <%if Request("NewsType") = "PicNews" then Response.Write("selected")%>>图片新闻</option>
          <option value="NewNews" <%if Request("NewsType") = "NewNews" then Response.Write("selected")%>>最新新闻</option>
          <option value="HotNews" <%if Request("NewsType") = "HotNews" then Response.Write("selected")%>>热点新闻</option>
          <option value="WordNews" <%if Request("NewsType") = "WordNews" then Response.Write("selected")%>>文字新闻</option>
          <option value="TitleNews" <%if Request("NewsType") = "TitleNews" then Response.Write("selected")%>>标题新闻</option>
          <option value="ProclaimNews" <%if Request("NewsType") = "ProclaimNews" then Response.Write("selected")%>>公告新闻</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;更多链接</td>
      <td> 
        <select name="MoreContent" id="MoreContent" style="width:90% " onChange="ChooseLink(this.options[this.selectedIndex].value);" <%If Types = "System" then Response.Write("disabled")%>>
          <option value="1" <%If Request("MoreContent")=1 or Request("MoreContent")="" then Response.Write("selected")%>>是</option>
          <option value="0" <%If Request("MoreContent")=0 then Response.Write("selected")%>>否</option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;链接字样</td>
      <td> 
        <input name="LinkWord" type="text" id="LinkWord" style="width:90%" value="<%=Request("LinkWord")%>" <%If Types = "System" then Response.Write("disabled")%>></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;新闻数量</td>
      <td> 
        <input name="NewsNum" type="text" id="NewsNum" style="width:90%" value="<%=Request("NewsNum")%>"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;每行数量</td>
      <td> 
        <input name="RowNum" type="text" id="RowNum" style="width:90%" value="<%=Request("RowNum")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;链接样式</td>
      <td> 
        <input name="LinkCSS" type="text" id="LinkCSS" style="width:90%" value="<%=Request("LinkCSS")%>" <%If Types = "System" then Response.Write("disabled")%>></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;标题字数</td>
      <td> 
        <input name="TitleNum" type="text" id="TitleNum" style="width:90%" value="<%=Request("TitleNum")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;图片宽度</td>
      <td> 
        <input name="PicWidth" type="text" id="PicWidth" style="width:90%" value="<%=Request("PicWidth")%>"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;标题样式</td>
      <td> 
        <input name="TitleCSS" type="text" id="TitleCSS" style="width:90%" value="<%=Request("TitleCSS")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;图片高度</td>
      <td> 
        <input name="PicHeight" type="text" id="PicHeight" style="width:90%" value="<%=Request("PicHeight")%>"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;新闻行距</td>
      <td> 
        <input name="RowSpace" type="text" id="RowSpace" style="width:90%" value="<%=Request("RowSpace")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;滚动速度</td>
      <td> 
        <input name="MarSpeed" type="text" id="MarSpeed" style="width:90%" value="<%=Request("MarSpeed")%>"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;滚动方向</td>
      <td> 
        <select name="MarDirection" id="MarDirection" style="width:90% ">
          <option value="up" <%If Request("MarDirection")="up" then Response.Write("selected")%>>向上</option>
          <option value="down" <%If Request("MarDirection")="down" then Response.Write("selected")%>>向下</option>
          <option value="left" <%If Request("MarDirection")="left" then Response.Write("selected")%>>向左</option>
          <option value="right" <%If Request("MarDirection")="right" then Response.Write("selected")%>>向右</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;公告宽度</td>
      <td> 
        <input name="MarWidth" type="text" id="MarWidth" style="width:90%" value="<%=Request("MarWidth")%>"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;公告高度</td>
      <td> 
        <input name="MarHeight" type="text" id="MarHeight" style="width:90%" value="<%=Request("MarHeight")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;显示标题</td>
      <td> 
        <select name="ShowTitle" id="ShowTitle" style="width:90%">
          <option value="1" <%If Request("ShowTitle")=1 then Response.Write("selected")%>>是</option>
          <option value="0" <%If Request("ShowTitle")=0 then Response.Write("selected")%>>否</option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;新开窗口</td>
      <td> 
        <select name="OpenMode" id="OpenMode" style="width:90%">
          <option value="1" <%If Request("OpenMode")=1 then Response.Write("selected")%>>是</option>
          <option value="0" <%If Request("OpenMode")=0 then Response.Write("selected")%>>否</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;导航图片</td>
      <td> 
        <input name="NaviPic" type="text" id="NaviPic" style="width:52%" value="<%=Request("NaviPic")%>"> 
        <input id="PicChooseButton" type="button" name="Submit34" value="选择图片" onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.ClassJSForm.NaviPic);"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;行间图片</td>
      <td> 
        <input name="RowBetween" type="text" id="RowBetween" style="width:52%" value="<%=Request("RowBetween")%>"> 
        <input id="PicChooseButton" type="button" name="Submit34" value="选择图片" onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<% = UpFiles %>',550,300,window,document.ClassJSForm.RowBetween);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;调用日期</td>
      <td> 
        <select name="DateType" id="DateType" style="width:90%">
          <option value="0">调用日期类型</option>
          <option value="1" <%if Request("DateType") = "1" then Response.Write("selected") end if%>><%=Year(Now)&"-"&Month(Now)&"-"&Day(Now)%></option>
          <option value="2" <%if Request("DateType") = "2" then Response.Write("selected") end if%>><%=Year(Now)&"."&Month(Now)&"."&Day(Now)%></option>
          <option value="3" <%if Request("DateType") = "3" then Response.Write("selected") end if%>><%=Year(Now)&"/"&Month(Now)&"/"&Day(Now)%></option>
          <option value="4" <%if Request("DateType") = "4" then Response.Write("selected") end if%>><%=Month(Now)&"/"&Day(Now)&"/"&Year(Now)%></option>
          <option value="5" <%if Request("DateType") = "5" then Response.Write("selected") end if%>><%=Day(Now)&"/"&Month(Now)&"/"&Year(Now)%></option>
          <option value="6" <%if Request("DateType") = "6" then Response.Write("selected") end if%>><%=Month(Now)&"-"&Day(Now)&"-"&Year(Now)%></option>
          <option value="7" <%if Request("DateType") = "7" then Response.Write("selected") end if%>><%=Month(Now)&"."&Day(Now)&"."&Year(Now)%></option>
          <option value="8" <%if Request("DateType") = "8" then Response.Write("selected") end if%>><%=Month(Now)&"-"&Day(Now)%></option>
          <option value="9" <%if Request("DateType") = "9" then Response.Write("selected") end if%>><%=Month(Now)&"/"&Day(Now)%></option>
          <option value="10" <%if Request("DateType") = "10" then Response.Write("selected") end if%>><%=Month(Now)&"."&Day(Now)%></option>
          <option value="11" <%if Request("DateType") = "11" then Response.Write("selected") end if%>><%=Month(Now)&"月"&Day(Now)&"日"%></option>
          <option value="12" <%if Request("DateType") = "12" then Response.Write("selected") end if%>><%=day(Now)&"日"&Hour(Now)&"时"%></option>
          <option value="13" <%if Request("DateType") = "13" then Response.Write("selected") end if%>><%=day(Now)&"日"&Hour(Now)&"点"%></option>
          <option value="14" <%if Request("DateType") = "14" then Response.Write("selected") end if%>><%=Hour(Now)&"时"&Minute(Now)&"分"%></option>
          <option value="15" <%if Request("DateType") = "15" then Response.Write("selected") end if%>><%=Hour(Now)&":"&Minute(Now)%></option>
          <option value="16" <%if Request("DateType") = "16" then Response.Write("selected") end if%>><%=Year(Now)&"年"&Month(Now)&"月"&Day(Now)&"日"%></option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;保存路径</td>
      <td> 
        <input name="SaveFilePath" type="text" id="SaveFilePath" style="width:52%" value="<%=Request("SaveFilePath")%>"> 
        <input type="button" name="Subsadfmit" value="选择路径" onClick="OpenWindowAndSetValue('../../FunPages/SelectPathFrame.asp?CurrPath=<%="/"&TeempSysRootDir&"JS"%>',400,300,window,document.ClassJSForm.SaveFilePath);document.ClassJSForm.SaveFilePath.focus();"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;日期样式</td>
      <td> 
        <input name="DateCSS" type="text" id="DateCSS" style="width:90%" value="<%=Request("DateCSS")%>"></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;显示栏目</td>
      <td> 
        <select name="ShowClassTF" id="ShowClassTF" style="width:90%">
          <option value="0" <%If Request("ShowClassTF")=0 or Request("ShowClassTF")="" then Response.Write("selected")%>>不显示</option>
          <option value="1" <%If Request("ShowClassTF")=1 then Response.Write("selected")%>>显示</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26">&nbsp;&nbsp;&nbsp;&nbsp;调用子类</td>
      <td> 
        <select name="SmallClass" id="SmallClass" style="width:90%" <%If Types = "System" then Response.Write("disabled")%>>
          <option value="1" <%If Request("SmallClass")=1 then Response.Write("selected")%>>是</option>
          <option value="0" <%If Request("SmallClass")=0 or Request("SmallClass")="" then Response.Write("selected")%>>否</option>
        </select></td>
      <td>&nbsp;&nbsp;&nbsp;&nbsp;日期右对齐</td>
      <td> 
        <select name="RightDate" id="RightDate" style="width:90%">
          <option value="1" <%If Request("RightDate")=1 then Response.Write("selected")%>>是</option>
          <option value="0" <%If Request("RightDate")=0 or Request("RightDate")="" then Response.Write("selected")%>>否</option>
        </select></td>
    </tr>
</table>
</form>
</body>
</html>
<%
If Request.Form("action") = "add" then
	Dim ResultStr,TempFsoObj,FileNameStr
	If Replace(Replace(Replace(Request.Form("FileName"),"/",""),"\",""),"'","")="" then
		Response.Write("<script>alert(""文件名称不能为空或是有非法字符"");</script>") '文件名称不能为空或是有非法字符
		Response.End
	Else
		FileNameStr = Replace(Replace(Replace(Request.Form("FileName"),"/",""),"\",""),"'","")
	End If
	If NoCSSHackAdmin(Request.Form("FileCName"),"中文名称")="" then
		Response.Write("<script>alert(""文件中文名称不能为空"");</script>") '文件名称不能为空或是有非法字符
		Response.End
	End If
	If Request.Form("SaveFilePath")="" then
		Response.Write("<script>alert(""未指定文件保存路径"");</script>") '文件名称不能为空或是有非法字符
		Response.End
	End If
	Set TempFsoObj = Conn.Execute("Select ID from FS_SysJs where FileName='"&FileNameStr&"'")
	If Not TempFsoObj.eof then
		Response.Write("<script>alert(""文件名称已经存在"");</script>") 
		Response.End
	end if
	TempFsoObj.Close
	Set TempFsoObj = Nothing
	If isnumeric(Request.Form("NewsNum"))=false then
		Response.Write("<script>alert(""调用新闻数量必须为数字型"");</script>") '调用新闻数量必须为数字型
		Response.End
	End If
	If isnumeric(Request.Form("TitleNum"))=false then
		Response.Write("<script>alert(""标题字数必须为数字型"");</script>")
		Response.End
	End If
	If isnumeric(Request.Form("RowNum"))=false then
		Response.Write("<script>alert(""新闻每行排列数量必须为数字型"");</script>") 
		Response.End
	End If
	If isnumeric(Request.Form("RowSpace"))=false then
		Response.Write("<script>alert(""新闻行距必须为数字型"");</script>") 
		Response.End
	End IF
	If Types="Class" and Request.Form("ClassID")="" then
		Response.Write("<script>alert(""栏目ID参数传递错误"");</script>") '栏目ID参数传递错误
		Response.End
	End If
	If Request.Form("NewsType")="PicNews" or Request.Form("NewsType")="FilterNews" then
		If isnumeric(Request.Form("PicWidth"))=false or isnumeric(Request.Form("PicHeight"))=false then
			Response.Write("<script>alert(""图片规格必须为数字型"");</script>") 
			Response.End
		End If
	End If
	If Request.Form("MoreContent")=1 then
		If Request.Form("LinkWord")="" then
			Response.Write("<script>alert(""请输入链接字样"");</script>") 
			Response.End
		End If
	End If
	If Request.Form("NewsType")="MarqueeNews" or Request.Form("NewsType")="ProclaimNews" then
		If isnumeric(Request.Form("MarSpeed"))=false then
			Response.Write("<script>alert(""新闻滚动速度必须为数字型"");</script>") 
			Response.End
		End If
	End If
	'插入数据库
	Dim ClassJsAddObj,RsClassSql
	Set ClassJsAddObj = Server.CreateObject(G_FS_RS)
	RsClassSql = "Select * from FS_SysJs where 1=0"
	ClassJsAddObj.Open RsClassSql,Conn,3,3
	ClassJsAddObj.AddNew
	ClassJsAddObj("FileName") = Cstr(FileNameStr)
	ClassJsAddObj("FileCName") = Request.Form("FileCName")
	If Types = "Class" then
		ClassJsAddObj("FileType") = 1
	else
		ClassJsAddObj("FileType") = 2
	End If
	If Types="Class" then
		ClassJsAddObj("ClassID") = Cstr(Request.Form("ClassID"))
	End If
	ClassJsAddObj("NewsType") = Request.Form("NewsType")
	ClassJsAddObj("NewsNum") = Cint(Request.Form("NewsNum"))
	ClassJsAddObj("TitleNum") = Cint(Request.Form("TitleNum"))
	ClassJsAddObj("TitleCSS") = Cstr(Request.Form("TitleCSS"))
	ClassJsAddObj("RowNum") = Cint(Request.Form("RowNum"))
	If Request.Form("NaviPic")<>"" then
		ClassJsAddObj("NaviPic") = Cstr(Request.Form("NaviPic"))
	End If
	If Request.Form("RowBetween")<>"" then
		ClassJsAddObj("RowBetween") = Cstr(Request.Form("RowBetween"))
	End If
	ClassJsAddObj("FileSavePath") = Cstr(Request.Form("SaveFilePath"))
	ClassJsAddObj("RowSpace") = Cint(Request.Form("RowSpace"))
	ClassJsAddObj("DateType") = Cint(Request.Form("DateType"))
	ClassJsAddObj("DateCSS") = Cstr(Request.Form("DateCSS"))
	If Request.Form("ShowClassTF")<>0 then
		ClassJsAddObj("ClassName") = 1
	Else
		ClassJsAddObj("ClassName") = 0
	End If
	If Request.Form("SmallClass")<>0 then
		ClassJsAddObj("SonClass") = 1
	Else
		ClassJsAddObj("SonClass") = 0
	End If
	If Request.Form("RightDate")<>0 then
		ClassJsAddObj("RightDate") = 1
	Else
		ClassJsAddObj("RightDate") = 0
	End If
	If Request.Form("MoreContent")<>"" and isnull(Request.Form("MoreContent"))=false then
		ClassJsAddObj("MoreContent") = Request.Form("MoreContent")
	End if
	If Request.Form("MoreContent")<>0 then
		ClassJsAddObj("LinkWord") = Request.Form("LinkWord")
		ClassJsAddObj("LinkCSS") = Request.Form("LinkCSS")
	End If
	If Request.Form("PicWidth")<>"" and isnull(Request.Form("PicWidth"))=false then
		ClassJsAddObj("PicWidth") = Cint(Request.Form("PicWidth"))
	End If
	If Request.Form("PicHeight")<>"" and isnull(Request.Form("PicHeight"))=false then
		ClassJsAddObj("PicHeight") = Cint(Request.Form("PicHeight"))
	End If
	If Request.Form("MarSpeed")<>"" and isnull(Request.Form("MarSpeed"))=false then
		ClassJsAddObj("MarSpeed") = Cint(Request.Form("MarSpeed"))
	End If
	If Request.Form("MarDirection")<>"" and isnull(Request.Form("MarDirection"))=false then
		ClassJsAddObj("MarDirection") = Cstr(Request.Form("MarDirection"))
	End If
	If Request.Form("ShowTitle")<>"" and isnull(Request.Form("ShowTitle"))=false then
		ClassJsAddObj("ShowTitle") = Request.Form("ShowTitle")
	End If
	If Request.Form("OpenMode")<>1 then
		ClassJsAddObj("OpenMode") = 0
	Else
		ClassJsAddObj("OpenMode") = 1
	End If
	If Request.Form("MarWidth")<>"" and isnull(Request.Form("MarWidth"))=false then
		ClassJsAddObj("MarWidth") = Request.Form("MarWidth")
	End If
	If Request.Form("MarHeight")<>"" and isnull(Request.Form("MarHeight"))=false then
		ClassJsAddObj("MarHeight") = Request.Form("MarHeight")
	End If
	ClassJsAddObj("AddTime") = Now()
	ClassJsAddObj.Update
	ClassJsAddObj.Close
	Set ClassJsAddObj = Nothing
	ResultStr = CreateSysJS(FileNameStr)
	if ResultStr = true then
		Response.Redirect("ClassJsList.asp?Types=" & Types)
	else
		Response.Write("<script>alert("""&ResultStr&""");location='ClassJsList.asp?Types=" & Types & "'</script>") 
	end if
	Response.End
End If
Conn.Close
Set Conn = Nothing
%>
<script>
function ChooseLink(Link)
{
	if (Link!=1)
	{
	 document.ClassJSForm.LinkWord.disabled=true;
	 document.ClassJSForm.LinkCSS.disabled=true;
	 }
	else
	{
	 document.ClassJSForm.LinkWord.disabled=false;
	 document.ClassJSForm.LinkCSS.disabled=false;
	 }
 }

function ChooseNewsType(NewsType)
{
 if ((NewsType!='MarqueeNews')&&(NewsType!='ProclaimNews'))
  {
	 document.ClassJSForm.MarSpeed.disabled=true;
	 document.ClassJSForm.MarDirection.disabled=true;
	 document.ClassJSForm.MarWidth.disabled=true;
	 document.ClassJSForm.MarHeight.disabled=true;
   }
  else
  {
	 document.ClassJSForm.MarSpeed.disabled=false;
	 document.ClassJSForm.MarDirection.disabled=false;
	 document.ClassJSForm.MarWidth.disabled=false;
	 document.ClassJSForm.MarHeight.disabled=false;
   }
  if ((NewsType!='PicNews')&&(NewsType!='FilterNews'))
  {
	 document.ClassJSForm.PicWidth.disabled=true;
	 document.ClassJSForm.PicHeight.disabled=true;
	 document.ClassJSForm.ShowTitle.disabled=true;
   }
  else
  {
	 document.ClassJSForm.PicWidth.disabled=false;
	 document.ClassJSForm.PicHeight.disabled=false;
	 document.ClassJSForm.ShowTitle.disabled=false;
   }
 }
</script>