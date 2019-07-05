<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Refresh/Function.asp" -->
<!--#include file="../Inc/Cls_JS.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../Inc/Cls_RefreshJs.asp" -->
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System(FoosunCMS V3.1.0930)
'最新更新：2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'商业注册联系：028-85098980-601,项目开发：028-85098980-606、609,客户支持：608
'产品咨询QQ：394226379,159410,125114015
'技术支持QQ：315485710,66252421 
'项目开发QQ：415637671，655071
'程序开发：四川风讯科技发展有限公司(Foosun Inc.)
'Email:service@Foosun.cn
'MSN：skoolls@hotmail.com
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.cn  演示站点：test.cooin.com 
'网站通系列(智能快速建站系列)：www.ewebs.cn
'==============================================================================
'免费版本请在程序首页保留版权信息，并做上本站LOGO友情连接
'风讯公司保留此程序的法律追究权利
'如需进行2次开发，必须经过风讯公司书面允许。否则将追究法律责任
'==============================================================================

if Not ((JudgePopedomTF(Session("Name"),"P010300")) OR (JudgePopedomTF(Session("Name"),"P010505"))) then Call ReturnError()
Dim NewsID,ClassID,DownLoadID,Operation,DelTypeStr
NewsID = Request("NewsID")
ClassID = Request("ClassID")
DownLoadID = Request("DownLoadID")
Operation = Request("Operation")
Dim DelNewsSysRootDir
if SysRootDir = "" then
	DelNewsSysRootDir = ""
else
	DelNewsSysRootDir = "/" & SysRootDir
end if
if Operation = "DelClass" then
	if Not JudgePopedomTF(Session("Name"),"P010300") then Call ReturnError()
	DelTypeStr = "栏目"
else
	if Not JudgePopedomTF(Session("Name"),"P010505") then Call ReturnError()
	DelTypeStr = "内容"
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>删除栏目或者新闻</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body>
<table width="100%" border="0" cellspacing="8" cellpadding="0">
 <form name="DelForm" method="post" action="">
  <tr> 
    <td width="21%">
<div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="79%" colspan="2">彻底删除单击删除按钮，放入回收站单击回收站按钮。
      确定要删除吗?
      <input name="OperateType" type="hidden" id="OperateType">
      <input name="Result" type="hidden" id="Result" value="Submit">
      <input name="NewsID" type="hidden" id="NewsID" value="<% = NewsID %>">
      <input name="DownLoadID" type="hidden" id="DownLoadID" value="<% = DownLoadID %>">
      <input name="ClassID" type="hidden" id="ClassID" value="<% = ClassID %>"></td>
    </tr>
  <tr> 
    <td colspan="3">
<div align="center">
          <input onClick="document.DelForm.OperateType.value='Del';" name="Submitsadf" type="submit" id="Submitsadf" value=" 删 除 ">
          <input onClick="document.DelForm.OperateType.value='Recycle';" type="submit" name="Submit2" value=" 回收站 ">
          <input type="button" onClick="window.close();" name="Submit3" value=" 取 消 ">
      </div></td>
    </tr>
 </form>
</table>
</body>
</html>
<%
Dim Result,MyFile
Set MyFile=Server.CreateObject(G_FS_FSO)
Result = Request.Form("Result")
if Result = "Submit" then
	Dim OperateType
	OperateType = Request.Form("OperateType")
	if ClassID <> "" then
		if JudgePopedomTF(Session("Name"),"P010300") then DelClass ClassID,OperateType
	end if
	if NewsID <> "" then
		if JudgePopedomTF(Session("Name"),"P010505") then DelNews NewsID,OperateType
	end if
	if DownLoadID <> "" then
		if JudgePopedomTF(Session("Name"),"P010704") then DelDownLoad DownLoadID,OperateType
	end if
	Response.Write("<script>window.close();</script>")
end if
Function DelClass(DelClassID,OpType)
	Dim DelClassIDArray,D_i,JSClassObj
	DelClassIDArray = Array("")
	DelClassIDArray = Split(DelClassID,"***")
	Set JSClassObj = New JSClass
	JSClassObj.SysRootDir = DelNewsSysRootDir
	For D_i = 0 to UBound(DelClassIDArray)
		if JudgePopedomTF(Session("Name"),""&DelClassIDArray(D_i)&"") then 
			Dim AllClassID,Sql
			AllClassID = "'" & DelClassIDArray(D_i) & "'" & ChildClassIDList(DelClassIDArray(D_i))
			On Error Resume Next
			if OpType = "Del" then
				Sql = "Delete from FS_News where ClassID in (" & AllClassID & ")"
				Conn.Execute(Sql)
				if Err.Number <> 0 then Alert "删除栏目下的新闻失败"
				Sql = "Delete from FS_Contribution where ClassID in (" & AllClassID & ")"
				Conn.Execute(Sql)
				if Err.Number <> 0 then Alert "删除栏目下的投稿失败"
				Sql = "Delete from FS_DownLoad where ClassID in (" & AllClassID & ")"
				Conn.Execute(Sql)
				if Err.Number <> 0 then Alert "删除栏目下的下载失败"
				'----------删除自由js中的相关记录及重新生成相关JS文件(FreeJsFile)----------
				Dim RsDelFreeJsObj,TempClassIDStr,FreeJsArr,FreeJsObj,Free_i
				TempClassIDStr = ""
				Set RsDelFreeJsObj = Conn.Execute("Select distinct JSName from FS_FreeJsFile where ClassID in (" & AllClassID & ") ")
				do while Not RsDelFreeJsObj.eof
					If TempClassIDStr = "" then
						TempClassIDStr = RsDelFreeJsObj("JSName")
					Else
						TempClassIDStr = TempClassIDStr &","&RsDelFreeJsObj("JSName")
					End If
					RsDelFreeJsObj.MoveNext
				Loop
				RsDelFreeJsObj.Close
				Set RsDelFreeJsObj = Nothing
				Conn.Execute("Delete from FS_FreeJsFile where ClassID in (" & AllClassID & ")")
				FreeJsArr = Array("")
				FreeJsArr = Split(TempClassIDStr,",")
				For Free_i=0 to UBound(FreeJsArr)
					Set FreeJsObj = Conn.Execute("Select EName,Manner from FS_FreeJS where EName='"&FreeJsArr(Free_i)&"'")
				  Select case FreeJsObj("Manner")
					 case "1"   JSClassObj.WCssA FreeJsObj("EName"),True
					 case "2"   JSClassObj.WCssB FreeJsObj("EName"),True
					 case "3"   JSClassObj.WCssC FreeJsObj("EName"),True
					 case "4"   JSClassObj.WCssD FreeJsObj("EName"),True
					 case "5"   JSClassObj.WCssE FreeJsObj("EName"),True
					 case "6"   JSClassObj.PCssA FreeJsObj("EName"),True
					 case "7"   JSClassObj.PCssB FreeJsObj("EName"),True
					 case "8"   JSClassObj.PCssC FreeJsObj("EName"),True
					 case "9"   JSClassObj.PCssD FreeJsObj("EName"),True
					 case "10"   JSClassObj.PCssE FreeJsObj("EName"),True
					 case "11"   JSClassObj.PCssF FreeJsObj("EName"),True
					 case "12"   JSClassObj.PCssG FreeJsObj("EName"),True
					 case "13"   JSClassObj.PCssH FreeJsObj("EName"),True
					 case "14"   JSClassObj.PCssI FreeJsObj("EName"),True
					 case "15"   JSClassObj.PCssJ FreeJsObj("EName"),True
					 case "16"   JSClassObj.PCssK FreeJsObj("EName"),True
					 case "17"   JSClassObj.PCssL FreeJsObj("EName"),True
				   End Select
				   FreeJsObj.Close
				   Set FreeJsObj = Nothing
				Next
				if Err.Number <> 0 then Alert "删除栏目下的自由JS新闻失败"
				'---------删除栏目时删除系统JS中的相关记录及文件(SysJs)---------------------
				Dim RsSysJsObj
				Set RsSysJsObj = Conn.Execute("Select FileName,FileSavePath from FS_SysJs where ClassID in ("&AllClassID&")")
				do while Not RsSysJsObj.eof
					If MyFile.FileExists(Server.Mappath(DelNewsSysRootDir&RsSysJsObj("FileSavePath"))&"/"&RsSysJsObj("FileName")&".js") then
						MyFile.DeleteFile(Server.Mappath(DelNewsSysRootDir&RsSysJsObj("FileSavePath"))&"/"&RsSysJsObj("FileName")&".js")
					End if
					RsSysJsObj.MoveNext
				loop
				RsSysJsObj.Close
				Set RsSysJsObj = Nothing
				Conn.Execute("Delete from FS_SysJs where ClassID in ("&AllClassID&")")
				if Err.Number <> 0 then Alert "删除栏目的系统JS信息失败"
				'---------------------物理文件删除-------------------------------------
				Dim DelClassFileObj
				Set DelClassFileObj = Conn.Execute("Select ClassEName,SaveFilePath from FS_NewsClass where ClassID in ("&AllClassID&")")
				'修改 栏目保存路径为/，删除时物理文件没有删除的情况
				Do while Not DelClassFileObj.eof
					If DelClassFileObj("SaveFilePath")<>"/" then
						If MyFile.FolderExists(Server.Mappath(DelNewsSysRootDir&DelClassFileObj("SaveFilePath")&"/"&DelClassFileObj("ClassEName"))) then
							MyFile.DeleteFolder(Server.Mappath(DelNewsSysRootDir&DelClassFileObj("SaveFilePath")&"/"&DelClassFileObj("ClassEName")))
						End if
					Else
						If MyFile.FolderExists(Server.Mappath(DelNewsSysRootDir&"/"&DelClassFileObj("ClassEName"))) then
							MyFile.DeleteFolder(Server.Mappath(DelNewsSysRootDir&"/"&DelClassFileObj("ClassEName")))
						End if
					End If
					DelClassFileObj.MoveNext
				Loop
				DelClassFileObj.Close
				Set DelClassFileObj = Nothing
				'----------------------------------------------------------------------
				Sql = "Delete from FS_NewsClass where ClassID in (" & AllClassID & ")"
				Conn.Execute(Sql)
				if Err.Number = 0 then
					Alert ""
				else
					Alert "删除失败"
				end if
			else
				'----------将自由js中的相关记录放入回收站及重新生成相关JS文件(FreeJsFile)----------
				TempClassIDStr = ""
				Set RsDelFreeJsObj = Conn.Execute("Select distinct JSName from FS_FreeJsFile where ClassID in (" & AllClassID & ") ")
				do while Not RsDelFreeJsObj.eof
					If TempClassIDStr = "" then
						TempClassIDStr = RsDelFreeJsObj("JSName")
					Else
						TempClassIDStr = TempClassIDStr &","&RsDelFreeJsObj("JSName")
					End If
					RsDelFreeJsObj.MoveNext
				Loop
				RsDelFreeJsObj.Close
				Set RsDelFreeJsObj = Nothing
				Conn.Execute("Update FS_FreeJsFile Set DelFlag=1 where ClassID in (" & AllClassID & ")")
				FreeJsArr = Array("")
				FreeJsArr = split(TempClassIDStr,",")
				For Free_i=0 to UBound(FreeJsArr)
					Set FreeJsObj = Conn.Execute("Select EName,Manner from FS_FreeJS where EName='"&FreeJsArr(Free_i)&"'")
				  Select case FreeJsObj("Manner")
					 case "1"   JSClassObj.WCssA FreeJsObj("EName"),True
					 case "2"   JSClassObj.WCssB FreeJsObj("EName"),True
					 case "3"   JSClassObj.WCssC FreeJsObj("EName"),True
					 case "4"   JSClassObj.WCssD FreeJsObj("EName"),True
					 case "5"   JSClassObj.WCssE FreeJsObj("EName"),True
					 case "6"   JSClassObj.PCssA FreeJsObj("EName"),True
					 case "7"   JSClassObj.PCssB FreeJsObj("EName"),True
					 case "8"   JSClassObj.PCssC FreeJsObj("EName"),True
					 case "9"   JSClassObj.PCssD FreeJsObj("EName"),True
					 case "10"   JSClassObj.PCssE FreeJsObj("EName"),True
					 case "11"   JSClassObj.PCssF FreeJsObj("EName"),True
					 case "12"   JSClassObj.PCssG FreeJsObj("EName"),True
					 case "13"   JSClassObj.PCssH FreeJsObj("EName"),True
					 case "14"   JSClassObj.PCssI FreeJsObj("EName"),True
					 case "15"   JSClassObj.PCssJ FreeJsObj("EName"),True
					 case "16"   JSClassObj.PCssK FreeJsObj("EName"),True
					 case "17"   JSClassObj.PCssL FreeJsObj("EName"),True
				   End Select
				   FreeJsObj.Close
				   Set FreeJsObj = Nothing
				Next
				if Err.Number <> 0 then Alert "删除栏目下的自由JS新闻失败"
				'---------删除栏目时删除系统JS中的相关记录及文件(SysJs)--------
				Set RsSysJsObj = Conn.Execute("Select FileName,FileSavePath from FS_SysJs where ClassID in ("&AllClassID&")")
				do while Not RsSysJsObj.eof
					If MyFile.FileExists(Server.Mappath(DelNewsSysRootDir&RsSysJsObj("FileSavePath"))&"/"&RsSysJsObj("FileName")&".js") then
						MyFile.DeleteFile(Server.Mappath(DelNewsSysRootDir&RsSysJsObj("FileSavePath"))&"/"&RsSysJsObj("FileName")&".js")
					End if
					RsSysJsObj.MoveNext
				loop
				RsSysJsObj.Close
				Set RsSysJsObj = Nothing
				Conn.Execute("Delete from FS_SysJs where ClassID in ("&AllClassID&")")
				if Err.Number <> 0 then Alert "删除栏目的系统JS信息失败"
				'-----------------------------------------------------------------
				Sql = "UpDate FS_News Set DelTF=1,DelTime='"&Now()&"' where ClassID in (" & AllClassID & ")"
				Conn.Execute(Sql)
				if Err.Number <> 0 then Alert "栏目下的新闻放入回收站失败"
				Sql = "UpDate FS_NewsClass Set DelFlag=1,DelTime='"&Now()&"' where ClassID in (" & AllClassID & ")"
				Conn.Execute(Sql)
				if Err.Number = 0 then
					Alert ""
				else
					Alert "放入回收站失败"
				end if
			end if
		End If
	Next
	Set JSClassObj = Nothing
End Function
Function DelNews(DelNewsID,OpType)
	Dim Sql,RikerClassIDObj,TempRikerID,RikerFileName,RikerFreeJsFileObj,RikerCreaFreeJsEName,RikerCreaFreeJsManner,JSClassObj
	'On Error Resume Next
	Dim DelNewsIDArray,DN_i
	DelNewsIDArray = Array("")
	DelNewsIDArray = Split(DelNewsID,"***")
	Set JSClassObj = New JSClass
	JSClassObj.SysRootDir = DelNewsSysRootDir

	For DN_i = 0 to UBound(DelNewsIDArray)
		if OpType = "Del" then
			Sql = "Delete from FS_News where NewsID='" & DelNewsIDArray(DN_i) & "'"
			'------------------------删除新闻物理文件-------------------
			Dim DelNewsClassFileObj,DelNewsFileObj,TempDelPath

			Set DelNewsFileObj = Conn.Execute("Select Path,FileName,FileExtName,ClassID from FS_News where NewsID='"&DelNewsIDArray(DN_i)&"'")
			If Not DelNewsFileObj.eof then
				Set DelNewsClassFileObj = Conn.execute("Select ClassEName,SaveFilePath from FS_NewsClass where ClassID='"&DelNewsFileObj("ClassID")&"'")
				If Not DelNewsFileObj.eof then
				'///////////////////////l
					Dim TempSaveFilePath
					If DelNewsClassFileObj("SaveFilePath")="/" then
						TempSaveFilePath=""
					Else
						TempSaveFilePath=DelNewsClassFileObj("SaveFilePath")
					End If
					if Application("UseDatePath")="1" then
						TempDelPath=DelNewsSysRootDir&TempSaveFilePath&"/"&DelNewsClassFileObj("ClassEName") & DelNewsFileObj("Path") & "/"&DelNewsFileObj("FileName")&"."&DelNewsFileObj("FileExtName")
					else
						TempDelPath=DelNewsSysRootDir&TempSaveFilePath&"/"&DelNewsClassFileObj("ClassEName")&"/"&DelNewsFileObj("FileName")&"."&DelNewsFileObj("FileExtName")
					end if
					If MyFile.FileExists(Server.Mappath(TempDelPath)) then
						MyFile.DeleteFile(Server.Mappath(TempDelPath))
					End if
				'///////////////
				End If
			End If
			'------------自由JS相关删除及重新生成----------------
			Set RikerClassIDObj = Conn.Execute("Select ClassID,FileName from FS_News where NewsID='"&DelNewsIDArray(DN_i)&"'")
			If Not RikerClassIDObj.eof then
				TempRikerID = RikerClassIDObj("ClassID")
				RikerFileName = RikerClassIDObj("FileName")
			End If
			RikerClassIDObj.Close
			Set RikerClassIDObj = Nothing
			
			Conn.Execute(Sql) '删除新闻

			Set RikerFreeJsFileObj = Conn.Execute("Select EName,Manner from FS_FreeJS where EName in (Select JSName from FS_FreeJsFile where FileName='"&RikerFileName&"')")
			RikerCreaFreeJsEName = ""
			RikerCreaFreeJsManner = ""
			do while Not RikerFreeJsFileObj.eof
				If RikerCreaFreeJsEName = "" then
					RikerCreaFreeJsEName = RikerFreeJsFileObj("EName")
					RikerCreaFreeJsManner = RikerFreeJsFileObj("Manner")
				Else
					RikerCreaFreeJsEName = RikerCreaFreeJsEName &","& RikerFreeJsFileObj("EName")
					RikerCreaFreeJsManner = RikerCreaFreeJsManner &","& RikerFreeJsFileObj("Manner")
				End If
					RikerFreeJsFileObj.MoveNext
			loop
			RikerFreeJsFileObj.Close
			Set RikerFreeJsFileObj = Nothing
			Conn.execute("Delete from FS_FreeJsFile where FileName='"&RikerFileName&"'")
			Dim RikerENameArr,RikerMannerArr,Riker_i
			RikerENameArr = Array("")
			RikerMannerArr = Array("")
			RikerENameArr = split(RikerCreaFreeJsEName,",")
			RikerMannerArr = split(RikerCreaFreeJsManner,",")
			For Riker_i=0 to UBound(RikerMannerArr)
				Select case RikerMannerArr(Riker_i)
					case "1"   JSClassObj.WCssA RikerENameArr(Riker_i),True
					case "2"   JSClassObj.WCssB RikerENameArr(Riker_i),True
					case "3"   JSClassObj.WCssC RikerENameArr(Riker_i),True
					case "4"   JSClassObj.WCssD RikerENameArr(Riker_i),True
					case "5"   JSClassObj.WCssE RikerENameArr(Riker_i),True
					case "6"   JSClassObj.PCssA RikerENameArr(Riker_i),True
					case "7"   JSClassObj.PCssB RikerENameArr(Riker_i),True
					case "8"  JSClassObj.PCssC RikerENameArr(Riker_i),True
					case "9"   JSClassObj.PCssD RikerENameArr(Riker_i),True
					case "10"   JSClassObj.PCssE RikerENameArr(Riker_i),True
					case "11"   JSClassObj.PCssF RikerENameArr(Riker_i),True
					case "12"   JSClassObj.PCssG RikerENameArr(Riker_i),True
					case "13"   JSClassObj.PCssH RikerENameArr(Riker_i),True
					case "14"   JSClassObj.PCssI RikerENameArr(Riker_i),True
					case "15"   JSClassObj.PCssJ RikerENameArr(Riker_i),True
					case "16"   JSClassObj.PCssK RikerENameArr(Riker_i),True
					case "17"   JSClassObj.PCssL RikerENameArr(Riker_i),True
				End Select
			Next
			'-----------------------------------------------------------
		else
			Sql = "Update FS_News Set DelTF=1,DelTime='"&Now()&"' where NewsID='" & DelNewsIDArray(DN_i) & "'"
			Conn.Execute(Sql)
			'------------自由JS相关删除及重新生成----------------
			Set RikerClassIDObj = Conn.Execute("Select ClassID,FileName from FS_News where NewsID='"&DelNewsIDArray(DN_i)&"'")
			If Not RikerClassIDObj.eof then
				TempRikerID = RikerClassIDObj("ClassID")
				RikerFileName = RikerClassIDObj("FileName")
				Conn.execute("Update FS_FreeJsFile set DelFlag=1 where FileName='"&RikerFileName&"'")
				'------------------重新生成相关系统栏目JS------------
				Dim RikerSysObj
				Set RikerSysObj = Conn.Execute("Select FileName from FS_SysJs where ClassID='"&TempRikerID&"'")
				If Not RikerSysObj.eof then
					CreateSysJS RikerSysObj("FileName")
				End If
				'------------------系统总JS相关------------
				'考虑中......
				'------------------自由JS------------------
				Set RikerFreeJsFileObj = Conn.Execute("Select EName,Manner from FS_FreeJS where EName in (Select JSName from FS_FreeJsFile where FileName='"&RikerFileName&"')")
				Do while Not RikerFreeJsFileObj.eof
					RikerCreaFreeJsEName = RikerFreeJsFileObj("EName")
					RikerCreaFreeJsManner = RikerFreeJsFileObj("Manner")
					Select case RikerCreaFreeJsManner
						case "1"   JSClassObj.WCssA RikerCreaFreeJsEName,True
						case "2"   JSClassObj.WCssB RikerCreaFreeJsEName,True
						case "3"   JSClassObj.WCssC RikerCreaFreeJsEName,True
						case "4"   JSClassObj.WCssD RikerCreaFreeJsEName,True
						case "5"   JSClassObj.WCssE RikerCreaFreeJsEName,True
						case "6"   JSClassObj.PCssA RikerCreaFreeJsEName,True
						case "7"   JSClassObj.PCssB RikerCreaFreeJsEName,True
						case "8"   JSClassObj.PCssC RikerCreaFreeJsEName,True
						case "9"   JSClassObj.PCssD RikerCreaFreeJsEName,True
						case "10"   JSClassObj.PCssE RikerCreaFreeJsEName,True
						case "11"   JSClassObj.PCssF RikerCreaFreeJsEName,True
						case "12"   JSClassObj.PCssG RikerCreaFreeJsEName,True
						case "13"   JSClassObj.PCssH RikerCreaFreeJsEName,True
						case "14"   JSClassObj.PCssI RikerCreaFreeJsEName,True
						case "15"   JSClassObj.PCssJ RikerCreaFreeJsEName,True
						case "16"   JSClassObj.PCssK RikerCreaFreeJsEName,True
						case "17"   JSClassObj.PCssL RikerCreaFreeJsEName,True
					End Select
					RikerFreeJsFileObj.MoveNext
				Loop
				RikerFreeJsFileObj.Close
				Set RikerFreeJsFileObj = Nothing
			End If
			RikerClassIDObj.Close
			Set RikerClassIDObj = Nothing
		end if
	Next
	Set JSClassObj = Nothing
	'------------------------------------------
	if Err.Number = 0 then
		Response.Write("<script>window.close();</script>")
		Response.End
	else
		Response.Write("<script>alert(""删除发生错误"");window.close();</script>")
		Response.End
	end if
End Function
Function DelDownLoad(DelDownLoadID,OpType)
  	Dim DelDownloadObj,DelDownClassObj,DDArray,DD_i
	DDArray = Array("")
	DDArray = Split(DelDownLoadID,"***")
	For DD_i = 0 to UBound(DDArray)
		Set DelDownloadObj = Conn.Execute("Select ClassID,FileName,FileExtName from FS_DownLoad where DownLoadID='"&DDArray(DD_i)&"'")
		If Not DelDownloadObj.eof then
			Set DelDownClassObj = Conn.Execute("Select ClassEName,SaveFilePath from FS_NewsClass where ClassID='"&DelDownloadObj("ClassID")&"'")
			If Not DelDownClassObj.eof then
				Dim TempSaveFilePath
				If DelDownClassObj("SaveFilePath")="/" then
					TempSaveFilePath=""
				Else
					TempSaveFilePath=DelDownClassObj("SaveFilePath")
				End If
				if MyFile.FileExists(Server.MapPath(DelNewsSysRootDir & TempSaveFilePath & "/"& DelDownClassObj("ClassEName")) & "/" & DelDownloadObj("FileName") & "." & DelDownloadObj("FileExtName")) then
					MyFile.DeleteFile (Server.MapPath(DelNewsSysRootDir & TempSaveFilePath & "/"& DelDownClassObj("ClassEName")) & "/" & DelDownloadObj("FileName") & "." & DelDownloadObj("FileExtName"))
				end if 
			End If
			DelDownClassObj.Close
			Set DelDownClassObj = Nothing
			Conn.Execute("Delete from FS_DownLoad where DownLoadID='" & DDArray(DD_i) & "'")
			Conn.Execute("Delete from FS_DownLoadAddress where DownLoadID='" & DDArray(DD_i) & "'")
		End If
		DelDownloadObj.Close
		Set DelDownloadObj = Nothing
	Next
End Function
Function ChildClassIDList(ClassID)
	Dim TempRs,TempStr
	Set TempRs = Conn.Execute("Select ClassID from FS_NewsClass where ParentID = '" & ClassID & "'")
	do while Not TempRs.Eof
		ChildClassIDList = ChildClassIDList & ",'" & TempRs("ClassID") & "'"
		ChildClassIDList = ChildClassIDList & ChildClassIDList(TempRs("ClassID"))
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function
Function Alert(InfoStr)
	%>
	<script language="JavaScript">
	<% if InfoStr <> "" then %>
		alert('<% = InfoStr %>');
	<% end if %>
	var LocationStr=dialogArguments.location.href;
	<% if Operation = "DelClass" then %>
		LocationStr=AddLocationStr(LocationStr,'<% = ParentClassIDList(ClassID) %>','OpenClassIDList');
		dialogArguments.location=LocationStr;
	<% else %>
		dialogArguments.location.reload();
	<% end if %>
	window.close();
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
	</script>
	<%
End Function
Function ParentClassIDList(ClassID)
	Dim TempRs,TempStr
	Set TempRs = Conn.Execute("Select ParentID from FS_NewsClass where ClassID = '" & ClassID & "'")
	Exit Function
	if Not TempRs.Eof then
		if TempRs("ParentID") <> "0" then
			ParentClassIDList =  TempRs("ParentID") & "," & ParentClassIDList
			ParentClassIDList = ParentClassIDList & ParentClassIDList(TempRs("ParentID"))
		end if
	end if
	TempRs.Close
	Set TempRs = Nothing
End Function
Set MyFile=nothing
Set Conn = Nothing
%>
