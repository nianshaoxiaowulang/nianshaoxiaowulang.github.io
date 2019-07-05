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
Class InfoClass
	Private Conn  
	Private IForm  
	
	Public Property Let TForm(ByRef ExteriorForm)
		if IsObject(ExteriorForm) then
			Set IForm = ExteriorForm
		else
			Class_Terminate()
		end if
	End Property
	
	Private Sub Class_initialize() 
		Dim DBC
		Set DBC = New DataBaseClass
		Set Conn = DBC.OpenConnection()
		Set DBC = Nothing
	End Sub 
	
	Private Sub Class_Terminate()
		Set Conn=Nothing
	End Sub 
	'Public Function Begin
	Private Function ParentClassIDList(ClassID)
		Dim TempRs,TempStr
		Set TempRs = Conn.Execute("Select ParentID from FS_NewsClass where ClassID = '" & ClassID & "'")
		if Not TempRs.Eof then
			if TempRs("ParentID") <> "0" then
				ParentClassIDList =  TempRs("ParentID") & "," & ParentClassIDList
				ParentClassIDList = ParentClassIDList & ParentClassIDList(TempRs("ParentID"))
			end if
		end if
		TempRs.Close
		Set TempRs = Nothing
	End Function
	
	Private Function ChildClassIDList(ClassID)
		Dim TempRs
		Set TempRs = Conn.Execute("Select * from FS_NewsClass where ParentID='" & ClassID & "'  order by ID desc")
		do while Not TempRs.Eof
			ChildClassIDList = ChildClassIDList & "," & TempRs("ClassID")
			ChildClassIDList = ChildClassIDList & ChildClassIDList(TempRs("ClassID"))
			TempRs.MoveNext
		loop
		TempRs.Close
		Set TempRs = Nothing
	End Function

	Private Function GetVisionStr()
		GetVisionStr = "<!--Published Date:"&Now&"   Powered by www.Foosun.net,Products:Foosun Content Manage system-->" & Chr(13) & Chr(10)
	End Function
	'Public Function End
	'Class Function Begin
	Private Function CheckClass(ClassID,ParentID,ClassEName,ClassCName,ClassTemp,SaveFilePath,FileExtName)
		Dim TempSqlClass,RsTempClassObj
		CheckClass = ""
		if ClassID <> "" then
			if (Conn.Execute("Select * from FS_NewsClass where ClassID = '" & ClassID & "'").Eof) then 
				CheckClass = CheckClass & "栏目不存在，可能已经被删除"  
				Exit Function
			end if
		end if
		if ParentID <> "0" then
			if (Conn.Execute("Select * from FS_NewsClass where ClassID = '" & ParentID & "'").Eof) then 
				CheckClass = CheckClass & "父栏目不存在，可能已经被删除"  
				Exit Function
			end if
		end if
		if ClassEName = "" then
		   CheckClass = CheckClass & "栏目的英文名没有填写！"
		end if
		if ClassCName = "" then
		   CheckClass = CheckClass & "栏目的中文名没有填写！"
		end if
		if ClassTemp = "" then
		   CheckClass = CheckClass & "栏目的模板没有填写！"
		end if
		if SaveFilePath = "" then
		   CheckClass = CheckClass & "栏目的保存路径没有填写！"
		end if
		if FileExtName = "" then
		   CheckClass = CheckClass & "栏目的生成扩展名没有填写！"
		end if
		if ClassID = "" then
			TempSqlClass = "Select ClassEName from FS_NewsClass where ClassEName='"& ClassEName & "'"
		else
			TempSqlClass = "Select ClassEName from FS_NewsClass where ClassEName='"& ClassEName & "' and ClassID<>'" & ClassID & "'"
		end if
		if Not (Conn.Execute(TempSqlClass).Eof) then
		   CheckClass = CheckClass & "栏目的英文名已经存在！"
		end if
	End Function
	
	Public Function AddAndModifyClass()
		Dim RsClassObj
		Dim SqlClass,ErrStr,EditTF,TTempSaveFilePath,TeempFileExtName
		Dim ParentID,ClassEName,ClassCName,ClassTemp,ClassID,Contribution,AddTime,SaveFilePath,FileExtName,BrowPop
		Dim TempSysRootDir,ShowTF,NewsTemp,DoMain,FileTime,Orders,DownLoadTemp,ProductTemp,RedirectList
		if SysRootDir = "" then
			TempSysRootDir = ""
		else
			TempSysRootDir = "/" & SysRootDir
		end if
		AddTime = IForm("AddTime")
		ClassID = IForm("ClassID")
		ParentID = IForm("ParentID")
		ClassEName = IForm("ClassEName")
		ClassCName = IForm("ClassCName")
		ClassTemp = IForm("ClassTemp")
		NewsTemp = IForm("NewsTemp")
		DownLoadTemp = IForm("DownLoadTemp")
		ProductTemp = IForm("ProductTemp")
		Contribution = IForm("Contribution")
		ShowTF = IForm("ShowTF")
		SaveFilePath = IForm("SaveFilePath") 
		FileExtName = IForm("FileExtName")
		BrowPop = IForm("BrowPop")
		DoMain = IForm("DoMain")
		FileTime = IForm("FileTime")
		Orders = IForm("Orders")
		RedirectList = IForm("RedirectList")
		if FileTime = "" then
			FileTime = 100
		else
			if Not IsNumeric(FileTime) then
				FileTime = 100
			else
				FileTime = CInt(FileTime)
			end if
		end if
		if BrowPop <> "" then
			FileExtName = "asp"
		end if
		if ClassID = "" then
			EditTF = False
		else
			EditTF = true
		end if
		ErrStr = CheckClass(ClassID,ParentID,ClassEName,ClassCName,ClassTemp,SaveFilePath,FileExtName)
		if ErrStr <> "" then
			AddAndModifyClass= ErrStr 
			Exit Function
		end if
		Set RsClassObj = Server.CreateObject(G_FS_RS)
		dim IsAdd
		IsAdd=0
		if ClassID = "" then
			IsAdd=1
			ClassID = GetRandomID18()
			SqlClass = "select * from FS_NewsClass"
			RsClassObj.Open SqlClass,Conn,3,3
			RsClassObj.AddNew
		else
			ClassID = ClassID
			SqlClass = "select * from FS_NewsClass where ClassID='" & ClassID & "'"
			RsClassObj.Open SqlClass,Conn,3,3
			TTempSaveFilePath = RsClassObj("SaveFilePath")
			TeempFileExtName = RsClassObj("FileExtName")
		end if
		RsClassObj("ClassID") = ClassID
		RsClassObj("ClassEName") = NoCSSHackAdmin(ClassEName,"英文名称")
		RsClassObj("ClassCName") = NoCSSHackAdmin(ClassCName,"中文名称")
		RsClassObj("ParentID") = ParentID
		RsClassObj("SaveFilePath") = SaveFilePath
		RsClassObj("FileExtName") = FileExtName
		if BrowPop <> "" then
			RsClassObj("BrowPop") = CInt(BrowPop)
		else
			RsClassObj("BrowPop") = 0
		end if
		if AddTime <> "" then
			RsClassObj("AddTime") = AddTime
		else
			RsClassObj("AddTime") = Now
		end if
		RsClassObj("ChildNum") = 0
		RsClassObj("ClassTemp") = ClassTemp
		RsClassObj("NewsTemp") = NewsTemp
		RsClassObj("DownLoadTemp") = DownLoadTemp
		If ProductTemp <> "" then
			RsClassObj("ProductTemp") = ProductTemp
		End if
		RsClassObj("RedirectList") = RedirectList
		if Contribution = "1" then
			RsClassObj("Contribution") = 1
		else
			RsClassObj("Contribution") = 0
		end if
		if ShowTF = "1" then
			RsClassObj("ShowTF") = 1
		else
			RsClassObj("ShowTF") = 0
		end if
		if DoMain <> "" then
			RsClassObj("DoMain") = DoMain
		else
			RsClassObj("DoMain") = ""
		end if
		if Orders <> "" then
			if IsNumeric(Orders) then RsClassObj("Orders") = Orders
		end if
		RsClassObj("FileTime") = FileTime
		RsClassObj.UpDate
		RsClassObj.Close
		set RsClassObj=Nothing
		Conn.Execute("Update FS_NewsClass Set ChildNum = ChildNum + 1 where ClassID = '" & ParentID & "'")
		Set RsClassObj = Nothing
		If ClassID <> "" then
			Dim MyFile
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If Cstr(TTempSaveFilePath) <> Cstr(SaveFilePath) then
				If MyFile.FolderExists(Server.Mappath(TempSysRootDir&TTempSaveFilePath&"/"&ClassEName)) then
					MyFile.DeleteFolder(Server.Mappath(TempSysRootDir&TTempSaveFilePath&"/"&ClassEName))
				End if
			End If
			If Cstr(TeempFileExtName) <> Cstr(FileExtName) then
				If MyFile.FileExists(Server.Mappath(TempSysRootDir&TTempSaveFilePath&"/"&ClassEName)&"/index."&TeempFileExtName) then
					MyFile.DeleteFile(Server.Mappath(TempSysRootDir&TTempSaveFilePath&"/"&ClassEName)&"/index."&TeempFileExtName)
				End if
			End If
			Set MyFile = Nothing
		End If

		if IsAdd=1 then 
			dim TemplRs,Lsql
			set templRs= Server.CreateObject(G_FS_RS)
			'Lsql="select Poplist,groupid from admin,admingroup where admin.groupid=admingroup.id and admin.name='" & session("Name") & "'"
			Lsql="select PopList from FS_admingroup where id=(select groupid from FS_admin where name='" & session("Name") & "')"
			TemplRs.Open Lsql,Conn,3,3
			if Not TemplRs.Eof then 
				TemplRs("Poplist") =TemplRs("Poplist") & "," & ClassID
				TemplRs.UpDate
			end if	
			Set TemplRs = Nothing
		end if 
		if err.Number = 0 then
			'if EditTF = True then
				'AddAndModifyClass = "Success||" & ClassID
			'else

				AddAndModifyClass = "Success||" & ParentClassIDList(ClassID) & ClassID
			'end if
			Exit Function
		else
			AddAndModifyClass = "失败"  
			Exit Function
		end if
	End Function
	
	Function GetRootClassSaveFilePath(IDArrays)
		Dim RsCheckObj,DoMain
		Set RsCheckObj = Conn.Execute("Select ParentID,DoMain,SaveFilePath from FS_NewsClass where ClassID in ('" & Replace(IDArrays,",","','") & "')")
		do while Not RsCheckObj.Eof
			if RsCheckObj("ParentID") = "0" then
				if (Not IsNull(RsCheckObj("DoMain"))) And (RsCheckObj("DoMain") <> "") then
					GetRootClassSaveFilePath = RsCheckObj("SaveFilePath")
				else
					GetRootClassSaveFilePath = ""
				end if
				Set RsCheckObj = Nothing
				Exit Function
			end if
			RsCheckObj.MoveNext
		Loop
		Set RsCheckObj = Nothing
		GetRootClassSaveFilePath = ""
	End Function
	
	Public Function MoveClass(SourceClassID,ObjectClassID)
		Dim RsClassObj,SqlClass,SourceClassParentID,TempArray,LoopVar
		TempArray = Split(SourceClassID,"***")
		for LoopVar = LBound(TempArray) to UBound(TempArray)
			SqlClass = "Select ParentID,DoMain from FS_NewsClass where ClassID='" & TempArray(LoopVar) & "'"
			Set RsClassObj = Conn.Execute(SqlClass)
			if Not RsClassObj.Eof then
				SourceClassParentID = RsClassObj("ParentID")
				MoveClass = "Update FS_NewsClass Set ParentID = '" & ObjectClassID & "' where ClassID = '" & TempArray(LoopVar) & "'"
				Conn.Execute(MoveClass)
				MoveClass = "Update FS_NewsClass Set DoMain='' where ClassID = '" & TempArray(LoopVar) & "'"
				Conn.Execute(MoveClass)
			else
				SourceClassParentID = ""
			end if
			if SourceClassParentID <> "" then
				MoveClass = "Update FS_NewsClass Set ChildNum=ChildNum-1 where ClassID = '" & SourceClassParentID & "';"
				Conn.Execute(MoveClass)
			end if
			MoveClass = "Update FS_NewsClass Set ChildNum=ChildNum+1 where ClassID = '" & ObjectClassID & "';"
			Conn.Execute(MoveClass)
		Next
	End Function
	'Class Function End
	'News Function End
	Public Function MoveNews(SourceNewsArray,ObjectClassID)
		Dim i 
		for i=LBound(SourceNewsArray) to UBound(SourceNewsArray)
			if SourceNewsArray(i) <> "" then
				Conn.Execute("Update FS_News set ClassID='" & ObjectClassID & "' Where NewsID='" & SourceNewsArray(i) & "'")
			end if
		next
	End Function

	Public Function CopyNews(SourceNewsArray,ObjectClassID)
		Dim i,j,RsNewsObj,CopyNewsObj,SqlNews,FiledObj
		Dim NewsFileNames,RsNewsConfigObj,TempNewsID,ConfigInfo
		Set RsNewsConfigObj = Conn.Execute("Select DoMain,NewsFileName from FS_Config")
		ConfigInfo = RsNewsConfigObj("NewsFileName")
		Set RsNewsConfigObj = Nothing
		for i = LBound(SourceNewsArray) to UBound(SourceNewsArray)
			Set RsNewsObj = Conn.Execute("Select * from FS_News where NewsID='" & SourceNewsArray(i) & "'")
			SqlNews = "Select * from FS_News where 1=0"
			Set CopyNewsObj = Server.CreateObject(G_FS_RS)
			CopyNewsObj.Open SqlNews,Conn,3,3
			CopyNewsObj.AddNew
			For Each FiledObj In CopyNewsObj.Fields
				if LCase(FiledObj.name) <> "id" then
					if LCase(FiledObj.name) = "newsid" then
						TempNewsID = GetRandomID18()
						CopyNewsObj("newsid") = TempNewsID
					elseif LCase(FiledObj.name) = "classid" then
						CopyNewsObj("classid") = ObjectClassID
					elseif LCase(FiledObj.name) = "filename" then
						NewsFileNames = NewsFileName(ConfigInfo,ObjectClassID,TempNewsID)
						CopyNewsObj("FileName") = NewsFileNames
					else
						CopyNewsObj(FiledObj.name) = RsNewsObj(FiledObj.name)
					end if
				end if
			Next
			CopyNewsObj.UpDate
			CopyNewsObj.Close
		next
		Set RsNewsObj = Nothing
		Set CopyNewsObj = Nothing
	End Function
	'DownLoad Function End
	Public Function MoveDownLoad(SourceNewsArray,ObjectClassID)
		Dim i 
		for i=LBound(SourceNewsArray) to UBound(SourceNewsArray)
			if SourceNewsArray(i) <> "" then
				Conn.Execute("Update FS_DownLoad set ClassID='" & ObjectClassID & "' Where DownLoadID='" & SourceNewsArray(i) & "'")
			end if
		next
	End Function

	Public Function CopyDownLoad(SourceNewsArray,ObjectClassID)
		Dim i,j,RsNewsObj,CopyNewsObj,SqlNews,FiledObj
		Dim NewsFileNames,RsNewsConfigObj,TempNewsID,ConfigInfo
		Set RsNewsConfigObj = Conn.Execute("Select DoMain,NewsFileName from FS_Config")
		ConfigInfo = RsNewsConfigObj("NewsFileName")
		Set RsNewsConfigObj = Nothing
		for i = LBound(SourceNewsArray) to UBound(SourceNewsArray)
			Set RsNewsObj = Conn.Execute("Select * from FS_DownLoad where DownLoadID='" & SourceNewsArray(i) & "'")
			SqlNews = "Select * from FS_DownLoad where 1=0"
			Set CopyNewsObj = Server.CreateObject(G_FS_RS)
			CopyNewsObj.Open SqlNews,Conn,3,3
			CopyNewsObj.AddNew
			For Each FiledObj In CopyNewsObj.Fields
				if LCase(FiledObj.name) <> "id" then
					if LCase(FiledObj.name) = "downloadid" then
						TempNewsID = GetRandomID18()
						CopyNewsObj("downloadid") = TempNewsID
					elseif LCase(FiledObj.name) = "classid" then
						CopyNewsObj("classid") = ObjectClassID
					elseif LCase(FiledObj.name) = "filename" then
						NewsFileNames = NewsFileName(ConfigInfo,ObjectClassID,TempNewsID)
						CopyNewsObj("FileName") = NewsFileNames
					else
						CopyNewsObj(FiledObj.name) = RsNewsObj(FiledObj.name)
					end if
				end if
			Next
			CopyNewsObj.UpDate
			CopyNewsObj.Close
		next
		Set RsNewsObj = Nothing
		Set CopyNewsObj = Nothing
	End Function
	'DownLoad Function End
End Class
%>
