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
Dim Path,OType,OldPathName,NewPathName,PhysicalPath,FsoObj,FileObj,TemPath,sRootDir
If SysRootDir="" then
	TemPath="/" & TempletDir
	sRootDir=""
Else
	TemPath="/" & SysRootDir & "/" & TempletDir
	sRootDir="/"&SysRootDir
End If
NewPathName = Request("NewPathName")
OldPathName = Request("OldPathName")
Set FsoObj = Server.CreateObject(G_FS_FSO)
OType = Request("Type")
	if OType = "FileReName" then
		Path = Request("Path")
		if Path <> "" then
			if (NewPathName <> "") And (OldPathName <> "") then
				PhysicalPath = Server.MapPath(Path) & "\" & OldPathName
				if FsoObj.FileExists(PhysicalPath) = True then
					PhysicalPath = Server.MapPath(Path) & "\" & NewPathName
					if FsoObj.FileExists(PhysicalPath) = False then
						Set FileObj = FsoObj.GetFile(Server.MapPath(Path) & "\" & OldPathName)
						FileObj.Name = NewPathName
						Set FileObj = Nothing
					end if
				end if
			end if
		end if
	elseif OType = "FolderReName" then
		Path = Request("Path")
		if Path <> "" then
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
	set FsoObj=nothing
Response.Redirect sRootDir &"/"& AdminDir  & "templet/Newstemplet_list.asp?path=" & TemPath
%>