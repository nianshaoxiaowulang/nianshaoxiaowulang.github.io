<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
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
%>
<!--#include file="../../Inc/Session.asp" -->
<%
Dim Path,FileName,OType,FsoObj,FolderObj,SubFolderObj,FolderItem,ReturnValue,FolderStr,Temp_Path
OType = Request("Type")
FileName = Request("FileName")
Path = Request("Path") 
If OType = "Del" and FileName="" then
	FolderStr = Right(Path,Len(Path)-InstrRev(Path,"/",-1))
	Path = Left(Path,InstrRev(Path,"/",-1))
End If
Path = Server.MapPath(Path)
if OType <> "" then
	Set FsoObj = Server.CreateObject(G_FS_FSO)
	if OType = "Del" then
		if Path <> "" then
			if FileName <> "" then
				Dim DelFileArray,DelFile_i,Temp_False,Temp_FileName
				Temp_False = 0
				Temp_FileName = ""
				DelFileArray = Array("")
				DelFileArray = Split(FileName,"***")
				For DelFile_i = 0 to UBound(DelFileArray)
					FileName = Path & "\" & DelFileArray(DelFile_i)
					if FsoObj.FileExists(FileName) then
						FsoObj.DeleteFile FileName
					else
						Temp_False = Temp_False + 1
						If Temp_FileName = "" then
							Temp_FileName = DelFileArray(DelFile_i)
						Else
							Temp_FileName = Temp_FileName &"|"& DelFileArray(DelFile_i)
						End If
					end if
				Next
				If Temp_False >= 1 then
					Response.Write("This File:"&Temp_FileName&" are not deleted")
				Else
					Response.Write("File delete successfully")
				End if
			else
				Dim DelFolderArray,DelFolder_i,DelFolder_False,DelFolder_Name
				DelFolder_False = 0
				DelFolder_Name = ""
				DelFolderArray = Array("")
				DelFolderArray = Split(FolderStr,"***")
				For DelFolder_i = 0 to UBound(DelFolderArray)
					if FsoObj.FolderExists(Path&"\"&DelFolderArray(DelFolder_i))=true then
						FsoObj.DeleteFolder Path&"\"&DelFolderArray(DelFolder_i)
						'Response.Write("Folder Delete Successfully")
					else
						DelFolder_False = DelFolder_False + 1
						If DelFolder_Name = "" then
							DelFolder_Name = DelFolderArray(DelFolder_i)
						Else
							DelFolder_Name = DelFolder_Name &"|"& DelFolderArray(DelFolder_i)
						End If
						'Response.Write("No Folder")
					end if
				Next
				If DelFolder_False >= 1 then
					Response.Write("This File:"&DelFolder_Name&" are not deleted")
				Else
					Response.Write("Folder delete successfully")
				End if
			end if
		else
			Response.Write("Parameter error,try again please")
		end if
	elseif OType = "AddFolder" then
		if FsoObj.FolderExists(Path) = True then
			Response.Write("Folder already exists")
		else
			FsoObj.CreateFolder Path
			Response.Write("Add folder Successfully")
		end if
	elseif OType = "ExtendFolder" then
		ReturnValue = ""
		Set FolderObj = FsoObj.GetFolder(Path)
		Set SubFolderObj = FolderObj.SubFolders
		for Each FolderItem In SubFolderObj
			if ReturnValue = "" then
				ReturnValue = FolderItem.Name
			else
				ReturnValue = ReturnValue & "$" & FolderItem.Name
			end if
		next
		if ReturnValue <> "" then
			Response.Write(EnCodeResponseStr(ReturnValue))
		else
			Response.Write(ReturnValue)
		end if
	else
		Response.Write("Parameter error,try again please")
	end if
	Set FsoObj = Nothing
else
	Response.Write("Parameter error,try again please")
end if

Function EnCodeResponseStr(Str)
	Dim i
	for i = 1 to Len(Str)
		if EnCodeResponseStr = "" then
			EnCodeResponseStr = AscW(Mid(Str,i,1))
		else
			EnCodeResponseStr = EnCodeResponseStr & "*" & AscW(Mid(Str,i,1))
		end if
	Next
End Function
%>