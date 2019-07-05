<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="Cls_Upfile.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Function.asp" -->
<!--#include file="../../Inc/ThumbnailFunction.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P990300") then Call ReturnError()

Dim FilePath,MaxFileSize,AllowFileExtStr,AutoReName,RsConfigObj,IsAddWaterMark
Dim FormName,Path
Dim UpFileObj
Dim ReturnValue
Set RsConfigObj = Conn.Execute("Select * from FS_Config")
if Not RsConfigObj.Eof then
	MaxFileSize = RsConfigObj("UpFileSize")
	AllowFileExtStr = RsConfigObj("UpFileType")
else
%>
<script language="JavaScript">
	alert('<% = "读取配置信息错误" %>');
	dialogArguments.location.reload();
	close();
</script>
<%
	Response.End
end if
Set RsConfigObj = Nothing
Set UpFileObj = New UpFileClass
UpFileObj.GetData
FilePath=Server.MapPath(UpFileObj.Form("Path")) & "\"
AutoReName = UpFileObj.Form("AutoRename")
IsAddWaterMark = UpFileObj.Form("chkAddWaterMark")
If IsAddWaterMark <> "1" Then	'生成是否要添加水印标记
	IsAddWaterMark = "0"
End if
ReturnValue = CheckUpFile(FilePath,MaxFileSize,AllowFileExtStr,AutoReName,IsAddWaterMark)
if ReturnValue <> "" then
%>
<script language="JavaScript">
	alert('<% = "以下文件上传失败，错误信息：\n" & ReturnValue %>');
	dialogArguments.location.reload();
	close();
</script>
<%
else
%>
<script language="JavaScript">
	dialogArguments.location.reload();
	close();
</script>
<%
end if
Set UpFileObj=Nothing


Function CheckUpFile(Path,FileSize,AllowExtStr,AutoReName,IsAddWaterMark)
	Dim ErrStr,NoUpFileTF,FsoObj,FileName,FileExtName,FileContent,SameFileExistTF
	NoUpFileTF = True
	ErrStr = ""
	Set FsoObj = Server.CreateObject(G_FS_FSO)
	For Each FormName in UpFileObj.File
		SameFileExistTF = False
		FileName = UpFileObj.File(FormName).FileName
		If NoIllegalStr(FileName)=False Then
			ErrStr=ErrStr&"文件：上传被禁止！\n"
		End If
		FileExtName = UpFileObj.File(FormName).FileExt
		FileContent = UpFileObj.File(FormName).FileData
		'是否存在重名文件
		if UpFileObj.File(FormName).FileSize > 1 then
			NoUpFileTF = False
			ErrStr = ""
			if UpFileObj.File(FormName).FileSize > CLng(FileSize)*1024 then
				ErrStr = ErrStr & FileName & "文件:超过了限制，最大只能上传" & FileSize & "K的文件\n"
			end if
			if AutoRename = "0" then
				If FsoObj.FileExists(Path & FileName) = True  then
					ErrStr = ErrStr & FileName & "文件:存在同名文件\n"
				else
					SameFileExistTF = True
				end if
			else
				SameFileExistTF = True
			End If
			if CheckFileType(AllowExtStr,FileExtName) = False then
				ErrStr = ErrStr & FileName & "文件:不允许上传,上传文件类型有" + AllowExtStr + "\n"
			end if
			if ErrStr = "" then
				if SameFileExistTF = True then
					SaveFile Path,FormName,AutoReName,IsAddWaterMark
				else
					SaveFile Path,FormName,"",IsAddWaterMark
				end if
			else
				CheckUpFile = CheckUpFile & ErrStr
			end if
		end if
	Next
	Set FsoObj = Nothing
	if NoUpFileTF = True then
		CheckUpFile = "没有上传文件"
	end if
End Function

Function CheckFileType(AllowExtStr,FileExtName)
	Dim i,AllowArray
	AllowArray = Split(AllowExtStr,",")
	FileExtName = LCase(FileExtName)
	CheckFileType = False
	For i = LBound(AllowArray) to UBound(AllowArray)
		if LCase(AllowArray(i)) = LCase(FileExtName) then
			CheckFileType = True
		end if
	Next
	if FileExtName="asp" or FileExtName="asa" or FileExtName="aspx" then
		CheckFileType = False
	end if
End Function
Function DealExtName(Byval UpFileExt)
		If IsEmpty(UpFileExt) Then Exit Function
		DealExtName = Lcase(UpFileExt)
		DealExtName = Replace(DealExtName,Chr(0),"")
		DealExtName = Replace(DealExtName,".","")
		DealExtName = Replace(DealExtName,"'","")
		DealExtName = Replace(DealExtName,"asp","")
		DealExtName = Replace(DealExtName,"asa","")
		DealExtName = Replace(DealExtName,"aspx","")
		DealExtName = Replace(DealExtName,"cer","")
		DealExtName = Replace(DealExtName,"cdx","")
		DealExtName = Replace(DealExtName,"htr","")
End Function

Function NoIllegalStr(Byval FileNameStr)
	Dim Str_Len,Str_Pos
	Str_Len=Len(FileNameStr)
	Str_Pos=InStr(FileNameStr,Chr(0))
	If Str_Pos=0 or Str_Pos=Str_Len then
	 	NoIllegalStr=True
	Else
	 	NoIllegalStr=False
	End If
End function

Function SaveFile(FilePath,FormNameItem,AutoNameType,IsAddWaterMark)
	Dim FileName,FileExtName,FileContent,FormName,RandomFigure
	Randomize 
	RandomFigure = CStr(Int((99999 * Rnd) + 1))
	FileName = UpFileObj.File(FormNameItem).FileName
	FileExtName = UpFileObj.File(FormNameItem).FileExt
	FileExtName=DealExtName(FileExtName)
	FileContent = UpFileObj.File(FormNameItem).FileData
	If AutoNameType = "1" Then
		FileName = FilePath & "副件" & FileName
	elseif AutoNameType = "2" Then
		FileName = FilePath & "1" & FileName 
	elseif AutoNameType = "3" Then
		FileName = FilePath & Year(Now())&Right("0"&Month(Now()),2)&Right("0"&Day(Now()),2)&Right("0"&Hour(Now()),2)&Right("0"&Minute(Now()),2)&Right("0"&Second(Now()),2)&RandomFigure&"."&FileExtName
	Else
		FileName = FilePath&FileName
	End If
	UpFileObj.File(FormNameItem).SaveToFile FileName
	If IsAddWaterMark = "1" Then   '在保存好的图片上添加水印
		AddWaterMark FileName
	End if
End Function
Set Conn = Nothing
%>