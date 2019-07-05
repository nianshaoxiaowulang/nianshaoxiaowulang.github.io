<%
Function GetPageContent(Url) 
	Dim HTTPObj
	On Error Resume Next
	Set HTTPObj = Server.CreateObject(TempHTTPObj) 
	With HTTPObj 
		.Open "Get", Url, False, "", "" 
		.Send 
	End With 
	if HTTPObj.Readystate <> 4 then
		Set HTTPObj = Nothing
		GetPageContent = False
		Exit Function
	end if
	GetPageContent = ResponseStrToStr(HTTPObj.ResponseBody)
	Set HTTPObj = Nothing
End Function

Function ResponseStrToStr(BodyStr)
	Dim ADOStreamObj
	Set ADOStreamObj = Server.CreateObject("Adodb.Stream")
	ADOStreamObj.Type = 1
	ADOStreamObj.Mode = 3
	ADOStreamObj.Open
	ADOStreamObj.Write BodyStr
	ADOStreamObj.Position = 0
	ADOStreamObj.Type = 2
	ADOStreamObj.Charset = "GB2312"
	ResponseStrToStr = ADOStreamObj.ReadText 
	ADOStreamObj.Close
	Set ADOStreamObj = Nothing
End Function

Function GetContent(Str,StartStr,LastStr,Flag)
	On Error Resume next
	if Instr(LCase(Str),LCase(StartStr)) > 0 then
		Dim regEx,SearchStr,Matches,Matche
		Str = Replace(Replace(Str,Chr(13),""),Chr(10),"")
		StartStr = Replace(StartStr,"[变量]",".*")
		LastStr = Replace(LastStr,"[变量]",".*")
		SearchStr = StartStr & ".*" & LastStr
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Pattern = SearchStr
		Set Matches = regEx.Execute(str)
		set Matche = Matches(0)
		Select Case Flag
			Case 0 '不包括首尾特征字符
				GetContent = Matche
				regEx.Pattern = StartStr
				GetContent = regEx.Replace(GetContent,"")
				regEx.Pattern = LastStr & ".*|\n"
				GetContent = regEx.Replace(GetContent,"")	
			Case 1 '包括首尾特征字符
				GetContent = Matche
			Case 2 '取开始字符后面的所有内容
				GetContent = Matche
				regEx.Pattern = StartStr
				GetContent = regEx.Replace(GetContent,"")
			Case else
				GetContent = ""
		End Select
	else
		GetContent = ""
	end if
	if Err then 
		Err.clear
		GetContent = ""
	End If
End Function

Function GetOtherContent(Str,StartStr,LastStr)
	On Error Resume Next
	Dim regEx,SearchStr,Matches,Matche
	Str = Replace(Replace(Str,Chr(13),""),Chr(10),"")
	StartStr = Replace(Replace(Replace(StartStr,"[变量]","(.*)"),Chr(13),""),Chr(10),"")
	LastStr = Replace(Replace(Replace(LastStr,"[变量]","(.*)"),Chr(13),""),Chr(10),"")
	SearchStr = StartStr & ".*" & LastStr
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Pattern = SearchStr
	Set Matches = regEx.Execute(str)
	For Each Matche In  Matches
		If Matche<>"" Then 
			GetOtherContent = Matche
			regEx.Pattern = StartStr
			GetOtherContent = regEx.Replace(GetOtherContent,"")
			regEx.Pattern = LastStr & ".*|\n"
			GetOtherContent = regEx.Replace(GetOtherContent,"")
		Else
			GetOtherContent = ""
		End If 
		If Err Then 
			Err.clear
			GetOtherContent = "" 
		End If
		Exit For
	Next
End Function

Function FormatUrl(NewsLinkStr,ObjURL)
	'///////
	'测试值
	'NewsLinkStr = "../aaa.htm"
	'CollectObjURL = "http://www.baidu.com/bbb/ccc/"
	'SiteUrl = "http://www.baidu.com"
	'/////
	Dim URLSearchLoc
	'NewsLinkStr = LCase(NewsLinkStr)
	if Left(NewsLinkStr,7) <> "http://" then
		Dim CheckURLStr,TempCollectObjURL,CheckObjURL
		NewsLinkStr = Replace(Replace(Replace(NewsLinkStr,"'",""),"""","")," ","")
		TempCollectObjURL = Left(ObjURL,InStrRev(ObjURL,"/"))
		CheckObjURL = NewsLinkStr
		CheckURLStr = Left(NewsLinkStr,3)
		if Left(NewsLinkStr,1) = "/" then
			URLSearchLoc = InStr(ObjURL,"//") + 2
			FormatUrl = Left(ObjURL,InStr(URLSearchLoc,ObjURL,"/") - 1)
			FormatUrl = FormatUrl & NewsLinkStr
		elseif CheckURLStr = "../" then
			do while Not CheckURLStr <> "../"
				CheckObjURL = Mid(CheckObjURL,4)
				if Right(TempCollectObjURL,1) = "/" then TempCollectObjURL = Left(TempCollectObjURL,Len(TempCollectObjURL) - 1)
				TempCollectObjURL = Left(TempCollectObjURL,InStrRev(TempCollectObjURL,"/"))
				CheckURLStr = Left(CheckObjURL,3)
			Loop
			FormatUrl = TempCollectObjURL & CheckObjURL
		else
			FormatUrl = TempCollectObjURL & NewsLinkStr
		end if
	else
		FormatUrl = NewsLinkStr
	end If
End Function


Function ReplaceIMGRemoteUrl(NewsContent,SaveFilePath,FunDoMain,DummyPath,NewsLinkStr,SaveRemotePic)  'ReplaceRemoteUrl变形
	Dim re,RemoteFile,RemoteFileurl,SaveFileName,FileName,FileExtName,SaveImagePath,ReplaceFileUrl,TempFileUrl
	Dim SaveIMGFileName,SourceFileUrl
	Set re = New RegExp
	re.IgnoreCase = True
	re.Global=True
	're.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}((\w)+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(gif|jpg|png|bmp|swf)))"
	re.Pattern = "(src\S+\.{1}(gif|jpg|png|bmp|swf)(""|\')?)"
	Set RemoteFile = re.Execute(NewsContent)
	Set re = Nothing
	For Each RemoteFileurl in RemoteFile
		ReplaceFileUrl = Replace(Replace(Replace(RemoteFileurl,"=",""),"'",""),"""","")
		SourceFileUrl = RemoteFileurl
		TempFileUrl = mid(ReplaceFileUrl,4)
		RemoteFileurl = FormatUrl(TempFileUrl,NewsLinkStr)
		If SaveRemotePic Then			
			SaveFileName = Mid(RemoteFileurl,InstrRev(RemoteFileurl,"/")+1)
			FileExtName = Mid(SaveFileName,InstrRev(SaveFileName,".")+1)
			SaveIMGFileName = GetRandomID18 & "." & FileExtName
			Call SaveRemoteFile(DummyPath & SaveFilePath & "/" & SaveIMGFileName,RemoteFileurl)
			NewsContent = Replace(NewsContent,SourceFileUrl, "src=""" & FunDoMain & SaveFilePath & "/" & SaveIMGFileName & """")
		Else
			NewsContent = Replace(NewsContent,SourceFileUrl, "src=""" & RemoteFileurl &"""")
			'不选择远程存图也替换图片地址为绝对地址2005.10.20
		End If		
	Next
	ReplaceIMGRemoteUrl = NewsContent
End Function

Function ReplaceContentStr(ContentStr)
	Dim TempContentStr
	TempContentStr = ContentStr
	if TextTF then
		TempContentStr = LoseHtml(TempContentStr)
	else
		if IsStyle = True then TempContentStr = LoseStyleTag(TempContentStr)
		if IsDiv = True then TempContentStr = LoseDivTag(TempContentStr)
		if IsA = True then TempContentStr = LoseATag(TempContentStr)
		if IsFont = True then TempContentStr = LoseFontTag(TempContentStr)
		if IsSpan = True then TempContentStr = LoseSpanTag(TempContentStr)
		if IsObjectTF = True then TempContentStr = LoseObjectTag(TempContentStr)
		if IsIFrame = True then TempContentStr = LoseIFrameTag(TempContentStr)
		if IsScript = True then TempContentStr = LoseScriptTag(TempContentStr)
		if IsClass = True then TempContentStr = LoseClassTag(TempContentStr)
	end if
	ReplaceContentStr = TempContentStr
End Function

Function LoseClassTag(ContentStr)
	Dim ClsTempLoseStr,regEx
	ClsTempLoseStr = Cstr(ContentStr)
	Set regEx = New RegExp
	regEx.Pattern = "(class=){1,}(""|\'){0,1}\S+(""|\'|>|\s){0,1}"
	regEx.IgnoreCase = True
	regEx.Global = True
	ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
	LoseClassTag = ClsTempLoseStr
	Set regEx = Nothing
End Function

Function LoseScriptTag(ContentStr)
	Dim ClsTempLoseStr,regEx
	ClsTempLoseStr = Cstr(ContentStr)
	Set regEx = New RegExp
	regEx.Pattern = "(<script){1,}[^<>]*>[^\0]*(<\/script>){1,}"
	regEx.IgnoreCase = True
	regEx.Global = True
	ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
	LoseScriptTag = ClsTempLoseStr
	Set regEx = Nothing
End Function

Function LoseIFrameTag(ContentStr)
	Dim ClsTempLoseStr,regEx
	ClsTempLoseStr = Cstr(ContentStr)
	Set regEx = New RegExp
	regEx.Pattern = "(<iframe){1,}[^<>]*>[^\0]*(<\/iframe>){1,}"
	regEx.IgnoreCase = True
	regEx.Global = True
	ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
	LoseIFrameTag = ClsTempLoseStr
	Set regEx = Nothing
End Function

Function LoseObjectTag(ContentStr)
	Dim ClsTempLoseStr,regEx
	ClsTempLoseStr = Cstr(ContentStr)
	Set regEx = New RegExp
	regEx.Pattern = "(<object){1,}[^<>]*>[^\0]*(<\/object>){1,}"
	regEx.IgnoreCase = True
	regEx.Global = True
	ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
	LoseObjectTag = ClsTempLoseStr
	Set regEx = Nothing
End Function

Function LoseSpanTag(ContentStr)
	Dim ClsTempLoseStr,regEx
	ClsTempLoseStr = Cstr(ContentStr)
	Set regEx = New RegExp
	regEx.Pattern = "<(\/){0,1}span[^<>]*>"
	regEx.IgnoreCase = True
	regEx.Global = True
	ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
	LoseSpanTag = ClsTempLoseStr
	Set regEx = Nothing
End Function

Function LoseFontTag(ContentStr)
	Dim ClsTempLoseStr,regEx
	ClsTempLoseStr = Cstr(ContentStr)
	Set regEx = New RegExp
	regEx.Pattern = "<(\/){0,1}font[^<>]*>"
	regEx.IgnoreCase = True
	regEx.Global = True
	ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
	LoseFontTag = ClsTempLoseStr
	Set regEx = Nothing
End Function

Function LoseATag(ContentStr)
	Dim ClsTempLoseStr,regEx
	ClsTempLoseStr = Cstr(ContentStr)
	Set regEx = New RegExp
	regEx.Pattern = "<(\/){0,1}a[^<>]*>"
	regEx.IgnoreCase = True
	regEx.Global = True
	ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
	LoseATag = ClsTempLoseStr
	Set regEx = Nothing
End Function

Function LoseDivTag(ContentStr)
	Dim ClsTempLoseStr,regEx
	ClsTempLoseStr = Cstr(ContentStr)
	Set regEx = New RegExp
	regEx.Pattern = "<(\/){0,1}div[^<>]*>"
	regEx.IgnoreCase = True
	regEx.Global = True
	ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
	LoseDivTag = ClsTempLoseStr
	Set regEx = Nothing
End Function

Function LoseStyleTag(ContentStr)
	Dim ClsTempLoseStr,regEx
	ClsTempLoseStr = Cstr(ContentStr)
	Set regEx = New RegExp
	regEx.Pattern = "(<style){1,}[^<>]*>[^\0]*(<\/style>){1,}"
	regEx.IgnoreCase = True
	regEx.Global = True
	ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
	LoseStyleTag = ClsTempLoseStr
	Set regEx = Nothing
End Function
%>