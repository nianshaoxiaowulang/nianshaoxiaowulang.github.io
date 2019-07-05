<%
Function HtmlFormcode(ByVal fString)
    if Not IsNull(fString) Then
        fString = Replace(fString, ">", "&gt;")
        fString = Replace(fString, "<", "&lt;")
        fString = Replace(fString, Chr(34), "&quot;")
        HtmlFormcode = fString
	else
        HtmlFormcode = fString
    end if
End Function

Function LoseHtml(ContentStr)
	Dim ClsTempLoseStr,regEx
	ClsTempLoseStr = Cstr(ContentStr)
	Set regEx = New RegExp
	regEx.Pattern = "<\/*[^<>]*>"
	regEx.IgnoreCase = True
	regEx.Global = True
	ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
	LoseHtml = ClsTempLoseStr
End function
Function ListTitle(TitleStr,TitleNum)
   Dim ClsTitleStr,ClsTitleNum,i,j,ClsTempNum,k,ClsTitleStrResult,LeftStr,RightStr
	   ClsTitleNum = Cint(TitleNum)
	   ClsTempNum = Len(Cstr(TitleStr))
	   if ClsTitleNum > ClsTempNum then
		   ClsTitleNum = ClsTempNum
	   end if
	   ClsTitleStr = Left(Cstr(TitleStr),ClsTitleNum)
	   Dim TempStr
	   For i = 1 to ClsTitleNum - 1
		   TempStr = TempStr & Mid(ClsTitleStr,i,1) & "<br>"
       Next
	   TempStr = TempStr & Right(ClsTitleStr,1)
	   ListTitle = TempStr
End Function
Function  DateFormat(DateStr,Types)
    Dim DateString
	if IsDate(DateStr) = False then
		DateString = ""
	end if
	Select Case Types
	  Case "1" 
		  DateString = Year(DateStr)&"-"&Month(DateStr)&"-"&Day(DateStr)
	  Case "2"
		  DateString = Year(DateStr)&"."&Month(DateStr)&"."&Day(DateStr)
	  Case "3"
		  DateString = Year(DateStr)&"/"&Month(DateStr)&"/"&Day(DateStr)
	  Case "4"
		  DateString = Month(DateStr)&"/"&Day(DateStr)&"/"&Year(DateStr)
	  Case "5"
		  DateString = Day(DateStr)&"/"&Month(DateStr)&"/"&Year(DateStr)
	  Case "6"
		  DateString = Month(DateStr)&"-"&Day(DateStr)&"-"&Year(DateStr)
	  Case "7"
		  DateString = Month(DateStr)&"."&Day(DateStr)&"."&Year(DateStr)
	  Case "8"
		  DateString = Month(DateStr)&"-"&Day(DateStr)
	  Case "9"
		  DateString = Month(DateStr)&"/"&Day(DateStr)
	  Case "10"
		  DateString = Month(DateStr)&"."&Day(DateStr)
	  Case "11"
		  DateString = Month(DateStr)&"月"&Day(DateStr)&"日"
	  Case "12"
		  DateString = Day(DateStr)&"日"&Hour(DateStr)&"时"
	  case "13"
		  DateString = Day(DateStr)&"日"&Hour(DateStr)&"点"
	  Case "14"
		  DateString = Hour(DateStr)&"时"&Minute(DateStr)&"分"
	  Case "15"
		  DateString = Hour(DateStr)&":"&Minute(DateStr)
	  Case "16"
		  DateString = Year(DateStr)&"年"&Month(DateStr)&"月"&Day(DateStr)&"日"
	  Case Else
	  	  DateString = DateStr
	 End Select
	 DateFormat = DateString
 End Function
Function GetRandomID18()
	Dim TempYear,TempMonth,TempDay,TempHour,TempMinute,TempSecond,RandomFigure
	Dim TempStr,NowTime
	NowTime = Now()
	TempYear =  Right(CStr(Year(NowTime)),2)
	TempMonth =  CStr(Month(NowTime))
	if Len(TempMonth) = 1 then
		TempHour = "0" & TempMonth
	end if
	TempDay =  CStr(Day(NowTime))
	if Len(TempDay) = 1 then
		TempHour = "0" & TempDay
	end if
	TempHour = CStr(Hour(NowTime))
	if Len(TempHour) = 1 then
		TempHour = "0" & TempHour
	end if
	TempMinute = CStr(Minute(NowTime))
	if Len(TempMinute) = 1 then
		TempMinute = "0" & TempMinute
	end if
	TempSecond = CStr(Second(NowTime))
	if Len(TempSecond) = 1 then
		TempSecond = "0" & TempSecond
	end if
	Randomize 
	RandomFigure = CStr(Int((99999 * Rnd) + 1))
	GetRandomID18 = TempYear & TempMonth & TempDay & TempHour & TempMinute & TempSecond & RandomFigure
End Function


Function NewsFileName(NewsFileNameStr,IClassID,INewsID)
		NewsFileName = ""
		If NewsFileNameStr="" or isnull(NewsFileNameStr) then
			NewsFileName = Hour(Now()) & Minute(Now()) & Second(Now()) & CStr(Int((999 * Rnd) + 1))
		Else
			If Instr(1,NewsFileNameStr,"U",1)<>0 then
				If Instr(1,NewsFileNameStr,"Y",1)<>0 then
					NewsFileName = NewsFileName & Year(Now())
				End If
				If Instr(1,NewsFileNameStr,"M",1)<>0 then
					If Len(Cstr(Month(Now()))) < 2 then
						NewsFileName = NewsFileName &"_0"& Month(Now())
					Else
						NewsFileName = NewsFileName &"_"& Month(Now())
					End If
				End If
				If Instr(1,NewsFileNameStr,"D",1)<>0 then
					If Len(Cstr(Day(Now())))<2 then
						NewsFileName = NewsFileName &"_0"& Day(Now())
					Else
						NewsFileName = NewsFileName &"_"& Day(Now())
					End if
				End If
				If Instr(1,NewsFileNameStr,"H",1)<>0 then
					If Len(Cstr(Hour(Now())))<2 then
						NewsFileName = NewsFileName &"_0"& Hour(Now())
					Else
						NewsFileName = NewsFileName &"_"& Hour(Now())
					End If
				End If
				If Instr(1,NewsFileNameStr,"I",1)<>0 then
					If Len(Cstr(Minute(Now())))<2 then
						NewsFileName = NewsFileName &"_0"& Minute(Now())
					Else
						NewsFileName = NewsFileName &"_"& Minute(Now())
					End If
				End If
				If Instr(1,NewsFileNameStr,"S",1)<>0 then
					If Len(Cstr(Second(Now())))<2 then
						NewsFileName = NewsFileName &"_0"& Second(Now())
					Else
						NewsFileName = NewsFileName &"_"& Second(Now())
					End If
				End If
				If Instr(1,NewsFileNameStr,"A",1)<>0 then
					NewsFileName = NewsFileName &"_"& IClassID
				End If
				If Instr(1,NewsFileNameStr,"N",1)<>0 then
					NewsFileName = NewsFileName &"_"& INewsID
				End If
				If Instr(1,NewsFileNameStr,"Z",1)<>0 then
					NewsFileName = NewsFileName &"_"& CStr(Int((99 * Rnd) + 1))
				End If
				If Instr(1,NewsFileNameStr,"X",1)<>0 then
					NewsFileName = NewsFileName &"_"& CStr(Int((999 * Rnd) + 1))
				End If
				If Instr(1,NewsFileNameStr,"C",1)<>0 then
					NewsFileName = NewsFileName &"_"& CStr(Int((9999 * Rnd) + 1))
				End If
				If Instr(1,NewsFileNameStr,"V",1)<>0 then
					NewsFileName = NewsFileName &"_"& CStr(Int((99999 * Rnd) + 1))
				End If
			Else
				If Instr(1,NewsFileNameStr,"Y",1)<>0 then
					NewsFileName = NewsFileName & Year(Now())
				End If
				If Instr(1,NewsFileNameStr,"M",1)<>0 then
					If Len(Cstr(Month(Now()))) < 2 then
						NewsFileName = NewsFileName &"0"& Month(Now())
					Else
						NewsFileName = NewsFileName & Month(Now())
					End If
				End If
				If Instr(1,NewsFileNameStr,"D",1)<>0 then
					If Len(Cstr(Day(Now())))<2 then
						NewsFileName = NewsFileName &"0"& Day(Now())
					Else
						NewsFileName = NewsFileName & Day(Now())
					End if
				End If
				If Instr(1,NewsFileNameStr,"H",1)<>0 then
					If Len(Cstr(Hour(Now())))<2 then
						NewsFileName = NewsFileName &"0"& Hour(Now())
					Else
						NewsFileName = NewsFileName & Hour(Now())
					End If
				End If
				If Instr(1,NewsFileNameStr,"I",1)<>0 then
					If Len(Cstr(Minute(Now())))<2 then
						NewsFileName = NewsFileName &"0"& Minute(Now())
					Else
						NewsFileName = NewsFileName & Minute(Now())
					End If
				End If
				If Instr(1,NewsFileNameStr,"S",1)<>0 then
					If Len(Cstr(Second(Now())))<2 then
						NewsFileName = NewsFileName &"0"& Second(Now())
					Else
						NewsFileName = NewsFileName & Second(Now())
					End If
				End If
				If Instr(1,NewsFileNameStr,"A",1)<>0 then
					NewsFileName = NewsFileName & IClassID
				End If
				If Instr(1,NewsFileNameStr,"N",1)<>0 then
					NewsFileName = NewsFileName & INewsID
				End If
				If Instr(1,NewsFileNameStr,"Z",1)<>0 then
					NewsFileName = NewsFileName & CStr(Int((99 * Rnd) + 1))
				End If
				If Instr(1,NewsFileNameStr,"X",1)<>0 then
					NewsFileName = NewsFileName & CStr(Int((999 * Rnd) + 1))
				End If
				If Instr(1,NewsFileNameStr,"C",1)<>0 then
					NewsFileName = NewsFileName & CStr(Int((9999 * Rnd) + 1))
				End If
				If Instr(1,NewsFileNameStr,"V",1)<>0 then
					NewsFileName = NewsFileName & CStr(Int((99999 * Rnd) + 1))
				End If
			End If
		End If
		If Left(NewsFileName,1)="_" then
			NewsFileName = Right(NewsFileName,Len(Cstr(NewsFileName))-1)
		End If
		NewsFileName = NewsFileName
End Function
Function GotTopic(Str,StrLen)
	Dim l,t,c, i,LableStr,regEx,Match,Matches
	If StrLen=0 then
		GotTopic=""
		exit function
	End If
	if IsNull(Str) then 
		GotTopic = ""
		Exit Function
	end if
	if Str = "" then
		GotTopic=""
		Exit Function
	end if
	'Set regEx = New RegExp
	'regEx.Pattern = "\[[^\[\]]*\]"
	'regEx.IgnoreCase = True
	'regEx.Global = True
	'Set Matches = regEx.Execute(Str)
	'For Each Match in Matches
	'		LableStr = LableStr & Match.Value
	'Next
	'Str = regEx.Replace(Str,"")
	Str=Replace(Replace(Replace(Replace(Str,"&nbsp;"," "),"&quot;",Chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	strlen=Clng(strLen)
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			GotTopic=left(str,i)
			exit for
		else
			GotTopic=str
		end if
	next
	GotTopic = Replace(Replace(Replace(Replace(GotTopic," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;")' & LableStr
end function

function IsValidEmail(email)
	dim names, name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	if UBound(names) <> 1 then
	   IsValidEmail = false
	   exit function
	end if
	for each name in names
		if Len(name) <= 0 then
			IsValidEmail = false
    		exit function
		end if
		for i = 1 to Len(name)
		    c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
		       IsValidEmail = false
		       exit function
		     end if
	   next
	   if Left(name, 1) = "." or Right(name, 1) = "." then
    	  IsValidEmail = false
	      exit function
	   end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmail = false
	   exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
	   IsValidEmail = false
	   exit function
	end if
	if InStr(email, "..") > 0 then
	   IsValidEmail = false
	end if
end function
Sub ReturnRefreshError()

End Sub
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
Function IsExpired(strClassString)
	On Error Resume Next
	IsExpired = True
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)

	If 0 = Err Then
		Select Case strClassString
			Case "Persits.Jpeg"
				If xTestObj.Expires > Now Then
					IsExpired = False
				End if
			Case "wsImage.Resize"
				If instr(xTestObj.errorinfo,"已经过期") = 0 Then
					IsExpired = False
				End if
			Case "SoftArtisans.ImageGen"
				xTestObj.CreateImage 500, 500, rgb(255,255,255)
				If Err = 0 Then
					IsExpired = False
				End if
		End Select
	End if
	Set xTestObj = Nothing
	Err = 0
End Function
function SendMail(MailtoAddress,MailtoName,Subject,MailBody,FromName,MailFrom,Priority)
	on error resume next
	Dim JMail
	Set JMail=Server.CreateObject("JMail.Message")
	if err then
		SendMail= "<br><li>没有安装JMail组件</li>"
		err.clear
		exit function
	end if
	JMail.Charset="gb2312"      
	JMail.silent=true
	JMail.ContentType = "text/html"  
	JMail.MailServerUserName = MailServerUserName  
   	JMail.MailServerPassWord = MailServerPassword      
  	JMail.MailDomain = MailDomain      
	JMail.AddRecipient MailtoAddress,MailtoName   
	JMail.Subject=Subject        
	JMail.HMTLBody=MailBody   
	JMail.Body=MailBody         
	JMail.FromName=FromName      
	JMail.From = MailFrom       
	JMail.Priority=Priority          
	JMail.Send(MailServer)
	SendMail =JMail.ErrorMessage
	JMail.Close
	Set JMail=nothing
end function

function ReplaceBadChar(strChar)
	if strChar="" then
		ReplaceBadChar=""
	else
		ReplaceBadChar=replace(replace(replace(replace(replace(replace(replace(strChar,"'",""),"*",""),"?",""),"(",""),")",""),"<",""),".","")
	end if
end function

Function ReplaceRemoteUrl(NewsContent,SaveFilePath,FunDoMain,DummyPath)
	Dim re,RemoteFile,RemoteFileurl,SaveFileName,FileName,FileExtName,SaveImagePath
	Set re = New RegExp
	re.IgnoreCase = True
	re.Global=True
	re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}((\w)+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(gif|jpg|png|bmp)))"
	Set RemoteFile = re.Execute(NewsContent)
	Set re = Nothing
	For Each RemoteFileurl in RemoteFile
		SaveFileName = Mid(RemoteFileurl,InstrRev(RemoteFileurl,"/")+1)
		FileExtName = Mid(SaveFileName,InstrRev(SaveFileName,".")+1)
		Call SaveRemoteFile(DummyPath & SaveFilePath & "/" & SaveFileName,RemoteFileurl)
		NewsContent = Replace(NewsContent,RemoteFileurl,FunDoMain & SaveFilePath & "/" & SaveFileName)
	Next
	ReplaceRemoteUrl = NewsContent
End Function

Sub SaveRemoteFile(LocalFileName,RemoteFileUrl)
	On Error Resume Next
	Dim StreamObj,Retrieval,GetRemoteData,TempHTTPObj
	TempHTTPObj = "MSXML2.XMLHTTP"
	Set Retrieval = Server.CreateObject(TempHTTPObj)
	With Retrieval
		.Open "Get", RemoteFileUrl, False, "", ""
		.Send
		if Err.Number <> 0 then
			Err.Clear
			Set Retrieval = Nothing
			Exit Sub
		end if
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing
	Set StreamObj = Server.CreateObject("Adodb.Stream")
	With StreamObj
		.Type = 1
		.Open
		.Write GetRemoteData
		.SaveToFile Server.MapPath(LocalFileName),2
		.Cancel()
		.Close()
	End With
	Set StreamObj = Nothing
End Sub

function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function
function strLength(str)
	ON ERROR RESUME NEXT
	dim WINNT_CHINESE
	WINNT_CHINESE    = (len("中国")=2)
	if WINNT_CHINESE then
        dim l,t,c
        dim i
        l=len(str)
        t=l
        for i=1 to l
        	c=asc(mid(str,i,1))
            if c<0 then c=c+65536
            if c>255 then
                t=t+1
            end if
        next
        strLength=t
    else 
        strLength=len(str)
    end if
    if err.number<>0 then err.clear
end function
Function InterceptStr(SourceStr,StrLen,AnnexStr)   
	InterceptStr = GotTopic(SourceStr,StrLen)
	if Not IsNull(AnnexStr) then
		 InterceptStr = InterceptStr & CStr(AnnexStr)
	end if
End Function
Function GetDoMain()
	Dim TempPath
	if Request.ServerVariables("SERVER_PORT")="80" then
		GetDoMain = Request.ServerVariables("SERVER_NAME")
	else
		GetDoMain = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")
	end if
	TempPath = Request.ServerVariables("APPL_MD_PATH")
	TempPath = Right(TempPath,Len(TempPath)-InStr(TempPath,"Root")-3)
	GetDoMain = "http://" & GetDoMain & TempPath
End Function
Function GetConfigDoMain()
	Dim ConfigSql,RsConfigObj
	ConfigSql = "Select DoMain,MakeType from FS_Config"
	Set RsConfigObj = Conn.Execute(ConfigSql)
	if Not RsConfigObj.Eof then
		GetConfigDoMain = RsConfigObj("DoMain")
	else
		GetConfigDoMain = GetDoMain
	end if
	Set RsConfigObj = Nothing
End Function

Function GetCurrentPath()
	Dim TempPath,Path
	TempPath = Request.ServerVariables("Path_info")
	Path = Left(TempPath,InstrRev(TempPath,"/"))
	GetCurrentPath = GetDoMain & Path
End Function
Function GetNewsTitlePara(Para)
	Dim TempArray
	if Len(Para) <> 9 then
		TempArray = Array("","","")
	else
		TempArray = Array(Left(Para,7),Mid(Para,8,1),Right(Para,1))
	end if
	GetNewsTitlePara = TempArray
End Function

Function GetHTMLTitle(StylePara,Title)
	Dim TempArray
	TempArray = GetNewsTitlePara(StylePara)
	if TempArray(0)<>"#UUUUUU" then
		GetHTMLTitle = "<font color=""" & TempArray(0) & """>" & Title & "</font>"
	else
		GetHTMLTitle =  Title
	end if
	if TempArray(1) = "1" then
		GetHTMLTitle = "<strong>" & GetHTMLTitle & "</strong>"
	end if
	if TempArray(2) = "1" then
		GetHTMLTitle = "<em>" & GetHTMLTitle & "</em>"
	end if
End Function
Function RemoveVirtualPath(Path)
	Dim PathInfo
	if Path <> "" then  
		if SysRootDir <> ""   then 
			PathInfo = Mid(Path,InStr(Path,SysRootDir)+Len(SysRootDir)+1)
		else
			PathInfo = Path
		end if
	else
		PathInfo = Path
	end if
	RemoveVirtualPath = PathInfo
End Function
Function RemoveSpecialStr(Path)
	Dim PathInfo
	if Path <> "" then  
		if SysRootDir <> ""   then 
			PathInfo = Mid(Path,InStr(Path,SysRootDir)+Len(SysRootDir))
		else
			PathInfo = Path
		end if
	else
		PathInfo = Path
	end if
	RemoveSpecialStr = PathInfo
End Function
Function GetVirtualPath()
	GetVirtualPath = Request.ServerVariables("APPL_MD_PATH")
	GetVirtualPath = Right(GetVirtualPath,Len(GetVirtualPath)-InStr(GetVirtualPath,"Root")-3)
End Function

Function SaveOption(AllOption,Optiontype)
	Dim StrOptionList,Inti,TempRs
	StrOptionList=split(Replace(Replace(AllOption,"""",""),"'",""),",")
	For Inti=0 to ubound(StrOptionList)
		Set TempRs=Conn.Execute("Select * from FS_Routine where name='"&StrOptionList(Inti)&"' and type="&Optiontype)
		If Temprs.eof and trim(StrOptionList(Inti))<>"" then 
			Conn.execute("insert into FS_Routine(name,url,type) values('"& StrOptionList(Inti) &"','',"&Optiontype&")" )
		End If
		TempRs.Close
	Next
	Set TempRs=Nothing
End Function

Function IStrLen(TempStr)
	Dim iLen,i,StrAsc
	iLen=0
	for i=1 to len(TempStr)
			StrAsc=Abs(Asc(Mid(TempStr,i,1)))
			if StrAsc>255 then
				iLen=iLen+2
			else
				iLen=iLen+1
			end if
	next
	IStrLen=iLen
End Function

Function ClassList(StrClass)
	Dim Rs,StrValue
	Set Rs = Conn.Execute("select ClassEName,ClassID,ClassCName from FS_newsclass where ParentID = '0' and DelFlag=0 and IsOutClass=0 order by AddTime desc")
	do while Not Rs.Eof
		StrValue=" Value=""" & RS(StrClass) & """"
		ClassList = ClassList & "<option"&StrValue&">" & Rs("ClassCName") & chr(10) & chr(13)
		ClassList = ClassList & ChildClassList(Rs("ClassID"),"",StrClass)
		Rs.MoveNext	
	loop
	Rs.Close
	Set Rs = Nothing
End Function
Function ChildClassList(ClassID,Temp,StrClass)
	Dim TempRs,TempStr,StrValue
	Set TempRs = Conn.Execute("Select ClassEName,ClassID,ClassCName,ChildNum from FS_NewsClass where ParentID = '" & ClassID & "' and DelFlag=0 order by AddTime desc ")
	TempStr = Temp & " -- "
	do while Not TempRs.Eof
		StrValue=" Value=""" & TempRs(StrClass) & """"
		ChildClassList = ChildClassList & "<option"&StrValue&">" & TempStr & TempRs("ClassCName") & "</option>"& chr(10) & chr(13)
		ChildClassList = ChildClassList & ChildClassList(TempRs("ClassID"),TempStr,StrClass)
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function

Function CreateDateDir(Path)
	Dim sBuild,FSO
	sBuild=path&"\"&year(Now())&"-"&month(now())
	Set FSO = Server.CreateObject(G_FS_FSO)
	If (FSO.FolderExists(sBuild)) then
	else
		FSO.CreateFolder(sBuild)
	End IF
	sBuild=sBuild&"\"&day(Now())
	If (FSO.FolderExists(sBuild)) then
	else
		FSO.CreateFolder(sBuild)
	End IF
	set FSO=Nothing
End Function
Function NoCSSHackAdmin(Str,StrTittle) '过滤跨站脚本和HTML标签
	Dim regEx
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Pattern = "<|>|\t"
	If regEx.Test(LCase(Str)) Then
		Response.Write "<script>alert('"& StrTittle &"含有非法字符(<,>,tab)');history.back();</script>"
		Response.End
	End If
	Set regEx = Nothing
	NoCSSHackAdmin = Str
End Function
Function NoCSSHackInput(Str) '过滤跨站脚本和HTML标签
	Dim regEx
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Pattern = "<|>|(script)|on(mouseover|mouseon|mouseout|click|dblclick|blur|focus|change)|url|eval|\t"
	If regEx.Test(LCase(Str)) Then
		Response.Write "<script>alert('你的输入含有非法字符(<,>,tab,script等)，请检查后再提交！');history.back();</script>"
		Response.End
	End If
	Set regEx = Nothing
	NoCSSHackInput = Str
End Function
Function NoCSSHackContent(Str) '过滤跨站脚本，只过滤脚本，对HTML不过滤
	Dim regEx
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Pattern = "(script)|on(mouseover|mouseon|mouseout|click|dblclick|blur|focus|change)|(url\()|eval"
	If regEx.Test(LCase(Str)) Then
		Response.Write "<script>alert('你提交的内容含有非法字符(不能包含脚本)，请检查后再提交！');history.back();</script>"
		Response.End
	End If
	Set regEx = Nothing
	NoCSSHackContent = Str
End Function
%>