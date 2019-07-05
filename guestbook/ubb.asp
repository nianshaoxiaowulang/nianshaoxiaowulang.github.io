<%
'**************************************
'**		UBB.asp
'**
'** 文件说明：UBB转换函数
'** 修改日期：2005-04-07
'**************************************

function UBBCode(strContent,issuper)

	strContent = HTMLEncode(strContent)
	dim re
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True

	re.Pattern="(javascript)"
	strContent=re.Replace(strContent,"&#106avascript")
	re.Pattern="(jscript:)"
	strContent=re.Replace(strContent,"&#106script:")
	re.Pattern="(js:)"
	strContent=re.Replace(strContent,"&#106s:")
	re.Pattern="(value)"
	strContent=re.Replace(strContent,"&#118alue")
	re.Pattern="(about:)"
	strContent=re.Replace(strContent,"about&#58")
	re.Pattern="(file:)"
	strContent=re.Replace(strContent,"file&#58")
	re.Pattern="(document.cookie)"
	strContent=re.Replace(strContent,"documents&#46cookie")
	re.Pattern="(vbscript:)"
	strContent=re.Replace(strContent,"&#118bscript:")
	re.Pattern="(vbs:)"
	strContent=re.Replace(strContent,"&#118bs:")
	re.Pattern="(on(mouse|exit|error|click|key))"
	strContent=re.Replace(strContent,"&#111n$2")

	if UBBcfg_face=1 or issuper=1 then
		re.Pattern="(\[face([0-9]{2})\])"
		strContent=re.Replace(strContent,"<img src=images/faces/$2.gif width=20 height=20 border=0 align=middle>")
	end if

	if UBBcfg_pic=1 or issuper=1 then
		re.Pattern="(\[IMG\])(.[^\[]*)(\[\/IMG\])"
		strContent=re.Replace(strContent,"<a href=""$2"" target=_blank><IMG border='0' SRC=""$2"" alt=点击在新窗口显示图片 onload=""javascript:if(this.width>screen.width-390)this.width=screen.width-390;if(this.height>480)this.height=480""></a>")
	end if

	if UBBcfg_swf=1 or issuper=1 then
		re.Pattern="(\[FLASH=([0-9]{1,3}),([0-9]{1,3})\])(.[^\[]*)(\[\/FLASH\])"
		strContent= re.Replace(strContent,"<OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$2 height=$3><PARAM NAME=movie VALUE=""$4""><PARAM NAME=quality VALUE=high><embed src=""$4"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$2 height=$3>$4</embed></OBJECT>")
	end if

	if UBBcfg_font=1 or issuper=1 then
		strContent = Replace(strContent,"[/face]","</font>")
		re.Pattern="(\[face=*(.[^\]]*)\])"
		strContent = re.Replace(strContent,"<font face=$2>")
	end if

	if UBBcfg_color=1 or issuper=1 then
		strContent = Replace(strContent,"[/color]","</font>")
		re.Pattern="(\[color=*(.[^\]]*)\])"
		strContent = re.Replace(strContent,"<font color=$2>")

	if UBBcfg_size=1 or issuper=1 then
		strContent = Replace(strContent,"[/size]","</font>")
		re.Pattern="(\[size=*(.[^\]]*)\])"
		strContent = re.Replace(strContent,"<font size=$2>")
	end if

	end if

	if UBBcfg_b=1 or issuper=1 then
		strContent = Replace(strContent,"[B]","<b>")
		strContent = Replace(strContent,"[/B]","</b>")
	end if

	if UBBcfg_i=1 or issuper=1 then
		strContent = Replace(strContent,"[I]","<i>")
		strContent = Replace(strContent,"[/I]","</i>")
	end if

	if UBBcfg_u=1 or issuper=1 then
		strContent = Replace(strContent,"[U]","<u>")
		strContent = Replace(strContent,"[/U]","</u>")
	end if

	if UBBcfg_center=1 or issuper=1 then
		strContent = Replace(strContent,"[center]","<center>")
		strContent = Replace(strContent,"[/center]","</center>")
	end if

	if UBBcfg_shadow=1 or issuper=1 then
		re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
		strContent=re.Replace(strContent,"<table width=$1><tr><td style=""filter:shadow(color=$2, strength=$3)"">$4</td></tr></table>")
	end if

	if UBBcfg_glow=1 or issuper=1 then
		re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
		strContent=re.Replace(strContent,"<table width=$1><tr><td style=""filter:glow(color=$2, strength=$3)"">$4</td></tr></table>")
	end if

	if UBBcfg_URL=1 or issuper=1 then
		re.Pattern="(\[URL\])(.[^\[]*)(\[\/URL\])"
		strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$2</A>")
		re.Pattern="(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
		strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$3</A>")
		re.Pattern = "^(http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
		strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
		re.Pattern = "(http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
		strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
		re.Pattern = "[^>=""](http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
		strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
		re.Pattern = "^(ftp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
		strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
		re.Pattern = "(ftp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
		strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
		re.Pattern = "[^>=""](ftp://[A-Za-z0-9\.\/=\?%\-&_~`@':+!]+)"
		strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
		re.Pattern = "^(rtsp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
		strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
		re.Pattern = "(rtsp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
		strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
		re.Pattern = "[^>=""](rtsp://[A-Za-z0-9\.\/=\?%\-&_~`@':+!]+)"
		strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
		re.Pattern = "^(mms://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
		strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
		re.Pattern = "(mms://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
		strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
		re.Pattern = "[^>=""](mms://[A-Za-z0-9\.\/=\?%\-&_~`@':+!]+)"
		strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
	end if

	if UBBcfg_email=1 or issuper=1 then
		re.Pattern="(\[EMAIL\])(.[^\]]*)(\[\/EMAIL\])"
		strContent= re.Replace(strContent,"<A HREF=""mailto:$2"">$2</A>")
		re.Pattern="(\[EMAIL=(.[^\]]*)\])(.[^\]]*)(\[\/EMAIL\])"
		strContent= re.Replace(strContent,"<A HREF=""mailto:$2"">$3</A>")
	end if

	set re=Nothing
	UBBCode=strContent

end function

function HTMLEncode(fString)
	if not isnull(fString) then
		fString = back_filter(fString)
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
		fString = Replace(fString, CHR(32), "&nbsp;")
		fString = Replace(fString, CHR(9), "&nbsp;")
		fString = Replace(fString, CHR(34), "&quot;")
		fString = Replace(fString, CHR(39), "&#39;")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
		fString = Replace(fString, CHR(10), "<BR> ")
		fString = Replace(fString, CHR(36), "&#36;")
		HTMLEncode = fString
	end if
end function
%>