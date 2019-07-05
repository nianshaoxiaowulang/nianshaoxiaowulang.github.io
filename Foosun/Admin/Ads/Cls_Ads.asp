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
	Dim MyFile,CrHNJS,TempStateFlag,DnsPath,GetUrl,LDnsPath,AdsJSStr,AdsTempStr,JsFileName,objStream,AdsTempStrRight
	Dim TempSysRootDir
	if SysRootDir = "" then
		TempSysRootDir = ""
	else
		TempSysRootDir = "/" & SysRootDir
	end if
	function AdsTempPicStr(Location)
	    dim FunLocation,FunAdsObj
		FunLocation = clng(Location)
		Set FunAdsObj = Conn.Execute("select * from FS_Ads where Location="&FunLocation&"")
		if InStr(1,LCase(FunAdsObj("LeftPicPath")),".swf",1)<>0 Then
			If InStr(1,LCase(FunAdsObj("LeftPicPath")),"http://")<>0 then
				AdsTempStr="<EMBED src="""& FunAdsObj("LeftPicPath") &""" quality=high WIDTH="""& FunAdsObj("PicWidth") &""" HEIGHT="""& FunAdsObj("PicHeight") &""" TYPE=""application/x-shockwave-flash"" PLUGINSPAGE=""http://www.macromedia.com/go/getflashplayer""></EMBED>"
			Else
				AdsTempStr="<EMBED src="""& TempSysRootDir & FunAdsObj("LeftPicPath") &""" quality=high WIDTH="""& FunAdsObj("PicWidth") &""" HEIGHT="""& FunAdsObj("PicHeight") &""" TYPE=""application/x-shockwave-flash"" PLUGINSPAGE=""http://www.macromedia.com/go/getflashplayer""></EMBED>"
			End If
		Else
			If InStr(1,LCase(FunAdsObj("LeftPicPath")),"http://")<>0 then
				AdsTempStr="<a href="""& TempSysRootDir & "/" & PlusDir &"/Ads/AdsClick.asp?Location="& FunLocation &""" target=_blank><img src="""& FunAdsObj("LeftPicPath") &""" border=0 width="""& FunAdsObj("PicWidth") &""" height="""& FunAdsObj("PicHeight") &""" alt="""& FunAdsObj("Explain") &""" align=top></a>"
			Else
				AdsTempStr="<a href="""& TempSysRootDir & "/" & PlusDir &"/Ads/AdsClick.asp?Location="& FunLocation &""" target=_blank><img src="""& TempSysRootDir & FunAdsObj("LeftPicPath") &""" border=0 width="""& FunAdsObj("PicWidth") &""" height="""& FunAdsObj("PicHeight") &""" alt="""& FunAdsObj("Explain") &""" align=top></a>"
			End If
		End If
		if InStr(1,LCase(FunAdsObj("RightPicPath")),".swf",1)<>0 Then
			If InStr(1,LCase(FunAdsObj("RightPicPath")),"http://")<>0 then
				AdsTempStrRight="<EMBED src="""& FunAdsObj("RightPicPath") &""" quality=high WIDTH="""& FunAdsObj("PicWidth") &""" HEIGHT="""& FunAdsObj("PicHeight") &""" TYPE=""application/x-shockwave-flash"" PLUGINSPAGE=""http://www.macromedia.com/go/getflashplayer""></EMBED>"
			Else
				AdsTempStrRight="<EMBED src="""& TempSysRootDir & FunAdsObj("RightPicPath") &""" quality=high WIDTH="""& FunAdsObj("PicWidth") &""" HEIGHT="""& FunAdsObj("PicHeight") &""" TYPE=""application/x-shockwave-flash"" PLUGINSPAGE=""http://www.macromedia.com/go/getflashplayer""></EMBED>"
			End If
		Else
			If InStr(1,LCase(FunAdsObj("RightPicPath")),"http://")<>0 then
				AdsTempStrRight="<a href="""& TempSysRootDir & "/" & PlusDir &"/Ads/AdsClick.asp?Location="& FunLocation &""" target=_blank><img src="""& FunAdsObj("RightPicPath") &""" border=0 width="""& FunAdsObj("PicWidth") &""" height="""& FunAdsObj("PicHeight") &""" alt="""& FunAdsObj("Explain") &""" align=top></a>"
			Else
				AdsTempStrRight="<a href="""& TempSysRootDir & "/" & PlusDir &"/Ads/AdsClick.asp?Location="& FunLocation &""" target=_blank><img src="""& TempSysRootDir & FunAdsObj("RightPicPath") &""" border=0 width="""& FunAdsObj("PicWidth") &""" height="""& FunAdsObj("PicHeight") &""" alt="""& FunAdsObj("Explain") &""" align=top></a>"
			End If
		End If
		FunAdsObj.close
		set FunAdsObj = nothing
	end function
 
        Sub ShowAds(TempLocation)
		    dim ShowAdsStr,ShowAdsLocation,ShowAdsObj
			ShowAdsLocation = clng(TempLocation)
			AdsTempPicStr(ShowAdsLocation)
			ShowAdsStr = AdsTempStr
			Set ShowAdsObj = Conn.Execute("select State from FS_Ads where Location="&ShowAdsLocation&"")
			if ShowAdsObj("State")<>"1" then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr = "document.write('"& ShowAdsStr &"');" & vbCrLf & _
				           "document.write('<script src="& TempSysRootDir & "/" & PlusDir &"/Ads/show.asp?Location="&ShowAdsLocation&"></script>');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")) = false then
				MyFile.CreateFolder(Server.MapPath(TempSysRootDir&"\JS\AdsJs"))
			End If
			 if MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& ShowAdsLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& ShowAdsLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& ShowAdsLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			ShowAdsObj.close
			Set ShowAdsObj = Nothing
		 End Sub
		 
		 Sub NewWindow(TempLocation)
		    dim NewWindowObj,NewWindowLocation,dialogConent,dialogConent1 ,sUrl
			NewWindowLocation = clng(TempLocation)
			Set NewWindowObj = Conn.Execute("select * from FS_Ads where Location="&NewWindowLocation&"")
			if NewWindowObj("State")<>"1" then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				If Instr(1,LCase(NewWindowObj("LeftPicPath")),"http://") <> 0 then
					AdsJSStr = "window.open('/foosun/admin/Ads/pic.asp?pic="&NewWindowLocation&"','','width="& NewWindowObj("PicWidth") &",height="& NewWindowObj("PicHeight") &",scrollbars=1');"
				Else
					AdsJSStr = "window.open('"& TempSysRootDir &"/foosun/admin/Ads/pic.asp?pic="&NewWindowLocation&"','','width="& NewWindowObj("PicWidth") &",height="& NewWindowObj("PicHeight") &",scrollbars=1');" & vbCrLf & _
					"document.write('<script src="& TempSysRootDir & "/" & PlusDir &"/Ads/show.asp?Location="&NewWindowLocation&"></script>');"
				End If
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")) = false then
				MyFile.CreateFolder(Server.MapPath(TempSysRootDir&"\JS\AdsJs"))
			End If
			 if MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& NewWindowLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& NewWindowLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& NewWindowLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=Nothing
			NewWindowObj.close
			Set NewWindowObj = Nothing
		 End Sub
		 
		 Sub OpenWindow(TempLocation)
		    dim OpenWindowObj,OpenWindowLocation
			OpenWindowLocation = clng(TempLocation)
			Set OpenWindowObj = Conn.Execute("select * from FS_Ads where Location="&OpenWindowLocation&"")
			if OpenWindowObj("State")<>"1" then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				If Instr(1,LCase(OpenWindowObj("LeftPicPath")),"http://") <> 0 then
					AdsJSStr = "window.open('/foosun/admin/Ads/pic.asp?pic="&OpenWindowLocation&"','_blank');" 
				Else
					AdsJSStr = "window.open('"& TempSysRootDir &"/foosun/admin/Ads/pic.asp?pic="&OpenWindowLocation&"','_blank');"
				End If
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")) = false then
				MyFile.CreateFolder(Server.MapPath(TempSysRootDir&"\JS\AdsJs"))
			End If
			 if MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& OpenWindowLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& OpenWindowLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& OpenWindowLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			OpenWindowObj.close
			Set OpenWindowObj = Nothing
		 End Sub
		 
		 Sub FilterAway(TempLocation)
		    dim FilterAwayStr,FilterAwayLocation,FilterAwayObj
			FilterAwayLocation = clng(TempLocation)
			AdsTempPicStr(FilterAwayLocation)
			FilterAwayStr = AdsTempStr
			Set FilterAwayObj = Conn.Execute("select * from FS_Ads where Location="&FilterAwayLocation&"")
			if FilterAwayObj("State")<>"1" then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr = "FilterAwayStr=(document.layers)?true:false;" & vbCrLf & _
						   "if(FilterAwayStr){document.write('<layer id=FilterAwayT onLoad=""moveToAbsolute(layer1.pageX-160,layer1.pageY);clip.height="& FilterAwayObj("PicHeight") &";clip.width="& FilterAwayObj("PicWidth") &"; visibility=show;""><layer id=FilterAwayF position:absolute; bottom:20; center:1>"& FilterAwayStr &"</layer></layer>');}" & vbCrLf & _
						   "else{document.write('<div style=""position:absolute;bottom:"& cint(FilterAwayObj("PicHeight")+20) &"; center:1;""><div id=FilterAwayT style=""position:absolute; width:"& FilterAwayObj("PicWidth") &"; height:"& FilterAwayObj("PicHeight") &";clip:rect(0,"& FilterAwayObj("PicWidth") &","& FilterAwayObj("PicHeight") &",0)""><div id=FilterAwayF style=""position:absolute;bottom:20; center:1"">"& FilterAwayStr &"</div></div></div>');} " & vbCrLf & _
						   "document.write('<script language=javascript src="& TempSysRootDir &"/Ads/CreateJs/FilterAway.js></script>');" & vbCrLf & _
				           "document.write('<script src="& TempSysRootDir & "/" & PlusDir &"/Ads/show.asp?Location="&FilterAwayLocation&"></script>');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")) = false then
				MyFile.CreateFolder(Server.MapPath(TempSysRootDir&"\JS\AdsJs"))
			End If
			 if MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& FilterAwayLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& FilterAwayLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& FilterAwayLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			FilterAwayObj.close
			Set FilterAwayObj = Nothing
		 End Sub
		 
		 Sub DialogBox(TempLocation)
		    dim DialogBoxObj,DialogBoxLocation
			DialogBoxLocation = clng(TempLocation)
			Set DialogBoxObj = Conn.Execute("select * from FS_Ads where Location="&DialogBoxLocation&"")
			if DialogBoxObj("State")<>"1" then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				If Instr(1,LCase(DialogBoxObj("LeftPicPath")),"http://") <> 0 then
					AdsJSStr = "window.showModalDialog('/foosun/admin/Ads/pic.asp?pic="&DialogBoxLocation&"','','dialogWidth:"& DialogBoxObj("PicWidth") &"px;dialogHeight:"& DialogBoxObj("PicHeight") &"px;center:0;status:no');" 
				Else
					AdsJSStr = "window.showModalDialog('"& TempSysRootDir & "/foosun/admin/Ads/pic.asp?pic="&DialogBoxLocation&"','','dialogWidth:"& DialogBoxObj("PicWidth") &"px;dialogHeight:"& DialogBoxObj("PicHeight") &"px;center:0;status:no');"
				End If
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")) = false then
				MyFile.CreateFolder(Server.MapPath(TempSysRootDir&"\JS\AdsJs"))
			End If
			 if MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& DialogBoxLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& DialogBoxLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& DialogBoxLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
		 	DialogBoxObj.Close
			Set DialogBoxObj = Nothing
		 End Sub
		 
		 Sub ClarityBox(TempLocation)
		    dim ClarityBoxObj,ClarityBoxLocation,ClarityBoxStr
			ClarityBoxLocation = clng(TempLocation)
			AdsTempPicStr(ClarityBoxLocation)
			ClarityBoxStr = AdsTempStr
			Set ClarityBoxObj = Conn.Execute("select * from FS_Ads where Location="&ClarityBoxLocation&"")
			if ClarityBoxObj("State")<>"1" then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr = "document.write('<script language=javascript src="& TempSysRootDir &"/Ads/CreateJs/ClarityBox.js></script>'); " & vbCrLf & _
						   "document.write('<div style=""position:absolute;left:300px;top:150px;width:"& ClarityBoxObj("PicWidth") &"; height:"& ClarityBoxObj("PicHeight") &";z-index:1;solid;filter:alpha(opacity=90)"" id=ClarityBoxID onmousedown=""ClarityBox(this)"" onmousemove=""ClarityBoxMove(this)"" onMouseOut=""down=false"" onmouseup=""down=false""><table cellpadding=0 border=0 cellspacing=1 width="& ClarityBoxObj("PicWidth") &" height="& cint(ClarityBoxObj("PicHeight")+20) &" bgcolor=#000000><tr><td height=20 align=right style=""cursor:move;""><a href=# style=""font-size: 9pt; color: white; text-decoration: none"" onClick=ClarityBoxclose(""ClarityBoxID"") >>>关闭>></a></td></tr><tr><td>"&ClarityBoxStr&"</td></tr></table></div>');" & vbCrLf & _
				           "document.write('<script src="& TempSysRootDir & "/" & PlusDir &"/Ads/show.asp?Location="&ClarityBoxLocation&"></script>');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")) = false then
				MyFile.CreateFolder(Server.MapPath(TempSysRootDir&"\JS\AdsJs"))
			End If
			 if MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& ClarityBoxLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& ClarityBoxLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& ClarityBoxLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			ClarityBoxObj.close
			Set ClarityBoxObj = Nothing
		 End Sub
		 
		 Sub RightBottom(TempLocation)
		    dim RightBottomStr,RightBottomLocation,RightBottomObj
			RightBottomLocation = clng(TempLocation)
			AdsTempPicStr(RightBottomLocation)
			RightBottomStr = AdsTempStr
			Set RightBottomObj = Conn.Execute("select * from FS_Ads where Location="&RightBottomLocation&"")
			if RightBottomObj("State")<>"1" then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr = "if (navigator.appName == 'Netscape')" & vbCrLf & _
						   "{document.write('<layer id=RightBottom top=150 width="& RightBottomObj("PicWidth") &" height="& RightBottomObj("PicHeight") &">"& RightBottomStr &"</layer>');}" & vbCrLf & _
						   "else{document.write('<div id=RightBottom style=""position: absolute;width:"& RightBottomObj("PicWidth") &";height:"& RightBottomObj("PicHeight") &";visibility: visible;z-index: 1"">"& RightBottomStr &"</div>');}" & vbCrLf & _
						   "document.write('<script language=javascript src="& TempSysRootDir &"/Ads/CreateJs/RightBottom.js></script>');" & vbCrLf & _
				           "document.write('<script src="& TempSysRootDir & "/" & PlusDir &"/Ads/show.asp?Location="&RightBottomLocation&"></script>');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")) = false then
				MyFile.CreateFolder(Server.MapPath(TempSysRootDir&"\JS\AdsJs"))
			End If
			 if MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& RightBottomLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& RightBottomLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& RightBottomLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			RightBottomObj.close
			Set RightBottomObj = Nothing
		 End Sub
		 
		 Sub DriftBox(TempLocation)
		    dim DriftBoxStr,DriftBoxLocation,DriftBoxObj
			DriftBoxLocation = clng(TempLocation)
			AdsTempPicStr(DriftBoxLocation)
			DriftBoxStr = AdsTempStr
			Set DriftBoxObj = Conn.Execute("select * from FS_Ads where Location="&DriftBoxLocation&"")
			if DriftBoxObj("State")<>"1" then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr = "DriftBoxStr=(document.layers)?true:false;" & vbCrLf & _
						   "if(DriftBoxStr){document.write('<layer id=DriftBox width="& DriftBoxObj("PicWidth") &" height="& DriftBoxObj("PicHeight") &" onmouseover=DriftBoxSM(""DriftBox"") onmouseout=movechip(""DriftBox"")>"& DriftBoxStr &"</layer>');}" & vbCrLf & _
						   "else{document.write('<div id=DriftBox style=""position:absolute; width:"& DriftBoxObj("PicWidth") &"px; height:"& DriftBoxObj("PicHeight") &"px; z-index:9; filter: Alpha(Opacity=90)"" onmouseover=DriftBoxSM(""DriftBox"") onmouseout=movechip(""DriftBox"")>"& DriftBoxStr &"</div>');}" & vbCrLf & _
						   "document.write('<script language=javascript src="& TempSysRootDir &"/Ads/CreateJs/DriftBox.js></script>');" & vbCrLf & _
				           "document.write('<script src="& TempSysRootDir & "/" & PlusDir &"/Ads/show.asp?Location="&DriftBoxLocation&"></script>');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")) = false then
				MyFile.CreateFolder(Server.MapPath(TempSysRootDir&"\JS\AdsJs"))
			End If
			 if MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& DriftBoxLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& DriftBoxLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& DriftBoxLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			DriftBoxObj.close
			Set DriftBoxObj = Nothing
		 End Sub
		 
		 Sub LeftBottom(TempLocation)
		    dim LeftBottomStr,LeftBottomLocation,LeftBottomObj
			LeftBottomLocation = clng(TempLocation)
			AdsTempPicStr(LeftBottomLocation)
			LeftBottomStr = AdsTempStr
			Set LeftBottomObj = Conn.Execute("select * from FS_Ads where Location="&LeftBottomLocation&"")
			if LeftBottomObj("State")<>"1" then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr = "if (navigator.appName == 'Netscape')" & vbCrLf & _
						   "{document.write('<layer id=LeftBottom top=150 width="& LeftBottomObj("PicWidth") &" height="& LeftBottomObj("PicHeight") &">"& LeftBottomStr &"</layer>');}" & vbCrLf & _
						   "else{document.write('<div id=LeftBottom style=""position: absolute;width:"& LeftBottomObj("PicWidth") &";height:"& LeftBottomObj("PicHeight") &";visibility: visible;z-index: 1"">"& LeftBottomStr &"</div>');}" & vbCrLf & _
						   "document.write('<script language=javascript src="& TempSysRootDir &"/Ads/CreateJs/LeftBottom.js></script>');" & vbCrLf & _
				           "document.write('<script src="& TempSysRootDir & "/" & PlusDir &"/Ads/show.asp?Location="&LeftBottomLocation&"></script>');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")) = false then
				MyFile.CreateFolder(Server.MapPath(TempSysRootDir&"\JS\AdsJs"))
			End If
			 if MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& LeftBottomLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& LeftBottomLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& LeftBottomLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			LeftBottomObj.close
			Set LeftBottomObj = Nothing
		 End Sub
		 
		 Sub Couplet(TempLocation)
		    dim CoupletLeftStr,CoupletLocation,CoupletRightStr,CoupletObj
			CoupletLocation = clng(TempLocation)
			AdsTempPicStr(CoupletLocation)
			CoupletLeftStr = AdsTempStr
			CoupletRightStr = AdsTempStrRight
			Set CoupletObj = Conn.Execute("select State from FS_Ads where Location="&CoupletLocation&"")
			if CoupletObj("State")<>"1" then
				AdsJSStr = "document.write('此广告已经暂停或是失效');"
			else
				AdsJSStr =  "function winload()" & vbCrLf & _
							"{" & vbCrLf & _
							"AdsLayerLeft.style.top=20;" & vbCrLf & _
							"AdsLayerLeft.style.left=5;" & vbCrLf & _
							"AdsLayerRight.style.top=20;" & vbCrLf & _
							"AdsLayerRight.style.right=5;" & vbCrLf & _
							"}" & vbCrLf & _
							"if(screen.availWidth>800){" & vbCrLf & _
							"{" & vbCrLf & _
							"document.write('<div id=AdsLayerLeft style=""position: absolute;visibility:visible;z-index:1""><table width=100  border=0 cellspacing=0 cellpadding=0><tr><td>" & CoupletLeftStr & "</td></tr></table></div>'" & vbCrLf & _
							"+'<div id=AdsLayerRight style=""position: absolute;visibility:visible;z-index:1""><table width=100  border=0 cellspacing=0 cellpadding=0><tr><td>" & CoupletRightStr & "</td></tr></table></div>');" & vbCrLf & _
							"}" & vbCrLf & _
							"document.write('<script language=javascript src="& TempSysRootDir &"/Ads/CreateJs/Couplet.js></script>');" & vbCrLf & _
							"winload()" & vbCrLf & _
							"}" & vbCrLf & _
				           "document.write('<script src="& TempSysRootDir & "/" & PlusDir &"/Ads/show.asp?Location="&CoupletLocation&"></script>');"
			end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")) = false then
				MyFile.CreateFolder(Server.MapPath(TempSysRootDir&"\JS\AdsJs"))
			End If
			 if MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& CoupletLocation &".js") then
				MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& CoupletLocation &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& CoupletLocation &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			CoupletObj.close
			Set CoupletObj = Nothing
		 End Sub
		 
		 Sub Cycle(ALocation,TempLocation)
		     dim CycleSelfObj,CycleSelfLocation,CycleLocation,CycleObj,NoCycleObj,JsFileName
			     CycleSelfLocation = clng(ALocation)
				 CycleLocation = clng(TempLocation)
			 Set CycleSelfObj = Conn.Execute("select * from FS_Ads where Location="&CycleSelfLocation&"")'自身查询
			 if CycleSelfObj("CycleTF") = "1" then '所有循环广告
					if CycleSelfObj("CycleLocation")<>"0" then '所有被添加到循环广告的非循环广告 
						Set NoCycleObj = Conn.Execute("select * from FS_Ads where State=1 and CycleTF=1 and (Location="&CycleSelfObj("CycleLocation")&" or CycleLocation="&CycleSelfObj("CycleLocation")&") order by AddTime desc")
					    Set CycleObj = Conn.Execute("select * from FS_Ads where Location="&CycleSelfObj("CycleLocation")&"")'所属循环位
					else  '所有TYPE=11的循环广告    
						Set NoCycleObj = Conn.Execute("select * from FS_Ads where State=1 and CycleTF=1 and (Location="&CycleSelfLocation&" or CycleLocation="&CycleSelfLocation&") order by AddTime desc")
					    Set CycleObj = Conn.Execute("select * from FS_Ads where Location="&CycleSelfLocation&"")'所属循环位
					end if
			  else '所有现在不是循环广告
			        if CycleLocation <> "0" then '以前是循环广告的非循环广告 
						Set NoCycleObj = Conn.Execute("select * from FS_Ads where State=1 and CycleTF=1 and (Location="&CycleLocation&" or CycleLocation="&CycleLocation&") order by AddTime desc")
					    Set CycleObj = Conn.Execute("select * from FS_Ads where Location="&CycleLocation&"")'所属循环位
					end if			  
			 end if
			   AdsJSStr = "document.write('<marquee onmouseout=start() onmouseover=stop() width="&CycleObj("PicWidth")&" height="&CycleObj("PicHeight")&" direction="&CycleObj("CycleDirection")&" scrollamount="&CycleObj("CycleSpeed")&">"
		     do while not NoCycleObj.eof 
			 	If Instr(1,LCase(NoCycleObj("LeftPicPath")),"http://") <> 0 then
				   AdsJSStr = AdsJSStr & " <a href="""& TempSysRootDir & "/" & PlusDir &"/Ads/AdsClick.asp?Location="& NoCycleObj("Location") &""" title="""&NoCycleObj("Explain")&""" target=_blank><img src="""& NoCycleObj("LeftPicPath")&""" width="""&CycleObj("PicWidth")&""" height="""&CycleObj("PicHeight")&""" border=""0""></a>"
			    Else
				   AdsJSStr = AdsJSStr & " <a href="""& TempSysRootDir & "/" & PlusDir &"/Ads/AdsClick.asp?Location="& NoCycleObj("Location") &""" title="""&NoCycleObj("Explain")&""" target=_blank><img src="""& TempSysRootDir & NoCycleObj("LeftPicPath")&""" width="""&CycleObj("PicWidth")&""" height="""&CycleObj("PicHeight")&""" border=""0""></a>"
				End If
			   NoCycleObj.movenext
			   if CycleObj("CycleDirection") = "up" or CycleObj("CycleDirection") = "down" then 
				   AdsJSStr = AdsJSStr & "<br><br>"
				else
				   AdsJSStr = AdsJSStr & "&nbsp;&nbsp;"
				end if
			   loop
			   AdsJSStr = AdsJSStr & "</marquee>');"
			  if CycleSelfObj("State")<>"1" and CycleSelfObj("Type")="11" then
				 AdsJSStr = "document.write('此广告已经暂停或是失效');"
			  end if
			  if CycleSelfObj("Type")<>"11" then
			     JsFileName = clng(CycleSelfObj("CycleLocation"))
			  else
			     JsFileName = clng(CycleSelfLocation)
			  end if
			Set MyFile=Server.CreateObject(G_FS_FSO)
			If MyFile.FolderExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")) = false then
				MyFile.CreateFolder(Server.MapPath(TempSysRootDir&"\JS\AdsJs"))
			End If
			 if MyFile.FileExists(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& JsFileName &".js") then
				MyFile.DeleteFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& JsFileName &".js")
			 end if
			set CrHNJS=MyFile.CreateTextFile(Server.MapPath(TempSysRootDir&"\JS\AdsJs")&"\"& JsFileName &".js")
				CrHNJS.write AdsJSStr
				set MyFile=nothing
			   NoCycleObj.close
			   CycleObj.close
			   CycleSelfObj.close
			   Set CycleSelfObj = Nothing
		 End Sub
'-------------生成广告JS结束----------------  
%>

