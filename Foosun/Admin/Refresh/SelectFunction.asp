<%
Function GetLableContent(LableStr)
	Dim ParaArray
	ParaArray = Split(LableStr,",")
	if UBound(ParaArray) = 0 then
		GetLableContent = ""
		Exit Function
	else
		Select Case LCase(ParaArray(0))
'******************************
'根据标签调用ypren()方法
'author:lino
'Start
'*****************************
Case "ypren"
  If UBound(ParaArray) = 1 then
    GetLableContent = ypren()
  Else
    Exit Function
  End if 
  
'**************************
'End
'**************************


			Case "freelable"
				if UBound(ParaArray) = 6 then
					GetLableContent = FreeLable(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "selfclass"
				if UBound(ParaArray) = 16 then
					GetLableContent = SelfClass(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11),ParaArray(12),ParaArray(13),ParaArray(14),ParaArray(15),ParaArray(16))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "childclasslist"
				if UBound(ParaArray) = 18 then
					GetLableContent = ChildClassList(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11),ParaArray(12),ParaArray(13),ParaArray(14),ParaArray(15),ParaArray(16),ParaArray(17),ParaArray(18))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "specialnewslist"
				if UBound(ParaArray) = 15 then
					GetLableContent = SpecialNewsList(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11),ParaArray(12),ParaArray(13),ParaArray(14),ParaArray(15))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "specialnavi"
				if UBound(ParaArray) = 7 then
					GetLableContent = specialnavi(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "location"
				if UBound(ParaArray) = 4 then
					GetLableContent = Location(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "locationnavi"
				if UBound(ParaArray) = 7 then
					GetLableContent = LocationNavi(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "classnavi"
				if UBound(ParaArray) = 6 then
					GetLableContent = ClassNavi(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "hotnews"
				if UBound(ParaArray) = 11 then
					GetLableContent = HotNews(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "lastnews"
				if UBound(ParaArray) = 11 then
					GetLableContent = LastNews(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "recnews"
				if UBound(ParaArray) = 11 then
					GetLableContent = RecNews(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "marqueenews"
				if UBound(ParaArray) = 10 then
					GetLableContent = MarqueeNews(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "classnewslist"
				if UBound(ParaArray) = 14 then
					GetLableContent = ClassNewsList(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11),ParaArray(12),ParaArray(13),ParaArray(14))
				else
					GetLableContent = Array("")
					Exit Function
				end if
			Case "search"
				GetLableContent = Search
			Case "advancedsearch"
				GetLableContent = AdvancedSearch
			Case "uselogin"
				GetLableContent = UseLogin
			Case "infostat"
				if UBound(ParaArray) = 3 then
					GetLableContent = InfoStat(ParaArray(1),ParaArray(2),ParaArray(3))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "downinfostat"
				if UBound(ParaArray) = 3 then
					GetLableContent = DownInfoStat(ParaArray(1),ParaArray(2),ParaArray(3))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "friendlink"
				if UBound(ParaArray) = 8 then
					GetLableContent = FriendLink(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "navireadnews"
				if UBound(ParaArray) = 13 then
					GetLableContent = NaviReadNews(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11),ParaArray(12),ParaArray(13))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "relatenews"
				if UBound(ParaArray) = 7 then
					GetLableContent = RelateNews(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "relatespecialnews"
				if UBound(ParaArray) = 7 then
					GetLableContent = RelateSpecialNews(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "picnews"
				if UBound(ParaArray) = 10 then
					GetLableContent = PicNews(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "filternews"
				if UBound(ParaArray) = 9 then
					GetLableContent = FilterNews(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "sitemap"
				if UBound(ParaArray) = 4 then
					GetLableContent = SiteMap(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "pagetitle"
				if UBound(ParaArray) = 1 then
					GetLableContent = PageTitle(ParaArray(1))
				else
					GetLableContent = ""
					Exit Function
				end if
			Case "copyrightstr"
				GetLableContent = CopyRightStr
			Case "focuspic"
				If UBound(ParaArray) = 11 then
					GetLableContent = FocusPic(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11))
				Else
					GetLableContent = ""
					Exit Function
				End If
			Case "recpic"
				If UBound(ParaArray) = 11 then
					GetLableContent = RecPic(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11))
				Else
					GetLableContent = ""
					Exit Function
				End If
			Case "classicalnews"
				If UBound(ParaArray) = 11 then
					GetLableContent = ClassicalNews(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11))
				Else
					GetLableContent = ""
					Exit Function
				End If
			Case "classicalpic"
				If UBound(ParaArray) = 11 then
					GetLableContent = ClassicalPic(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11))
				Else
					GetLableContent = ""
					Exit Function
				End If
			Case "lastclasspic"
				If UBound(ParaArray) = 10 then
					GetLableContent = LastClassPic(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10))
				Else
					GetLableContent = Array("")
					Exit Function
				End If
			Case "classdownload"
				If UBound(ParaArray) = 16 then
					GetLableContent = ClassDownLoad(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11),ParaArray(12),ParaArray(13),ParaArray(14),ParaArray(15),ParaArray(16))
				Else
					GetLableContent = ""
					Exit Function
				End If
			Case "downloadlist"
				If UBound(ParaArray) = 12 then
					GetLableContent = DownLoadList(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11),ParaArray(12))
				Else
					GetLableContent = Array("")
					Exit Function
				End If
			Case "lastdownlist"
				If UBound(ParaArray) = 11 then
					GetLableContent = LastDownList(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11))
				Else
					GetLableContent = ""
					Exit Function
				End If
			Case "recdownlist"
				If UBound(ParaArray) = 11 then
					GetLableContent = RecDownList(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11))
				Else
					GetLableContent = ""
					Exit Function
				End If
			Case "hotdownlist"
				If UBound(ParaArray) = 11 then
					GetLableContent = HotDownList(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11))
				Else
					GetLableContent = ""
					Exit Function
				End If
			Case "todaynews"
				If UBound(ParaArray) = 11 then
					GetLableContent = TodayNews(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11))
				Else
					GetLableContent = ""
					Exit Function
				End If
			Case "speciallastnewslist"
				If UBound(ParaArray) = 13 then
					GetLableContent = SpecialLastNewsList(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11),ParaArray(12),ParaArray(13))
				Else
					GetLableContent = Array("")
					Exit Function
				End If
			Case "download_pic"
				If UBound(ParaArray) = 2 then
					GetLableContent = DownLoad_Pic(ParaArray(1),ParaArray(2))
				Else
					GetLableContent = ""
					Exit Function
				End If
			Case "lablefile"
				If UBound(ParaArray) = 12 then
					GetLableContent = LableFile(ParaArray(1),ParaArray(2),ParaArray(3),ParaArray(4),ParaArray(5),ParaArray(6),ParaArray(7),ParaArray(8),ParaArray(9),ParaArray(10),ParaArray(11),ParaArray(12))
				Else
					Exit Function
				End If
			Case Else
				GetLableContent = ""
				Exit Function
		End Select
	end if
End Function
%>