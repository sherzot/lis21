<%
'******************************************************************************
'作成日：2014/12/17
'概　要：求人票詳細ページの求人情報と企業情報のタブ
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vType			：表示中情報の種類 ["0"]職種情報 ["1"]会社情報 ["2"]インタビュー
'　　　：vAccessCount	：表示中求人票のアクセス回数
'作成者：Lis K.Kaz
'備　考：
'使用元：しごとナビ/order/order_detail.asp
'******************************************************************************
Function DspOrderShowTypeSwitch_addon(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vType, ByVal vAccessCount)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode
	Dim sOrderType
	Dim sJobTypeDetail
	Dim sUpdateDay
	Dim dbTopInterviewFlag
	Dim dbPlanType

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sOrderCode = rRS.Collect("OrderCode")
	sOrderType = rRS.Collect("OrderType")
	dbPlanType = ChkStr(rRS.Collect("PlanTypeName"))
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	'具体的職種名
	sJobTypeDetail = rRS.Collect("JobTypeDetail")
	'更新日
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")
	'トップインタビュー
	dbTopInterviewFlag = rRS.Collect("TopInterviewFlag")

	If sJobTypeDetail <> "" Then sJobTypeDetail = sJobTypeDetail & "のお仕事情報詳細"

	Response.Write "<div id=""tab_switch"">"
	Response.Write "<div class=""left"">"

    ' 2014/12/17　画像なし版タブ化

	If vType = "0" Then
        '仕事情報を表示中の場合
		'Response.Write "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_orderdetail_on.gif"" alt=""" & sJobTypeDetail & """ border=""0"" width=""93"" height=""22""></div>"
        Response.Write "<p class=""nolink"">しごと情報</p>"


		If sOrderType = "0" Then
		'一般の求人広告の場合は会社情報へのリンクを表示
		'Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./company_order.asp?poc=" & sOrderCode & """ title=""会社情報""><img src=""/img/order/tab_companyinfo_off.gif"" alt=""会社情報"" border=""0"" width=""93"" height=""22""></a></div>"
		Response.Write "<a class=""tablink"" href=""./company_order.asp?poc=" & sOrderCode & """>企業情報</a>"
        End If

		If sOrderType = "0" And dbTopInterviewFlag = "1" Then
			Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""/order/order_interview.asp?ordercode=" & sOrderCode & """ title=""会社情報""><img src=""/img/order/tab_interview_off.gif"" alt=""インタビュー"" border=""0"" width=""93"" height=""22""></a></div>"
		End If
        
        Response.Write "<div class=""autoprintspace""><a href=""./order_detail_autoprint.asp?ordercode=" & sOrderCode & """ target=""_blank""><p class=""Lautoprint""><img src=""/img/order/order_print.png"" class=""autoprint"">印刷</p></a></div>"
                'Response.Write "<p style=""float: left;margin-right: 4px;margin-top: 2px;background-color: #ffb200;color: #FFF;font-size: 5px;display: block;width: 30px;height: 30px;-webkit-border-radius: 7px;-moz-border-radius: 7px;border-radius: 7px;""><img src=""/img/order/printer87.png"" style=""width:20px;margin:5px 10px 10px 5px;""></p>"
                'Response.Write "<p style=""float: left;margin-right: 4px;margin-top: 2px;background-color: #ffb200;color: #FFF;font-size: 5px;display: block;width: 30px;height: 30px;-webkit-border-radius: 50%;-moz-border-radius: 50%;border-radius: 50%;""><img src=""/img/order/printer87.png"" style=""width:20px;margin:5px 10px 10px 5px;""></p>"
                'Response.Write "<p style=""float: left;margin-right: 4px;margin-top: 2px;background-color: #ffb200;color: #FFF;font-size: 5px;display: block;width: 50px;height: 50px;-webkit-border-radius: 50%;-moz-border-radius: 50%;border-radius: 50%;""><img src=""/img/order/office material7 (2).png"" style=""width:30px;margin:10px 10px 10px 10px;""></p>"
	ElseIf vType = "1" Then
		'会社情報を表示中の場合
		'Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./order_detail.asp?ordercode=" & sOrderCode & """><img src=""/img/order/tab_orderdetail_off.gif"" alt=""" & sJobTypeDetail & """ border=""0"" width=""93"" height=""22""></a></div>"
        Response.Write "<a class=""tablink"" href=""./order_detail.asp?ordercode=" & sOrderCode & """>しごと情報</a>"
		If sOrderType = "0" Then
			'一般の求人広告の場合は会社情報を表示
			'Response.Write "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_companyinfo_on.gif"" alt=""会社情報"" border=""0"" width=""93"" height=""22""></div>"
		     Response.Write "<p class=""nolink"">企業情報</p>"
        End If

		If sOrderType = "0" And dbTopInterviewFlag = "1" Then
			Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""/order/order_interview.asp?ordercode=" & sOrderCode & """ title=""会社情報""><img src=""/img/order/tab_interview_off.gif"" alt=""インタビュー"" border=""0"" width=""93"" height=""22""></a></div>"
		End If

    ' 2014/12/17　vType=2は扱い不明、放置

	ElseIf vType = "2" Then
		'インタビューを表示中の場合
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./order_detail.asp?ordercode=" & sOrderCode & """><img src=""/img/order/tab_orderdetail_off.gif"" alt=""" & sJobTypeDetail & """ border=""0"" width=""93"" height=""22""></a></div>"
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./company_order.asp?poc=" & sOrderCode & """ title=""会社情報""><img src=""/img/order/tab_companyinfo_off.gif"" alt=""会社情報"" border=""0"" width=""93"" height=""22""></a></div>"
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_interview_on.gif"" alt=""会社情報"" border=""0"" width=""93"" height=""22""></div>"
	End If

	Response.Write "</div>"


	Response.Write "<br clear=""both""></div>" & vbCrLf
    Response.Write "<br clear=""both"">"
End Function

'******************************************************************************
'作成日：2014/12/17
'概　要：求人票のキャッチコピー部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：ナビ/order/order_detail.asp
'備　考：2

'******************************************************************************
Function DspOrderCatchCopy2_addon(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vAccessCount, ByVal vCategoryCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode
    Dim dbCompanyName
    Dim dbCatchCopy

	Dim sOptionNo			'大きい写真の番号
	Dim sCompanyPictureFlag	'企業写真フラグ ["1"]有 ["0"]無
	Dim sImg1,sCap1

    Dim sImg2,sImg3,sImg4,sCap2,sCap3,sCap4 'その他の3枚の画像の番号

	Dim sClass
	Dim sImgSpeciality

	Dim sUpdateDay
	Dim sPublishLimitStr
	Dim sCautionStr
	Dim flgNowPublic
	
    Dim HimlOiwai

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")
    dbCatchCopy = rRS.Collect("CatchCopy")
    'お祝い金設定
    HimlOiwai = rRS.Collect("CongratulationPrice")

		sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

        if dbOrderType = "0" Then
            dbCompanyName = rRS.Collect("CompanyName")
        else
            dbCompanyName = ""
        End if
        %>
        <div class="addcatch"><%= dbCatchCopy %></div>
        <%
	'******************************************************************************
	'大きい画像 start
	'------------------------------------------------------------------------------
	dbImageLimit = rRS.Collect("ImageLimit")
	sOptionNo = ""
	sImg1 = ""

	If dbImageLimit > 0 Then
		If dbImageLimit > 1 Then
			sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("OptionNo1")) <> "" Then
					sOptionNo = oRS.Collect("OptionNo1")
					sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & sOptionNo
				End If
			End If
		End If

		If sImg1 = "" And dbOrderType = "0" Then
			sSQL = "sp_GetDataPicture '" & dbCompanyCode & "', '1'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'大きい画像 end
	'******************************************************************************

    '******************************************************************************
	'その他3枚の画像 start
	'------------------------------------------------------------------------------
    If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo2")) <> "" Then
				sImg2 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo2")
				sCap2 = ChkStr(oRS.Collect("Caption2"))
			End If
			If ChkStr(oRS.Collect("PicNo3")) <> "" Then
				sImg3 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo3")
				sCap3 = ChkStr(oRS.Collect("Caption3"))
			End If
			If ChkStr(oRS.Collect("PicNo4")) <> "" Then
				sImg4 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo4")
				sCap4 = ChkStr(oRS.Collect("Caption4"))
			End If
		End If
		Call RSClose(oRS)

	ElseIf dbImageLimit > 1 Then
		sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', '" & vCategoryCode & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then

			If ChkStr(oRS.Collect("OptionNo2")) <> "" Then
				sImg2 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo2")
				sCap2 = ChkStr(oRS.Collect("Caption2"))
			End If
			If ChkStr(oRS.Collect("OptionNo3")) <> "" Then
				sImg3 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo3")
				sCap3 = ChkStr(oRS.Collect("Caption3"))
			End If
			If ChkStr(oRS.Collect("OptionNo4")) <> "" Then
				sImg4 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo4")
				sCap4 = ChkStr(oRS.Collect("Caption4"))
			End If
		End If
	End If
    '------------------------------------------------------------------------------
	'その他3枚の画像 end
	'******************************************************************************

	'更新日
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'求人票掲載期限 start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">※期限前に掲載終了する場合があります。</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'掲載中 or 非掲載
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'社外案件ならDspPublicLimitDayを、社内案件ならPublicLimitDayを表示
	'社外案件 OrderType = 0
	'社内案件 OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
			'ライセンス切れのときは"掲載終了"と表示
			sPublishLimitStr = "掲載終了"
			sCautionStr = ""
		Else
			sPublishLimitStr = "募集中"
		End If
	End If
	'------------------------------------------------------------------------------
	'求人票掲載期限 end
	'******************************************************************************

	'<社内案件用写真>
	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo1")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo1")
			End If
		End If
		Call RSClose(oRS)
	End If
	'</社内案件用写真>

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)


	If sImg1 <> "" Then
		Response.Write "<div id=""catchcopy"">"

		Response.Write "<div class=""main_pics"">"
		'Response.Write "<img src=""" & sImg1 & """ alt="""" id=""big_pics"">"

        '画像をサムネイル表示：gallery.js
        '<!-- デフォルト画像 -->
        Response.Write "<img src=""" & sImg1 & """ alt="""" class=""mainImage"" />"
                        Response.Write "<br style=""clear:both;"" />"
        '<!-- 表示させるテキスト -->
        Response.Write "<div class=""messageBox"">"
            Response.Write "<p id=""pict1"" ></p>"
            Response.Write "<p id=""pict2"" class=""invisible"">" & sCap2 & "</p>"
            Response.Write "<p id=""pict3"" class=""invisible"">" & sCap3 & "</p>"
            Response.Write "<p id=""pict4"" class=""invisible"">" & sCap4 & "</p>"
            Response.Write "</div>"

            

        '画像が2枚目以降存在すれば、1つずつサムネに追加していく
        If sImg2 & sImg3 & sImg4 <> "" Then
        Response.Write "<hr />"
        Response.Write "<img src=""" & sImg1 & """ alt="""" class=""thumb"" rel=""pict1"" />"
        	If sImg2 <> "" Then
        Response.Write "<img src=""" & sImg2 & """ alt=""" & sCap2 & """ class=""thumb"" rel=""pict2"" />"
        	End If
        	If sImg3 <> "" Then
        Response.Write "<img src=""" & sImg3 & """ alt=""" & sCap3 & """ class=""thumb"" rel=""pict3"" />"
            End If
            If sImg4 <> "" Then
        Response.Write "<img src=""" & sImg4 & """ alt=""" & sCap4 & """ class=""thumb"" rel=""pict4"" />"
            End If
        End If
        Response.Write "</div>"

'pic.js
'※使ってません
    if 0 = 1 then
        Response.Write "<div id=""demo"">"
	
        Response.Write "<div id=""MainBox"">"
        'ここに画像が表示される
	    Response.Write "<p class=""large-image""><img src=""" & sImg1 & """ alt=""写真1"" class=""mainImage""></p>"
	    Response.Write "<p id=""note0"" class=""viewText""></p>"
	    Response.Write "<p id=""note1"" class=""viewText invisible""></p>"
	    Response.Write "<p id=""note2"" class=""viewText invisible"">" & sCap2 & "</p>"
	    Response.Write "<p id=""note3"" class=""viewText invisible"">" & sCap3 & "</p>"
	    Response.Write "<p id=""note4"" class=""viewText invisible"">" & sCap4 & "</p>"
        Response.Write "</div>"

	        Response.Write "<ul class=""small-images clearfix"">"
		        Response.Write "<li><img src=""" & sImg1 & """ alt=""写真1"" class=""thumb"" rel=""note1""></li>"
		        Response.Write "<li><img src=""" & sImg2 & """ alt=""写真2"" class=""thumb"" rel=""note2""></li>"
		        Response.Write "<li><img src=""" & sImg3 & """ alt=""写真3"" class=""thumb"" rel=""note3""></li>"
		        Response.Write "<li><img src=""" & sImg4 & """ alt=""写真4"" class=""thumb"" rel=""note4""></li>"
	        Response.Write "</ul>"

        Response.Write "</div>"
    End if
        %>

            

        	<div id="lissapo">
            <a href="#detail_waku" class="deju">募集要項を見る<span style="float:right;color:#FFF;">>></span><br style="clear:both;"></a>

            <div class="circ1">
                            <img src="/img/neo/webjobad.png">
                            <br style="clear:both;">
            </div>

                <table style="table-layout:fixed;width:100%;margin:4px 0;"><tbody>
			            <tr><td class="cal1_a">掲載期限</td>
                        <td class="cal2_a"><%= sPublishLimitStr %></td></tr>

                        <tr><td class="cal1_a">更新日</td>
                        <td class="cal2_a"><%= sUpdateDay %></td></tr>

                        <tr><td class="cal1_b">情報コード</td>
                        <td class="cal2_b"><%= dbOrderCode %></td></tr>
                        <tr><td colspan=2 class="cal_long"><img src="/img/order/cau2.png" /></td></tr>

            </tbody></table>
			</div>

           <br clear="all">



           <% If G_USERTYPE = "" Then %> 
            <div id="top_reg_button">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="履歴書登録して応募" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="ログインして応募" border="0">
            </a>
			</div>
            <% End If %> 
		<%	

		Response.Write "<br clear=""all"">"
		Response.Write "</div>"

	Else
		Response.Write "<div id=""catchcopy"">"
			%>
            
            <div class="main_pics">
            <table><tbody><tr ><td colspan="2">
            
                                <a href="#detail_waku" class="deju2">募集要項を見る<span style="float:right;color:#FFF;">>></span><br style="clear:both;"></a></td></tr>


                    <tr><td style="width:50%;">
			<div class="circ2_o">
                            <img src="/img/neo/webjobad.png" />
			</div>
                            </td>
                    </tr>


            </tbody></table>
            </div>



            <div id="lissapo">
                <table style="table-layout:fixed;width:100%;margin:4px 0;"><tbody>
			            <tr><td class="cal1_a">掲載期限</td>
                        <td class="cal2_a"><%= sPublishLimitStr %></td></tr>

                        <tr><td class="cal1_a">更新日</td>
                        <td class="cal2_a"><%= sUpdateDay %></td></tr>

                        <tr><td class="cal1_b">情報コード</td>
                        <td class="cal2_b"><%= dbOrderCode %></td></tr>
                        <tr><td colspan="2" class="cal_long"><img src="/img/order/cau2.png" /></td></tr>

                </tbody></table>
			</div>
            </div>

           <br clear="all">

		   <% If G_USERTYPE = "" Then %> 
			
<div class="center">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="履歴書登録して応募" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="ログインして応募" border="0">
            </a>
			</div>

		<% End If 

	End If

    '※使ってません
    if 0 = 1 then
        'さらに他の画像が存在するなら
        If sImg2 & sImg3 & sImg4 <> "" Then
		Response.Write "<div id=""sub_pics"">"
		Response.Write "<div class=""auto"">"

            If sImg1 <> "" Then
			Response.Write "<div class=""sub_pics sub_pics1""><img src=""" & sImg1 & """>"
            Response.Write "</div>"
		    End If

		    If sImg2 <> "" Then
			Response.Write "<div class=""sub_pics sub_pics1""><img src=""" & sImg2 & """ alt=""" & sCap2 & """>"
			Response.Write "<p class=""m0"" align=""left"">" & sCap2 & "</p>"
            Response.Write "</div>"
		    End If

		    If sImg3 <> "" Then
			Response.Write "<div class=""sub_pics sub_pics2""><img src=""" & sImg3 & """ alt=""" & sCap3 & """>"
			Response.Write "<p class=""m0"" align=""left"">" & sCap3 & "</p>"
			Response.Write "</div>"
		    End If

		    If sImg4 <> "" Then
			Response.Write "<div class=""sub_pics sub_pics3""><img src=""" & sImg4 & """ alt=""" & sCap4 & """>"
			Response.Write "<p class=""m0"" align=""left"">" & sCap4 & "</p>"
			Response.Write "</div>"
		    End If

            Response.Write "<br clear=""all"">"
		    Response.Write "</div>"
		    Response.Write "</div>"
        End If
   End If

End Function

'******************************************************************************
'作成日：2014/12/17
'概　要：求人票のタイトル部分を出力（募集職種、キャッチコピー、求人特徴）
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：ナビ/order/order_detail.asp
'備　考：2

'******************************************************************************
Function DspOrderTitle_addon(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vAccessCount)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode
    Dim dbCompanyName

	Dim sOptionNo			'大きい写真の番号
	Dim sCompanyPictureFlag	'企業写真フラグ ["1"]有 ["0"]無
	Dim sImg1
	Dim sClass
	Dim sImgSpeciality

	Dim sUpdateDay
	Dim sPublishLimitStr
	Dim sCautionStr
	Dim flgNowPublic
	
	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")

		sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

        if dbOrderType = "0" Then
            dbCompanyName = rRS.Collect("CompanyName")
        else
            dbCompanyName = ""
        End if

	'******************************************************************************
	'大きい画像 start
	'------------------------------------------------------------------------------
	dbImageLimit = rRS.Collect("ImageLimit")
	sOptionNo = ""
	sImg1 = ""
	If dbImageLimit > 0 Then
		If dbImageLimit > 1 Then
			sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("OptionNo1")) <> "" Then
					sOptionNo = oRS.Collect("OptionNo1")
					sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & sOptionNo
				End If
			End If
		End If

		If sImg1 = "" And dbOrderType = "0" Then
			sSQL = "sp_GetDataPicture '" & dbCompanyCode & "', '1'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'大きい画像 end
	'******************************************************************************

	'更新日
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'求人票掲載期限 start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">※期限前に掲載終了する場合があります。</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'掲載中 or 非掲載
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'社外案件ならDspPublicLimitDayを、社内案件ならPublicLimitDayを表示
	'社外案件 OrderType = 0
	'社内案件 OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
			'ライセンス切れのときは"掲載終了"と表示
			sPublishLimitStr = "掲載終了"
			sCautionStr = ""
		Else
			sPublishLimitStr = "常時募集中"
		End If
	End If
	'------------------------------------------------------------------------------
	'求人票掲載期限 end
	'******************************************************************************

	'<社内案件用写真>ここは不要だが企業ロゴ出力時に応用したいので据え置き
	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo1")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo1")
			End If
		End If
		Call RSClose(oRS)
	End If
	'</社内案件用写真>

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)

    'タイトル出力部（募集職種とキャッチコピー）
    Response.Write "<div class=""titleon"">"
        '求人広告なら企業名を表示するが、それ以外なら伏せる
        If dbOrderType = "0" Then
             Response.Write "<p class=""comadd""><a href=""./company_order.asp?poc=" & sOrderCode & """ class=""c_c"">" & rRS.Collect("CompanyName") & "</a></p>"
             Response.Write "<p class=""jobadd"">" & rRS.Collect("JobTypeDetail") & "</p>"
        Else
            Response.Write "<p class=""comadd"">" & rRS.Collect("CompanySpeciality") & "</p>"
            Response.Write "<p class=""jobadd"">" & rRS.Collect("JobTypeDetail") & "</p>"
        End If
        '求人の特徴があれば表示
		If sImgSpeciality <> "" Then
			Response.Write "<div class=""ordersp_a"">"
            Response.Write "<div class=""ordersp_b"">"
			'Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">【募集の特徴】</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
            Response.Write "</div>"
		End If
    Response.Write "</div>"

End Function

'******************************************************************************
'概　要：求人票のキャッチコピー部分を出力
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：ナビ/order/order_detail.asp
'備　考：リスサポート案件用

'******************************************************************************
Function DspOrderCatchCopy3_addon(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vAccessCount, ByVal vCategoryCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbImageLimit
	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode
    Dim dbCatchCopy		

	Dim sOptionNo			'大きい写真の番号
	Dim sCompanyPictureFlag		'企業写真フラグ ["1"]有 ["0"]無
	Dim sImg1,sCap1

    Dim sImg2,sImg3,sImg4,sCap2,sCap3,sCap4 'その他の3枚の画像の番号

	Dim sClass
	Dim sImgSpeciality

	Dim sUpdateDay
	Dim sPublishLimitStr
	Dim sCautionStr
	Dim flgNowPublic
	Dim dbCompanyName 'リス自体の求人に使う会社名

	Dim HimlOiwai

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")
    'キャッチコピー設定
	dbCatchCopy = rRS.Collect("CatchCopy")
	'お祝い金設定
    HimlOiwai = rRS.Collect("CongratulationPrice")

		sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")
	
	'自社求人用、使わないけど残す
        if dbCompanyCode = "C0001533" Then
            dbCompanyName = rRS.Collect("CompanyName")
            %>
            <div id="c_name"><%= dbCompanyName %></div>
            <%
        else
            dbCompanyName = ""
        End if

	    %>
        <div class="addcatch"><%= dbCatchCopy %></div>
        <%
	'******************************************************************************
	'大きい画像 start
	'------------------------------------------------------------------------------
	dbImageLimit = rRS.Collect("ImageLimit")
	sOptionNo = ""
	sImg1 = ""
	If dbImageLimit > 0 Then
		If dbImageLimit > 1 Then
			sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', 'orderpicture'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				If ChkStr(oRS.Collect("OptionNo1")) <> "" Then
					sOptionNo = oRS.Collect("OptionNo1")
					sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & sOptionNo
				End If
			End If
		End If

		If sImg1 = "" And dbOrderType = "0" Then
			sSQL = "sp_GetDataPicture '" & dbCompanyCode & "', '1'"
			flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
			If GetRSState(oRS) = True Then
				sImg1 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=1"
			End If
		End If
	End If
	'------------------------------------------------------------------------------
	'大きい画像 end
	'******************************************************************************

	'******************************************************************************
	'その他3枚の画像 start
	'------------------------------------------------------------------------------
    If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo2")) <> "" Then
				sImg2 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo2")
				sCap2 = ChkStr(oRS.Collect("Caption2"))
			End If
			If ChkStr(oRS.Collect("PicNo3")) <> "" Then
				sImg3 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo3")
				sCap3 = ChkStr(oRS.Collect("Caption3"))
			End If
			If ChkStr(oRS.Collect("PicNo4")) <> "" Then
				sImg4 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo4")
				sCap4 = ChkStr(oRS.Collect("Caption4"))
			End If
		End If
		Call RSClose(oRS)

	ElseIf dbImageLimit > 1 Then
		sSQL = "up_GetListOrderPictureNow '" & dbCompanyCode & "', '" & dbOrderCode & "', '" & vCategoryCode & "'"
		flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
		If GetRSState(oRS) = True Then

			If ChkStr(oRS.Collect("OptionNo2")) <> "" Then
				sImg2 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo2")
				sCap2 = ChkStr(oRS.Collect("Caption2"))
			End If
			If ChkStr(oRS.Collect("OptionNo3")) <> "" Then
				sImg3 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo3")
				sCap3 = ChkStr(oRS.Collect("Caption3"))
			End If
			If ChkStr(oRS.Collect("OptionNo4")) <> "" Then
				sImg4 = "/company/imgdsp.asp?companycode=" & dbCompanyCode & "&amp;optionno=" & oRS.Collect("OptionNo4")
				sCap4 = ChkStr(oRS.Collect("Caption4"))
			End If
		End If
	End If
    '------------------------------------------------------------------------------
	'その他3枚の画像 end
	'******************************************************************************

	'更新日
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'求人票掲載期限 start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">※期限前に掲載終了する場合があります。</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'掲載中 or 非掲載
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'社外案件ならDspPublicLimitDayを、社内案件ならPublicLimitDayを表示
	'社外案件 OrderType = 0
	'社内案件 OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
			'ライセンス切れのときは"掲載終了"と表示
			sPublishLimitStr = "掲載終了"
			sCautionStr = ""
		Else
			sPublishLimitStr = "常時募集中"
		End If
	End If

    '<無期限延長機能対応>
    '2016/04/01 池田改修
    If sPublishLimitStr = "9999/12/31" Then
        '無期限の場合は、掲載期限に月末を指定。更新日に月初を指定。
        sPublishLimitStr = DateSerial(Year(Date()), Month(Date()) + 1, 0)
        sUpdateDay       = DateSerial(Year(Date()), Month(Date()), 1)
    End If
    '</無期限延長機能対応>

	'------------------------------------------------------------------------------
	'求人票掲載期限 end
	'******************************************************************************

	'<社内案件用写真>
	If dbOrderType <> "0" Then
		sSQL = "EXEC up_DtlC_PictureLIS '" & dbOrderCode & "';"
		flgQE = QUERYEXE(dbconn,oRS,sSQL,sError)
		If GetRSState(oRS) = True Then
			If ChkStr(oRS.Collect("PicNo1")) <> "" Then
				sImg1 = "/img/order/lisimgdsp.asp?companycode=" & dbCompanyCode & "&amp;picno=" & oRS.Collect("PicNo1")
			End If
		End If
		Call RSClose(oRS)
	End If
	'</社内案件用写真>

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)


	If sImg1 <> "" Then
		Response.Write "<div id=""catchcopy"">"

		Response.Write "<div class=""main_pics"">"
		'Response.Write "<img src=""" & sImg1 & """ alt="""" id=""big_pics"">"

        '画像をサムネイル表示：gallery.js
        '<!-- デフォルト画像 -->
        Response.Write "<img src=""" & sImg1 & """ alt="""" class=""mainImage"" />"
                        Response.Write "<br style=""clear:both;"" />"
        '<!-- 表示させるテキスト -->
        Response.Write "<div class=""messageBox"">"
            Response.Write "<p id=""pict1"" ></p>"
            Response.Write "<p id=""pict2"" class=""invisible"">" & sCap2 & "</p>"
            Response.Write "<p id=""pict3"" class=""invisible"">" & sCap3 & "</p>"
            Response.Write "<p id=""pict4"" class=""invisible"">" & sCap4 & "</p>"
            Response.Write "</div>"

            

        '画像が2枚目以降存在すれば、1つずつサムネに追加していく
        If sImg2 & sImg3 & sImg4 <> "" Then
        Response.Write "<hr />"
        Response.Write "<img src=""" & sImg1 & """ width=""40"" height=""40"" alt="""" class=""thumb"" rel=""pict1"" />"
        	If sImg2 <> "" Then
        Response.Write "<img src=""" & sImg2 & """ width=""40"" height=""40"" alt=""" & sCap2 & """ class=""thumb"" rel=""pict2"" />"
        	End If
        	If sImg3 <> "" Then
        Response.Write "<img src=""" & sImg3 & """ width=""40"" height=""40"" alt=""" & sCap3 & """ class=""thumb"" rel=""pict3"" />"
            End If
            If sImg4 <> "" Then
        Response.Write "<img src=""" & sImg4 & """ width=""40"" height=""40"" alt=""" & sCap4 & """ class=""thumb"" rel=""pict4"" />"
            End If
        End If
        Response.Write "</div>"


		%>
			<div id="lissapo">
                    <a href="#detail_waku" class="deju">募集要項を見る<span style="float:right;color:#FFF;">>></span><br style="clear:both;"></a>
                    <!-- 2014/12/19 -->
<!--                    <div class="cl-effect-13" style="text-align:center;"><a href="#detail_waku" style="background-color:#ffb200;display:block;">募集要項を見る</a></div>
                    -->
                    <div class="circ1">
                    <% If sOrderType = "2" Then %> 	
					<img src="/img/neo/support_shokai.png">
                    <% ElseIf sOrderType = "1" or "3" Then %>
					<img src="/img/neo/support_haken.png">
                    <% End If %> 
					<a href="<%= HTTPS_CURRENTURL %>point/pr/"><img src="/img/neo/oiwai_point.png"></a>
			        </div>


                <table style="table-layout:fixed;width:100%;margin:4px 0;"><tbody>
			            <tr><td class="cal1_a">掲載期限</td>
                        <td class="cal2_a"><%= sPublishLimitStr %></td></tr>

                        <tr><td class="cal1_a">更新日</td>
                        <td class="cal2_a"><%= sUpdateDay %></td></tr>

                        <tr><td class="cal1_b">情報コード</td>
                        <td class="cal2_b"><%= dbOrderCode %></td></tr>
                        <tr><td colspan="2" class="cal_long"><img src="/img/order/cau0.png" /></td></tr>
            </tbody></table>
			</div>

           <br clear="all">


           <% If G_USERTYPE = "" Then %> 
            <div id="top_reg_button">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="履歴書登録して応募" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="ログインして応募" border="0">
            </a>

                <!-- 2016/04/14 池田改修 -->
                <% If True Then %>

                    <!-- ログインしてない場合 -->
                    <a href="#" onclick="window.open('<%= HTTPS_NAVI_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=staff/mailtocompany.asp?ordercode=<%= sOrderCode %>', '_blank');return false;">
                    <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/conBtn.png" alt="求人への問合せ" border="0">
                    </a>

                <% Else %>

                    <!-- ログイン済みの場合 -->
                    <a href="#" onclick="window.open('<%= HTTPS_NAVI_CURRENTURL %>staff/mailtocompany.asp?ordercode=<%= sOrderCode %>','_blank');return false;">
                    <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/conBtn.png" alt="求人への問合せ" border="0">
                    </a>
                <% End If %>

			</div>
            <% End If %>
		

		<br clear="all">
		</div>
       
    <%
	Else
		Response.Write "<div id=""catchcopy"">"
		
			%>
            <div class="main_pics">
                    <a href="#detail_waku" class="deju2">募集要項を見る<span style="float:right;color:#FFF;">>></span><br style="clear:both;"></a>
                    <div class="circ2" style="display:block !important;">

                    <% If sOrderType = "2" Then %> 	
					<img src="/img/neo/support_shokai.png">
                    <% ElseIf sOrderType = "1" or "3" Then %>
					<img src="/img/neo/support_haken.png">
                    <% End If %> 
					<a href="<%= HTTPS_CURRENTURL %>point/pr/"><img src="/img/neo/oiwai_point.png"></a>

			        </div>
            </div>

            <div id="lissapo">
                <table style="table-layout:fixed;width:100%;margin:4px 0;"><tbody>
			            <tr><td class="cal1_a">掲載期限</td>
                        <td class="cal2_a"><%= sPublishLimitStr %></td></tr>

                        <tr><td class="cal1_a">更新日</td>
                        <td class="cal2_a"><%= sUpdateDay %></td></tr>

                        <tr><td class="cal1_b">情報コード</td>
                        <td class="cal2_b"><%= dbOrderCode %></td></tr>
                        <tr><td colspan="2" class="cal_long"><img src="/img/order/cau0.png" /></td></tr>

                </tbody></table>
			</div>
            </div>

           <br clear="all">


		<%
		  If G_USERTYPE = "" Then  %>
			
<div class="center">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="履歴書登録して応募" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="ログインして応募" border="0">
            </a>
			</div>
			
		<% End If 

	End If
End Function

'******************************************************************************
'作成日：2014/12/17
'概　要：求人票詳細ページの求人情報と企業情報のタブ
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vType			：表示中情報の種類 ["0"]職種情報 ["1"]会社情報 ["2"]インタビュー
'　　　：vAccessCount	：表示中求人票のアクセス回数
'作成者：Lis K.Kaz
'備　考：
'使用元：しごとナビ/order/order_detail.asp
'******************************************************************************
Function DspOrderShowTypeSwitchComapny_addon(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vType, ByVal vAccessCount)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderCode
	Dim sOrderType
	Dim sJobTypeDetail
	Dim sUpdateDay
	Dim dbTopInterviewFlag
	Dim dbPlanType

	If GetRSState(rRS) = False Then Exit Function

	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sOrderCode = rRS.Collect("OrderCode")
	sOrderType = rRS.Collect("OrderType")
	dbPlanType = ChkStr(rRS.Collect("PlanTypeName"))
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

	'具体的職種名
	sJobTypeDetail = rRS.Collect("JobTypeDetail")
	'更新日
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")
	'トップインタビュー
	dbTopInterviewFlag = rRS.Collect("TopInterviewFlag")

	If sJobTypeDetail <> "" Then sJobTypeDetail = sJobTypeDetail & "のお仕事情報詳細"

	Response.Write "<div id=""tab_switch"">"
	Response.Write "<div class=""left"">"

    ' 2014/12/17　画像なし版タブ化


		'会社情報を表示中の場合
		'Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./order_detail.asp?ordercode=" & sOrderCode & """><img src=""/img/order/tab_orderdetail_off.gif"" alt=""" & sJobTypeDetail & """ border=""0"" width=""93"" height=""22""></a></div>"
        Response.Write "<a class=""tablink_comp"" href=""./order_detail.asp?ordercode=" & sOrderCode & """>しごと情報</a>"
		If sOrderType = "0" Then
			'一般の求人広告の場合は会社情報を表示
			'Response.Write "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_companyinfo_on.gif"" alt=""会社情報"" border=""0"" width=""93"" height=""22""></div>"
		     Response.Write "<p class=""nolink_comp"" style=""padding:10px 16px 4px 16px;"">企業情報</p>"
        End If

		If sOrderType = "0" And dbTopInterviewFlag = "1" Then
			Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""/order/order_interview.asp?ordercode=" & sOrderCode & """ title=""会社情報""><img src=""/img/order/tab_interview_off.gif"" alt=""インタビュー"" border=""0"" width=""93"" height=""22""></a></div>"
		End If

	Response.Write "</div>"


	Response.Write "<br clear=""both""></div>" & vbCrLf
    Response.Write "<br clear=""both"">"
End Function

'******************************************************************************
'概　要：求人票のスクロール量に応じたナビゲーションバー
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_DtlOrderで生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'使　用：ナビ/order/order_detail.asp
'******************************************************************************
Function DspOrderScrollNavi_nov(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vAccessCount, ByVal vMyOrder, ByVal vJobTypeLimitFlag)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	Dim sOrderType

	Dim dbOrderCode
	Dim dbOrderType
	Dim dbCompanyCode
    Dim dbCompanyName
    Dim sImgSpeciality

    Dim sUpdateDay
	Dim sPublishLimitStr
	Dim sCautionStr
	Dim flgNowPublic

        '<ログイン後ボタン表示用>
        Dim sPermitFlag			'掲載許可フラグ
	    Dim sPublicFlag			'掲載フラグ
	    Dim sRiyoFlag			'掲載開始日
	    Dim sHakouFlag			'利用開始日（ライセンス発効日）
        Dim flgAddWatchList
	    Dim iMailTemplateCnt	'メールテンプレートの件数
        '</ログイン後ボタン表示用>

    If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")
    sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)

    sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

    '<ログイン後ボタン表示用>
	'******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	sPermitFlag = rRS.Collect("PermitFlag")
	sPublicFlag = rRS.Collect("PublicFlag")
	sRiyoFlag = rRS.Collect("RiyoFlag")
	sHakouFlag = rRS.Collect("HakouFlag")
	iMailTemplateCnt = rRS.Collect("MailTemplateCnt")
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************
    
    '******************************************************************************
	'企業コード start
	'------------------------------------------------------------------------------
	flgAddWatchList = False
	sSQL = "EXEC up_ChkWatchListExists_Staff '" & vUserID & "', '" & sOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("ExistsFlag") = "1" Then flgAddWatchList = True
	End If
	Call RSClose(oRS)
	'------------------------------------------------------------------------------
	'企業コード end
	'******************************************************************************

    Dim qsOrderCode				'オーダーコード(受注表番号)
	Dim iDetail				'求人票詳細からのフラグ
	
	qsOrderCode = GetForm("ordercode", 2)
	iDetail = GetForm("Detail", 2)

    '</ログイン後ボタン表示用>


    '******************************************************************************
	'求人票掲載期限 start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">※期限前に掲載終了する場合があります。</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'掲載中 or 非掲載
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'社外案件ならDspPublicLimitDayを、社内案件ならPublicLimitDayを表示
	'社外案件 OrderType = 0
	'社内案件 OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
			'ライセンス切れのときは"掲載終了"と表示
			sPublishLimitStr = "掲載終了"
			sCautionStr = ""
		Else
			sPublishLimitStr = "常時募集中"
		End If
	End If
	'------------------------------------------------------------------------------
	'求人票掲載期限 end
	'******************************************************************************
   

    Response.Write "<nav class=""scr""><p class=""navtitle"">"& rRS.Collect("JobTypeDetail") &"</p>"

    %>

    <% '会員登録、またはログインをしていない場合のボタン（会員登録してから応募、ログインしてから応募）
    If G_USERTYPE = "" Then
    %>

    <div class="navibutton">
    <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>" class="button_reg_n">会員登録して応募</a>
    <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>"  class="button_rec_n">ログインして応募</a>
    </div>
    <% End If %>

        <% '会員登録をしている場合のボタン（応募とお気に入り追加）
        If vUserType = "staff" Then
        %>

        <div class="navibutton">
                    <% If rRS.Collect("NowPublicFlag") = "1" Then %>
                    <% If flgAddWatchList = True Then %>

                        <span class="kentozumi2_n">お気に入りに追加済み</span>

                        <% Else
				        response.write "<form id=""frmSendMailJobOfferAddress"" name=""frmSendMailJobOfferAddress"" method=""post"" action=""../staff/watchlist_register.asp"" style=""width:70%;float:left"" onSubmit=""return Submit();"">"
				        'Response.Write "<a href=""#"" onclick=""document.forms.frmSendMailJobOfferAddress.submit();return false;"" class=""kento2_n"">お気に入りに追加</a>"
                        Response.Write "<a href=""#"" onclick=""document.frmSendMailJobOfferAddress.submit();return false;"" class=""kento2_n"">お気に入りに追加</a>"
				        response.write "<input type=""hidden"" name=""CONF_OrderCode"" value='"& qsOrderCode &"'>"
				        'response.write "<input type=""text"" name=""dummy"" style=""display:none;"" />"
				        'response.write "<input type=""text2"" name=""dummy2"" style=""display:none;"" />"
				        response.write "</form>"
                        End If %>

                    <% If dbOrderType = "0" Then %>
                        <a href="#" onclick="contactCompanyAdv('');return false;" class="button_obo_n">
                        この求人に応募する
                        </a>

                    <% Else %>
                        <a href="#" onclick="contactCompanyLis('');return false;" class="button_obo_n">
                        この求人に応募する
                        </a>
                        <%
                        'Response.Write "<a href=""#"" onclick=""contactCompany('1');return false;"" class=""button_que_n"">"
                        'Response.Write "この求人について質問する"
                        'Response.Write "</a>"
                        %>
                    <% End If %>
            

                        
                    <% Else %>
			        <li id="finKokoku_n"><div class="description" align="center"><b>この求人票は掲載が終了しています。メール送信はできません。</b></div></li>
		            <% End If %>

        <% End If %>
    
    
    
    <a href="#waku_<%= HimlOiwai %>" class="button_det_n">求人詳細</a>
    <!-- <li><a href="../index.asp">home</a></li> -->
    </div>
    <%
    Response.Write "</nav>"


End Function

'******************************************************************************
'概　要：印刷ログの書き込み
'引　数：rDB			：接続中のDBConnection
'　　　：rRS			：up_SearchOrder or 求人票詳細検索SQL で生成されたレコードセットオブジェクト
'　　　：vUserType		：利用中ユーザのユーザ種類 [Session("usertype")]
'　　　：vUserID		：利用中ユーザのユーザID [Session("userid")]
'　　　：vOrderCode		：閲覧中求人票
'備　考：
'使用元：order/order_detail_autoprint.asp
'履　歴：2015/08/06 Kimura func_order.aspからほぼ流用
'******************************************************************************
Function AutoPrintHistoryOrder(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	'社内からのアクセスと、たたろうさん(S0018066)からのアクセスはログに残さない
	'If IsRE(G_IPADDRESS, "^192.168.", True) = False And vUserID <> "S0018066" Then
    If vUserID <> "" Then
		If vUserType = "staff" Then
			sSQL = "up_Reg_LOG_AutoPrintHistoryOrder '" & vOrderCode & "', '" & vUserID & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
		ElseIf IsRE(Request.Cookies("id_memory"), "^S\d\d\d\d\d\d\d$", True) = True Then
			sSQL = "up_Reg_LOG_AutoPrintHistoryOrder '" & vOrderCode & "', '" & Request.Cookies("id_memory") & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
		ElseIf IsRE(GetForm("uc",2), "^S\d\d\d\d\d\d\d$", True) = True Then
			sSQL = "up_Reg_LOG_AutoPrintHistoryOrder '" & vOrderCode & "', '" & GetForm("uc",2) & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
			sSQL = "update P_Userinfo set lastaccessday = getdate() where staffcode = '" & GetForm("uc",2) & "'"
			flgQE = QUERYEXE(dbconn, oRS, sSQL, sError)
			Call RSClose(oRS)
		End If
	End If
End Function

%>