<%
'******************************************************************************
'�쐬���F2014/12/17
'�T�@�v�F���l�[�ڍ׃y�[�W�̋��l���Ɗ�Ə��̃^�u
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvType			�F�\�������̎�� ["0"]�E���� ["1"]��Џ�� ["2"]�C���^�r���[
'�@�@�@�FvAccessCount	�F�\�������l�[�̃A�N�Z�X��
'�쐬�ҁFLis K.Kaz
'���@�l�F
'�g�p���F�����ƃi�r/order/order_detail.asp
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
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	sOrderCode = rRS.Collect("OrderCode")
	sOrderType = rRS.Collect("OrderType")
	dbPlanType = ChkStr(rRS.Collect("PlanTypeName"))
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'��̓I�E�햼
	sJobTypeDetail = rRS.Collect("JobTypeDetail")
	'�X�V��
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")
	'�g�b�v�C���^�r���[
	dbTopInterviewFlag = rRS.Collect("TopInterviewFlag")

	If sJobTypeDetail <> "" Then sJobTypeDetail = sJobTypeDetail & "�̂��d�����ڍ�"

	Response.Write "<div id=""tab_switch"">"
	Response.Write "<div class=""left"">"

    ' 2014/12/17�@�摜�Ȃ��Ń^�u��

	If vType = "0" Then
        '�d������\�����̏ꍇ
		'Response.Write "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_orderdetail_on.gif"" alt=""" & sJobTypeDetail & """ border=""0"" width=""93"" height=""22""></div>"
        Response.Write "<p class=""nolink"">�����Ə��</p>"


		If sOrderType = "0" Then
		'��ʂ̋��l�L���̏ꍇ�͉�Џ��ւ̃����N��\��
		'Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./company_order.asp?poc=" & sOrderCode & """ title=""��Џ��""><img src=""/img/order/tab_companyinfo_off.gif"" alt=""��Џ��"" border=""0"" width=""93"" height=""22""></a></div>"
		Response.Write "<a class=""tablink"" href=""./company_order.asp?poc=" & sOrderCode & """>��Ə��</a>"
        End If

		If sOrderType = "0" And dbTopInterviewFlag = "1" Then
			Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""/order/order_interview.asp?ordercode=" & sOrderCode & """ title=""��Џ��""><img src=""/img/order/tab_interview_off.gif"" alt=""�C���^�r���["" border=""0"" width=""93"" height=""22""></a></div>"
		End If
        
        Response.Write "<div class=""autoprintspace""><a href=""./order_detail_autoprint.asp?ordercode=" & sOrderCode & """ target=""_blank""><p class=""Lautoprint""><img src=""/img/order/order_print.png"" class=""autoprint"">���</p></a></div>"
                'Response.Write "<p style=""float: left;margin-right: 4px;margin-top: 2px;background-color: #ffb200;color: #FFF;font-size: 5px;display: block;width: 30px;height: 30px;-webkit-border-radius: 7px;-moz-border-radius: 7px;border-radius: 7px;""><img src=""/img/order/printer87.png"" style=""width:20px;margin:5px 10px 10px 5px;""></p>"
                'Response.Write "<p style=""float: left;margin-right: 4px;margin-top: 2px;background-color: #ffb200;color: #FFF;font-size: 5px;display: block;width: 30px;height: 30px;-webkit-border-radius: 50%;-moz-border-radius: 50%;border-radius: 50%;""><img src=""/img/order/printer87.png"" style=""width:20px;margin:5px 10px 10px 5px;""></p>"
                'Response.Write "<p style=""float: left;margin-right: 4px;margin-top: 2px;background-color: #ffb200;color: #FFF;font-size: 5px;display: block;width: 50px;height: 50px;-webkit-border-radius: 50%;-moz-border-radius: 50%;border-radius: 50%;""><img src=""/img/order/office material7 (2).png"" style=""width:30px;margin:10px 10px 10px 10px;""></p>"
	ElseIf vType = "1" Then
		'��Џ���\�����̏ꍇ
		'Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./order_detail.asp?ordercode=" & sOrderCode & """><img src=""/img/order/tab_orderdetail_off.gif"" alt=""" & sJobTypeDetail & """ border=""0"" width=""93"" height=""22""></a></div>"
        Response.Write "<a class=""tablink"" href=""./order_detail.asp?ordercode=" & sOrderCode & """>�����Ə��</a>"
		If sOrderType = "0" Then
			'��ʂ̋��l�L���̏ꍇ�͉�Џ���\��
			'Response.Write "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_companyinfo_on.gif"" alt=""��Џ��"" border=""0"" width=""93"" height=""22""></div>"
		     Response.Write "<p class=""nolink"">��Ə��</p>"
        End If

		If sOrderType = "0" And dbTopInterviewFlag = "1" Then
			Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""/order/order_interview.asp?ordercode=" & sOrderCode & """ title=""��Џ��""><img src=""/img/order/tab_interview_off.gif"" alt=""�C���^�r���["" border=""0"" width=""93"" height=""22""></a></div>"
		End If

    ' 2014/12/17�@vType=2�͈����s���A���u

	ElseIf vType = "2" Then
		'�C���^�r���[��\�����̏ꍇ
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./order_detail.asp?ordercode=" & sOrderCode & """><img src=""/img/order/tab_orderdetail_off.gif"" alt=""" & sJobTypeDetail & """ border=""0"" width=""93"" height=""22""></a></div>"
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./company_order.asp?poc=" & sOrderCode & """ title=""��Џ��""><img src=""/img/order/tab_companyinfo_off.gif"" alt=""��Џ��"" border=""0"" width=""93"" height=""22""></a></div>"
		Response.Write "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_interview_on.gif"" alt=""��Џ��"" border=""0"" width=""93"" height=""22""></div>"
	End If

	Response.Write "</div>"


	Response.Write "<br clear=""both""></div>" & vbCrLf
    Response.Write "<br clear=""both"">"
End Function

'******************************************************************************
'�쐬���F2014/12/17
'�T�@�v�F���l�[�̃L���b�`�R�s�[�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�i�r/order/order_detail.asp
'���@�l�F2

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

	Dim sOptionNo			'�傫���ʐ^�̔ԍ�
	Dim sCompanyPictureFlag	'��Ǝʐ^�t���O ["1"]�L ["0"]��
	Dim sImg1,sCap1

    Dim sImg2,sImg3,sImg4,sCap2,sCap3,sCap4 '���̑���3���̉摜�̔ԍ�

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
    '���j�����ݒ�
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
	'�傫���摜 start
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
	'�傫���摜 end
	'******************************************************************************

    '******************************************************************************
	'���̑�3���̉摜 start
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
	'���̑�3���̉摜 end
	'******************************************************************************

	'�X�V��
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'���l�[�f�ڊ��� start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">�������O�Ɍf�ڏI������ꍇ������܂��B</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'�f�ڒ� or ��f��
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'�ЊO�Č��Ȃ�DspPublicLimitDay���A�Г��Č��Ȃ�PublicLimitDay��\��
	'�ЊO�Č� OrderType = 0
	'�Г��Č� OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
			'���C�Z���X�؂�̂Ƃ���"�f�ڏI��"�ƕ\��
			sPublishLimitStr = "�f�ڏI��"
			sCautionStr = ""
		Else
			sPublishLimitStr = "��W��"
		End If
	End If
	'------------------------------------------------------------------------------
	'���l�[�f�ڊ��� end
	'******************************************************************************

	'<�Г��Č��p�ʐ^>
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
	'</�Г��Č��p�ʐ^>

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)


	If sImg1 <> "" Then
		Response.Write "<div id=""catchcopy"">"

		Response.Write "<div class=""main_pics"">"
		'Response.Write "<img src=""" & sImg1 & """ alt="""" id=""big_pics"">"

        '�摜���T���l�C���\���Fgallery.js
        '<!-- �f�t�H���g�摜 -->
        Response.Write "<img src=""" & sImg1 & """ alt="""" class=""mainImage"" />"
                        Response.Write "<br style=""clear:both;"" />"
        '<!-- �\��������e�L�X�g -->
        Response.Write "<div class=""messageBox"">"
            Response.Write "<p id=""pict1"" ></p>"
            Response.Write "<p id=""pict2"" class=""invisible"">" & sCap2 & "</p>"
            Response.Write "<p id=""pict3"" class=""invisible"">" & sCap3 & "</p>"
            Response.Write "<p id=""pict4"" class=""invisible"">" & sCap4 & "</p>"
            Response.Write "</div>"

            

        '�摜��2���ڈȍ~���݂���΁A1���T���l�ɒǉ����Ă���
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
'���g���Ă܂���
    if 0 = 1 then
        Response.Write "<div id=""demo"">"
	
        Response.Write "<div id=""MainBox"">"
        '�����ɉ摜���\�������
	    Response.Write "<p class=""large-image""><img src=""" & sImg1 & """ alt=""�ʐ^1"" class=""mainImage""></p>"
	    Response.Write "<p id=""note0"" class=""viewText""></p>"
	    Response.Write "<p id=""note1"" class=""viewText invisible""></p>"
	    Response.Write "<p id=""note2"" class=""viewText invisible"">" & sCap2 & "</p>"
	    Response.Write "<p id=""note3"" class=""viewText invisible"">" & sCap3 & "</p>"
	    Response.Write "<p id=""note4"" class=""viewText invisible"">" & sCap4 & "</p>"
        Response.Write "</div>"

	        Response.Write "<ul class=""small-images clearfix"">"
		        Response.Write "<li><img src=""" & sImg1 & """ alt=""�ʐ^1"" class=""thumb"" rel=""note1""></li>"
		        Response.Write "<li><img src=""" & sImg2 & """ alt=""�ʐ^2"" class=""thumb"" rel=""note2""></li>"
		        Response.Write "<li><img src=""" & sImg3 & """ alt=""�ʐ^3"" class=""thumb"" rel=""note3""></li>"
		        Response.Write "<li><img src=""" & sImg4 & """ alt=""�ʐ^4"" class=""thumb"" rel=""note4""></li>"
	        Response.Write "</ul>"

        Response.Write "</div>"
    End if
        %>

            

        	<div id="lissapo">
            <a href="#detail_waku" class="deju">��W�v��������<span style="float:right;color:#FFF;">>></span><br style="clear:both;"></a>

            <div class="circ1">
                            <img src="/img/neo/webjobad.png">
                            <br style="clear:both;">
            </div>

                <table style="table-layout:fixed;width:100%;margin:4px 0;"><tbody>
			            <tr><td class="cal1_a">�f�ڊ���</td>
                        <td class="cal2_a"><%= sPublishLimitStr %></td></tr>

                        <tr><td class="cal1_a">�X�V��</td>
                        <td class="cal2_a"><%= sUpdateDay %></td></tr>

                        <tr><td class="cal1_b">���R�[�h</td>
                        <td class="cal2_b"><%= dbOrderCode %></td></tr>
                        <tr><td colspan=2 class="cal_long"><img src="/img/order/cau2.png" /></td></tr>

            </tbody></table>
			</div>

           <br clear="all">



           <% If G_USERTYPE = "" Then %> 
            <div id="top_reg_button">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="�������o�^���ĉ���" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="���O�C�����ĉ���" border="0">
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
            
                                <a href="#detail_waku" class="deju2">��W�v��������<span style="float:right;color:#FFF;">>></span><br style="clear:both;"></a></td></tr>


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
			            <tr><td class="cal1_a">�f�ڊ���</td>
                        <td class="cal2_a"><%= sPublishLimitStr %></td></tr>

                        <tr><td class="cal1_a">�X�V��</td>
                        <td class="cal2_a"><%= sUpdateDay %></td></tr>

                        <tr><td class="cal1_b">���R�[�h</td>
                        <td class="cal2_b"><%= dbOrderCode %></td></tr>
                        <tr><td colspan="2" class="cal_long"><img src="/img/order/cau2.png" /></td></tr>

                </tbody></table>
			</div>
            </div>

           <br clear="all">

		   <% If G_USERTYPE = "" Then %> 
			
<div class="center">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="�������o�^���ĉ���" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="���O�C�����ĉ���" border="0">
            </a>
			</div>

		<% End If 

	End If

    '���g���Ă܂���
    if 0 = 1 then
        '����ɑ��̉摜�����݂���Ȃ�
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
'�쐬���F2014/12/17
'�T�@�v�F���l�[�̃^�C�g���������o�́i��W�E��A�L���b�`�R�s�[�A���l�����j
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�i�r/order/order_detail.asp
'���@�l�F2

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

	Dim sOptionNo			'�傫���ʐ^�̔ԍ�
	Dim sCompanyPictureFlag	'��Ǝʐ^�t���O ["1"]�L ["0"]��
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
	'�傫���摜 start
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
	'�傫���摜 end
	'******************************************************************************

	'�X�V��
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'���l�[�f�ڊ��� start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">�������O�Ɍf�ڏI������ꍇ������܂��B</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'�f�ڒ� or ��f��
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'�ЊO�Č��Ȃ�DspPublicLimitDay���A�Г��Č��Ȃ�PublicLimitDay��\��
	'�ЊO�Č� OrderType = 0
	'�Г��Č� OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
			'���C�Z���X�؂�̂Ƃ���"�f�ڏI��"�ƕ\��
			sPublishLimitStr = "�f�ڏI��"
			sCautionStr = ""
		Else
			sPublishLimitStr = "�펞��W��"
		End If
	End If
	'------------------------------------------------------------------------------
	'���l�[�f�ڊ��� end
	'******************************************************************************

	'<�Г��Č��p�ʐ^>�����͕s�v������ƃ��S�o�͎��ɉ��p�������̂Ő����u��
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
	'</�Г��Č��p�ʐ^>

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)

    '�^�C�g���o�͕��i��W�E��ƃL���b�`�R�s�[�j
    Response.Write "<div class=""titleon"">"
        '���l�L���Ȃ��Ɩ���\�����邪�A����ȊO�Ȃ畚����
        If dbOrderType = "0" Then
             Response.Write "<p class=""comadd""><a href=""./company_order.asp?poc=" & sOrderCode & """ class=""c_c"">" & rRS.Collect("CompanyName") & "</a></p>"
             Response.Write "<p class=""jobadd"">" & rRS.Collect("JobTypeDetail") & "</p>"
        Else
            Response.Write "<p class=""comadd"">" & rRS.Collect("CompanySpeciality") & "</p>"
            Response.Write "<p class=""jobadd"">" & rRS.Collect("JobTypeDetail") & "</p>"
        End If
        '���l�̓���������Ε\��
		If sImgSpeciality <> "" Then
			Response.Write "<div class=""ordersp_a"">"
            Response.Write "<div class=""ordersp_b"">"
			'Response.Write "<div style=""font-size:12px;font-weight:normal;color:#008900;"">�y��W�̓����z</div>"
			Response.Write sImgSpeciality
			Response.Write "</div>"
            Response.Write "</div>"
		End If
    Response.Write "</div>"

End Function

'******************************************************************************
'�T�@�v�F���l�[�̃L���b�`�R�s�[�������o��
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�i�r/order/order_detail.asp
'���@�l�F���X�T�|�[�g�Č��p

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

	Dim sOptionNo			'�傫���ʐ^�̔ԍ�
	Dim sCompanyPictureFlag		'��Ǝʐ^�t���O ["1"]�L ["0"]��
	Dim sImg1,sCap1

    Dim sImg2,sImg3,sImg4,sCap2,sCap3,sCap4 '���̑���3���̉摜�̔ԍ�

	Dim sClass
	Dim sImgSpeciality

	Dim sUpdateDay
	Dim sPublishLimitStr
	Dim sCautionStr
	Dim flgNowPublic
	Dim dbCompanyName '���X���̂̋��l�Ɏg����Ж�

	Dim HimlOiwai

	If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")
    '�L���b�`�R�s�[�ݒ�
	dbCatchCopy = rRS.Collect("CatchCopy")
	'���j�����ݒ�
    HimlOiwai = rRS.Collect("CongratulationPrice")

		sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")
	
	'���Ћ��l�p�A�g��Ȃ����ǎc��
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
	'�傫���摜 start
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
	'�傫���摜 end
	'******************************************************************************

	'******************************************************************************
	'���̑�3���̉摜 start
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
	'���̑�3���̉摜 end
	'******************************************************************************

	'�X�V��
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'******************************************************************************
	'���l�[�f�ڊ��� start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">�������O�Ɍf�ڏI������ꍇ������܂��B</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'�f�ڒ� or ��f��
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'�ЊO�Č��Ȃ�DspPublicLimitDay���A�Г��Č��Ȃ�PublicLimitDay��\��
	'�ЊO�Č� OrderType = 0
	'�Г��Č� OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
			'���C�Z���X�؂�̂Ƃ���"�f�ڏI��"�ƕ\��
			sPublishLimitStr = "�f�ڏI��"
			sCautionStr = ""
		Else
			sPublishLimitStr = "�펞��W��"
		End If
	End If

    '<�����������@�\�Ή�>
    '2016/04/01 �r�c���C
    If sPublishLimitStr = "9999/12/31" Then
        '�������̏ꍇ�́A�f�ڊ����Ɍ������w��B�X�V���Ɍ������w��B
        sPublishLimitStr = DateSerial(Year(Date()), Month(Date()) + 1, 0)
        sUpdateDay       = DateSerial(Year(Date()), Month(Date()), 1)
    End If
    '</�����������@�\�Ή�>

	'------------------------------------------------------------------------------
	'���l�[�f�ڊ��� end
	'******************************************************************************

	'<�Г��Č��p�ʐ^>
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
	'</�Г��Č��p�ʐ^>

	sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)


	If sImg1 <> "" Then
		Response.Write "<div id=""catchcopy"">"

		Response.Write "<div class=""main_pics"">"
		'Response.Write "<img src=""" & sImg1 & """ alt="""" id=""big_pics"">"

        '�摜���T���l�C���\���Fgallery.js
        '<!-- �f�t�H���g�摜 -->
        Response.Write "<img src=""" & sImg1 & """ alt="""" class=""mainImage"" />"
                        Response.Write "<br style=""clear:both;"" />"
        '<!-- �\��������e�L�X�g -->
        Response.Write "<div class=""messageBox"">"
            Response.Write "<p id=""pict1"" ></p>"
            Response.Write "<p id=""pict2"" class=""invisible"">" & sCap2 & "</p>"
            Response.Write "<p id=""pict3"" class=""invisible"">" & sCap3 & "</p>"
            Response.Write "<p id=""pict4"" class=""invisible"">" & sCap4 & "</p>"
            Response.Write "</div>"

            

        '�摜��2���ڈȍ~���݂���΁A1���T���l�ɒǉ����Ă���
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
                    <a href="#detail_waku" class="deju">��W�v��������<span style="float:right;color:#FFF;">>></span><br style="clear:both;"></a>
                    <!-- 2014/12/19 -->
<!--                    <div class="cl-effect-13" style="text-align:center;"><a href="#detail_waku" style="background-color:#ffb200;display:block;">��W�v��������</a></div>
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
			            <tr><td class="cal1_a">�f�ڊ���</td>
                        <td class="cal2_a"><%= sPublishLimitStr %></td></tr>

                        <tr><td class="cal1_a">�X�V��</td>
                        <td class="cal2_a"><%= sUpdateDay %></td></tr>

                        <tr><td class="cal1_b">���R�[�h</td>
                        <td class="cal2_b"><%= dbOrderCode %></td></tr>
                        <tr><td colspan="2" class="cal_long"><img src="/img/order/cau0.png" /></td></tr>
            </tbody></table>
			</div>

           <br clear="all">


           <% If G_USERTYPE = "" Then %> 
            <div id="top_reg_button">
            <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="�������o�^���ĉ���" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="���O�C�����ĉ���" border="0">
            </a>

                <!-- 2016/04/14 �r�c���C -->
                <% If True Then %>

                    <!-- ���O�C�����ĂȂ��ꍇ -->
                    <a href="#" onclick="window.open('<%= HTTPS_NAVI_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=staff/mailtocompany.asp?ordercode=<%= sOrderCode %>', '_blank');return false;">
                    <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/conBtn.png" alt="���l�ւ̖⍇��" border="0">
                    </a>

                <% Else %>

                    <!-- ���O�C���ς݂̏ꍇ -->
                    <a href="#" onclick="window.open('<%= HTTPS_NAVI_CURRENTURL %>staff/mailtocompany.asp?ordercode=<%= sOrderCode %>','_blank');return false;">
                    <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/conBtn.png" alt="���l�ւ̖⍇��" border="0">
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
                    <a href="#detail_waku" class="deju2">��W�v��������<span style="float:right;color:#FFF;">>></span><br style="clear:both;"></a>
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
			            <tr><td class="cal1_a">�f�ڊ���</td>
                        <td class="cal2_a"><%= sPublishLimitStr %></td></tr>

                        <tr><td class="cal1_a">�X�V��</td>
                        <td class="cal2_a"><%= sUpdateDay %></td></tr>

                        <tr><td class="cal1_b">���R�[�h</td>
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
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/regBtn.png" alt="�������o�^���ĉ���" border="0">
            </a>
            
            <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>">
            <img src="<%= HTTP_NAVI_CURRENTURL %>img/newbutton/loginBtn.png" alt="���O�C�����ĉ���" border="0">
            </a>
			</div>
			
		<% End If 

	End If
End Function

'******************************************************************************
'�쐬���F2014/12/17
'�T�@�v�F���l�[�ڍ׃y�[�W�̋��l���Ɗ�Ə��̃^�u
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvType			�F�\�������̎�� ["0"]�E���� ["1"]��Џ�� ["2"]�C���^�r���[
'�@�@�@�FvAccessCount	�F�\�������l�[�̃A�N�Z�X��
'�쐬�ҁFLis K.Kaz
'���@�l�F
'�g�p���F�����ƃi�r/order/order_detail.asp
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
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	sOrderCode = rRS.Collect("OrderCode")
	sOrderType = rRS.Collect("OrderType")
	dbPlanType = ChkStr(rRS.Collect("PlanTypeName"))
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

	'��̓I�E�햼
	sJobTypeDetail = rRS.Collect("JobTypeDetail")
	'�X�V��
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")
	'�g�b�v�C���^�r���[
	dbTopInterviewFlag = rRS.Collect("TopInterviewFlag")

	If sJobTypeDetail <> "" Then sJobTypeDetail = sJobTypeDetail & "�̂��d�����ڍ�"

	Response.Write "<div id=""tab_switch"">"
	Response.Write "<div class=""left"">"

    ' 2014/12/17�@�摜�Ȃ��Ń^�u��


		'��Џ���\�����̏ꍇ
		'Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""./order_detail.asp?ordercode=" & sOrderCode & """><img src=""/img/order/tab_orderdetail_off.gif"" alt=""" & sJobTypeDetail & """ border=""0"" width=""93"" height=""22""></a></div>"
        Response.Write "<a class=""tablink_comp"" href=""./order_detail.asp?ordercode=" & sOrderCode & """>�����Ə��</a>"
		If sOrderType = "0" Then
			'��ʂ̋��l�L���̏ꍇ�͉�Џ���\��
			'Response.Write "<div style=""float:left; width:93px; margin:0px;""><img src=""/img/order/tab_companyinfo_on.gif"" alt=""��Џ��"" border=""0"" width=""93"" height=""22""></div>"
		     Response.Write "<p class=""nolink_comp"" style=""padding:10px 16px 4px 16px;"">��Ə��</p>"
        End If

		If sOrderType = "0" And dbTopInterviewFlag = "1" Then
			Response.Write "<div style=""float:left; width:93px; margin:0px;""><a href=""/order/order_interview.asp?ordercode=" & sOrderCode & """ title=""��Џ��""><img src=""/img/order/tab_interview_off.gif"" alt=""�C���^�r���["" border=""0"" width=""93"" height=""22""></a></div>"
		End If

	Response.Write "</div>"


	Response.Write "<br clear=""both""></div>" & vbCrLf
    Response.Write "<br clear=""both"">"
End Function

'******************************************************************************
'�T�@�v�F���l�[�̃X�N���[���ʂɉ������i�r�Q�[�V�����o�[
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_DtlOrder�Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�g�@�p�F�i�r/order/order_detail.asp
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

        '<���O�C����{�^���\���p>
        Dim sPermitFlag			'�f�ڋ��t���O
	    Dim sPublicFlag			'�f�ڃt���O
	    Dim sRiyoFlag			'�f�ڊJ�n��
	    Dim sHakouFlag			'���p�J�n���i���C�Z���X�������j
        Dim flgAddWatchList
	    Dim iMailTemplateCnt	'���[���e���v���[�g�̌���
        '</���O�C����{�^���\���p>

    If GetRSState(rRS) = False Then Exit Function

	dbOrderCode = rRS.Collect("OrderCode")
	dbOrderType = rRS.Collect("OrderType")
	dbCompanyCode = rRS.Collect("CompanyCode")
    sImgSpeciality = GetImgOrderSpeciality(rDB, rRS)

    sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

    '<���O�C����{�^���\���p>
	'******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	sPermitFlag = rRS.Collect("PermitFlag")
	sPublicFlag = rRS.Collect("PublicFlag")
	sRiyoFlag = rRS.Collect("RiyoFlag")
	sHakouFlag = rRS.Collect("HakouFlag")
	iMailTemplateCnt = rRS.Collect("MailTemplateCnt")
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************
    
    '******************************************************************************
	'��ƃR�[�h start
	'------------------------------------------------------------------------------
	flgAddWatchList = False
	sSQL = "EXEC up_ChkWatchListExists_Staff '" & vUserID & "', '" & sOrderCode & "';"
	flgQE = QUERYEXE(rDB, oRS, sSQL, sError)
	If GetRSState(oRS) = True Then
		If oRS.Collect("ExistsFlag") = "1" Then flgAddWatchList = True
	End If
	Call RSClose(oRS)
	'------------------------------------------------------------------------------
	'��ƃR�[�h end
	'******************************************************************************

    Dim qsOrderCode				'�I�[�_�[�R�[�h(�󒍕\�ԍ�)
	Dim iDetail				'���l�[�ڍׂ���̃t���O
	
	qsOrderCode = GetForm("ordercode", 2)
	iDetail = GetForm("Detail", 2)

    '</���O�C����{�^���\���p>


    '******************************************************************************
	'���l�[�f�ڊ��� start
	'------------------------------------------------------------------------------
	sCautionStr = "<p class=""m0"" style=""padding-left:12px;line-height:11px;text-align:left;font-size:10px;color:gray;text-indent:-1em"">�������O�Ɍf�ڏI������ꍇ������܂��B</p>"
	
	sUpdateDay = GetDateStr(rRS.Fields("UpdateDay").Value, "/")

	'�f�ڒ� or ��f��
	flgNowPublic = False
	If rRS.Collect("NowPublicFlag") = "1" Then flgNowPublic = True

	'�ЊO�Č��Ȃ�DspPublicLimitDay���A�Г��Č��Ȃ�PublicLimitDay��\��
	'�ЊO�Č� OrderType = 0
	'�Г��Č� OrderType <> 0
	If sOrderType = "0" Then
		sPublishLimitStr = GetDateStr(ChkStr(rRS.Collect("DspPublicLimitDay")), "/")
	Else
		sPublishLimitStr = ChkStr(rRS.Collect("PublicLimitDay"))
	End If

	If IsNull(sPublishLimitStr) = True Or sPublishLimitStr = "" Then
		If rRS.Collect("NowPublicFlag") = "0" Then
			'���C�Z���X�؂�̂Ƃ���"�f�ڏI��"�ƕ\��
			sPublishLimitStr = "�f�ڏI��"
			sCautionStr = ""
		Else
			sPublishLimitStr = "�펞��W��"
		End If
	End If
	'------------------------------------------------------------------------------
	'���l�[�f�ڊ��� end
	'******************************************************************************
   

    Response.Write "<nav class=""scr""><p class=""navtitle"">"& rRS.Collect("JobTypeDetail") &"</p>"

    %>

    <% '����o�^�A�܂��̓��O�C�������Ă��Ȃ��ꍇ�̃{�^���i����o�^���Ă��牞��A���O�C�����Ă��牞��j
    If G_USERTYPE = "" Then
    %>

    <div class="navibutton">
    <a href="<%= HTTPS_CURRENTURL %>staff/person_reg1.asp?ordercode=<%= dbOrderCode %>" class="button_reg_n">����o�^���ĉ���</a>
    <a href="<%= HTTPS_CURRENTURL %>login_menu.asp?JUMP_URL_FLAG=True&amp;JUMP_URL=/order/order_detail.asp&amp;ordercode=<%= dbOrderCode %>"  class="button_rec_n">���O�C�����ĉ���</a>
    </div>
    <% End If %>

        <% '����o�^�����Ă���ꍇ�̃{�^���i����Ƃ��C�ɓ���ǉ��j
        If vUserType = "staff" Then
        %>

        <div class="navibutton">
                    <% If rRS.Collect("NowPublicFlag") = "1" Then %>
                    <% If flgAddWatchList = True Then %>

                        <span class="kentozumi2_n">���C�ɓ���ɒǉ��ς�</span>

                        <% Else
				        response.write "<form id=""frmSendMailJobOfferAddress"" name=""frmSendMailJobOfferAddress"" method=""post"" action=""../staff/watchlist_register.asp"" style=""width:70%;float:left"" onSubmit=""return Submit();"">"
				        'Response.Write "<a href=""#"" onclick=""document.forms.frmSendMailJobOfferAddress.submit();return false;"" class=""kento2_n"">���C�ɓ���ɒǉ�</a>"
                        Response.Write "<a href=""#"" onclick=""document.frmSendMailJobOfferAddress.submit();return false;"" class=""kento2_n"">���C�ɓ���ɒǉ�</a>"
				        response.write "<input type=""hidden"" name=""CONF_OrderCode"" value='"& qsOrderCode &"'>"
				        'response.write "<input type=""text"" name=""dummy"" style=""display:none;"" />"
				        'response.write "<input type=""text2"" name=""dummy2"" style=""display:none;"" />"
				        response.write "</form>"
                        End If %>

                    <% If dbOrderType = "0" Then %>
                        <a href="#" onclick="contactCompanyAdv('');return false;" class="button_obo_n">
                        ���̋��l�ɉ��傷��
                        </a>

                    <% Else %>
                        <a href="#" onclick="contactCompanyLis('');return false;" class="button_obo_n">
                        ���̋��l�ɉ��傷��
                        </a>
                        <%
                        'Response.Write "<a href=""#"" onclick=""contactCompany('1');return false;"" class=""button_que_n"">"
                        'Response.Write "���̋��l�ɂ��Ď��₷��"
                        'Response.Write "</a>"
                        %>
                    <% End If %>
            

                        
                    <% Else %>
			        <li id="finKokoku_n"><div class="description" align="center"><b>���̋��l�[�͌f�ڂ��I�����Ă��܂��B���[�����M�͂ł��܂���B</b></div></li>
		            <% End If %>

        <% End If %>
    
    
    
    <a href="#waku_<%= HimlOiwai %>" class="button_det_n">���l�ڍ�</a>
    <!-- <li><a href="../index.asp">home</a></li> -->
    </div>
    <%
    Response.Write "</nav>"


End Function

'******************************************************************************
'�T�@�v�F������O�̏�������
'���@���FrDB			�F�ڑ�����DBConnection
'�@�@�@�FrRS			�Fup_SearchOrder or ���l�[�ڍ׌���SQL �Ő������ꂽ���R�[�h�Z�b�g�I�u�W�F�N�g
'�@�@�@�FvUserType		�F���p�����[�U�̃��[�U��� [Session("usertype")]
'�@�@�@�FvUserID		�F���p�����[�U�̃��[�UID [Session("userid")]
'�@�@�@�FvOrderCode		�F�{�������l�[
'���@�l�F
'�g�p���Forder/order_detail_autoprint.asp
'���@���F2015/08/06 Kimura func_order.asp����قڗ��p
'******************************************************************************
Function AutoPrintHistoryOrder(ByRef rDB, ByRef rRS, ByVal vUserType, ByVal vUserID, ByVal vOrderCode)
	Dim sSQL
	Dim oRS
	Dim flgQE
	Dim sError

	'�Г�����̃A�N�Z�X�ƁA�����낤����(S0018066)����̃A�N�Z�X�̓��O�Ɏc���Ȃ�
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