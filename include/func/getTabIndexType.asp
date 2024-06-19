<%
'*******************************************************************************
'概　要：タブIndexを取得
'引　数：
'戻り値：String
'備　考：
'履　歴：2010/05/13 LIS K.Kokubo 作成
'*******************************************************************************
Function getTabIndexType(ByVal vURL)
	vURL = LCase(vURL)
	Select Case vURL
		'はじめての方タブ
		Case "/tab/index1.asp","/staff/person_reg1.asp","/staff/passwordreminder.asp","/staff/s_aboutnavi.asp","/staff/qa.asp","/staff/s_kiyaku.asp","/lis/lis.asp","/shigotonavi/sitemap.asp","/privacy/privacy.asp","/staff/passwordreminder.asp","/lis/service-development.asp","/lis/lis_annai.asp","/privacy/sengen.asp","/privacy/privacymark.asp","/promotion/mobilepromotion.asp","/promotion/conpri_riyou.asp","/logout.asp","/error/404.html","/error/400.html","/infomation/nenmatsunenshi/index.asp","/order/order_search_detail.asp","/staff/kiyakus.asp","/promotion/s_conpri_riyou.asp","/pr/search_index.asp"
		: getTabIndexType = 0
		'求人を探すタブ
		Case "/support/index.asp","/tab/index8.asp","/search/index.asp","/city/minatoku.asp","/city/chiyodaku.asp","/city/chuouku.asp","/city/shinjukuku.asp","/city/koutouku.asp","/railway/railway_search1.asp","/lis/lis_saiyou.asp","/order/order_list.asp","/railway/railway_search2.asp","/lis/regional_manager_recruit.asp","/lis/recruit_manager_message.asp","/railway/railway_search3.asp","/order/company_order.asp","/lis/regional_manager_message.asp","/order/order_interview.asp","/staff/searchordercondition/index.asp","/order/order_list_accesscount.asp","/s_contents/accounting_special.asp","/s_contents/s_localgoverment_hokkaido_tohoku.asp","/s_contents/s_localgoverment_kanto_koushinetsu.asp","/s_contents/s_localgoverment_tokai_hokuriku.asp","/s_contents/s_localgoverment_kinki.asp","/s_contents/s_localgoverment_shikoku_chugoku.asp","/s_contents/s_localgoverment_kyushu_okinawa.asp","/error/trouble.asp","/promotion/lis_recruit.asp","/infomation/20110318.asp","/staff/s_resume.asp","/staff/s_resume_kakikata.asp","/staff/s_resume_kakikata2.asp","/staff/s_resume_qa.asp","/staff/s_careersheet.asp","/staff/s_careersheet_kakikata_1.asp","/s_contents/motive_index.asp","/s_contents/s_jikopr.asp","/s_contents/s_taishokunegai.asp","/s_contents/s_year_calculation.asp","/s_contents/s_year_calculation2.asp","/s_contents/s_year_calculation3.asp","/promotion/conpripromotion.asp","/s_contents/s_introduce.asp","/s_contents/s_temporary.asp","/s_contents/s_temptoperm.asp","/staff/s_resume_xperror.asp","/type_map.asp","/area_map.asp","/keyword_map.asp","/news.asp","/neo/howabout/index.asp","/recruit/index.asp"
		: getTabIndexType = 1
		'便利ツールタブ
		Case "/tab/index2.asp","/staff/s_careersheet2.asp","/s_contents/motive_tool.asp","/conpri/index.html","/conpri/step1.html","/conpri/step2.html","/conpri/step3.html","/staff/s_careersheet_kakikata_b_employment.asp","/staff/s_careersheet_kakikata_s_employment.asp","/staff/s_careersheet_kakikata_clericalwork.asp","/s_contents/s_logicalpoint.asp","/link_collection.asp","/link_collect01.asp","/link_collect02.asp","/link_collect03.asp","/link_collect04.asp","/link_collect05.asp","/link_collect06.asp","/link_collect07.asp","/link_collect08.asp","/link_collect09.asp","/link_collect10.asp","/conpri/help_netprint.asp","/promotion/blog_part.asp","/staff/s_resume_sheet.asp","/staff/oshigotochecker_download.asp","/api/index.asp","/api/rule/index.asp","/api/key/index.asp","/api/manual/index.asp","/api/approval/index.asp": getTabIndexType = 2
		'転職サポートタブ
		Case "/tab/index3.asp","/s_contents/navistep_seme.asp","/s_contents/navistep_mamori.asp","/s_contents/navistep_saport.asp","/staff/access.asp","/s_contents/s_changejob_support.asp","/s_contents/request_service.asp","/column/column_motive_example.asp","/column/column_motive.asp","/column/column_motive_reverse.asp","/column/column_apply_decision.asp","/column/column_aptitudetest.asp","/column/column_drawup_document.asp","/column/column_bad_example.asp","/column/column_vision.asp","/column/column_uneasiness.asp","/column/column_interview_finish.asp","/column/column_1.asp","/s_contents/s_books.asp","/staff/jobcon/index.asp","/staff/schedule/index.asp","/staff/schedule/day.asp","/staff/schedule/reg.asp","/s_contents/s_mynavi_result.asp","/s_contents/s_mynavi_tool.asp","/promotion_itfair06/main.asp","/promotion_itfair06/picture_album_06.asp","/promotion_itfair06/register.asp","/promotion_itfair06/company_list.asp","/promotion_itfair06/company.asp","/staff/wa_help.asp","/consultant/consultantbranch.asp","/consultant/consultantlist.asp","/consultant/consultantdetail.asp","/consultant/consultantinquiry.asp","/form_mail.asp","/staff/jobcon/careeranalyzer/motive.asp","/staff/jobcon/careeranalyzer/ideal.asp","/staff/jobcon/interviewsimulator/question.asp","/staff/jobcon/interviewsimulator/answer.asp","/staff/jobcon/careerconsultation/confirm.asp","/staff/jobcon/careerconsultation/regist.asp","/staff/jobcon/careerconsultation/thanks.asp","/staff/s_searchexplanation.asp","/s_contents/s_mynavi_request_service.asp","/fb/nayami/index.asp","/skill/index.asp","/skill/list/itpassport.asp"
		: getTabIndexType = 3
		'コミュニティタブ
		Case "/tab/index4.asp","/staff/ranking.asp","/infomation/info.asp","/cafe/cafe_list.asp","/cafe/cafe_detail.asp","/cafe/cafe_help.asp","/cafe/cafe_rule_dsp.asp","/cafe/cafe_topic_edit.asp","/cafe/cafe_comment_edit.asp","/cafe/cafe_profile_edit.asp","/shigotonavi/navi_revo_main.asp","/shigotonavi/navi_revo_1.asp","/company/contents/satisfaction.asp","/shigotonavi/navi_revo_2.asp","/shigotonavi/navi_revo_3.asp","/company/inquiry.asp","/shigotonavi/navi_revo_4.asp","/shigotonavi/navi_revo_5.asp","/staff/ranking_index.asp","/staff/ranking_jobtype.asp","/staff/ranking_industry.asp","/staff/ranking_station.asp","/staff/ranking_word.asp","/ranking/ranking_graduationuniversity.asp","/link.asp","/infomation/renew.asp","/shigotonavi/kaikaku.asp","/shigotonavi/kaikaku2.asp","/shigotonavi/kaikaku3.asp","/shigotonavi/kaikaku4.asp","/shigotonavi/kaikaku5.asp","/shigotonavi/kaikaku6.asp","/shigotonavi/kaikaku7.asp","/shigotonavi/kaikaku8.asp","/promotion/20100614.asp","/promotion/twitter.asp","/lis/blog_kimura.asp": getTabIndexType = 4
		'採用ご担当者タブ
		Case "/company/about.asp","/company/index.asp","/company/c_hajime.asp","/company/c_function.asp","/company/c_voice.asp","/company/research.asp","/company/c_staffdata.asp","/company/charge.asp","/company/autoestimate.asp","/jinzai/jinzai.asp","/company/qa.asp","/company/c_kiyaku.asp","/company/access.asp","/company/c_introduce.asp","/company/c_temptoperm.asp","/company/c_dispatch.asp","/company/c_contents01_00.asp","/company/c_contents02_00.asp","/company/c_contents03_00.asp","/company/c_contents04_00.asp","/company/c_ad.asp","/jinzai/jinzai_reg.asp","/company/c_contents01_01.asp","/company/c_contents01_02.asp","/company/c_contents01_03.asp","/company/c_contents01_04.asp","/company/c_contents01_05.asp","/company/c_contents01_06.asp","/company/c_contents01_07.asp","/company/c_contents01_08.asp","/company/c_contents01_09.asp","/company/c_contents01_10.asp","/company/c_contents01_11.asp","/company/c_contents01_12.asp","/company/c_contents01_13.asp","/company/c_contents02_01.asp","/company/c_contents02_02.asp","/company/c_contents02_03.asp","/company/c_contents02_04.asp","/company/c_contents02_05.asp","/company/c_contents02_06.asp","/company/c_contents02_07.asp","/company/c_contents02_08.asp","/company/c_contents02_09.asp","/company/c_contents02_10.asp","/company/c_contents02_11.asp","/company/c_contents02_12.asp","/company/c_contents02_13.asp","/company/c_contents02_14.asp","/company/c_contents02_15.asp","/company/c_contents02_16.asp","/company/c_contents02_17.asp","/company/c_contents02_18.asp","/company/c_contents02_19.asp","/company/c_contents02_20.asp","/company/c_contents02_21.asp","/company/c_contents03_01.asp","/company/c_contents03_02.asp","/company/c_contents03_03.asp","/company/c_contents03_04.asp","/company/c_contents03_05.asp","/company/c_contents03_06.asp","/company/c_contents04_01.asp","/company/c_contents04_02.asp","/company/c_contents04_03.asp","/company/c_contents04_04.asp","/company/c_contents04_05.asp","/company/c_contents04_06.asp","/company/c_contents04_07.asp","/company/c_contents04_08.asp","/company/c_contents04_09.asp","/company/c_contents04_10.asp","/company/c_contents04_11.asp","/company/c_contents04_12.asp","/company/c_contents04_13.asp","/company/c_contents04_14.asp","/company/c_contents04_15.asp","/company/c_contents04_16.asp","/company/c_contents04_17.asp","/company/c_contents04_18.asp","/company/c_contents04_19.asp","/company/c_contents04_20.asp","/company/c_contents04_21.asp","/company/c_contents04_22.asp","/company/c_contents04_23.asp","/company/c_scout3point.asp","/company/c_successpoint.asp","/company/request01.asp","/company/request02.asp","/company/request03.asp","/company/request04.asp","/agency/index.asp","/agency/login/index.asp","/agency/login/detail.asp","/jinzaisearch/index.asp","/jinzaisearch/person_detail.asp","/gyoumu_teikei.asp","/company/twitter/inquiry.asp","/promotion/introtwitter.asp","/company/c_point.asp","/company/10yearsanniversary/index.asp","/company/10yearsanniversary/bannerapplication.asp","/company/10yearsanniversary/request01.asp","/company/10yearsanniversary/request02.asp","/company/10yearsanniversary/request03.asp","/company/10yearsanniversary/request04.asp","/company/10yearsanniversary/requestprintview.asp","/company/c_login_help.asp","/company/costperformance/index.asp","/neo/shoukai/index.asp","/neo/kokoku/index.asp","/staff/kiyaku.asp"
		: getTabIndexType = 5
		'My Page(求職者)
		Case "/tab/index6.asp","/staff/s_login.asp","/staff/mailhistory_person.asp","/staff/watchlist.asp","/staff/resume_print.asp","/staff/resume_picture.asp","/staff/changepassword.asp","/staff/changepassword_register.asp","/staff/wa_help.asp","/staff/mail_detail_person.asp","/staff/person_edit1.asp","/staff/person_edit2.asp","/staff/person_edit3.asp","/staff/person_edit4.asp","/staff/person_edit5.asp","/staff/person_edit6.asp","/staff/person_edit7.asp","/staff/person_edit8.asp","/staff/person_edit9.asp","/staff/person_edit9_1.asp","/staff/mailtocompany.asp","/staff/resume.asp","/staff/footprint.asp","/suspension/questionnarie.asp","/suspension/already.asp","/staff/person_registed.asp","/infomation/20110318.asp","/staff/edit_list.asp","/staff/my_footprint.asp","/staff/step2a.asp","/staff/step2c.asp"
		: getTabIndexType = 6
		'My Page(企業)
		Case "/tab/index7.asp","/company/c_login.asp","/map/maphelp.asp","/company/watchlist.asp","/company/mailhistory_company.asp","/company/company_reg1.asp","/order/company_order.asp","/company/orderedit/base.asp","/company/orderedit/edit01.asp","/company/orderedit/edit02.asp","/company/orderedit/edit03.asp","/company/orderedit/edit04.asp","/company/orderedit/edit05.asp","/company/orderedit/edit06.asp","/company/orderedit/edit07.asp","/company/orderedit/edit08.asp","/company/orderedit/edit09.asp","/company/orderedit/edit10.asp","/company/orderedit/edit11.asp","/company/orderedit/edit12.asp","/company/orderedit/edit13.asp","/company/orderedit/edit13_1.asp","/company/orderedit/edit13_2.asp","/company/orderedit/edit14.asp","/company/orderedit/edit15.asp","/company/orderedit/edit16.asp","/company/orderedit/edit17.asp","/company/orderedit/edit18.asp","/company/orderedit/edit19.asp","/company/orderedit/edit20.asp","/company/orderedit/edit21.asp","/company/orderedit/edit22.asp","/map/map.asp","/company/orderedit/matching/index.asp","/company/orderedit/matching/edit01.asp","/company/orderedit/matching/edit02.asp","/company/orderedit/matching/edit03.asp","/company/orderedit/matching/edit04.asp","/company/orderedit/matching/edit05.asp","/company/orderedit/matching/edit06.asp","/company/orderedit/matching/edit07.asp","/company/orderedit/matching/edit08.asp","/company/orderedit/matching/edit09.asp","/company/orderedit/matching/edit10.asp","/company/orderedit/matching/edit11.asp","/company/orderedit/matching/edit12.asp","/company/orderedit/matching/edit13.asp","/company/orderedit/matching/edit14.asp","/company/orderedit/matching/edit15.asp","/company/orderedit/matching/edit16.asp","/company/orderedit/matching/edit17.asp","/company/orderedit/matching/edit18.asp","/company/orderedit/matching/edit19.asp","/company/orderedit/matching/edit20.asp","/company/orderedit/matching/edit21.asp","/company/orderedit/matching/edit22.asp","/company/orderedit/matching/edit23.asp","/company/company_reg2.asp","/company/img_upload.asp","/company/company_img_list.asp","/company/myorderlist.asp","/license/license_manager.asp","/company/mail_detail_company.asp","/company/mailhistory_progress.asp","/company/company_reg3.asp","/license/license_confirm.asp","/company/company_img_detail.asp","/license/license_register.asp","/staff/person_search_detail.asp","/company/order_img_list.asp","/staff/person_list.asp","/mailtemplate/manager.asp","/mailtemplate/regist.asp","/mailtemplate/copy.asp","/mailtemplate/delete.asp","/mailtemplate/copyreference.asp","/company/mailtoperson.asp","/company/costperformance/explain.asp","/company/costperformance/manage/index.asp","/company/costperformance/manage/reg_branch.asp","/company/costperformance/manage/reg_media.asp","/company/costperformance/manage/search_media.asp","/company/costperformance/manage/reg_adoptplan.asp","/company/costperformance/manage/reg_adoptresult.asp","/company/costperformance/branch.asp","/company/costperformance/media.asp","/company/costperformance/simulation/index.asp","/company/costperformance/simulation/reference.asp","/company/costperformance/year/index.asp","/company/costperformance/year/branch.asp","/company/orderedit/new.asp","/company/elderinterview/list.asp","/company/elderinterview/reg.asp","/company/elderinterview/picture_reg.asp","/company/report/footprint.asp","/company/lumpmail/list.asp","/company/lumpmail/send.asp","/company/searchstaffcondition/list.asp","/company/license/mailplan_status.asp","/company/point/index.asp","/company/point/offer.asp","/company/changepassword.asp","/company/topinterview/reg.asp","/company/topinterview/picture_reg.asp","/company/topinterview/qa_reg.asp": getTabIndexType = 7
		
		'学ぶ
		Case "/s_contents/businesscolumns/com01.asp","/s_contents/businesscolumns/index.asp","/manabu/column/index.asp","/manabu/knowhow/index.asp","/manabu/skillup/index.asp","/manabu/jiko/index.asp","/manabu/index.asp","/staff/jobcon/introduction.asp","/staff/jobcon/careeranalyzer/index.asp","/staff/jobcon/searchadvice/index.asp","/s_contents/s_kyuuyomeisai.asp","/s_contents/labo/scoutlabo.asp","/staff/notification_mail_service.asp","/staff/jobcon/interviewsimulator/index.asp","/s_contents/s_ready.asp","/s_contents/s_proce.asp","/s_contents/s_goukaku.asp","/s_contents/navistep_index.asp","/column/column_index.asp","/s_contents/s_localgoverment.asp","/s_contents/businesscolumns/com02.asp","/s_contents/businesscolumns/com03.asp","/s_contents/businesscolumns/com04.asp","/s_contents/businesscolumns/com05.asp","/s_contents/businesscolumns/men01.asp","/s_contents/businesscolumns/men02.asp","/s_contents/businesscolumns/men03.asp","/s_contents/businesscolumns/pre01.asp","/s_contents/businesscolumns/pre02.asp","/s_contents/businesscolumns/pre03.asp","/s_contents/warmreception/detail.asp","/s_contents/s_mensetsu_index.asp","/s_contents/s_logical.asp","/s_contents/s_neet.asp","/s_contents/warmreception/offer.asp"
		: getTabIndexType = 8
		
		'TOP
		Case"/index.asp"
		: getTabIndexType = 10
		
		'交流
		Case "/koryu/index.asp"
		: getTabIndexType = 11
		
		'リンク
		Case "/link/index.asp","/s_contents/enquete.asp","/s_contents/s_mynavi.asp","/s_contents/s_introduce_swf.asp","/link/tieup/processlink.asp"
		: getTabIndexType = 12
		
		Case "/login_menu.asp","/login_menu_in.asp"
		: getTabIndexType = 13

		Case "/order/order_detail.asp"
		: getTabIndexType = 14


		Case Else: getTabIndexType = -1
	End Select

	'<ＡＢテストなどの一時的なページはコチラで指定>
	Select Case vURL
		'はじめての方タブ
		'Case "": getTabIndexType = 0
		'求人を探すタブ
		'Case "": getTabIndexType = 1
		'便利ツールタブ
		Case "/staff/s_resume3.asp": getTabIndexType = 2
		'転職サポートタブ
		'Case "": getTabIndexType = 3
		'コミュニティタブ
		'Case "": getTabIndexType = 4
		'採用ご担当者タブ
		'Case "": getTabIndexType = 5
		'My Page(求職者)
		'Case "": getTabIndexType = 6
		'My Page(企業)
		'Case "": getTabIndexType = 7
		'Case Else: getTabIndexType = -1
	End Select
	'</ＡＢテストなどの一時的なページはコチラで指定>

	If getTabIndexType > 0 Then Exit Function

	'Select Case では指定し辛いURLの場合
	If IsRE(vURL,"/lis/recruit_shinsotsu\d\d((\_index)|(\_explanation)|(\_info)|(\_interview)|(\_form)).asp",False) = True _
	Or IsRE(vURL,"/s_contents/license/\d\d\d\d\d\d\d.asp",False) = True _
	Or IsRE(vURL,"/order/special/",False) = True Then
		'求人を探すタブ
		getTabIndexType = 1
	ElseIf (vURL = "/company/costperformance/index.asp" And G_USERTYPE = "") Then
		'採用ご担当者タブ
		getTabIndexType = 5
	ElseIf (vURL = "/staff/person_detail.asp" And G_USERTYPE = "staff") _
	Or (vURL = "/logout.asp" And G_USERTYPE = "staff") Then
		'My Pageタブ(求職者)
		getTabIndexType = 6
	ElseIf (vURL = "/staff/person_detail.asp" And G_USERTYPE = "company") _
	Or (vURL = "/logout.asp" And G_USERTYPE = "company") Then
		'My Pageタブ(企業)
		getTabIndexType = 7
	End If
End Function
%>
