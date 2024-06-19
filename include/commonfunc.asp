<!-- #INCLUDE FILE="../../commonfunc.asp" -->
<!-- #INCLUDE FILE="../../funcCreateMailBody.asp" -->
<!-- #INCLUDE FILE="func_autologin.asp" -->
<!-- #INCLUDE FILE="func/chkSmartPhone.asp" -->
<!-- #INCLUDE FILE="func/dspTopicPath.asp" -->
<!-- #INCLUDE FILE="func/getURLDecode.asp" -->
<!-- #INCLUDE FILE="func/getHeaderText.asp" -->
<!-- #INCLUDE FILE="func/getTabIndexType.asp" -->
<!-- #INCLUDE FILE="func/htmlHeader.asp" -->
<!-- #INCLUDE FILE="func/htmlFooter.asp" -->
<!-- #INCLUDE FILE="func/htmlTabIndex.asp" -->
<!-- #INCLUDE FILE="func/scrGoogleAdwords_convertion.asp" -->
<!-- #INCLUDE FILE="func/scrOverT_opti_convertion.asp" -->
<!-- #INCLUDE FILE="func/scrOverT_opti_shopper.asp" -->
<!-- #INCLUDE FILE="func/scrOverT_opti_univeral.asp" -->
<!-- #INCLUDE FILE="func/scrRefAll.asp" -->
<!-- #INCLUDE FILE="func/scrTwitterFollowBadge.asp" -->
<!-- #INCLUDE FILE="func/scrIntroTwitterFollowBadge.asp" -->
<!-- #INCLUDE FILE="func/scrgoogleremarketing.asp" -->
<!-- #INCLUDE FILE="func/scrgoogleremarketing_thanks.asp" -->
<!-- #INCLUDE FILE="func/getStrJoinSep.asp" -->
<!-- #INCLUDE FILE="func/scrYahoo_convertion.asp" -->

<% 
'メンテナンスページへのリダイレクト設定
'メンテナンス時間の指定

    Dim MainteStDay 
    Dim MainteStTime
    Dim MainteEdDay
    Dim MainteEdTime
    Dim MainteStTimeE
    Dim MainteEdTimeS
    Dim curTime,curDate
    Dim Mainte201706flag
    Mainte201706flag = "0"

MainteStDay = clng("20170623")
MainteEdDay = clng("20170623")
    MainteStTime = clng("0500")
    MainteStTimeE = clng("2359")
    MainteEdTimeS = clng("0000")
    MainteEdTime = clng("2000")

'テスト用
'MainteStDay = clng("20170623")
'MainteEdDay = clng("20170625")
'    MainteStTime = clng("0500")
'    MainteStTimeE = clng("2359")
'    MainteEdTimeS = clng("0000")
'    MainteEdTime = clng("2000")
'
'    response.write " St " & MainteStDay & " & " & MainteStTime&"<BR>"
'    response.write " Ed " & MainteEdDay & " & " & MainteEdTime&"<BR>"

'    curDate = Mid(Date(),1,4) & Mid(Date(),6,2) & Mid(Date(),9,2)
'    curTime = Mid(Time(),1,2) & Mid(Time(),4,2)
'    curDate = Mid(NOW(),1,10)
    curDate = clng(Replace(Mid(NOW(),1,10),"/",""))
    curTime = clng(Replace(Mid(NOW(),11,6),":",""))

'    response.write " Nw " & curDate & " & " & curTime&"<BR>"
'    response.write " i1 "& MainteStDay &"="& curDate & "<BR>"
'    response.write "  "& MainteStTime &"<="& curTime &"<="& MainteStTimeE & "<BR>"
'    response.write " i2 "& curDate &"="& MainteEdDay & "<BR>"
'    response.write "  "& MainteEdTimeS &"<="& curTime &"<="& MainteEdTime & "<BR>"

'Response.write HTTPS_CURRENTURL & "maintenance/index.asp"& "<BR>"
'Response.Write G_IPADDRESS


    If(MainteStDay = curDate)  Then
        If(MainteStTime <= curTime ) And (curTime <= MainteStTimeE) Then
            Mainte201706flag = "1"
        End If
    ElseIf(curDate = MainteEdDay) Then
   		If (MainteEdTimeS <= curTime) And (curTime <= MainteEdTime) Then
            Mainte201706flag = "1"
        End If
    End if
		
		if (Mainte201706flag = "1") Then
           If IsRE(G_IPADDRESS, "^114.147.197.173", True) = true Then 
                'response.redirect HTTPS_CURRENTURL & "maintenance/index.asp"
       	   Else
	         response.write "しごとナビ"
		    End if
        End if            



%>



