<%@ Language="VBScript" CodePage="932" %><% Option Explicit %>
<!-- #INCLUDE virtual="/config/personnel.asp" -->
<!-- #INCLUDE virtual="/include/connect.asp" -->
<!-- #INCLUDE virtual="/include/commonfunc.asp" -->
<%
Dim oRS,sSQL,sError
if IsMainCode(Session("userid")) = true then
	sSQL = "Select IsNull(PortableMailAddress,'') as PortableMailAddress  from P_Info Where StaffCode='" & Session("userid") & "'"
	Call QUERYEXE(dbconn,oRS,sSQL,sError)
	If GetRSState(oRS) = true then
		If ChkPortableMail(oRS.Collect("PortableMailAddress")) = true then
			Select Case Right(oRS.Collect("PortableMailAddress"),Len(oRS.Collect("PortableMailAddress"))-InStr(oRS.Collect("PortableMailAddress"),"@"))
			
			Case "docomo.ne.jp","disney.ne.jp"
			
				Response.Write "<div style=""margin-top:10px;padding:6px;line-height:2.5em;text-align:center;border:solid 1px lightgray"">"
				Response.Write "<img src=""/img/promotion/mobilepromotion/docomo_logo.gif"" align=""left"" style=""margin-left20px;margin-left:40px;"">"
				Response.Write "<b style=""color:gray"">「ドコモR」の「ｉモードR」メニューサイト「しごとナビモバイル」なら同一ＩＤで携帯からも利用可能</b><br>"
				Response.Write "<span style="""">[ドコモｉＭｅｎｕ]→[メニューリスト]→[働く/住む/学ぶ]→[就職/転職]→[しごとナビモバイル] </span>"
				Response.Write "</div>"
				
			Case "ezweb.ne.jp"
				Response.Write "<div style=""margin-top:10px;padding:6px;line-height:2.5em;text-align:center;border:solid 1px lightgray"">"
				Response.Write "<img src=""/img/promotion/mobilepromotion/ezweb_logo.gif"" align=""left"" style=""margin:5px 20px;margin-left:40px;"">"
				Response.Write "<b style=""color:gray"">au公式サイトの「しごとナビモバイル」なら同一ＩＤで携帯からもご利用できます。</b><br>"
				Response.Write "<span style="""">[au公式サイト]→[カテゴリで探す]→[仕事・学習]→[仕事・資格]→[しごとナビモバイル] </span>"
				Response.Write "</div>"
				
			Case "softbank.ne.jp","i.softbank.jp","t.vodafone.ne.jp","d.vodafone.ne.jp","h.vodafone.ne.jp","c.vodafone.ne.jp","k.vodafone.ne.jp","r.vodafone.ne.jp","n.vodafone.ne.jp","s.vodafone.ne.jp","q.vodafone.ne.jp"
	
				Response.Write "<div style=""margin-top:10px;padding:6px;line-height:2.5em;text-align:center;border:solid 1px lightgray"">"
				Response.Write "<img src=""/img/promotion/mobilepromotion/softbank_logo.gif"" align=""left"" style=""margin-top:10px;margin-left:40px;"">"
				Response.Write "<b style=""color:gray"">SoftBankMobileの公式サイト「しごとナビモバイル」なら同一ＩＤで携帯から利用可能</b><br>"
				Response.Write "<span style="""">[Yahoo！ケータイ]→[メニューリスト]→[働く・住む・学ぶ]→[就職・転職・バイト]→[しごとナビモバイル] </span>"
				Response.Write "</div>"
			
			Case "willcom.com","wm.pdx.ne.jp","dj.pdx.ne.jp","di.pdx.ne.jp","dk.pdx.ne.jp","pdx.ne.jp"
				
				Response.Write "<div style=""margin-top:10px;padding:6px;line-height:2.5em;text-align:center;border:solid 1px lightgray;"">"
				Response.Write "<img src=""/img/promotion/mobilepromotion/willcomLOGO.gif"" align=""left"" style=""margin-top:10px;margin-left:40px;"">"
				Response.Write "<b style=""color:gray"">WILLCOM公式サイトの「しごとナビモバイル」なら同一ＩＤで携帯からもご利用できます。</b><br>"
				Response.Write "<span style="""">[WILLCOM公式サイト]→[ライフ＆ショッピング]→[仕事・資格]→[しごとナビモバイル] </span>"
				Response.Write "</div>"
			End Select
		End If
	End If
end if
%>