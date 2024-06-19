<%
'*******************************************************************************
'概　要：企業からのメール着信通知メールの文言を生成
'引　数：
'戻り値：
'備　考：
'履　歴：2011/04/26 LIS K.Kokubo 作成
'*******************************************************************************
Function scrGoogleAdwords_convertion()
	Response.Write "<!--アドワーズコンバージョンタグ-->" & vbCrLf
	Response.Write "<div style=""display:none;"">" & vbCrLf
	%>
	<!-- Google Code for &#30331;&#37682;&#23436;&#20102; Conversion Page -->
	<script type="text/javascript">
    /* <![CDATA[ */
    var google_conversion_id = 1070319369;
    var google_conversion_language = "en";
    var google_conversion_format = "2";
    var google_conversion_color = "ffffff";
    var google_conversion_label = "crcfCInTngQQiY6v_gM";
    var google_conversion_value = 0;
    /* ]]> */
    </script>
    <script type="text/javascript" src="https://www.googleadservices.com/pagead/conversion.js">
    </script>
    <noscript>
    <div style="display:inline;">
    <img height="1" width="1" style="border-style:none;" alt="" src="https://www.googleadservices.com/pagead/conversion/1070319369/?value=0&amp;label=crcfCInTngQQiY6v_gM&amp;guid=ON&amp;script=0"/>
    </div>
    </noscript>
	<%
	Response.Write "</div>" & vbCrLf
End Function
%>
