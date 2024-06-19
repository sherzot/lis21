<%
	Dim sSQL_input_check	: sSQL_input_check = "sp_GetPersonBase '" & Session("userid") & "'"
	Dim oRS_input_check		: Set oRS_input_check = dbconn.Execute(sSQL_input_check)
	Dim oRS_index_check2
	Dim aHopeWorkingType_input_check(2)
	Dim sHopePrefectureCode
	Dim idx_input_check

	Dim UnAge_input_check
%>
function DataCheckPerreg1()
{
/***********************************/
/*		入力チェック			 */
/***********************************/

	if (!necessaryInputCheck(document.forms[0].Name_K_f,"Text","姓（漢字）を入力して下さい。")) return false;
	if (!necessaryInputCheck(document.forms[0].Name_F_f,"Text","姓（フリガナ）を入力して下さい。")) return false;
	if (!necessaryInputCheck(document.forms[0].Name_K_b,"Text","名（漢字）を入力して下さい。")) return false;
	if (!necessaryInputCheck(document.forms[0].Name_F_b,"Text","名（フリガナ）を入力して下さい。")) return false;
	if (!necessaryInputCheck(document.forms[0].Post_U,"Text","郵便番号を全て入力して下さい。")) return false;
	if (!necessaryInputCheck(document.forms[0].Post_L,"Text","郵便番号を全て入力して下さい。")) return false;
	if (!necessaryInputCheck(document.forms[0].Prefecture,"Select","都道府県を入力して下さい。")) return false;
	if (!necessaryInputCheck(document.forms[0].City_K,"Text","市区郡（漢字）を入力して下さい。")) return false;
	if (!necessaryInputCheck(document.forms[0].MailAddress,"Text","メールアドレスを入力して下さい。")) return false;
	if (!necessaryInputCheck(document.forms[0].MailAddressConf,"Text","確認用メールアドレスを入力して下さい。")) return false;
	if (!necessaryInputCheck(document.forms[0].Birthday,"Text","誕生日を入力して下さい。")) return false;

	if (!stringCheckWithAlert(document.forms[0].HomeTelphoneNo,"telephone","自宅電話番号は半角数字または - を入力して下さい。")) return false;
	if (!stringCheckWithAlert(document.forms[0].PortableTelphoneNo,"telephone","携帯電話番号は半角数字または - を入力して下さい。")) return false;

	if (!datecheck(document.forms[0].Birthday,"誕生日を正しい日付で入力されていますか？\n\n・yyyy/mm/dd 形式で入力して下さい。\n・半角で入力してください。")) return false;
//	if (!necessaryInputCheck(document.forms[0].site_aim,"Radio","アンケートにお答えください")) return false;
<%
	UnAge_input_check = year(now) - 15
%>
	nen = document.forms[0].Birthday.value
	AGE = nen.split("/");
	if (AGE[0] >= <%= UnAge_input_check %>){
		if (AGE[0] == <%= UnAge_input_check %>){
			if (AGE[1] == "01" | AGE[1] == "02" | AGE[1] == "03"){
			}else{
				 alert('労働基準法に基づき、15才未満の方のご登録はご遠慮ください。');
				 document.forms[0].Birthday.focus();
				  return false;
			}
		}else{
			 alert('労働基準法に基づき、15才未満の方のご登録はご遠慮ください。');
			 document.forms[0].Birthday.focus();
			  return false;
		}
	}

	if (!necessaryInputCheck(document.forms[0].Sex,"Radio","性別を選択してください")) return false;

// メールアドレスチェック
	if (document.forms[0].MailAddress.value != document.forms[0].MailAddressConf.value)
	{
	 alert('メールアドレスが一致しません');
	  return false;
	}
	if (!mailAddressCheck(document.forms[0].MailAddress,'メールアドレスが正しく入力されていません。') ) return false;		//@@@2003/01/17 TASC Uda Add

// 職種チェック
	fObj = document.forms[0];
	if (fObj.occupationCode1.value == "") {
		alert('希望職種を選択してください');
		fObj.occupationButton1.focus();
		return false;
	}

//勤務形態チェック
	if ((fObj.HopeWorkingType1.selectedIndex == 0) & (fObj.HopeWorkingType2.selectedIndex == 0) & (fObj.HopeWorkingType3.selectedIndex == 0)) {
		alert('希望勤務形態を最低一つ選択してください');
		fObj.HopeWorkingType1.focus();
		return false;
	}

//勤務地チェック
	if (fObj.HopePrefecture1.selectedIndex == 0){
		alert('希望勤務地を選択してください');
		fObj.HopePrefecture1.focus();
		return false;
	}

//パスワードチェック
	if (document.forms[0].password.value != document.forms[0].passwordconf.value)
	{
		 alert('パスワードが違います');
		 return false;
	}
	else if (document.forms[0].password.value == "")
	{
		 alert('希望するパスワードを入力してください');
		 return false;
	}
	else if (!stringCheck(document.forms[0].password))
	{
			//@@@2003/01/17 TASC Uda Mod alert('パスワードに「ブランク」は使用できません。\n再度パスワードの設定をお願いします');
		alert('パスワード半角英数字で入力してください。\nまた、「ブランク」は使用できません。');

		 return false;
	}
	else
	{
/***********************************/
/*		入力妥当性チェック */
/***********************************/
		if (document.forms[0].password.value.length < 3) {
			alert("パスワードは3文字以上入力してください。");
			return false;
		}
		if (document.forms[0].password.value.length > 20) {
			alert("パスワードは20文字以下で入力してください。");
			return false;
		}
	}
/***********************************/
/*		市区郡文字数チェック			 */
/***********************************/

	Br = navigator.appName;
	City_K_Mojisu = document.forms[0].City_K.value.length
	if (Br == "Netscape") {
		City_K_Mojisu = City_K_Mojisu / 2
		if (City_K_Mojisu > 7) {
				document.forms[0].City_K.focus();
			//alert('住所・市区郡の項目には「市区郡」までを入力してください。\n市区郡以降につきましては、詳細ページで入力できます。');
			if (window.confirm('市区郡欄の入力が長いようです。\n住所・市区郡の項目は「市区郡」までを入力してください。\nこのまま進みますか？')==false){
			return false;
			}
		}
	} else {
		if (City_K_Mojisu > 7) {
			//alert('住所・市区郡の項目には「市区郡」までを入力してください。\n市区郡以降につきましては、詳細ページで入力できます。');
				document.forms[0].City_K.focus();
			if (window.confirm('市区郡欄の入力が長いようです。\n住所・市区郡の項目は「市区郡」までを入力してください。\nこのまま進みますか？')==false){
			return false;
			}
		}
	}

document.forms[0].submit()
}


function WorkingTypeCheck(type) {
	target = type[type.selectedIndex].text;


	if ( target == '派遣' ) {
		result = confirm('リス株式会社に派遣社員として仮登録されます。\nその後、来社本登録が必要ですが、\nよろしいですか？');
		if ( !result ) {
			type.selectedIndex = 0;
		}
	}

	wObj = document.forms[0];
	wObj1 = wObj.HopeWorkingType1;	// 希望勤務形態1
	wObj2 = wObj.HopeWorkingType2;	// 希望勤務形態2
	wObj3 = wObj.HopeWorkingType3;	// 希望勤務形態3
	// 全部選択されていない状態、3つのうちいずれかに「派遣」「正社員」「契約社員」が選択されている状態の場合は、正社員紹介予定派遣の希望を選択可能
	if ( ( wObj1.selectedIndex == 0 & wObj2.selectedIndex == 0 & wObj3.selectedIndex == 0 ) || ( wObj1[wObj1.selectedIndex].text == '派遣' || wObj1[wObj1.selectedIndex].text == '正社員' || wObj1[wObj1.selectedIndex].text == '契約社員' ) || ( wObj2[wObj2.selectedIndex].text == '派遣' || wObj2[wObj2.selectedIndex].text == '正社員' || wObj2[wObj2.selectedIndex].text == '契約社員' ) || ( wObj3[wObj3.selectedIndex].text == '派遣' || wObj3[wObj3.selectedIndex].text == '正社員' || wObj3[wObj3.selectedIndex].text == '契約社員' )) {
		wObj.TempTo[0].disabled = false;
		wObj.TempTo[1].disabled = false;
	} else {
		wObj.TempTo[0].checked = false;
		wObj.TempTo[1].checked = true;
		wObj.TempTo[0].disabled = true;
		wObj.TempTo[1].disabled = true;
	}
}


function setOccupation(count,flag) {
	document.forms[0].occupationCount.value = count;
	document.forms[0].occupationFlag.value = flag;
	selectWin = window.open("./../select_jobtype.asp?regflag=1","list","width=700,height=300,resizable=yes");
}

function setDefaultValue()
{
<%
If (Session("usertype") = "" And Request.Form("PersonCode") <> "") Then
%>
	document.forms[0].OparateClass_Web.selectedIndex = <%= oRS_input_check.Fields("OperateClassWebCode").Value %>;
	document.forms[0].Prefecture.selectedIndex = <%= oRS_input_check.Fields("PrefectureCode").Value %>;
<%
End If
If Request.Form("pid") <> "" Then
	'▼▼▼ 希望勤務形態取得
	sSQL_input_check = "sp_GetWorkingType '" & Session("userid") & "'"
	Set oRS_index_check2 = dbconn.Execute(sSQL_input_check)
	idx_input_check = 0
	Do Until oRS_index_check2.EOF
		aHopeWorkingType_input_check(idx_input_check) = oRS_index_check2.Fields("WorkingTypeCode").Value
		idx_input_check = idx_input_check + 1
		If idx_input_check = 3 Then Exit Do
	Loop
	oRS_index_check2.Close
	'▲▲▲ 希望勤務形態取得

	'▼▼▼ 希望勤務都道府県取得
	sSQL_input_check = "sp_GetPrefecture '" & Session("userid") & "'"
	Set oRS_index_check2 = dbconn.Execute(sSQL_input_check)
	If Not oRS_index_check2.EOF Then
		sHopePrefectureCode = oRS_index_check2.Fields("PrefectureCode").Value
	End If
	oRS_index_check2.Close
	'▲▲▲ 希望勤務都道府県取得
%>
		document.forms[0].HopeWorkingType1.selectedIndex = <%= aHopeWorkingType_input_check(0) %>;
		document.forms[0].HopeWorkingType2.selectedIndex = <%= aHopeWorkingType_input_check(1) %>;
		document.forms[0].HopeWorkingType3.selectedIndex = <%= aHopeWorkingType_input_check(2) %>;
		document.forms[0].HopePrefecture1.selectedIndex = <%= sHopePrefectureCode %>;
<%
End If
oRS_input_check.Close
%>
}
