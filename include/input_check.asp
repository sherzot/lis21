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
/*		���̓`�F�b�N			 */
/***********************************/

	if (!necessaryInputCheck(document.forms[0].Name_K_f,"Text","���i�����j����͂��ĉ������B")) return false;
	if (!necessaryInputCheck(document.forms[0].Name_F_f,"Text","���i�t���K�i�j����͂��ĉ������B")) return false;
	if (!necessaryInputCheck(document.forms[0].Name_K_b,"Text","���i�����j����͂��ĉ������B")) return false;
	if (!necessaryInputCheck(document.forms[0].Name_F_b,"Text","���i�t���K�i�j����͂��ĉ������B")) return false;
	if (!necessaryInputCheck(document.forms[0].Post_U,"Text","�X�֔ԍ���S�ē��͂��ĉ������B")) return false;
	if (!necessaryInputCheck(document.forms[0].Post_L,"Text","�X�֔ԍ���S�ē��͂��ĉ������B")) return false;
	if (!necessaryInputCheck(document.forms[0].Prefecture,"Select","�s���{������͂��ĉ������B")) return false;
	if (!necessaryInputCheck(document.forms[0].City_K,"Text","�s��S�i�����j����͂��ĉ������B")) return false;
	if (!necessaryInputCheck(document.forms[0].MailAddress,"Text","���[���A�h���X����͂��ĉ������B")) return false;
	if (!necessaryInputCheck(document.forms[0].MailAddressConf,"Text","�m�F�p���[���A�h���X����͂��ĉ������B")) return false;
	if (!necessaryInputCheck(document.forms[0].Birthday,"Text","�a��������͂��ĉ������B")) return false;

	if (!stringCheckWithAlert(document.forms[0].HomeTelphoneNo,"telephone","����d�b�ԍ��͔��p�����܂��� - ����͂��ĉ������B")) return false;
	if (!stringCheckWithAlert(document.forms[0].PortableTelphoneNo,"telephone","�g�ѓd�b�ԍ��͔��p�����܂��� - ����͂��ĉ������B")) return false;

	if (!datecheck(document.forms[0].Birthday,"�a�����𐳂������t�œ��͂���Ă��܂����H\n\n�Eyyyy/mm/dd �`���œ��͂��ĉ������B\n�E���p�œ��͂��Ă��������B")) return false;
//	if (!necessaryInputCheck(document.forms[0].site_aim,"Radio","�A���P�[�g�ɂ�������������")) return false;
<%
	UnAge_input_check = year(now) - 15
%>
	nen = document.forms[0].Birthday.value
	AGE = nen.split("/");
	if (AGE[0] >= <%= UnAge_input_check %>){
		if (AGE[0] == <%= UnAge_input_check %>){
			if (AGE[1] == "01" | AGE[1] == "02" | AGE[1] == "03"){
			}else{
				 alert('�J����@�Ɋ�Â��A15�˖����̕��̂��o�^�͂��������������B');
				 document.forms[0].Birthday.focus();
				  return false;
			}
		}else{
			 alert('�J����@�Ɋ�Â��A15�˖����̕��̂��o�^�͂��������������B');
			 document.forms[0].Birthday.focus();
			  return false;
		}
	}

	if (!necessaryInputCheck(document.forms[0].Sex,"Radio","���ʂ�I�����Ă�������")) return false;

// ���[���A�h���X�`�F�b�N
	if (document.forms[0].MailAddress.value != document.forms[0].MailAddressConf.value)
	{
	 alert('���[���A�h���X����v���܂���');
	  return false;
	}
	if (!mailAddressCheck(document.forms[0].MailAddress,'���[���A�h���X�����������͂���Ă��܂���B') ) return false;		//@@@2003/01/17 TASC Uda Add

// �E��`�F�b�N
	fObj = document.forms[0];
	if (fObj.occupationCode1.value == "") {
		alert('��]�E���I�����Ă�������');
		fObj.occupationButton1.focus();
		return false;
	}

//�Ζ��`�ԃ`�F�b�N
	if ((fObj.HopeWorkingType1.selectedIndex == 0) & (fObj.HopeWorkingType2.selectedIndex == 0) & (fObj.HopeWorkingType3.selectedIndex == 0)) {
		alert('��]�Ζ��`�Ԃ��Œ��I�����Ă�������');
		fObj.HopeWorkingType1.focus();
		return false;
	}

//�Ζ��n�`�F�b�N
	if (fObj.HopePrefecture1.selectedIndex == 0){
		alert('��]�Ζ��n��I�����Ă�������');
		fObj.HopePrefecture1.focus();
		return false;
	}

//�p�X���[�h�`�F�b�N
	if (document.forms[0].password.value != document.forms[0].passwordconf.value)
	{
		 alert('�p�X���[�h���Ⴂ�܂�');
		 return false;
	}
	else if (document.forms[0].password.value == "")
	{
		 alert('��]����p�X���[�h����͂��Ă�������');
		 return false;
	}
	else if (!stringCheck(document.forms[0].password))
	{
			//@@@2003/01/17 TASC Uda Mod alert('�p�X���[�h�Ɂu�u�����N�v�͎g�p�ł��܂���B\n�ēx�p�X���[�h�̐ݒ�����肢���܂�');
		alert('�p�X���[�h���p�p�����œ��͂��Ă��������B\n�܂��A�u�u�����N�v�͎g�p�ł��܂���B');

		 return false;
	}
	else
	{
/***********************************/
/*		���͑Ó����`�F�b�N */
/***********************************/
		if (document.forms[0].password.value.length < 3) {
			alert("�p�X���[�h��3�����ȏ���͂��Ă��������B");
			return false;
		}
		if (document.forms[0].password.value.length > 20) {
			alert("�p�X���[�h��20�����ȉ��œ��͂��Ă��������B");
			return false;
		}
	}
/***********************************/
/*		�s��S�������`�F�b�N			 */
/***********************************/

	Br = navigator.appName;
	City_K_Mojisu = document.forms[0].City_K.value.length
	if (Br == "Netscape") {
		City_K_Mojisu = City_K_Mojisu / 2
		if (City_K_Mojisu > 7) {
				document.forms[0].City_K.focus();
			//alert('�Z���E�s��S�̍��ڂɂ́u�s��S�v�܂ł���͂��Ă��������B\n�s��S�ȍ~�ɂ��܂��ẮA�ڍ׃y�[�W�œ��͂ł��܂��B');
			if (window.confirm('�s��S���̓��͂������悤�ł��B\n�Z���E�s��S�̍��ڂ́u�s��S�v�܂ł���͂��Ă��������B\n���̂܂ܐi�݂܂����H')==false){
			return false;
			}
		}
	} else {
		if (City_K_Mojisu > 7) {
			//alert('�Z���E�s��S�̍��ڂɂ́u�s��S�v�܂ł���͂��Ă��������B\n�s��S�ȍ~�ɂ��܂��ẮA�ڍ׃y�[�W�œ��͂ł��܂��B');
				document.forms[0].City_K.focus();
			if (window.confirm('�s��S���̓��͂������悤�ł��B\n�Z���E�s��S�̍��ڂ́u�s��S�v�܂ł���͂��Ă��������B\n���̂܂ܐi�݂܂����H')==false){
			return false;
			}
		}
	}

document.forms[0].submit()
}


function WorkingTypeCheck(type) {
	target = type[type.selectedIndex].text;


	if ( target == '�h��' ) {
		result = confirm('���X������Ђɔh���Ј��Ƃ��ĉ��o�^����܂��B\n���̌�A���Ж{�o�^���K�v�ł����A\n��낵���ł����H');
		if ( !result ) {
			type.selectedIndex = 0;
		}
	}

	wObj = document.forms[0];
	wObj1 = wObj.HopeWorkingType1;	// ��]�Ζ��`��1
	wObj2 = wObj.HopeWorkingType2;	// ��]�Ζ��`��2
	wObj3 = wObj.HopeWorkingType3;	// ��]�Ζ��`��3
	// �S���I������Ă��Ȃ���ԁA3�̂��������ꂩ�Ɂu�h���v�u���Ј��v�u�_��Ј��v���I������Ă����Ԃ̏ꍇ�́A���Ј��Љ�\��h���̊�]��I���\
	if ( ( wObj1.selectedIndex == 0 & wObj2.selectedIndex == 0 & wObj3.selectedIndex == 0 ) || ( wObj1[wObj1.selectedIndex].text == '�h��' || wObj1[wObj1.selectedIndex].text == '���Ј�' || wObj1[wObj1.selectedIndex].text == '�_��Ј�' ) || ( wObj2[wObj2.selectedIndex].text == '�h��' || wObj2[wObj2.selectedIndex].text == '���Ј�' || wObj2[wObj2.selectedIndex].text == '�_��Ј�' ) || ( wObj3[wObj3.selectedIndex].text == '�h��' || wObj3[wObj3.selectedIndex].text == '���Ј�' || wObj3[wObj3.selectedIndex].text == '�_��Ј�' )) {
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
	'������ ��]�Ζ��`�Ԏ擾
	sSQL_input_check = "sp_GetWorkingType '" & Session("userid") & "'"
	Set oRS_index_check2 = dbconn.Execute(sSQL_input_check)
	idx_input_check = 0
	Do Until oRS_index_check2.EOF
		aHopeWorkingType_input_check(idx_input_check) = oRS_index_check2.Fields("WorkingTypeCode").Value
		idx_input_check = idx_input_check + 1
		If idx_input_check = 3 Then Exit Do
	Loop
	oRS_index_check2.Close
	'������ ��]�Ζ��`�Ԏ擾

	'������ ��]�Ζ��s���{���擾
	sSQL_input_check = "sp_GetPrefecture '" & Session("userid") & "'"
	Set oRS_index_check2 = dbconn.Execute(sSQL_input_check)
	If Not oRS_index_check2.EOF Then
		sHopePrefectureCode = oRS_index_check2.Fields("PrefectureCode").Value
	End If
	oRS_index_check2.Close
	'������ ��]�Ζ��s���{���擾
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
