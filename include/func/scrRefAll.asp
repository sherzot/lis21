<%
'*******************************************************************************
'�T�@�v�F���t�@����̓^�O�擾
'���@���F
'�o�@�́F
'�߂�l�FString
'���@�l�F
'���@���F2010/05/11 LIS K.Kokubo �쐬
'*******************************************************************************
Function scrRefAll()
	Dim sScript

	Dim id
	Dim refer

	id = request.querystring("id")	'ID����

	If IsNumeric(id) = True Then
		If id = 1 Then
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume1.js""></script>"
		ElseIf id = 2 Then 'Overture�u�������v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume2.js""></script>"
		ElseIf id = 3 Then
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume3.js""></script>"
		ElseIf id = 4 Then
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume4.js""></script>"
		ElseIf id = 5 Then
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume5.js""></script>"
		ElseIf id = 6 Then 'Google�u�]�E�v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume6.js""></script>"
		ElseIf id = 7 Then 'Google�u�A�E�v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume7.js""></script>"
		ElseIf id = 8 Then 'Google�u���l�v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume8.js""></script>"
		ElseIf id = 9 Then 'Google�u�ʐځv
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume9.js""></script>"
		ElseIf id = 10 Then '�����C���^���N�e�B�u���[��
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume10.js""></script>"
		ElseIf id = 11 Then 'Overture[�E���o����] 2003/09/12ADD
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume11.js""></script>"
		ElseIf id = 12 Then 'JListing[������] 2004/04/28ADD
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume12.js""></script>"
		ElseIf id = 13 Then 'JListing[�E���o����] 2004/04/28ADD
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume13.js""></script>"
		ElseIf id = 14 Then 'Overture[�ʐ�] 2004/07/22ADD
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume14.js""></script>"
		ElseIf id = 15 Then 'Overture[�ސE] 2004/07/22ADD
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume15.js""></script>"
		ElseIf id = 16 Then 'Overture[�������̏�����] 2004/07/22ADD
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume16.js""></script>"
		ElseIf id = 17 Then '�v���W�f���g�r�W���� 2004/09/3�z�M
			Session("ref") = "president"
		ElseIf id = 18 Then 'JINZAI�����}�K 2004/09/6�z�M
			Session("ref") = "jinzai-mail"
		ElseIf id = 19 Then 'Overture�u�l�ށv 2004/09/16
			Session("ref") = "overture_01"
		ElseIf id = 20 Then 'Overture�u���̑��v 2004/09/16
			Session("ref") = "overture_02"
		ElseIf id = 21 Then '�T�C�{�E�Y�@�e�L�X�g�L�� 2004/09/20�`26
			Session("ref") = "cybozu"
		ElseIf id = 22 Then '�����Ə��l�b�g 2004/10/22
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume17.js""></script>"
		ElseIf id = 23 Then 'e-words 2004/11/24
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume18.js""></script>"
		ElseIf id = 24 Then 'Overture�y12���ǉ�C�����z 2004/12/22
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume19.js""></script>"
		ElseIf id = 25 Then 'MSN�̃I�X�X��
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume25.js""></script>"
		ElseIf id = 26 Then 'Adwords�V�ǉ�������
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume1.js""></script>"
			'sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume26.js""></script>"
		ElseIf id = 27 Then 'oveture�V�ǉ�������
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume2.js""></script>"
			'sScript = "<script type=""text/javascript"" src=""/java-script/refer_s_resume27.js""></script>"
		ElseIf id = 28 Then 'oveture�u�u�]���@�v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_column_1_ot.js""></script>"
		ElseIf id = 29 Then 'adwords�u�u�]���@�v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_column_1_aw.js""></script>"
		ElseIf id = 30 Then 'overture�u�E��ʃy�[�W�v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_shoku_search_1_ot.js""></script>"
		ElseIf id = 31 Then 'adwords�u�E��ʃy�[�W�v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_shoku_search_1_aw.js""></script>"
		ElseIf id = 32 Then 'J���X�e�B���O�u�u�]���@�v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_column_1_jlisting.js""></script>"
		ElseIf id = 33 Then 'overture�u�ސE��v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_retire_want_ot.js""></script>"
		ElseIf id = 34 Then 'adword�u�ސE��v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_retire_want_ad.js""></script>"
		ElseIf id = 35 Then 'overture�u����PR�v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_Mypr_ot.js""></script>"
		ElseIf id = 36 Then 'adword�u����PR�v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_Mypr_ad.js""></script>"
		ElseIf id = 37 Then 'overture�u�j�[�g�v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_neet_ot.js""></script>"
		ElseIf id = 38 Then 'adword�u�j�[�g�v
			sScript = "<script type=""text/javascript"" src=""/java-script/refer_neet_ad.js""></script>"
		Else 
			refer = Request.ServerVariables("HTTP_REFERER") 
			If Left(refer,22) <> "http://www.shigotonavi" Then
				If Left(refer,23) <> "https://www.shigotonavi" Then
					If Left(refer,22) <> "www.shigotonavi.co.jp/" Then
						If refer <> "" Then
							sScript = "<script type=""text/javascript"" src=""/java-script/refer_c_hajime.js""></script>"
						End If
					End If
				End If
			End If
		End If
	End If

	scrRefAll = sScript & vbCrLf
End Function
%>
