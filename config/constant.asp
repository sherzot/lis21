<%
'************************************************
'*	�Œ�ϐ��錾�t�@�C��						*
'************************************************

'*****���[�����*****
Const Cnt_MailServer = "172.16.1.39"		'���[���T�[�o�[��
Const Cnt_LisMailAddress = "�����ƃi�r�E���X <lis@lis21.co.jp>"		'���X�l��\���[���A�h���X
Const Cnt_NaviMailAddress = "�����ƃi�r <info@shigotonavi.jp>"		'�����ƃi�r�ʒm�p���[���A�h���X

'******************************************************************************
'** �A�v���[�`���[���ɂĎg�p(staff/mailtocompany.asp)
'******************************************************************************
Dim MAIL_URL_STAFF: MAIL_URL_STAFF = "https://www.shigotonavi.co.jp/staff/mailhistory_person.asp"	'���E�҃��[���Ǘ��y�[�W
Dim MAIL_URL_COMPANY: MAIL_URL_COMPANY = "https://www.shigotonavi.co.jp/company/mailhistory_company.asp"	'���l��ƃ��[���Ǘ��y�[�W
Dim MAIL_URL_TALENT: MAIL_URL_TALENT = "https://www.shigotonavi.co.jp/talent/mailhistory_talent.asp"		'�l�ޕۗL��ƃ��[���Ǘ��y�[�W

'************************************************
'*	���[�����M���̎Q�Ɛ�URL						*
'************************************************

'************************************************
'*	���E��(staff)								*
'************************************************
'�V�K�X�^�b�t�o�^������эX�V���Ƀ��X�a�փ��[�����M�i�h���̏ꍇ�̂݁j
'(LISPROJE/staff/RegisterStaff_sendmail.asp)
Const Cnt_RegistMail_StaffURL = "http://bi-b1.lis21.co.jp/INCLUDE/Staff_detail.asp"
'���{�ԓ����O��Cnt_RegistMail_StaffURL�͕K��
'"http://bi.lis21.co.jp/INCLUDE/Staff_detail.asp"�i�{�ԗp�j�ɂ��Ă��������B(�h�����o�^�҂̓��͒ʒm���[���p�A�h���X)

'************************************************
'*	���l���(company)							*
'************************************************

'************************************************
'*	�l�ޕۗL���(talent)						*
'************************************************
'�l�ޕۗL��Ƃ���Џ����X�V�����ۂɃ��X�a�փ��[�����M
'(LISPROJE/talent/t_company_regist_sendmail.asp)
'@@@2003/02/27 TASc Uda Del		Cnt_RegistMail_TCompanyURL = "https://www.shigotonavi.co.jp/talent/t_company_regist.asp"

'�l�ޕۗL��Ƃ��l�ޏ����X�V�����ۂɃ��X�a�փ��[�����M
'(LISPROJE/talent/t_person_reg_sendmail.asp)
'@@@2003/02/27 TASC Uda Del		Cnt_RegistMail_TPersonURL = "https://www.shigotonavi.co.jp/talent/t_person_reg1l.asp"

'******************************************************************************
'** �X�^�b�t�o�^�m�F���[��(staff/person_reg1_register.asp)
'******************************************************************************
'�^�C�g��
Const MAIL_STAFFREG_SUBJECT = "�y�����ƃi�r�z���o�^�����̂��ē� ���������쐬�̎菇���f��"
'���[���{��
Dim MAIL_STAFFREG_BODY: MAIL_STAFFREG_BODY = "" & _
	"���l�T�C�g�u�����ƃi�r�v���^�c���Ă��郊�X������Ђł��B" & vbCrLf & _
	"���̓x�͂����ƃi�r�ւ̂��o�^���肪�Ƃ��������܂����B" & vbCrLf & _
	"�M���l�̓o�^�́A���������v���܂����B" & vbCrLf & _
	vbCrLf & _
	"����A�u�����ƃi�r�v�̃T�[�r�X���j���[�������p�����������߁A" & vbCrLf & _
	"���L�̂h�c�ƃp�X���[�h�𔭍s�v���܂��B��؂ɕۊǂ��Ă��������B" & vbCrLf & _
	vbCrLf & _
	"�y�M���l�̂h�c�z������������������������������������������������������"
'���[���t�b�^
Dim MAIL_STAFFREG_FOOTER: MAIL_STAFFREG_FOOTER = "" & _
	"���u�x�~�E�މ�v�ɂ���" & vbCrLf & vbCrLf & _
	"�����ƃi�r�̂����p���K�v�ȏꍇ�́A��LID�ƃp�X���[�h�Ń��O�C����A" & vbCrLf & _
	"���j���[���́u�x�~�E�މ�v�������ĉ������B" & vbCrLf & vbCrLf & _
	"�����ƃi�r�Ɋւ��Ă��s���ȓ_���������܂�����A" & vbCrLf & _
	"���萔�ł͂������܂������L�܂Ń��[���ɂĂ��⍇�����������B" & vbCrLf & vbCrLf & _
	"����������������������������������������������������������������������" & vbCrLf & vbCrLf & _
	"�͂��炭�l�̃\�[�V�����R�~���j�e�B�[�u�����ƃi�r�v" & vbCrLf & _
	"�^�c��ЁF���X�������" & vbCrLf & _
	"http://www.shigotonavi.co.jp/" & vbCrLf & _
	"���₢���킹�Flis@lis21.co.jp"

'******************************************************************************
'** �A�t�B���G�C�g�o�^�җp�F�؃��[��(staff/person_reg1_register.asp)
'******************************************************************************
Dim MAIL_STAFFREG_AFFILIATE_HEADER: MAIL_STAFFREG_AFFILIATE_HEADER = "" & _
	"���l��āu��������ށv��ؽ������Ђł��B" & vbCrLf & _
	"���̓x�͂�������ނւ̂��o�^���肪�Ƃ��������܂����B" & vbCrLf & _
	"���L��URL��د�����ƁAҰق̔F�؂������������܂��B" & vbCrLf & _
	"--------------------" & vbCrLf & vbCrLf & _
	"�F�؊m��y�[�W��" & vbCrLf

Dim MAIL_STAFFREG_AFFILIATE_FOOTER: MAIL_STAFFREG_AFFILIATE_FOOTER = "" & _
	"--------------------" & vbCrLf & _
	"�͂��炭�l�̃\�[�V�����R�~���j�e�B�[�u�����ƃi�r�v" & vbCrLf & _
	"�^�c��ЁF���X�������" & vbCrLf & _
	"http://www.shigotonavi.co.jp/" & vbCrLf & _
	"���₢���킹�Flis@lis21.co.jp"

'******************************************************************************
'** �y���E�҂��狁�l��Ƃւ̃A�v���[�`���[�� (staff/mailtocompany.asp�ɂĎg�p)�z
'******************************************************************************
'�^�C�g��
Const MAIL_FROM_STAFF_SUBJECT = "�y�����ƃi�r�z���E�҂��烁�[�����͂��܂���"
'���[���{��
Dim MAIL_FROM_STAFF_BODY: MAIL_FROM_STAFF_BODY = "" & _
	"�����u�����ƃi�r�v�������p���������܂��Ă��肪�Ƃ��������܂��B" & vbCrLf & vbCrLf & _
	"�u�����ƃi�r�v���^�c���Ă���܂����X������Ђł��B" & vbCrLf & vbCrLf & vbCrLf & _
	"�M�Ђ̋��l���֋��E�҂��牞�傪����܂����̂ł��m�点�v���܂��B" & vbCrLf & vbCrLf & vbCrLf & _
	"��������e�́A���L��URL��育���������B" & vbCrLf & vbCrLf & _
	"�����E�҂̕��ւ��A�����̂��Ή����X�������肢�\���グ�܂��B" & vbCrLf & vbCrLf & vbCrLf & _
	"���E�҂���̃��[�����e�͂����炩�炲�m�F��������" & vbCrLf & _
	"����"
'���[���t�b�^
Dim MAIL_FROM_STAFF_FOOTER: MAIL_FROM_STAFF_FOOTER = "" & vbCrLf & vbCrLf & _
	"-------------------------------" & vbCrLf & _
	"�͂��炭�l�̃\�[�V�����R�~���j�e�B�[�u�����ƃi�r�v" & vbCrLf & _
	"�^�c��ЁF���X�������" & vbCrLf & _
	"http://www.shigotonavi.co.jp/" & vbCrLf & _
	"���₢���킹�Flis@lis21.co.jp"

'******************************************************************************
'** �y���E�҂��狁�l��Ƃւ̃A�v���[�`���[�� (staff/mailtocompany.asp�ɂĎg�p)�z
'******************************************************************************
'�^�C�g��
Const MAIL_FROM_COMPANY_SUBJECT = "�y�����ƃi�r�z���l��Ƃ��烁�[�����͂��܂���"
'���[���{��
Dim MAIL_FROM_COMPANY_BODY: MAIL_FROM_COMPANY_BODY = "" & _
	"���������p���������܂��Ă��肪�Ƃ��������܂��B" & vbCrLf & _
	"���o�^���������Ă���܂��u�����ƃi�r�v�i���X������Ёj�ł��B"  & vbCrLf & vbCrLf & _
	"��قǁA���l��Ƃ���X�J�E�g�E�A�����[�������a���肵�܂����B" & vbCrLf & _
	"�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P" & vbCrLf & _
	"�����[�������������ɂ͉���URL�����O�C�������[���Ǘ������[���^�C�g��" & vbCrLf & _
	"�@���N���b�N���ĉ������B" & vbCrLf & _
	"���ԐM�́A��L���@�Ń��[���������ɂȂ�u�ԐM�v�{�^������A" & vbCrLf & _
	"�@���[�����쐬�E���M�ł��܂��B" & vbCrLf & _
	"����Ƃ���̑�؂ȃ��[���ł��B���Ђ��Ԏ������肢�������܂��B" & vbCrLf & vbCrLf & _
	"���C�ɂȂ�u���[�����e�v�Ɓu�ԐM�v�͂�������N���b�N���ĉ������I"
'���[���t�b�^
Dim MAIL_FROM_COMPANY_FOOTER: MAIL_FROM_COMPANY_FOOTER = "" & vbCrLf & vbCrLf & _
	"�����̃��[���ɒ��ڕԐM����Ă���Ƃɂ͓͂��܂���̂ł����Ӊ������B��" & vbCrLf & vbCrLf & _
	"���X�J�E�g�����Ȃ��I����Ȏ��́A�o�^��񂪑���Ȃ����Ƃ�����܂��B" & vbCrLf & _
	"�@������x�o�^�����������A�ǉ����ĉ������B" & vbCrLf & _
	"����Ƃ֒��ڃA�v���[�`���ł��܂��B���O�C�����d���������E�\�����A" & vbCrLf & _
	"�@�u���̉�Ђփ��[���𑗐M����v�{�^�����烁�[���������A�ϋɓI��" & vbCrLf & _
	"�@�A�v���[�`���ĉ������B" & vbCrLf & _
	"�������ƃi�r�̃��X�ł́A���d���̂��Љ���s�Ȃ��Ă���܂��B" & vbCrLf & _
	"�@���d�����������ہA�d�b�⃁�[���ɂĘA������ő��T�[�r�X�����Ă���܂��B" & vbCrLf & _
	"���X�J�E�g�A�����[���́A�����ƃi�r�������p�̂��ׂĂ̊F�l�ɓ͂�" & vbCrLf & _
	"�@�`�����X������܂��B" & vbCrLf & vbCrLf & _
	"�EID��p�X���[�h�Ȃǂ�������Ȃ����͂����炩�����ł��܂���" & vbCrLf & _
	"https://www.shigotonavi.co.jp/staff/passwordreminder.asp" & vbCrLf & vbCrLf & _
	"��--------------------------------------------" & vbCrLf & _
	"�͂��炭�l�̃\�[�V�����R�~���j�e�B�[�u�����ƃi�r�v" & vbCrLf & _
	"�^�c��ЁF���X�������" & vbCrLf & _
	"http://www.shigotonavi.co.jp/" & vbCrLf & _
	"���₢���킹�Flis@lis21.co.jp"

'�y��ƊԃR���{�ł̃A�v���[�`���[��(/LISPROJE/sendmail_company.asp�ɂĎg�p)�z
'�^�C�g��
Const Cnt_sendmail_Advert_Subject = "�y�����ƃi�r�z�R���{���[�V�����̃��[�����͂��܂����B"
'���[���{��
Dim Cnt_sendmail_Advert_Body
Cnt_sendmail_Advert_Body = "���������p���肪�Ƃ��������܂��B" & vbCrLf & _
	"�����ƃi�r�ł��B" & vbCrLf & vbCrLf & _
	"�����ƃi�r�̃R���{�@�\�𗘗p���ċM�Ђփ��[�����͂��Ă���܂��B" & vbCrLf & _
	"���m�F�̏�A���Ԏ��Ȃǂ̂��Ή��̒��A�X�������肢�\���グ�܂��B" & vbCrLf & _
	"���[�����e�͂����炩�火"
'���[���t�b�^
Dim Cnt_sendmail_Advert_Fut
Cnt_sendmail_Advert_Fut = vbCrLf & vbCrLf & "-------------------------------" & vbCrLf & _
	"�͂��炭�l�̃\�[�V�����R�~���j�e�B�[�u�����ƃi�r�v" & vbCrLf & _
	"�^�c��ЁF���X�������" & vbCrLf & _
	"http://www.shigotonavi.co.jp/" & vbCrLf & _
	"���₢���킹�Flis@lis21.co.jp"


'�y���E�҃R���{�ł̃A�v���[�`���[���z
'�^�C�g��
Const Cnt_sendmail_Advert2_Subject = "�y�����ƃi�r�z�R���{���[�V�����̃��[�����͂��܂����B"
'���[���{��
Dim Cnt_sendmail_Advert2_Body
Cnt_sendmail_Advert2_Body = "���������p���肪�Ƃ��������܂��B" & vbCrLf & _
	"�����ƃi�r�ł��B" & vbCrLf & _
	"�����ƃi�r�����g���̋��E�ҁi�������͐l�ޕۗL��Ɓj���炠�Ȃ��ցA" & vbCrLf & _
	"�����ƃi�r�R���{���[�V�����@�\�𗘗p���ă��[�����͂��Ă���܂��B" & vbCrLf & _
	"���m�F�̏�A���Ԏ��Ȃǂ̂��Ή��̒��A�X�������肢�\���グ�܂��B" & vbCrLf & _
	"���[�����e�͂����炩�火"
'���[���t�b�^
Dim Cnt_sendmail_Advert2_Fut
Cnt_sendmail_Advert2_Fut = vbCrLf & vbCrLf & "���R���{���[�V�����@�\" & vbCrLf & _
	"�u���Ɓv�u�����J���v�̈Ӗ��B�����ƃi�r�ł͂r�n�g�n�̕���A" & vbCrLf & _
	"�Ɨ��u���̕��A�l�ޕۗL��ƂƂ����d�������߂Ă��Ȃ���A" & vbCrLf & _
	"����ŋ����ō�Ƃ��s�Ȃ����Ȃǂ����l�������s�Ȃ����Ƃ��ł��܂��B" & vbCrLf & _
	"���̋@�\���R���{���[�V�����@�\�ƌĂ�ł���܂��B" & vbCrLf & vbCrLf & _
	vbCrLf & _
	"�͂��炭�l�̃\�[�V�����R�~���j�e�B�[�u�����ƃi�r�v" & vbCrLf & _
	"�^�c��ЁF���X�������" & vbCrLf & _
	"http://www.shigotonavi.co.jp/" & vbCrLf & _
	"���₢���킹�Flis@lis21.co.jp"


'�y���E�҃R���{���[�V�����\���������[��(/LISPROJE/Collabo_Entry_Reg.asp�ɂĎg�p)�z
'�^�C�g��
Const Cnt_Collabo_Entry_Subject = "�y�����ƃi�r�z���E�҃R���{���[�V�����\���̂��m�点"
'���[���{��
Dim Cnt_Collabo_Entry_Body
Cnt_Collabo_Entry_Body = "���E�҃R���{���[�V�����̐\��������܂����B" & vbCrLf & vbCrLf & _
	"�S���̕��́u�Г��V�X�e���v���" & vbCrLf & _
	"�u�����ƃi�r�Ǘ��v���u���C�Z���X���ێ�v��" & vbCrLf & _
	"���C�Z���X���s���s�Ȃ��Ă��������B" & vbCrLf & _
	"�i�����_�ł̓��C�Z���X�f�[�^�͓����Ă��܂���j" & vbCrLf & _
	"���s��A�L���\�����𑗕t���Ă��������B" & vbCrLf

'�y���E�҃R���{���[�V�����p�����[��(/LISPROJE/Collabo_Entry_Reg.asp�ɂĎg�p)�z
'�^�C�g��
Const Cnt_Collabo_keizoku_Subject = "�y�����ƃi�r�z���E�҃R���{���[�V�����p���\��"
'���[���{��
Dim Cnt_Collabo_keizoku_Body
Cnt_Collabo_keizoku_Body = "���E�҃R���{���[�V�����̌p���\��������܂����B" & vbCrLf & _
	"���łɌp�����C�Z���X�͔��s�ς݂ł��̂ŁA�m�F�Ɛ\�����̑��t�����肢�������܂��B" & vbCrLf & vbCrLf & _
	"���łɃ��C�Z���X���͓����Ă���܂��B" & vbCrLf & _
	"�S���̕��͏��m�F��A�L���\�����𑗕t���Ă��������B" & vbCrLf & _
	"�u�Г��V�X�e���v���u�����ƃi�r�Ǘ��v���u���C�Z���X���ێ�v" & vbCrLf & _
	"�Ŋm�F���\�ł��B" & vbCrLf


'�y��ƊԃR���{���[�V�����\���������[��(/LISPROJE/company/Collabo_Entry_Reg.asp�ɂĎg�p)�z
'�^�C�g��
Const Cnt_Company_Collabo_Entry_Subject = "�y�����ƃi�r�z�r�W�l�X�R���{���[�V�����\���̂��m�点"
'���[���{��
Dim Cnt_Company_Collabo_Entry_Body
Cnt_Company_Collabo_Entry_Body = "�r�W�l�X�R���{���[�V�����̐\��������܂����B" & vbCrLf & vbCrLf & _
	"�S���̕��́u�Г��V�X�e���v���" & vbCrLf & _
	"�u�����ƃi�r�Ǘ��v���u���C�Z���X���ێ�v��" & vbCrLf & _
	"���C�Z���X���s���s�Ȃ��Ă��������B" & vbCrLf & _
	"�i�����_�ł̓��C�Z���X�f�[�^�͓����Ă��܂���j" & vbCrLf & _
	"���s��A�L���\�����𑗕t���Ă��������B" & vbCrLf

'�y��ƊԃR���{���[�V�����p�����[��(/LISPROJE/company/Collabo_Entry_Reg.asp�ɂĎg�p)�z
'�^�C�g��
Const Cnt_Company_Collabo_keizoku_Subject = "�y�����ƃi�r�z�r�W�l�X�R���{���[�V�����p���\���̂��m�点"
'���[���{��
Dim Cnt_Company_Collabo_keizoku_Body
Cnt_Company_Collabo_keizoku_Body = "�r�W�l�X�R���{���[�V�����̌p���\��������܂����B" & vbCrLf & _
	"���łɌp�����C�Z���X�͔��s�ς݂ł��̂ŁA�m�F�Ɛ\�������t�����肢�������܂��B" & vbCrLf & vbCrLf & _
	"���łɃ��C�Z���X���͓����Ă���܂��B" & vbCrLf & _
	"�S���̕��͏��m�F��A�L���\�����𑗕t���Ă��������B" & vbCrLf & _
	"�u�Г��V�X�e���v���u�����ƃi�r�Ǘ��v���u���C�Z���X���ێ�v" & vbCrLf & _
	"�Ŋm�F���\�ł��B" & vbCrLf

'�y���l�L���p�����p�z
'�^�C�g��
Const Cnt_Company_keizoku_Subject = "�y�����ƃi�r�z���l�L�����p�p���\�����݂̂��m�点"
'���[���{��
Dim Cnt_Company_keizoku_Body
Cnt_Company_keizoku_Body = "���l�L���̗��p�p���̐\��������܂����B" & vbCrLf & _
	"���łɌp�����C�Z���X�͔��s�ς݂ł��̂ŁA�m�F�Ɛ\�������t�����肢�������܂��B" & vbCrLf & vbCrLf & _
	"���łɃ��C�Z���X���͓����Ă���܂��B" & vbCrLf & _
	"�S���̕��͏��m�F��A�L���\�����𑗕t���Ă��������B" & vbCrLf & _
	"�u�Г��V�X�e���v���u�����ƃi�r�Ǘ��v���u���C�Z���X���ێ�v" & vbCrLf & _
	"�Ŋm�F���\�ł��B" & vbCrLf


'�z�M���~�t�q�k
Dim Cnt_Jinzai_Stop_URL
Cnt_Jinzai_Stop_URL = HTTP_CURRENTURL & "jinzai/jinzai_stop_reg.asp"

'�y���[���}�K�W���o�^�m�F���[��(/LISPROJECT/JINZAI/Jinzai_Reg.asp�ɂĎg�p)�z
'�^�C�g��
Dim Cnt_Jinzai_Entry_Subject
Cnt_Jinzai_Entry_Subject = "���������}�K�u�i�h�m�y�`�h�v�o�^�m�F"
'���[���{��
Dim Cnt_Jinzai_Entry_Body
Cnt_Jinzai_Entry_Body = "�����p���肪�Ƃ��������܂��B" & vbCrLf & _
	"���������}�K�u�i�h�m�y�`�h�v�ł��B" & vbCrLf & vbCrLf & _
	"���[���}�K�W���̓o�^�m�F���[���ł��B" & vbCrLf & _
	"���̓��e�Ń��[���}�K�W����o�^�������܂����B" & vbCrLf
'���[���t�b�^
Dim Cnt_Jinzai_Entry_Fut
Cnt_Jinzai_Entry_Fut = vbCrLf & "-------------------------------" & vbCrLf & _
	"���₢���킹�Flis@lis21.co.jp"

'�y�����ƃi�r�o�^�\�����[��(/LISPROJECT/JINZAI/Jinzai_EntryNavi_Reg.asp�ɂĎg�p)�z
'�^�C�g��
Dim Cnt_Jinzai_Navi_Subject
Cnt_Jinzai_Navi_Subject 	= "�y�����}�K�����ƃi�r�\���݁z"
'���[���{��
Dim Cnt_Jinzai_Navi_Body
Cnt_Jinzai_Navi_Body = "���[���}�K�W���̍w�Ǌ�Ƃ��炵���ƃi�r�o�^��" & vbCrLf & _
	"�\���݂�����܂����B" & vbCrLf & vbCrLf & _
	"�S���̕��͂��̊�ƂɃA�v���[�`���s���Ă��������B" & vbCrLf & _
	"�����W��A�u�Г��V�X�e���v���" & vbCrLf & _
	"�u�����ƃi�r�Ǘ��v���u���p�ҏ��ێ�v�Ŋ�Ə��o�^�A" & vbCrLf & _
	"�@�@�@�@�@�@�@�@�@���u���C�Z���X���ێ�v�Ń��C�Z���X���s�A" & vbCrLf & _
	"�@�@�@�@�@�@�@�@�@���u�F�؂h�c���s�v�Ə�����i�߁A" & vbCrLf & _
	"��Ƃ֗��p�J�n��A�����Ă��������B" & vbCrLf & vbCrLf & _
	"�y�\���̂�������Ƃ͂����火�z" & vbCrLf
'���[���t�b�^
Dim Cnt_Jinzai_Navi_Fut
Cnt_Jinzai_Navi_Fut = ""

'�y���[���}�K�W���o�^�m�F���[��(/LISPROJECT/JINZAI/Jinzai_Reg.asp�ɂĎg�p)�z
	'�^�C�g��
Dim Cnt_Jinzai_GetID_Subject
Cnt_Jinzai_GetID_Subject = "���������}�K�u�i�h�m�y�`�h�v�h�c�̑��t"
	'���[���{��
Dim Cnt_Jinzai_GetID_Body
Cnt_Jinzai_GetID_Body = vbCrLf & "���o�^���肪�Ƃ��������܂��B" & vbCrLf & _
	"�l�ޑ�����u�i�h�m�y�`�h�v�ł��B" & vbCrLf & vbCrLf & _
	"�ȉ��̒ʂ育�o�^���󂯕t���v���܂����̂łh�c�������t" & vbCrLf & _
	"�����Ă��������܂��B" & vbCrLf & vbCrLf & _
	"���M���l�̂h�c�F"
	'���[���t�b�^
Dim Cnt_Jinzai_GetID_Fut
Cnt_Jinzai_GetID_Fut = vbCrLf & "��L�h�c���ȉ��̕��@�ł��ݒ艺�����B" & vbCrLf & vbCrLf & _
	"1.�u�i�h�m�y�`�h�I�v�𗧂��グ��" & vbCrLf & _
	"2.��ʍ���́u�t�@�C���v���u�ݒ�v���N���b�N" & vbCrLf & _
	"3.���[���}�K�W���h�c�ɏ�L�h�c��o�^����" & vbCrLf & vbCrLf & _
	"�ȏ�Ŋ����ł��B" & vbCrLf & vbCrLf & _
	"����A��ЂɓK�C�̗D�G�Ȑl�ނ����l���������B" & vbCrLf & vbCrLf & _
	"-------------------------------" & vbCrLf & _
	"���₢���킹" & vbCrLf & _
	"���X������Ё@Web�헪��" & vbCrLf & _
	"���[���A�h���X�@lis@lis21.co.jp" & vbCrLf

'�y�\�t�g�o�^�m�F���[��(/LISPROJECT/JINZAI/JinzaiSoft_Reg2.asp�ɂĎg�p)�z
Dim Cnt_JinzaiSoft_Entry_Body
Cnt_JinzaiSoft_Entry_Body = "�����p���肪�Ƃ��������܂��B" & vbCrLf & _
"�u�i�h�m�y�`�h�I�v�ł��B" & vbCrLf & vbCrLf & _
"�����o�^�m�F���[���ł��B" & vbCrLf & _
"�ȉ��̓��e�œo�^�������܂����B" & vbCrLf

'*****���[������*****
'1�y�[�W����̍ő�\������(/LISPROJE/company/mail_company.asp)
Const Cnt_DispNum = 20

'*** ���C�Z���X���E�p�� ***
'(/LISPROJE/License_Continue.asp�ALicense_ContinueEnd.asp�ɂĎg�p)
Const License_DispNum					= 12	'1�y�[�W����̍ő�\������
Const License_Zeiritu					= 5		'�ŗ�
Const License_CompanySimebi			= 31	'��Ƃ̒��ߓ��i���ߓ������͂̏ꍇ�g�p�j
Const License_PersonSimebi			= 20	'���E�҂̒��ߓ�


'***��ƌ���***
Const Company_List_DispNum			= 10	'1�y�[�W������̕\���ő匏��
Const Company_List_ShowPageNum 		= 10	'�w�萔���y�[�W�ԍ���\��
Const Company_DspModePop				= 1		'�|�b�v�A�b�v
Const Company_DspModeSam				= 2		'�������
Const Company_JyucyuFlag				= 1		'�󒍊֌W�̉�ʂ���БI����ʂֈڍs�����ꍇ�̃t���O


'*** ���E�ғo�^ ***
'���E�ғo�^�ɂ����āA��]�Ζ��`�Ԃɔh�����I������Ă����ꍇ�g�p(/LISPROJE/staff/person_reg5l.asp�ɂĎg�p)
Const Cnt_Haken_OparateClass_Com = "100"
Const Cnt_Haken_OparateClass_ComMoji = "���ʐ�"
%>
