<%
Function getWarmreceptionMaster()
	Dim aData(16)	'���i�E���薼�őI�ԃf�B�N�V���i���z��
	Dim idx

	idx = -1

	'<�e���v���[�g>
	'idx = idx + 1
	'Set aData(idx) = Server.CreateObject("scripting.dictionary")
	'aData(idx).Add "Category","license"
	'aData(idx).Add "ID",""
	'aData(idx).Add "�D��",""
	'aData(idx).Add "���",""
	'aData(idx).Add "����",""
	'aData(idx).Add "�T�v",""
	'aData(idx).Add "�ڍ�",""
	'aData(idx).Add "��p",""
	'aData(idx).Add "���i��",""
	'aData(idx).Add "�c�̖�",""
	'aData(idx).Add "����@��",""
	'aData(idx).Add "�u�����e",""
	'aData(idx).Add "���i",""
	'aData(idx).Add "���T",""
	'aData(idx).Add "�N�[�|��",""
	'aData(idx).Add "�ΏۋƎ�",""
	'aData(idx).Add "���i2",""
	'aData(idx).Add "�T�v2",""
	'aData(idx).Add "���T2",""
	'</�e���v���[�g>

	'<�鏑����F�莎��1��>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0401"
	aData(idx).Add "�D��",""
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","�鏑����F�莎���i1���j"
	aData(idx).Add "����TITLE","�鏑����F�莎���i1���j"
	aData(idx).Add "�T�v","�y�S�Ǝ�Ή��̎��i�z" & vbCrLf & vbCrLf & "�w�����Ȋw�ȔF��̌��I���i�x�Ƃ��āA�鏑�݂̂Ȃ炸�����E�S�ʂ̍����X�L�����ؖ��ł��鎑�i"
	aData(idx).Add "�ڍ�","�鏑�Ƃ������ʂȐE�������łȂ��A���i���邱�Ƃň�ʎ����E�̒m���ƋZ�\�������Ă���ؖ��Ƃ��Ē��ڂ���Ă��܂��B���������E��񏈗��E�ڋ��̃G�L�X�p�[�g�Ƃ��āA������g�D�Ŋ���ł������L���邱�Ƃ��ł��܂��B"
	aData(idx).Add "��p","�󌱗���" & vbCrLf & "1���@6,000�~" & vbCrLf & "��1���@4,800�~" & vbCrLf & "2���@3,700�~" & vbCrlf & "3���@2,500�~"
	aData(idx).Add "���i��","1���@25�D3��" & vbCrlf & "��1���@29�D8��" & vbCrlf & "2���@43�D5��" & vbCrlf & "3���@64�D0��" & vbCrlf & "�i08�N11���j"
	aData(idx).Add "�c�̖�","���c�@�l�����Z�\���苦��"
	aData(idx).Add "����@��","�Y�Ɣ\����w����������"
	aData(idx).Add "�u�����e","���ʐM����" & vbCrLf & "�i÷�āE�Y��L��j"
	aData(idx).Add "���i","1���@26,250�~" & vbCrLf & "��1���@26,250�~"
	aData(idx).Add "���T","��u��20��OFF" & vbCrLf & "���x�����@�F������������ɂ���u�ƂȂ�܂��B"
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�ΏۋƎ�","�S�Ǝ�Ή��̎��i"
	aData(idx).Add "���i2","1���@26,250�~<br>��1���@26,250�~"
	aData(idx).Add "�T�v2","�w�����Ȋw�ȔF��̌��I���i�x�Ƃ��āA�鏑�݂̂Ȃ炸�����E�S�ʂ̍����X�L�����ؖ��ł��鎑�i"
	aData(idx).Add "���T2","��u��<br>20��OFF"
	'</�鏑����F�莎��1��>

	'<�鏑����F�莎��2��>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0402"
	aData(idx).Add "�D��",""
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","�鏑����F�莎���i2���j"
	aData(idx).Add "����TITLE","�鏑����F�莎���i2���j"
	aData(idx).Add "�T�v","�y�S�Ǝ�Ή��̎��i�z" & vbCrLf & vbCrLf & "�w�����Ȋw�ȔF��̌��I���i�x�Ƃ��āA�鏑�݂̂Ȃ炸�����E�S�ʂ̍����X�L�����ؖ��ł��鎑�i"
	aData(idx).Add "�ڍ�","�鏑�Ƃ������ʂȐE�������łȂ��A���i���邱�Ƃň�ʎ����E�̒m���ƋZ�\�������Ă���ؖ��Ƃ��Ē��ڂ���Ă��܂��B���������E��񏈗��E�ڋ��̃G�L�X�p�[�g�Ƃ��āA������g�D�Ŋ���ł������L���邱�Ƃ��ł��܂��B"
	aData(idx).Add "��p","�󌱗���" & vbCrLf & "1���@6,000�~" & vbCrLf & "��1���@4,800�~" & vbCrLf & "2���@3,700�~" & vbCrlf & "3���@2,500�~"
	aData(idx).Add "���i��","1���@25�D3��" & vbCrlf & "��1���@29�D8��" & vbCrlf & "2���@43�D5��" & vbCrlf & "3���@64�D0��" & vbCrlf & "�i08�N11���j"
	aData(idx).Add "�c�̖�","���c�@�l�����Z�\���苦��"
	aData(idx).Add "����@��","�Y�Ɣ\����w����������"
	aData(idx).Add "�u�����e","���ʐM����" & vbCrLf & "�i÷�āE�Y��L��j"
	aData(idx).Add "���i","2���@21,000�~"
	aData(idx).Add "���T","��u��20��OFF" & vbCrLf & "���x�����@�F������������ɂ���u�ƂȂ�܂��B"
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�ΏۋƎ�","�S�Ǝ�Ή��̎��i"
	aData(idx).Add "���i2","2���@21,000�~"
	aData(idx).Add "�T�v2","�����Ȋw�ȔF��̌��I���i�x�Ƃ��āA�鏑�݂̂Ȃ炸�����E�S�ʂ̍����X�L�����ؖ��ł��鎑�i"
	aData(idx).Add "���T2","��u��<br>20��OFF"

	'</�鏑����F�莎��2��>

	'<�鏑����F�莎��3��>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0403"
	aData(idx).Add "�D��",""
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","�鏑����F�莎���i3���j"
	aData(idx).Add "����TITLE","�鏑����F�莎���i3���j"
	aData(idx).Add "�T�v","�y�S�Ǝ�Ή��̎��i�z" & vbCrLf & vbCrLf & "�w�����Ȋw�ȔF��̌��I���i�x�Ƃ��āA�鏑�݂̂Ȃ炸�����E�S�ʂ̍����X�L�����ؖ��ł��鎑�i"
	aData(idx).Add "�ڍ�","�鏑�Ƃ������ʂȐE�������łȂ��A���i���邱�Ƃň�ʎ����E�̒m���ƋZ�\�������Ă���ؖ��Ƃ��Ē��ڂ���Ă��܂��B���������E��񏈗��E�ڋ��̃G�L�X�p�[�g�Ƃ��āA������g�D�Ŋ���ł������L���邱�Ƃ��ł��܂��B"
	aData(idx).Add "��p","�󌱗���" & vbCrLf & "1���@6,000�~" & vbCrLf & "��1���@4,800�~" & vbCrLf & "2���@3,700�~" & vbCrlf & "3���@2,500�~"
	aData(idx).Add "���i��","1���@25�D3��" & vbCrlf & "��1���@29�D8��" & vbCrlf & "2���@43�D5��" & vbCrlf & "3���@64�D0��" & vbCrlf & "�i08�N11���j"
	aData(idx).Add "�c�̖�","���c�@�l�����Z�\���苦��"
	aData(idx).Add "����@��","�Y�Ɣ\����w����������"
	aData(idx).Add "�u�����e","���ʐM����" & vbCrLf & "�i÷�āE�Y��L��j"
	aData(idx).Add "���i","3���@25,200�~"
	aData(idx).Add "���T","��u��20��OFF" & vbCrLf & "���x�����@�F������������ɂ���u�ƂȂ�܂��B"
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�ΏۋƎ�","�S�Ǝ�Ή��̎��i"
	aData(idx).Add "���i2","3���@25,200�~"
	aData(idx).Add "�T�v2","�w�����Ȋw�ȔF��̌��I���i�x�Ƃ��āA�鏑�݂̂Ȃ炸�����E�S�ʂ̍����X�L�����ؖ��ł��鎑�i"
	aData(idx).Add "���T2","��u��<br>20��OFF"

	'</�鏑����F�莎��3��>

	'<�����ٽ�E�ȼ���Č���T��>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0501"
	aData(idx).Add "�D��",""
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","�����^���w���X�E�}�l�W�����g����i�T��j"
	aData(idx).Add "����TITLE","�����ٽ�E�ȼ���Č���i�T��j"
	aData(idx).Add "�T�v","�y�����E�l���E�J���E�����̕��ɂ��E�߂̎��i�z" & vbCrLf & vbCrLf & "��Ɠ��ł̓K�؂ȃ����^���w���X�΍���u���鑍���E�l���J���Ǘ��Ɍg�����ɂ����߂̌���"
	aData(idx).Add "�ڍ�","�S�̕a�������J���҂̑������Љ��艻���Ă���A�����l�́u�S�̌��N�Ǘ��v�Ɋ֐S�����܂��Ă��܂��B�����J���Ȃł͐E��ɂ�����K�؂��L���ȃ����^���w���X�΍�̎��{�𐄐i���Ă��܂��B�{���莎���́A�����J���Ȃ́u�J���҂̐S�̌��N�̕ێ����i�̂��߂̎w�j�v�Ɋ�Â��č\�z����Ă��܂��B"
	aData(idx).Add "��p","�T��@10,500�~" & vbCrLf & "�U��@�@6,300�~" & vbCrlf & "�V��@�@4,200�~"
	aData(idx).Add "���i��","08�N�x" & vbCrLf & "�T��@11�D1��" & vbCrLf & "�U��@70�D7��" & vbCrLf & "�V��@87�D4��"
	aData(idx).Add "�c�̖�","��㏤�H��c��"
	aData(idx).Add "����@��","�Y�Ɣ\����w����������"
	aData(idx).Add "�u�����e","���ʐM����" & vbCrLf & "�i÷�āE�Y��L��j"
	aData(idx).Add "���i","��u��" & vbCrLf & "�T��@24,1500�~"
	aData(idx).Add "���T","���ʎ�u��" & vbCrLf & "�T��21,000�~" & vbCrLf & "���x�����@�F������������ɂ���u�ƂȂ�܂��B"
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�ΏۋƎ�","�����E�l���E�J���E�����̕��ɂ��E��"
	aData(idx).Add "���i2","�T��@24,1500�~"
	aData(idx).Add "�T�v2","��Ɠ��ł̓K�؂ȃ����^���w���X�΍���u���鑍���E�l���J���Ǘ��Ɍg�����ɂ����߂̌���"
	aData(idx).Add "���T2","���ʎ�u��<br>�T��21,000�~"

	'</�����ٽ�E�ȼ���Č���T��>

	'<�����ٽ�E�ȼ���Č���U��>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0502"
	aData(idx).Add "�D��",""
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","�����^���w���X�E�}�l�W�����g����i�U��j"
	aData(idx).Add "����TITLE","�����ٽ�E�ȼ���Č���i�U��j"
	aData(idx).Add "�T�v","�y�����E�l���E�J���E�����̕��ɂ��E�߂̎��i�z" & vbCrLf & vbCrLf & "��Ɠ��ł̓K�؂ȃ����^���w���X�΍���u���鑍���E�l���J���Ǘ��Ɍg�����ɂ����߂̌���"
	aData(idx).Add "�ڍ�","�S�̕a�������J���҂̑������Љ��艻���Ă���A�����l�́u�S�̌��N�Ǘ��v�Ɋ֐S�����܂��Ă��܂��B�����J���Ȃł͐E��ɂ�����K�؂��L���ȃ����^���w���X�΍�̎��{�𐄐i���Ă��܂��B�{���莎���́A�����J���Ȃ́u�J���҂̐S�̌��N�̕ێ����i�̂��߂̎w�j�v�Ɋ�Â��č\�z����Ă��܂��B"
	aData(idx).Add "��p","�T��@10,500�~" & vbCrLf & "�U��@�@6,300�~" & vbCrlf & "�V��@�@4,200�~"
	aData(idx).Add "���i��","08�N�x" & vbCrLf & "�T��@11�D1��" & vbCrLf & "�U��@70�D7��" & vbCrLf & "�V��@87�D4��"
	aData(idx).Add "�c�̖�","��㏤�H��c��"
	aData(idx).Add "����@��","�Y�Ɣ\����w����������"
	aData(idx).Add "�u�����e","���ʐM����" & vbCrLf & "�i÷�āE�Y��L��j"
	aData(idx).Add "���i","��u��" & vbCrLf & "�U��@12,600�~"
	aData(idx).Add "���T","���ʎ�u��" & vbCrLf & "�U��@9,450�~" & vbCrLf & "���x�����@�F������������ɂ���u�ƂȂ�܂��B"
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�ΏۋƎ�","�����E�l���E�J���E�����̕��ɂ��E��"
	aData(idx).Add "���i2","�U��@12,600�~"
	aData(idx).Add "�T�v2","��Ɠ��ł̓K�؂ȃ����^���w���X�΍���u���鑍���E�l���J���Ǘ��Ɍg�����ɂ����߂̌���"
	aData(idx).Add "���T2","���ʎ�u��<br>�U��@9,450�~"

	'</�����ٽ�E�ȼ���Č���U��>

	'<�����ٽ�E�ȼ���Č���V��>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0503"
	aData(idx).Add "�D��",""
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","�����^���w���X" & vbCrLf & "�}�l�W�����g����i�V��j"
	aData(idx).Add "����TITLE","�����ٽ�E�ȼ���Č���i�V��j"
	aData(idx).Add "�T�v","�y�����E�l���E�J���E�����̕��ɂ��E�߂̎��i�z" & vbCrLf & vbCrLf & "��Ɠ��ł̓K�؂ȃ����^���w���X�΍���u���鑍���E�l���J���Ǘ��Ɍg�����ɂ����߂̌���"
	aData(idx).Add "�ڍ�","�S�̕a�������J���҂̑������Љ��艻���Ă���A�����l�́u�S�̌��N�Ǘ��v�Ɋ֐S�����܂��Ă��܂��B�����J���Ȃł͐E��ɂ�����K�؂��L���ȃ����^���w���X�΍�̎��{�𐄐i���Ă��܂��B�{���莎���́A�����J���Ȃ́u�J���҂̐S�̌��N�̕ێ����i�̂��߂̎w�j�v�Ɋ�Â��č\�z����Ă��܂��B"
	aData(idx).Add "��p","�T��@10,500�~" & vbCrLf & "�U��@�@6,300�~" & vbCrlf & "�V��@�@4,200�~"
	aData(idx).Add "���i��","08�N�x" & vbCrLf & "�T��@11�D1��" & vbCrLf & "�U��@70�D7��" & vbCrLf & "�V��@87�D4��"
	aData(idx).Add "�c�̖�","��㏤�H��c��"
	aData(idx).Add "����@��","�Y�Ɣ\����w����������"
	aData(idx).Add "�u�����e","���ʐM����" & vbCrLf & "�i÷�āE�Y��L��j"
	aData(idx).Add "���i","��u��" & vbCrLf & "�V��@11,970�~"
	aData(idx).Add "���T","���ʎ�u��" & vbCrLf & "�V��@8,820�~" & vbCrLf & "���x�����@�F������������ɂ���u�ƂȂ�܂��B"
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�ΏۋƎ�","�����E�l���E�J���E�����̕��ɂ��E��"
	aData(idx).Add "���i2","�V��@11,970�~"
	aData(idx).Add "�T�v2","��Ɠ��ł̓K�؂ȃ����^���w���X�΍���u���鑍���E�l���J���Ǘ��Ɍg�����ɂ����߂̌���"
	aData(idx).Add "���T2","���ʎ�u��<br>�V��@8,820�~"

	'</�����ٽ�E�ȼ���Č���V��>

	'<������L����2��>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0602"
	aData(idx).Add "�D��",""
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","������L���莎���i2���j"
	aData(idx).Add "����TITLE","������L���莎���i2���j"
	aData(idx).Add "�T�v","�y�o���͂������̂��Ƒ���Ɩ��Ŋ������鎑�i�z" & vbCrLf & vbCrLf & "��Ќo�c�̐����𗝉����邤���ŕK�{�̃X�L��"
	aData(idx).Add "�ڍ�","��ƋK�͂�Ǝ�E�ƑԂ��킸�A�o�c�������L�^�E�v�Z�E�������o�c���тƍ�����Ԃ𖾂炩�ɂ��邾���łȂ��A�����̌o�c��Ԃ����c���ł���Z�\�ł��邱�Ƃ���o���S���҂����łȂ����L���r�W�l�X�X�L���Ƃ��Ė𗧂B�����i�Ƃ̑g�ݍ��킹�ɂ��L�����A�A�b�v��ڎw�����Ƃ��\�B"
	aData(idx).Add "��p","�󌱗��́A2���@4,500�~" & vbCrLf & "3���@2,500�~"
	aData(idx).Add "���i��","2��" & vbCrlf & "�g20.11���{�@29.6��" & vbCrLf & "�g20.06���{�@31.3��" & vbCrLf & vbCrLf & "3��" & vbCrLf & "�g20.11���{�@40.2��" & vbCrlf & "�g20.06���{�@29.5��"
	aData(idx).Add "�c�̖�","���{���H��c��"
	aData(idx).Add "����@��","�Y�Ɣ\����w����������"
	aData(idx).Add "�u�����e","���ʐM����" & vbCrlf & "�i÷�āE�Y��L��j"
	aData(idx).Add "���i","��u��" & vbCrLf & "2��22,050�~"
	aData(idx).Add "���T","��25��OFF" & vbCrLf & "���ʎ�u��" & vbCrLf & "2��16,800�~" & vbCrLf & "�������̎��{�����ɍ��킹�āA�󌱐\�����@�Ȃǂ̏������͂��I" & vbCrLf & "���x�����@�F������������ɂ���u�ƂȂ�܂��B"
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�ΏۋƎ�","�o���͂������̂��Ƒ���Ɩ��Ŋ�������"
	aData(idx).Add "���i2","2��22,050�~"
	aData(idx).Add "�T�v2","��Ќo�c�̐����𗝉����邤���ŕK�{�̃X�L��"
	aData(idx).Add "���T2","���ʎ�u��<br>2��16,800�~"

	'</������L����2��>

	'<������L����3��>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0603"
	aData(idx).Add "�D��",""
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","������L���莎���i3���j"
	aData(idx).Add "����TITLE","������L���莎���i3���j"
	aData(idx).Add "�T�v","�y�o���͂������̂��Ƒ���Ɩ��Ŋ������鎑�i�z" & vbCrLf & vbCrLf & "��Ќo�c�̐����𗝉����邤���ŕK�{�̃X�L��"
	aData(idx).Add "�ڍ�","��ƋK�͂�Ǝ�E�ƑԂ��킸�A�o�c�������L�^�E�v�Z�E�������o�c���тƍ�����Ԃ𖾂炩�ɂ��邾���łȂ��A�����̌o�c��Ԃ����c���ł���Z�\�ł��邱�Ƃ���o���S���҂����łȂ����L���r�W�l�X�X�L���Ƃ��Ė𗧂B�����i�Ƃ̑g�ݍ��킹�ɂ��L�����A�A�b�v��ڎw�����Ƃ��\�B"
	aData(idx).Add "��p","�󌱗��́A2���@4,500�~" & vbCrLf & "3���@2,500�~"
	aData(idx).Add "���i��","2��" & vbCrlf & "�g20.11���{�@29.6��" & vbCrLf & "�g20.06���{�@31.3��" & vbCrLf & vbCrLf & "3��" & vbCrLf & "�g20.11���{�@40.2��" & vbCrlf & "�g20.06���{�@29.5��"
	aData(idx).Add "�c�̖�","���{���H��c��"
	aData(idx).Add "����@��","�Y�Ɣ\����w����������"
	aData(idx).Add "�u�����e","���ʐM����" & vbCrlf & "�i÷�āE�Y��L��j"
	aData(idx).Add "���i","��u��" & vbCrLf & "3��19,950�~�@�i���<del>19,950�~</del>�j"
	aData(idx).Add "���T","��25��OFF" & vbCrLf & "���ʎ�u��" & vbCrLf & "3��14,700�~" & vbCrLf & "�������̎��{�����ɍ��킹�āA�󌱐\�����@�Ȃǂ̏������͂��I" & vbCrLf & "���x�����@�F������������ɂ���u�ƂȂ�܂��B"
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�ΏۋƎ�","�o���͂������̂��Ƒ���Ɩ��Ŋ�������"
	aData(idx).Add "���i2","3��19,950�~"
	aData(idx).Add "�T�v2","��Ќo�c�̐����𗝉����邤���ŕK�{�̃X�L��"
	aData(idx).Add "���T2","���ʎ�u��<br>3��14,700�~"

	'</������L����3��>

	'<�r�W�l�X�E�L�����A���莎��2��>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0102"
	aData(idx).Add "�D��","1"
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","�L�����A���莎���i2���j"
	aData(idx).Add "����TITLE","�L�����A���莎���i2���j"
	aData(idx).Add "�T�v","�y����Ǝ�Ή��̎��i�z" & vbCrLf & vbCrLf & "�����n�E�����L���ԗ������B��̌��I���i����" &vbCrLf & "�E���ʂɕK�v�ȃX�L���̌n���\�z���A��Ǝ����ɑ��������I�m���E�\�͂��q�ϓI�ɕ]���������i" & vbCrLf & vbCrLf & "���x���̃C���[�W" & vbCrLf & "�E���Ɋ֘A���镝�L�������I�Ȑ��m������ɁA�O���[�v��`�[���̒��S�����o�[�Ƃ��āA�n�ӍH�v���Â炵�A����I�Ȕ��f�E���P�E��Ă��s���Ȃ���Ɩ��𐋍s���邱�Ƃ��ł���B�i�Ⴆ�΁A�ے��A�}�l�[�W���[����ڎw���l�A���̓V�j�A�E�X�^�b�t�j" & vbCrLf & vbCrLf & "������͕��L���I" & vbCrLf & "�l���l�ފJ���E�J���Ǘ��E�o���E�����Ǘ��E�o�c���V�X�e���A�c�ƃ}�[�P�e�B���O" & vbCrLf & "�ȂǁB"
	aData(idx).Add "�ڍ�","�r�W�l�X�L�����A����Ƃ�&nbsp;�|&nbsp;�l�ޗ͂����߁A��Ɨ͂����߂�r�W�l�X�L�����A����" & vbCrLf & vbCrLf & "1�D���̒�߂��ɏ������������E�r�W�l�X�L�����A����" & vbCrLf & "��(�����J����)���A�r�W�l�X�p�[�\���̐E��(�Z�N�V����)�ʂɕK�v�ȃX�L���̌n(�K�C�h���C��)���\�z���A���̃X�L���̌n(�K�C�h���C��)����Ɍ��I���i�Ƃ��Ă̌��莎�������{�B" & vbCrLf & "" & vbCrLf & "2�D�E���𕝍L���J�o�[�����B��̌��莎���E�r�W�l�X�L�����A����" & vbCrLf & "��(�����J����)�Ɗw���o���҃O���[�v���A�r�W�l�X�p�[�\���̐E��(�Z�N�V����)�ʂɕK�v�ȃX�L���v�f�𒊏o���A�����X�L���v�f�̊֘A����d�v���Ȃǂ���̌n���\�z���Ă���̂ŐE���S�̂̋Ɩ����s�̂��߂ɕK�v�ȃX�L�����S�ԗ�����Ă��錟�莎���B" & vbCrLf & "" & vbCrLf & "3�D�����\�͂̕]�����d�����������E�r�W�l�X�L�����A����" & vbCrLf & "�e�E���ɕK�v�Ȓm���C�����͂��߁A�����́A1���A2���y��3���̃��x���ɑ̌n������A�����ɑ��������I�m����\�͂��q�ϓI�ɕ]�����A��Ƃł͎Ј��̎����\�͂̋q�ϓI�ȕ]����l�ފJ�����ɁA�l�ɂƂ��āA�L�����A�A�b�v�Ȃǂɕ��L�����p�ł��鎎���B" & vbCrLf & "" & vbCrLf & "4�D�w�K���₷���̐��E�r�W�l�X�L�����A����" & vbCrLf & "�����J���Ȃ��玎�����K�C�h���C���ɏ��������W���e�L�X�g���������{�@�ւł��钆���E�Ɣ\�͊J��������s���A�l�̎��w�K���ނ�ʐM�A�ʊw�p�u���̋��ށA��Ɠ����C�ł̋��ނƂ��Ċ��p�ł��A�����E�Ɣ\�͊J������F�肵�������Ή��u�������p���Ċw�K���邱�Ƃ��ł���"
	aData(idx).Add "��p","����������F" & vbCrLf & "�l���l�ފJ���i1.2.3���j�A�J���Ǘ��i1.2.3���j�A�o���i1.2.3���j�A�����Ǘ��i1.2.3���j�A�o�c���V�X�e���i1.2.3���j�A�c�ƃ}�[�P�e�B���O�i1.2.3���j" & vbCrLf & vbCrLf & "����p�F" & vbCrlf & "1����7,850�~�A2����5,250�~�A3����4,200�~"
	aData(idx).Add "���i��","�l���l�ފJ���i1.2.3���j���e20��.35��.43��" & vbCrLf & "�J���Ǘ��i2.3���j���e24��.45��" & vbCrLf & "�o��1��. �o��2���i������v�j. �o��3���i�����v�Z�j���e20��.39��.53��" & vbCrLf & "�����Ǘ��i2.3���j���e18��.68��" & vbCrLf & "�o�c���V�X�e���i1���j��18��" & vbCrLf & "�c�Ɓi1.2.3���j��20��.44��.41��" & vbCrLf & "���̑��ꕔ����J" & vbCrLf & "�i08�N�O���j"
	aData(idx).Add "�c�̖�","�����E�Ɣ\�͊J������"
	aData(idx).Add "����@��","�y�ʐM�u���z" & vbCrLf & "�y�ʊw�X�N�[���z" & vbCrLf & "�m�l�q�r�W�l�X�L�����A�w�@"
	aData(idx).Add "�u�����e","���ʐM����i�F��u���j" & vbCrLf & "�����i�R�[�X" & vbCrLf & "�����Ȍ[���R�[�X"
	aData(idx).Add "���i","����u��" & vbCrLf & "����ʁE�R�[�X�ʂɂ��قȂ�i22,000�~�`�j" & vbCrLf & vbCrLf & "���c�ƁE�}�[�P�e�B���O����" & vbCrLf & "�@�E�}�[�P�e�B���O2���@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & "�@�E�c��2���@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & vbCrLf & "���l���E�l�ފJ���E�J���Ǘ�����" & vbCrLf & "�@�E�l���E�l�ފJ��2���@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & "�@�E�J���Ǘ�2���@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & vbCrLf & "����Ɩ@���E��������" & vbCrLf & "�@�E��Ɩ@��2��(����@��)�@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & "�@�E��Ɩ@��2��(�g�D�@��)�@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & "�@�E����2���@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & vbCrLf & "���o���E�����Ǘ�����" & vbCrLf & "�@�E�o��2��(������v)�@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & "�@�E�����Ǘ�2��(�����Ǘ��E�Ǘ���v)�@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & vbCrLf & "���o�c�헪����" & vbCrLf & "�@�E�o�c�헪2���@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & vbCrLf & "���o�c���V�X�e������" & vbCrLf & "�@�E�o�c���V�X�e��2��(��񉻊��)�@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & "�@�E�o�c���V�X�e��2��(��񉻊��p)�@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & vbCrLf & "�����W�X�e�B�N�X����" & vbCrLf & "�@�E���W�X�e�B�N�X�Ǘ�2���@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & "�@�E���W�X�e�B�N�X�E�I�y���[�V����2���@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & vbCrLf & "�����Y�Ǘ�����" & vbCrLf & "�@�E���Y�Ǘ��v�����j���O2��(���Y�V�X�e���E���Y�Ǘ�)�@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & "�@�E���Y�Ǘ��v�����j���O2��(���i���E�݌v�Ǘ�)�@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & "�@�E���Y�Ǘ��I�y���[�V����2��(�w���E�����E�݌ɊǗ�)�@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~" & vbCrLf & "�@�E���Y�Ǘ��I�y���[�V����2��(��ƁE�H���E�ݔ��Ǘ�)�@���i�R�[�X 38,000�~�@���Ȍ[���R�[�X 22,000�~"
	aData(idx).Add "���T",""
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�N�[�|�����ӓ_",""
	aData(idx).Add "�ΏۋƎ�","����Ǝ�Ή�"
	aData(idx).Add "���i2","�R�[�X�ʂɂ��قȂ�<br>�i22,000�~�`�j"
	aData(idx).Add "�T�v2","�����n�E�����L���ԗ������B��̌��I���i����"
	aData(idx).Add "���T2",""

	'</�r�W�l�X�E�L�����A���莎��2��>

	'<�r�W�l�X�E�L�����A���莎��3��>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0103"
	aData(idx).Add "�D��","1"
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","�L�����A���莎���i3���j"
	aData(idx).Add "����TITLE","�L�����A���莎���i3���j"
	aData(idx).Add "�T�v","�y����Ǝ�Ή��̎��i�z" & vbCrLf & vbCrLf & "�����n�E�����L���ԗ������B��̌��I���i����" &vbCrLf & "�E���ʂɕK�v�ȃX�L���̌n���\�z���A��Ǝ����ɑ��������I�m���E�\�͂��q�ϓI�ɕ]���������i" & vbCrLf & vbCrLf & "���x���̃C���[�W" & vbCrLf & "�E���S�ʂɊւ��镝�L�����m������ɁA�S���҂Ƃ��ď�i�̎w���E�����𓥂܂��A������ӎ���������^�I�Ɩ����m���ɐ��s���邱�Ƃ��ł���B�i�Ⴆ�΁A�W���A���[�_�[����ڎw���l�A���͒S���Ɩ���I�m�ɐ��s�ł��邱�Ƃ�ڎw���l�j" & vbCrLf & vbCrLf & "������͕��L���I" & vbCrLf & "�l���l�ފJ���E�J���Ǘ��E�o���E�����Ǘ��E�o�c���V�X�e���A�c�ƃ}�[�P�e�B���O" & vbCrLf & "�ȂǁB"
	aData(idx).Add "�ڍ�","�r�W�l�X�L�����A����Ƃ�&nbsp;�|&nbsp;�l�ޗ͂����߁A��Ɨ͂����߂�r�W�l�X�L�����A����" & vbCrLf & vbCrLf & "1�D���̒�߂��ɏ������������E�r�W�l�X�L�����A����" & vbCrLf & "��(�����J����)���A�r�W�l�X�p�[�\���̐E��(�Z�N�V����)�ʂɕK�v�ȃX�L���̌n(�K�C�h���C��)���\�z���A���̃X�L���̌n(�K�C�h���C��)����Ɍ��I���i�Ƃ��Ă̌��莎�������{�B" & vbCrLf & "" & vbCrLf & "2�D�E���𕝍L���J�o�[�����B��̌��莎���E�r�W�l�X�L�����A����" & vbCrLf & "��(�����J����)�Ɗw���o���҃O���[�v���A�r�W�l�X�p�[�\���̐E��(�Z�N�V����)�ʂɕK�v�ȃX�L���v�f�𒊏o���A�����X�L���v�f�̊֘A����d�v���Ȃǂ���̌n���\�z���Ă���̂ŐE���S�̂̋Ɩ����s�̂��߂ɕK�v�ȃX�L�����S�ԗ�����Ă��錟�莎���B" & vbCrLf & "" & vbCrLf & "3�D�����\�͂̕]�����d�����������E�r�W�l�X�L�����A����" & vbCrLf & "�e�E���ɕK�v�Ȓm���C�����͂��߁A�����́A1���A2���y��3���̃��x���ɑ̌n������A�����ɑ��������I�m����\�͂��q�ϓI�ɕ]�����A��Ƃł͎Ј��̎����\�͂̋q�ϓI�ȕ]����l�ފJ�����ɁA�l�ɂƂ��āA�L�����A�A�b�v�Ȃǂɕ��L�����p�ł��鎎���B" & vbCrLf & "" & vbCrLf & "4�D�w�K���₷���̐��E�r�W�l�X�L�����A����" & vbCrLf & "�����J���Ȃ��玎�����K�C�h���C���ɏ��������W���e�L�X�g���������{�@�ւł��钆���E�Ɣ\�͊J��������s���A�l�̎��w�K���ނ�ʐM�A�ʊw�p�u���̋��ށA��Ɠ����C�ł̋��ނƂ��Ċ��p�ł��A�����E�Ɣ\�͊J������F�肵�������Ή��u�������p���Ċw�K���邱�Ƃ��ł���"
	aData(idx).Add "��p","����������F" & vbCrLf & "�l���l�ފJ���i1.2.3���j�A�J���Ǘ��i1.2.3���j�A�o���i1.2.3���j�A�����Ǘ��i1.2.3���j�A�o�c���V�X�e���i1.2.3���j�A�c�ƃ}�[�P�e�B���O�i1.2.3���j" & vbCrLf & vbCrLf & "����p�F" & vbCrlf & "1����7,850�~�A2����5,250�~�A3����4,200�~"
	aData(idx).Add "���i��","�l���l�ފJ���i1.2.3���j���e20��.35��.43��" & vbCrLf & "�J���Ǘ��i2.3���j���e24��.45��" & vbCrLf & "�o��1��. �o��2���i������v�j. �o��3���i�����v�Z�j���e20��.39��.53��" & vbCrLf & "�����Ǘ��i2.3���j���e18��.68��" & vbCrLf & "�o�c���V�X�e���i1���j��18��" & vbCrLf & "�c�Ɓi1.2.3���j��20��.44��.41��" & vbCrLf & "���̑��ꕔ����J" & vbCrLf & "�i08�N�O���j"
	aData(idx).Add "�c�̖�","�����E�Ɣ\�͊J������"
	aData(idx).Add "����@��","�y�ʐM�u���z" & vbCrLf & "�y�ʊw�X�N�[���z" & vbCrLf & "�m�l�q�r�W�l�X�L�����A�w�@"
	aData(idx).Add "�u�����e","���ʐM����i�F��u���j" & vbCrLf & "�����i�R�[�X" & vbCrLf & "�����Ȍ[���R�[�X"
	aData(idx).Add "���i","����u��" & vbCrLf & "����ʁE�R�[�X�ʂɂ��قȂ�i19,500�~�`�j" & vbCrLf & vbCrLf & "���c�ƁE�}�[�P�e�B���O����" & vbCrLf & "�@�E�}�[�P�e�B���O3���@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~" & vbCrLf & "�@�E�c��3���@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~" & vbCrLf & vbCrLf & "���l���E�l�ފJ���E�J���Ǘ�����" & vbCrLf & "�@�E�l���E�l�ފJ��3���@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~" & vbCrLf & "�@�E�J���Ǘ�3���@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~" & vbCrLf & vbCrLf & "����Ɩ@���E��������" & vbCrLf & "�@�E��Ɩ@��3���@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~" & vbCrLf & "�@�E����3���@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~" & vbCrLf & vbCrLf & "���o���E�����Ǘ�����" & vbCrLf & "�@�E�o��3��(��L�E�������\)�@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~" & vbCrLf & "�@�E�o��3��(�����v�Z)�@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~" & vbCrLf & "�@�E�����Ǘ�3���@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~" & vbCrLf & vbCrLf & "���o�c�헪����" & vbCrLf & "�@�E�o�c�헪3���@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~" & vbCrLf & vbCrLf & "���o�c���V�X�e������" & vbCrLf & "�@�E�o�c���V�X�e��3���@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~" & vbCrLf & vbCrLf & "�����W�X�e�B�N�X����" & vbCrLf & "�@�E���W�X�e�B�N�X�Ǘ�3���@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~" & vbCrLf & "�@�E���W�X�e�B�N�X�E�I�y���[�V����3���@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~" & vbCrLf & vbCrLf & "�����Y�Ǘ�����" & vbCrLf & "�@�E���Y�Ǘ��v�����j���O3���@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~" & vbCrLf & "�@�E���Y�Ǘ��I�y���[�V����3���@���i�R�[�X 33,000�~�@���Ȍ[���R�[�X 19,500�~"
	aData(idx).Add "���T",""
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�N�[�|�����ӓ_",""
	aData(idx).Add "�ΏۋƎ�","����Ǝ�Ή�"
	aData(idx).Add "���i2","�R�[�X�ʂɂ��قȂ�<br>�i19,500�~�`�j"
	aData(idx).Add "�T�v2","�����n�E�����L���ԗ������B��̌��I���i����"
	aData(idx).Add "���T2",""

	'</�r�W�l�X�E�L�����A���莎��3��>

	'<�q���Ǘ���1��>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0201"
	aData(idx).Add "�D��",""
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","�q���Ǘ��ҁi1��j"
	aData(idx).Add "����TITLE","�q���Ǘ��ҁi1��j"
	aData(idx).Add "�T�v","�y�S�Ǝ�Ή��̎��i�z" & vbCrLf & "�y�����E�l���E�����̕��ɂ��E�߁z" & vbCrLf & vbCrLf & "�w���Ǝ��i�x(�����J���ȔF��j�Ƃ��āA�S�Ǝ�ɑI�C���`���Â����Ă��鎑�i"
	aData(idx).Add "�ڍ�","�펞50�l�ȏ�̘J���҂��g�p���鎖�Ə�ł́A�q���Ǘ��ҖƋ���L����҂̂�������J���Ґ��ɉ����Ĉ�萔�ȏ�̉q���Ǘ��҂�I�C���A���S�q���Ɩ��̂����A�q���Ɋւ��Z�p�I�Ȏ������Ǘ������邱�Ƃ��K�v�ɂȂ�܂��B" & vbCrLf & "��1��͑S�Ă̎��Ə��ɂ����ĊǗ��҂ɂȂ�܂��B��2��́A�L�Q�Ɩ��Ɗ֘A�̔������ʐM�Ƃ���Z�ƂȂǂ̈��̋Ǝ�̎��Ə�ɂ����Ă̂݊Ǘ��҂ɂȂ�܂��B�����ȐE���́A�J���҂̌��N��Q��h�~���邽�߂̍�Ɗ��Ǘ��E��ƊǗ��E���N�Ǘ��E�J���q������̎��{�E���N�ێ����i�[�u�Ȃǂł��B"
	aData(idx).Add "��p","�󌱗���8,300�~�B" & vbCrLf & "���i��A�o�^�萔���󎆑�1,500�~����������B"
	aData(idx).Add "���i��","��1��F54.7��" & vbCrLf & "��2��F65.6��" & vbCrLf & "�i07�N�x�j"
	aData(idx).Add "�c�̖�","�i���j���S�q���Z�p��������"
	aData(idx).Add "����@��","�Y�Ɣ\����w����������"
	aData(idx).Add "�u�����e","���ʐM����" & vbCrLf & "�i÷�āE�Y��L��j" & vbCrLf & "���\�z���W�E�p�����W�t"
	aData(idx).Add "���i","����u��" & vbCrLf & "1��@25,200�~" & vbCrLf & "���{�����ɑΉ��������H�I���|�[�g���Ŋm���ɍ��i�x���I" & vbCrLf & "���x�����@�F������������ɂ���u�ƂȂ�܂��B"
	aData(idx).Add "���T",""
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�ΏۋƎ�","�S�Ǝ�Ή�"
	aData(idx).Add "���i2","1��@25,200�~"
	aData(idx).Add "�T�v2","�w���Ǝ��i�x(�����J���ȔF��j�Ƃ��āA�S�Ǝ�ɑI�C���`���Â����Ă��鎑�i"
	aData(idx).Add "���T2",""

	'</�q���Ǘ���1��>

	'<�q���Ǘ���2��>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0202"
	aData(idx).Add "�D��",""
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","�q���Ǘ��ҁi2��j"
	aData(idx).Add "����TITLE","�q���Ǘ��ҁi2��j"
	aData(idx).Add "�T�v","�y�S�Ǝ�Ή��̎��i�z" & vbCrLf & "�y�����E�l���E�����̕��ɂ��E�߁z" & vbCrLf & vbCrLf & "�w���Ǝ��i�x(�����J���ȔF��j�Ƃ��āA�S�Ǝ�ɑI�C���`���t�����Ă��鎑�i"
	aData(idx).Add "�ڍ�","�펞50�l�ȏ�̘J���҂��g�p���鎖�Ə�ł́A�q���Ǘ��ҖƋ���L����҂̂�������J���Ґ��ɉ����Ĉ�萔�ȏ�̉q���Ǘ��҂�I�C���A���S�q���Ɩ��̂����A�q���Ɋւ��Z�p�I�Ȏ������Ǘ������邱�Ƃ��K�v�ɂȂ�܂��B" & vbCrLf & "��1��͑S�Ă̎��Ə��ɂ����ĊǗ��҂ɂȂ�܂��B��2��́A�L�Q�Ɩ��Ɗ֘A�̔������ʐM�Ƃ���Z�ƂȂǂ̈��̋Ǝ�̎��Ə�ɂ����Ă̂݊Ǘ��҂ɂȂ�܂��B�����ȐE���́A�J���҂̌��N��Q��h�~���邽�߂̍�Ɗ��Ǘ��E��ƊǗ��E���N�Ǘ��E�J���q������̎��{�E���N�ێ����i�[�u�Ȃǂł��B"
	aData(idx).Add "��p","�󌱗���8,300�~�B" & vbCrLf & "���i��A�o�^�萔���󎆑�1,500�~����������B"
	aData(idx).Add "���i��","��1��F54.7��" & vbCrLf & "��2��F65.6��" & vbCrLf & "�i07�N�x�j"
	aData(idx).Add "�c�̖�","�i���j���S�q���Z�p��������"
	aData(idx).Add "����@��","�Y�Ɣ\����w����������"
	aData(idx).Add "�u�����e","���ʐM����" & vbCrLf & "�i÷�āE�Y��L��j" & vbCrLf & "���\�z���W�E�p�����W�t"
	aData(idx).Add "���i","����u��" & vbCrLf & "2��@23,100�~" & vbCrLf & "���{�����ɑΉ��������H�I���|�[�g���Ŋm���ɍ��i�x���I" & vbCrLf & "���x�����@�F������������ɂ���u�ƂȂ�܂��B"
	aData(idx).Add "���T",""
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�ΏۋƎ�","�S�Ǝ�Ή�"
	aData(idx).Add "���i2","2��@23,100�~"
	aData(idx).Add "�T�v2","�w���Ǝ��i�x(�����J���ȔF��j�Ƃ��āA�S�Ǝ�ɑI�C���`���t�����Ă��鎑�i"
	aData(idx).Add "���T2",""

	'</�q���Ǘ���2��>

	'<��񏈗��Z�p�Ҏ����@�h�s�p�X�|�[�g>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0301"
	aData(idx).Add "�D��",""
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","��񏈗��Z�p�Ҏ���"
	aData(idx).Add "����TITLE","��񏈗��Z�p�Ҏ���"
	aData(idx).Add "�T�v","�y�S�Ǝ�Ή��̎��i�z" & vbCrLf & vbCrLf & "�w���Ǝ��i�x�i�o�ώY�ƏȔF��j�Ƃ��āA�����V�X�A�h�����ɑ���p�\�R�����g���S�Ă̐l���Ώۂ̎��i"
	aData(idx).Add "�ڍ�","�E�Ɛl�N�������ʂɔ����Ă����ׂ����Z�p�Ɋւ����b�I�Ȓm���𑪂�A��񏈗��Z�p�Ҏ����̃��x��1�̎����B" & vbCrLf & "���o�ح�����޼�Ƚ�����{�����A���P�[�g�����i2009�N�Łu���鎑�i�A����Ȃ����i�v�j�ł́A�c�ƐE�Ɏ�点�������i�̏��10�ʒ�6����񏈗��Z�p�Ҏ�������߂Ă����BIT�p�X�|�[�g���i�͑�2�ʁB" & vbCrLf & "��1�ʂ́A��񏈗��Z�p�Ҏ����@��{���Z�p�Ҏ����i���x��2�j"
	aData(idx).Add "��p","�󌱗��́@5,100�~�B"
	aData(idx).Add "���i��","31�D0���i��������ގ����j"
	aData(idx).Add "�c�̖�","�i�Ɓj��񏈗����i�@�\"
	aData(idx).Add "����@��","�Y�Ɣ\����w����������"
	aData(idx).Add "�u�����e","���ʐM����" & vbCrLf & "�i�e�L�X�g�E���W�E���̓e�X�g3��t���j"
	aData(idx).Add "���i","����u��" & vbCrLf & "18,900�~" & vbCrLf & "���V�������x�̃G���g�����x���ɊY�����鎎�����i��ڎw���R�[�X�I" & vbCrLf & "���x�����@�F������������ɂ���u�ƂȂ�܂��B"
	aData(idx).Add "���T",""
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�ΏۋƎ�","�S�Ǝ�Ή�"
	aData(idx).Add "���i2","18,900�~"
	aData(idx).Add "�T�v2","�w���Ǝ��i�x�i�o�ώY�ƏȔF��j�Ƃ��āA�����V�X�A�h�����ɑ���p�\�R�����g���S�Ă̐l���Ώۂ̎��i"
	aData(idx).Add "���T2",""

	'</��񏈗��Z�p�Ҏ����@�h�s�p�X�|�[�g>

	'<�o�^�̔���>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0601"
	aData(idx).Add "�D��",""
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","�o�^�̔���"
	aData(idx).Add "����TITLE","�o�^�̔���"
	aData(idx).Add "�T�v","�y��ǂŎs�̖�̔��̂��d��������]�̕��ɂ��E�߂̎��i�z" & vbCrLf & vbCrLf & "�ŐV�I2009�N6�������򎖖@�̐V���i�I"
	aData(idx).Add "�ڍ�","2009�N�x�������򎖖@���{�s����Ă��܂��B��ʗp���i�͕���p���X�N�ɉ�����3���ނ���A���X�N�̒Ⴂ���ށE��O�ނ́u�o�^�̔��ҁv���̔��ł���悤�ɂȂ�܂����B" & vbCrLf & "�o�^�̔��҂͓s���{�������{����o�^�̔��Ҏ����ɍ��i���邱�Ƃɂ��擾�ł��܂��B" & vbCrLf & "�h���b�O�X�g�A�E�R���r�j�ƊE�ł́A���́u�o�^�̔��ҁv�̊m�ۂƈ琬���d�v�Ȑl�ފJ���ۑ�ɂȂ��Ă������Ƃ��\�z����Ă��܂��I"
	aData(idx).Add "��p",""
	aData(idx).Add "���i��","2008�N�x�@��1�񎎌����i���i��ǐV��2009�N4��1���t���j" & vbCrLf & vbCrLf & "�����@82�D3 (��)" & vbCrLf & vbCrLf & "��ʁ@77�D0 (��)" & vbCrLf & "��t�@80�D0 (��)" & vbCrLf & "�_�ސ�@84�D5 (��)" & vbCrLf & vbCrLf & "�k�C���@54�D8 (��)" & vbCrLf & "�X�@53�D1 (��)" & vbCrLf & "���@43�D0 (��)" & vbCrLf & "�{��@53�D6 (��)" & vbCrLf & "�H�c�@52�D9 (��)" & vbCrLf & "�R�`�@47�D5 (��)" & vbCrLf & "�����@52�D2 (��)" & vbCrLf & vbCrLf & "�V���@75�D4 (��)" & vbCrLf & "�R���@66�D6 (��)" & vbCrLf & "����@75�D5 (��)" & vbCrLf & "���@73�D8 (��)" & vbCrLf & "�Ȗ؁@71�D1 (��)" & vbCrLf & "�Q�n�@77�D6 (��)" & vbCrLf & vbCrLf & "�����@63�D2 (��)" & vbCrLf & "����@55�D7 (��)" & vbCrLf & "�F�{�@62�D9 (��)" & vbCrLf & "�啪�@54�D5 (��)" & vbCrLf & "�{��@63�D9 (��)" & vbCrLf & "�������@56�D0 (��)" & vbCrLf & "����@47�D8 (��)"
	aData(idx).Add "�c�̖�","�e�s���{��"
	aData(idx).Add "����@��","�Y�Ɣ\����w����������"
	aData(idx).Add "�u�����e","���ʐM����" & vbCrLf & "�i÷�āE�Y�킠��j"
	aData(idx).Add "���i","��u��" & vbCrLf & "22,050�~"
	aData(idx).Add "���T","���ʎ�u��" & vbCrLf & "16,800�~"
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�ΏۋƎ�","��ǂŎs�̖�̔��̂��d��������]�̕�"
	aData(idx).Add "���i2","22,050�~"
	aData(idx).Add "�T�v2","2009�N�x�������򎖖@���{�s����Ă��܂��B��ʗp���i�͕���p���X�N�ɉ�����3���ނ���A���X�N�̒Ⴂ���ށE��O�ނ́u�o�^�̔��ҁv���̔��ł���悤�ɂȂ�܂����B"
	aData(idx).Add "���T2","���ʎ�u��<br>16,800�~"

	'</�o�^�̔���>

	'<�e���v���[�g>
	'idx = idx + 1
	'Set aData(idx) = Server.CreateObject("scripting.dictionary")
	'aData(idx).Add "Category","skillup"
	'aData(idx).Add "ID",""
	'aData(idx).Add "�D��",""
	'aData(idx).Add "���",""
	'aData(idx).Add "����",""
	'aData(idx).Add "����@��",""
	'aData(idx).Add "�u�����e",""
	'aData(idx).Add "���i",""
	'aData(idx).Add "���T",""
	'aData(idx).Add "�N�[�|��",""
	'aData(idx).Add "�N�[�|�����ӓ_",""
	'aData(idx).Add "�ΏۋƎ�",""
	'aData(idx).Add "���i2",""
	'aData(idx).Add "�T�v2",""
	'aData(idx).Add "���T2",""

	'</�e���v���[�g>

	'<��w�X�L��>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","skillup"
	aData(idx).Add "ID","0101"
	aData(idx).Add "�D��","1"
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʊw"
	aData(idx).Add "����","�p��b�i�}���c�[�}���j"
	aData(idx).Add "����TITLE","�p��b�i�}���c�[�}���j"
	aData(idx).Add "����@��","Gaba�}���c�[�}���p��b"
	aData(idx).Add "�u�����e","�y�ʊw�X�N�[���z"
	aData(idx).Add "���i",""
	aData(idx).Add "���T","1�D����� ��10,500 �i�ʏ� ��31,500�j" & vbCrLf & "2�D���b�X���������� ��21,000OFF�i�ΏۃR�[�X�F63��ȏ�j" & vbCrLf & "3�DLesson Anywhere �����i��i�ʏ� ��4,200�j" & vbCrLf & "�@�@���ǂ̃X�N�[���ł����b�X������u�ł���I�v�V�����ł��B"
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�N�[�|�����ӓ_",""
	aData(idx).Add "�ΏۋƎ�","�S�Ǝ�Ή�"
	aData(idx).Add "���i2",""
	aData(idx).Add "�T�v2",""
	aData(idx).Add "���T2","<div style=""text-align:left; font-size:11px;"">1�D����� ��10,500<br>2�D���b�X������<br>��21,000OFF<br>3�DLesson Anywhere<br>�����i��"

	'</��w�X�L��>

	'<�p�\�R���X�L���A�b�v>
	'idx = idx + 1
	'Set aData(idx) = Server.CreateObject("scripting.dictionary")
	'aData(idx).Add "Category","skillup"
	'aData(idx).Add "ID","0301"
	'aData(idx).Add "�D��",""
	'aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʊw"
	'aData(idx).Add "����","�p�\�R���X�L���A�b�v"
	'aData(idx).Add "����@��","�A�r�o"
	'aData(idx).Add "�u�����e","�y�ʊw�X�N�[���z"
	'aData(idx).Add "���i",""
	'aData(idx).Add "���T","�S��160�Z�ǂ��ł����T���p�\" & vbCrLf & "�����2���~�ȏ�OFF" & vbCrLf & "��u�����ʊ���5��OFF�I"
	'aData(idx).Add "�N�[�|��","1"
	'aData(idx).Add "�N�[�|�����ӓ_","�E�N�[�|���������p�̍ۂ́A�K���g���ؖ����������Q�������B" & vbCrLf & "�E�X�N�[���ŃJ�E���Z�����O���s���A�R�[�X�ݒ��������œ��T���K�p����܂��B"
	'aData(idx).Add "�ΏۋƎ�",""
	'aData(idx).Add "���i2",""
	'aData(idx).Add "�T�v2",""
	'aData(idx).Add "���T2",""

	'</�p�\�R���X�L���A�b�v>

	'<�p�\�R���X�L���A�b�v>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","skillup"
	aData(idx).Add "ID","0401"
	aData(idx).Add "�D��",""
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʊw"
	aData(idx).Add "����","�p�\�R���X�L���A�b�v"
	aData(idx).Add "����TITLE","�p�\�R���X�L���A�b�v"
	aData(idx).Add "����@��","�Y�Ɣ\����w����������"
	aData(idx).Add "�u�����e","�y�ʐM�u���z"
	aData(idx).Add "���i","�������R�[�X" & vbCrLf & "�@�Z����V���[�Y" & vbCrLf & "�@�@�yOffice�Z����qOffice2003�E2002�E2000�r�z�@��ʉ��i 16,800�~ ��30���ȏ�OFF" & vbCrLf & "�@�@��Word�EExcel�EPowerPoint���g�����Ȃ��R�[�X�ł��I" & vbCrLf & "" & vbCrLf & "��Excel�EWord�R�[�X" & vbCrLf & "�@���āE����āE�ȒP" & vbCrLf & "�@�@�yExcel2002�����z�@��ʉ��i 19,950�~" & vbCrLf & "�@�@�yWord2002�����z�@��ʉ��i 19,950�~" & vbCrLf & "�@�@�yExcel�EWord2002��b�z�@��ʉ��i 19,950�~" & vbCrLf & "�@�@�yExcel�EWord2003���p�z�@��ʉ��i 19,950�~" & vbCrLf & "" & vbCrLf & "��PowerPoint�R�[�X" & vbCrLf & "�@�y�g����v���[���IPowerPoint�z�@��ʉ��i 19,950�~" & vbCrLf & "" & vbCrLf & "���C���^�[�l�b�g�R�[�X" & vbCrLf & "�@�y�C���^�[�l�b�g�Z�L�����e�B�[�h���h����z�@��ʉ��i 13,650�~" & vbCrLf & "�@�y�C���^�[�l�b�g���p�z�@��ʉ��i 13,650�~"
	aData(idx).Add "���T","�������R�[�X" & vbCrLf & "�@�Z����V���[�Y" & vbCrLf & "�@�@�yOffice�Z����qOffice2003�E2002�E2000�r�z�@���ʉ��i 11,550�~(�ʏ�F16,800�~) ��30���ȏ�OFF" & vbCrLf & "�@�@��Word�EExcel�EPower�o�����������g�����Ȃ��R�[�X�ł��I" & vbCrLf & "" & vbCrLf & "��Excel�EWord�R�[�X" & vbCrLf & "�@���āE����āE�ȒP" & vbCrLf & "�@�@�yExcel2002�����z�@���ʉ��i 14,700�~(�ʏ�F19,950�~) ��26��OFF" & vbCrLf & "�@�@�yWord2002�����z�@���ʉ��i 14,700�~(�ʏ�F19,950�~) ��26��OFF" & vbCrLf & "�@�@�yExcel�EWord2002��b�z�@���ʉ��i 14,700�~(�ʏ�F19,950�~) ��26��OFF" & vbCrLf & "�@�@�yExcel�EWord2003���p�z�@���ʉ��i 14,700�~(�ʏ�F19,950�~) ��26��OFF" & vbCrLf & "" & vbCrLf & "��PowerPoint�R�[�X" & vbCrLf & "�@�y�g����v���[���IPowerPoint�z�@���ʉ��i 14,700�~(�ʏ�F19,950�~) ��26��OFF" & vbCrLf & "" & vbCrLf & "���C���^�[�l�b�g�R�[�X" & vbCrLf & "�@�y�C���^�[�l�b�g�Z�L�����e�B�[�h���h����z�@���ʉ��i 8,400�~(�ʏ�F13,650�~) ��38��OFF" & vbCrLf & "�@�y�C���^�[�l�b�g���p�z�@���ʉ��i 8,400�~(�ʏ�F13,650�~) ��38��OFF"
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�N�[�|�����ӓ_",""
	aData(idx).Add "�ΏۋƎ�","�S�Ǝ�Ή�"
	aData(idx).Add "���i2","Office�Z����R�[�X�Ȃ�<br>(�ʏ�F16,800�~�`)"
	aData(idx).Add "�T�v2",""
	aData(idx).Add "���T2","���ʉ��i<br>�ő�38��OFF<br>�Ȃ�"

	'</�p�\�R���X�L���A�b�v>

	'<��w�X�L��>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","skillup"
	aData(idx).Add "ID","0201"
	aData(idx).Add "�D��",""
	aData(idx).Add "���","�r�W�l�X(���i)�n �^ �ʐM����"
	aData(idx).Add "����","TOEIC���H�g���[�j���O" & vbCrLf & "750�ر�E650�ر�E550�ر�E450�ر�̊e�R�[�X�ݒ�"
	aData(idx).Add "����TITLE","TOEIC���H�g���[�j���O"
	aData(idx).Add "����@��","�Y�Ɣ\����w����������"
	aData(idx).Add "�u�����e","�y�ʐM�u���z"
	aData(idx).Add "���i","750�N���A �� 31,500�~" & vbCrLf & "650�N���A �� 23,100�~" & vbCrLf & "550�N���A �� 22,050�~" & vbCrLf & "450�N���A �� 21,000�~" & vbCrLf & "���x�����@�F������������ɂ���u�ƂȂ�܂��B"
	aData(idx).Add "���T",""
	aData(idx).Add "�N�[�|��",""
	aData(idx).Add "�N�[�|�����ӓ_",""
	aData(idx).Add "�ΏۋƎ�","�S�Ǝ�Ή�"
	aData(idx).Add "���i2","750�N���A �� 31,500�~<br>650�N���A �� 23,100�~<br>550�N���A �� 22,050�~<br>450�N���A �� 21,000�~"
	aData(idx).Add "�T�v2",""
	aData(idx).Add "���T2",""

	'</��w�X�L��>

	getWarmreceptionMaster = aData
End Function
%>
