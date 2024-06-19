<%
'*****************************************************
'** 企業向けしごとナビ応募情報のファックス通知機能
'** 
'** 変数一覧
'**		vCn	:	会社名
'**		vPn	:	企業担当者氏名
'**		vSc	:	スタッフコード
'**		vOc	:	求人情報コード
'**		vSj	:	件名
'**		vBd	:	本文
'**		vFn	:	ファックス番号
'**	返り値
'** 	正常終了 : true 、異常 : Err.Number
'*****************************************************
function FaxLetterNaviCompany(vCn,vPn,vSc,vOc,vSj,vBd,vFn)
	
	On Error Resume Next
	
'*******************************************
'FAX送信用文書をExcelフォーマットで生成する。
'*******************************************
'	ReportFolder			 各種帳票フォーマット保管場所 personnel.aspに記述
	'作成ファイル名
	Dim wOutFileName	:	wOutFileName	=	"応通" & vOc & year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) & Minute(Now()) & Second(Now())
	'保存先のフォルダ
	Dim wSaveFolder		:	wSaveFolder		=	"\\192.168.10.61\fax送信文書"
	
	
	response.write Err.Number
	response.write Err.Description
	
	Dim wErrNo
	Dim wMsg
	Dim Xlsx1
	
	'ExcelCreatorオブジェクト生成
	Set Xlsx1 = Server.CreateObject("XlsxCrt.XlsxCrtCtrl.1")
	
	'Excelファイル（売上伝票）オーバレイオープン
	Xlsx1.OpenBook wSaveFolder & "\" & wOutFileName & ".xlsx", ReportFolder & "\しごとナビ応募連絡通知文書（企業向け）.xlsx"
	
	Xlsx1.SheetNo = 0
    
    Xlsx1.Cell("A6").Value = vCn & "様。"
    Xlsx1.Cell("A7").Value = "求職者（" & vSc & ")から、しごとナビを通じて" & _
    						vOc & "の求人について連絡が入りましたのでＦＡＸにてお伝え致します。"
    
    Xlsx1.Cell("B11").Value = vSj
    Xlsx1.Cell("B12").Value = vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd & vBd
    
	If wErrNo <> 0 Then
		FaxLetterNaviCompany = "ExcelCreator3.6 エラーメッセージ：" & Xlsx1.ErrorMessage
		exit function
	End If
	Xlsx1.CloseBook
	Set Xlsx1 = Nothing
	
'*******************************************
'FAX送信用CSVファイルを生成する。
'*******************************************
	Dim wCsvFolder	:	wCsvFolder = "\\192.168.10.61\CsvShare"
	
	Set objFS = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFS.GetFolder(wCsvFolder)
	Set objFile = objFolder.CreateTextFile(wOutFileName & ".csv")
	objFile.WriteLine("""C:\FAX送信文書\" & wOutFileName & ".xlsx" & """,""" & vFn & """,""" & vPn & ""","" " & "様" & " "",""" & vCn & ""","""","""",""""")
	objFile.Close
	
'*******************************************
'エラーチェック
'*******************************************
	response.write wSaveFolder & "\" & wOutFileName
	response.write "<br>"
	response.write Err.Number
	response.write Err.Description
	
end function
%>