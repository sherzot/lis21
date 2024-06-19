<%
Sub RegPersonEdit8()
	Dim oP_ResumeStudent:	Set oP_ResumeStudent = New clsP_ResumeStudent
	Dim oRS
	Dim idx
	Dim bError:	bError = False
	Dim sResult

	Call oP_ResumeStudent.Initialize()

	Session("reg_page") = "pe8"
	If IsObject(Session("pe8")) = False Then Set Session("pe8") = Server.CreateObject("scripting.dictionary")

	'******************************************************************************
	'** form Žæ“¾ start
	'******************************************************************************
	For idx = 1 To Request.Form.Count
		If Session("pe8").Exists(Request.Form.Key(idx)) = True Then
			Session("pe8")(Request.Form.Key(idx)) = Request.Form(idx)
		Else
			Call Session("pe8").Add(Request.Form.Key(idx), Request.Form(idx))
		End If
	Next
	'******************************************************************************
	'** form Žæ“¾ end
	'******************************************************************************

	'******************************************************************************
	'** ƒGƒ‰[ˆ— Žæ“¾ end
	'******************************************************************************
	Session.Contents.Remove("errstyle")
	Set Session("errstyle") = oP_ResumeStudent.ErrStyle
	Session("errstyle").CompareMode = 1
	Session("errorstring") = oP_ResumeStudent.Err

	If oP_ResumeStudent.IsData = False Then Response.Redirect "./pe8.asp"
	'******************************************************************************
	'** ƒGƒ‰[ˆ— Žæ“¾ end
	'******************************************************************************

	Set oRS = QEXE(dbconn, oP_ResumeStudent.GetRegSQL(Session("userid")))
	If GetRSState(oRS) = True Then
		sResult = ChkStr(oRS.Collect("StaffCode"))
		Session.Contents.Remove("pe8")
	End If

	If sResult = "" Then
		Session("errorstring") = "“o˜^‚ÉŽ¸”s‚µ‚Ü‚µ‚½B"
		Response.Redirect "./person_edit8.asp"
	End If
End If
%>
