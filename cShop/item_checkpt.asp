<!-- #Include virtual="/inc/conn.asp" -->
<%
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = -1
Response.CacheControl = "no-cache"

Dim PostID, PostKEY, PostWORLD
Dim CheckStr, Account_Code, Account_Price, Error_Msg

PostID = GetPostField("id")
PostKEY = GetPostField("key")
PostWORLD = GetPostInt("world")

'PostID = "caonimabi"
'PostKEY = "53dc3b60f2d40cd4"
'PostWORLD = 1



If PostID = "" or PostKEY = "" or PostWORLD = 0 Then
	response.Write("FAIL46")
	response.End()
End If
If len(PostID) < 4 or len(PostID) > 12 Then
	response.Write("FAIL200")
	response.End()
End If
If len(PostKEY) <> 16 Then
	response.Write("FAIL210")
	response.End()
End If
If PostWORLD < 0 or PostWORLD > 10 Then
	response.Write("FAIL220")
	response.End()
End If

CheckStr = 0
If ChkInvaildWord(PostID) = True Then CheckStr = 1
If ChkInvaildWord(PostKEY) = True Then CheckStr = 1
If ChkInvaildWord(PostWORLD) = True Then CheckStr = 1

If CheckStr = 1 Then
	response.Write("FAIL230")
	response.End()
Else
	Set rsConn = server.createobject("ADODB.Recordset")
	rsSQL = "DECLARE @return_value int;EXEC @return_value = [dbo].[uspn_get_user_authority_check] '"& PostID &"','"& PostKEY &"','"& getip() &"';SELECT 'Return_Value' = @return_value"
	'response.Write(rsSQL & vbCrLf)
	rsConn.open rsSQL,Conn,3,2
	Error_Msg = rsConn("Return_Value")
	set rsSQL = NoThing
	rsConn.close
	Select Case Error_Msg
		Case 0
			Set rsConn = server.createobject("ADODB.Recordset")
			rsSQL = "SELECT ACCOUNT_CODE FROM [dbo].[tbl_account] WHERE ACCOUNT_GID = '"& PostID &"'"
			'response.Write(rsSQL & vbCrLf)
			rsConn.open rsSQL,Conn,1,3
			Account_Code = rsConn("ACCOUNT_CODE")
			set rsSQL = NoThing
			rsConn.close

			Set rsConn = server.createobject("ADODB.Recordset")
			rsSQL = "SELECT PRICE FROM [dbo].[_stt_sell_webshop] WHERE ACCOUNT_CODE = "& Account_Code &""
			'response.Write(rsSQL & vbCrLf)
			rsConn.open rsSQL,Conn,1,3
			If rsConn.EOF Then
				Account_Price = 0
			Else
				Account_Price = rsConn("PRICE")
			End If
			set rsSQL = NoThing
			rsConn.close

			response.Write(Account_Price)
			'response.Write("item_checkpt.asp" & vbCrLf)
			'response.Write(Request.Form() & vbCrLf)
		Case 103
			response.Write("FAIL40")
			response.End()
		Case Else
			response.Write("FAIL" & Error_Msg)
			response.End()
	End Select
	If Error_Msg <> 0 Then
		'response.Write(Error_Msg & vbCrLf)
	End If
End If
%>
<!-- #Include virtual="/inc/close.asp" -->