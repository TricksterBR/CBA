<!-- #Include virtual="/inc/conn.asp" -->
<%
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim PostID, PostKEY, PostWORLD
Dim CheckStr, Error_Msg

PostID = GetPostField("id")
PostKEY = GetPostField("key")
PostWORLD = GetPostInt("world")

'PostID = "ntreev"
'PostKEY = "965eb72c92a549dd"
'PostWORLD = 1

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
			response.Write("<body scroll=""no"" oncontextmenu=""window.event.returnValue=false"">" & vbCrLf)
			response.Write("<div align=""center"">功能暫未開放！</div>" & vbCrLf)
			response.Write("</body>" & vbCrLf)
			'response.Write("gacha.asp" & vbCrLf)
			'response.Write(Request.Form() & vbCrLf)
		Case 103
			response.Write("FAIL40")
		Case Else
			response.Write("FAIL" & Error_Msg)
	End Select
	If Error_Msg <> 0 Then
		'response.Write(Error_Msg & vbCrLf)
	End If
End If
%>
<!-- #Include virtual="/inc/close.asp" -->