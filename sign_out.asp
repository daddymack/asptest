<!--#include file="oauth/_inc/_base.asp"-->
<%
	Session(TWITTER_SCREEN_NAME) = ""
	Session.Contents.Remove(TWITTER_SCREEN_NAME)
	Session.Contents.RemoveAll()
	Session.Abandon()

	Response.ContentType = "text/javascript"
	Response.Write "<script></script>"
	Response.End
%>
