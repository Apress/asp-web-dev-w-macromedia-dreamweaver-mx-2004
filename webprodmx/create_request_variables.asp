<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Response.Cookies("ckHello") = "Hello" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Create request variables</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<p><a href="create_request_variables.asp?age=30&name=rob">create query_string parameters</a></p>
<form name="form1" method="post" action="">
  <input name="textfield" type="text" value="<%= Request.QueryString("name") %>">
  <input name="hiddenField" type="hidden" value="<%= Request.QueryString("age") %>">
</form>
<hr>
<form name="form2" method="post" action="create_request_variables.asp">
<p>Username: <input name="username" type="text" id="username" /></p>
<p>Password: <input name="password" type="password" id="password" /></p>
<p><input type="submit" name="Submit" value="Submit" /></p>
</form>
<hr>
<p><%= Request.Cookies("ckHello") %></p>
<hr>
<p><%= Request.ServerVariables("remote_host") %></p>
</body>
</html>