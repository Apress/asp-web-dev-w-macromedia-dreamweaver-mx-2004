<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/conn_webprodmx.asp" -->
<%
Dim rsBooks__varCategoryID
rsBooks__varCategoryID = "0"
If (Request.QueryString("CategoryID")  <> "") Then 
  rsBooks__varCategoryID = Request.QueryString("CategoryID") 
End If
%>
<%
Dim rsBooks
Dim rsBooks_numRows

Set rsBooks = Server.CreateObject("ADODB.Recordset")
rsBooks.ActiveConnection = MM_conn_webprodmx_STRING
rsBooks.Source = "SELECT BookID, BookTitle, BookAuthorFirstName, BookAuthorLastName, BookPrice, BookImage  FROM dbo.tbl_Books  WHERE CategoryID = " + Replace(rsBooks__varCategoryID, "'", "''") + "  ORDER BY BookTitle"
rsBooks.CursorType = 0
rsBooks.CursorLocation = 2
rsBooks.LockType = 1
rsBooks.Open()

rsBooks_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 4
Repeat1__index = 0
rsBooks_numRows = rsBooks_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsBooks_total
Dim rsBooks_first
Dim rsBooks_last

' set the record count
rsBooks_total = rsBooks.RecordCount

' set the number of rows displayed on this page
If (rsBooks_numRows < 0) Then
  rsBooks_numRows = rsBooks_total
Elseif (rsBooks_numRows = 0) Then
  rsBooks_numRows = 1
End If

' set the first and last displayed record
rsBooks_first = 1
rsBooks_last  = rsBooks_first + rsBooks_numRows - 1

' if we have the correct record count, check the other stats
If (rsBooks_total <> -1) Then
  If (rsBooks_first > rsBooks_total) Then
    rsBooks_first = rsBooks_total
  End If
  If (rsBooks_last > rsBooks_total) Then
    rsBooks_last = rsBooks_total
  End If
  If (rsBooks_numRows > rsBooks_total) Then
    rsBooks_numRows = rsBooks_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsBooks_total = -1) Then

  ' count the total records by iterating through the recordset
  rsBooks_total=0
  While (Not rsBooks.EOF)
    rsBooks_total = rsBooks_total + 1
    rsBooks.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsBooks.CursorType > 0) Then
    rsBooks.MoveFirst
  Else
    rsBooks.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsBooks_numRows < 0 Or rsBooks_numRows > rsBooks_total) Then
    rsBooks_numRows = rsBooks_total
  End If

  ' set the first and last displayed record
  rsBooks_first = 1
  rsBooks_last = rsBooks_first + rsBooks_numRows - 1
  
  If (rsBooks_first > rsBooks_total) Then
    rsBooks_first = rsBooks_total
  End If
  If (rsBooks_last > rsBooks_total) Then
    rsBooks_last = rsBooks_total
  End If

End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = rsBooks
MM_rsCount   = rsBooks_total
MM_size      = rsBooks_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsBooks_first = MM_offset + 1
rsBooks_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsBooks_first > MM_rsCount) Then
    rsBooks_first = MM_rsCount
  End If
  If (rsBooks_last > MM_rsCount) Then
    rsBooks_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Book details</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/JavaScript">
<!--
function tmt_confirm(msg){
	document.MM_returnValue=(confirm(unescape(msg)));
}
//-->
</script>
</head>
<body>
<a href="categories3.asp">Back to categories</a><table width="450" border="0" cellspacing="0" cellpadding="0">
  <% If rsBooks.EOF And rsBooks.BOF Then %>
  <tr>
    <td>&nbsp;</td>
    <td>Sorry, no records found </td>
  </tr>
  <% End If ' end rsBooks.EOF And rsBooks.BOF %>
  <% If Not rsBooks.EOF Or Not rsBooks.BOF Then %>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp; Records <%=(rsBooks_first)%> to <%=(rsBooks_last)%> of <%=(rsBooks_total)%> </td>
  </tr>
  <tr>
    <td>
      <% If MM_offset <> 0 Then %>
          <A HREF="<%=MM_movePrev%>">Previous</A>
          <% End If ' end MM_offset <> 0 %></td>
    <td>
      <% If MM_atTotal Then %>
          <A HREF="<%=MM_moveNext%>">Next</A>
          <% End If ' end MM_atTotal %></td>
  </tr>
  <tr>
    <td width="150">Book cover </td>
    <td width="300">Book details </td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsBooks.EOF)) 
%>
  <tr align="left" valign="top">
    <td><img src="media/bookimages/<%=(rsBooks.Fields.Item("BookImage").Value)%>"></td>
    <td><p><%=(rsBooks.Fields.Item("BookTitle").Value)%></p>
        <p><%=(rsBooks.Fields.Item("BookAuthorFirstName").Value)%>&nbsp;<%=(rsBooks.Fields.Item("BookAuthorLastName").Value)%></p>
        <p><%= FormatCurrency((rsBooks.Fields.Item("BookPrice").Value), -1, -2, -2, -2) %></p>
        <p><a href="update_books.asp?BookID=<%=(rsBooks.Fields.Item("BookID").Value)%>">Update</a>&nbsp;<a href="delete_book.asp?BookID=<%=(rsBooks.Fields.Item("BookID").Value)%>" onClick="tmt_confirm('Are%20you%20sure?');return document.MM_returnValue">Delete</a></p></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsBooks.MoveNext()
Wend
%>
<% End If ' end Not rsBooks.EOF Or NOT rsBooks.BOF %>
</table>
</body>
</html>
<%
rsBooks.Close()
Set rsBooks = Nothing
%>