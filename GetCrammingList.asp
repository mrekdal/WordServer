<!-- #include virtual="Include/Config.asp" -->
<%

Response.AddHeader "Access-Control-Allow-Origin","*"
Response.ContentType = "application/json"
Response.AddHeader "Content-Type", "application/json;charset=UTF-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8"

  Dim strSQL
  Dim oConn
  Dim oRs
  Dim lAntallPoster
  
  Set oConn = Server.CreateObject("ADODB.Connection")
  oConn.Open PUBLIC_CONNECTION
  
  strSQL = "EXEC WebService.GetCrammingList " & CInt( Request("ListId") & ",2,3" )
  Set oRs = oConn.Execute( strSQL, lAntallPoster, 1 )
  do while not oRs.EOF
    for each fld in oRs.Fields
      Response.Write( fld.Value )
    next
    oRs.MoveNext
  loop
%>
