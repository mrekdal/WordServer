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

  Const PUBLIC_CONNECTION = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=WordTrainer;Data Source=tcp:route55.database.windows.net,1433;User ID=WebPublicUser@route55;Password=rXhsQiwXCNAmmNrb6OmF1KUPtZn4qV6TD9xbkHdNxeU=;"
  
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
