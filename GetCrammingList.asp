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

  Const	PUBLIC_CONNECTION = "Server=tcp:route55.database.windows.net,1433;Initial Catalog=WordTrainer;Persist Security Info=False;User ID=WebPublicUser;Password=rXhsQiwXCNAmmNrb6OmF1KUPtZn4qV6TD9xbkHdNxeU=;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"
  
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
