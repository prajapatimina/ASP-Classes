<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>ADO-Database</h2>
        <%
        dim conn,rs,x
       set conn=server.CreateObject("ADODB.Connection")
        
        conn.ConnectionString="Provider=SQLOLEDB;Server=.\SQLEXPRESS;Initial Catalog=MBM; user id='mina12';password='123';"
        conn.Open
        response.write "Connected Successfully."

        set rs =Server.CreateObject("ADODB.Recordset")
        rs.open "Select * From Student1",conn

        do until rs.EOF
            for each x in rs.Fields
                response.write x.name &"="
                response.write x.value & "<br/>"

            next
            rs.MoveNext
        loop
        rs.close
        conn.close
          
        %>
        
        
    </body>
    
    <!--   SQL server authentication ko lagi
    conn.ConnectionString="Provider=SQLOLEDB; Server=.\SQLEXPRESS;"_
       "Database=MBM;User Id=Mina;Password=12345;"
        -->
    </html>