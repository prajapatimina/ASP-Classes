<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>ADO- connecting to Database</h2>
         <%
        dim conn,rs,x,sql,rowsAffected
       set conn=server.CreateObject("ADODB.Connection")
        
        conn.ConnectionString="Provider=SQLOLEDB;Server=.\SQLEXPRESS;Initial Catalog=MBM; user id='mina12';password='123';"
        conn.Open

        response.write "Connected"

        sql= "Insert into Student1(Id,FirstName, LastName,Address,Phone) Values (5,'john', 'Elton','baneswor',12330098)"
        conn.Execute sql, rowsAffected
        response.write rowsAffected & "Rows Inserted"
        conn.close
        %>
        </body>
        </html>