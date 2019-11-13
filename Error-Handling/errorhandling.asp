<% @ language="VBScript" %>
<% option explicit %>

<html>
    <head>
        <title>ASP introduction</title>
    </head>
    <body>
        <h1>Welcome To ASP Programming</h1>
        <h2>Error handling</h2>
        <%
            dim fso, textstream, allcontent, filePath
            filePath = "D:\asd.txt"
            on error resume next
                set fso = server.createobject("scripting.filesystemobject")
                set textstream = fso.opentextfile(filepath,1)
                allcontent = textstream.readall()
                textstream.close()
            if Err.Number <> 0 then
                response.write "Error :" & Err.Number & "<br>"
                response.write "Source :" & Err.Source & "<br>"
                response.write "Description :" & Err.Description & "<br>"
                response.write filePath & "does not exist" & "<br>"
            end if
        %>