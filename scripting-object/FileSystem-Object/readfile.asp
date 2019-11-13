<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>ASP-Read text file</h2>

        <%
            dim filesystem ,file
            set filesystem = Server.CreateObject("Scripting.FileSystemObject")
            '.OpenTextFile(fname,mode,create,format)
            'mode
            '1=ForReading - Open a file for reading. You cannot write to this file.
            '2=ForWriting - open a file for writing.
            '3=ForAppending -open a file and write to the end of the file.

            set file= filesystem.OpenTextFile("F:\abc.txt",1)
            response.write file.ReadAll()

            set file = nothing
        %>
        <hr/>
        <%
            dim f
            set f= filesystem.GetFile("F:\abc.txt")
            response.write("File created:" & f.DateCreated & "<br>")
            response.write("File Last Modified:" & f.DateLastModified & "<br>")
            response.write("File Type:" & f.Type & "<br>")
            response.write("File size (bytes): " & f.size & "<br>")

        %>
