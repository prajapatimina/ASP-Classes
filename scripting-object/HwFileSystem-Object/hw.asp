<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1> 
        <%
        dim i, currentFolder, nextFolder, nextFile, fileSystem, folder, file

        set fileSystem = Server.CreateObject("Scripting.FileSystemObject")
        currentFolder = Server.MapPath(".")

        for i=97 to 122
            nextFolder = currentFolder & "/" & Chr(i)

             if fileSystem.FolderExists(nextFolder) = false then
                set folder = fileSystem.CreateFolder(nextFolder)
                set file= fileSystem.CreateTextFile(nextFolder & "\" & Chr(i) & ".txt")
            end if


    
        next

       

        %>