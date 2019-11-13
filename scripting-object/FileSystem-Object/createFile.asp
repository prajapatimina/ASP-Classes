<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>ASP-Create and write text file</h2>
        <%
            'The FileSystemObject object is used to access the file system on a server.
            'We can manipulate files, folders, drives and directory paths.

            dim fileSystem , currDirPath,file,folderPath,folder
        
            currDirPath = Server.MapPath(".")
            folderPath = currDirPath & "/TextFiles"

            set fileSystem = Server.CreateObject("Scripting.FileSystemObject")
            if fileSystem.FolderExists(folderPath) = false then
                set folder=fileSystem.CreateFolder(folderPath)
            end if
            set file= fileSystem.CreateTextFile(currDirPath & "\test.txt")

            file.write("Hello World")   'writes in same line
            file.writeLine("Hello World")
             file.writeLine("Hello World")

            file.Close
            set file=nothing

            set folder = nothing
            set fileSystem = nothing
            response.write "text file created successfully." 
            
              %>
            
            
            
