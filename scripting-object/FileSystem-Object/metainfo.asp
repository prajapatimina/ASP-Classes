<% @ language="VBScript" %>
<% option explicit %>

<html>
    <head>
        <title>ASP introduction</title>
    </head>
    <body>
        <h1>Welcome To ASP Programming</h1>
        <h2>ASP- Create and write text file</h2>

        <%
            dim fileSystem,f
            set fileSystem = server.createObject("scripting.FileSystemObject")
            set f = fileSystem.GetFile("F:\ASP classes\scripting-Object\FileSystem-Object\Test.txt")
            
            response.write("File created: " & f.DateCreated & "</br>")
            response.write("File Modified: " & f.DateLastModified & "</br>")
            response.write("File Type: " & f.Type & "</br>")
            response.write("File Size(byte): " & f.Size & "</br>")

            set f = nothing
        %>

        <hr/>
    
        <%
            dim fo, x

            set fo = fileSystem.getFolder("F:\Asp Classes\scripting-Object")
            response.write("Folder Name: " & fo.shortName & "<br/>")
            response.write("Folder Created: " & fo.DateCreated & "<br/>")
            response.write("Folder Last Modified: " & fo.DateLastModified & "<br/>")
            response.write("Size(bytes): " & fo.Size & "<br/>")

           ' set fo=nothing

          
            
            %>
            <h4>  Files in this folder:</h4>
            <table>
            <tr>
                <th>Name</th>
                <th>File Type</th>
                <th>Size</th>
                <th>Created Date</th>
                <th>Last Modified</th>
            </tr>

          <%  for each x in fo.Files%>
            <tr>
                
                <td><%=x.Name %></td>
                <td><%=x.Type %></td>
                  <td><%=x.Size %></td>
                  <td><%=x.DateCreated %></td>
                 <td><%=x.DateLastModified %> </td>
                 </tr>
               <%next%>
            </table>
        


    
    </body>
</html>


