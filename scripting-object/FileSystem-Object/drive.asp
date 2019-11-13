<% @ language="VBScript" %>
<% option explicit %>

<html>
    <head>
        <title>ASP introduction</title>
    </head>
    <body>
        <h1>Welcome To ASP Programming</h1>
        <h2>ASP- Drive meta information</h2>

        <%
        dim fileSystem, drive ,driveletter
        set fileSystem = Server.createObject("Scripting.FileSystemObject")
 
        for each driveletter in fileSystem.drives 
            set drive = fileSystem.GetDrive(driveletter)
            if drive.IsReady=true then
                response.write drive.DriveLetter & "<br>"
                response.write drive.DriveType & "<br>"
                response.write drive.FileSystem & "<br>"
                response.write drive.TotalSize & "<br>"
                response.write drive.AvailableSpace & "<br><hr>"
            end if
        next
            '0=unknown
            '1=removable
            '2= fixed
            '3= network
            '4= CD-ROM
            '5= RAM disk
        %>
             <h4>AvailableSpace and TotalSize</h4>
            <table>
            <tr>
                <th>DriveLetter</th>
                <th>DriveType</th>
                <th>FileSystem</th>
                <th>TotalSize</th>
                <th>AvailableSpace</th>
            </tr>
        <% 
            for each driveletter in fileSystem.drives 
            set drive = fileSystem.GetDrive(driveletter)
            if drive.IsReady then
            
        %>
        <tr>
            <td><%=drive.DriveLetter%></td>
            <td><%=drive.DriveType%></td>
            <td><%=drive.FileSystem%></td>
            <td><%=drive.TotalSize / (1024 * 1024 *1024)%></td>
            <td><%=drive.AvailableSpace / (1024 * 1024 *1024)%></td>

        </tr>
        <%
        end if
        next 

        function GetDriveType(driveletter)
            select case driveletter 
                case 0:GetDriveType = "unknown" 
                case 1:GetDriveType = "removable" 
                case 2:GetDriveType = "fixed" 
                case 3:GetDriveType = "network" 
                case 4:GetDriveType = "CD-ROM" 
                case 5:GetDriveType = "RAm disk" 
            end select
        end function

        %>  
            </table>
    
        
    
    </body>
    
    </html>
    