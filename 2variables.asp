<%@Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1> 
        <h2>VBScript Variables</h2>
        <%
            'dim = dimension, declare in memory
            'only one type of varaible  - variant type (dynamic type)
            'which can hold 3 kinds  of values :scalar values,arrays and object pointer

            'scalar types 
            dim x, y, z,a, b, c
            x=10
            y=34.5
            z="hey"
            a= true
            b= #2019/12/23#
            c= #23:34:23#

            Response.Write x
            Response.Write"<br/>"
            Response.Write y
            Response.Write "<br/>"
            Response.Write z
            Response.Write "<br/>"
            Response.Write a
            Response.Write "<br/>"
            Response.Write b
            Response.Write "<br/>"
            Response.Write c
            
            


        %>
        </body>
    </html>