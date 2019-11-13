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
            dim x, y, z,a, b, c, m
            x=10
            y=34.5
            z="hey"
            a= true
            b= #2019/12/23#
            c= #23:34:23#

            
            
            'Array Types
            dim arr(10)
            arr(0)="first item"
            arr(2)=2367
            Response.Write"<br/>"&arr(0)

            dim arr1
            arr1= Array(123, 322, 3423,544)

            dim mat(2,3) '3 rows,  columns
            mat(0,0) = "Apple"
            mat(0,1) = "Apple"
            mat(0,1) = "Apple"
            mat(0,2) = "Apple"
            mat(0,3) = "Apple"
            mat(1,0) = "Apple"

            Response.Write LBound(arr)
            Response.Write UBound(arr)

            'object pointers
            dim objDictionary
            set objDictionary = Server.createobject("Scripting.Dictionary")
            set objDictionary = nothing

        %>
        </body>
    </html>