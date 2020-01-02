<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>

<%@ page import="java.sql.*" %>


<html>
	<head>
		<meta charset="utf-8">
		<title>Student Login Form</title>
		<link rel="stylesheet" href="style.css">
	</head>
	<body background="bg4.jpg">
	
         <div class="menuBox">
			<h1>You have been successfully Registered.</h1>
			<h1><a href="login.jsp">Click here</a> to login</h1>
		</div>
		
		<%
	  String name = request.getParameter("t1");
	  String email = request.getParameter("t2");
	  
	  Connection conn = null;
	  Statement st = null;
	  
	  try{
		  Class.forName("com.mysql.jdbc.Driver");
		  conn = (Connection)DriverManager.getConnection("jdbc:mysql://localhost:3306/jspproject","root","");
		  st = conn.createStatement();
		  
		  String qry = "INSERT INTO `tbl_user` (`id`, `name`, `email`) VALUES (NULL, '"+name+"', '"+email+"')"; 
		  
		  st.executeUpdate(qry); 
		  //out.print("Data is successfully inserted");
		  
	  }catch(Exception ex)
	  {
		  out.print(ex.getMessage()); 
	  }
	
	
	%>
		
	</body>
</html>