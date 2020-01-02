<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
    
<%@ page import="java.sql.*" %>

<html>
	<head>
		<meta charset="utf-8">
		<title>Admin Login Form</title>
		<link rel="stylesheet" href="style.css">
	</head>
	<body background="bg.jpg">
	
         <div class="menuBox">
			<h1>Data is successfully inserted</h1>
			<h1><a href="adminSubject.jsp">Click here</a> to go your site.</h1>
		</div>
		
		<%
		
		for(int i=0;i<10;i++)
		{
			Connection conn = null;
		 	Statement st = null;
		 	
			String nmber_qstn = "";
			nmber_qstn = Integer.toString(i);
			String qstn_no = request.getParameter(nmber_qstn);
			
			
			String[] op = new String[10];
			
			for(int j=1;j<=4;j++)
			{
				String nmber_optn = "", nm = "";
				nmber_optn = Integer.toString(j);
				nm += nmber_qstn;
				nm += nmber_optn;
				
				String optn_no = request.getParameter(nm);
				op[j-1] = optn_no;
			}
			
			
			
			String nm_ans = "";
			String nmber_ans = Integer.toString(5);
			nm_ans += nmber_qstn;
			nm_ans += nmber_ans;
			
			String ans_no = request.getParameter(nm_ans);
			
			
			try{
				  Class.forName("com.mysql.jdbc.Driver");
				  conn = (Connection)DriverManager.getConnection("jdbc:mysql://localhost:3306/jspproject","root","");
				  st = conn.createStatement();
				  
				  String qry = "INSERT INTO `add_qstn4` (`id`, `question`, `Op1`, `Op2`, `Op3`, `Op4`, `answer`) VALUES (NULL, '"+qstn_no+"', '"+op[0]+"', '"+op[1]+"', '"+op[2]+"', '"+op[3]+"', '"+ans_no+"')"; 
				  
				  st.executeUpdate(qry); 
				  //out.print("Data is successfully inserted");
				  
			  }catch(Exception ex)
			  {
				  out.print(ex.getMessage()); 
			  }
			
		}
	  
	  
	
	
	%>
		
	</body>
</html>

