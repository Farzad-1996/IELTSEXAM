<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>

<%@ page import="java.sql.*" %>

<!DOCTYPE html>
<html>
	<head>
		<meta charset="utf-8">
		<title>Welcome Page</title>
		<link rel="stylesheet" href="style.css">
    </head>
    
	<body background="bg4.jpg">
	
		<ul>
		    <li><a href="logoutStudent.html">Logout</a></li>
		    <li><a href="score.jsp">My Score</a></li>
		    <li><a href="welcome.jsp" class="active">Subjects</a></li>
            <li><a href="#">About Us</a></li>
            <li><a href="index.html">Home</a></li> 
        </ul>

       <div class="menuBox">
       <p><center><h1 style="color:#9AC303;">Sit for exam on the following subjects</h1></center></p>
			
				<%-- Listening Start --%>
				<%
					String email = (String) session.getAttribute("email");	
    		   
    		  		Connection conn = null;
    			  	Statement st = null;
    			  	ResultSet rs = null;
    			  	
    			  	Class.forName("com.mysql.jdbc.Driver");
    				conn = (Connection)DriverManager.getConnection("jdbc:mysql://localhost:3306/jspproject","root","");
    				st = conn.createStatement();
    				  
    				String qry = "select * from `listening`";
    				rs = st.executeQuery(qry);
    				
    				  String tMail = "";
    				  boolean ok = false;
    				  while(rs.next())
    				  {
    					  tMail = rs.getString("email");
    					  if(tMail.equals(email))
    					  {
    						  ok = true;
    						  break;
    					  }
    				  }
    				  
    				  if(ok){
    					  %>
    					  
    					  	<div id="listening" >
							</div>
						  
    					  <% 
    				  }
    				  
    				  else{
    				  
    				  %>
    				  	<a href="timer.jsp">
    				  		<div class = "box" id="listening" >
						
							</div>
						</a>
    				  <% 
    				  }
    				  %>
				
			<%-- Listening End--%>

			
			<%-- Reading Start --%>
				<%
					email = (String) session.getAttribute("email");	
    		   
					conn = null;
			  		st = null;
			  		rs = null;
    			  	
    			  	Class.forName("com.mysql.jdbc.Driver");
    				conn = (Connection)DriverManager.getConnection("jdbc:mysql://localhost:3306/jspproject","root","");
    				st = conn.createStatement();
    				  
    				qry = "select * from `reading`";
    				rs = st.executeQuery(qry);
    				
    				  tMail = "";
    				  ok = false;
    				  while(rs.next())
    				  {
    					  tMail = rs.getString("email");
    					  if(tMail.equals(email))
    					  {
    						  ok = true;
    						  break;
    					  }
    				  }
    				  
    				  if(ok){
    					  %>
    					  
    					  	<div id="reading" >
							</div>
						  
    					  <% 
    				  }
    				  
    				  else{
    				  
    				  %>
    				  	<a href="timer2.jsp">
    				  		<div class = "box" id="reading" >
						
							</div>
						</a>
    				  <% 
    				  }
    				  %>
				
			<%-- Reading End--%>
			
			
			
			<%-- Writing Start --%>
				<%
					email = (String) session.getAttribute("email");	
    		   
    		  		conn = null;
    			    st = null;
    			  	rs = null;
    			  	
    			  	Class.forName("com.mysql.jdbc.Driver");
    				conn = (Connection)DriverManager.getConnection("jdbc:mysql://localhost:3306/jspproject","root","");
    				st = conn.createStatement();
    				  
    				qry = "select * from `writing`";
    				rs = st.executeQuery(qry);
    				
    				  tMail = "";
    				  ok = false;
    				  while(rs.next())
    				  {
    					  tMail = rs.getString("email");
    					  if(tMail.equals(email))
    					  {
    						  ok = true;
    						  break;
    					  }
    				  }
    				  
    				  if(ok){
    					  %>
    					  
    					  	<div id="writing" >
							</div>
						  
    					  <% 
    				  }
    				  
    				  else{
    				  
    				  %>
    				  	<a href="timer3.jsp">
    				  		<div class = "box" id="writing" >
						
							</div>
						</a>
    				  <% 
    				  }
    				  %>
				
			<%-- Writing End--%>
			
			
			
			<%-- Speaking Start --%>
				<%
					email = (String) session.getAttribute("email");	
    		   
    		  		conn = null;
    			    st = null;
    			  	rs = null;
    			  	
    			  	Class.forName("com.mysql.jdbc.Driver");
    				conn = (Connection)DriverManager.getConnection("jdbc:mysql://localhost:3306/jspproject","root","");
    				st = conn.createStatement();
    				  
    				qry = "select * from `speaking`";
    				rs = st.executeQuery(qry);
    				
    				  tMail = "";
    				  ok = false;
    				  while(rs.next())
    				  {
    					  tMail = rs.getString("email");
    					  if(tMail.equals(email))
    					  {
    						  ok = true;
    						  break;
    					  }
    				  }
    				  
    				  if(ok){
    					  %>
    					  
    					  	<div id="speaking" >
							</div>
						  
    					  <% 
    				  }
    				  
    				  else{
    				  
    				  %>
    				  	<a href="timer4.jsp">
    				  		<div class = "box" id="speaking" >
						
							</div>
						</a>
    				  <% 
    				  }
    				  %>
				
			<%-- Speaking End--%>
			
			
			
		</div>
		
		
		
     
     <form action="timer.jsp" method="post">
	<%
		if(session.getAttribute("email") == null)
		{
			response.sendRedirect("login.jsp");
		}
	
		String name1 = (String) session.getAttribute("uname");
		String email1 = (String) session.getAttribute("email");
		
		//out.println("name: "+ name1);
		//out.print("email: "+ email1);
	
		HttpSession session1 = request.getSession();
		session.setAttribute("uname", name1);
		session.setAttribute("email", email1);
		
				  
	%>
				  
				
	</form>
</body>
</html>