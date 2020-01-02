<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
    
    <%@ page import="java.sql.*" %>
    <%@ page import="java.util.*" %>
    
    
<html>
	<head>
		<meta charset="utf-8">
		<title>Done Exam</title>
		<link rel="stylesheet" href="style.css">
	</head>
	<body background="bg6.jpg">
	
         <div class="menuBox">
			<h1>You have been successfully submitted your papers.</h1>
			<h1><a href="welcome.jsp">Click here</a> to go welcome page and see your scores.</h1>
		</div>
		
		<%
			
				List<Integer> list1 = new ArrayList<Integer>();
				List<Integer> list2 = new ArrayList<Integer>();
				
				int score = 0;
			
				String test, tot_ans = "";
				for(int i=0;i<10;i++)
				{
					test = "radio";
					String op = Integer.toString(i);
					test += op;
					
					if(request.getParameter(test) != null)
					{
						String myAns = request.getParameter(test);
						tot_ans += myAns;
						
						if(i!=10-1){
							tot_ans += "@";
						}
						
						int ans = Integer.parseInt(myAns);
						list1.add(ans);
						//out.print(myAns);
					}
					
					else
					{
						list1.add(5);
						//out.print("hi");	
					}
					
				}
				
				//out.print("total ans: " + tot_ans);
				
				String name = (String) session.getAttribute("uname");
				String email = (String) session.getAttribute("email");
				  
				//out.println("name: "+ name);
				//out.print("email: "+ email);	
				
				Connection conn = null;
				Statement st = null;
				ResultSet rs = null;
				
				try
				{
					Class.forName("com.mysql.jdbc.Driver");
					conn = (Connection)DriverManager.getConnection("jdbc:mysql://localhost:3306/jspproject","root","");
					st = conn.createStatement();
					
					String qry; 
				    qry = "";
				     
					qry = "select * from `add_qstn4`";
				 	rs = st.executeQuery(qry);
				 	
				 	while(rs.next())
				 	{
				 		int ans = Integer.parseInt(rs.getString("answer"));
				 		list2.add(ans);
				 	}
				 	
				 	
				 	
				 	for (int i = 0; i < list1.size(); i++) {
				 	    
				 		int op_list1 = list1.get(i);
				 		int op_list2 = list2.get(i);
				 		
				 		if(op_list1 == op_list2)
				 			score += 1;
				 		
				 	}
				 	
				}catch(Exception e){
					e.printStackTrace();
				}
				
				//out.println();
				//out.print("your score is: "+score);
				
				Connection conn1 = null;
				Statement st1 = null;
				ResultSet rs1 = null;
				
				
				try{
					  Class.forName("com.mysql.jdbc.Driver");
					  conn1 = (Connection)DriverManager.getConnection("jdbc:mysql://localhost:3306/jspproject","root","");
					  st1 = conn.createStatement();
					  
					  String qry = "INSERT INTO `speaking` (`id`, `name`, `email`, `answer`, `score`) VALUES (NULL, '"+name+"', '"+email+"', '"+tot_ans+"', '"+score+"')"; 
					  st1.executeUpdate(qry); 
					  //out.print("Data is successfully inserted");
					  
				  }catch(Exception ex)
				  {
					  out.print(ex.getMessage()); 
				  }
			%> 
		
	</body>
</html>

