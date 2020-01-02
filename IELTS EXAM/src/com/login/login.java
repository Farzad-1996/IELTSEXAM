package com.login;

import java.io.IOException;
import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import java.sql.*;

/**
 * Servlet implementation class login
 */
@WebServlet("/login")
public class login extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    
    public login() {
        super();
        // TODO Auto-generated constructor stub
    }

	
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		response.getWriter().append("Served at: ").append(request.getContextPath());
	}

	
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		  String name = request.getParameter("uname");
		  String mail = request.getParameter("email");
		  
		  Connection conn = null;
		  Statement st = null;
		  ResultSet rs = null;
		  
		  try{
			  Class.forName("com.mysql.jdbc.Driver");
			  conn = (Connection)DriverManager.getConnection("jdbc:mysql://localhost:3306/jspproject","root","");
			  st = conn.createStatement();
			  
			  String qry = "select * from `tbl_user`";
			  rs = st.executeQuery(qry);
			  
			  String tName = "";
			  String tMail = "";
			  
			  String ok = "no";
			  
			  while(rs.next())
			  {
				  tName = rs.getString(2);
				  tMail = rs.getString(3);
				  if(tName.equals(name) && tMail.equals(mail))
				  {
					  ok = "yes";
					  break;
				  }
			  }
			  
			  if(ok.equals("yes"))
			  {
				  System.out.println("Hello I am here");
				  
				  HttpSession session = request.getSession();
				  session.setAttribute("email", mail);
				  session.setAttribute("uname", name);
				  response.sendRedirect("welcome.jsp");
			  }
			  else
			  {
				  response.sendRedirect("login.jsp");
			  }
		  }catch(Exception ex)
		  {
			  System.out.println("Data not found");
		  }
	
	}

}
