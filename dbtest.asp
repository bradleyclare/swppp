<html><body>

<asp:button id="btnTest" runat="server" onclick="btnTest_Click" text="Test Database Connection" />
<%



protected void btnTest_Click(object sender, EventArgs e)
{
	SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["NorthwindConnectionString"].ConnectionString);
	connection.Open();
	if ((connection.State & ConnectionState.Open) > 0)
	{
		Response.Write("Connection OK!");
		connection.Close();
	}
	else
	{
		Response.Write("Connection no good!");
	}
}

%>
</body></html>