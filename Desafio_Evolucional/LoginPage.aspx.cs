using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Security;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;

namespace Desafio_Evolucional
{

    public partial class WebForm1 : System.Web.UI.Page
    {
        string connectionString = ConfigurationManager.ConnectionStrings["myCon"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {
            if(Session["ErrorMessage"] != null)
            {
                LabelError.Visible = true;
                LabelError.Text = Session["ErrorMessage"].ToString();
                Session["ErrorMessage"] = null;
            }
        }

        protected void LoginButton_Click(object sender, EventArgs e)
        {
            if(ValidateUser(UserName.Text, Password.Text))
            {
                FormsAuthentication.RedirectFromLoginPage(UserName.Text, true);
            }
            else
            {
                Response.Redirect("LoginPage.aspx", false);
                Context.ApplicationInstance.CompleteRequest();
            }
        }

        private bool ValidateUser(string Username, string Password)
        {
            try
            {
                using (SqlConnection sql_con = new SqlConnection(connectionString))
                {
                    string sql_query = "SELECT * FROM Users_Table WHERE Username = @username AND Password = @password";
                    using (SqlCommand sql_cmd = new SqlCommand(sql_query, sql_con))
                    {
                        sql_cmd.Parameters.Add(new SqlParameter("@username", SqlDbType.VarChar, 50) { Value = Username });
                        sql_cmd.Parameters.Add(new SqlParameter("@password", SqlDbType.VarChar, 30) { Value = Password });

                        sql_con.Open();

                        DataSet ds = new DataSet();
                        SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                        da.Fill(ds);
                        sql_con.Close();

                        bool loginSuccessful = ((ds.Tables.Count > 0) && (ds.Tables[0].Rows.Count > 0));

                        if (loginSuccessful)
                        {
                            System.Diagnostics.Debug.WriteLine("Success!");
                            return true;
                        }
                        else
                        {
                            Session["ErrorMessage"] = "Credenciais inválidas";
                            System.Diagnostics.Debug.WriteLine("Invalid username or password");
                            return false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Exception occurred: " + ex.Message);
                return false;
            }
        }
    }   
}