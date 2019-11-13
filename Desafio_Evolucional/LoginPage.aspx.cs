//Inserção das namespaces utilizadas
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
        //Declara e atribui a string de conexão, estabelecida na Web.config
        string connectionString = ConfigurationManager.ConnectionStrings["myCon"].ConnectionString;

        //Função de carregamento da página
        protected void Page_Load(object sender, EventArgs e)
        {
            //Se houver um erro na sessão, insira-o no texto da Label de erro
            if(Session["ErrorMessage"] != null)
            {
                LabelError.Visible = true;
                LabelError.Text = Session["ErrorMessage"].ToString();
                Session["ErrorMessage"] = null;
            }
        }

        //Função do botão Login
        protected void LoginButton_Click(object sender, EventArgs e)
        {
            //Se a função de validação do usuário retornar verdadeiro, redireciona-o à proxima página
            if(ValidateUser(UserName.Text, Password.Text))
            {
                System.Diagnostics.Debug.WriteLine("Success!");
                FormsAuthentication.RedirectFromLoginPage(UserName.Text, true);
            }
            //Se não, retorne à página de login e estabelece um erro de sessão
            else
            {
                Session["ErrorMessage"] = "Credenciais inválidas";
                System.Diagnostics.Debug.WriteLine("Invalid username or password");
                Response.Redirect("LoginPage.aspx", false);
                Context.ApplicationInstance.CompleteRequest();
            }
        }

        //Função de validação do usuário. Recebe duas strings, "nome de usuário" e "senha". 
        //No caso dessa página, essas informações são inseridas nas TextBoxes correspondentes
        private bool ValidateUser(string Username, string Password)
        {
            try
            {
                //Tenta-se estabelecer uma conexão com o banco de dados SQL para uma consulta, utilizando a string de conexão
                using (SqlConnection sql_con = new SqlConnection(connectionString))
                {
                    //Estabelece o comando para procurar no banco de dados pela combinação de usuário e senha inseridos como parâmetros
                    //OBS: Seria adequado criptografar as senhas
                    string sql_query = "SELECT * FROM Users_Table WHERE Username = @username AND Password = @password";

                    using (SqlCommand sql_cmd = new SqlCommand(sql_query, sql_con))
                    {
                        //São atribuídos os valores dos parâmetros da consulta de acordo com as strings inseridas como argumento da função
                        sql_cmd.Parameters.Add(new SqlParameter("@username", SqlDbType.VarChar, 50) { Value = Username });
                        sql_cmd.Parameters.Add(new SqlParameter("@password", SqlDbType.VarChar, 30) { Value = Password });

                        //Abre-se a conexão com o banco de dados
                        sql_con.Open();

                        //Inicializa o objeto DataSet que irá ser preenchido com o valor de retorno da consulta
                        DataSet ds = new DataSet();
                        //Inicializa o objeto SqlDataAdapter que irá retornar o valor de pesquisa de acordo com o comando declarado
                        SqlDataAdapter da = new SqlDataAdapter(sql_cmd);
                        //Preenche o DataSet com o valor de retorno da pesquisa
                        da.Fill(ds);
                        //A conexão com o banco de dados é fechada
                        sql_con.Close();
                        //Elimina-se o DataAdapter.
                        da.Dispose();
                        // OBS: O uso da keywork "using" ao declarar os objetos SQL deixa implícito que eles serão despejados ao final do uso

                        //Se houver ao menos um resultado para os parâmetros definidos, então o login foi realizado com sucesso
                        bool loginSuccessful = ((ds.Tables.Count > 0) && (ds.Tables[0].Rows.Count > 0));
                        //Retorna o resultado da validação
                        return loginSuccessful;
                    }
                }
            }
            //Caso ocorra alguma exceção durante o acesso ao banco de dados SQL, devemos pegá-la
            //OBS: Seria melhor explicitar que tipos de exceções estamos esperando
            catch (Exception ex)
            {
                //Aqui deveria ocorrer algum tratamento de exceção
                System.Diagnostics.Debug.WriteLine("Exception occurred: " + ex.Message);
                //Caso ocorra uma exceção retornamos falso, invalidando o login
                return false;
            }
        }
    }   
}