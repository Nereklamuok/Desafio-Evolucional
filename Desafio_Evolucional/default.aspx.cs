using System;
using System.IO;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Security;
using System.Data;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace Desafio_Evolucional
{
    public partial class WebForm2 : System.Web.UI.Page
    {
        string connectionString = ConfigurationManager.ConnectionStrings["myCon"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void ButtonLogout_Click(object sender, EventArgs e)
        {
            FormsAuthentication.SignOut();
            Response.Redirect("LoginPage.aspx", false);
            Context.ApplicationInstance.CompleteRequest();
        }

        protected void InsertEntriesBtn_Click(object sender, EventArgs e)
        {
            List<Aluno> alunos = GerarListaAlunos(1000);                       

            try
            {
                using (SqlConnection sql_con = new SqlConnection(connectionString))
                {
                    string sql_query = "IF OBJECT_ID('Alunos', 'U') IS NOT NULL " +
                        "DROP TABLE Alunos " + "CREATE TABLE Alunos(Nome varchar(255), Matematica DECIMAL(4,2)," +
                        "Portugues DECIMAL(4,2), Geografia DECIMAL(4,2), Ingles DECIMAL(4,2), Biologia DECIMAL(4,2)," +
                        "Filosofia DECIMAL(4,2), Fisica DECIMAL(4,2), Quimica DECIMAL(4,2))";
                    using (SqlCommand sql_cmd = new SqlCommand(sql_query, sql_con))
                    { 
                        sql_con.Open();
                        sql_cmd.ExecuteNonQuery();
                        sql_con.Close();
                    }

                    DataTable tabela = GerarTabelaAlunos(alunos);

                    using(SqlBulkCopy sql_bulkCopy = new SqlBulkCopy(sql_con))
                    {
                        sql_bulkCopy.DestinationTableName = "Alunos";
                        sql_con.Open();
                        sql_bulkCopy.WriteToServer(tabela);
                        sql_con.Close();
                    }
                    Response.Write("<script>alert('Tabela SQL gerada com sucesso!');</script>");
                }
            }
            catch (Exception ex)
            {
                Response.Write("<script>alert('Erro no acesso ao banco de dados SQL!');</script>");
                System.Diagnostics.Debug.WriteLine("Exception occurred: " + ex.Message);
            }
        }

        protected void GenerateExcelBtn_Click(object sender, EventArgs e)
        {

            Excel.Application excelApp;
            Excel._Workbook excelWorkbook;
            Excel._Worksheet excelWorksheet;

            try
            {                
                excelApp = new Excel.Application();

                if (excelApp == null)
                {
                    Response.Write("<script>alert('Excel não instalado!');</script>");
                    return;
                }

                excelApp.Visible = false;
                excelWorkbook = (Excel._Workbook)(excelApp.Workbooks.Add(System.Reflection.Missing.Value));
                excelWorksheet = (Excel._Worksheet)excelWorkbook.ActiveSheet;

                if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "Alunos.xlsx"))
                {
                    System.Diagnostics.Debug.WriteLine("Tabela existente deletada");
                    File.Delete(AppDomain.CurrentDomain.BaseDirectory + "Alunos.xlsx");
                }

                DataTable tableHeader = new DataTable();

                using (SqlConnection sql_con = new SqlConnection(connectionString))
                {
                    string sql_query = "SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('Alunos')";
                    using (SqlDataAdapter sql_da = new SqlDataAdapter(sql_query, sql_con))
                    {
                        sql_con.Open();
                        sql_da.Fill(tableHeader);
                        sql_con.Close();

                        sql_da.Dispose();
                    }

                    for (int i = 0; i < tableHeader.Rows.Count; i++)
                    {
                        excelWorksheet.Cells[1, i + 1] = tableHeader.Rows[i][0].ToString();
                    }
                    excelWorksheet.Cells[1, tableHeader.Rows.Count + 1] = "Média";

                    DataTable tableContent = new DataTable();

                    sql_query = "SELECT * from Alunos";
                    using(SqlDataAdapter sql_da = new SqlDataAdapter(sql_query, sql_con))
                    {
                        sql_con.Open();
                        sql_da.Fill(tableContent);
                        sql_con.Close();

                        sql_da.Dispose();
                    }

                    for (int i = 0; i < tableContent.Rows.Count; i++)
                    {   
                        for(int j = 0; j < tableContent.Columns.Count; j++)
                        {
                            string cellVal = tableContent.Rows[i][j].ToString();
                            Decimal decCellVal;
                            if(Decimal.TryParse(cellVal, out decCellVal)){
                                excelWorksheet.Cells[i + 2, j + 1] = decCellVal;
                            }
                            else
                            {
                                excelWorksheet.Cells[i + 2, j + 1] = cellVal;
                            }
                        }
                    }

                    Excel.Range range = (Excel.Range)excelWorksheet.Range[excelWorksheet.Cells[2, tableHeader.Rows.Count + 1], excelWorksheet.Cells[tableContent.Rows.Count + 1, tableHeader.Rows.Count + 1]];
                    range.Formula = "=AVERAGE(B2:I2)";
                    range = (Excel.Range)excelWorksheet.Range[excelWorksheet.Cells[1, 1], excelWorksheet.Cells[tableContent.Rows.Count + 1, tableHeader.Rows.Count+1]];
                    range.EntireColumn.AutoFit();

                    Response.Write("<script>alert('Tabela Excel gerada com sucesso!');</script>");
                    excelWorkbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "Alunos.xlsx", Excel.XlFileFormat.xlWorkbookDefault);
                    excelWorkbook.Close();
                    excelApp.Quit();
                }

            }
            catch (Exception ex)
            {
                Response.Write("<script>alert('Um erro ocorreu na geração do arquivo Excel');</script>");
                System.Diagnostics.Debug.WriteLine("Exception ocurred: " + ex.Message);
            }

        }

        public IEnumerable<string> listaNomes = new List<string>()
        {
            "José",
            "Maria",
            "João",
            "Ana",
            "Antônio",
            "Francisca",
            "Francisco",
            "Carlos",
            "Adriana",
            "Paulo",
            "Márcia",
            "Lucas",
            "Fernanda",
            "Luiz",
            "Patrícia",
            "Marcos",
            "Aline",
            "Luís",
            "Sandra",
            "Gabriel",
            "Camila",
            "Rafael",
            "Amanda",
            "Daniel",
            "Bruna",
            "Marcelo",
            "Jéssica",
            "Bruno",
            "Letícia",
            "Eduardo",
            "Julia",
            "Felipe",
            "Luciana",
            "Raimundo",
            "Vanessa",
            "Rodrigo",
            "Mariana",
            "Cleyton",
            "Marta",
            "Joyce"
        };

        public IEnumerable<string> listaSobrenomes = new List<string>()
        {
            "Silva",
            "Souza",
            "Costa",
            "Santos",
            "Oliveira",
            "Pereira",
            "Rodrigues",
            "Almeida",
            "Nascimento",
            "Lima",
            "Araújo",
            "Fernandes",
            "Carvalho",
            "Gomes",
            "Martins",
            "Rocha",
            "Ribeiro",
            "Alves",
            "Monteiro",
            "Mendes",
            "Barros",
            "Freitas",
            "Barbosa",
            "Moura",
            "Cavalcanti",
            "Dias",
            "Castro",
            "Campos",
            "Cardoso",
            "Moraes",
            "Navarro",
            "Lopes",
            "Corrêa",
            "Salgado"
        };

        DataTable GerarTabelaAlunos(List<Aluno> listaAlunos)
        {
            DataTable table = new DataTable();

            table.Columns.Add("Nome", typeof(string));

            foreach(string nomedisciplina in listaDisciplinas)
            {
                table.Columns.Add(nomedisciplina, typeof(string));
            }

            object[] rowArray = new object[table.Columns.Count];

            foreach(Aluno aluno in listaAlunos)
            {
                rowArray[0] = aluno.Nome;
                for(int i = 1; i < rowArray.Length; i++)
                {
                    rowArray[i] = aluno.Disciplinas[i - 1].Nota;
                }

                DataRow row = table.NewRow();
                row.ItemArray = rowArray;
                table.Rows.Add(row);
            }

            return table;
        }

        List<Aluno> GerarListaAlunos(int num_alunos)
        {
            Random rnd = new Random();

            List<Aluno> alunos = new List<Aluno>();

            while (alunos.Count < num_alunos)
            {
                int idx1, idx2;
                string nome;

                do
                {
                    idx1 = rnd.Next(0, listaNomes.Count());
                    idx2 = rnd.Next(0, listaSobrenomes.Count());

                    nome = listaNomes.ElementAt(idx1) + " " + listaSobrenomes.ElementAt(idx2);
                }
                while (alunos.Find(a => a.Nome.Equals(nome)) != null);

                Aluno aluno = new Aluno(nome);
                aluno.GerarNotas(listaDisciplinas, rnd);
                alunos.Add(aluno);
            }
            return alunos;
        }

        class Aluno
        {
            public string Nome { get; set; }

            public List<Disciplina> Disciplinas { get; set; }

            public Aluno(string nome)
            {
                this.Nome = nome;

                Disciplinas = new List<Disciplina>();
            }

            public void GerarNotas(IEnumerable<string> disciplinas, Random rnd)
            {
                foreach(string discStr in disciplinas)
                {
                    string nomeDisciplina = discStr;
                    double rndDouble = rnd.NextDouble() * 10;
                    double teste = Math.Round(rndDouble, 2);
                    decimal notaDisciplina = Convert.ToDecimal(Math.Round(teste, 2));
                    Disciplina disciplina = new Disciplina(nomeDisciplina, notaDisciplina);
                    this.Disciplinas.Add(disciplina);
                }
            }
        }

        public IEnumerable<string> listaDisciplinas = new List<string>()
        {
            "Matemática",
            "Português",
            "Geografia",
            "Inglês",
            "Biologia",
            "Filosofia",
            "Física",
            "Química"
        };
        
        class Disciplina
        {
            public string Nome_Disciplina;
            public decimal Nota;

            public Disciplina(string disciplina, decimal nota)
            {
                this.Nome_Disciplina = disciplina;
                this.Nota = nota;
            }
        }

        protected void ButtonDownload_Click(object sender, EventArgs e)
        {
            if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "Alunos.xlsx"))
            {
                try
                {
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AppendHeader("content-disposition", "attachment; filename=Alunos");
                    Response.TransmitFile(AppDomain.CurrentDomain.BaseDirectory + "Alunos.xlsx");
                    Response.End();
                }
                catch (Exception ex)
                {
                    Response.Write("<script>alert('Erro na transferência do arquivo!');</script>");
                    System.Diagnostics.Debug.WriteLine("Exception ocurred: " + ex.Message);
                }
            }
            else
            {
                Response.Write("<script>alert('O arquivo Excel ainda não foi gerado!');</script>");
            }
        }
    }
}