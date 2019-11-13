//Inserção das namespaces utilizadas
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
        //Declara e atribui a string de conexão, estabelecida na Web.config
        string connectionString = ConfigurationManager.ConnectionStrings["myCon"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        //Função do botão de logout
        protected void ButtonLogout_Click(object sender, EventArgs e)
        {
            //Realiza o logout do usuário e o redireciona para a página de login
            FormsAuthentication.SignOut();
            Response.Redirect("LoginPage.aspx", false);
            Context.ApplicationInstance.CompleteRequest();
        }

        //Função do botão "SQL", que insere os dados de alunos aleatórios em uma tabela
        protected void InsertEntriesBtn_Click(object sender, EventArgs e)
        {
            //Declara e inicializa uma lista de objetos da classe Aluno utilizando-se da função "GerarListaAlunos".
            //No caso o argumento da função é um inteiro de valor "1000", o que retornará uma lista com 1000 instâncias
            List<Aluno> alunos = GerarListaAlunos(1000);                       

            try
            {
                //Realiza a tentativa de conexão com o banco de dados através da string de conexão
                using (SqlConnection sql_con = new SqlConnection(connectionString))
                {
                    //Define a string de comando para a consulta.
                    //Aqui, verifica-se se a tabela "Alunos" já existe. Caso positivo, ela é apagada (DROP)
                    //Então, cria-se uma tabela de nome Alunos com as seguintes colunas:
                    //Nome (varchar), Matemática, Português, Geografia, Inglês, 
                    string sql_query = "IF OBJECT_ID('Alunos', 'U') IS NOT NULL " +
                        "DROP TABLE Alunos " + "CREATE TABLE Alunos(Nome varchar(255),";

                    //Utiliza um loop para inserir uma coluna para cada disciplina na lista de disciplinas
                    foreach(string nome_disciplina in listaDisciplinas)
                    {

                        if(nome_disciplina == listaDisciplinas.Last())
                        {
                            sql_query += nome_disciplina + "DECIMAL(4,2))";
                        }
                        else
                        {
                            sql_query += nome_disciplina + "DECIMAL(4,2),";
                        }
                    }

                    /*sql_query += 
                        "Matematica DECIMAL(4,2)," +
                        "Portugues DECIMAL(4,2), Geografia DECIMAL(4,2), Ingles DECIMAL(4,2), Biologia DECIMAL(4,2)," +
                        "Filosofia DECIMAL(4,2), Fisica DECIMAL(4,2), Quimica DECIMAL(4,2))";*/

                    //O comando é utilizado para executar a instrução
                    using (SqlCommand sql_cmd = new SqlCommand(sql_query, sql_con))
                    { 
                        //Abre a conexão com o banco de dados
                        sql_con.Open();
                        //Executa o comando
                        sql_cmd.ExecuteNonQuery();
                        //Fecha a conexão
                        sql_con.Close();
                    }

                    /*Declaramos uma DataTable e a populamos utilizando uma função que retorna
                     uma DataTable preenchida de acordo com uma lista de Alunos*/

                    DataTable tabela = GerarTabelaAlunos(alunos);

                    //Realizamos a cópia em massa dos dados do objeto DataTable para a tabela Alunos no banco de dados SQL
                    using(SqlBulkCopy sql_bulkCopy = new SqlBulkCopy(sql_con))
                    {
                        //Define a tabela de destino como "Alunos"
                        sql_bulkCopy.DestinationTableName = "Alunos";
                        //Abre a conexão
                        sql_con.Open();
                        //Realiza a escrita da tabela
                        sql_bulkCopy.WriteToServer(tabela);
                        //Fecha conexão
                        sql_con.Close();
                    }
                    //Avisa ao usuário através de uma janela de alerta que a tabela foi gerada com sucesso.
                    Response.Write("<script>alert('Tabela SQL gerada com sucesso!');</script>");
                }
            }
            //Caso haja alguma exceção, devemos pegá-la
            //OBS: Seria melhor explicitar que tipos de exceções estamos esperando
            catch (Exception ex)
            {
                //Aqui deveria ocorrer algum tratamento de exceção
                //Avisamos ao usuário que houve um erro durante a geração dos dados ou sua transferência para o banco de dados
                Response.Write("<script>alert('Erro no acesso ao banco de dados SQL!');</script>");
                System.Diagnostics.Debug.WriteLine("Exception occurred: " + ex.Message);
            }
        }

        //Botão que gera o arquivo Excel
        protected void GenerateExcelBtn_Click(object sender, EventArgs e)
        {
            //Declara as variáveis relevantes da classe Excel
            Excel.Application excelApp;
            Excel._Workbook excelWorkbook;
            Excel._Worksheet excelWorksheet;

            try
            {                
                //Inicializamos um novo aplicativo Excel
                excelApp = new Excel.Application();

                /*Se após a inicialização, o valor de excelApp permanecer nulo, então o software Excel não foi instalado
                Uma alternativa seria formatar os dados como CSV (Comma Separated Values) em vez de escrevê-los diretamente
                em um arquivo xlsx, o qual o usuário poderia posteriormente converter para o formato desejado.
                Porém um dos requisitos do desafio exige o arquivo .xlsx*/
                if (excelApp == null)
                {
                    //Avisamos ao usuário a ausência da instalação do Excel e retornamos
                    Response.Write("<script>alert('Excel não instalado!');</script>");
                    return;
                }

                //Torna o aplicativo Excel invisível para que o processamento seja feito no "background"
                excelApp.Visible = false;
                //Inicializamos o Workbook e Worksheet trabalhados
                excelWorkbook = (Excel._Workbook)(excelApp.Workbooks.Add(System.Reflection.Missing.Value));
                excelWorksheet = (Excel._Worksheet)excelWorkbook.ActiveSheet;

                //Se já existir um arquivo xlsx previamente gerado no diretório do projeto, nós o deletamos
                if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "Alunos.xlsx"))
                {
                    System.Diagnostics.Debug.WriteLine("Tabela existente deletada");
                    File.Delete(AppDomain.CurrentDomain.BaseDirectory + "Alunos.xlsx");
                }

                //Declara uma DataTable que irá conter o cabeçalho da tabela
                DataTable tableHeader = new DataTable();

                //Acessamos o banco de dados SQL utilizando a string de conexão
                using (SqlConnection sql_con = new SqlConnection(connectionString))
                {
                    //A consulta deve retornar os nomes das colunas na tabela "Alunos"
                    string sql_query = "SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('Alunos')";
                    using (SqlDataAdapter sql_da = new SqlDataAdapter(sql_query, sql_con))
                    {
                        //Abre-se a conexão
                        sql_con.Open();
                        //Preenche-se a tabela do cabeçalho
                        sql_da.Fill(tableHeader);
                        //Fecha-se a conexão
                        sql_con.Close();
                    }

                    //Adiciona-se os nomes das colunas às respectivas células do cabeçalho no arquivo Excel
                    for (int i = 0; i < tableHeader.Rows.Count; i++)
                    {
                        excelWorksheet.Cells[1, i + 1] = tableHeader.Rows[i][0].ToString();
                    }
                    //É adicionada uma coluna extra, chamada "Média"
                    excelWorksheet.Cells[1, tableHeader.Rows.Count + 1] = "Média";

                    //Cria-se uma DataTable que irá conter o conteúdo da tabela de alunos (nomes e notas)
                    DataTable tableContent = new DataTable();

                    //Seleciona todo o conteúdo da tabela "Alunos"
                    sql_query = "SELECT * from Alunos";
                    using(SqlDataAdapter sql_da = new SqlDataAdapter(sql_query, sql_con))
                    {
                        //Abre a conexão
                        sql_con.Open();
                        //Preenche a DataTable com o conteúdo retornado pela query
                        sql_da.Fill(tableContent);
                        //Fecha conexão
                        sql_con.Close();
                    }

                    //Preenche iterativamente as células da tabela com o conteúdo da DataTable
                    for (int i = 0; i < tableContent.Rows.Count; i++)
                    {   
                        for(int j = 0; j < tableContent.Columns.Count; j++)
                        {
                            //O valor armazenado na DataTable é convertido em string
                            string cellVal = tableContent.Rows[i][j].ToString();
                            Decimal decCellVal;
                            //Tenta-se converter esse valor para Decimal. Caso possível, ele é convertido e inserido na tabela.
                            //Isso quer dizer que o valor é uma nota de alguma disciplina
                            if(Decimal.TryParse(cellVal, out decCellVal)){
                                excelWorksheet.Cells[i + 2, j + 1] = decCellVal;
                            }
                            //Caso não, ele é enviado como string (ou seja, é o nome do aluno)
                            else
                            {
                                excelWorksheet.Cells[i + 2, j + 1] = cellVal;
                            }
                        }
                    }
                    //Estabelece um range de células correspondente à coluna "Média"
                    Excel.Range range = (Excel.Range)excelWorksheet.Range[excelWorksheet.Cells[2, tableHeader.Rows.Count + 1], excelWorksheet.Cells[tableContent.Rows.Count + 1, tableHeader.Rows.Count + 1]];
                    //Insere uma fórmula para todo esse range, que irá calcular a média dos alunos
                    //OBS: Seria melhor elaborar uma função que retorne os IDs das células em vez de deixá-lo "hard-coded" aqui.
                    //Para o propósito atual irá funcionar. Porém caso sejam inseridas/retiradas matérias, haverá um problema.
                    range.Formula = "=AVERAGE(B2:I2)";
                    //Calcula o range que encobre todos as colunas modificadas na tabela
                    range = (Excel.Range)excelWorksheet.Range[excelWorksheet.Cells[1, 1], excelWorksheet.Cells[tableContent.Rows.Count + 1, tableHeader.Rows.Count+1]];
                    //Aplica o "Autofit" à todas as colunas, ajustando automaticamente sua largura.
                    range.EntireColumn.AutoFit();

                    //Envia um alerta ao usuário dizendo que a tabela foi gerada
                    Response.Write("<script>alert('Tabela Excel gerada com sucesso!');</script>");
                    //Salva essa tabela no diretório local com o nome de "Alunos.xlsx"
                    excelWorkbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "Alunos.xlsx", Excel.XlFileFormat.xlWorkbookDefault);
                    //Fecha a tabela
                    excelWorkbook.Close();
                    //Fecha o aplicativo Excel
                    excelApp.Quit();
                }

            }
            //Caso ocorra alguma exceção, enviamos um alerta de erro ao usuário
            catch (Exception ex)
            {
                Response.Write("<script>alert('Um erro ocorreu na geração do arquivo Excel');</script>");
                System.Diagnostics.Debug.WriteLine("Exception ocurred: " + ex.Message);
            }

        }
        //Função do botão de download
        protected void ButtonDownload_Click(object sender, EventArgs e)
        {
            //Caso exista um arquivo de nome Alunos.xlsx no diretório local, então prosseguimos para seu envio
            if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "Alunos.xlsx"))
            {
                try
                {
                    //Estabelece o tipo de conteúdo a ser enviado pela resposta. No caso, uma tabela Excel
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    //Gera o cabeçalho HTTP da resposta
                    Response.AppendHeader("content-disposition", "attachment; filename=Alunos");
                    //Envia o arquivo "Alunos.xlsx" que se encontra no diretório local
                    Response.TransmitFile(AppDomain.CurrentDomain.BaseDirectory + "Alunos.xlsx");
                    //Finaliza a request
                    Context.ApplicationInstance.CompleteRequest();
                }
                //Caso ocorra algum erro na transferência do arquivo, avisamos o usuário
                catch (Exception ex)
                {
                    Response.Write("<script>alert('Erro na transferência do arquivo!');</script>");
                    System.Diagnostics.Debug.WriteLine("Exception ocurred: " + ex.Message);
                }
            }
            //Caso não exista o arquivo, então ele ainda tem de ser gerado. Portanto, enviamos um alerta para o usuário.
            else
            {
                Response.Write("<script>alert('O arquivo Excel ainda não foi gerado!');</script>");
            }
        }

        //Lista com alguns nomes para a geração aleatória de alunos
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

        //Lista com alguns sobrenomes para a geração aleatória de alunos
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

        //Função que gera a tabela dos alunos
        DataTable GerarTabelaAlunos(List<Aluno> listaAlunos)
        {
            //Declara e inicializa a Datatable
            DataTable table = new DataTable();

            //É adicionada a primeira coluna, "Nome";
            table.Columns.Add("Nome", typeof(string));

            //Para cada uma das disciplinas na lista de disciplinas, é adicionada uma coluna com o nome correspondente
            foreach(string nomedisciplina in listaDisciplinas)
            {
                table.Columns.Add(nomedisciplina, typeof(string));
            }

            //Inicializado um array de objetos com o tamanho igual ao número de colunas da tabela
            object[] rowArray = new object[table.Columns.Count];

            /*Para cada um dos alunos na lista de Alunos que serve como argumento da função, atribui-se ao index 0 do array
              o nome do aluno, e aos demais índices as notas das respectivas matérias*/
            foreach(Aluno aluno in listaAlunos)
            {
                //Insere nome do aluno no primeiro índice
                rowArray[0] = aluno.Nome;
                //Insere a nota do aluno nas respectivas disciplinas nos demais índices
                for(int i = 1; i < rowArray.Length; i++)
                {
                    rowArray[i] = aluno.Disciplinas[i - 1].Nota;
                }
                //Declara uma nova linha na DataTable
                DataRow row = table.NewRow();
                //Atribui à essa linha os valores do array (nome e notas)
                row.ItemArray = rowArray;
                //Adiciona essa linha à tabela
                table.Rows.Add(row);
            }
            //Retorna a tabela preenchida
            return table;
        }

        //Função que gera uma lista de alunos com nomes e notas aleatórias
        List<Aluno> GerarListaAlunos(int num_alunos)
        {
            //Declara um novo objeto Random, responsável pela geração de valores pseudo-aleatórios
            Random rnd = new Random();

            //Declara uma nova lista de alunos
            List<Aluno> alunos = new List<Aluno>();

            //Gera um novo aluno e o insere na lista enquanto o número de alunos na lista
            //for menor do que o número de alunos solicitado no argumento da função (num_alunos)
            while (alunos.Count < num_alunos)
            {
                //idx1 e idx2 serão os índices utilizados para escolher um nome e um sobrenome, respectivamente
                int idx1, idx2;
                string nome;

                do
                {
                    //Atribui um valor inteiro aleatório para idx1, indo de 0 até o último índice da lista de nomes
                    idx1 = rnd.Next(0, listaNomes.Count());
                    //Atribui um valor inteiro aleatório para idx2, indo de 0 até o último índice da lista de sobrenomes
                    idx2 = rnd.Next(0, listaSobrenomes.Count());

                    //Concatena nome e sobrenome
                    nome = listaNomes.ElementAt(idx1) + " " + listaSobrenomes.ElementAt(idx2);
                }
                //Esse loop é repetido enquanto o nome do aluno não for válido,
                //ou seja, sempre que o nome gerado já se encontrar na tabela
                while (alunos.Find(a => a.Nome.Equals(nome)) != null);

                //OBS: no caso desse software, o número de nomes e sobrenomes inseridos possibilita um número de combinações
                //suficiente para povoar a tabela com 1000 alunos únicos.
                //Porém, seria prudente colocar alguma condição para impedir um loop infinito no caso do número de combinações for insuficiente

                //Declara um novo aluno e o inicializa, utilizando o nome aleatório gerado
                Aluno aluno = new Aluno(nome);
                //Gera notas aleatórias para esse aluno
                aluno.GerarNotas(listaDisciplinas, rnd);
                //Adiciona esse aluno na lista de alunos
                alunos.Add(aluno);
            }
            //Retorna a lista de alunos
            return alunos;
        }

        //Classe Aluno
        //Contém o nome, a lista de Disciplinas e as funções relevantes 
        class Aluno
        {
            public string Nome { get; set; }

            public List<Disciplina> Disciplinas { get; set; }

            public Aluno(string nome)
            {
                this.Nome = nome;

                Disciplinas = new List<Disciplina>();
            }

            //Função que gera notas aleatórias para o aluno
            public void GerarNotas(IEnumerable<string> disciplinas, Random rnd)
            {
                //Para cada disciplina na lista de disciplinas, é gerada uma nota aleatória
                foreach(string discStr in disciplinas)
                {
                    //Armazena o nome da disciplina
                    string nomeDisciplina = discStr;
                    //Gera um double aleatório entre 0 e 1 e o multiplica por 10, obtendo um valor entre 0 e 10
                    double rndDouble = rnd.NextDouble() * 10;
                    //Arredonda esse double para 2 casas decimais
                    rndDouble = Math.Round(rndDouble, 2);
                    //Converte o double em um valor decimal
                    decimal notaDisciplina = Convert.ToDecimal(rndDouble);
                    //Gera um novo objeto disciplina, contendo o nome e a respectiva nota
                    Disciplina disciplina = new Disciplina(nomeDisciplina, notaDisciplina);
                    //Adiciona esse objeto à lista de disciplinas do aluno
                    this.Disciplinas.Add(disciplina);
                }
            }
        }

        //Lista de disciplinas
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
        
        //Classe disciplina. Contém o nome da disciplina ea respectiva nota
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
    }
}