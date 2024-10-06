using System.Data;
using System.IO;
using System.Text;
using System.Xml;
using ExcelDataReader;

internal class Program
{
    private static void Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        if (args.Length == 0)
        {
            Console.WriteLine("Por favor, forneca o caminho do arquivo.");
            return;
        }

        string caminho_arquivos = args[0];

        if (!File.Exists(caminho_arquivos))
        {
            Console.WriteLine($"O arquivo no caminho {caminho_arquivos} nao foi encontrado");
            return;
        }

        var dados = ler_excel(caminho_arquivos);
        int escolha;
        Console.WriteLine("Arquivo lido com sucesso.");
        Console.ReadKey();
        Console.Clear();

        while (true)
        {
            Console.WriteLine("""
                O que gostaria de listar?
                [1] - Almoços menores do que 55 minutos.
                [2] - Pontos impares.
                [3] - Pontos repetidos / diferença muito pequena.
                [4] - Funcionarios sem horario de almoço.
                [0] - Sair
                """);
            escolha = int.Parse(Console.ReadLine()!);
            if (escolha == 1) { comparar_almocos(dados); }
            if (escolha == 2) { pontos_impares(dados); }
            if (escolha == 3) { horario_igual(dados); }
            if (escolha == 4) { sem_almoco(dados); }
            if (escolha == 0) { break; }
        }

        List<Dictionary<string, object>> ler_excel(string caminho_arquivos)
        {
            var rowsList = new List<Dictionary<string, object>>();

            try
            {
                using (var stream = File.Open(caminho_arquivos, FileMode.Open, FileAccess.Read))
                {
                    // Configura a leitura
                    using (var reader = ExcelReaderFactory.CreateBinaryReader(stream)) // Para XLS use CreateBinaryReader
                    {
                        // Configura as opções para retornar o DataSet
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true // Se o arquivo tiver cabeçalho
                            }
                        });

                        // Percorre as tabelas e dados do arquivo
                        foreach (DataTable table in result.Tables)
                        {
                            foreach (DataRow row in table.Rows)
                            {
                                var rowDict = new Dictionary<string, object>();


                                rowDict["NomeFuncionario"] = row[4];
                                rowDict["CodigoFuncionario"] = row[3];
                                rowDict["PostoFuncionario"] = row[5];
                                rowDict["PrimeiroPonto"] = row[9];
                                rowDict["InicioAlmoco"] = row[10];
                                rowDict["FimAlmoco"] = row[11];
                                rowDict["QuartoPonto"] = row[12];
                                rowDict["QuintoPonto"] = row[13];

                                rowsList.Add(rowDict);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao ler o arquivo: {ex.Message}");
            }

            return rowsList; // Retorna a lista de dicionários
        }


        void comparar_almocos(List<Dictionary<String, object>> dados)
        {
            if (dados != null)
            {
                Console.WriteLine("Almoços menores do que 55 minutos.");
                Console.WriteLine("------------------------");
                foreach (var dado in dados)
                {
                    if (DateTime.TryParse(dado["InicioAlmoco"].ToString(), out DateTime horarioInicio) &&
                        DateTime.TryParse(dado["FimAlmoco"].ToString(), out DateTime horarioFim))
                    {
                        TimeSpan diferenca = horarioFim - horarioInicio;

                        if (diferenca.TotalMinutes < 55)
                        {
                            Console.WriteLine($"Funcionario: {dado["NomeFuncionario"]} \n Matricula: {dado["CodigoFuncionario"]} \n Posto: {dado["PostoFuncionario"]}.");
                            Console.WriteLine("------------------------");
                        }

                    }
                }
                Console.WriteLine("Fim dos erros no almoço");
                Console.WriteLine("------------------------");
            }
        }

        void pontos_impares(List<Dictionary<String, object>> dados)
        {
            if (dados != null) 
            {
                Console.WriteLine("Pontos impares.");
                Console.WriteLine("------------------------");
                foreach(var dado in dados)
                {
                 // Ponto impar com um ponto
                    if (DateTime.TryParse(dado["PrimeiroPonto"].ToString(), out DateTime primeiroPonto) != false && DateTime.TryParse(dado["InicioAlmoco"].ToString(), out DateTime segundoPonto) == false)
                    {
                        Console.WriteLine($"""
                        Apenas uma marcação:
                        Funcionario: {dado["NomeFuncionario"]}
                        Matricula: {dado["CodigoFuncionario"]}
                        Posto: {dado["PostoFuncionario"]}.
                        ------------------------
                        """);
                    }
                    // Ponto impar com 3 marcacoes
                    else if (DateTime.TryParse(dado["FimAlmoco"].ToString(), out DateTime terceiroPonto) != false && DateTime.TryParse(dado["QuartoPonto"].ToString(), out DateTime quartoPonto) == false)
                    {
                        Console.WriteLine($"""
                        Tres marcacoes:
                        Funcionario: {dado["NomeFuncionario"]}
                        Matricula: {dado["CodigoFuncionario"]}
                        Posto: {dado["PostoFuncionario"]}.
                        ------------------------
                        """);
                    }else if(DateTime.TryParse(dado["QuintoPonto"].ToString(), out DateTime quintoPonto) != false)
                    {
                       Console.WriteLine($"""
                       Cinco marcacoes ou mais:
                       Funcionario: {dado["NomeFuncionario"]}
                       Matricula: {dado["CodigoFuncionario"]}
                       Posto: {dado["PostoFuncionario"]}.
                       ------------------------
                      """);
                    }
                }
            }
        }

        void horario_igual(List<Dictionary<String, object>> dados)
        {
            if (dados != null) 
            {
                Console.WriteLine("Pontos com horarios iguais.");
                Console.WriteLine("------------------------");
                foreach (var dado in dados) 
                {
                    if (DateTime.TryParse(dado["PrimeiroPonto"].ToString(), out DateTime primeiroPonto) != false &&
                    DateTime.TryParse(dado["InicioAlmoco"].ToString(), out DateTime segundoPonto) != false &&
                    DateTime.TryParse(dado["FimAlmoco"].ToString(), out DateTime terceiroPonto) != false &&
                    DateTime.TryParse(dado["QuartoPonto"].ToString(), out DateTime quartoPonto) != false)
                    {

                        TimeSpan diferenca;
                        // Dois pontos iguais ou com diferença muito próxima
                        diferenca = segundoPonto - primeiroPonto;
                        if (diferenca.TotalMinutes <= 60)
                        {
                            Console.WriteLine($"""
                        Dois pontos com diferença menor do que 1h:
                        Funcionario: {dado["NomeFuncionario"]}
                        Matricula: {dado["CodigoFuncionario"]}
                        Posto: {dado["PostoFuncionario"]}.
                        ------------------------
                        """);
                        }
                        // Horarios repitidos
                        if (terceiroPonto == quartoPonto)
                        {
                            Console.WriteLine($"""
                        Ponto repetido:
                        Funcionario: {dado["NomeFuncionario"]}
                        Matricula: {dado["CodigoFuncionario"]}
                        Posto: {dado["PostoFuncionario"]}.
                        ------------------------
                        """);
                        }
                    }
                }
            }
        }

        void sem_almoco(List<Dictionary<String, object>> dados)
        {
            if (dados != null)
            {
                Console.WriteLine("Pontos sem almoco");
                Console.WriteLine("------------------------");
                foreach(var dado in dados)
                {
                    if(DateTime.TryParse(dado["PrimeiroPonto"].ToString(), out DateTime primeiroPonto) != false && DateTime.TryParse(dado["InicioAlmoco"].ToString(), out DateTime segundoPonto) != false &&
                       DateTime.TryParse(dado["FimAlmoco"].ToString(), out DateTime terceiroPonto) == false)
                    {
                        TimeSpan diferenca = segundoPonto - primeiroPonto;
                        if(diferenca.TotalMinutes > 405)
                        {
                            Console.WriteLine($"""
                            Sem horario de almoco:
                            Funcionario: {dado["NomeFuncionario"]}
                            Matricula: {dado["CodigoFuncionario"]}
                            Posto: {dado["PostoFuncionario"]}.
                            ------------------------
                            """);
                        }
                    }
                }
            }
        }
    }
}