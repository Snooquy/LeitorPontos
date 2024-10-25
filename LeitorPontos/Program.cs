using System.Collections.Generic;
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

        // VARIAVEIS NECESSARIAS
        var dados = ler_excel(caminho_arquivos);
        int escolha;
        var funcionarios_listados = new List<Dictionary<string, object>>();


        Console.WriteLine("Arquivo lido com sucesso.");
        Console.ReadKey();
        Console.Clear();

        while (true)
        {
            Console.WriteLine("""
                O que gostaria de listar?
                [1] - Localiza - DF
                [2] - Localiza - ES
                [3] - Localiza - GO
                [4] - Localiza - MG
                [5] - Localiza - MT
                [6] - Localiza - MTS
                [7] - Localiza - RJ
                [8] - Localiza - SP
                [9] - Localiza - TO
                [10] - MOVIDA - MOVIMENTAÇÃO BA
                [11] - MOVIDA - MOVIMENTACAO SP
                [0] - Sair
                """);
            escolha = int.Parse(Console.ReadLine()!);
            if (escolha == 0) { break; }
            listar(escolha, dados);
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


        void comparar_almocos(List<Dictionary<String, object>> dados,string posto)
        {
            if (dados != null)
            {
                Console.WriteLine("Almoços menores do que 55 minutos.");
                Console.WriteLine("------------------------");
                foreach (var dado in dados)
                {
                    if (dado["PostoFuncionario"].ToString() == posto)
                    {
                        // VERIFICA SE O FUNCIONARIO JA FOI LISTADO ANTES
                        if (!funcionarios_listados.Any(f => f["CodigoFuncionario"].ToString() == dado["CodigoFuncionario"].ToString()))
                        {
                            if (DateTime.TryParse(dado["InicioAlmoco"].ToString(), out DateTime horarioInicio) != false &&
                            DateTime.TryParse(dado["FimAlmoco"].ToString(), out DateTime horarioFim) != false)
                            {
                                TimeSpan diferenca = horarioFim - horarioInicio;

                                if (diferenca.TotalMinutes < 55)
                                {
                                    Console.WriteLine($"Funcionario: {dado["NomeFuncionario"]} \n Matricula: {dado["CodigoFuncionario"]} \n Posto: {dado["PostoFuncionario"]}.");
                                    Console.WriteLine("------------------------");
                                    funcionarios_listados.Add(dado);
                                }

                            }
                        }
                    }
                }
                Console.WriteLine("Fim dos erros no almoço");
                Console.WriteLine("------------------------");
            }
        }

        void pontos_impares(List<Dictionary<String, object>> dados, string posto)
        {
            if (dados != null) 
            {
                Console.WriteLine("Pontos impares.");
                Console.WriteLine("------------------------");
                foreach (var dado in dados)
                {
                    if (posto == dado["PostoFuncionario"].ToString())
                    {
                        if (!funcionarios_listados.Any(f => f["CodigoFuncionario"].ToString() == dado["CodigoFuncionario"].ToString()))
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
                                funcionarios_listados.Add(dado);
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
                                funcionarios_listados.Add(dado);
                            }
                            else if (DateTime.TryParse(dado["QuintoPonto"].ToString(), out DateTime quintoPonto) != false)
                            {
                                Console.WriteLine($"""
                       Cinco marcacoes ou mais:
                       Funcionario: {dado["NomeFuncionario"]}
                       Matricula: {dado["CodigoFuncionario"]}
                       Posto: {dado["PostoFuncionario"]}.
                       ------------------------
                       """);
                                funcionarios_listados.Add(dado);
                            }
                        }
                    }
                }
            }
        }

        void horario_igual(List<Dictionary<String, object>> dados, string posto)
        {
            if (dados != null) 
            {
                Console.WriteLine("Pontos com horarios iguais.");
                Console.WriteLine("------------------------");
                foreach (var dado in dados)
                {
                    if (posto == dado["PostoFuncionario"].ToString())
                    {
                        if (!funcionarios_listados.Any(f => f["CodigoFuncionario"].ToString() == dado["CodigoFuncionario"].ToString()))
                        {
                            if (DateTime.TryParse(dado["PrimeiroPonto"].ToString(), out DateTime primeiroPonto) != false &&
                                DateTime.TryParse(dado["InicioAlmoco"].ToString(), out DateTime segundoPonto) != false &&
                                DateTime.TryParse(dado["FimAlmoco"].ToString(), out DateTime terceiroPonto) != false &&
                                DateTime.TryParse(dado["QuartoPonto"].ToString(), out DateTime quartoPonto) != false)
                            {

                                TimeSpan diferenca;
                                TimeSpan diferenca2;
                                // Dois pontos iguais ou com diferença muito próxima
                                diferenca = segundoPonto - primeiroPonto;
                                diferenca2 = quartoPonto - terceiroPonto;
                                if (diferenca.TotalMinutes <= 60 && diferenca.TotalMinutes >= 0)
                                {
                                    Console.WriteLine($"""
                            Primeiro ponto muito proximo do segundo:
                            Funcionario: {dado["NomeFuncionario"]}
                            Matricula: {dado["CodigoFuncionario"]}
                            Posto: {dado["PostoFuncionario"]}.
                            ------------------------
                            """);
                                    funcionarios_listados.Add(dado);
                                }

                                else if (diferenca2.TotalMinutes <= 60 && diferenca2.TotalMinutes >= 0)
                                {
                                    Console.WriteLine($"""
                            Terceiro ponto muito proximo do quarto:
                            Funcionario: {dado["NomeFuncionario"]}
                            Matricula: {dado["CodigoFuncionario"]}
                            Posto: {dado["PostoFuncionario"]}.
                            ------------------------
                            """);
                                    funcionarios_listados.Add(dado);
                                }
                            }
                        }
                    }
                }
            }
        }

        void sem_almoco(List<Dictionary<String, object>> dados, string posto)
        {
            if (dados != null)
            {
                Console.WriteLine("Pontos sem almoco");
                Console.WriteLine("------------------------");
                foreach (var dado in dados)
                {
                    if (dado["PostoFuncionario"].ToString() == posto)
                    {
                        if (!funcionarios_listados.Any(f => f["CodigoFuncionario"].ToString() == dado["CodigoFuncionario"].ToString()))
                        {
                            if (DateTime.TryParse(dado["PrimeiroPonto"].ToString(), out DateTime primeiroPonto) != false && DateTime.TryParse(dado["InicioAlmoco"].ToString(), out DateTime segundoPonto) != false &&
                           DateTime.TryParse(dado["FimAlmoco"].ToString(), out DateTime terceiroPonto) == false)
                            {
                                TimeSpan diferenca = segundoPonto - primeiroPonto;
                                if (diferenca.TotalMinutes > 405)
                                {
                                    Console.WriteLine($"""
                            Sem horario de almoco:
                            Funcionario: {dado["NomeFuncionario"]}
                            Matricula: {dado["CodigoFuncionario"]}
                            Posto: {dado["PostoFuncionario"]}.
                            ------------------------
                            """);
                                    funcionarios_listados.Add(dado);
                                }
                            }
                        }
                    }
                }
            }
        }

        void listar(int escolha, List<Dictionary<String, object>> dados)
        {
            Dictionary<int, string> postos = new Dictionary<int, string>{
                { 1, "Localiza - DF" },
                { 2, "Localiza - ES" },
                { 3, "Localiza - GO" },
                { 4, "Localiza - MG" },
                { 5, "Localiza - MT" },
                { 6, "Localiza - MTS" },
                { 7, "Localiza - RJ" },
                { 8, "Localiza - SP" },
                { 9, "Localiza - TO" },
                { 10, "MOVIDA - MOVIMENTAÇÃO BA" },
                { 11, "MOVIDA - MOVIMENTACAO SP" }
            };
            string escolha_posto = postos[escolha].ToUpper();
            comparar_almocos(dados, escolha_posto);
            pontos_impares(dados, escolha_posto);
            horario_igual(dados, escolha_posto);
            sem_almoco(dados, escolha_posto);
        }
    }
}