using ExcelDataReader;
using OfficeOpenXml;
using System.ComponentModel;
using System.Data;
using System.Text;

internal partial class Program
{
    private static void Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        if (args.Length == 0)
        {
            Console.WriteLine("Por favor, forneça o caminho do arquivo.");
            return;
        }

        string caminho_arquivos = args[0];
        string caminho_saida = @"C:\Users\isnoo\OneDrive\Documentos\Erros\ErrosPontos.xlsx";

        if (!File.Exists(caminho_arquivos))
        {
            Console.WriteLine($"O arquivo no caminho {caminho_arquivos} não foi encontrado.");
            return;
        }

        // Carrega os dados e realiza as análises
        var dados = LerExcel(caminho_arquivos);
        Console.WriteLine($"Total de registros lidos: {dados.Count}");

        if (dados.Count == 0)
        {
            Console.WriteLine("Nenhum dado foi lido do arquivo Excel.");
            return;
        }

        var funcionariosListados = new List<Dictionary<string, object>>();
        AnalisarErros(dados, funcionariosListados);

        if (funcionariosListados.Count == 0)
        {
            Console.WriteLine("Nenhum erro encontrado para gerar a planilha.");
            return;
        }

        // Gera a nova planilha com os erros
        GerarPlanilhaErros(funcionariosListados, caminho_saida);
        Console.WriteLine("Planilha de erros gerada com sucesso.");
        Console.ReadKey(); // Mover para o final para ver a última mensagem
    }

    private static List<Dictionary<string, object>> LerExcel(string caminhoArquivos)
    {
        var rowsList = new List<Dictionary<string, object>>();

        try
        {
            using (var stream = File.Open(caminhoArquivos, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    foreach (DataTable table in result.Tables)
                    {
                        foreach (DataRow row in table.Rows)
                        {
                            var rowDict = new Dictionary<string, object>
                            {
                                ["NomeFuncionario"] = row[4],
                                ["CodigoFuncionario"] = row[3],
                                ["PostoFuncionario"] = row[5],
                                ["Escala"] = row[6],
                                ["PrimeiroPonto"] = row[9],
                                ["InicioAlmoco"] = row[10],
                                ["FimAlmoco"] = row[11],
                                ["QuartoPonto"] = row[12],
                                ["QuintoPonto"] = row[13]
                            };

                            rowsList.Add(rowDict);
                        }
                    }
                }
            }
            Console.WriteLine("Arquivo lido com sucesso!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao ler o arquivo: {ex.Message}");
        }

        return rowsList;
    }

    private static List<Dictionary<string, object>> AnalisarErros(List<Dictionary<string, object>> dados, List<Dictionary<string, object>> funcionariosListados)
    {
        foreach (var dado in dados)
        {
            string? erro = null;

            // Verificar almoços menores de 55 minutos e maiores de 126 minutos
            if (DateTime.TryParse(dado["InicioAlmoco"].ToString(), out DateTime inicioAlmoco) &&
                DateTime.TryParse(dado["FimAlmoco"].ToString(), out DateTime fimAlmoco))
            {
                TimeSpan diferenca = fimAlmoco - inicioAlmoco;
                if ((diferenca.TotalMinutes < 55 && diferenca.TotalMinutes >= 0) || diferenca.TotalMinutes > 126)
                {
                    erro = "Duração do almoço fora do intervalo.";
                }
            }

            // Verificar pontos ímpares (1, 3, ou 5 marcações)
            int marcacoes = new[] { dado["PrimeiroPonto"], dado["InicioAlmoco"], dado["FimAlmoco"], dado["QuartoPonto"], dado["QuintoPonto"] }
                .Count(x => DateTime.TryParse(x?.ToString(), out _));

            if (marcacoes == 1 || marcacoes == 3 || marcacoes >= 5)
            {
                erro = "Número de marcações ímpares.";
            }

            // Verificar pontos com horários muito próximos (menos de 50 minutos de intervalo)
            if (DateTime.TryParse(dado["PrimeiroPonto"].ToString(), out DateTime primeiroPonto) &&
                DateTime.TryParse(dado["InicioAlmoco"].ToString(), out DateTime segundoPonto) &&
                DateTime.TryParse(dado["FimAlmoco"].ToString(), out DateTime terceiroPonto) &&
                DateTime.TryParse(dado["QuartoPonto"].ToString(), out DateTime quartoPonto))
            {
                if ((segundoPonto - primeiroPonto).TotalMinutes <= 50 || (quartoPonto - terceiroPonto).TotalMinutes <= 50)
                {
                    erro = "Intervalo de pontos muito curto.";
                }
            }

            // Verificar falta de horário de almoço (intervalo maior que 405 minutos sem almoço)
            if (DateTime.TryParse(dado["PrimeiroPonto"].ToString(), out primeiroPonto) &&
                DateTime.TryParse(dado["InicioAlmoco"].ToString(), out segundoPonto) &&
                !DateTime.TryParse(dado["FimAlmoco"].ToString(), out _))
            {
                if ((segundoPonto - primeiroPonto).TotalMinutes > 405)
                {
                    erro = "Falta de horário de almoço.";
                }
            }

            if (erro != null)
            {
                var funcionarioComErro = new Dictionary<string, object>(dado)
                {
                    ["Erro"] = erro
                };
                funcionariosListados.Add(funcionarioComErro);
                Console.WriteLine($"Erro encontrado para {dado["NomeFuncionario"]}: {erro}");
            }
        }
        return funcionariosListados;
    }

    private static void GerarPlanilhaErros(List<Dictionary<string, object>> funcionariosListados, string caminhoSaida)
    {
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Erros de Pontos");

            // Cabeçalho
            worksheet.Cells[1, 1].Value = "Nome do Funcionário";
            worksheet.Cells[1, 2].Value = "Código do Funcionário";
            worksheet.Cells[1, 3].Value = "Posto";
            worksheet.Cells[1, 4].Value = "Escala";
            worksheet.Cells[1, 5].Value = "Primeiro Ponto";
            worksheet.Cells[1, 6].Value = "Início do Almoço";
            worksheet.Cells[1, 7].Value = "Fim do Almoço";
            worksheet.Cells[1, 8].Value = "Quarto Ponto";
            worksheet.Cells[1, 9].Value = "Quinto Ponto";
            worksheet.Cells[1, 10].Value = "Erro"; // Nova coluna para erros

            int row = 2;

            foreach (var funcionario in funcionariosListados)
            {
                worksheet.Cells[row, 1].Value = funcionario["NomeFuncionario"];
                worksheet.Cells[row, 2].Value = funcionario["CodigoFuncionario"];
                worksheet.Cells[row, 3].Value = funcionario["PostoFuncionario"];
                worksheet.Cells[row, 4].Value = funcionario["Escala"];
                worksheet.Cells[row, 5].Value = funcionario["PrimeiroPonto"];
                worksheet.Cells[row, 6].Value = funcionario["InicioAlmoco"];
                worksheet.Cells[row, 7].Value = funcionario["FimAlmoco"];
                worksheet.Cells[row, 8].Value = funcionario["QuartoPonto"];
                worksheet.Cells[row, 9].Value = funcionario["QuintoPonto"];
                worksheet.Cells[row, 10].Value = funcionario["Erro"]; // Preencher a coluna de erro
                row++;
            }

            package.SaveAs(new FileInfo(caminhoSaida));
        }
    }
}