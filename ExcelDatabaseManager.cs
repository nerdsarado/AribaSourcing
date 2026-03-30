// ExcelDatabaseManager.cs
using OfficeOpenXml;
using System.IO;

namespace CotacoesAriba
{
    public static class ExcelHelper
    {
        private static bool _licenseConfigured = false;

        public static void ConfigureLicense()
        {
            if (!_licenseConfigured)
            {
                try
                {
                    // Para uso não comercial
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    _licenseConfigured = true;
                    Console.WriteLine("✅ Licenca EPPlus configurada com sucesso");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ Erro ao configurar licenca EPPlus: {ex.Message}");
                    throw;
                }
            }
        }
    }

    public class ExcelDatabaseManager
    {
        private readonly string _caminhoPlanilha;
        private ExcelPackage _excelPackage;
        private FileInfo _arquivoInfo;

        public ExcelDatabaseManager(string caminhoPlanilha)
        {
            try
            {
                Console.WriteLine($"Inicializando ExcelDatabaseManager...");

                // Configurar licenca primeiro
                ExcelHelper.ConfigureLicense();

                // Verificar se o arquivo existe
                _arquivoInfo = new FileInfo(caminhoPlanilha);
                if (!_arquivoInfo.Exists)
                {
                    throw new FileNotFoundException($"Planilha nao encontrada: {caminhoPlanilha}");
                }

                Console.WriteLine($"Carregando planilha: {_arquivoInfo.Name}");

                // Abrir o arquivo com FileStream para garantir acesso
                using (var stream = new FileStream(caminhoPlanilha, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    _excelPackage = new ExcelPackage(stream);
                }

                _caminhoPlanilha = caminhoPlanilha;

                Console.WriteLine($"✅ Planilha carregada com sucesso!");
                Console.WriteLine($"   📊 Arquivo: {_arquivoInfo.Name}");
                Console.WriteLine($"   📁 Tamanho: {_arquivoInfo.Length / 1024} KB");

                // Verificar se tem planilhas
                if (_excelPackage.Workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine($"⚠️ Atencao: Planilha nao contem abas/worksheets");
                }
                else
                {
                    Console.WriteLine($"   📋 Total de abas: {_excelPackage.Workbook.Worksheets.Count}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ ERRO ao carregar planilha: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"   📌 Inner Exception: {ex.InnerException.Message}");
                }
                throw;
            }
        }

        public bool CotacaoJaExiste(string numeroCotacao)
        {
            try
            {
                Console.WriteLine($"\n🔍 VERIFICANDO COTAÇÃO NO BANCO DE DADOS...");
                Console.WriteLine($"   📄 Planilha: {Path.GetFileName(_caminhoPlanilha)}");
                Console.WriteLine($"   🔍 Número: {numeroCotacao}");

                // Verificar se a planilha existe
                if (!File.Exists(_caminhoPlanilha))
                {
                    Console.WriteLine($"   ⚠ Planilha não encontrada no caminho: {_caminhoPlanilha}");
                    Console.WriteLine($"   📌 Verifique se o diretório de rede está acessível");
                    return false; // Se não existe, considera que não foi processada
                }

                using (var package = new ExcelPackage(new FileInfo(_caminhoPlanilha)))
                {
                    var workbook = package.Workbook;

                    // Listar todas as planilhas para debug
                    Console.WriteLine($"   📊 Planilhas encontradas:");
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        Console.WriteLine($"      - {worksheet.Name}");
                    }

                    // Procurar em TODAS as planilhas
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        Console.WriteLine($"\n   🔎 Buscando na planilha: {worksheet.Name}");

                        var dimension = worksheet.Dimension;
                        if (dimension == null)
                        {
                            Console.WriteLine($"      ⏭️ Planilha vazia, pulando...");
                            continue;
                        }

                        int rowCount = dimension.Rows;
                        int colCount = dimension.Columns;

                        Console.WriteLine($"      📈 Dimensões: {rowCount} linhas x {colCount} colunas");

                        // Procurar o número da cotação em TODAS as células
                        bool encontrou = false;
                        int linhaEncontrada = -1;
                        int colunaEncontrada = -1;

                        for (int row = 1; row <= rowCount; row++)
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
                                var cellValue = worksheet.Cells[row, col].Text?.Trim() ?? "";

                                // Verificar se a célula contém o número da cotação
                                if (cellValue.Contains(numeroCotacao))
                                {
                                    encontrou = true;
                                    linhaEncontrada = row;
                                    colunaEncontrada = col;

                                    // Verificar células ao redor para contexto
                                    string contexto = ObterContextoCelula(worksheet, row, col);
                                    Console.WriteLine($"      ✅ ENCONTRADO na linha {row}, coluna {col}");
                                    Console.WriteLine($"      📋 Contexto: {contexto}");

                                    return true;
                                }
                            }

                            // A cada 100 linhas, mostrar progresso
                            if (row % 100 == 0)
                            {
                                Console.WriteLine($"      🔍 Verificadas {row} linhas...");
                            }
                        }

                        if (!encontrou)
                        {
                            Console.WriteLine($"      ❌ Não encontrado na planilha {worksheet.Name}");
                        }
                    }

                    Console.WriteLine($"   ✅ Cotação NÃO encontrada na planilha - pode ser processada");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ❌ ERRO ao acessar planilha: {ex.Message}");
                Console.WriteLine($"   📌 Tipo: {ex.GetType().Name}");

                // Em caso de erro, considera que não foi processada para não bloquear
                return false;
            }
        }

        private string ObterContextoCelula(ExcelWorksheet worksheet, int row, int col)
        {
            try
            {
                // Pegar algumas células ao redor para contexto
                var contexto = new List<string>();

                // Célula anterior na mesma linha
                if (col > 1)
                {
                    var anterior = worksheet.Cells[row, col - 1].Text?.Trim();
                    if (!string.IsNullOrEmpty(anterior))
                        contexto.Add($"← {anterior}");
                }

                // Célula atual
                var atual = worksheet.Cells[row, col].Text?.Trim();
                contexto.Add($"[{atual}]");

                // Célula seguinte na mesma linha
                var seguinte = worksheet.Cells[row, col + 1].Text?.Trim();
                if (!string.IsNullOrEmpty(seguinte))
                    contexto.Add($"→ {seguinte}");

                return string.Join(" ", contexto);
            }
            catch
            {
                return "[contexto não disponível]";
            }
        }

        public bool AdicionarCotacaoNaPlanilha(
     string numeroCotacao,
     string portal,
     string cliente,
     string dataVencimento,
     string horarioVencimento,
     string produtos,
     string empresaResposta,
     string vendedor = "")
        {
            try
            {
                Console.WriteLine($"\n📝 ADICIONANDO COTAÇÃO NA PLANILHA...");
                Console.WriteLine($"   🔢 Número: {numeroCotacao}");
                Console.WriteLine($"   🏢 Empresa: {empresaResposta}");

                if (!File.Exists(_caminhoPlanilha))
                {
                    Console.WriteLine($"   ❌ Planilha não encontrada!");
                    return false;
                }

                using (var package = new ExcelPackage(new FileInfo(_caminhoPlanilha)))
                {
                    // Tentar encontrar a planilha correta
                    var worksheet = package.Workbook.Worksheets["COTAÇÕES 2025"]
                                  ?? package.Workbook.Worksheets[0];

                    if (worksheet == null)
                    {
                        Console.WriteLine($"   ❌ Nenhuma planilha encontrada");
                        return false;
                    }

                    Console.WriteLine($"   📋 Planilha: {worksheet.Name}");

                    // Encontrar a próxima linha vazia
                    int ultimaLinha = worksheet.Dimension?.End.Row ?? 1;
                    ultimaLinha = ultimaLinha > 0 ? ultimaLinha : 1;

                    // Começar da linha 2 se for cabeçalho
                    if (ultimaLinha == 1)
                    {
                        ultimaLinha = 2;
                    }
                    else
                    {
                        // Encontrar a primeira linha vazia após a última linha com dados
                        for (int row = ultimaLinha; row >= 1; row--)
                        {
                            var cellValue = worksheet.Cells[row, 1].Text?.Trim();
                            if (!string.IsNullOrEmpty(cellValue))
                            {
                                ultimaLinha = row + 1;
                                break;
                            }
                        }
                    }

                    Console.WriteLine($"   📍 Adicionando na linha: {ultimaLinha}");

                    // Mapear colunas
                    Dictionary<string, int> colunas = MapearColunasPlanilha(worksheet);

                    // DEBUG: Mostrar mapeamento
                    Console.WriteLine($"   🗺️  Mapeamento de colunas encontrado:");
                    foreach (var kvp in colunas.OrderBy(c => c.Value))
                    {
                        Console.WriteLine($"      Col{kvp.Value}: '{kvp.Key}'");
                    }

                    // Preencher os dados baseado nos cabeçalhos reais
                    foreach (var kvp in colunas)
                    {
                        string cabecalho = kvp.Key;
                        int coluna = kvp.Value;

                        // Determinar qual valor colocar baseado no cabeçalho
                        string valor = "";

                        if (cabecalho.Contains("CLCL") || cabecalho.Contains("COTA") || cabecalho.Contains("NÚMERO") || cabecalho.Contains("6000"))
                        {
                            valor = numeroCotacao;
                            Console.WriteLine($"      ✅ Número na coluna {coluna} ({cabecalho})");
                        }
                        else if (cabecalho.Contains("PORTAL"))
                        {
                            valor = portal;
                        }
                        else if (cabecalho.Contains("CLIENTE"))
                        {
                            valor = cliente;
                        }
                        else if (cabecalho.Contains("DATA") && cabecalho.Contains("VENCIMENTO"))
                        {
                            valor = dataVencimento;
                        }
                        else if (cabecalho.Contains("HORÁRIO") && cabecalho.Contains("VENCIMENTO"))
                        {
                            valor = horarioVencimento;
                        }
                        else if (cabecalho.Contains("PRODUTO"))
                        {
                            valor = produtos;
                        }
                        else if (cabecalho.Contains("DATA") && cabecalho.Contains("ENTREGA"))
                        {
                            valor = DateTime.Now.ToString("dd/MM/yyyy");
                        }
                        else if (cabecalho.Contains("HORÁRIO") && cabecalho.Contains("ENTREGA"))
                        {
                            valor = DateTime.Now.ToString("HH:mm");
                        }
                        else if (cabecalho.Contains("POR QUAL") || cabecalho.Contains("EMPRESA"))
                        {
                            valor = empresaResposta;
                        }
                        else if (cabecalho.Contains("VENDEDOR"))
                        {
                            valor = vendedor;
                        }

                        if (!string.IsNullOrEmpty(valor))
                        {
                            worksheet.Cells[ultimaLinha, coluna].Value = valor;
                            Console.WriteLine($"      📝 '{cabecalho}': '{LimitarTexto(valor, 30)}'");
                        }
                    }

                    // Salvar a planilha
                    package.Save();

                    Console.WriteLine($"\n   ✅ Dados adicionados na linha {ultimaLinha}!");
                    Console.WriteLine($"   💾 Planilha salva: {Path.GetFileName(_caminhoPlanilha)}");

                    // Mostrar debug
                    DebugColocacaoDados(worksheet, ultimaLinha, colunas);

                    // Verificar se foi salvo corretamente
                    return VerificarSeCotacaoFoiAdicionada(numeroCotacao, ultimaLinha);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ❌ ERRO ao adicionar na planilha: {ex.Message}");
                Console.WriteLine($"   📌 StackTrace: {ex.StackTrace}");
                return false;
            }
        }
        // Adicione este método para debug
        private void DebugColocacaoDados(ExcelWorksheet worksheet, int linha, Dictionary<string, int> colunas)
        {
            Console.WriteLine($"\n   🔍 DEBUG - ONDE OS DADOS FORAM COLOCADOS:");

            foreach (var kvp in colunas)
            {
                var valor = worksheet.Cells[linha, kvp.Value].Text?.Trim() ?? "[vazio]";
                Console.WriteLine($"      Coluna {kvp.Value} ({kvp.Key}): '{LimitarTexto(valor, 30)}'");
            }
        }
        private Dictionary<string, int> MapearColunasPlanilha(ExcelWorksheet worksheet)
        {
            var colunas = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            if (worksheet.Dimension == null)
                return colunas;

            // Procurar cabeçalhos na primeira linha
            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                string cabecalho = worksheet.Cells[1, col].Text?.Trim() ?? "";

                if (!string.IsNullOrEmpty(cabecalho))
                {
                    colunas[cabecalho.ToUpper()] = col;
                    Console.WriteLine($"      Coluna {col}: '{cabecalho}'");
                }
            }

            return colunas;
        }
        private bool VerificarSeCotacaoFoiAdicionada(string numeroCotacao, int linhaEsperada)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(_caminhoPlanilha)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension?.Rows ?? 0;
                    int colCount = worksheet.Dimension?.Columns ?? 0;

                    // Verificar na linha esperada (procurar em TODAS as colunas)
                    bool encontradaNaLinhaEsperada = false;
                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cells[linhaEsperada, col].Text?.Trim() ?? "";
                        if (cellValue.Contains(numeroCotacao))
                        {
                            Console.WriteLine($"   ✅ Verificação: Cotação {numeroCotacao} encontrada na linha {linhaEsperada}, coluna {col}");
                            return true;
                        }
                    }

                    if (!encontradaNaLinhaEsperada)
                    {
                        Console.WriteLine($"   ⚠ Aviso: Cotação não encontrada na linha esperada. Procurando em toda planilha...");

                        // Procurar em toda planilha
                        for (int row = 1; row <= rowCount; row++)
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
                                var cellValue = worksheet.Cells[row, col].Text?.Trim() ?? "";
                                if (cellValue.Contains(numeroCotacao))
                                {
                                    Console.WriteLine($"   ✅ Encontrada na linha {row}, coluna {col}");
                                    return true;
                                }
                            }
                        }

                        Console.WriteLine($"   ❌ Cotação {numeroCotacao} NÃO encontrada na planilha após busca completa");
                        return false;
                    }

                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠ Erro na verificação: {ex.Message}");
                return false;
            }
        }
        private string LimitarTexto(string texto, int maxLength)
        {
            if (string.IsNullOrEmpty(texto)) return "";
            return texto.Length <= maxLength ? texto : texto.Substring(0, maxLength) + "...";
        }
    }
}