using CotacoesAriba;
using OfficeOpenXml;
using System.Text;

class Program
{
    private static List<string> _empresasPrioritarias = new List<string>();
    private static CancellationTokenSource _cts;

    [STAThread]
    static async Task Main(string[] args)
    {
        await Log();
    }
    private static async Task Iniciar()
    {
        // Configurar prioridades UMA VEZ no início
        ConfigurarPrioridades();

        while (true) // Loop infinito - sempre processa com as mesmas prioridades
        {
            try
            {
                Console.WriteLine("\n" + new string('=', 60));
                Console.WriteLine("🔄 INICIANDO CICLO DE PROCESSAMENTO");
                Console.WriteLine($"Data/Hora: {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
                Console.WriteLine($"Prioridade: {string.Join(" > ", _empresasPrioritarias)}");
                Console.WriteLine(new string('=', 60));

                // Configurar licença EPPlus
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Criar e executar processador
                var processador = new ProcessadorAutomatico(_empresasPrioritarias);
                var estatisticas = await processador.ExecutarProcessamentoCompletoAsync();

                Console.WriteLine($"\n✅ Ciclo concluído em {estatisticas.DuracaoTotal:hh\\:mm\\:ss}");
                Console.WriteLine($"📊 Total de cotações: {estatisticas.TotalCotacoes}");

                // Fechar recursos
                await processador.Fechar();

                // Aguardar próximo ciclo (5 minutos)
                Console.WriteLine($"\n⏳ Aguardando 5 minutos para próximo ciclo...");

                // Mostrar contagem regressiva
                for (int i = 300; i > 0; i--)
                {
                    if (Console.KeyAvailable) // Permitir interrupção com tecla
                    {
                        var key = Console.ReadKey(true);
                        if (key.Key == ConsoleKey.Escape)
                        {
                            Console.WriteLine("\n🚪 Sistema interrompido pelo usuário");
                            return;
                        }
                    }

                    Console.Write($"\rPróximo ciclo em: {i / 60:D2}:{i % 60:D2} segundos (ESC para sair)");
                    await Task.Delay(1000);
                }

                Console.WriteLine(); // Nova linha
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n💥 ERRO CRÍTICO: {ex.Message}");
                Console.WriteLine(ex.StackTrace);

                // Aguardar 30 segundos antes de tentar novamente
                Console.WriteLine($"\n⏳ Tentando novamente em 30 segundos...");
                await Task.Delay(30000);
            }
        }
    }
    private static void ConfigurarPrioridades()
    {
        Console.Clear();
        Console.WriteLine("╔══════════════════════════════════════════════════════════╗");
        Console.WriteLine("║         CONFIGURAÇÃO DE PRIORIDADE DE EMPRESAS          ║");
        Console.WriteLine("╚══════════════════════════════════════════════════════════╝");
        Console.WriteLine("\nDefina a ordem de prioridade (separada por espaços):");
        Console.WriteLine("\n[1] ALIANÇA");
        Console.WriteLine("[2] VENTURA");
        Console.WriteLine("[3] UNIÃO");
        Console.WriteLine("\nExemplo: '1 2 3' para: ALIANÇA > VENTURA > UNIÃO");
        Console.WriteLine("Exemplo: '3 1' para: UNIÃO > ALIANÇA (VENTURA ignorada)");
        Console.WriteLine("\n" + new string('─', 60));
        Console.Write("Prioridade: ");

        string input = Console.ReadLine()?.Trim() ?? "";
        var numeros = input.Split(' ', StringSplitOptions.RemoveEmptyEntries);

        _empresasPrioritarias.Clear();

        foreach (var num in numeros)
        {
            switch (num)
            {
                case "1":
                    if (!_empresasPrioritarias.Contains("ALIANÇA"))
                        _empresasPrioritarias.Add("ALIANÇA");
                    break;
                case "2":
                    if (!_empresasPrioritarias.Contains("VENTURA"))
                        _empresasPrioritarias.Add("VENTURA");
                    break;
                case "3":
                    if (!_empresasPrioritarias.Contains("UNIÃO"))
                        _empresasPrioritarias.Add("UNIÃO");
                    break;
            }
        }

        // Se nenhuma selecionada, usar todas
        if (_empresasPrioritarias.Count == 0)
        {
            _empresasPrioritarias = new List<string> { "ALIANÇA", "VENTURA", "UNIÃO" };
            Console.WriteLine("\n⚠️ Usando todas as empresas por padrão");
        }

        Console.WriteLine($"\n✅ Prioridade configurada: {string.Join(" > ", _empresasPrioritarias)}");
        Console.WriteLine("\n" + new string('─', 60));
        Console.WriteLine("Pressione qualquer tecla para iniciar...");
        Console.Read();
    }
    static async Task Log()
    {

        using var logger = new ConsoleFileLogger(@"\\SERVIDOR2\Publico\ALLAN\Logs");

        Console.WriteLine("=== INICIANDO APLICAÇÃO ===");
        Console.WriteLine($"Data: {DateTime.Now:F}");
        Console.WriteLine();

        try
        {
            Console.WriteLine("Chamando Iniciar...");
            await Iniciar();

            Console.WriteLine("Processamento concluído com sucesso!");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"!!! ERRO CAPTURADO !!!");
            Console.Error.WriteLine($"Mensagem: {ex.Message}");
            Console.Error.WriteLine($"StackTrace: {ex.StackTrace}");
        }

        Console.WriteLine();
        Console.WriteLine("=== APLICAÇÃO FINALIZADA ===");
    }
}
public class ConsoleFileLogger : IDisposable
{
    private readonly string _logDirectory;
    private readonly StreamWriter _fileWriter;
    private readonly TextWriter _originalOutput;
    private readonly TextWriter _originalError;
    private readonly MultiTextWriter _multiOutput;
    private readonly MultiTextWriter _multiError;

    public ConsoleFileLogger(string logDirectory)
    {
        _logDirectory = logDirectory;
        Directory.CreateDirectory(logDirectory);

        // Salva os escritores originais
        _originalOutput = Console.Out;
        _originalError = Console.Error;

        // Cria o arquivo de log com data no nome
        var logFile = Path.Combine(logDirectory, $"aribasourcing.txt");

        // StreamWriter com AutoFlush = true para escrever IMEDIATAMENTE
        _fileWriter = new StreamWriter(logFile, append: true)
        {
            AutoFlush = true  // <--- ESSENCIAL para escrever continuamente
        };

        // Escreve cabeçalho no início do log
        _fileWriter.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] === SESSÃO INICIADA ===");

        // Cria escritores que escrevem tanto no console quanto no arquivo
        _multiOutput = new MultiTextWriter(_originalOutput, _fileWriter);
        _multiError = new MultiTextWriter(_originalError, _fileWriter);

        // Redireciona o console
        Console.SetOut(_multiOutput);
        Console.SetError(_multiError);
    }

    public void Dispose()
    {
        // Escreve rodapé no final do log
        _fileWriter.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] === SESSÃO FINALIZADA ===");
        _fileWriter.WriteLine();

        // Restaura o console original
        Console.SetOut(_originalOutput);
        Console.SetError(_originalError);
        _fileWriter?.Dispose();
    }
}

public class MultiTextWriter : TextWriter
{
    private readonly TextWriter[] _writers;

    public MultiTextWriter(params TextWriter[] writers)
    {
        _writers = writers;
    }

    public override void Write(char value)
    {
        foreach (var writer in _writers)
        {
            writer.Write(value);
        }
    }

    public override void Write(string value)
    {
        foreach (var writer in _writers)
        {
            writer.Write(value);
        }
    }

    public override void WriteLine(string value)
    {
        foreach (var writer in _writers)
        {
            writer.WriteLine(value);
        }
    }

    public override Encoding Encoding => Encoding.UTF8;
}

