using CotacoesAriba;
using OfficeOpenXml;

class Program
{
    private static List<string> _empresasPrioritarias = new List<string>();
    private static CancellationTokenSource _cts;

    [STAThread]
    static async Task Main(string[] args)
    {
        Console.Clear();

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
        Console.ReadKey();
    }
}