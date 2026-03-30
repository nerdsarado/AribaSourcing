using CotacoesAriba;
using System.Text;

public class ConfiguracoesSistema
{
    public bool ModoHeadless { get; set; } = true;
    public bool MostrarConsole { get; set; } = false;
    public int IntervaloEntreCiclosMinutos { get; set; } = 5;
    public int IntervaloEntreContasSegundos { get; set; } = 10;
    public int TentativasPorConta { get; set; } = 3;
    public string PastaLogs { get; set; } = "Cotacoes_Ariba/Logs";
    public List<string> EmpresasPrioritarias { get; set; } = new List<string>();
}
public class FileLogger
{
    private readonly string _logDirectory;
    private readonly string _logFile;
    private readonly bool _ativo;

    public FileLogger(string logDirectory = "Logs")
    {
        _logDirectory = logDirectory;
        _logFile = Path.Combine(logDirectory, $"log_{DateTime.Now:yyyyMMdd}.txt");
        _ativo = true;

        Directory.CreateDirectory(logDirectory);
    }

    public void LogInfo(string mensagem)
    {
        Log("INFO", mensagem);
    }

    public void LogErro(string mensagem, Exception ex = null)
    {
        string detalhes = ex != null ? $"{ex.Message}\n{ex.StackTrace}" : "";
        Log("ERRO", $"{mensagem} {detalhes}");
    }

    public void LogSucesso(string mensagem)
    {
        Log("SUCESSO", mensagem);
    }

    public void LogCiclo(int numeroCiclo, int cotacoesProcessadas, TimeSpan duracao)
    {
        string mensagem = $"Ciclo {numeroCiclo} concluido. " +
                         $"Cotacoes: {cotacoesProcessadas}. " +
                         $"Duracao: {duracao:hh\\:mm\\:ss}";
        Log("CICLO", mensagem);
    }

    private void Log(string nivel, string mensagem)
    {
        if (!_ativo) return;

        try
        {
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{nivel}] {mensagem}";

            // Escrever no arquivo
            File.AppendAllText(_logFile, logEntry + Environment.NewLine);

            // Se quiser também mostrar no console (opcional)
            if (nivel == "ERRO")
            {
                Console.Error.WriteLine(logEntry);
            }
        }
        catch
        {
            // Ignorar erros de logging
        }
    }
    public class SistemaCotacoesHeadless
    {
        private readonly ProcessadorAutomatico _processador;
        private readonly FileLogger _logger;
        private readonly ConfiguracoesSistema _config;
        private CancellationTokenSource _cancellationTokenSource;
        private bool _executando;

        private int _totalCiclos;
        private int _totalCotacoes;
        private DateTime _inicioExecucao;

        public SistemaCotacoesHeadless(ConfiguracoesSistema config = null)
        {
            _config = config ?? new ConfiguracoesSistema();
            _logger = new FileLogger(_config.PastaLogs);
            _processador = new ProcessadorAutomatico();

            // Configurar navegador headless se necessário
            ConfigurarNavegadorHeadless();
        }

        private void ConfigurarNavegadorHeadless()
        {
            // Esta configuração será aplicada quando o navegador for criado
            // O ProcessadorAutomatico deve usar essas configurações
        }

        public async Task IniciarAsync()
        {
            _executando = true;
            _cancellationTokenSource = new CancellationTokenSource();
            _inicioExecucao = DateTime.Now;

            _logger.LogInfo("Sistema de Cotacoes Ariba iniciado");
            _logger.LogInfo($"Modo: Headless ({(_config.ModoHeadless ? "SIM" : "NAO")})");
            _logger.LogInfo($"Intervalo entre ciclos: {_config.IntervaloEntreCiclosMinutos} minutos");

            try
            {
                while (_executando && !_cancellationTokenSource.Token.IsCancellationRequested)
                {
                    try
                    {
                        await ExecutarCicloAsync(_cancellationTokenSource.Token);

                        if (_executando && !_cancellationTokenSource.Token.IsCancellationRequested)
                        {
                            await AguardarProximoCicloAsync(_cancellationTokenSource.Token);
                        }
                    }
                    catch (TaskCanceledException)
                    {
                        _logger.LogInfo("Ciclo cancelado");
                        break;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogErro("Erro durante execucao do ciclo", ex);
                        await Task.Delay(TimeSpan.FromMinutes(1));
                    }
                }
            }
            finally
            {
                await FinalizarAsync();
            }
        }

        private async Task ExecutarCicloAsync(CancellationToken cancellationToken)
        {
            _totalCiclos++;
            DateTime inicioCiclo = DateTime.Now;
            int cotacoesEsteCiclo = 0;

            _logger.LogInfo($"Iniciando ciclo {_totalCiclos}");

            try
            {
                for (int contaNumero = 1; contaNumero <= 4; contaNumero++)
                {
                    if (cancellationToken.IsCancellationRequested) break;

                    string chaveConta = contaNumero.ToString();
                    _logger.LogInfo($"Processando conta {chaveConta}");

                    try
                    {
                        var estatisticas = await _processador.ExecutarProcessamentoCompletoAsync();

                        if (estatisticas.CotacoesRegistradas > 0)
                        {
                            cotacoesEsteCiclo += estatisticas.CotacoesRegistradas;
                            _totalCotacoes += estatisticas.CotacoesRegistradas;

                            _logger.LogSucesso($"Conta {chaveConta}: {estatisticas.CotacoesRegistradas} cotacoes processadas");
                        }
                        else
                        {
                            _logger.LogInfo($"Conta {chaveConta}: Nenhuma cotacao nova encontrada");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogErro($"Erro na conta {chaveConta}", ex);
                    }

                    if (contaNumero < 4 && !cancellationToken.IsCancellationRequested)
                    {
                        await Task.Delay(TimeSpan.FromSeconds(_config.IntervaloEntreContasSegundos), cancellationToken);
                    }
                }

                TimeSpan duracaoCiclo = DateTime.Now - inicioCiclo;
                _logger.LogCiclo(_totalCiclos, cotacoesEsteCiclo, duracaoCiclo);

                // Salvar status periodicamente
                if (_totalCiclos % 10 == 0)
                {
                    SalvarStatusAtual();
                }
            }
            catch (Exception ex)
            {
                _logger.LogErro($"Erro critico no ciclo {_totalCiclos}", ex);
                throw;
            }
        }

        private async Task AguardarProximoCicloAsync(CancellationToken cancellationToken)
        {
            DateTime proximoCiclo = DateTime.Now.AddMinutes(_config.IntervaloEntreCiclosMinutos);
            _logger.LogInfo($"Proximo ciclo em: {proximoCiclo:HH:mm:ss}");

            try
            {
                await Task.Delay(TimeSpan.FromMinutes(_config.IntervaloEntreCiclosMinutos), cancellationToken);
            }
            catch (TaskCanceledException)
            {
                // Cancelamento normal
            }
        }

        private void SalvarStatusAtual()
        {
            try
            {
                string statusFile = Path.Combine(_config.PastaLogs, "status_atual.json");
                var status = new
                {
                    InicioExecucao = _inicioExecucao,
                    TotalCiclos = _totalCiclos,
                    TotalCotacoes = _totalCotacoes,
                    UltimaAtualizacao = DateTime.Now,
                    Status = _executando ? "EXECUTANDO" : "PARADO"
                };

                string json = System.Text.Json.JsonSerializer.Serialize(status, new System.Text.Json.JsonSerializerOptions
                {
                    WriteIndented = true
                });

                File.WriteAllText(statusFile, json);
            }
            catch
            {
                // Ignorar erros
            }
        }

        public async Task PararAsync()
        {
            _executando = false;
            _cancellationTokenSource?.Cancel();

            _logger.LogInfo("Solicitando parada do sistema...");
            await Task.Delay(1000);
        }

        private async Task FinalizarAsync()
        {
            try
            {
                _logger.LogInfo("Finalizando sistema...");

                TimeSpan tempoTotal = DateTime.Now - _inicioExecucao;

                // Salvar relatório final
                SalvarRelatorioFinal(tempoTotal);

                await _processador.Fechar();

                _logger.LogInfo($"Sistema finalizado. Tempo total: {tempoTotal:hh\\:mm\\:ss}");
                _logger.LogInfo($"Ciclos completos: {_totalCiclos}");
                _logger.LogInfo($"Cotacoes processadas: {_totalCotacoes}");
            }
            catch (Exception ex)
            {
                _logger.LogErro("Erro ao finalizar sistema", ex);
            }
        }

        private void SalvarRelatorioFinal(TimeSpan tempoTotal)
        {
            try
            {
                string relatorioFile = Path.Combine(_config.PastaLogs, $"relatorio_final_{DateTime.Now:yyyyMMdd_HHmmss}.txt");

                string relatorio = $@"
========================================
RELATORIO FINAL - SISTEMA DE COTAÇÕES ARIBA
========================================
Data/Hora: {DateTime.Now:dd/MM/yyyy HH:mm:ss}
Tempo total de execucao: {tempoTotal:hh\:mm\:ss}
Ciclos completos: {_totalCiclos}
Cotacoes processadas: {_totalCotacoes}
Inicio: {_inicioExecucao:dd/MM/yyyy HH:mm:ss}
Termino: {DateTime.Now:dd/MM/yyyy HH:mm:ss}

CONFIGURACOES:
- Modo Headless: {_config.ModoHeadless}
- Intervalo entre ciclos: {_config.IntervaloEntreCiclosMinutos} minutos
- Tentativas por conta: {_config.TentativasPorConta}

========================================
SISTEMA FINALIZADO COM SUCESSO
========================================
";

                File.WriteAllText(relatorioFile, relatorio, Encoding.UTF8);
            }
            catch
            {
                // Ignorar erros
            }
        }
    }
}