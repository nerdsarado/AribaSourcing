// Notificador.cs - Correção completa
using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace CotacoesAriba
{
    public class Notificador
    {
        // Importações do Windows API para manipulação de janela
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool FlashWindow(IntPtr hWnd, bool bInvert);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        // Constantes
        private const int SW_RESTORE = 9;
        private const int SW_SHOW = 5;
        private const int SW_SHOWMAXIMIZED = 3;
        private const int WM_SYSCOMMAND = 0x0112;
        private const int SC_RESTORE = 0xF120;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);

        private static IntPtr _consoleWindow = IntPtr.Zero;
        private static readonly object _lock = new object();

        public enum TipoNotificacao
        {
            COTACAO_PRIORITARIA,
            COTACAO_UNICA_NAO_PRIORITARIA,
            ERRO,
            AVISO
        }

        static Notificador()
        {
            _consoleWindow = GetConsoleWindow();
        }

        // Método principal que realmente traz a janela para frente
        public static void NotificarCotacaoPrioritaria(string numeroCotacao, string empresa, int totalItens)
        {
            string titulo = $"✅ COTAÇÃO PRIORITÁRIA PROCESSADA - {empresa}";
            string mensagem = $"📋 Cotação: {numeroCotacao}\n🏢 Empresa: {empresa}\n📦 Itens: {totalItens}\n⏰ {DateTime.Now:HH:mm:ss}\n\n✅ Documento baixado e registrado no banco!";

            MostrarNotificacaoForcada(titulo, mensagem, TipoNotificacao.COTACAO_PRIORITARIA);
        }

        public static void NotificarCotacaoUnicaNaoPrioritaria(string numeroCotacao, string empresa, int totalUnicas)
        {
            string titulo = $"⚠️ COTAÇÃO ÚNICA ENCONTRADA - {empresa}";
            string mensagem = $"📋 Cotação: {numeroCotacao}\n🏢 Empresa: {empresa}\n📊 Total únicas: {totalUnicas}\n⏰ {DateTime.Now:HH:mm:ss}\n\n⚠️ Existe APENAS em empresa não prioritária!";

            MostrarNotificacaoForcada(titulo, mensagem, TipoNotificacao.COTACAO_UNICA_NAO_PRIORITARIA);
        }

        public static void NotificarErro(string operacao, string erro)
        {
            string titulo = $"❌ ERRO NO PROCESSAMENTO";
            string mensagem = $"🔧 Operação: {operacao}\n💥 Erro: {erro}\n⏰ {DateTime.Now:HH:mm:ss}\n\n⚠️ Verifique o sistema!";

            MostrarNotificacaoForcada(titulo, mensagem, TipoNotificacao.ERRO);
        }

        private static void MostrarNotificacaoForcada(string titulo, string mensagem, TipoNotificacao tipo)
        {
            // Executar em uma thread separada para não bloquear o processamento
            ThreadPool.QueueUserWorkItem(state =>
            {
                try
                {
                    lock (_lock)
                    {
                        // 1. Tocar som ALTO para alertar
                        TocarSomAlerta(tipo);

                        // 2. Trazer console para frente FORÇADAMENTE
                        ForcarConsoleParaFrente();

                        // 3. Aguardar um pouco para garantir que está na frente
                        Thread.Sleep(500);

                        // 4. Mostrar notificação colorida
                        MostrarNotificacaoColorida(titulo, mensagem, tipo);

                        // 5. Manter na frente por alguns segundos
                        Thread.Sleep(3000);

                        // 6. Logar no arquivo
                        LogarNotificacao(titulo, mensagem, tipo);
                    }
                }
                catch (Exception ex)
                {
                    // Se falhar a notificação fancy, pelo menos logar
                    Console.WriteLine($"[ERRO NOTIFICAÇÃO] {ex.Message}");
                    LogarNotificacao("ERRO NA NOTIFICAÇÃO", ex.Message, TipoNotificacao.ERRO);
                }
            });
        }

        private static void TocarSomAlerta(TipoNotificacao tipo)
        {
            try
            {
                // Usar múltiplos beeps para chamar atenção
                for (int i = 0; i < 3; i++)
                {
                    switch (tipo)
                    {
                        case TipoNotificacao.COTACAO_PRIORITARIA:
                            Console.Beep(1000, 200);
                            Thread.Sleep(50);
                            Console.Beep(1200, 300);
                            break;
                        case TipoNotificacao.COTACAO_UNICA_NAO_PRIORITARIA:
                            Console.Beep(800, 300);
                            Thread.Sleep(100);
                            Console.Beep(800, 300);
                            Thread.Sleep(100);
                            Console.Beep(800, 300);
                            break;
                        case TipoNotificacao.ERRO:
                            Console.Beep(500, 400);
                            Thread.Sleep(50);
                            Console.Beep(400, 400);
                            Thread.Sleep(50);
                            Console.Beep(300, 600);
                            break;
                        default:
                            Console.Beep();
                            Console.Beep();
                            break;
                    }
                    Thread.Sleep(100);
                }
            }
            catch
            {
                // Fallback simples
                for (int i = 0; i < 5; i++)
                {
                    Console.Beep();
                    Thread.Sleep(100);
                }
            }
        }

        private static void ForcarConsoleParaFrente()
        {
            try
            {
                if (_consoleWindow == IntPtr.Zero)
                    _consoleWindow = GetConsoleWindow();

                if (_consoleWindow != IntPtr.Zero)
                {
                    // Restaurar se minimizado
                    ShowWindow(_consoleWindow, SW_RESTORE);

                    // Trazer para frente
                    SetForegroundWindow(_consoleWindow);

                    // Garantir que está no topo
                    SetWindowPos(_consoleWindow, HWND_TOPMOST, 0, 0, 0, 0,
                        SWP_NOSIZE | SWP_NOMOVE | SWP_SHOWWINDOW);

                    // Piscar na barra de tarefas
                    FlashWindow(_consoleWindow, true);

                    // Forçar foco
                    SetForegroundWindow(_consoleWindow);

                    // Aguardar um pouco
                    Thread.Sleep(100);

                    // Forçar novamente (às vezes precisa de múltiplas tentativas)
                    SetForegroundWindow(_consoleWindow);
                }
            }
            catch
            {
                // Se tudo falhar, pelo menos tenta o básico
                try
                {
                    if (_consoleWindow != IntPtr.Zero)
                    {
                        ShowWindow(_consoleWindow, SW_SHOWMAXIMIZED);
                        SetForegroundWindow(_consoleWindow);
                    }
                }
                catch
                {
                    // Último recurso
                }
            }
        }

        private static void MostrarNotificacaoColorida(string titulo, string mensagem, TipoNotificacao tipo)
        {
            ConsoleColor oldForeground = Console.ForegroundColor;
            ConsoleColor oldBackground = Console.BackgroundColor;

            try
            {
                // Posicionar no topo da tela
                Console.SetCursorPosition(0, 0);

                // Determinar cores
                ConsoleColor bgColor, fgColor, borderColor;

                switch (tipo)
                {
                    case TipoNotificacao.COTACAO_PRIORITARIA:
                        bgColor = ConsoleColor.DarkGreen;
                        fgColor = ConsoleColor.White;
                        borderColor = ConsoleColor.Green;
                        break;
                    case TipoNotificacao.COTACAO_UNICA_NAO_PRIORITARIA:
                        bgColor = ConsoleColor.DarkYellow;
                        fgColor = ConsoleColor.Black;
                        borderColor = ConsoleColor.Yellow;
                        break;
                    case TipoNotificacao.ERRO:
                        bgColor = ConsoleColor.DarkRed;
                        fgColor = ConsoleColor.White;
                        borderColor = ConsoleColor.Red;
                        break;
                    default:
                        bgColor = ConsoleColor.DarkBlue;
                        fgColor = ConsoleColor.White;
                        borderColor = ConsoleColor.Cyan;
                        break;
                }

                int largura = Console.WindowWidth;
                string linhaBorda = new string('═', largura - 2);
                string linhaEspaco = new string(' ', largura - 2);

                // Limpar área da notificação
                Console.BackgroundColor = bgColor;
                Console.ForegroundColor = borderColor;

                // Borda superior
                Console.WriteLine("╔" + linhaBorda + "╗");

                // Título centralizado
                Console.Write("║");
                Console.BackgroundColor = bgColor;
                Console.ForegroundColor = fgColor;
                int espacosTitulo = (largura - 2 - titulo.Length) / 2;
                Console.Write(new string(' ', espacosTitulo));
                Console.Write(titulo);
                Console.Write(new string(' ', largura - 2 - titulo.Length - espacosTitulo));
                Console.BackgroundColor = bgColor;
                Console.ForegroundColor = borderColor;
                Console.WriteLine("║");

                // Linha divisória
                Console.WriteLine("╠" + linhaBorda + "╣");

                // Mensagem
                string[] linhasMensagem = mensagem.Split('\n');
                foreach (var linha in linhasMensagem)
                {
                    Console.Write("║");
                    Console.BackgroundColor = bgColor;
                    Console.ForegroundColor = fgColor;

                    if (linha.Length <= largura - 2)
                    {
                        Console.Write(linha.PadRight(largura - 2));
                    }
                    else
                    {
                        Console.Write(linha.Substring(0, largura - 2));
                    }

                    Console.BackgroundColor = bgColor;
                    Console.ForegroundColor = borderColor;
                    Console.WriteLine("║");
                }

                // Rodapé com hora
                Console.Write("║");
                Console.BackgroundColor = bgColor;
                Console.ForegroundColor = fgColor;
                string rodape = $"[{DateTime.Now:HH:mm:ss}] Pressione qualquer tecla para continuar...";
                Console.Write(rodape.PadRight(largura - 2));
                Console.BackgroundColor = bgColor;
                Console.ForegroundColor = borderColor;
                Console.WriteLine("║");

                // Borda inferior
                Console.WriteLine("╚" + linhaBorda + "╝");

                // Forçar exibição
                Console.Out.Flush();

                // Aguardar tecla sem bloquear o console inteiro
                WaitForAnyKeyNonBlocking(10000); // Timeout de 10 segundos

                // Limpar notificação
                Console.SetCursorPosition(0, 0);
                for (int i = 0; i < 5 + linhasMensagem.Length; i++)
                {
                    Console.WriteLine(new string(' ', largura));
                }

                // Restaurar posição do cursor
                Console.SetCursorPosition(0, 0);
            }
            finally
            {
                Console.ForegroundColor = oldForeground;
                Console.BackgroundColor = oldBackground;
            }
        }

        private static void WaitForAnyKeyNonBlocking(int timeoutMilliseconds)
        {
            DateTime start = DateTime.Now;

            while ((DateTime.Now - start).TotalMilliseconds < timeoutMilliseconds)
            {
                if (Console.KeyAvailable)
                {
                    Console.ReadKey(true); // Consumir a tecla
                    break;
                }
                Thread.Sleep(100);
            }
        }

        private static void LogarNotificacao(string titulo, string mensagem, TipoNotificacao tipo)
        {
            try
            {
                string logDir = "Logs_Notificacoes";
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);

                string logFile = Path.Combine(logDir, $"notificacoes_{DateTime.Now:yyyyMMdd}.log");
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{tipo}] {titulo} - {mensagem.Replace("\n", " | ")}\n";

                File.AppendAllText(logFile, logEntry);
            }
            catch
            {
                // Ignorar erros de log
            }
        }

        // Método alternativo SIMPLES que SEMPRE funciona
        public static void NotificarSimples(string titulo, string mensagem, TipoNotificacao tipo)
        {
            try
            {
                // Abrir uma nova janela de console separada
                var process = new System.Diagnostics.Process();
                process.StartInfo.FileName = "cmd.exe";
                process.StartInfo.Arguments = $"/c title {titulo} && echo {mensagem} && echo. && echo Pressione qualquer tecla para continuar... && pause >nul";
                process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
                process.StartInfo.UseShellExecute = true;
                process.Start();

                // Logar
                LogarNotificacao(titulo, mensagem, tipo);
            }
            catch (Exception ex)
            {
                // Último recurso: escrever no console atual
                Console.WriteLine($"\n🚨 NOTIFICAÇÃO: {titulo}");
                Console.WriteLine($"📝 {mensagem}");
                Console.WriteLine("⚠️ Pressione ENTER para continuar...");
                Console.ReadLine();
            }
        }

        // Métodos públicos que usam a abordagem simples
        public static void NotificarCotacaoPrioritariaSimples(string numeroCotacao, string empresa, int totalItens)
        {
            string titulo = $"COTAÇÃO PRIORITÁRIA - {empresa}";
            string mensagem = $"Cotação {numeroCotacao} processada com {totalItens} itens às {DateTime.Now:HH:mm:ss}";
            NotificarSimples(titulo, mensagem, TipoNotificacao.COTACAO_PRIORITARIA);
        }

        public static void NotificarCotacaoUnicaNaoPrioritariaSimples(string numeroCotacao, string empresa, int totalUnicas)
        {
            string titulo = $"COTAÇÃO ÚNICA - {empresa}";
            string mensagem = $"Cotação {numeroCotacao} encontrada apenas em {empresa} ({totalUnicas} únicas) às {DateTime.Now:HH:mm:ss}";
            NotificarSimples(titulo, mensagem, TipoNotificacao.COTACAO_UNICA_NAO_PRIORITARIA);
        }
    }
}