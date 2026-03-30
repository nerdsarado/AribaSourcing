// GerenciadorLogin.cs
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Playwright;
using OfficeOpenXml;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;


namespace CotacoesAriba
{
    public class GerenciadorLogin
    {
        private IPlaywright _playwright;
        private IBrowser _browser;
        private ExcelDatabaseManager _excelDb;
        private string _empresaAtual = "AEGEA";
        

        

        public void SetEmpresaAtual(string empresa)
        {
            if (empresa == "AEGEA" || empresa == "ESTÁCIO")
            {
                _empresaAtual = empresa;
                Console.WriteLine($"🏢 Empresa definida para: {empresa}");
            }
            else
            {
                Console.WriteLine($"⚠️ Empresa inválida: {empresa}, usando padrão AEGEA");
                _empresaAtual = "AEGEA";
            }
        }
        // Definição das contas
        public readonly Dictionary<string, Conta> Contas = new()
        {
            ["1"] = new Conta("ALIANÇA", "alianca@venturainformatica.com.br", "Alianca26*", "CNPJ_ALIANCA"),
            ["2"] = new Conta("ALIANÇA", "vendas@venturainformatica.com.br", "Alianca@2026**", "CNPJ_ALIANCA"),
            ["3"] = new Conta("VENTURA", "vendas2@venturainformatica.com.br", "Ventura@2026*", "CNPJ_VENTURA"),
            ["4"] = new Conta("UNIÃO", "uniao@venturainformatica.com.br", "Uniao2026@", "CNPJ_UNIAO")
        };

        // Classe para representar um item da cotação
        public class ItemCotacao
        {
            public string DescricaoOriginal { get; set; } = ""; // Com quantidade
            public string DescricaoLimpa { get; set; } = "";    // Sem quantidade (para a planilha)
            public string Quantidade { get; set; } = "";
            public string Unidade { get; set; } = "unidade";
        }


        // Classe para representar uma conta
        public record Conta(string Nome, string Email, string Senha, string Cnpj = null);
        public GerenciadorLogin()
        {
            try
            {
                // Configurar licença do EPPlus se ainda não foi configurada
                if (ExcelPackage.LicenseContext == LicenseContext.NonCommercial)
                {
                    Console.WriteLine("Licenca EPPlus configurada (NonCommercial)");
                }
                // Tente diferentes extensões possíveis
                string[] possiveisExtensoes = { ".xlsx", ".xls", ".xlsm" };
            string caminhoBase = @"\\SERVIDOR2\Publico\ANMYNA\CONTROLE PEDIDOS EQUIPE ANMYNA 2025 (Salvo automaticamente)";
            string caminhoCompleto = null;

            foreach (var extensao in possiveisExtensoes)
            {
                string caminhoTeste = caminhoBase + extensao;
                if (File.Exists(caminhoTeste))
                {
                    caminhoCompleto = caminhoTeste;
                    Console.WriteLine($"✅ Planilha encontrada: {caminhoTeste}");
                    break;
                }
            }

            if (caminhoCompleto == null)
            {
                // Se não encontrou com extensão, tenta sem extensão (pode ser o nome completo)
                if (File.Exists(caminhoBase))
                {
                    caminhoCompleto = caminhoBase;
                    Console.WriteLine($"✅ Planilha encontrada (sem extensão explícita): {caminhoBase}");
                }
                else
                {
                    Console.WriteLine($"⚠ AVISO: Não foi possível encontrar a planilha");
                    Console.WriteLine($"📌 Caminho base: {caminhoBase}");
                    Console.WriteLine($"🔍 Verifique se o servidor Z: está mapeado e acessível");
                }
            }

            _excelDb = new ExcelDatabaseManager(caminhoCompleto ?? caminhoBase + ".xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERRO ao inicializar GerenciadorLogin: {ex.Message}");
                throw;
            }

        }
        public async Task<IPage> RealizarLoginAsync(string chaveConta)
        {
            try
            {
                // Verificar se a conta existe no dicionário
                if (!Contas.TryGetValue(chaveConta, out var conta))
                {
                    Console.WriteLine($"ERRO: Conta {chaveConta} nao encontrada");
                    return null;
                }

                Console.WriteLine($"Tentando login: {conta.Nome}");

                // Verificar se já existe um contexto aberto
                if (_browser == null)
                {
                    Console.WriteLine($"Iniciando navegador em modo headless...");
                    _playwright = await Playwright.CreateAsync();
                    _browser = await _playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
                    {
                        Headless = false
                    });
                }

                // Criar nova página
                var page = await _browser.NewPageAsync();
                await page.SetViewportSizeAsync(1920, 1080);

                // URL de login padrão do Ariba
                string urlLogin = "https://service.ariba.com/Sourcing.aw/109521004/aw?awh=r&awssk=M1Ju378m&dard=1";

                Console.WriteLine($"Navegando para pagina de login...");
                await page.GotoAsync(urlLogin, new PageGotoOptions
                {
                    WaitUntil = WaitUntilState.DOMContentLoaded,
                    Timeout = 60000
                });


                try
                {
                    Console.WriteLine($"   🔍 Procurando campos de login do Ariba...");

                    // Aguardar o formulário carregar completamente
                    await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                    await Task.Delay(3000);

                    // Verificar se estamos na página correta
                    var pageTitle = await page.TitleAsync();
                    var pageUrl = page.Url;
                    Console.WriteLine($"   📄 Título: {pageTitle}");
                    Console.WriteLine($"   🔗 URL: {pageUrl}");

                    // CAPTURA 1: Campo de Email/Usuário
                    Console.WriteLine($"   🔑 Localizando campo de email...");

                    // Tentar múltiplos seletores para o campo de email
                    var emailSelectors = new[]
                    {
                "input[name='UserName']", // Seletor específico do HTML que você forneceu
                "input[name='userName']",
                "input[name='username']",
                "input[id='UserName']",
                "input[id='userName']",
                "input[id='username']",
                "input[type='email']",
                "input[autocomplete='username']",
                "input[placeholder*='email' i], input[placeholder*='user' i]",
                "input.w-txt-dsize" // Classe do campo de email
            };

                    bool emailPreenchido = false;
                    foreach (var selector in emailSelectors)
                    {
                        try
                        {
                            var emailField = page.Locator(selector);
                            var count = await emailField.CountAsync();

                            if (count > 0)
                            {
                                Console.WriteLine($"      ✅ Campo encontrado com: {selector}");

                                // Verificar se está visível
                                var isVisible = await emailField.First.IsVisibleAsync();
                                if (!isVisible)
                                {
                                    Console.WriteLine($"      👁️  Campo não visível, tentando scroll...");
                                    await emailField.First.ScrollIntoViewIfNeededAsync();
                                    await Task.Delay(1000);
                                }

                                // Preencher email
                                Console.WriteLine($"      📧 Preenchendo email: {conta.Email}");
                                await emailField.First.FillAsync(conta.Email);
                                await Task.Delay(1000);

                                emailPreenchido = true;
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"      ⚠️ Erro com seletor {selector}: {ex.Message}");
                            continue;
                        }
                    }

                    if (!emailPreenchido)
                    {
                        Console.WriteLine($"   ❌ Campo de email não encontrado");
                        throw new Exception("Campo de email não encontrado");
                          
                    }

                    // CAPTURA 2: Campo de Senha
                    Console.WriteLine($"   🔑 Localizando campo de senha...");

                    var passwordSelectors = new[]
                    {
                "input[name='Password']", // Seletor específico do HTML que você forneceu
                "input[name='password']",
                "input[id='Password']",
                "input[id='password']",
                "input[type='password']",
                "input.w-psw", // Classe do campo de senha
                "input[autocomplete='current-password']",
                "input[placeholder*='password' i], input[placeholder*='senha' i]"
            };

                    bool passwordPreenchido = false;
                    foreach (var selector in passwordSelectors)
                    {
                        try
                        {
                            var passwordField = page.Locator(selector);
                            var count = await passwordField.CountAsync();

                            if (count > 0)
                            {
                                Console.WriteLine($"      ✅ Campo encontrado com: {selector}");

                                // Verificar se está visível
                                var isVisible = await passwordField.First.IsVisibleAsync();
                                if (!isVisible)
                                {
                                    await passwordField.First.ScrollIntoViewIfNeededAsync();
                                    await Task.Delay(1000);
                                }

                                // Preencher senha
                                Console.WriteLine($"      🔒 Preenchendo senha");
                                await passwordField.First.FillAsync(conta.Senha);
                                await Task.Delay(1000);

                                passwordPreenchido = true;
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"      ⚠️ Erro com seletor {selector}: {ex.Message}");
                            continue;
                        }
                    }

                    if (!passwordPreenchido)
                    {
                        Console.WriteLine($"   ❌ Campo de senha não encontrado");
                        throw new Exception("Campo de senha não encontrado");
                    }

                    // CAPTURA 3: Botão de Login
                    Console.WriteLine($"   🔘 Localizando botão de login...");

                    var buttonSelectors = new[]
                    {
                "button[type='submit']",
                "input[type='submit']",
                "button:has-text('Log In')",
                "button:has-text('Sign In')",
                "button:has-text('Entrar')",
                "button:has-text('Login')",
                "[class*='btn-login']",
                "[class*='btn-submit']",
                "[onclick*='login']",
                "[onclick*='submit']"
            };

                    bool botaoClicado = false;
                    foreach (var selector in buttonSelectors)
                    {
                        try
                        {
                            var loginButton = page.Locator(selector);
                            var count = await loginButton.CountAsync();

                            if (count > 0)
                            {
                                Console.WriteLine($"      ✅ Botão encontrado com: {selector}");

                                // Verificar se está visível
                                var isVisible = await loginButton.First.IsVisibleAsync();
                                if (!isVisible)
                                {
                                    await loginButton.First.ScrollIntoViewIfNeededAsync();
                                    await Task.Delay(1000);
                                }

                                // Clicar no botão
                                Console.WriteLine($"      🖱️ Clicando no botão de login...");
                                await loginButton.First.ClickAsync(new LocatorClickOptions
                                {
                                    Force = true,
                                    Timeout = 10000
                                });

                                botaoClicado = true;
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"      ⚠️ Erro com seletor {selector}: {ex.Message}");
                            continue;
                        }
                    }

                    if (!botaoClicado)
                    {
                        // Tentativa final: pressionar Enter
                        Console.WriteLine($"      ⌨️ Pressionando Enter...");
                        await page.Keyboard.PressAsync("Enter");
                    }

                    Console.WriteLine($"   ✅ Login submetido com sucesso");

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"   ⚠️ Erro no processo de login: {ex.Message}");
                    Console.WriteLine($"   🔄 Tentando método alternativo...");

                    // Método alternativo: usar os seletores exatos do HTML que você forneceu
                    try
                    {
                        Console.WriteLine($"   🎯 Usando seletores exatos do HTML fornecido...");

                        // Campo de email exato
                        var emailField = page.Locator("input[name='UserName']");
                        if (await emailField.CountAsync() > 0)
                        {
                            await emailField.First.FillAsync(conta.Email);
                            await Task.Delay(1000);
                        }

                        // Campo de senha exato
                        var passwordField = page.Locator("input[name='Password']");
                        if (await passwordField.CountAsync() > 0)
                        {
                            await passwordField.First.FillAsync(conta.Senha);
                            await Task.Delay(1000);

                            // Clicar Enter na senha
                            await page.Keyboard.PressAsync("Enter");
                            Console.WriteLine($"   ✅ Login usando seletores exatos");
                        }
                    }
                    catch (Exception ex2)
                    {
                        Console.WriteLine($"   ❌ Método alternativo também falhou: {ex2.Message}");
                    }
                }

                // Aguardar login completar
                Console.WriteLine($"   ⏳ Aguardando login completar (20 segundos)...");
                await Task.Delay(20000); // Aguardar 20 segundos para login processar

                // Verificar se login foi bem-sucedido
                Console.WriteLine($"   🔍 Verificando status do login...");

                // Verificar URL atual
                var currentUrl = page.Url;
                Console.WriteLine($"   🔗 URL atual: {currentUrl}");

                // Verificar se há erros na página
                var pageContent = await page.ContentAsync();
                bool hasError = pageContent.Contains("error", StringComparison.OrdinalIgnoreCase) ||
                               pageContent.Contains("invalid", StringComparison.OrdinalIgnoreCase) ||
                               pageContent.Contains("incorrect", StringComparison.OrdinalIgnoreCase);

                if (hasError)
                {
                    Console.WriteLine($"   ❌ Possível erro detectado na página");

                    // Procurar mensagem de erro específica
                    var errorElements = page.Locator("[class*='error'], [class*='alert'], [class*='message']:has-text('error'), [class*='message']:has-text('invalid')");
                    var errorCount = await errorElements.CountAsync();

                    if (errorCount > 0)
                    {
                        for (int i = 0; i < Math.Min(errorCount, 3); i++)
                        {
                            var errorText = await errorElements.Nth(i).TextContentAsync();
                            Console.WriteLine($"   💬 Erro {i + 1}: {errorText}");
                        }
                    }
                    await page.CloseAsync();
                    return null;
                }

                // Verificar sinais de login bem-sucedido - CORRIGIDO
                bool loginSucesso = false;

                // Verificação 1: URL
                if (currentUrl.Contains("Supplier.aw") && !currentUrl.Contains("login"))
                {
                    loginSucesso = true;
                    Console.WriteLine($"   ✅ Indicador 1: URL correta");
                }

                // Verificação 2: Elementos visíveis
                if (!loginSucesso)
                {
                    var successElements = new[]
                    {
                "text='My Inbox'",
                "text='Sourcing'",
                "text='RFx'",
                "[class*='supplier-dashboard']",
                "[class*='welcome']",
                "text='Dashboard'",
                "text='Supplier'"
            };

                    foreach (var selector in successElements)
                    {
                        try
                        {
                            var element = page.Locator(selector);
                            if (await element.CountAsync() > 0 && await element.First.IsVisibleAsync())
                            {
                                loginSucesso = true;
                                Console.WriteLine($"   ✅ Indicador 2: Elemento '{selector}' encontrado");
                                break;
                            }
                        }
                        catch
                        {
                            // Continuar com próximo seletor
                        }
                    }
                }

                // Verificação 3: Frames (comum no Ariba)
                if (!loginSucesso)
                {
                    var frames = page.Frames;
                    if (frames.Count > 1)
                    {
                        loginSucesso = true;
                        Console.WriteLine($"   ✅ Indicador 3: {frames.Count} frames detectados");
                    }
                }

                // Verificação 4: Conteúdo da página
                if (!loginSucesso && pageContent.Length > 5000)
                {
                    loginSucesso = true;
                    Console.WriteLine($"   ✅ Indicador 4: Conteúdo extenso ({pageContent.Length} caracteres)");
                }

                if (loginSucesso)
                {
                    Console.WriteLine($"   ✅ Login confirmado!");
                    Console.WriteLine($"   ⏳ Aguardando 5 segundos para estabilização...");
                    await Task.Delay(5000);
                    return page;
                }
                else
                {
                    Console.WriteLine($"   ⚠️ Login não confirmado, mas continuando...");
                    Console.WriteLine($"   ⏳ Aguardando 5 segundos...");
                    await Task.Delay(5000);
                    await page.CloseAsync();
                    return page; // Continuar mesmo sem confirmação explícita
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"💥 ERRO CRÍTICO no login: {ex.Message}");
                Console.WriteLine($"📌 StackTrace: {ex.StackTrace}");
                return null;
            }
        }

        public async Task<bool> NavegarParaAegeaAsync(IPage page)
        {
            try
            {
                Console.WriteLine("\n  🚀 Iniciando navegação para Aegea...");

                // Aguardar um pouco mais para garantir que a página carregou completamente
                await Task.Delay(3000);

                // Esperar pelo menu "Mais..."
                await page.WaitForSelectorAsync("a.w-tabitem-a[href='#']", new PageWaitForSelectorOptions
                {
                    Timeout = 15000, // Aumentei de 10 para 15 segundos
                    State = WaitForSelectorState.Visible
                });

                await page.ClickAsync("a.w-tabitem-a[href='#']");
                Console.WriteLine("  ✓ Menu 'Mais...' clicado");

                // Aguardar um pouco mais após clicar no menu
                await Task.Delay(2500); // Aumentei de 1500 para 2500 ms

                // Esperar pelo link da AEGEA
                await page.WaitForSelectorAsync("a#_c0b6dd", new PageWaitForSelectorOptions
                {
                    Timeout = 10000, // Aumentei de 8000 para 10000 ms
                    State = WaitForSelectorState.Visible
                });

                await page.ClickAsync("a#_c0b6dd");
                Console.WriteLine("  ✓ 'AEGEA SANEAMENTO E PARTICIPACOES S.A.' selecionado");

                // Aguardar a navegação completar com um tempo maior
                await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                await Task.Delay(6000); // Aumentei de 4000 para 6000 ms

                // Verificar se a navegação foi bem-sucedida
                string urlAtual = page.Url;
                Console.WriteLine($"  📍 URL atual: {urlAtual}");

                // Verificar se há algum indicador de sucesso na página
                try
                {
                    // Esperar por algum elemento que indique que a página da AEGEA carregou
                    await page.WaitForSelectorAsync("iframe[src*='SupplierFrame'], #SupplierFrame, iframe[name='SupplierFrame']",
                        new PageWaitForSelectorOptions { Timeout = 8000 });
                    Console.WriteLine("  ✅ Frame SupplierFrame detectado!");
                }
                catch
                {
                    Console.WriteLine("  ⚠ Frame SupplierFrame não encontrado, mas continuando...");
                }

                Console.WriteLine("  ✅ Navegação para AEGEA concluída!");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ❌ Erro durante navegação: {ex.Message}");
                return false;
            }
        }

        public async Task<IFrame> ObterFrameSupplierSeguroAsync(IPage page)
        {
            try
            {
                IFrame frame = page.Frame("SupplierFrame");
                if (frame != null)
                {
                    Console.WriteLine("   ✅ Frame encontrado pelo nome");
                    return frame;
                }

                frame = page.Frames.FirstOrDefault(f =>
                    f.Url?.Contains("SupplierFrame", StringComparison.OrdinalIgnoreCase) == true);
                if (frame != null)
                {
                    Console.WriteLine("   ✅ Frame encontrado pela URL");
                    return frame;
                }

                if (page.Frames.Count > 1)
                {
                    frame = page.Frames[1];
                    Console.WriteLine($"   ✅ Frame encontrado pelo índice [1]");
                    return frame;
                }

                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠ Erro ao obter frame: {ex.Message}");
                return null;
            }
        }
        private string ObterNomeEmpresa(string chaveConta)
        {
            // Esta função retorna apenas o nome da EMPRESA (Ventura, Aliança ou União)
            // Independente se é AEGEA ou ESTÁCIO
            return chaveConta switch
            {
                "1" => "ALIANÇA",  // ALIANÇA ESTÁCIO → ALIANÇA
                "2" => "ALIANÇA",  // ALIANÇA AEGEA → ALIANÇA  
                "3" => "VENTURA",  // VENTURA AEGEA/ESTACIO → VENTURA
                "4" => "UNIÃO",    // UNIÃO AEGEA/ESTÁCIO → UNIÃO
                _ => $"Empresa {chaveConta}"
            };
        }
        private string ObterClienteParaPlanilha(string empresaAtual)
        {
            // Esta função retorna se é AEGEA ou ESTÁCIO como cliente
            return empresaAtual == "ESTÁCIO" ? "ESTÁCIO" : "AEGEA";
        }

        public bool VerificarSeCotacaoJaFoiProcessada(string numeroCotacao)
        {
            if (string.IsNullOrEmpty(numeroCotacao))
                return false;

            return _excelDb.CotacaoJaExiste(numeroCotacao);
        }



        public async Task<bool> RegistrarCotacaoNoBancoDadosAsync(
    string numeroCotacao,
    string dataVencimento,
    string horarioVencimento,
    List<ItemCotacao> itens,
    string contaChave,
    string empresaAtual)
        {
            try
            {
                Console.WriteLine($"\n🗄️  REGISTRANDO COTAÇÃO NO BANCO DE DADOS...");

                // Obter nomes corretos baseados na conta e empresa atual
                string nomeEmpresa = ObterNomeEmpresa(contaChave);
                string cliente = ObterClienteParaPlanilha(empresaAtual);
                string empresaParaPlanilha = $"{nomeEmpresa} {cliente}";
                string portal = "ARIBA";

                Console.WriteLine($"   🔢 Número: {numeroCotacao}");
                Console.WriteLine($"   🏢 Conta: {contaChave}");
                Console.WriteLine($"   🔧 Empresa atual: {nomeEmpresa}");
                Console.WriteLine($"   📋 Para planilha: {nomeEmpresa}");
                Console.WriteLine($"   👥 Cliente: {cliente}");
                Console.WriteLine($"   📅 Vencimento: {dataVencimento} às {horarioVencimento}");
                Console.WriteLine($"   📦 Total de itens: {itens.Count}");

                // FORMATAR PRODUTOS COMO SOLICITADO: "Primeiro Item (X Itens)"
                string produtosParaPlanilha = "";

                if (itens.Count > 0)
                {
                    // Pegar o primeiro item (descrição limpa se disponível, senão a original)
                    string primeiroItem = !string.IsNullOrWhiteSpace(itens[0].DescricaoLimpa)
                        ? itens[0].DescricaoLimpa.Trim()
                        : itens[0].DescricaoOriginal.Trim();

                    // Se o primeiro item for muito longo, truncar
                    if (primeiroItem.Length > 50)
                    {
                        primeiroItem = primeiroItem.Substring(0, 50) + "...";
                    }

                    // Formatar como "Primeiro Item (X Itens)"
                    produtosParaPlanilha = $"{primeiroItem} ({itens.Count} Itens)";

                    Console.WriteLine($"   📝 Produtos formatados: {produtosParaPlanilha}");

                    // Log detalhado para debugging
                    Console.WriteLine($"   📋 Itens encontrados:");
                    for (int i = 0; i < Math.Min(itens.Count, 5); i++) // Mostrar até 5 itens
                    {
                        var item = itens[i];
                        string descricao = !string.IsNullOrWhiteSpace(item.DescricaoLimpa)
                            ? item.DescricaoLimpa
                            : item.DescricaoOriginal;
                        Console.WriteLine($"     {i + 1}. {descricao}");
                    }
                    if (itens.Count > 5)
                    {
                        Console.WriteLine($"     ... e mais {itens.Count - 5} itens");
                    }
                }
                else
                {
                    produtosParaPlanilha = "Nenhum item identificado";
                    Console.WriteLine($"   ⚠️ Nenhum item encontrado para formatar");
                }

                // Adicionar na planilha com o formato novo
                bool sucesso = _excelDb.AdicionarCotacaoNaPlanilha(
                    numeroCotacao: numeroCotacao,
                    portal: portal,
                    cliente: cliente,
                    dataVencimento: dataVencimento,
                    horarioVencimento: horarioVencimento,
                    produtos: produtosParaPlanilha, // NOVO FORMATO
                    empresaResposta: nomeEmpresa,
                    vendedor: "" // Vazio até ser distribuído
                );

                if (sucesso)
                {
                    Console.WriteLine($"   ✅ Cotação registrada no banco de dados!");
                    Console.WriteLine($"   📋 Formato: {produtosParaPlanilha}");

                    // Criar log do registro (com quantidades) - AJUSTADO PARA 6 PARÂMETROS
                    await CriarLogRegistroCotacaoAsync(
                        numeroCotacao,
                        nomeEmpresa,
                        itens,
                        produtosParaPlanilha,
                        contaChave,
                        cliente);

                    return true;
                }
                else
                {
                    Console.WriteLine($"   ❌ Falha ao registrar no banco de dados!");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ❌ ERRO no registro: {ex.Message}");
                Console.WriteLine($"   📌 StackTrace: {ex.StackTrace}");
                return false;
            }
        }

        private async Task CriarLogRegistroCotacaoAsync(
    string numeroCotacao,
    string empresaPlanilha,
    List<ItemCotacao> itens,
    string produtosComQuantidade,
    string nomeConta,
    string cliente) // ADICIONADO: cliente
        {
            try
            {
                string logFileName = $"Registro_{numeroCotacao}_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                string logPath = Path.Combine("Cotacoes_Ariba", "Logs_Registros", logFileName);

                Directory.CreateDirectory(Path.GetDirectoryName(logPath));

                StringBuilder log = new StringBuilder();
                log.AppendLine("=".PadRight(60, '='));
                log.AppendLine("📋 LOG DE REGISTRO DE COTAÇÃO");
                log.AppendLine("=".PadRight(60, '='));
                log.AppendLine($"Número: {numeroCotacao}");
                log.AppendLine($"Conta: {nomeConta}");
                log.AppendLine($"Empresa (Planilha): {empresaPlanilha}");
                log.AppendLine($"Cliente: {cliente}");
                log.AppendLine($"Data/Hora Registro: {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
                log.AppendLine();
                log.AppendLine("📦 ITENS DA COTAÇÃO:");
                log.AppendLine("-".PadRight(60, '-'));

                foreach (var item in itens)
                {
                    string unidade = string.IsNullOrWhiteSpace(item.Unidade) ? "unidade" : item.Unidade;
                    log.AppendLine($"• {item.DescricaoOriginal}");
                    log.AppendLine($"  Quantidade: {item.Quantidade} {unidade}");
                    log.AppendLine($"  Descrição Limpa: {item.DescricaoLimpa}");
                    log.AppendLine();
                }

                log.AppendLine("=".PadRight(60, '='));
                log.AppendLine("✅ REGISTRO CONCLUÍDO COM SUCESSO");
                log.AppendLine("=".PadRight(60, '='));

                await File.WriteAllTextAsync(logPath, log.ToString(), Encoding.UTF8);
                Console.WriteLine($"   📋 Log salvo: {logFileName}");

                // Também salvar em JSON para facilitar processamento futuro
                await SalvarItensJsonAsync(numeroCotacao, itens, empresaPlanilha, nomeConta);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️ Erro ao criar log: {ex.Message}");
            }
        }

        public async Task<bool> BaixarDocumentoImpressaoAsync(IPage page, string numeroCotacao, string contaChave, string empresa)
        {
            try
            {
                Console.WriteLine($"\n" + new string('=', 60));
                Console.WriteLine($"BAIXANDO DOCUMENTO DA COTAÇÃO {numeroCotacao}");
                Console.WriteLine(new string('=', 60));

                // Configurar o download ANTES de clicar
                var downloadTask = page.WaitForDownloadAsync();

                Console.WriteLine($"🔍 Buscando botão de impressão...");

                // Lista de seletores em ordem de prioridade - CORRIGIDO
                var seletores = new[]
                {
            "button[title=\"Imprimir informações do evento em um documento\"]",  // Título exato
            "button:has-text('Imprimir informações do evento')",  // Texto do botão
            "button.w-btn.aw7_w-btn",  // Classes
            "button[id='_1jn$t']",  // ID com $ escapado
            "button.w-btn:has-text('Imprimir')",
            "[class*='print']",
            "button[onclick*='print']"
        };

                ILocator botaoImprimir = null;
                bool encontrado = false;

                foreach (var seletor in seletores)
                {
                    try
                    {
                        Console.WriteLine($"   🔍 Tentando seletor: {seletor}");
                        botaoImprimir = page.Locator(seletor);

                        var count = await botaoImprimir.CountAsync();
                        if (count > 0)
                        {
                            Console.WriteLine($"   ✅ Botão encontrado com: {seletor}");
                            encontrado = true;
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ⚠️ Erro com seletor '{seletor}': {ex.Message}");
                        continue;
                    }
                }

                if (!encontrado)
                {
                    Console.WriteLine($"❌ Nenhum botão de impressão encontrado");

                    // DEBUG adicional
                    Console.WriteLine($"\n🔍 DEBUG: Procurando por texto 'Imprimir'...");
                    var elementosComTexto = page.Locator(":has-text('Imprimir'), :has-text('imprimir')");
                    var total = await elementosComTexto.CountAsync();
                    Console.WriteLine($"   Elementos com 'Imprimir': {total}");

                    for (int i = 0; i < Math.Min(total, 5); i++)
                    {
                        try
                        {
                            var elemento = elementosComTexto.Nth(i);
                            var tagName = await elemento.EvaluateAsync<string>("el => el.tagName.toLowerCase()");
                            var texto = (await elemento.TextContentAsync() ?? "").Trim();
                            var titulo = await elemento.GetAttributeAsync("title") ?? "";

                            Console.WriteLine($"   [{i}] <{tagName}> Texto='{texto}', Título='{titulo}'");

                            if (tagName == "button" && texto.Contains("Imprimir"))
                            {
                                Console.WriteLine($"   ⭐ Botão manual identificado, tentando clicar...");
                                await elemento.ClickAsync();
                                encontrado = true;
                                break;
                            }
                        }
                        catch { }
                    }

                    if (!encontrado)
                    {
                        return false;
                    }
                }

                // Se encontramos o botão, clicar
                if (encontrado && botaoImprimir != null)
                {
                    Console.WriteLine($"🖱️ Clicando no botão...");
                    try
                    {
                        await botaoImprimir.First.ClickAsync(new LocatorClickOptions
                        {
                            Force = true,
                            Timeout = 10000
                        });
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"⚠️ Erro ao clicar: {ex.Message}");
                        return false;
                    }
                }

                Console.WriteLine($"✅ Botão clicado!");
                Console.WriteLine($"⏳ Aguardando download...");

                try
                {
                    // Aguardar download com timeout
                    var download = await downloadTask.WaitAsync(TimeSpan.FromSeconds(30));

                    string nomeArquivo = download.SuggestedFilename;
                    Console.WriteLine($"\n📥 Download iniciado: {nomeArquivo}");

                    // Criar diretório para a empresa
                    Console.WriteLine($"   🏢 Empresa: {empresa}");
                    string pastaDownloads = Path.Combine("Cotacoes_Ariba", "Downloads", empresa);
                    Directory.CreateDirectory(pastaDownloads);

                    // Definir caminho
                    string caminhoCompleto = Path.Combine(pastaDownloads, nomeArquivo);

                    // Salvar arquivo
                    await download.SaveAsAsync(caminhoCompleto);

                    if (File.Exists(caminhoCompleto))
                    {
                        var fileInfo = new FileInfo(caminhoCompleto);
                        Console.WriteLine($"✅ Download concluído!");
                        Console.WriteLine($"   📍 Local: {caminhoCompleto}");
                        Console.WriteLine($"   📊 Tamanho: {fileInfo.Length} bytes");

                        // Processar arquivo
                        await ProcessarArquivoWordComCabecalhoAsync(caminhoCompleto, numeroCotacao, empresa, contaChave);

                        return true;
                    }

                    return false;
                }
                catch (TimeoutException)
                {
                    Console.WriteLine($"⏱️ Timeout aguardando download");
                    return false;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ Erro no download: {ex.Message}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ ERRO no download: {ex.Message}");
                return false;
            }
        }
        private async Task<bool> ProcessarArquivoWordComCabecalhoAsync(string caminhoArquivo, string numeroCotacao, string empresaAtual, string contaChave)
        {
            try
            {
                Console.WriteLine($"\n📝 PROCESSANDO ARQUIVO WORD COM CABEÇALHO CORRETO...");
                Console.WriteLine($"   📄 Arquivo: {Path.GetFileName(caminhoArquivo)}");

                // Ler o conteúdo do arquivo
                string conteudo = await File.ReadAllTextAsync(caminhoArquivo, Encoding.UTF8);

                // Obter nomes corretos
                string nomeConta = ObterNomeContaCorreto(contaChave);
                string empresaPlanilha = ObterEmpresaParaPlanilha(contaChave, empresaAtual);

                Console.WriteLine($"   🏢 Conta: {nomeConta}");
                Console.WriteLine($"   📋 Empresa na planilha: {empresaPlanilha}");
                Console.WriteLine($"   🔧 Empresa atual: {empresaAtual}");

                // Verificar se já tem o cabeçalho personalizado
                if (conteudo.Contains("INFORMAÇÕES DA COTAÇÃO") && conteudo.Contains("Documento processado automaticamente"))
                {
                    Console.WriteLine($"   ✅ Arquivo já possui cabeçalho personalizado");
                    return true;
                }

                // Criar o cabeçalho HTML personalizado com informações corretas
                string cabecalhoHtml = $@"
<!-- ============================================================ -->
<!-- CABEÇALHO DE INFORMAÇÕES DA COTAÇÃO - ADICIONADO AUTOMATICAMENTE -->
<!-- ============================================================ -->

<div class=""MsoNormal"" style=""margin-bottom:12.0pt"">
  <table class=""MsoNormalTable"" border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" style=""width:100.0%;border-collapse:collapse"">
    <tr style=""height:15.75pt"">
      <td width=""100%"" colspan=""2"" style=""width:100.0%;border:solid windowtext 1.0pt;background:#4472C4;padding:0cm 5.4pt 0cm 5.4pt;height:15.75pt"">
        <p class=""MsoNormal"" style=""margin-bottom:0cm""><b><span style=""font-size:12.0pt;font-family:&quot;Calibri&quot;,sans-serif;color:white"">INFORMAÇÕES DA COTAÇÃO</span></b></p>
      </td>
    </tr>
    <tr style=""height:15.75pt"">
      <td width=""25%"" style=""width:25.0%;border:solid windowtext 1.0pt;border-top:none;background:#D9E2F3;padding:0cm 5.4pt 0cm 5.4pt;height:15.75pt"">
        <p class=""MsoNormal"" style=""margin-bottom:0cm""><b><span style=""font-size:10.0pt;font-family:&quot;Calibri&quot;,sans-serif;color:#1F3864"">Cotação:</span></b></p>
      </td>
      <td width=""75%"" style=""width:75.0%;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#E7E6E6;padding:0cm 5.4pt 0cm 5.4pt;height:15.75pt"">
        <p class=""MsoNormal"" style=""margin-bottom:0cm""><span style=""font-size:10.0pt;font-family:&quot;Calibri&quot;,sans-serif"">{numeroCotacao}</span></p>
      </td>
    </tr>
    <tr style=""height:15.75pt"">
      <td width=""25%"" style=""width:25.0%;border:solid windowtext 1.0pt;border-top:none;background:#D9E2F3;padding:0cm 5.4pt 0cm 5.4pt;height:15.75pt"">
        <p class=""MsoNormal"" style=""margin-bottom:0cm""><b><span style=""font-size:10.0pt;font-family:&quot;Calibri&quot;,sans-serif;color:#1F3864"">Cliente:</span></b></p>
      </td>
      <td width=""75%"" style=""width:75.0%;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#E7E6E6;padding:0cm 5.4pt 0cm 5.4pt;height:15.75pt"">
        <p class=""MsoNormal"" style=""margin-bottom:0cm""><span style=""font-size:10.0pt;font-family:&quot;Calibri&quot;,sans-serif"">{empresaAtual}</span></p>
      </td>
    </tr>
    <tr style=""height:15.75pt"">
      <td width=""25%"" style=""width:25.0%;border:solid windowtext 1.0pt;border-top:none;background:#D9E2F3;padding:0cm 5.4pt 0cm 5.4pt;height:15.75pt"">
        <p class=""MsoNormal"" style=""margin-bottom:0cm""><b><span style=""font-size:10.0pt;font-family:&quot;Calibri&quot;,sans-serif;color:#1F3864"">Fornecedor:</span></b></p>
      </td>
      <td width=""75%"" style=""width:75.0%;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#E7E6E6;padding:0cm 5.4pt 0cm 5.4pt;height:15.75pt"">
        <p class=""MsoNormal"" style=""margin-bottom:0cm""><span style=""font-size:10.0pt;font-family:&quot;Calibri&quot;,sans-serif"">{nomeConta}</span></p>
      </td>
    </tr>
    <tr style=""height:15.75pt"">
      <td width=""25%"" style=""width:25.0%;border:solid windowtext 1.0pt;border-top:none;background:#D9E2F3;padding:0cm 5.4pt 0cm 5.4pt;height:15.75pt"">
        <p class=""MsoNormal"" style=""margin-bottom:0cm""><b><span style=""font-size:10.0pt;font-family:&quot;Calibri&quot;,sans-serif;color:#1F3864"">Número:</span></b></p>
      </td>
      <td width=""75%"" style=""width:75.0%;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#E7E6E6;padding:0cm 5.4pt 0cm 5.4pt;height:15.75pt"">
        <p class=""MsoNormal"" style=""margin-bottom:0cm""><span style=""font-size:10.0pt;font-family:&quot;Calibri&quot;,sans-serif"">{numeroCotacao}</span></p>
      </td>
    </tr>
    <tr style=""height:15.75pt"">
      <td width=""25%"" style=""width:25.0%;border:solid windowtext 1.0pt;border-top:none;background:#D9E2F3;padding:0cm 5.4pt 0cm 5.4pt;height:15.75pt"">
        <p class=""MsoNormal"" style=""margin-bottom:0cm""><b><span style=""font-size:10.0pt;font-family:&quot;Calibri&quot;,sans-serif;color:#1F3864"">Processado em:</span></b></p>
      </td>
      <td width=""75%"" style=""width:75.0%;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;background:#E7E6E6;padding:0cm 5.4pt 0cm 5.4pt;height:15.75pt"">
        <p class=""MsoNormal"" style=""margin-bottom:0cm""><span style=""font-size:10.0pt;font-family:&quot;Calibri&quot;,sans-serif"">{DateTime.Now:dd/MM/yyyy HH:mm:ss}</span></p>
      </td>
    </tr>
    <tr style=""height:15.0pt"">
      <td width=""100%"" colspan=""2"" style=""width:100.0%;border:solid windowtext 1.0pt;border-top:none;background:#F2F2F2;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt"">
        <p class=""MsoNormal"" align=""center"" style=""margin-bottom:0cm;text-align:center""><i><span style=""font-size:9.0pt;font-family:&quot;Calibri&quot;,sans-serif;color:#7F7F7F"">Documento processado automaticamente pelo Sistema de Cotações Ariba</span></i></p>
      </td>
    </tr>
  </table>
</div>

<p class=""MsoNormal"" style=""margin-bottom:12.0pt"">&nbsp;</p>

<!-- ============================================================ -->
<!-- FIM DO CABEÇALHO - CONTEÚDO ORIGINAL DO DOCUMENTO ABAIXO -->
<!-- ============================================================ -->
";

                // Encontrar posição para inserir
                int posicaoInsercao;
                if (conteudo.Contains("<div class=\"WordSection1\">"))
                {
                    posicaoInsercao = conteudo.IndexOf("<div class=\"WordSection1\">") + "<div class=\"WordSection1\">".Length;
                }
                else if (conteudo.Contains("<body>"))
                {
                    posicaoInsercao = conteudo.IndexOf("<body>") + "<body>".Length;
                }
                else
                {
                    posicaoInsercao = conteudo.IndexOf("</head>") + "</head>".Length;
                }

                // Inserir o cabeçalho
                string novoConteudo = conteudo.Insert(posicaoInsercao, cabecalhoHtml);

                // Atualizar título do documento
                string novoTitle = $"<title>Cotação {numeroCotacao} - {empresaPlanilha} - Processado em {DateTime.Now:dd/MM/yyyy HH:mm}</title>";

                if (conteudo.Contains("<title>"))
                {
                    int inicioTitle = novoConteudo.IndexOf("<title>");
                    int fimTitle = novoConteudo.IndexOf("</title>") + "</title>".Length;
                    if (inicioTitle >= 0 && fimTitle > inicioTitle)
                    {
                        string tituloAtual = novoConteudo.Substring(inicioTitle, fimTitle - inicioTitle);
                        novoConteudo = novoConteudo.Replace(tituloAtual, novoTitle);
                    }
                }

                // Salvar o arquivo
                await File.WriteAllTextAsync(caminhoArquivo, novoConteudo, Encoding.UTF8);

                Console.WriteLine($"   ✅ Cabeçalho adicionado com informações corretas:");
                Console.WriteLine($"      📋 Cotação: {empresaPlanilha}/{numeroCotacao}");
                Console.WriteLine($"      👥 Fornecedor: {nomeConta}");
                Console.WriteLine($"      🏢 Cliente: {empresaAtual}");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ❌ Erro ao processar arquivo Word: {ex.Message}");
                return false;
            }
        }
        private async Task SalvarItensJsonAsync(
    string numeroCotacao,
    List<ItemCotacao> itens,
    string empresaPlanilha,
    string nomeConta)
        {
            try
            {
                var jsonData = new
                {
                    NumeroCotacao = numeroCotacao,
                    EmpresaPlanilha = empresaPlanilha,
                    Conta = nomeConta,
                    DataRegistro = DateTime.Now,
                    TotalItens = itens.Count,
                    Itens = itens.Select(i => new
                    {
                        i.DescricaoOriginal,
                        i.DescricaoLimpa,
                        i.Quantidade,
                        i.Unidade
                    }).ToArray()
                };

                string json = System.Text.Json.JsonSerializer.Serialize(jsonData, new System.Text.Json.JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });

                string jsonFileName = $"Itens_{numeroCotacao}_{DateTime.Now:yyyyMMdd_HHmmss}.json";
                string jsonPath = Path.Combine("Cotacoes_Ariba", "Logs_Registros", jsonFileName);

                await File.WriteAllTextAsync(jsonPath, json, Encoding.UTF8);
                Console.WriteLine($"   📄 JSON salvo: {jsonFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️ Erro ao salvar JSON: {ex.Message}");
            }
        }

        private string LimitarTexto(string texto, int maxLength)
        {
            if (string.IsNullOrEmpty(texto)) return texto;
            return texto.Length <= maxLength ? texto : texto.Substring(0, maxLength) + "...";
        }
        private string ObterNomeContaCorreto(string chaveConta)
        {
            // Mapeamento correto baseado no dicionário Contas
            return chaveConta switch
            {
                "1" => "ALIANÇA",
                "2" => "ALIANÇA",
                "3" => "VENTURA",
                "4" => "UNIÃO",
                _ => $"Conta {chaveConta}"
            };
        }

        private string ObterEmpresaParaPlanilha(string contaChave, string empresaAtual)
        {
            // Determinar qual empresa colocar na planilha baseado na conta e empresa atual
            return contaChave switch
            {
                "1" => "ALIANÇA", // Conta 1 sempre é ALIANÇA
                "2" => "ALIANÇA",    // Conta 2 sempre é ALIANÇA
                "3" => "VENTURA",
                "4" => "UNIÃO",
                _ => empresaAtual
            };
        }
        public async Task FecharAsync()
        {
            if (_browser != null)
                await _browser.CloseAsync();
            _playwright?.Dispose();
        }
    }
}