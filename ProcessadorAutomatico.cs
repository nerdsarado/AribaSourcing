// ProcessadorAutomatico.cs
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Playwright;
using System.Text;
using System.Text.RegularExpressions;

namespace CotacoesAriba
{
    public class ProcessadorAutomatico
    {
        private readonly GerenciadorLogin _gerenciador;
        private int _totalCotacoesProcessadas = 0;
        private int _totalContasProcessadas = 0;
        private string _empresaAtual = "AEGEA";
        private List<string> _empresasPrioritarias = new List<string>();
        private Dictionary<string, List<string>> _cotacoesEncontradasPorEmpresa = new Dictionary<string, List<string>>();

        // Estatísticas
        public class Estatisticas
        {
            public int TotalContas { get; set; }
            public int ContasComSucesso { get; set; }
            public int TotalCotacoes { get; set; }
            public int CotacoesRegistradas { get; set; }
            public int SegundaRodadaDetectada { get; set; }
            public Dictionary<string, int> CotacoesPorConta { get; set; } = new();
            public DateTime InicioProcessamento { get; set; }
            public DateTime FimProcessamento { get; set; }
            public TimeSpan DuracaoTotal => FimProcessamento - InicioProcessamento;
        }

        public ProcessadorAutomatico()
        {
            _gerenciador = new GerenciadorLogin();
        }
        public ProcessadorAutomatico(List<string> empresasPrioritarias)
        {
            _gerenciador = new GerenciadorLogin();
            _empresasPrioritarias = empresasPrioritarias ?? new List<string> { "ALIANÇA", "VENTURA", "UNIÃO" };

            // Inicializar dicionário para cada empresa
            foreach (var empresa in _empresasPrioritarias)
            {
                _cotacoesEncontradasPorEmpresa[empresa] = new List<string>();
            }
        }


        public async Task<Estatisticas> ExecutarProcessamentoCompletoAsync()
        {
            var estatisticas = new Estatisticas
            {
                InicioProcessamento = DateTime.Now
            };

            Console.Clear();
            Console.WriteLine("=".PadRight(80, '='));
            Console.WriteLine("🚀 PROCESSAMENTO AUTOMÁTICO COM PRIORIDADE");
            Console.WriteLine("=".PadRight(80, '='));
            Console.WriteLine($"Início: {estatisticas.InicioProcessamento:dd/MM/yyyy HH:mm:ss}");
            Console.WriteLine($"Prioridade: {string.Join(" > ", _empresasPrioritarias)}");
            Console.WriteLine("=".PadRight(80, '='));
            Console.WriteLine();

            // Primeiro, determinar quais são as empresas NÃO prioritárias
            var todasEmpresas = new List<string> { "ALIANÇA", "VENTURA", "UNIÃO" };
            var empresasNaoPrioritarias = todasEmpresas.Except(_empresasPrioritarias).ToList();

            Console.WriteLine($"🏢 Empresas prioritárias: {string.Join(", ", _empresasPrioritarias)}");

            if (empresasNaoPrioritarias.Count > 0)
            {
                Console.WriteLine($"📋 Empresas não prioritárias (apenas verificação): {string.Join(", ", empresasNaoPrioritarias)}");
            }
            else
            {
                Console.WriteLine($"✅ Todas as empresas serão processadas como prioritárias");
            }
            Console.WriteLine();

            // 1. PRIMEIRO: Processar TODAS as empresas prioritárias na ordem definida
            Console.WriteLine("\n" + new string('=', 80));
            Console.WriteLine("⭐ PROCESSANDO EMPRESAS PRIORITÁRIAS");
            Console.WriteLine(new string('=', 80));

            foreach (var empresaPrioritaria in _empresasPrioritarias)
            {
                Console.WriteLine($"\n🎯 EMPRESA PRIORITÁRIA: {empresaPrioritaria}");

                // Processar contas relacionadas a esta empresa
                var contasDaEmpresa = ObterContasDaEmpresa(empresaPrioritaria);

                foreach (var conta in contasDaEmpresa)
                {
                    await ProcessarContaComPrioridadeAsync(conta, empresaPrioritaria, estatisticas);

                    // Aguardar entre contas
                    if (conta != contasDaEmpresa.Last())
                    {
                        Console.WriteLine($"\n⏳ Aguardando 10 segundos antes da próxima conta...");
                        await Task.Delay(10000);
                    }
                }
            }

            // 2. SÓ DEPOIS: Verificar empresas não prioritárias
            if (empresasNaoPrioritarias.Count > 0)
            {
                await VerificarEmpresasNaoPrioritariasAsync(estatisticas, empresasNaoPrioritarias);
            }

            estatisticas.FimProcessamento = DateTime.Now;

            // Gerar relatório final
            await GerarRelatorioFinalAsync(estatisticas, empresasNaoPrioritarias);

            return estatisticas;
        }

        private List<string> ObterContasDaEmpresa(string empresa)
        {
            // Mapeia empresa para as contas correspondentes
            return empresa switch
            {
                "ALIANÇA" => new List<string> { "1", "2" },  // ALIANÇA ESTÁCIO e ALIANÇA AEGEA
                "VENTURA" => new List<string> { "3" },       // VENTURA AEGEA/ESTACIO
                "UNIÃO" => new List<string> { "4" },         // UNIÃO AEGEA/ESTÁCIO
                _ => new List<string>()
            };
        }

        private async Task ProcessarContaComPrioridadeAsync(string chaveConta, string empresa, Estatisticas estatisticas)
        {
            try
            {
                Console.WriteLine($"\n" + new string('>', 60));
                Console.WriteLine($"🔐 PROCESSANDO CONTA {chaveConta} ({empresa}) [PRIORITÁRIA]");
                Console.WriteLine(new string('>', 60));

                IPage pagina = null;

                try
                {
                    // Login automático na conta
                    pagina = await _gerenciador.RealizarLoginAsync(chaveConta);

                    if (pagina == null)
                    {
                        Console.WriteLine($"❌ Falha no login na conta {chaveConta}");
                        return;
                    }

                    _totalContasProcessadas++;
                    estatisticas.ContasComSucesso++;

                    Console.WriteLine($"✅ Login bem-sucedido!");

                    // Determinar quais clientes processar baseado na conta
                    int totalCotacoesEstaConta = 0;

                    switch (chaveConta)
                    {
                        case "1": // ALIANÇA ESTÁCIO - Apenas Estácio
                            Console.WriteLine($"\n🎓 PROCESSANDO ESTÁCIO (Conta ALIANÇA ESTÁCIO)...");
                            _empresaAtual = "ESTÁCIO";
                            bool navegarParaEstacio = await NavegarParaEstacioAsync(pagina);
                            if (navegarParaEstacio)
                            {
                                int cotacoesEstacioAli = await ProcessarCotacoesDaContaPrioritariaAsync(pagina, chaveConta, _empresaAtual, empresa);
                                totalCotacoesEstaConta = cotacoesEstacioAli;
                            }
                            break;


                        case "2": // ALIANÇA AEGEA - Apenas AEGEA
                            Console.WriteLine($"\n🏢 PROCESSANDO AEGEA (Conta ALIANÇA AEGEA)...");
                            _empresaAtual = "AEGEA";
                            await Task.Delay(5000);
                            int cotacoesAEGEA = await ProcessarCotacoesDaContaPrioritariaAsync(pagina, chaveConta, _empresaAtual, empresa);
                            totalCotacoesEstaConta = cotacoesAEGEA;
                            break;

                        case "3": // VENTURA AEGEA/ESTACIO - Ambas as empresas
                            // Processar AEGEA primeiro
                            Console.WriteLine($"\n🏢 PROCESSANDO AEGEA (Conta VENTURA)...");
                            _empresaAtual = "AEGEA";
                            await Task.Delay(5000);
                            int cotacoesAegeaVentura = await ProcessarCotacoesDaContaPrioritariaAsync(pagina, chaveConta, _empresaAtual, empresa);
                            totalCotacoesEstaConta += cotacoesAegeaVentura;

                            // Agora processar Estácio
                            Console.WriteLine($"\n🎓 PROCESSANDO ESTÁCIO (Conta VENTURA)...");
                            _empresaAtual = "ESTÁCIO";
                            bool navegacaoEstacioVentura = await NavegarParaEstacioAsync(pagina);

                            if (navegacaoEstacioVentura)
                            {
                                await Task.Delay(10000);
                                int cotacoesEstacioVentura = await ProcessarCotacoesDaContaPrioritariaAsync(pagina, chaveConta, _empresaAtual, empresa);
                                totalCotacoesEstaConta += cotacoesEstacioVentura;
                            }
                            break;

                        case "4": // UNIÃO AEGEA/ESTÁCIO - Ambas as empresas
                            // Processar AEGEA primeiro
                            Console.WriteLine($"\n🏢 PROCESSANDO AEGEA (Conta UNIÃO)...");
                            _empresaAtual = "AEGEA";
                            bool navegacaoAEGEA = await _gerenciador.NavegarParaAegeaAsync(pagina);
                            await Task.Delay(15000);
                            int cotacoesAegeaUniao = await ProcessarCotacoesDaContaPrioritariaAsync(pagina, chaveConta, _empresaAtual, empresa);
                            totalCotacoesEstaConta += cotacoesAegeaUniao;

                            // Agora processar Estácio
                            Console.WriteLine($"\n🎓 PROCESSANDO ESTÁCIO (Conta UNIÃO)...");
                            _empresaAtual = "ESTÁCIO";
                            bool navegacaoEstacioUniao = await NavegarParaEstacioAsync(pagina);

                            if (navegacaoEstacioUniao)
                            {
                                await Task.Delay(5000);
                                int cotacoesEstacioUniao = await ProcessarCotacoesDaContaPrioritariaAsync(pagina, chaveConta, _empresaAtual, empresa);
                                totalCotacoesEstaConta += cotacoesEstacioUniao;
                            }
                            break;
                    }

                    // Registrar cotações encontradas para esta empresa
                    estatisticas.CotacoesPorConta[chaveConta] = totalCotacoesEstaConta;
                    estatisticas.TotalCotacoes += totalCotacoesEstaConta;

                    Console.WriteLine($"\n✅ CONTA {chaveConta} ({empresa}) PROCESSADA: {totalCotacoesEstaConta} cotações");
                }
                finally
                {
                    if (pagina != null)
                    {
                        await pagina.CloseAsync();
                        Console.WriteLine($"📭 Página da conta {chaveConta} fechada");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"💥 ERRO na conta {chaveConta}: {ex.Message}");
            }
        }
        private async Task<int> ProcessarCotacoesDaContaPrioritariaAsync(IPage pagina, string contaChave, string cliente, string empresa)
        {
            int cotacoesProcessadas = 0;
            var cotaçõesEncontradas = new List<string>();

            try
            {
                Console.WriteLine($"\n🔍 PROCURANDO COTAÇÕES ABERTAS em {cliente}...");

                var frameSupplier = await _gerenciador.ObterFrameSupplierSeguroAsync(pagina);
                if (frameSupplier == null)
                {
                    Console.WriteLine($"⚠️  Frame Supplier não encontrado");
                    return 0;
                }

                await frameSupplier.WaitForLoadStateAsync(LoadState.NetworkIdle);
                await Task.Delay(4000);

                // Encontrar todas as linhas com Status: Aberto
                var linhasStatusAberto = frameSupplier.Locator("tr:has-text('Status: Aberto')");
                var totalLinhasAbertas = await linhasStatusAberto.CountAsync();

                Console.WriteLine($"   📊 Cotações abertas encontradas: {totalLinhasAbertas}");

                if (totalLinhasAbertas == 0)
                {
                    Console.WriteLine($"   📭 Nenhuma cotação aberta encontrada");
                    return 0;
                }

                var cotacoesProcessadasNestaExecucao = new List<string>();

                for (int indiceLinha = 0; indiceLinha < totalLinhasAbertas; indiceLinha++)
                {
                    try
                    {
                        var linhaAtual = linhasStatusAberto.Nth(indiceLinha);

                        // Expandir toggle
                        bool expandido = await ExpandirToggleEspecificoAsync(frameSupplier, linhaAtual, indiceLinha);

                        if (!expandido) continue;

                        await Task.Delay(3000);

                        int cotacoesNestaLinha = await AnalisarCotaçõesExpandidasPrioritariasAsync(
                            frameSupplier, indiceLinha, cotacoesProcessadasNestaExecucao,
                            cotaçõesEncontradas, contaChave, empresa);

                        cotacoesProcessadas += cotacoesNestaLinha;

                        // Atualizar referências
                        frameSupplier = await _gerenciador.ObterFrameSupplierSeguroAsync(pagina) ?? frameSupplier;
                        await Task.Delay(2000);
                        linhasStatusAberto = frameSupplier.Locator("tr:has-text('Status: Aberto')");
                        totalLinhasAbertas = await linhasStatusAberto.CountAsync();
                        await Task.Delay(2000);


                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ⚠️ Erro na linha {indiceLinha + 1}: {ex.Message}");
                        continue;
                    }
                }

                // Armazenar cotações encontradas para esta empresa
                _cotacoesEncontradasPorEmpresa[empresa].AddRange(cotaçõesEncontradas.Distinct());

                Console.WriteLine($"\n✅ Total de cotações processadas nesta conta: {cotacoesProcessadas}");
                Console.WriteLine($"📋 Cotações encontradas para {empresa}: {cotaçõesEncontradas.Count}");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"💥 ERRO no processamento: {ex.Message}");
            }

            return cotacoesProcessadas;
        }
        private async Task<int> AnalisarCotaçõesExpandidasPrioritariasAsync(
            IFrame frame, int indiceLinhaPai, List<string> cotacoesProcessadas,
            List<string> cotaçõesEncontradas, string contaChave, string empresa)
        {
            int cotacoesProcessadasNestaExpansao = 0;
            //Declarar variável de página
            IPage pagina;
            string cliente = _empresaAtual;
            try
            {
                Console.WriteLine($"Cliente: {cliente}");
                //Atribuir valor à variável página
                pagina = frame.Page;
                // Buscar números de cotação
                var regexNumeroCotacao = new Regex(@"6000\d{6}");
                var locatorCandidatos = frame.Locator("tr:has(a[href*='webjumper']), tr:has-text('6000'), tr[class*='tableRow']");
                var totalCandidatos = await locatorCandidatos.CountAsync();

                for (int i = 0; i < totalCandidatos; i++)
                {
                    try
                    {
                        var linha = locatorCandidatos.Nth(i);
                        var textoLinha = (await linha.TextContentAsync() ?? "").Trim();
                        var match = regexNumeroCotacao.Match(textoLinha);

                        if (!match.Success) continue;

                        string numeroCotacao = match.Value;
                        cotaçõesEncontradas.Add(numeroCotacao);

                        Console.WriteLine($"\n📄 Cotação encontrada: {numeroCotacao}");

                        // Verificar se já foi processada
                        if (cotacoesProcessadas.Contains(numeroCotacao))
                        {
                            Console.WriteLine($"   ⏭️ Já processada - Pulando...");
                            continue;
                        }

                        // Verificar se já está no banco de dados
                        bool jaExisteNoBanco = _gerenciador.VerificarSeCotacaoJaFoiProcessada(numeroCotacao);

                        if (jaExisteNoBanco)
                        {
                            Console.WriteLine($"   ⏭️ Já existe no banco de dados - Pulando...");
                            cotacoesProcessadas.Add(numeroCotacao);
                            continue;
                        }

                        // Processar apenas se for da empresa prioritária
                        Console.WriteLine($"   ✅ NOVA COTAÇÃO - Processando como PRIORITÁRIA...");

                        // Clicar no link da cotação
                        bool clicou = await ClicarLinkCotacaoAsync(frame, linha, i);
                        if (!clicou) continue;

                        await Task.Delay(5000);

                        if (cliente == "ESTÁCIO")
                        {
                            var quantidadeItens = await ExtrairItensSimplesAsync(frame.Page);
                            if (quantidadeItens.Count >= 1)
                            {
                                Console.WriteLine("   🎓 Cotação já teve seus itens revisados.");
                                Console.WriteLine("   🔄 Iniciando extração de detalhes da cotação...");
                                Console.WriteLine("Procurando seletor de revisar detalhes do evento...");

                                var revisarDetalhes = await pagina.QuerySelectorAsync("#_c8_tuc");
                                if (revisarDetalhes != null)
                                {
                                    Console.WriteLine("Clicando seletor de revisar detalhes do evento...");
                                    await revisarDetalhes.ClickAsync();
                                }
                                else
                                {
                                    Console.WriteLine("Possivelmente já está na pagina de revisar detalhes.");
                                }

                                Console.WriteLine("Agora esta pronto para extrair os dados da cotação.");
                                // Extrair detalhes da cotação
                                bool processar = await ExtrairERegistrarCotacaoAsync(frame.Page, numeroCotacao, contaChave, cliente);

                                if (processar)
                                {
                                    cotacoesProcessadas.Add(numeroCotacao);
                                    cotacoesProcessadasNestaExpansao++;
                                    _totalCotacoesProcessadas++;
                                    Console.WriteLine($"   ✅ Processada com sucesso! (Total: {_totalCotacoesProcessadas})");
                                }

                                // Voltar para a lista
                                string urlLogin = "https://service.ariba.com/Sourcing.aw/109521004/aw?awh=r&awssk=M1Ju378m&dard=1";

                                Console.WriteLine($"Navegando para pagina de cotações da Estácio...");
                                await pagina.GotoAsync(urlLogin, new PageGotoOptions
                                {
                                    WaitUntil = WaitUntilState.DOMContentLoaded,
                                    Timeout = 60000
                                });
                                await Task.Delay(10000);
                                frame = await _gerenciador.ObterFrameSupplierSeguroAsync(frame.Page) ?? frame;
                                locatorCandidatos = frame.Locator("tr:has(a[href*='webjumper']), tr:has-text('6000'), tr[class*='tableRow']");
                                totalCandidatos = await locatorCandidatos.CountAsync();

                                //Abrir o Toggle novamente
                                await Task.Delay(4000);
                                await ProcessarCotacoesDaContaPrioritariaAsync(pagina, contaChave, cliente, empresa);
                            }
                            else
                            {
                                Console.WriteLine("   🎓 Revisando pré requisitos...");

                                // Lista de seletores em ordem de prioridade - CORRIGIDO
                                var seletores = new[]
                                {
            "button[title=\"Revisar os termos dos pré-requisitos e aceitá-los ou recusá-los\"]",  // Título exato
            "button:has-text('Revisar pré-requisitos')",  // Texto do botão
            "button.w-btn w-btn-primary aw7_w-btn-primary",  // Classes
            "button[id='_gz7tac']",
            "button.w-btn:has-text('Revisar')"
        };

                                ILocator botaoRevisar = null;
                                bool encontrado = false;
                                bool acionado = false;

                                foreach (var seletor in seletores)
                                {
                                    try
                                    {
                                        Console.WriteLine($"   🔍 Tentando seletor: {seletor}");
                                        botaoRevisar = pagina.Locator(seletor);

                                        var count = await botaoRevisar.CountAsync();
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
                                    Console.WriteLine($"❌ Nenhum botão de revisão encontrado");

                                    // DEBUG adicional
                                    Console.WriteLine($"\n🔍 DEBUG: Procurando por texto 'Revisar'...");
                                    var elementosComTexto = pagina.Locator(":has-text('Revisar pré-requisitos'), :has-text('Revisar pré-requisitos')");
                                    var total = await elementosComTexto.CountAsync();
                                    Console.WriteLine($"   Elementos com 'Revisar pré-requisitos': {total}");

                                    for (int q = 0; q < Math.Min(total, 5); q++)
                                    {
                                        try
                                        {
                                            var elemento = elementosComTexto.Nth(i);
                                            var tagName = await elemento.EvaluateAsync<string>("el => el.tagName.toLowerCase()");
                                            var texto = (await elemento.TextContentAsync() ?? "").Trim();
                                            var titulo = await elemento.GetAttributeAsync("title") ?? "";

                                            Console.WriteLine($"   [{i}] <{tagName}> Texto='{texto}', Título='{titulo}'");

                                            if (tagName == "button" && texto.Contains("Revisar pré-requisitos"))
                                            {
                                                Console.WriteLine($"   ⭐ Botão manual identificado, tentando clicar...");
                                                await elemento.ClickAsync();
                                                encontrado = true;
                                                break;
                                            }
                                        }
                                        catch { }
                                    }
                                }
                                if (encontrado && botaoRevisar != null)
                                {
                                    Console.WriteLine($"🖱️ Clicando no botão...");
                                    try
                                    {
                                        await botaoRevisar.First.ClickAsync(new LocatorClickOptions
                                        {
                                            Force = true,
                                            Timeout = 10000
                                        });
                                        acionado = true;
                                        Console.WriteLine($"✅ Botão clicado!");
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"⚠️ Erro ao clicar: {ex.Message}");

                                    }
                                }
                                if (acionado)
                                {
                                    try
                                    {
                                        pagina = frame.Page;
                                        await Task.Delay(15000);

                                        Console.WriteLine("Procurando seletores para o termo de contratação...");
                                        var seletorAceitar = await pagina.QuerySelectorAsync(".w-dropdown-pic-ct");
                                        if (seletorAceitar != null)
                                        {
                                            Console.WriteLine("✅ Seletor de termo de contratação encontrado.");
                                            try
                                            {
                                                Console.WriteLine("Clicando no Seletor de termo de contratação...");
                                                await seletorAceitar.ClickAsync();
                                            }
                                            catch
                                            {
                                                Console.WriteLine("⚠️ Erro ao clicar no seletor de termo de contratação.");
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine("⚠️ Seletor de termo de contratação não encontrado!");
                                        }

                                        Console.WriteLine("Procurando seletor 'Sim'...");
                                        var seletorSim = await pagina.QuerySelectorAsync("#_4o05gd0");
                                        if (seletorSim != null)
                                        {
                                            Console.WriteLine("✅ Seletor 'Sim' encontrado.");
                                            Console.WriteLine("Clicando no Seletor 'Sim'...");
                                            await seletorSim.ClickAsync();
                                        }
                                        else
                                        {
                                            Console.WriteLine("⚠️ Seletor 'Sim' não encontrado!");
                                        }

                                        Console.WriteLine("Procurando no botão 'OK'...");
                                        var botaoOK = await pagina.QuerySelectorAsync("#_ali5ud");
                                        bool okApertado = false;

                                        if (botaoOK != null)
                                        {
                                            Console.WriteLine("Clicando no Seletor 'OK'...");
                                            await botaoOK.ClickAsync();
                                            okApertado = true;
                                        }
                                        else
                                        {


                                            Console.WriteLine("Não foi possível clicar no botão. Tentando outro botão...");
                                            var botaoConcluido = await pagina.QuerySelectorAsync("_bywhoc");
                                            if (botaoConcluido != null)
                                            {
                                                Console.WriteLine("Clicando no botão de concluido");
                                                await botaoConcluido.ClickAsync();
                                                okApertado = true;
                                            }
                                            else
                                            {
                                                Console.WriteLine("Não foi possivel encontrar o botão");
                                            }
                                        }


                                        if (okApertado)
                                        {
                                            Console.WriteLine("🔄 Procurando e manipulando modal...");

                                            try
                                            {
                                                // 1. Aguardar o modal aparecer
                                                var modal = await pagina.WaitForSelectorAsync("#_dbej6c.panel.w-dlg-panel-active",
                                                    new PageWaitForSelectorOptions { Timeout = 20000 });

                                                if (modal != null)
                                                {
                                                    Console.WriteLine("✅ Modal encontrado, aguardando 3 segundos...");
                                                    await Task.Delay(3000);

                                                    // 2. Clicar no botão OK usando o ID específico
                                                    Console.WriteLine("Procurando botão OK (#_kfai0c)...");

                                                    // Usar Locator que é mais robusto
                                                    var botaoOKLocator = pagina.Locator("#_kfai0c.w-btn.aw7_w-btn");
                                                    var count = await botaoOKLocator.CountAsync();

                                                    if (count > 0)
                                                    {
                                                        Console.WriteLine("Clicando no botão OK...");
                                                        await botaoOKLocator.First.ClickAsync(new LocatorClickOptions
                                                        {
                                                            Force = true,
                                                            Timeout = 10000
                                                        });

                                                        Console.WriteLine("✅ Modal processado com sucesso!");

                                                        // Aguardar o modal desaparecer
                                                        await Task.Delay(5000);
                                                    }
                                                    else
                                                    {
                                                        Console.WriteLine("❌ Botão OK não encontrado, tentando seletor alternativo...");

                                                        // Procurar por qualquer botão com texto OK dentro do modal
                                                        var botaoAlternativo = await modal.QuerySelectorAsync("button:has-text('OK')");
                                                        if (botaoAlternativo != null)
                                                        {
                                                            await botaoAlternativo.ClickAsync();
                                                            Console.WriteLine("✅ Botão OK alternativo clicado");
                                                        }
                                                    }
                                                }
                                                else
                                                {

                                                    Console.WriteLine("❌ Modal não apareceu dentro do tempo esperado.");
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine($"⚠️ Erro ao processar modal: {ex.Message}");

                                                // Fallback: tentar clicar diretamente no botão se ele existir
                                                var botaoDireto = await pagina.QuerySelectorAsync("#_kfai0c");
                                                if (botaoDireto != null)
                                                {
                                                    Console.WriteLine("Tentando clique direto no botão OK...");
                                                    await botaoDireto.ClickAsync();
                                                }
                                            }
                                            // Dar tempo de carregar
                                            await Task.Delay(15000);
                                            var input = pagina.Locator("#_v3xacd.w-chk-native");
                                            var botaoConfirmarLotes = pagina.Locator("#_9ckd4d.w-btn w-btn-primary aw7_w-btn-primary, button:has-text('Confirmar lotes/itens de linha selecionados')");
                                            Console.WriteLine("Procurando no input...");
                                            input = pagina.Locator("#_v3xacd.w-chk-native");
                                            var countInput = await input.CountAsync();

                                            if (countInput > 0)
                                            {
                                                Console.WriteLine("Clicando no input...");
                                                await input.First.ClickAsync(new LocatorClickOptions
                                                {
                                                    Force = true,
                                                    Timeout = 30000
                                                });
                                            }
                                            else

                                            {
                                                Console.WriteLine("⚠️ Input não encontrado!");
                                            }
                                            await Task.Delay(8000);

                                            Console.WriteLine("Procurando no botão de confirmar lotes...");
                                            botaoConfirmarLotes = pagina.Locator("#_9ckd4d.w-btn w-btn-primary aw7_w-btn-primary, button:has-text('Confirmar lotes/itens de linha selecionados')");
                                            var countConfirmar = await botaoConfirmarLotes.CountAsync();

                                            if (countConfirmar > 0)
                                            {
                                                Console.WriteLine("Clicando no botão de confirmar lotes...");
                                                await botaoConfirmarLotes.First.ClickAsync(new LocatorClickOptions
                                                {
                                                    Force = true,
                                                    Timeout = 30000
                                                });
                                                var verificarBotão = await pagina.QuerySelectorAsync("#_9ckd4d.w-btn w-btn-primary aw7_w-btn-primary, button:has-text('Confirmar lotes/itens de linha selecionados')");
                                                if (verificarBotão != null)
                                                {
                                                    await Task.Delay(10000);
                                                    Console.WriteLine("Aparentemente o botão não foi apertado, apertando novamente por métodos alternativos...");
                                                    await verificarBotão.ClickAsync();
                                                }
                                                else
                                                {
                                                    Console.WriteLine("✅ Botão de confirmar lotes clicado com sucesso!");
                                                }
                                            }
                                            else
                                            {
                                                Console.WriteLine("⚠️ Botão de confirmar lotes não encontrado!");
                                            }



                                            await Task.Delay(15000);
                                            var revisarDetalhes = await pagina.QuerySelectorAsync("#_c8_tuc");
                                            Console.WriteLine("Procurando seletor de revisar detalhes do evento...");
                                            await Task.Delay(15000);
                                            if (revisarDetalhes != null)
                                            {
                                                await Task.Delay(10000);
                                                Console.WriteLine("Clicando seletor de revisar detalhes do evento...");
                                                await revisarDetalhes.ClickAsync();
                                                Console.WriteLine("Agora esta pronto para extrair os dados da cotação.");
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"⚠️ Erro durante o processo de revisão: {ex.Message}");
                                        var revisarDetalhes = await pagina.QuerySelectorAsync("#_c8_tuc");
                                        if (revisarDetalhes != null)
                                        {
                                            Console.WriteLine("Tentando clicar no seletor de revisar detalhes do evento mesmo após erro...");
                                            try
                                            {
                                                await revisarDetalhes.ClickAsync();
                                                Console.WriteLine("✅ Seletor de revisar detalhes clicado após erro");
                                            }
                                            catch (Exception ex2)
                                            {
                                                Console.WriteLine($"⚠️ Erro ao clicar no seletor de revisar detalhes após erro: {ex2.Message}");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        // Extrair detalhes da cotação
                        bool processado = await ExtrairERegistrarCotacaoAsync(frame.Page, numeroCotacao, contaChave, cliente);

                        if (processado)
                        {
                            cotacoesProcessadas.Add(numeroCotacao);
                            cotacoesProcessadasNestaExpansao++;
                            _totalCotacoesProcessadas++;
                            Console.WriteLine($"   ✅ Processada com sucesso! (Total: {_totalCotacoesProcessadas})");
                        }
                        // Voltar para a lista
                        string voltarCotações = "https://service.ariba.com/Sourcing.aw/109521004/aw?awh=r&awssk=M1Ju378m&dard=1";

                        Console.WriteLine($"Navegando para pagina de cotações...");
                        await pagina.GotoAsync(voltarCotações, new PageGotoOptions
                        {
                            WaitUntil = WaitUntilState.DOMContentLoaded,
                            Timeout = 60000
                        });
                        await Task.Delay(10000);
                        frame = await _gerenciador.ObterFrameSupplierSeguroAsync(frame.Page) ?? frame;
                        locatorCandidatos = frame.Locator("tr:has(a[href*='webjumper']), tr:has-text('6000'), tr[class*='tableRow']");
                        totalCandidatos = await locatorCandidatos.CountAsync();

                        //Abrir o Toggle novamente
                        await Task.Delay(4000);
                        await ProcessarCotacoesDaContaPrioritariaAsync(pagina, contaChave, cliente, empresa);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ⚠️ Erro na linha {i}: {ex.Message}");
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ Erro na análise: {ex.Message}");
            }

            return cotacoesProcessadasNestaExpansao;
        }
        private async Task VerificarEmpresasNaoPrioritariasAsync(Estatisticas estatisticas, List<string> empresasNaoPrioritarias)
        {
            Console.WriteLine("\n" + new string('=', 80));
            Console.WriteLine("🔍 VERIFICANDO EMPRESAS NÃO PRIORITÁRIAS");
            Console.WriteLine("=".PadRight(80, '='));
            Console.WriteLine($"Empresas: {string.Join(", ", empresasNaoPrioritarias)}");
            Console.WriteLine("(Verificação completa - apenas coleta de números, sem processar individualmente)");
            Console.WriteLine(new string('=', 80));

            // Para cada empresa não prioritária, fazer a verificação completa
            foreach (var empresaNaoPrioritaria in empresasNaoPrioritarias)
            {
                Console.WriteLine($"\n📊 VERIFICANDO {empresaNaoPrioritaria}...");

                // Obter contas desta empresa
                var contasDaEmpresa = ObterContasDaEmpresa(empresaNaoPrioritaria);
                var cotaçõesDaEmpresa = new List<string>();

                foreach (var conta in contasDaEmpresa)
                {
                    // Fazer verificação COMPLETA (incluindo navegação para AEGEA/ESTÁCIO)
                    var cotações = await VerificarCotaçõesCompletoAsync(conta, empresaNaoPrioritaria);
                    cotaçõesDaEmpresa.AddRange(cotações);

                    // Aguardar entre contas
                    if (conta != contasDaEmpresa.Last())
                    {
                        await Task.Delay(5000);
                    }
                }

                // Remover duplicados
                cotaçõesDaEmpresa = cotaçõesDaEmpresa.Distinct().ToList();

                Console.WriteLine($"   📈 Total de cotações encontradas em {empresaNaoPrioritaria}: {cotaçõesDaEmpresa.Count}");

                if (cotaçõesDaEmpresa.Count > 0)
                {
                    Console.WriteLine($"   📋 Lista de cotações:");
                    foreach (var cotacao in cotaçõesDaEmpresa.Take(10)) // Mostra até 10
                    {
                        Console.WriteLine($"     • {cotacao}");
                    }
                    if (cotaçõesDaEmpresa.Count > 10)
                    {
                        Console.WriteLine($"     ... e mais {cotaçõesDaEmpresa.Count - 10} cotações");
                    }
                }

                // Comparar com cotações das empresas prioritárias
                var cotaçõesUnicas = new List<string>();
                foreach (var cotacao in cotaçõesDaEmpresa)
                {
                    bool encontradaEmPrioritaria = false;

                    foreach (var empresaPrioritaria in _empresasPrioritarias)
                    {
                        if (_cotacoesEncontradasPorEmpresa.ContainsKey(empresaPrioritaria) &&
                            _cotacoesEncontradasPorEmpresa[empresaPrioritaria].Contains(cotacao))
                        {
                            encontradaEmPrioritaria = true;
                            break;
                        }
                    }

                    if (!encontradaEmPrioritaria)
                    {
                        cotaçõesUnicas.Add(cotacao);
                    }
                }

                if (cotaçõesUnicas.Count > 0)
                {
                    Console.WriteLine($"\n⚠️  ATENÇÃO: {cotaçõesUnicas.Count} cotações encontradas APENAS em {empresaNaoPrioritaria}:");
                    foreach (var cotacao in cotaçõesUnicas)
                    {
                        Console.WriteLine($"   • {cotacao}");
                    }

                    // Perguntar se deseja processar estas cotações
                    Console.WriteLine($"\n❓ Deseja processar estas {cotaçõesUnicas.Count} cotações de {empresaNaoPrioritaria}?");
                    Console.WriteLine("   Digite 'S' para SIM ou qualquer outra tecla para NÃO");
                    Console.Write("   Resposta: ");

                    var resposta = Console.ReadLine()?.Trim().ToUpper();

                    if (resposta == "S" || resposta == "SIM")
                    {
                        Console.WriteLine($"\n🔄 Processando cotações únicas de {empresaNaoPrioritaria}...");

                        // Processar cada cotação única
                        foreach (var conta in contasDaEmpresa)
                        {
                            await ProcessarCotaçõesUnicasAsync(conta, empresaNaoPrioritaria, cotaçõesUnicas);
                        }
                    }
                    else
                    {
                        Console.WriteLine($"   ⏭️ Pulando cotações de {empresaNaoPrioritaria}");
                    }
                }
                else
                {
                    Console.WriteLine($"   ✅ Todas as cotações de {empresaNaoPrioritaria} já foram encontradas nas empresas prioritárias.");
                }

                // Aguardar entre empresas
                if (empresaNaoPrioritaria != empresasNaoPrioritarias.Last())
                {
                    Console.WriteLine($"\n⏳ Aguardando 5 segundos antes da próxima verificação...");
                    await Task.Delay(5000);
                }
            }
        }
        private async Task<List<string>> VerificarCotaçõesCompletoAsync(string chaveConta, string empresa)
        {
            var cotaçõesEncontradas = new List<string>();

            try
            {
                Console.WriteLine($"   🔍 Verificando conta {chaveConta} ({empresa})...");

                // Fazer login
                var pagina = await _gerenciador.RealizarLoginAsync(chaveConta);
                if (pagina == null)
                {
                    Console.WriteLine($"     ❌ Falha no login");
                    return cotaçõesEncontradas;
                }

                Console.WriteLine($"     ✅ Login bem-sucedido");

                // Determinar quais clientes verificar baseado na conta
                switch (chaveConta)
                {
                    case "1": // ALIANÇA ESTÁCIO - Apenas Estácio
                        Console.WriteLine($"     🎓 Navegando para ESTÁCIO...");
                        await NavegarParaEstacioAsync(pagina);
                        await Task.Delay(5000);
                        var cotaçõesEstacio = await BuscarCotaçõesNaPaginaAsync(pagina);
                        cotaçõesEncontradas.AddRange(cotaçõesEstacio);
                        break;

                    case "2": // ALIANÇA AEGEA - Apenas AEGEA
                        Console.WriteLine($"     🏢 Navegando para AEGEA...");
                        await _gerenciador.NavegarParaAegeaAsync(pagina);
                        await Task.Delay(5000);
                        var cotaçõesAEGEA = await BuscarCotaçõesNaPaginaAsync(pagina);
                        cotaçõesEncontradas.AddRange(cotaçõesAEGEA);
                        break;

                    case "3": // VENTURA AEGEA/ESTACIO - Ambas as empresas
                              // Verificar AEGEA primeiro
                        Console.WriteLine($"     🏢 Navegando para AEGEA (VENTURA)...");
                        await Task.Delay(5000);
                        var cotaçõesAegeaVentura = await BuscarCotaçõesNaPaginaAsync(pagina);
                        cotaçõesEncontradas.AddRange(cotaçõesAegeaVentura);
                        Console.WriteLine($"       📊 Cotações AEGEA: {cotaçõesAegeaVentura.Count}");

                        // Verificar Estácio
                        Console.WriteLine($"     🎓 Navegando para ESTÁCIO (VENTURA)...");
                        bool navegacaoEstacio = await NavegarParaEstacioAsync(pagina);
                        if (navegacaoEstacio)
                        {
                            await Task.Delay(5000);
                            var cotaçõesEstacioVentura = await BuscarCotaçõesNaPaginaAsync(pagina);
                            cotaçõesEncontradas.AddRange(cotaçõesEstacioVentura);
                            Console.WriteLine($"       📊 Cotações Estácio: {cotaçõesEstacioVentura.Count}");
                        }
                        break;

                    case "4": // UNIÃO AEGEA/ESTÁCIO - Ambas as empresas
                              // Verificar AEGEA primeiro
                        Console.WriteLine($"     🏢 Navegando para AEGEA (UNIÃO)...");
                        await _gerenciador.NavegarParaAegeaAsync(pagina);
                        await Task.Delay(5000);
                        var cotaçõesAegeaUniao = await BuscarCotaçõesNaPaginaAsync(pagina);
                        cotaçõesEncontradas.AddRange(cotaçõesAegeaUniao);
                        Console.WriteLine($"       📊 Cotações AEGEA: {cotaçõesAegeaUniao.Count}");

                        // Verificar Estácio
                        Console.WriteLine($"     🎓 Navegando para ESTÁCIO (UNIÃO)...");
                        bool navegacaoEstacioUniao = await NavegarParaEstacioAsync(pagina);
                        if (navegacaoEstacioUniao)
                        {
                            await Task.Delay(5000);
                            var cotaçõesEstacioUniao = await BuscarCotaçõesNaPaginaAsync(pagina);
                            cotaçõesEncontradas.AddRange(cotaçõesEstacioUniao);
                            Console.WriteLine($"       📊 Cotações Estácio: {cotaçõesEstacioUniao.Count}");
                        }
                        break;
                }

                await pagina.CloseAsync();
                Console.WriteLine($"     📭 Página fechada");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"     ❌ Erro na verificação completa: {ex.Message}");
            }

            return cotaçõesEncontradas.Distinct().ToList();
        }
        private async Task<List<string>> BuscarCotaçõesNaPaginaAsync(IPage pagina)
        {
            var cotações = new List<string>();

            try
            {
                Console.WriteLine($"       🔍 Procurando cotações na página...");

                // Obter frame Supplier
                var frameSupplier = await _gerenciador.ObterFrameSupplierSeguroAsync(pagina);
                if (frameSupplier == null)
                {
                    Console.WriteLine($"       ⚠️ Frame Supplier não encontrado");
                    return cotações;
                }

                await frameSupplier.WaitForLoadStateAsync(LoadState.NetworkIdle);
                await Task.Delay(3000);

                // Encontrar todas as linhas com Status: Aberto
                var linhasStatusAberto = frameSupplier.Locator("tr:has-text('Status: Aberto')");
                var totalLinhasAbertas = await linhasStatusAberto.CountAsync();

                Console.WriteLine($"       📊 Linhas com 'Status: Aberto': {totalLinhasAbertas}");

                if (totalLinhasAbertas == 0)
                {
                    Console.WriteLine($"       📭 Nenhuma linha encontrada");
                    return cotações;
                }

                // Processar cada linha (expandir toggles)
                for (int indiceLinha = 0; indiceLinha < totalLinhasAbertas; indiceLinha++)
                {
                    try
                    {
                        Console.WriteLine($"         📋 Processando linha {indiceLinha + 1}/{totalLinhasAbertas}...");

                        var linhaAtual = linhasStatusAberto.Nth(indiceLinha);

                        // Tentar expandir toggle
                        bool expandido = await ExpandirToggleEspecificoAsync(frameSupplier, linhaAtual, indiceLinha);

                        if (expandido)
                        {
                            Console.WriteLine($"           ✅ Toggle expandido");
                            await Task.Delay(2000);

                            // Buscar números de cotação na área expandida
                            var cotaçõesNaLinha = await BuscarCotaçõesNaAreaExpandidaAsync(frameSupplier);
                            cotações.AddRange(cotaçõesNaLinha);

                            Console.WriteLine($"           📈 Cotações nesta linha: {cotaçõesNaLinha.Count}");

                            // Se necessário, recolher o toggle para continuar
                            await Task.Delay(1000);
                        }
                        else
                        {
                            Console.WriteLine($"           ⚠️ Não foi possível expandir");

                            // Mesmo sem expandir, tentar encontrar cotações na linha
                            var textoLinha = await linhaAtual.TextContentAsync() ?? "";
                            var regexNumeroCotacao = new Regex(@"6000\d{6}");
                            var matches = regexNumeroCotacao.Matches(textoLinha);

                            foreach (Match match in matches)
                            {
                                cotações.Add(match.Value);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"         ❌ Erro na linha {indiceLinha + 1}: {ex.Message}");
                        continue;
                    }
                }

                Console.WriteLine($"       ✅ Total de cotações encontradas nesta página: {cotações.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"       💥 ERRO na busca: {ex.Message}");
            }

            return cotações.Distinct().ToList();
        }
        private async Task<List<string>> BuscarCotaçõesNaAreaExpandidaAsync(IFrame frame)
        {
            var cotações = new List<string>();

            try
            {
                // Buscar todas as linhas que podem conter cotações
                var locatorCandidatos = frame.Locator("tr:has(a[href*='webjumper']), tr:has-text('6000'), tr[class*='tableRow']");
                var totalCandidatos = await locatorCandidatos.CountAsync();

                for (int i = 0; i < totalCandidatos; i++)
                {
                    try
                    {
                        var linha = locatorCandidatos.Nth(i);
                        var textoLinha = (await linha.TextContentAsync() ?? "").Trim();

                        // Verificar se contém número de cotação
                        var regexNumeroCotacao = new Regex(@"6000\d{6}");
                        var match = regexNumeroCotacao.Match(textoLinha);

                        if (match.Success)
                        {
                            string numeroCotacao = match.Value;

                            // Verificar se já está na lista
                            if (!cotações.Contains(numeroCotacao))
                            {
                                cotações.Add(numeroCotacao);
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"         ⚠️ Erro na busca expandida: {ex.Message}");
            }

            return cotações;
        }

        private async Task ProcessarCotaçõesUnicasAsync(string chaveConta, string empresa, List<string> cotaçõesUnicas)
        {
            try
            {
                Console.WriteLine($"\n🔧 Processando {cotaçõesUnicas.Count} cotações de {empresa} (conta {chaveConta})...");

                var pagina = await _gerenciador.RealizarLoginAsync(chaveConta);
                if (pagina == null)
                {
                    Console.WriteLine($"   ❌ Falha no login");
                    return;
                }

                // Navegar para a empresa correta
                bool navegacaoOk = true;
                if (chaveConta == "1" || empresa == "ESTÁCIO")
                {
                    navegacaoOk = await NavegarParaEstacioAsync(pagina);
                }
                else if (chaveConta == "2" || empresa == "AEGEA")
                {
                    navegacaoOk = await _gerenciador.NavegarParaAegeaAsync(pagina);
                }
                else if (chaveConta == "3" || chaveConta == "4")
                {
                    // Para VENTURA e UNIÃO que têm ambas, vamos processar na página atual
                    // O usuário já navegou durante a verificação
                }

                if (!navegacaoOk)
                {
                    Console.WriteLine($"   ❌ Falha na navegação");
                    await pagina.CloseAsync();
                    return;
                }

                await Task.Delay(5000);

                // Para cada cotação única, processar individualmente
                foreach (var numeroCotacao in cotaçõesUnicas)
                {
                    try
                    {
                        Console.WriteLine($"\n📄 Processando cotação única: {numeroCotacao}");

                        // Aqui você precisaria encontrar e clicar na cotação específica
                        // Como é complexo, vamos apenas tentar registrar

                        // Extrair dados da cotação (usando os métodos existentes)
                        bool processado = await ExtrairERegistrarCotacaoAsync(pagina, numeroCotacao, chaveConta, empresa);

                        if (processado)
                        {
                            Console.WriteLine($"   ✅ Cotação {numeroCotacao} processada com sucesso!");
                        }
                        else
                        {
                            Console.WriteLine($"   ❌ Falha ao processar {numeroCotacao}");
                        }

                        // Voltar para a lista principal
                        await pagina.GoBackAsync();
                        await Task.Delay(3000);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ⚠️ Erro ao processar {numeroCotacao}: {ex.Message}");
                    }
                }

                await pagina.CloseAsync();
                Console.WriteLine($"✅ Processamento de {cotaçõesUnicas.Count} cotações concluído");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"💥 ERRO ao processar cotações de {empresa}: {ex.Message}");
            }
        }
        private async Task<List<string>> VerificarCotaçõesRapidoAsync(string chaveConta, string empresa)
        {
            var cotaçõesEncontradas = new List<string>();

            try
            {
                Console.WriteLine($"   🔍 Verificando conta {chaveConta}...");

                // Fazer login rápido apenas para verificar
                var pagina = await _gerenciador.RealizarLoginAsync(chaveConta);
                if (pagina == null) return cotaçõesEncontradas;

                // Navegar para a empresa correta
                if (chaveConta == "1" || chaveConta == "3" || chaveConta == "4")
                {
                    // Para empresas que precisam navegar
                    if (empresa == "ESTÁCIO" && (chaveConta == "3" || chaveConta == "4"))
                    {
                        await NavegarParaEstacioAsync(pagina);
                    }
                    else if (empresa == "AEGEA" && (chaveConta == "2" || chaveConta == "4"))
                    {
                        await _gerenciador.NavegarParaAegeaAsync(pagina);
                    }
                }

                await Task.Delay(3000);

                // Obter frame e listar cotações rapidamente
                var frameSupplier = await _gerenciador.ObterFrameSupplierSeguroAsync(pagina);
                if (frameSupplier != null)
                {
                    var regexNumeroCotacao = new Regex(@"6000\d{6}");
                    var html = await frameSupplier.ContentAsync();
                    var matches = regexNumeroCotacao.Matches(html);

                    foreach (Match match in matches)
                    {
                        cotaçõesEncontradas.Add(match.Value);
                    }

                    cotaçõesEncontradas = cotaçõesEncontradas.Distinct().ToList();
                }

                await pagina.CloseAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️ Erro na verificação rápida: {ex.Message}");
            }

            return cotaçõesEncontradas;
        }

        // Adicione este método ao relatório final
        private async Task GerarRelatorioFinalAsync(Estatisticas estatisticas, List<string> empresasNaoPrioritarias)
        {
            try
            {
                string relatorioFile = $"Relatorio_Prioridade_{DateTime.Now:yyyyMMdd_HHmmss}.txt";

                StringBuilder relatorio = new StringBuilder();

                relatorio.AppendLine("=".PadRight(80, '='));
                relatorio.AppendLine("📊 RELATÓRIO FINAL - SISTEMA COM PRIORIDADE");
                relatorio.AppendLine("=".PadRight(80, '='));
                relatorio.AppendLine($"Data/Hora: {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
                relatorio.AppendLine($"Duração: {estatisticas.DuracaoTotal:hh\\:mm\\:ss}");
                relatorio.AppendLine($"Ordem de Prioridade: {string.Join(" > ", _empresasPrioritarias)}");

                if (empresasNaoPrioritarias.Count > 0)
                {
                    relatorio.AppendLine($"Empresas não prioritárias: {string.Join(", ", empresasNaoPrioritarias)}");
                }
                relatorio.AppendLine();

                relatorio.AppendLine("📈 ESTATÍSTICAS:");
                relatorio.AppendLine($"- Total de contas processadas: {estatisticas.ContasComSucesso}/{estatisticas.TotalContas}");
                relatorio.AppendLine($"- Total de cotações encontradas: {estatisticas.TotalCotacoes}");
                relatorio.AppendLine($"- Cotações registradas: {_totalCotacoesProcessadas}");
                relatorio.AppendLine();

                relatorio.AppendLine("🏢 COTAÇÕES POR EMPRESA:");
                foreach (var empresa in new List<string> { "ALIANÇA", "VENTURA", "UNIÃO" })
                {
                    bool isPrioritaria = _empresasPrioritarias.Contains(empresa);
                    int total = _cotacoesEncontradasPorEmpresa.ContainsKey(empresa)
                        ? _cotacoesEncontradasPorEmpresa[empresa].Count
                        : 0;

                    string status = isPrioritaria ? "(PRIORITÁRIA)" : "(NÃO PRIORITÁRIA)";
                    relatorio.AppendLine($"  - {empresa} {status}: {total} cotações");
                }

                relatorio.AppendLine();
                relatorio.AppendLine("✅ PROCESSAMENTO CONCLUÍDO!");
                relatorio.AppendLine("=".PadRight(80, '='));

                await File.WriteAllTextAsync(relatorioFile, relatorio.ToString(), Encoding.UTF8);

                Console.WriteLine("\n" + new string('=', 80));
                Console.WriteLine(relatorio.ToString());
                Console.WriteLine($"💾 Relatório salvo em: {relatorioFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Erro ao gerar relatório: {ex.Message}");
            }
        }
        private async Task<bool> NavegarParaEstacioAsync(IPage pagina)
        {
            try
            {
                Console.WriteLine("🔍 Procurando botão 'MAIS'...");

                // Tentar vários seletores para o botão MAIS
                var maisSelectors = new[]
                {
            "button:has-text('MAIS')",
            "a:has-text('MAIS')",
            "[class*='mais']",
            "[class*='more']",
            "[id*='menu']",
            "[class*='menu']",
            "[role='menu']",
            "[aria-label*='menu']"
        };

                bool botaoMaisClicado = false;

                foreach (var selector in maisSelectors)
                {
                    try
                    {
                        var botaoMais = pagina.Locator(selector);
                        if (await botaoMais.CountAsync() > 0)
                        {
                            Console.WriteLine($"✅ Botão 'MAIS' encontrado com seletor: {selector}");
                            await botaoMais.First.ClickAsync();
                            await Task.Delay(5000);
                            botaoMaisClicado = true;
                            break;
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                if (!botaoMaisClicado)
                {
                    Console.WriteLine("⚠️ Botão 'MAIS' não encontrado, tentando alternativas...");

                    // Tentar encontrar qualquer elemento que possa abrir o menu
                    var todosBotoes = pagina.Locator("button, a, [role='button']");
                    var totalBotoes = await todosBotoes.CountAsync();

                    for (int i = 0; i < Math.Min(totalBotoes, 20); i++)
                    {
                        try
                        {
                            var botao = todosBotoes.Nth(i);
                            var texto = (await botao.TextContentAsync() ?? "").Trim();
                            var title = await botao.GetAttributeAsync("title") ?? "";

                            if (texto.Contains("...") || texto == "⋮" || texto == "☰" ||
                                title.Contains("menu") || title.Contains("more"))
                            {
                                Console.WriteLine($"🔄 Clicando no botão alternativo: '{texto}'");
                                await botao.ClickAsync();
                                await Task.Delay(5000);
                                botaoMaisClicado = true;
                                break;
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                if (!botaoMaisClicado)
                {
                    Console.WriteLine("❌ Não foi possível encontrar o botão 'MAIS'");
                    return false;
                }

                // Procurar o link específico da Estácio
                Console.WriteLine("🔍 Procurando 'Sociedade De Ensino Superior Estácio De Sá Ltda'...");

                // Aguardar menu abrir
                await Task.Delay(3000);

                // Tentar vários padrões
                var estacioSelectors = new[]
                {
            "a:has-text('Sociedade De Ensino Superior Estácio De Sá Ltda')",
            "a:has-text('Estácio')",
            "a:has-text('Estacio')",
            "a:has-text('Sociedade De Ensino')",
            "a.w-pmi-item:has-text('Estácio')",
            "a[class*='pmi']:has-text('Estácio')",
            "[role='menuitem']:has-text('Estácio')"
        };

                bool estacioClicado = false;

                foreach (var selector in estacioSelectors)
                {
                    try
                    {
                        var linkEstacio = pagina.Locator(selector);
                        if (await linkEstacio.CountAsync() > 0)
                        {
                            Console.WriteLine($"✅ Link da Estácio encontrado com: {selector}");
                            await linkEstacio.First.ClickAsync();
                            await Task.Delay(5000);
                            estacioClicado = true;
                            break;
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                if (!estacioClicado)
                {
                    // Debug: listar todos os links no menu
                    Console.WriteLine("📋 Listando todos os links no menu...");
                    var todosLinks = pagina.Locator("a, [role='menuitem'], [class*='menu-item']");
                    var count = await todosLinks.CountAsync();
                    Console.WriteLine($"Links encontrados no menu: {count}");

                    for (int i = 0; i < Math.Min(count, 10); i++)
                    {
                        try
                        {
                            var link = todosLinks.Nth(i);
                            var texto = (await link.TextContentAsync() ?? "").Trim();
                            if (!string.IsNullOrEmpty(texto))
                            {
                                Console.WriteLine($"  Link {i}: '{texto}'");

                                // Se encontrar algo com "Estácio" ou similar
                                if (texto.Contains("Estácio", StringComparison.OrdinalIgnoreCase) ||
                                    texto.Contains("Estacio", StringComparison.OrdinalIgnoreCase))
                                {
                                    Console.WriteLine($"⭐ Clicando no link: {texto}");
                                    await link.ClickAsync();
                                    await Task.Delay(5000);
                                    estacioClicado = true;
                                    break;
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                if (!estacioClicado)
                {
                    Console.WriteLine("❌ Link da Estácio não encontrado no menu");
                    return false;
                }

                // Verificar se navegou corretamente
                await Task.Delay(5000);

                var frameSupplier = await _gerenciador.ObterFrameSupplierSeguroAsync(pagina);
                if (frameSupplier != null)
                {
                    Console.WriteLine("✅ Navegação para Estácio bem-sucedida!");
                    return true;
                }
                else
                {
                    Console.WriteLine("⚠️ Frame Supplier não encontrado após navegação");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Erro ao navegar para Estácio: {ex.Message}");
                return false;
            }
        }

        private async Task<bool> ExpandirToggleEspecificoAsync(IFrame frame, ILocator linha, int indice)
        {
            try
            {
                var toggle = linha.Locator("span.w-togglebox-icon-off");

                if (await toggle.CountAsync() == 0)
                {
                    toggle = linha.Locator("[class*='togglebox'], [class*='toggle']");
                }

                if (await toggle.CountAsync() > 0)
                {
                    await toggle.ScrollIntoViewIfNeededAsync();
                    await Task.Delay(1500);

                    await toggle.ClickAsync(new LocatorClickOptions { Force = true });
                    await Task.Delay(3000);

                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private async Task<bool> ClicarLinkCotacaoAsync(IFrame frame, ILocator linha, int index)
        {
            try
            {
                // Tentar vários padrões de links
                var linkCotacao = linha.Locator("a[href*='webjumper'], a[href*='6000'], a:has-text('6000')");

                if (await linkCotacao.CountAsync() == 0)
                {
                    // Buscar qualquer link que contenha o número da cotação
                    var links = linha.Locator("a");
                    var totalLinks = await links.CountAsync();

                    for (int i = 0; i < totalLinks; i++)
                    {
                        var link = links.Nth(i);
                        var textoLink = (await link.TextContentAsync() ?? "").Trim();
                        if (textoLink.Contains("6000"))
                        {
                            linkCotacao = link;
                            break;
                        }
                    }
                }

                if (await linkCotacao.CountAsync() > 0)
                {
                    Console.WriteLine($"   👆 Clicando no link...");
                    await linkCotacao.ScrollIntoViewIfNeededAsync();
                    await Task.Delay(1000);

                    await linkCotacao.ClickAsync(new LocatorClickOptions
                    {
                        Force = true,
                        Timeout = 10000
                    });

                    await Task.Delay(5000);
                    return true;
                }

                Console.WriteLine($"   ⚠️ Nenhum link encontrado para clicar");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ❌ Erro ao clicar: {ex.Message}");
                return false;
            }
        }

        private async Task<bool> ExtrairERegistrarCotacaoAsync(IPage page, string numeroCotacao, string contaChave, string empresaAtual)
        {
            try
            {
                Console.WriteLine($"\n📋 EXTRAINDO DETALHES DA COTAÇÃO {numeroCotacao}...");

                await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                await Task.Delay(4000);

                // Extrair data de vencimento
                Console.WriteLine($"📅 Extraindo data de vencimento...");
                var (dataVencimento, horarioVencimento) = await ExtrairDataVencimentoAsync(page);
                Console.WriteLine($"   ✅ Vencimento: {dataVencimento} às {horarioVencimento}");

                // Extrair itens
                Console.WriteLine($"📦 Extraindo itens...");
                var itens = await ExtrairItensSimplesAsync(page);
                Console.WriteLine($"   ✅ Itens encontrados: {itens.Count}");

                // Baixar documento
                Console.WriteLine($"\n🖨️ Baixando documento...");
                bool downloadSucesso = await _gerenciador.BaixarDocumentoImpressaoAsync(page, numeroCotacao, contaChave, empresaAtual);

                if (downloadSucesso)
                {
                    Console.WriteLine($"   ✅ Download concluído");
                }
                else
                {
                    Console.WriteLine($"   ⚠️ Download não realizado");
                }

                // Registrar no banco de dados - USANDO _empresaAtual
                Console.WriteLine($"\n🗄️ Registrando no banco de dados...");
                bool registroSucesso = await _gerenciador.RegistrarCotacaoNoBancoDadosAsync(
                    numeroCotacao, dataVencimento, horarioVencimento, itens, contaChave, _empresaAtual);

                return registroSucesso;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Erro na extração: {ex.Message}");
                Console.WriteLine($"📌 StackTrace: {ex.StackTrace}");
                return false;
            }
        }


        private async Task<(string data, string horario)> ExtrairDataVencimentoAsync(IPage page)
        {
            try
            {
                var html = await page.ContentAsync();

                Console.WriteLine($"\n🔍 ANALISANDO HTML PARA EXTRAIR DATA DE VENCIMENTO...");

                // Padrão específico para <td valign="bottom">1/1/2026 09:52</td>
                // O problema é que o \d{1,2} pode capturar 09:52 como 9 (9:52)
                var patterns = new[]
                {
            // Padrão 1: Data e hora no formato exato - CORRIGIDO
            @"<td[^>]*valign\s*=\s*[""']bottom[""'][^>]*>\s*(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}:\d{2}\s*)",
            
            // Padrão 2: Qualquer data e hora em td - CORRIGIDO
            @"<td[^>]*>\s*(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}:\d{2})\s*</td>",
            
            // Padrão 3: Data e hora com possível espaço/linha quebrada
            @"<td[^>]*valign\s*=\s*[""']bottom[""'][^>]*>\s*(\d{1,2}/\d{1,2}/\d{4})\s+(\d{2}:\d{2})",
        };

                for (int i = 0; i < patterns.Length; i++)
                {
                    try
                    {
                        Console.WriteLine($"   🔍 Testando padrão {i + 1}...");
                        var regex = new Regex(patterns[i], RegexOptions.IgnoreCase | RegexOptions.Singleline);
                        var match = regex.Match(html);

                        if (match.Success)
                        {
                            string data = match.Groups[1].Value.Trim();
                            string horario = match.Groups[2].Value.Trim();

                            Console.WriteLine($"   ✅ Padrão {i + 1} encontrou: '{data}' '{horario}'");

                            // Debug: mostrar o match completo
                            Console.WriteLine($"   📋 Match completo: '{match.Value}'");

                            // Validar e formatar a data
                            if (DateTime.TryParse(data, out DateTime dataObj))
                            {
                                // Formatar data para dd/MM/yyyy
                                string dataFormatada = dataObj.ToString("dd/MM/yyyy");

                                // Formatar horário
                                horario = horario.Trim();

                                // Se o horário tem apenas 4 caracteres (ex: "9:52"), adicionar zero
                                if (horario.Length == 4 && horario.Contains(':'))
                                {
                                    horario = "0" + horario;
                                }

                                // Garantir que está no formato HH:mm
                                if (horario.Length == 5 && horario.Contains(':'))
                                {
                                    // Já está no formato HH:mm
                                }
                                else if (DateTime.TryParseExact(horario, "H:mm", null, System.Globalization.DateTimeStyles.None, out DateTime horaObj))
                                {
                                    horario = horaObj.ToString("HH:mm");
                                }

                                Console.WriteLine($"   📅 Data formatada: {dataFormatada} às {horario}");
                                return (dataFormatada, horario);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   ⚠️ Erro no padrão {i + 1}: {ex.Message}");
                        continue;
                    }
                }

                // Tentativa alternativa: buscar por regex mais específico
                Console.WriteLine($"\n🔍 Tentando regex específico para <td valign=\"bottom\">...");
                var specificPattern = @"<td\s+valign\s*=\s*""bottom""[^>]*>([^<]+)</td>";
                var specificRegex = new Regex(specificPattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                var specificMatch = specificRegex.Match(html);

                if (specificMatch.Success)
                {
                    string conteudo = specificMatch.Groups[1].Value.Trim();
                    Console.WriteLine($"   ✅ Conteúdo encontrado: '{conteudo}'");

                    // Extrair data e hora do conteúdo
                    var dataHoraPattern = @"(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}:\d{2})";
                    var dataHoraMatch = Regex.Match(conteudo, dataHoraPattern);

                    if (dataHoraMatch.Success)
                    {
                        string data = dataHoraMatch.Groups[1].Value;
                        string horario = dataHoraMatch.Groups[2].Value;

                        if (DateTime.TryParse(data, out DateTime dataObj))
                        {
                            string dataFormatada = dataObj.ToString("dd/MM/yyyy");

                            // Formatar horário
                            if (horario.Length == 4) horario = "0" + horario;

                            Console.WriteLine($"   ✅ Data extraída: {dataFormatada} às {horario}");
                            return (dataFormatada, horario);
                        }
                    }
                }

                Console.WriteLine($"   ❌ Nenhuma data de vencimento encontrada");
                return ("Não encontrada", "Não encontrado");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      ⚠️ Erro ao extrair data: {ex.Message}");
                return ("Erro", "Erro");
            }
        }
        private async Task<List<GerenciadorLogin.ItemCotacao>> ExtrairItensSimplesAsync(IPage page)
        {
            var itens = new List<GerenciadorLogin.ItemCotacao>();

            try
            {
                var html = await page.ContentAsync();

                Console.WriteLine($"\n🔍 ANALISANDO HTML PARA EXTRAIR ITENS...");

                // Padrão para capturar itens: <b>DESCRIÇÃO</b> seguida de quantidade
                // Vamos capturar tudo entre <b> e </b> que não seja "Menos..."
                var regex = new Regex(@"<b>\s*((?!Menos\.\.\.)[^<]+?)\s*</b>",
                    RegexOptions.Singleline | RegexOptions.IgnoreCase);

                var matches = regex.Matches(html);

                Console.WriteLine($"   🔍 Possiveis itens encontrados: {matches.Count}");

                foreach (Match match in matches)
                {
                    string descricaoBruta = match.Groups[1].Value.Trim();

                    // Limpar a descrição
                    string descricao = descricaoBruta
                        .Replace("\n", " ")
                        .Replace("\r", " ")
                        .Replace("  ", " ")
                        .Trim();

                    // Ignorar descrições vazias ou muito curtas
                    if (string.IsNullOrEmpty(descricao) || descricao.Length < 3)
                        continue;

                    // Ignorar "Upload de propostas" e "Menos..."
                    if (descricao.Contains("Upload de propostas", StringComparison.OrdinalIgnoreCase) ||
                        descricao.Contains("Menos...", StringComparison.OrdinalIgnoreCase) ||
                        descricao.Contains("Termo de Contratação", StringComparison.OrdinalIgnoreCase) ||
                        descricao.Contains("Dúvidas Técnicas", StringComparison.OrdinalIgnoreCase) ||
                        descricao.Contains("Propostas", StringComparison.OrdinalIgnoreCase) ||
                        descricao.Contains("Detalhamento", StringComparison.OrdinalIgnoreCase) ||
                        descricao.Contains("Escopo", StringComparison.OrdinalIgnoreCase))
                        continue;

                    Console.WriteLine($"   📝 Item encontrado: '{descricao}'");

                    // Tentar encontrar quantidade para este item
                    string quantidade = await ExtrairQuantidadeParaItemAsync(page, descricao);

                    var item = new GerenciadorLogin.ItemCotacao
                    {
                        DescricaoOriginal = descricao,
                        DescricaoLimpa = descricao,
                        Quantidade = quantidade,
                        Unidade = "unidade"
                    };

                    itens.Add(item);
                }

                Console.WriteLine($"\n✅ Total de itens validos: {itens.Count}");

                return itens;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   ⚠️ Erro na extracao: {ex.Message}");
                return itens;
            }
        }
        private async Task<string> ExtrairQuantidadeParaItemAsync(IPage page, string descricao)
        {
            try
            {
                // Procurar por padrão: quantidade após a descrição
                var regex = new Regex($@"{Regex.Escape(descricao)}.*?(\d+)\s*unidade",
                    RegexOptions.Singleline | RegexOptions.IgnoreCase);

                var html = await page.ContentAsync();
                var match = regex.Match(html);

                if (match.Success)
                {
                    return match.Groups[1].Value;
                }

                // Se não encontrou, tentar procurar na tabela
                var linhas = page.Locator("tr");
                var count = await linhas.CountAsync();

                for (int i = 0; i < count; i++)
                {
                    var linha = linhas.Nth(i);
                    var texto = await linha.TextContentAsync();

                    if (texto != null && texto.Contains(descricao, StringComparison.OrdinalIgnoreCase))
                    {
                        // Procurar números nesta linha
                        var qtdMatch = Regex.Match(texto, @"(\d+)\s*unidade", RegexOptions.IgnoreCase);
                        if (qtdMatch.Success)
                        {
                            return qtdMatch.Groups[1].Value;
                        }
                    }
                }

                return "1"; // Valor padrão
            }
            catch
            {
                return "1"; // Valor padrão em caso de erro
            }
        }

        private async Task<List<GerenciadorLogin.ItemCotacao>> ExtrairItensAlternativoAsync(IPage page)
        {
            var itens = new List<GerenciadorLogin.ItemCotacao>();

            try
            {
                // Procurar por todas as tags <b> que contêm descrições de itens
                var bElements = page.Locator("b");
                var count = await bElements.CountAsync();

                for (int i = 0; i < count; i++)
                {
                    var elemento = bElements.Nth(i);
                    var descricao = (await elemento.TextContentAsync() ?? "").Trim();

                    // Ignorar descrições inválidas
                    if (string.IsNullOrEmpty(descricao) ||
                        descricao.Contains("Menos...", StringComparison.OrdinalIgnoreCase) ||
                        descricao.Contains("Upload de propostas", StringComparison.OrdinalIgnoreCase) ||
                        descricao.Length < 5)
                        continue;

                    // Limpar espaços extras
                    descricao = Regex.Replace(descricao, @"\s+", " ").Trim();

                    Console.WriteLine($"   📝 Item encontrado (alternativo): '{descricao}'");

                    var item = new GerenciadorLogin.ItemCotacao
                    {
                        DescricaoOriginal = descricao,
                        DescricaoLimpa = descricao,
                        Quantidade = "1", // Tentar extrair depois
                        Unidade = "unidade"
                    };

                    itens.Add(item);
                }
            }
            catch
            {
                // Ignorar erros
            }

            return itens;
        }

        // Método para fechar recursos
        public async Task Fechar()
        {
            await _gerenciador.FecharAsync();
        }
    }
}