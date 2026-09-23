function enviarEditaisPorAbas() {
  var ui = SpreadsheetApp.getUi();

  var cotaGoogle = MailApp.getRemainingDailyQuota();

  var respCheckbox = ui.alert(
    'Confirmação de Envio',
    '📧 COTA DO GOOGLE: Você ainda tem ' + cotaGoogle + ' envios disponíveis hoje.\n\nDeseja enviar os editais marcados na Coluna H para os fornecedores?',
    ui.ButtonSet.YES_NO
  );

  if (respCheckbox == ui.Button.NO || respCheckbox == ui.Button.CLOSE) {
    ui.alert('Operação cancelada.');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Otimizando motor de busca...", "🤖 Iniciando", -1);

  function removerAcentos(texto) {
    if (!texto) return "";
    return String(texto)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
  }

  var mapaDDD = {
    11:"SP", 12:"SP", 13:"SP", 14:"SP", 15:"SP", 16:"SP", 17:"SP", 18:"SP", 19:"SP",
    21:"RJ", 22:"RJ", 24:"RJ", 27:"ES", 28:"ES",
    31:"MG", 32:"MG", 33:"MG", 34:"MG", 35:"MG", 37:"MG", 38:"MG",
    41:"PR", 42:"PR", 43:"PR", 44:"PR", 45:"PR", 46:"PR",
    47:"SC", 48:"SC", 49:"SC",
    51:"RS", 53:"RS", 54:"RS", 55:"RS",
    61:"DF", 62:"GO", 64:"GO", 63:"TO",
    65:"MT", 66:"MT", 67:"MS",
    68:"AC", 69:"RO",
    71:"BA", 73:"BA", 74:"BA", 75:"BA", 77:"BA", 79:"SE",
    81:"PE", 87:"PE", 82:"AL", 83:"PB", 84:"RN", 85:"CE", 88:"CE", 86:"PI", 89:"PI",
    91:"PA", 93:"PA", 94:"PA", 92:"AM", 97:"AM", 95:"RR", 96:"AP", 98:"MA", 99:"MA"
  };

  function obterEstadoPorDDD(telefone) {
    if (!telefone) return null;
    var numeros = String(telefone).replace(/\D/g, "");
    if (numeros.length < 2) return null;
    var ddd = parseInt(numeros.substring(0, 2));
    return mapaDDD[ddd] || null;
  }

  var mapaEstados = {
    "AC": "ACRE", "AL": "ALAGOAS", "AP": "AMAPA", "AM": "AMAZONAS",
    "BA": "BAHIA", "CE": "CEARA", "DF": "DISTRITO FEDERAL", "ES": "ESPIRITO SANTO",
    "GO": "GOIAS", "MA": "MARANHAO", "MT": "MATO GROSSO", "MS": "MATO GROSSO DO SUL",
    "MG": "MINAS GERAIS", "PA": "PARA", "PB": "PARAIBA", "PR": "PARANA",
    "PE": "PERNAMBUCO", "PI": "PIAUI", "RJ": "RIO DE JANEIRO", "RN": "RIO GRANDE DO NORTE",
    "RS": "RIO GRANDE DO SUL", "RO": "RONDONIA", "RR": "RORAIMA", "SC": "SANTA CATARINA",
    "SP": "SAO PAULO", "SE": "SERGIPE", "TO": "TOCANTINS"
  };
  var mapaAcentos = {
    "SP": "SÃO PAULO", "GO": "GOIÁS", "MA": "MARANHÃO", "PA": "PARÁ",
    "RO": "RONDÔNIA", "PI": "PIAUÍ", "CE": "CEARÁ", "PB": "PARAÍBA", "PR": "PARANÁ"
  };

  function obterTermosDeEstado(celula) {
    if (!celula || String(celula).trim() === "") return [];
    var siglas = String(celula).toUpperCase().split(/[\s,;]+/).filter(function(s){ return s !== ""; });
    var termosFinais = [];
    siglas.forEach(function(s) {
      termosFinais.push(s);
      if (mapaEstados[s]) termosFinais.push(mapaEstados[s]);
      if (mapaAcentos[s]) termosFinais.push(mapaAcentos[s]);
    });
    return termosFinais;
  }

  // =========================================================================
  // STOP WORDS — ignoradas tanto na keyword quanto no texto do fornecedor
  // =========================================================================
  var stopWords = ["de", "da", "do", "das", "dos", "e", "em", "na", "no", "com", "para", "por", "a", "o", "as", "os"];

  // Extrai palavras significativas (remove stop words e tokens curtos demais)
  function extrairTokensSignificativos(texto) {
    return removerAcentos(texto)
      .replace(/[^\w\s]/gi, " ")
      .split(/\s+/)
      .filter(function(w) {
        return w.length > 2 && stopWords.indexOf(w) === -1;
      });
  }

  // Normaliza token removendo plural e sufixos comuns
  function normalizarToken(w) {
    if (w.endsWith("oes")) return w.slice(0, -3) + "ao";
    if (w.endsWith("ais")) return w.slice(0, -2) + "al";
    if (w.endsWith("eis")) return w.slice(0, -2) + "el";
    if (w.endsWith("es")) return w.slice(0, -2);
    if (w.endsWith("s")) return w.slice(0, -1);
    return w;
  }

  // =========================================================================
  // MATCH ESTRITO
  //
  // Keyword e linha do fornecedor precisam ter os mesmos tokens significativos,
  // sem sobras em nenhum dos lados. Isso evita falsos positivos como:
  //   keyword "MONITORAMENTO REMOTO" batendo em "PEÇAS DE MONITORAMENTO REMOTO"
  //   keyword "INFORMATICA" batendo em "MANUTENÇÃO DE INFORMATICA"
  //
  // Exemplos:
  //   keyword "MONITORAMENTO REMOTO"
  //     linha "MONITORAMENTO REMOTO"          → ✅
  //     linha "PEÇAS DE MONITORAMENTO REMOTO" → ❌ ("pecas" sobra)
  //
  //   keyword "AUTOPEÇAS"
  //     linha "AUTOPEÇAS"                     → ✅
  //     linha "PEÇAS E SERVIÇOS AUTOMOTIVOS"  → ❌ (tokens diferentes)
  //
  //   keyword "PEÇAS E SERVIÇOS PARA MAQUINAS PESADAS"
  //     linha "PEÇAS E SERVIÇOS PARA MAQUINAS PESADAS" → ✅
  // =========================================================================
  function matchEstrito(tokensKeyword, textoLinha) {
    var tokensTexto = extrairTokensSignificativos(textoLinha);
    if (tokensTexto.length === 0) return false;

    var tokensKwNorm  = tokensKeyword.map(normalizarToken);
    var tokensTexNorm = tokensTexto.map(normalizarToken);

    // Todo token do TEXTO deve estar coberto pela keyword
    for (var t = 0; t < tokensTexNorm.length; t++) {
      var encontrado = false;
      for (var k = 0; k < tokensKwNorm.length; k++) {
        if (tokensKwNorm[k].indexOf(tokensTexNorm[t]) === 0 ||
            tokensTexNorm[t].indexOf(tokensKwNorm[k]) === 0) {
          encontrado = true;
          break;
        }
      }
      if (!encontrado) return false;
    }

    // Todo token da KEYWORD deve estar presente no texto
    for (var k = 0; k < tokensKwNorm.length; k++) {
      var encontrado = false;
      for (var t = 0; t < tokensTexNorm.length; t++) {
        if (tokensKwNorm[k].indexOf(tokensTexNorm[t]) === 0 ||
            tokensTexNorm[t].indexOf(tokensKwNorm[k]) === 0) {
          encontrado = true;
          break;
        }
      }
      if (!encontrado) return false;
    }

    return true;
  }

  // =========================================================================
  // MATCH POR LINHAS
  //
  // Testa cada linha da célula individualmente. Resolve o caso de
  // fornecedores com múltiplos produtos separados por quebra de linha.
  //
  //   célula: "BATERIA AUTOMOTIVA\nAUTOPEÇAS\nSERVIÇOS MECÂNICOS"
  //   keyword "AUTOPEÇAS"          → bate na linha "AUTOPEÇAS"          ✅
  //   keyword "BATERIA AUTOMOTIVA" → bate na linha "BATERIA AUTOMOTIVA" ✅
  //   keyword "INFORMATICA"        → não bate em nenhuma linha          ❌
  // =========================================================================
  function temMatchEmAlgumaLinha(tokensKeyword, textoCelulaBruto) {
    var linhas = String(textoCelulaBruto).split(/[\n\r]+/);
    for (var l = 0; l < linhas.length; l++) {
      var linha = linhas[l].replace(/[^\w\s\u00C0-\u00FF]/g, " ").replace(/\s+/g, " ").trim();
      if (linha === "") continue;
      if (matchEstrito(tokensKeyword, linha)) return true;
    }
    return false;
  }

  // --- CONFIGURAÇÃO ---
  var MODO_TESTE = false;
  var SEU_EMAIL_TESTE = "enricomenicucci818@gmail.com";
  var limiteDiarioScript = 1500;
  var URL_IMAGEM_CABECALHO = "https://i.imgur.com/VnzKnsC.png";

  var imagemBlob = null;
  var imagemInlineObj = {};

  try {
    if (URL_IMAGEM_CABECALHO && URL_IMAGEM_CABECALHO.startsWith("http")) {
      imagemBlob = UrlFetchApp.fetch(URL_IMAGEM_CABECALHO).getBlob().setName("bannerHeader.jpeg");
      imagemInlineObj = { bannerHeader: imagemBlob };
    }
  } catch(e) {
    Logger.log("Erro ao baixar imagem: " + e.message);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var todasAbas = ss.getSheets();
  var abaEditais = ss.getSheetByName("Editais");

  if (!abaEditais) {
    Browser.msgBox("ERRO: Aba 'Editais' não encontrada.");
    return;
  }

  var dadosEditais = abaEditais.getRange(2, 1, abaEditais.getLastRow() - 1, 8).getValues();

  var editaisProcessados = dadosEditais
    .filter(function(linha) {
      return linha[7] === true;
    })
    .map(function(linha) {
      var rawKeys = String(linha[5]).split(",");
      var positivas = [];
      var negativas = [];

      rawKeys.forEach(function(k) {
        var termo = k.trim();
        if (termo === "") return;
        if (termo.startsWith("-")) {
          var semSinal = termo.substring(1);
          negativas.push(removerAcentos(semSinal).replace(/[^\w\s]/gi, " ").trim());
        } else {
          positivas.push(termo);
        }
      });

      // Pré-processa cada keyword em tokens significativos
      var keywordTokens = [];
      positivas.forEach(function(kword) {
        var kwordTratada = kword.replace(/\bp\//gi, " ");
        var tokens = extrairTokensSignificativos(kwordTratada);
        if (tokens.length > 0) {
          keywordTokens.push(tokens);
        }
      });

      return {
        cidade: linha[0],
        numEdital: linha[1],
        objeto: String(linha[2]).trim(),
        dataLimite: linha[3],
        linkDrive: linha[4],
        keywordTokens: keywordTokens,
        termosProibidos: negativas,
        termosEstadoAlvo: obterTermosDeEstado(linha[6])
      };
    });

  if (editaisProcessados.length === 0) {
    ui.alert("Nenhum edital selecionado!");
    return;
  }

  var totalLinhasParaLer = 0;
  for (var k = 0; k < todasAbas.length; k++) {
    var nomeAba = todasAbas[k].getName().toLowerCase();
    if (nomeAba === "editais" || nomeAba === "0" || nomeAba === "consultas" || nomeAba === "resumo") continue;
    var lin = todasAbas[k].getLastRow();
    if (lin > 1) totalLinhasParaLer += (lin - 1);
  }

  var sentCount = 0;
  var linhasLidasCount = 0;

  for (var k = 0; k < todasAbas.length; k++) {
    var abaAtual = todasAbas[k];
    var nomeAba = abaAtual.getName().toLowerCase();

    if (nomeAba === "editais" || nomeAba === "0" || nomeAba === "consultas" || nomeAba === "resumo") continue;

    var ultimaLinha = abaAtual.getLastRow();
    if (ultimaLinha < 2) continue;

    var rangeDados = abaAtual.getRange(2, 1, ultimaLinha - 1, 9);
    var dadosFornecedores = rangeDados.getValues();
    var coresFundo = rangeDados.getBackgrounds();

    var novosStatus = [];
    var houveAlteracao = false;
    var categoriaMestra = "";
    var categoriaMestraBruta = "";

    for (var i = 0; i < dadosFornecedores.length; i++) {
      linhasLidasCount++;

      if (linhasLidasCount % 1000 === 0) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          "Analisando: " + linhasLidasCount + " / " + totalLinhasParaLer + " fornecedores",
          "🔍 Varrendo Planilha...",
          -1
        );
      }

      var numeroLinha = String(dadosFornecedores[i][0]).trim();

      // Preserva o valor BRUTO (com \n) para o split por linha no match
      var produtoCelulaBruto = String(dadosFornecedores[i][4]);

      // Versão limpa usada apenas para filtros de estado e palavras proibidas
      var produtoCelulaLimpo = produtoCelulaBruto
        .replace(/[\n\r]+/g, " ")
        .replace(/[^\w\s\u00C0-\u00FF]/g, " ")
        .replace(/\s+/g, " ")
        .trim();

      if (numeroLinha === "1" || numeroLinha === "1.0") {
        categoriaMestra = produtoCelulaLimpo;
        categoriaMestraBruta = produtoCelulaBruto;
      }

      // Texto unificado apenas para filtros (estado e palavras proibidas)
      var textoCompletoContexto = categoriaMestra;
      if (produtoCelulaLimpo !== "" && produtoCelulaLimpo !== categoriaMestra) {
        textoCompletoContexto += " " + produtoCelulaLimpo;
      }

      var textoFornecedorMaiusculo = textoCompletoContexto.toUpperCase();
      var textoFornecedorNormalizado = removerAcentos(textoCompletoContexto);

      var statusAtual = dadosFornecedores[i][8];
      var corFundoEmail = coresFundo[i][5];
      var coresProibidas = ["#ff0000", "#cc0000"];

      if ((!MODO_TESTE && statusAtual == "Enviado") ||
          sentCount >= limiteDiarioScript ||
          coresProibidas.indexOf(corFundoEmail) > -1) {
        novosStatus.push([statusAtual]);
        continue;
      }

      var nomeEmpresa = dadosFornecedores[i][1];
      var emailOriginal = String(dadosFornecedores[i][5]);

      if (!nomeEmpresa && emailOriginal.trim() === "") {
        novosStatus.push([statusAtual]);
        continue;
      }

      var telefone = dadosFornecedores[i][3];
      var estadoPorDDD = obterEstadoPorDDD(telefone);
      var rawG = String(dadosFornecedores[i][6]).toUpperCase();
      var estadosParticipacao = rawG.replace(/[\/;,]/g, " ");
      var estadosParticipacaoNorm = removerAcentos(estadosParticipacao);

      var htmlEditaisAba = "";
      var contagemEditais = 0;
      var nomesObjetos = [];

      for (var j = 0; j < editaisProcessados.length; j++) {
        var edital = editaisProcessados[j];

        // --- FILTRO DE ESTADO ---
        if (edital.termosEstadoAlvo.length > 0) {
          var fornecedorEhDoPublicoAlvo = false;
          for (var t = 0; t < edital.termosEstadoAlvo.length; t++) {
            var termoAlvo = edital.termosEstadoAlvo[t];
            if (textoFornecedorMaiusculo.indexOf(termoAlvo) > -1 ||
                estadosParticipacao.indexOf(termoAlvo) > -1 ||
                estadosParticipacaoNorm.indexOf(termoAlvo) > -1) {
              fornecedorEhDoPublicoAlvo = true;
              break;
            }
          }
          if (!fornecedorEhDoPublicoAlvo && estadoPorDDD) {
            if (edital.termosEstadoAlvo.indexOf(estadoPorDDD) > -1) fornecedorEhDoPublicoAlvo = true;
          }
          if (!fornecedorEhDoPublicoAlvo) continue;
        }

        // --- TRAVA DE PALAVRAS PROIBIDAS ---
        var temPalavraProibida = false;
        if (edital.termosProibidos.length > 0) {
          for (var n = 0; n < edital.termosProibidos.length; n++) {
            if (textoFornecedorNormalizado.indexOf(edital.termosProibidos[n].toLowerCase()) > -1) {
              temPalavraProibida = true;
              break;
            }
          }
        }
        if (temPalavraProibida) continue;

        // =====================================================================
        // MATCH DE PALAVRAS-CHAVE
        //
        // Cada keyword é testada linha a linha contra o produto e a categoria.
        // O envio ocorre se QUALQUER keyword der match em QUALQUER linha.
        //
        // keyword "MONITORAMENTO REMOTO":
        //   linha "MONITORAMENTO REMOTO"          → ✅ envia
        //   linha "PEÇAS DE MONITORAMENTO REMOTO" → ❌ bloqueia ("pecas" sobra)
        //
        // keyword "AUTOPEÇAS, PEÇAS E SERVIÇOS PARA MAQUINAS PESADAS":
        //   linha "AUTOPEÇAS"                              → ✅ envia (1ª keyword)
        //   linha "PEÇAS E SERVIÇOS PARA MAQUINAS PESADAS" → ✅ envia (2ª keyword)
        //   linha "BATERIA AUTOMOTIVA"                     → ❌ nenhuma keyword bate
        // =====================================================================
        var temMatchKeyword = false;

        for (var r = 0; r < edital.keywordTokens.length; r++) {
          var tokens = edital.keywordTokens[r];

          if (temMatchEmAlgumaLinha(tokens, produtoCelulaBruto) ||
              temMatchEmAlgumaLinha(tokens, categoriaMestraBruta)) {
            temMatchKeyword = true;
            break;
          }
        }

        if (temMatchKeyword) {
          var novoObjeto = edital.objeto.trim();
          if (novoObjeto !== "") {
            var jaExiste = nomesObjetos.some(function(o) {
              return o.toLowerCase().includes(novoObjeto.toLowerCase()) || novoObjeto.toLowerCase().includes(o.toLowerCase());
            });
            if (!jaExiste) nomesObjetos.push(novoObjeto);
          }

          var dataFormatada = edital.dataLimite;
          if (edital.dataLimite instanceof Date) {
            dataFormatada = Utilities.formatDate(edital.dataLimite, "GMT-3", "dd/MM/yyyy HH:mm");
          }

          var botoesHtml = "";
          var arrayLinks = String(edital.linkDrive).split(/[\s,;]+/).filter(function(l){ return l.toLowerCase().indexOf("http") === 0; });

          if (arrayLinks.length > 1) {
            for (var k_link = 0; k_link < arrayLinks.length; k_link++) {
              botoesHtml += `<a href="${arrayLinks[k_link]}" style="background-color: #0056b3; color: #ffffff; padding: 8px 15px; text-decoration: none; border-radius: 4px; font-weight: bold; margin-right: 5px; font-size: 12px;">PARTE ${k_link + 1}</a> `;
            }
          } else if (arrayLinks.length === 1) {
            botoesHtml = `<a href="${arrayLinks[0]}" style="background-color: #0056b3; color: #ffffff; padding: 8px 15px; text-decoration: none; border-radius: 4px; font-weight: bold;">BAIXAR EDITAL</a>`;
          } else {
            botoesHtml = `<span style="color: #999;">Link indisponível</span>`;
          }

          htmlEditaisAba += `
            <div style="margin-bottom: 20px; border-bottom: 1px solid #ccc; padding-bottom: 15px;">
              <p style="margin: 0; font-weight: bold; font-size: 16px; color: #2c3e50;">${edital.cidade} - Edital ${edital.numEdital}</p>
              <p style="margin: 5px 0;"><strong>Objeto:</strong> ${edital.objeto}</p>
              <p style="margin: 5px 0;"><strong>Recebimento de propostas até:</strong> ${dataFormatada}</p>
              <div style="margin-top: 10px;">
                 ${botoesHtml}
              </div>
            </div>
          `;
          contagemEditais++;
        }
      }

      if (contagemEditais === 0) {
        novosStatus.push([statusAtual]);
        continue;
      }

      var emailTexto = String(emailOriginal);
      var listaEmails = emailTexto.split(/[\s,;]+/).filter(function(e) {
        return e.indexOf('@') > -1;
      });

      if (listaEmails.length === 0) {
        novosStatus.push([statusAtual]);
        continue;
      }

      var destinatarioReal = listaEmails.join(',');
      var destinatarioFinal = MODO_TESTE ? SEU_EMAIL_TESTE : destinatarioReal;

      var dataHoje = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");
      var assunto = "Editais Órgãos Públicos - Plataforma BBMNET Licitações - " + dataHoje;

      var imgTag = "";
      if (imagemBlob) {
        imgTag = `<div style="text-align: center; margin-bottom: 20px;">
                    <img src="cid:bannerHeader" alt="Cabeçalho" width="600" height="150" style="display: block; max-width: 100%; height: auto; margin: 0 auto; border-radius: 4px; border: 0;">
                  </div>`;
      }

      var corpoEmail = `
        <div style="font-family: 'Segoe UI', Arial, sans-serif; color: #202124; font-size: 16px; max-width: 600px; line-height: 1.6; margin: 0 auto;">
          ${imgTag}
          <p>Prezado(a),</p>
          <p>Segue os editais das Prefeituras com oportunidade de vendas para:</p>
          <hr style="border: 0; border-top: 2px solid #666; margin: 20px 0;">
          ${htmlEditaisAba}
          <div style="background-color: #f1f3f4; border: 1px solid #e0e0e0; border-radius: 8px; padding: 25px; margin-top: 40px; font-size: 14px;">
            <h3 style="margin-top: 0; color: #444; font-size: 18px; text-align: center; margin-bottom: 10px;">Acesso à Plataforma BBMNET</h3>
            <p style="text-align: center; margin-bottom: 20px; color: #555;">
              Para se cadastrar, renovar ou entrar no sistema, utilize o botão abaixo:
            </p>
            <div style="text-align: center; margin-bottom: 25px;">
              <a href="https://novobbmnet.com.br/licitante/" target="_blank"
                 style="background-color: #28a745; color: #ffffff; padding: 14px 28px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 16px; display: inline-block; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                 ACESSAR PLATAFORMA
              </a>
            </div>
            <div style="border-top: 1px solid #ddd; padding-top: 15px;">
              <p style="margin: 8px 0;"><strong>📺 Como operar a Plataforma?</strong> <a href="https://novobbmnet.com.br/como-operar/#manuais-como-operar-v2" style="color: #0056b3; text-decoration: none;">Clique aqui e assista aos vídeos de orientação</a></p>
              <p style="margin: 8px 0;"><strong>📞 Suporte Técnico: </strong> (11) 3181-8214 / (11) 91666-9860</p>
              <p style="margin: 8px 0;"><strong>Dificuldade no cadastro? </strong> (34) 99995-0156 (WhatsApp)</p>
              <p style="margin: 8px 0; font-size: 13px; color: #666;"><em>Dúvidas sobre o edital? Envie diretamente no campo (ESCLARECIMENTOS) da plataforma.</em></p>
            </div>
            <div style="background-color: #e8f0fe; border: 2px solid #aecbfa; color: #1967d2; padding: 10px; margin-top: 20px; text-align: center; font-weight: bold; font-size: 14px; border-radius: 6px;">
               ACESSO: VIA REPRESENTANTE - SAGA CORRETAGEM
            </div>
          </div>
          <p style="font-size: 12px; color: #888; margin-top: 20px; text-align: center;">Atenciosamente,<br>BBMNET Licitações</p>
        </div>
      `;

      try {
        var cotaAtual = cotaGoogle - sentCount;
        SpreadsheetApp.getActiveSpreadsheet().toast("Destino: " + nomeEmpresa, "🚀 Enviando (Restam: " + cotaAtual + ")", -1);

        MailApp.sendEmail({
          to: destinatarioFinal,
          subject: assunto,
          htmlBody: corpoEmail,
          inlineImages: imagemInlineObj
        });

        novosStatus.push(["Enviado"]);
        houveAlteracao = true;
        sentCount++;
        Utilities.sleep(1000);

      } catch (e) {
        var erroMsg = String(e.message);

        if (erroMsg.toLowerCase().indexOf("limit") > -1 ||
            erroMsg.toLowerCase().indexOf("quota") > -1 ||
            erroMsg.toLowerCase().indexOf("muitas vezes") > -1 ||
            erroMsg.toLowerCase().indexOf("too many times") > -1) {

          novosStatus.push(["Erro: Bloqueio do Google 🛑"]);

          if (novosStatus.length > 0) {
            abaAtual.getRange(2, 9, novosStatus.length, 1).setValues(novosStatus);
          }

          ui.alert("🚨 PARADA DE EMERGÊNCIA\n\nO Google bloqueou a operação. O erro exato retornado foi:\n\n👉 " + erroMsg + "\n\nCota Restante Atual: " + MailApp.getRemainingDailyQuota());
          return;
        }

        novosStatus.push(["Erro: " + e.message]);
        houveAlteracao = true;
      }
    }

    if (houveAlteracao && novosStatus.length > 0) {
      abaAtual.getRange(2, 9, novosStatus.length, 1).setValues(novosStatus);
    }

    if (sentCount >= limiteDiarioScript) break;
  }

  var cotaRestante = MailApp.getRemainingDailyQuota();

  var msgFinal = "";
  if (sentCount === 0) {
    msgFinal = "⚠️ ALERTA: Nenhum fornecedor encontrado.\nVerifique Palavras-Chave e Filtro de Estado.";
  } else if (sentCount >= limiteDiarioScript) {
    msgFinal = "⚠️ Limite atingido (" + limiteDiarioScript + ").";
  } else {
    msgFinal = "✅ Processo concluído com sucesso!";
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Foram lidos " + linhasLidasCount + " fornecedores.", "✅ Concluído", 5);

  ui.alert(msgFinal + "\n\n📤 Total enviados: " + sentCount + "\n📧 Cota Restante: " + cotaRestante);
}