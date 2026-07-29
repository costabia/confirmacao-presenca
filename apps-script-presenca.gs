function normalizarTexto(texto) {
  if (!texto) return "";
  return texto.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function separarAcompanhantes(valor) {
  if (!valor) return [];
  return valor.toString().split(";").map(function(nome) {
    return nome.trim();
  }).filter(function(nome) {
    return nome !== "";
  });
}

function buscarConvitePorNome(sheetLista, nomeNormalizado) {
  const dataLista = sheetLista.getDataRange().getValues();

  for (let i = 1; i < dataLista.length; i++) {
    const nomeTitular = dataLista[i][0];
    if (normalizarTexto(nomeTitular) === nomeNormalizado) {
      return {
        row: i + 1,
        nome: nomeTitular,
        permitidos: Number(dataLista[i][1]) || 0,
        acompanhantes: separarAcompanhantes(dataLista[i][2])
      };
    }
  }

  return null;
}

function buscarConfirmacaoPorNome(sheetConfirmados, nomeNormalizado) {
  const dataConfirmados = sheetConfirmados.getDataRange().getValues();

  for (let i = 1; i < dataConfirmados.length; i++) {
    if (normalizarTexto(dataConfirmados[i][0]) === nomeNormalizado) {
      return {
        row: i + 1,
        data: dataConfirmados[i]
      };
    }
  }

  return null;
}

function contarLinhasComValor(data, coluna) {
  let total = 0;
  for (let i = 1; i < data.length; i++) {
    if (normalizarTexto(data[i][coluna])) total++;
  }
  return total;
}

function contarPessoasNaListaOriginal(dataLista) {
  let total = 0;

  for (let i = 1; i < dataLista.length; i++) {
    if (!normalizarTexto(dataLista[i][0])) continue;
    total += 1 + separarAcompanhantes(dataLista[i][2]).length;
  }

  return total;
}

function montarListaConfirmados(dataConfirmados) {
  const confirmados = [];

  for (let i = 1; i < dataConfirmados.length; i++) {
    if (!normalizarTexto(dataConfirmados[i][0])) continue;

    const todosConfirmadosRaw = dataConfirmados[i][6] ? dataConfirmados[i][6].toString() : "";
    const todosConfirmados = separarAcompanhantes(todosConfirmadosRaw);

    confirmados.push({
      nome_principal: dataConfirmados[i][0],
      acompanhantes_permitidos: Number(dataConfirmados[i][1]) || 0,
      acompanhantes_confirmados: Number(dataConfirmados[i][2]) || 0,
      data_confirmacao: dataConfirmados[i][3] || "",
      titular_vai: dataConfirmados[i][4] || "",
      acompanhantes: separarAcompanhantes(dataConfirmados[i][5]),
      todos_confirmados: todosConfirmados
    });
  }

  return confirmados;
}

function montarListaPendentes(dataLista, dataConfirmados) {
  const pendentes = [];
  const confirmadosPorNome = {};

  for (let i = 1; i < dataConfirmados.length; i++) {
    const nomeConfirmado = normalizarTexto(dataConfirmados[i][0]);
    if (nomeConfirmado) {
      confirmadosPorNome[nomeConfirmado] = true;
    }
  }

  for (let i = 1; i < dataLista.length; i++) {
    if (!normalizarTexto(dataLista[i][0])) continue;

    const nomeNormalizado = normalizarTexto(dataLista[i][0]);
    if (!confirmadosPorNome[nomeNormalizado]) {
      pendentes.push({
        nome_principal: dataLista[i][0],
        acompanhantes_permitidos: Number(dataLista[i][1]) || 0,
        acompanhantes: separarAcompanhantes(dataLista[i][2])
      });
    }
  }

  return pendentes;
}

function doGetAdmin(e) {
  const senha = e && e.parameter ? e.parameter.senha : "";
  if (senha !== "0909") {
    return jsonResponse({ status: "unauthorized", message: "Senha inválida." });
  }

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLista = spreadSheet.getSheetByName("Lista_Original");
  const sheetConfirmados = spreadSheet.getSheetByName("Confirmados");
  const sheetPortaria = spreadSheet.getSheetByName("Portaria");

  if (!sheetLista || !sheetConfirmados || !sheetPortaria) {
    return jsonResponse({ status: "error", message: "A planilha está sem as abas Lista_Original, Confirmados ou Portaria." });
  }

  const dataLista = sheetLista.getDataRange().getValues();
  const dataConfirmados = sheetConfirmados.getDataRange().getValues();
  const dataPortaria = sheetPortaria.getDataRange().getValues();
  const confirmados = montarListaConfirmados(dataConfirmados);
  const pendentes = montarListaPendentes(dataLista, dataConfirmados);
  const totalConvites = contarLinhasComValor(dataLista, 0);
  const convitesConfirmados = confirmados.length;
  const pessoasConfirmadas = contarLinhasComValor(dataPortaria, 0);
  const totalPessoas = contarPessoasNaListaOriginal(dataLista);

  return jsonResponse({
    status: "success",
    atualizado_em: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
    resumo: {
      total_convites: totalConvites,
      convites_confirmados: convitesConfirmados,
      convites_pendentes: Math.max(totalConvites - convitesConfirmados, 0),
      pessoas_confirmadas: pessoasConfirmadas,
      pessoas_nao_confirmadas: Math.max(totalPessoas - pessoasConfirmadas, 0),
      total_pessoas: totalPessoas
    },
    confirmados: confirmados,
    pendentes: pendentes
  });
}

function removerPortariaDaFamilia(sheetPortaria, nomesPossiveisNormalizados) {
  const dataPortaria = sheetPortaria.getDataRange().getValues();

  for (let j = dataPortaria.length - 1; j >= 1; j--) {
    const nomePortariaNormalizado = normalizarTexto(dataPortaria[j][0]);
    if (nomesPossiveisNormalizados.includes(nomePortariaNormalizado)) {
      sheetPortaria.deleteRow(j + 1);
    }
  }
}

function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : "";
  if (action === "admin") {
    return doGetAdmin(e);
  }

  const nomeBusca = e && e.parameter ? e.parameter.nome : "";
  if (!nomeBusca || !normalizarTexto(nomeBusca)) {
    return jsonResponse({ status: "error", message: "Nome não fornecido." });
  }

  const nomeNormalizado = normalizarTexto(nomeBusca);
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLista = spreadSheet.getSheetByName("Lista_Original");
  const sheetConfirmados = spreadSheet.getSheetByName("Confirmados");

  if (!sheetLista || !sheetConfirmados) {
    return jsonResponse({ status: "error", message: "A planilha está sem as abas necessárias." });
  }

  const convite = buscarConvitePorNome(sheetLista, nomeNormalizado);
  if (!convite) {
    return jsonResponse({ status: "not_found" });
  }

  const confirmacao = buscarConfirmacaoPorNome(sheetConfirmados, normalizarTexto(convite.nome));
  if (confirmacao) {
    const qtdeAcompanhantes = confirmacao.data[2];
    return jsonResponse({
      status: "already_confirmed",
      nome_exato: convite.nome,
      detalhes: qtdeAcompanhantes !== "" ? qtdeAcompanhantes + " acompanhante(s)" : "Apenas o titular"
    });
  }

  return jsonResponse({
    status: "found",
    nome_exato: convite.nome,
    acompanhantes_permitidos: convite.permitidos,
    acompanhantes: convite.acompanhantes
  });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: "error", message: "Dados não recebidos." });
    }

    const data = JSON.parse(e.postData.contents);
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetConfirmados = spreadSheet.getSheetByName("Confirmados");
    const sheetPortaria = spreadSheet.getSheetByName("Portaria");
    const sheetLista = spreadSheet.getSheetByName("Lista_Original");

    if (!sheetConfirmados || !sheetPortaria || !sheetLista) {
      return jsonResponse({ status: "error", message: "A planilha está sem as abas Confirmados, Portaria ou Lista_Original." });
    }

    const nomeNormalizado = normalizarTexto(data.nome_principal);
    if (!nomeNormalizado) {
      return jsonResponse({ status: "error", message: "Nome principal não recebido." });
    }

    const convite = buscarConvitePorNome(sheetLista, nomeNormalizado);
    if (!convite) {
      return jsonResponse({ status: "error", message: "Convite não encontrado na lista original." });
    }

    const convidadosPermitidosNormalizados = convite.acompanhantes.map(function(nome) {
      return normalizarTexto(nome);
    });

    const acompanhantesRecebidos = Array.isArray(data.acompanhantes_que_vao) ? data.acompanhantes_que_vao : [];
    const acompanhantesValidos = [];
    const acompanhantesJaIncluidos = [];

    acompanhantesRecebidos.forEach(function(nomeAcomp) {
      const nomeAcompNormalizado = normalizarTexto(nomeAcomp);
      if (!nomeAcompNormalizado) return;
      if (!convidadosPermitidosNormalizados.includes(nomeAcompNormalizado)) return;
      if (acompanhantesJaIncluidos.includes(nomeAcompNormalizado)) return;

      const nomeOriginal = convite.acompanhantes.find(function(nomePermitido) {
        return normalizarTexto(nomePermitido) === nomeAcompNormalizado;
      });
      acompanhantesValidos.push(nomeOriginal || nomeAcomp);
      acompanhantesJaIncluidos.push(nomeAcompNormalizado);
    });

    if (acompanhantesValidos.length > convite.permitidos) {
      return jsonResponse({ status: "error", message: "Foram selecionados mais acompanhantes do que o permitido." });
    }

    const convidadoPrincipalVai = data.convidado_principal_vai === true;
    if (!convidadoPrincipalVai && acompanhantesValidos.length === 0) {
      return jsonResponse({ status: "error", message: "Selecione pelo menos uma pessoa para confirmar presença." });
    }

    const dataHoraAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    const confirmacao = buscarConfirmacaoPorNome(sheetConfirmados, normalizarTexto(convite.nome));
    if (confirmacao) {
      return jsonResponse({ status: "already_confirmed", message: "Presença já confirmada anteriormente." });
    }

    const nomesConfirmados = [];
    if (convidadoPrincipalVai) {
      nomesConfirmados.push(convite.nome);
      sheetPortaria.appendRow([convite.nome]);
    }

    acompanhantesValidos.forEach(function(nomeAcomp) {
      nomesConfirmados.push(nomeAcomp);
      sheetPortaria.appendRow([nomeAcomp]);
    });

    const linhaConfirmados = [
      convite.nome,
      convite.permitidos,
      acompanhantesValidos.length,
      dataHoraAtual,
      convidadoPrincipalVai ? "Sim" : "Não",
      acompanhantesValidos.join("; "),
      nomesConfirmados.join("; ")
    ];

    sheetConfirmados.appendRow(linhaConfirmados);

    return jsonResponse({ status: "success" });
  } catch (error) {
    return jsonResponse({ status: "error", message: error.toString() });
  }
}
