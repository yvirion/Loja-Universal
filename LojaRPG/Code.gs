/*******************************************************
 * 🏰 LOJA MEDIEVAL DE RPG
 * Backend - Google Apps Script
 *
 * Banco de dados:
 *   JOGADORES
 *   ESTOQUE
 *   INVENTARIO
 *   TRANSACOES
 *
 * Frontend:
 *   Index.html
 *   CSS.html
 *   JS.html
 *******************************************************/


/* =====================================================
   CONFIGURAÇÕES
   ===================================================== */

const CONFIG = {
  NOME_LOJA: "A Loja do Dragão Dourado",

  TEMPO_SESSAO: 21600, // 6 horas

  SHEETS: {
    JOGADORES: "JOGADORES",
    ESTOQUE: "ESTOQUE",
    INVENTARIO: "INVENTARIO",
    TRANSACOES: "TRANSACOES"
  }
};


/* =====================================================
   WEB APP
   ===================================================== */

function doGet() {

  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle(CONFIG.NOME_LOJA)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/* =====================================================
   INCLUIR ARQUIVOS HTML
   ===================================================== */

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}


/* =====================================================
   CONFIGURAÇÃO INICIAL
   ===================================================== */

function configurarSistema() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAba(
    ss,
    CONFIG.SHEETS.JOGADORES,
    [
      "ID",
      "NOME",
      "USUARIO",
      "SENHA",
      "OURO",
      "PRATA",
      "COBRE",
      "STATUS"
    ]
  );


  criarAba(
    ss,
    CONFIG.SHEETS.ESTOQUE,
    [
      "ID",
      "NOME",
      "CATEGORIA",
      "DESCRICAO",
      "PRECO",
      "ESTOQUE",
      "IMAGEM",
      "ATIVO"
    ]
  );


  criarAba(
    ss,
    CONFIG.SHEETS.INVENTARIO,
    [
      "ID",
      "JOGADOR_ID",
      "ITEM_ID",
      "ITEM_NOME",
      "QUANTIDADE"
    ]
  );


  criarAba(
    ss,
    CONFIG.SHEETS.TRANSACOES,
    [
      "ID",
      "DATA",
      "TIPO",
      "JOGADOR_ID",
      "JOGADOR",
      "ITEM_ID",
      "ITEM",
      "QUANTIDADE",
      "VALOR",
      "SALDO_ANTES",
      "SALDO_DEPOIS",
      "RESPONSAVEL"
    ]
  );


  criarDadosExemplo(ss);

  return {
    sucesso: true,
    mensagem: "Sistema configurado com sucesso."
  };
}


/* =====================================================
   CRIAR ABA
   ===================================================== */

function criarAba(ss, nome, cabecalhos) {

  let sheet = ss.getSheetByName(nome);

  if (!sheet) {

    sheet = ss.insertSheet(nome);

    sheet
      .getRange(1, 1, 1, cabecalhos.length)
      .setValues([cabecalhos]);

    sheet
      .getRange(1, 1, 1, cabecalhos.length)
      .setFontWeight("bold");

    sheet.setFrozenRows(1);
  }
}


/* =====================================================
   DADOS DE EXEMPLO
   ===================================================== */

function criarDadosExemplo(ss) {

  const jogadores =
    ss.getSheetByName(CONFIG.SHEETS.JOGADORES);

  const estoque =
    ss.getSheetByName(CONFIG.SHEETS.ESTOQUE);


  /*
   * Verifica se já existem jogadores.
   */

  if (jogadores.getLastRow() <= 1) {

    jogadores.getRange(2, 1, 3, 8).setValues([

      [
        "P001",
        "Leonardo",
        "leonardo",
        "1234",
        100,
        20,
        50,
        "ATIVO"
      ],

      [
        "P002",
        "Aragorn",
        "aragorn",
        "1234",
        80,
        10,
        0,
        "ATIVO"
      ],

      [
        "P003",
        "Gandalf",
        "gandalf",
        "1234",
        500,
        50,
        0,
        "ATIVO"
      ]

    ]);
  }


  /*
   * Produtos de exemplo.
   */

  if (estoque.getLastRow() <= 1) {

    estoque.getRange(2, 1, 8, 8).setValues([

      [
        "I001",
        "Espada Longa",
        "Armas",
        "Uma espada confiável para aventureiros.",
        50,
        8,
        "",
        "SIM"
      ],

      [
        "I002",
        "Escudo de Ferro",
        "Armaduras",
        "Escudo pesado feito de ferro.",
        35,
        5,
        "",
        "SIM"
      ],

      [
        "I003",
        "Poção de Cura",
        "Poções",
        "Recupera parte dos pontos de vida.",
        15,
        20,
        "",
        "SIM"
      ],

      [
        "I004",
        "Arco Élfico",
        "Armas",
        "Arco de madeira élfica.",
        120,
        2,
        "",
        "SIM"
      ],

      [
        "I005",
        "Adaga Sombria",
        "Armas",
        "Uma pequena lâmina de aparência misteriosa.",
        40,
        6,
        "",
        "SIM"
      ],

      [
        "I006",
        "Cota de Malha",
        "Armaduras",
        "Proteção utilizada por guerreiros.",
        150,
        3,
        "",
        "SIM"
      ],

      [
        "I007",
        "Pergaminho de Fogo",
        "Magia",
        "Contém uma magia ofensiva.",
        75,
        4,
        "",
        "SIM"
      ],

      [
        "I008",
        "Corda de Cânhamo",
        "Utilidades",
        "Uma corda resistente de 15 metros.",
        10,
        15,
        "",
        "SIM"
      ]

    ]);
  }
}


/* =====================================================
   LOGIN
   ===================================================== */

function login(usuario, senha) {

  usuario = String(usuario || "").trim();
  senha = String(senha || "").trim();

  if (!usuario || !senha) {

    return {
      sucesso: false,
      mensagem: "Informe usuário e senha."
    };
  }


  /*
   * MESTRE
   *
   * Para a versão inicial:
   *
   * usuário: mestre
   * senha:   dragao
   *
   * Em uma versão posterior podemos migrar isso
   * para uma tabela própria de administradores.
   */

  if (usuario === "mestre" && senha === "dragao") {

    const token = criarSessao({
      tipo: "MESTRE",
      usuario: "mestre",
      jogadorId: null
    });

    return {
      sucesso: true,
      token: token,
      tipo: "MESTRE",
      usuario: "mestre"
    };
  }


  /*
   * LOGIN DOS JOGADORES
   */

  const jogador =
    encontrarJogadorPorUsuario(usuario);

  if (!jogador) {

    return {
      sucesso: false,
      mensagem: "Jogador não encontrado."
    };
  }


  if (String(jogador.SENHA) !== senha) {

    return {
      sucesso: false,
      mensagem: "Senha incorreta."
    };
  }


  if (String(jogador.STATUS).toUpperCase() !== "ATIVO") {

    return {
      sucesso: false,
      mensagem: "Este jogador está desativado."
    };
  }


  const token = criarSessao({

    tipo: "JOGADOR",

    usuario: jogador.USUARIO,

    jogadorId: jogador.ID
  });


  return {

    sucesso: true,

    token: token,

    tipo: "JOGADOR",

    jogadorId: jogador.ID,

    usuario: jogador.USUARIO,

    nome: jogador.NOME
  };
}


/* =====================================================
   SESSÃO
   ===================================================== */

function criarSessao(dados) {

  const token =
    Utilities.getUuid();

  const cache =
    CacheService.getScriptCache();

  cache.put(
    "SESSION_" + token,
    JSON.stringify(dados),
    CONFIG.TEMPO_SESSAO
  );

  return token;
}


function obterSessao(token) {

  if (!token) {

    throw new Error(
      "Sessão inválida."
    );
  }


  const cache =
    CacheService.getScriptCache();

  const dados =
    cache.get("SESSION_" + token);


  if (!dados) {

    throw new Error(
      "Sessão expirada. Faça login novamente."
    );
  }


  return JSON.parse(dados);
}


function logout(token) {

  if (token) {

    CacheService
      .getScriptCache()
      .remove("SESSION_" + token);
  }

  return {
    sucesso: true
  };
}


/* =====================================================
   LISTAR PRODUTOS
   ===================================================== */

function listarProdutos(token) {

  const sessao =
    obterSessao(token);

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.ESTOQUE);


  const dados =
    sheet
      .getDataRange()
      .getValues();


  if (dados.length <= 1) {

    return [];
  }


  const cabecalhos =
    dados[0];


  return dados
    .slice(1)
    .filter(linha => {

      return String(linha[7]).toUpperCase() === "SIM";
    })
    .map(linha => {

      const item = {};

      cabecalhos.forEach((cab, i) => {

        item[cab] = linha[i];

      });

      return item;
    });
}


/* =====================================================
   LISTAR CATEGORIAS
   ===================================================== */

function listarCategorias(token) {

  obterSessao(token);

  const produtos =
    listarProdutos(token);

  const categorias =
    produtos.map(
      item => item.CATEGORIA
    );

  return [
    ...new Set(categorias)
  ];
}


/* =====================================================
   DADOS DO JOGADOR
   ===================================================== */

function obterDadosJogador(token) {

  const sessao =
    obterSessao(token);


  if (sessao.tipo !== "JOGADOR") {

    throw new Error(
      "Apenas jogadores possuem esta área."
    );
  }


  const jogador =
    encontrarJogadorPorId(
      sessao.jogadorId
    );


  if (!jogador) {

    throw new Error(
      "Jogador não encontrado."
    );
  }


  const inventario =
    obterInventarioJogador(
      sessao.jogadorId
    );


  return {

    id: jogador.ID,

    nome: jogador.NOME,

    usuario: jogador.USUARIO,

    ouro: Number(jogador.OURO) || 0,

    prata: Number(jogador.PRATA) || 0,

    cobre: Number(jogador.COBRE) || 0,

    inventario: inventario
  };
}


/* =====================================================
   INVENTÁRIO
   ===================================================== */

function obterInventarioJogador(jogadorId) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.INVENTARIO);


  const dados =
    sheet
      .getDataRange()
      .getValues();


  if (dados.length <= 1) {

    return [];
  }


  return dados
    .slice(1)
    .filter(linha => {

      return String(linha[1]) ===
        String(jogadorId);
    })
    .map(linha => {

      return {

        id: linha[0],

        jogadorId: linha[1],

        itemId: linha[2],

        itemNome: linha[3],

        quantidade: Number(linha[4]) || 0

      };
    });
}


/* =====================================================
   COMPRA
   ===================================================== */

function comprarItem(token, itemId, quantidade) {

  const sessao =
    obterSessao(token);


  if (sessao.tipo !== "JOGADOR") {

    throw new Error(
      "Somente jogadores podem comprar itens."
    );
  }


  quantidade =
    Number(quantidade);


  if (!Number.isInteger(quantidade) ||
      quantidade <= 0) {

    throw new Error(
      "Quantidade inválida."
    );
  }


  /*
   * Lock evita problemas caso dois jogadores
   * tentem comprar simultaneamente.
   */

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);


  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const jogadores =
      ss.getSheetByName(
        CONFIG.SHEETS.JOGADORES
      );


    const estoque =
      ss.getSheetByName(
        CONFIG.SHEETS.ESTOQUE
      );


    const inventario =
      ss.getSheetByName(
        CONFIG.SHEETS.INVENTARIO
      );


    const transacoes =
      ss.getSheetByName(
        CONFIG.SHEETS.TRANSACOES
      );


    /*
     * Encontrar jogador
     */

    const jogador =
      encontrarJogadorPorId(
        sessao.jogadorId
      );


    if (!jogador) {

      throw new Error(
        "Jogador não encontrado."
      );
    }


    /*
     * Encontrar item
     */

    const item =
      encontrarItemPorId(itemId);


    if (!item) {

      throw new Error(
        "Item não encontrado."
      );
    }


    /*
     * Estoque
     */

    const estoqueAtual =
      Number(item.ESTOQUE) || 0;


    if (estoqueAtual < quantidade) {

      throw new Error(
        "Estoque insuficiente."
      );
    }


    /*
     * PREÇO
     */

    const preco =
      Number(item.PRECO) || 0;


    const valorTotal =
      preco * quantidade;


    /*
     * SALDO
     *
     * Nesta versão o preço está em OURO.
     */

    const saldoAntes =
      Number(jogador.OURO) || 0;


    if (saldoAntes < valorTotal) {

      throw new Error(
        "Você não possui ouro suficiente."
      );
    }


    const saldoDepois =
      saldoAntes - valorTotal;


    /*
     * ATUALIZA JOGADOR
     */

    jogadores
      .getRange(
        jogador.LINHA,
        5
      )
      .setValue(saldoDepois);


    /*
     * ATUALIZA ESTOQUE
     */

    estoque
      .getRange(
        item.LINHA,
        6
      )
      .setValue(
        estoqueAtual - quantidade
      );


    /*
     * ATUALIZA INVENTÁRIO
     */

    adicionarAoInventario(

      inventario,

      sessao.jogadorId,

      item.ID,

      item.NOME,

      quantidade
    );


    /*
     * REGISTRA TRANSAÇÃO
     */

    transacoes.appendRow([

      Utilities.getUuid(),

      new Date(),

      "COMPRA",

      jogador.ID,

      jogador.NOME,

      item.ID,

      item.NOME,

      quantidade,

      -valorTotal,

      saldoAntes,

      saldoDepois,

      jogador.NOME

    ]);


    return {

      sucesso: true,

      mensagem:
        "Compra realizada com sucesso!",

      item: item.NOME,

      quantidade: quantidade,

      valor: valorTotal,

      saldo: saldoDepois

    };


  } finally {

    lock.releaseLock();

  }
}


/* =====================================================
   ADICIONAR AO INVENTÁRIO
   ===================================================== */

function adicionarAoInventario(
  sheet,
  jogadorId,
  itemId,
  itemNome,
  quantidade
) {

  const dados =
    sheet
      .getDataRange()
      .getValues();


  /*
   * Procurar item existente.
   */

  for (let i = 1; i < dados.length; i++) {

    const jogador =
      String(dados[i][1]);

    const item =
      String(dados[i][2]);


    if (
      jogador === String(jogadorId) &&
      item === String(itemId)
    ) {

      const linha =
        i + 1;


      const quantidadeAtual =
        Number(dados[i][4]) || 0;


      sheet
        .getRange(linha, 5)
        .setValue(
          quantidadeAtual + quantidade
        );


      return;
    }
  }


  /*
   * Criar novo registro.
   */

  sheet.appendRow([

    Utilities.getUuid(),

    jogadorId,

    itemId,

    itemNome,

    quantidade

  ]);
}


/* =====================================================
   MESTRE - LISTAR JOGADORES
   ===================================================== */

function listarJogadores(token) {

  const sessao =
    obterSessao(token);


  if (sessao.tipo !== "MESTRE") {

    throw new Error(
      "Acesso restrito ao mestre."
    );
  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.JOGADORES
      );


  const dados =
    sheet
      .getDataRange()
      .getValues();


  if (dados.length <= 1) {

    return [];
  }


  return dados
    .slice(1)
    .map(linha => {

      return {

        id: linha[0],

        nome: linha[1],

        usuario: linha[2],

        ouro: Number(linha[4]) || 0,

        prata: Number(linha[5]) || 0,

        cobre: Number(linha[6]) || 0,

        status: linha[7]

      };
    });
}


/* =====================================================
   MESTRE - ALTERAR OURO
   ===================================================== */

function alterarOuro(
  token,
  jogadorId,
  valor
) {

  const sessao =
    obterSessao(token);


  if (sessao.tipo !== "MESTRE") {

    throw new Error(
      "Acesso restrito ao mestre."
    );
  }


  valor =
    Number(valor);


  if (!Number.isFinite(valor) ||
      valor === 0) {

    throw new Error(
      "Valor inválido."
    );
  }


  const lock =
    LockService.getScriptLock();


  lock.waitLock(30000);


  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const sheet =
      ss.getSheetByName(
        CONFIG.SHEETS.JOGADORES
      );


    const jogador =
      encontrarJogadorPorId(
        jogadorId
      );


    if (!jogador) {

      throw new Error(
        "Jogador não encontrado."
      );
    }


    const saldoAntes =
      Number(jogador.OURO) || 0;


    const saldoDepois =
      saldoAntes + valor;


    if (saldoDepois < 0) {

      throw new Error(
        "O saldo não pode ficar negativo."
      );
    }


    sheet
      .getRange(
        jogador.LINHA,
        5
      )
      .setValue(
        saldoDepois
      );


    registrarTransacao(

      "MESTRE_MOEDAS",

      jogador,

      null,

      null,

      valor,

      saldoAntes,

      saldoDepois,

      "MESTRE"

    );


    return {

      sucesso: true,

      jogador: jogador.NOME,

      saldo: saldoDepois

    };


  } finally {

    lock.releaseLock();

  }
}


/* =====================================================
   MESTRE - ADICIONAR ESTOQUE
   ===================================================== */

function alterarEstoque(
  token,
  itemId,
  quantidade
) {

  const sessao =
    obterSessao(token);


  if (sessao.tipo !== "MESTRE") {

    throw new Error(
      "Acesso restrito ao mestre."
    );
  }


  quantidade =
    Number(quantidade);


  if (!Number.isInteger(quantidade) ||
      quantidade === 0) {

    throw new Error(
      "Quantidade inválida."
    );
  }


  const lock =
    LockService.getScriptLock();


  lock.waitLock(30000);


  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const sheet =
      ss.getSheetByName(
        CONFIG.SHEETS.ESTOQUE
      );


    const item =
      encontrarItemPorId(itemId);


    if (!item) {

      throw new Error(
        "Item não encontrado."
      );
    }


    const estoqueAntes =
      Number(item.ESTOQUE) || 0;


    const estoqueDepois =
      estoqueAntes + quantidade;


    if (estoqueDepois < 0) {

      throw new Error(
        "O estoque não pode ficar negativo."
      );
    }


    sheet
      .getRange(
        item.LINHA,
        6
      )
      .setValue(
        estoqueDepois
      );


    return {

      sucesso: true,

      item: item.NOME,

      estoque: estoqueDepois

    };


  } finally {

    lock.releaseLock();

  }
}


/* =====================================================
   HISTÓRICO
   ===================================================== */

function listarTransacoes(token) {

  const sessao =
    obterSessao(token);


  if (sessao.tipo !== "MESTRE") {

    throw new Error(
      "Acesso restrito ao mestre."
    );
  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.TRANSACOES
      );


  const dados =
    sheet
      .getDataRange()
      .getValues();


  if (dados.length <= 1) {

    return [];
  }


  return dados
    .slice(1)
    .reverse()
    .slice(0, 100)
    .map(linha => {

      return {

        data:
          Utilities.formatDate(
            new Date(linha[1]),
            Session.getScriptTimeZone(),
            "dd/MM/yyyy HH:mm"
          ),

        tipo: linha[2],

        jogador: linha[4],

        item: linha[6],

        quantidade: linha[7],

        valor: linha[8],

        saldo: linha[10],

        responsavel: linha[11]

      };
    });
}


/* =====================================================
   FUNÇÃO AUXILIAR - TRANSAÇÃO
   ===================================================== */

function registrarTransacao(
  tipo,
  jogador,
  itemId,
  itemNome,
  valor,
  saldoAntes,
  saldoDepois,
  responsavel
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.TRANSACOES
      );


  sheet.appendRow([

    Utilities.getUuid(),

    new Date(),

    tipo,

    jogador.ID,

    jogador.NOME,

    itemId || "",

    itemNome || "",

    "",

    valor,

    saldoAntes,

    saldoDepois,

    responsavel

  ]);
}


/* =====================================================
   BUSCAR JOGADOR
   ===================================================== */

function encontrarJogadorPorUsuario(usuario) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.JOGADORES
      );


  const dados =
    sheet
      .getDataRange()
      .getValues();


  for (let i = 1; i < dados.length; i++) {

    if (
      String(dados[i][2]).toLowerCase() ===
      String(usuario).toLowerCase()
    ) {

      return {

        ID: dados[i][0],

        NOME: dados[i][1],

        USUARIO: dados[i][2],

        SENHA: dados[i][3],

        OURO: dados[i][4],

        PRATA: dados[i][5],

        COBRE: dados[i][6],

        STATUS: dados[i][7],

        LINHA: i + 1

      };
    }
  }


  return null;
}


/* =====================================================
   BUSCAR JOGADOR POR ID
   ===================================================== */

function encontrarJogadorPorId(id) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.JOGADORES
      );


  const dados =
    sheet
      .getDataRange()
      .getValues();


  for (let i = 1; i < dados.length; i++) {

    if (
      String(dados[i][0]) ===
      String(id)
    ) {

      return {

        ID: dados[i][0],

        NOME: dados[i][1],

        USUARIO: dados[i][2],

        SENHA: dados[i][3],

        OURO: dados[i][4],

        PRATA: dados[i][5],

        COBRE: dados[i][6],

        STATUS: dados[i][7],

        LINHA: i + 1

      };
    }
  }


  return null;
}


/* =====================================================
   BUSCAR ITEM
   ===================================================== */

function encontrarItemPorId(id) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.ESTOQUE
      );


  const dados =
    sheet
      .getDataRange()
      .getValues();


  for (let i = 1; i < dados.length; i++) {

    if (
      String(dados[i][0]) ===
      String(id)
    ) {

      return {

        ID: dados[i][0],

        NOME: dados[i][1],

        CATEGORIA: dados[i][2],

        DESCRICAO: dados[i][3],

        PRECO: dados[i][4],

        ESTOQUE: dados[i][5],

        IMAGEM: dados[i][6],

        ATIVO: dados[i][7],

        LINHA: i + 1

      };
    }
  }


  return null;
}
