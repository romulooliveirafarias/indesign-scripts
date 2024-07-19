// Versão atualizada em 19.07.24 por Rômulo Farias
// Reinserção do trecho que remove seções da página dentrod a função encadear

#target indesign;
#targetengine "session";

var caminhoCSV;
var caminhoIndt;
var caminhoIndd;
var temCSV;
var temINDT;
var temINDD;

//----------------- INTERFACE -----------------//

var w = new Window("palette");
w.text = "Montar Blocos";
w.preferredSize.width = 300;

//Adicionando um dropdownlist e botões de seleção de pastas e arquivos
w.selectTipoBloco = w.add("dropdownlist", undefined, [
  "Multiplos Blocos",
  "Blocão"
]);
w.selectTipoBloco.selection = 0;
w.selectTipoBloco.alignment = ["fill", "top"];

var selectCSV = w.add("button", undefined, "Selecionar CSV");
selectCSV.alignment = ["fill", "top"];

var selectINDT = w.add("button", undefined, "Selecionar INDT");
selectINDT.alignment = ["fill", "top"];

var selectINDD = w.add("button", undefined, "Selecionar INDD");
selectINDD.alignment = ["fill", "top"];

//Adicionando um grupo e um painel
w.gp = w.add("group");
w.gp.orientation = "stack";
w.gp.alignment = ["fill", "top"];
w.gp.alignChildren = ["fill", "top"];
w.gp.p1 = w.gp.add("panel", undefined, "Múltiplos blocos");
w.gp.p1.spacing = 10;
w.gp.p1.margins = 10;

// Adicionando botão "Gerar" para iniciar as funções
var gerarButton = w.add("button", undefined, "Gerar");
gerarButton.alignment = ["middle", "top"];
gerarButton.size = [100, 40];
gerarButton.enabled = false;

//Adicionando os textos estáticos para dar  feedback das seleções
var staticTextCSV = w.gp.p1.add("statictext", undefined, "");
staticTextCSV.alignment = ["fill", "top"];

var staticTextINDT = w.gp.p1.add("statictext", undefined, "");
staticTextINDT.alignment = ["fill", "top"];

var staticTextINDD = w.gp.p1.add("statictext", undefined, "");
staticTextINDD.alignment = ["fill", "top"];

//----------------- FUNÇÕES QUE MODIFICAM A INTERFACE -----------------//

//Função que seleciona CSV
selectCSV.onClick = function () {
  caminhoCSV = File.openDialog("Selecione o arquivo CSV (*.csv)", "*.csv");

  if (caminhoCSV) {
    staticTextCSV.text = "Caminho CSV: OK ";
    temCSV = 1;
    checarCaminhos();
  }
};

//Função que seleciona pasta com INDT
selectINDT.onClick = function () {
  caminhoIndt = Folder.selectDialog(
    "Selecione a pasta contendo o arquivo Bloco.indt"
  ); // Retorna String com caminho

  if (caminhoIndt) {
    if (File(caminhoIndt + "/Bloco.indt").exists) {
      staticTextINDT.text = "Caminho INDT: OK ";
      temINDT = 1;
      checarCaminhos();
    } else {
      alert(
        "O arquivo Bloco.indt não está presente na pasta selecionada. Tente novamente. "
      );
    }
  }
};

//Função que seleciona pasta onde estão os itens em INDD
selectINDD.onClick = function () {
  caminhoIndd = Folder.selectDialog(
    "Selecione a pasta com os arquivos em .indd"
  ); // Retorna String com caminho

  if (caminhoIndd) {
    staticTextINDD.text = "Caminho INDDs: OK ";
    checarArquivosIndd();
    temINDD = 1;
    checarCaminhos();
  }
};

// Função que modifica o painel a partir da seleção do dropdownlist
w.selectTipoBloco.onChange = function () {
  temCSV = 0;
  temINDT = 0;
  temINDD = 0;
  checarCaminhos();
  if (w.selectTipoBloco.selection == 0) {
    // Múltiplos Blocos selecionado
    selectINDT.enabled = true;
    w.gp.p1.text = "Múltiplos Blocos";
    staticTextCSV.text = "";
    staticTextINDT.text = "";
    staticTextINDD.text = "";
    w.layout.layout(1);
  } else {
    // Blocão selecionado
    w.gp.p1.text = "Blocão";
    selectINDT.enabled = false;
    staticTextCSV.text = "";
    staticTextINDT.text = "";
    staticTextINDD.text = "";

    if (app.documents.length > 0) {
      alert(
        "O arquivo " +
          app.activeDocument.name +
          "está aberto bara geração do bloco."
      );
    } else {
      alert("Não ha arquivos abertos. Abra um arquivo para continuar.");
    }
    staticTextINDT.text = app.activeDocument.name;
  }
};

// Função do botão gerar
gerarButton.onClick = function () {
  gerarBlocos();
};

w.show();

//----------------- INTERFACE DA JANELA DE PROGRESSO -----------------//
var janelaProgresso = new Window("palette");
janelaProgresso.text = "Progresso";
janelaProgresso.preferredSize.width = 600;

var staticTextProgresso = janelaProgresso.add("statictext", undefined, "");
staticTextProgresso.preferredSize.width = 600;
staticTextProgresso.justify = "center";
staticTextProgresso.text = "Executando";

// Adicionando botão "Gerar" para iniciar as funções
var okProgresso = janelaProgresso.add("button", undefined, "OK");
okProgresso.enabled = false;

okProgresso.onClick = function () {
  janelaProgresso.close();
};

//----------------- FUNÇÃO PRINCIPAL -----------------//

//Função que gera os blocos
function gerarBlocos() {
  w.close();
  janelaProgresso.show();

  // Passo 1: Ler o conteúdo do arquivo CSV
  // Passo 2: Abrir o arquivo INDT para cada linha do CSV ou usar o arquivo aberto
  // Passo 3: Gerar o bloco para cada linha do CSV
  // Passo 4: Salvar o arquivo INDD para cada linha do CSV

  // Executando o passo 1 - Ler o conteúdo do arquivo CSV
  var linhasCSV = lerArquivoCSV();

  for (var i = 0; i < linhasCSV.length - 1; i++) {
    var linhaBloco = linhasCSV[i];
    var nomeBloco = linhaBloco.split(";")[0];
    staticTextProgresso.text = "Gerando o bloco " + nomeBloco;
    var itensBloco = linhaBloco.split(";");
    var documento;

    // Executando o passo 2 - Abrir o arquivo INDT para cada linha do CSV ou usar o arquivo aberto
    var arquivoIndt = caminhoIndt + "/Bloco.indt";
    if (w.selectTipoBloco.selection == 0) {
      var documento = app.open(File(arquivoIndt));
    } else {
      var documento = app.activeDocument;
    }

    var docName = documento.name;

    for (var j = 1; j < itensBloco.length; j++) {
      var item = itensBloco[j];
      if (item.match(".indd")) {
        var arquivoIndd = caminhoIndd + "/" + item;
        // Executando o passo 3 - Gerar o bloco para cada linha do CSV
        app.open(File(arquivoIndd));

        var blocoArquivo = app.documents.item(docName);
        var itemArquivo = app.documents.item(item);
        var itemPagina = itemArquivo.pages.item(0);
        itemPagina.duplicate(LocationOptions.AFTER, itemArquivo.pages.item(0));
        itemPagina.move(LocationOptions.AFTER, blocoArquivo.pages.item(-1));

        // Adicionar uma linha em branco separadora ao final do texto na última caixa de texto da página atual
        var lastTextFrame =
          blocoArquivo.pages.item(-1).textFrames[
            blocoArquivo.pages.item(-1).textFrames.length - 1
          ];
        if (lastTextFrame.parentStory) {
          lastTextFrame.parentStory.insertionPoints[-1].contents = "\r"; // Adiciona um caractere de quebra de linha
        }

        // Fecha o arquivo que foi aberto sem salvar (para evitar problemas de memória)
        itemArquivo.close(SaveOptions.NO);
      }
    }
    encadear(docName);
    removerMasters(docName);
    // Executando o passo 4 - Salvar o arquivo INDD para cada linha do CSV
    salvarBloco(docName, nomeBloco);
  }
  staticTextProgresso.text = "Todos os blocos foram gerados";
  okProgresso.enabled = true;
}

//----------------- FUNÇÕES GERAIS -----------------//

// Função para checar as seleções do usuário para ativar/desativar o botão gerar
function checarCaminhos() {
  if (w.selectTipoBloco.selection == 0) {
    if (temCSV == 1 && temINDT == 1 && temINDD == 1) {
      gerarButton.enabled = true;
    } else {
      gerarButton.enabled = false;
    }
  } else {
    if (temCSV == 1 && temINDD == 1) {
      gerarButton.enabled = true;
    } else {
      gerarButton.enabled = false;
    }
  }
}

// Função para ler o conteúdo de um arquivo CSV e armazenar as linhas
function lerArquivoCSV() {
  var arquivo = File(caminhoCSV);
  arquivo.open("r");
  var linhasbloco = [];
  var linhas = arquivo.read().split("\n");
  arquivo.close();
  return linhas;
}

// Função para encadear as caixas de texto
function encadear(documento) {
  var docBloco = app.documents.item(documento);

  var docBlocoPaginas = docBloco.pages;

  for (var p = 1; p < docBlocoPaginas.length; p++) {
    var previousPage = docBlocoPaginas[p - 1];
    var currentPage = docBlocoPaginas[p];
    var lastTextFrame =
      previousPage.textFrames[previousPage.textFrames.length - 1];
    var firstTextFrame = currentPage.textFrames[0];
    lastTextFrame.nextTextFrame = firstTextFrame;
  }

  // Remova a seção da página
if (currentPage.appliedSection) {
  currentPage.appliedSection.remove();
}
}

//Função que remove as páginas mestras excessivas
function removerMasters(documento) {
  var docBloco = app.documents.item(documento);
  var masters = docBloco.masterSpreads;

  for (i = masters.length - 1; i > 1; i--) {
    masters[i].remove();
  }
}

//Função que salva o arquivo INDD
function salvarBloco(documento, bloco) {
  var docBloco = app.documents.item(documento);
  docBloco.save(File(caminhoIndt + "/" + bloco + ".indd"));
  docBloco.close(SaveOptions.YES);
}

//Função que checa se os arquivos INDD listados no CSV estão todos presentes na pasta informada pleo usuário.
function checarArquivosIndd() {
  var itensFaltantes = [];
  var arquivoCSV = File(caminhoCSV);
  arquivoCSV.open("r");
  var linhasCSV = arquivoCSV.read().split("\n");
  for (var i = 0; i < linhasCSV.length - 1; i++) {
    var linhaCSV = linhasCSV[i].split(";");
    for (var k = 1; k < linhaCSV.length; k++) {
      var itemCSV = linhaCSV[k];

      if (itemCSV.match(".indd")) {
        var arquivoIndd = caminhoIndd + "/" + itemCSV;
        if (File(arquivoIndd).exists) {
        } else {
          itensFaltantes.push(itemCSV);
        }
      }
    }
  }

  function removeDuplicate(a) {
    for (var i = 0; i < a.length; i++) {
      for (var j = 0; j < a.length; j++) {
        if (i != j && a[i] == a[j]) {
          a.splice(j, 1);
        }
      }
    }
    return a;
  }

  var itensFaltantesUnicos = removeDuplicate(itensFaltantes);

  if (itensFaltantes.length > 0) {
    alert(
      "Os seguintes itens estão faltando na pasta selecionada: " +
        itensFaltantesUnicos
    );
  } else {
    alert(
      "Todos os itens listados no CSV estão presentes na pasta selecionada."
    );
  }
}
