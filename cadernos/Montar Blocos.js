#target indesign;
#targetengine "session";

var caminhoCSV;
var caminhoIndt;
var caminhoIndd;

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

//------------------------------------------------------------------------------------------------------------------------------------------------------------------------

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
gerarButton.onClick = function () {
  if (w.selectTipoBloco.selection == 0) {
    // Opção A selecionada
    alert("selecionado A");
    /* menu1(caminhoCSV, caminhoIndt, caminhoIndd); */
  } else {
    alert("selecionado B");
    /* menu2(caminhoCSV, caminhoIndd); */
  }
};

//------------------------------------------------------------//
//Adicionando os textos estáticos para dar  feedback das seleções
var staticTextBlocao = w.gp.p1.add("statictext", undefined, "");
    staticTextBlocao.alignment = ["fill", "top"];

var staticTextCSV = w.gp.p1.add("statictext", undefined, "");
    staticTextCSV.alignment = ["fill", "top"];

var staticTextINDT = w.gp.p1.add("statictext", undefined, "");
    staticTextINDT.alignment = ["fill", "top"];

var staticTextINDD = w.gp.p1.add("statictext", undefined, "");
    staticTextINDD.alignment = ["fill", "top"];

//Adicionando funções de seleção de pastas e arquivos
selectCSV.onClick = function () {
    caminhoCSV = File.openDialog("Selecione o arquivo CSV (*.csv)", "*.csv");
  
    if (caminhoCSV) {
        staticTextCSV.text = "Caminho CSV: OK ";
    }
  };

selectINDT.onClick = function () {
    caminhoIndt = Folder.selectDialog(
      "Selecione a pasta contendo o arquivo Bloco.indt"
    ); // Retorna String com caminho
  
    if (caminhoIndt) {
        staticTextINDT.text = "Caminho INDT: OK ";
  
    }
  };

selectINDD.onClick = function () {
    caminhoIndd = Folder.selectDialog(
      "Selecione a pasta com os arquivos em .indd"
    ); // Retorna String com caminho
  
    if (caminhoIndd) {
        staticTextINDD.text = "Caminho INDDs: OK ";
    }
  };

w.selectTipoBloco.onChange = function () {
  if (w.selectTipoBloco.selection == 0) {
    // Múltiplos Blocos selecionado
    selectINDT.enabled = true;
    w.gp.p1.text = "Múltiplos Blocos";
    staticTextBlocao.text = "";
    staticTextBlocao.maximumSize.height = 0;
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
    staticTextBlocao.maximumSize.height = 40;

    if (app.documents.length > 0) {
        staticTextBlocao.maximumSize.height = 40;
        staticTextBlocao.text  = "Arquivo do Blocão Aberto"
        } else {
        alert('Não ha arquivos abertos. Abra um arquivo para continuar.')
        staticTextBlocao.maximumSize.height = 40;
        staticTextBlocao.text  = "Não esquecer: Gerar com o .indd aberto"
        }
  }
};

w.show();

// Função para ler o conteúdo de um arquivo CSV e armazenar as linhas
function lerArquivoCSV() {
  var arquivo = File(caminhoCSV);
  arquivo.open("r");
  var linhasbloco = [];

  if (!caminhoCSV) {
    // Encerrar o script, já que não há arquivo selecionado
    alert("Nenhum arquivo CSV selecionado.");
    exit();
  } else {
    var linhas = arquivo.read().split("\n"); // Verificar melhor uso.
    for (var i = 0; i < linhas.length; i++) {
      // Dividir a linha em elementos com base em uma vírgula (pode ser ajustado)
      var elementos = linhas[i].split(",");
      // Adicionar os elementos ao array linhasbloco
      linhasbloco.push(elementos);
    }
    arquivo.close();

    return linhasbloco;
  }
}

//--------------------------------------------Montagem de vários blocos-------------------------------------------------------------------------------

function menu1(caminhoCSV, caminhoIndt, caminhoIndd) {
  // Função para ler o conteúdo de um arquivo de texto e armazenar os nomes dos arquivos

  // Função para manipular o conteúdo do documento InDesign
  function manipularDocumentoIndesign(
    arquivoIndt,
    listaLinha,
    blocos,
    caminhoIndd
  ) {
    w.close();

    documento = app.open(File(arquivoIndt));

    // lista os arquivos de caminhoIndd
    var fileList = caminhoIndd.getFiles();

    //Ordena os arquivos
    fileList.sort();

    // Filtrando os arquivos existentes na pasta de origem (parte copiada)
    var ListaArqPast = []; // Retorna lista filtrada
    var naoindd = []; // Outros arquivos

    //adiciona a lista sortedFileList, arquivos indd e a naoindd arquivos diversos
    for (var h = 0; h <= fileList.length; h++) {
      var file = fileList[h];
      if (file instanceof File && file.name.match(/\.indd$/i)) {
        ListaArqPast.push(file);
      } else {
        naoindd.push(file);
      }
    }

    // Verificado 20032024 12:37
    var verifyListCsv = [];
    // Ordenando os arquivos INDD de acordo com a ordem especificada no .csv (parte copiada)
    // Linhas separada por ";" e em nova Array

    for (var t = 0; t < listaLinha.length; t++) {
      var LinhaSep = [];
      // Dividir a string em valores usando o separador ";"
      var valores = listaLinha[t].split(";");

      // Adicionar os valores à nova array
      for (var z = 0; z < valores.length; z++) {
        LinhaSep.push(valores[z]);
      }
    }

    for (var j = 0; j < LinhaSep.length; j++) {
      var fileName = LinhaSep[j];

      for (var f = 0; f < ListaArqPast.length; f++) {
        if (ListaArqPast[f].name.match(new RegExp("^" + fileName + "$", "i"))) {
          verifyListCsv.push(ListaArqPast[f]);
          break;
        }
      }
    }

    // percorrer cada arquivo (parte copiada)
    for (var k = 0; k < verifyListCsv.length; k++) {
      var source_file = verifyListCsv[k]; //Pega posição K, e ordena

      if (source_file instanceof File && source_file.name.match(/\.indd$/i)) {
        app.open(source_file);
        var source_doc = app.documents.item(source_file.name);
        var sourcePages = source_doc.pages.item(0);

        // Duplica o arquivo original (Com os erros) e ele move para
        // o arquivo de destino.
        sourcePages.duplicate(LocationOptions.AFTER, source_doc.pages.item(0));
        sourcePages.move(LocationOptions.AFTER, documento.pages.item(-1));

        // Adicionar uma linha em branco separadora ao final do texto na última caixa de texto da página atual
        var lastTextFrame =
          documento.pages.item(-1).textFrames[
            documento.pages.item(-1).textFrames.length - 1
          ];
        if (lastTextFrame.parentStory) {
          lastTextFrame.parentStory.insertionPoints[-1].contents = "\r"; // Adiciona um caractere de quebra de linha
        }

        // Fecha o arquivo que foi aberto sem salvar (para evitar problemas de memória)
        app.activeDocument.close(SaveOptions.NO);
      }
    }

    // Encadeamento de caixas de texto
    var destinationPages = documento.pages;
    for (var p = 1; p < destinationPages.length; p++) {
      var previousPage = destinationPages[p - 1];
      var currentPage = destinationPages[p];
      var lastTextFrame =
        previousPage.textFrames[previousPage.textFrames.length - 1];
      var firstTextFrame = currentPage.textFrames[0];
      lastTextFrame.nextTextFrame = firstTextFrame;
    }

    // Verifique se o aplicativo InDesign está aberto
    if (app && app.name == "Adobe InDesign") {
      // Verifique se há um documento aberto
      if (app.documents.length > 0) {
        var doc = app.activeDocument;

        // Verifique se há mais de uma página no documento
        if (doc.pages.length > 1) {
          // Loop através de todas as páginas, exceto a primeira
          for (var o = 1; o < doc.pages.length; o++) {
            var currentPage = doc.pages[o];

            // Remova a seção da página
            if (currentPage.appliedSection) {
              currentPage.appliedSection.remove();
            }

            var masters = doc.masterSpreads;

            for (i = masters.length - 1; i > 1; i--) {
              masters[i].remove();
            }
          }
        }
      }
    }

    // Salva o documento original com o nome baseado em 'blocos'
    var pasta = new File(arquivoIndt).parent;
    documento.save(File(pasta + "/" + blocos + ".indd"));
    documento.close(SaveOptions.YES);
  }

  function laco(listaLinhas, blocos, caminhoIndt, caminhoIndd) {
    //Etapa 4: Início do loop de criação de blocos

    if (File(caminhoIndt + "/Bloco.indt").exists) {
      arquivoIndt = caminhoIndt + "/Bloco.indt"; //defindo caminho bloco
    } else {
      alert("O arquivo Bloco.indt não está presente na pasta.");
      exit();
    }

    // asubtração do tamanho da lista se deve a linha em braco
    for (var i = 0; i < listaLinhas.length - 1; i++) {
      //Etapa 5: Manipular o arquivo InDesign
      manipularDocumentoIndesign(
        arquivoIndt,
        listaLinhas[i],
        blocos[i],
        caminhoIndd
      );
    }
    // Alerta de fim de execução

    alert("Blocos montados");
    exit();
  }
  var blocos = lerArquivoTexto(); //Recebe lista (X,Y,Z)

  var listaLinhas = lerArquivoCSV(); //Recebe tupla [(a,b,c),(d,e,f),(g,h,i,j)]

  laco(listaLinhas, blocos, caminhoIndt, caminhoIndd);
}

//-------------------------------------------------------------------------------  Bloção  -------------------------------------------------------------------------------

function menu2(caminhoCSV, caminhoIndd) {
  w.close();

  // Função para manipular o conteúdo do documento InDesign

  function manipularDocumentoIndesign2(
    arquivoIndt2,
    listaLinha2,
    caminhoIndd2
  ) {
    w.close();

    documento2 = arquivoIndt2;

    // lista os arquivos de caminhoIndd
    var fileList2 = caminhoIndd2.getFiles();

    //Ordena os arquivos
    fileList2.sort();

    // Filtrando os arquivos existentes na pasta de origem (parte copiada)
    var ListaArqPast2 = []; // Retorna lista filtrada
    var naoindd2 = []; // Outros arquivos

    //adiciona a lista sortedFileList, arquivos indd e a naoindd arquivos diversos
    for (var h = 0; h <= fileList2.length; h++) {
      var file = fileList2[h];
      if (file instanceof File && file.name.match(/\.indd$/i)) {
        ListaArqPast2.push(file);
      } else {
        naoindd2.push(file);
      }
    }

    // Verificado 20032024 12:37
    var verifyListCsv2 = [];
    // Ordenando os arquivos INDD de acordo com a ordem especificada no .csv (parte copiada)
    // Linhas separada por ";" e em nova Array

    for (var t = 0; t < listaLinha2.length; t++) {
      var LinhaSep2 = [];
      // Dividir a string em valores usando o separador ";"
      var valores2 = listaLinha2[t].split(";");

      // Adicionar os valores à nova array
      for (var z = 0; z < valores2.length; z++) {
        LinhaSep2.push(valores2[z]);
      }
    }

    for (var j = 0; j < LinhaSep2.length; j++) {
      var fileName2 = LinhaSep2[j];

      for (var f = 0; f < ListaArqPast2.length; f++) {
        if (
          ListaArqPast2[f].name.match(new RegExp("^" + fileName2 + "$", "i"))
        ) {
          verifyListCsv2.push(ListaArqPast2[f]);
          break;
        }
      }
    }

    // percorrer cada arquivo (parte copiada)
    for (var k = 0; k < verifyListCsv2.length; k++) {
      var source_file2 = verifyListCsv2[k]; //Pega posição K, e ordena

      if (source_file2 instanceof File && source_file2.name.match(/\.indd$/i)) {
        try {
          app.open(source_file2);
          var source_doc2 = app.documents.item(source_file2.name);
          var sourcePages2 = source_doc2.pages.item(0);

          // Duplica o arquivo original (Com os erros) e ele move para
          // o arquivo de destino.
          sourcePages2.duplicate(
            LocationOptions.AFTER,
            source_doc2.pages.item(0)
          );
          sourcePages2.move(LocationOptions.AFTER, documento2.pages.item(-1));

          // Adicionar uma linha em branco separadora ao final do texto na última caixa de texto da página atual
          var lastTextFrame2 =
            documento2.pages.item(-1).textFrames[
              documento2.pages.item(-1).textFrames.length - 1
            ];
          if (lastTextFrame2.parentStory) {
            lastTextFrame2.parentStory.insertionPoints[-1].contents = "\r"; // Adiciona um caractere de quebra de linha
          }

          // Fecha o arquivo que foi aberto sem salvar (para evitar problemas de memória)
          app.activeDocument.close(SaveOptions.NO);
        } catch (e) {
          //alert("Erro ao abrir ou manipular o arquivo: " + source_file2.name + "\n" + e.message);
        }
      } else {
        //alert("Arquivo inválido ou não encontrado: " + source_file2.name);
      }
    }

    // Encadeamento de caixas de texto
    var destinationPages = documento2.pages;
    for (var p = 1; p < destinationPages.length; p++) {
      var previousPage = destinationPages[p - 1];
      var currentPage = destinationPages[p];
      var lastTextFrame =
        previousPage.textFrames[previousPage.textFrames.length - 1];
      var firstTextFrame = currentPage.textFrames[0];
      lastTextFrame.nextTextFrame = firstTextFrame;
    }

    // Verifique se o aplicativo InDesign está aberto
    if (app && app.name == "Adobe InDesign") {
      // Verifique se há um documento aberto
      if (app.documents.length > 0) {
        var doc2 = app.activeDocument;

        // Verifique se há mais de uma página no documento
        if (doc2.pages.length > 1) {
          // Loop através de todas as páginas, exceto a primeira
          for (var o = 1; o < doc2.pages.length; o++) {
            var currentPage2 = doc2.pages[o];

            // Remova a seção da página
            if (currentPage2.appliedSection) {
              currentPage2.appliedSection.remove();
            }

            var masters = documento2.masterSpreads;

            for (i = masters.length - 1; i > 1; i--) {
              masters[i].remove();
            }
          }
        }
      }
    }

    exit();
  }

  function laco2(listaLinhas2, caminhoIndd2) {
    if (app.documents.length > 0) {
      arquivoIndt2 = app.activeDocument;
    } else {
      alert("Nenhum documento aberto!");
      exit();
    }

    // asubtração do tamanho da lista se deve a linha em braco
    for (var i = 0; i < listaLinhas2.length - 1; i++) {
      //Etapa 5: Manipular o arquivo InDesign
      manipularDocumentoIndesign2(
        arquivoIndt2,
        listaLinhas2[i],
        caminhoIndd2
      );
    }
    // Alerta de fim de execução
    alert("Bloco montado");
    exit();
  }

  var listaLinhas2 = lerArqCSV(); //Recebe tupla [(a,b,c),(d,e,f),(g,h,i,j)]

  laco2(listaLinhas2, caminhoIndd2);
}
