// Versão 3.0 - 17.06.24
//----------------------------------
// Notas da versão:
// Inserção do seletor de preset de pdf.
//----------------------------------
//VARIÁVEIS GLOBAIS
#target indesign;
#targetengine "session"
//----------------------------------
var TipoCaderno = 0;
var precapa = "";
var poscapa = "";
var codcapa = 0;
var precabecalho = "";
var poscabecalho = "";
var codcabecalho = 0;
var setsection = false;
var continuar = true;
//Variável Global que vai definir o número inicial do código do Caderno
var NumeroCadernoInicial = 1;
var padraoEncontrado = false;
var objChkBxArray = new Array();
var estadoInicial = true;
//número de casas_decimais decimais para o projeto
var casas_decimais = 0;
//tralhas
var tralhas = "";

var pdfPresets = app.pdfExportPresets.everyItem().name;
pdfPresets.unshift("- Select Preset -");

var selecionaCSV = File.openDialog(
  "Selecione um arquivo CSV",
  "selecionaCSV CSV:*.csv",
  false
);
var myFolder = Folder.selectDialog("Selecione a pasta dos blocos:");

var arquivos = myFolder.getFiles();
var tralhas_substituidas = 0;
var total_tralhas = 0;

for (var k = 0; k < arquivos.length; k++) {
  arquivo = arquivos[k];

  if (arquivo instanceof File && arquivo.name.match(/\.indd$/i)) {
    app.open(arquivo);

    var doc = app.activeDocument;
    var cabStyle = doc.paragraphStyles.item("cabecalho-rodape");

    if (cabStyle === null) {
      doc.paragraphStyles.add({
        name: "cabecalho-rodape",
        appliedFont: "Arial",
        fontStyle: "Regular",
        pointSize: 8,
        leading: 8,
        justification: Justification.CENTER_ALIGN
      });
    }

    checkTralhas();
    countTralhas();

    function checkTralhas() {
      app.findGrepPreferences = app.changeGrepPreferences = null;

      var tralhas_buscadas = "^\\d\\d\\)";

      app.findGrepPreferences.findWhat = tralhas_buscadas;

      var tralhas_encontradas = doc.findGrep();

      for (var j = 0; j < tralhas_encontradas.length; j++) {
        var tralha = tralhas_encontradas[j];

        tralha.select();

        app.changeGrepPreferences.changeTo = "##)";

        doc.changeGrep();
        tralhas_substituidas = tralhas_substituidas + 1;
      }
    }

    function countTralhas() {
      app.findGrepPreferences = app.changeGrepPreferences = null;

      var tralhas_buscadas = "^##\\)";

      app.findGrepPreferences.findWhat = tralhas_buscadas;

      var tralhas_encontradas = doc.findGrep();

      for (var j = 0; j < tralhas_encontradas.length; j++) {
        
        total_tralhas = total_tralhas + 1;
      }
    }

    doc.save();
    doc.close();
  }
}

var ArrayPrincipal = carregaCSV(selecionaCSV);

// Criando Formulário
// A Janela principal posuirá 4 painéis:
// 1. Informações sobre o caderno
// 2. Informações sobre a capa
// 3. Informações sobre o cabeçalho
// 4. Informações sobre o preset de pdf
// 5. Informações sobre tralhas e itens

// Dimensões padrão
var label_size = [0, 0, 400, 20];
var input_size = [0, 0, 160, 20];
var total_size = [0, 0, 560, 20];

//// Janela principal
var janelaConfig = new Window(
  "palette",
  "Configuração da montagem dos cadernos"
);

//// Painel 1 - CONFIGURAÇÕES DO CADERNO
var painelCadernos = janelaConfig.add(
  "panel",
  undefined,
  "CONFIGURAÇÕES DO CADERNO"
);
painelCadernos.spacing = 10;
painelCadernos.margins = 30;

var grupoTipoCaderno = painelCadernos.add("group");
var textoTipoCaderno = grupoTipoCaderno.add(
  "statictext",
  label_size,
  "Informe o tipo de caderno:"
);
var selectTipoCaderno = grupoTipoCaderno.add("dropdownlist", input_size, [
  "Comum",
  "Múltiplo de 4",
  "Múltiplo de 8"
]);

var grupoNumeracao = painelCadernos.add("group");
var statictextNumeracao = grupoNumeracao.add(
  "statictext",
  label_size,
  "Numeração Inicial do Caderno: (Ex.: X= 9, C1209)"
);
var editTextNumeracao = grupoNumeracao.add("edittext", input_size, "1");

var grupoDigitos = painelCadernos.add("group");
var statictextDigitos = grupoDigitos.add(
  "statictext",
  label_size,
  "Número de Dígitos do Código (capa e cabeçalho): (Ex.: X= 5, C1209)"
);
var editTextDigitos = grupoDigitos.add("edittext", input_size, "5");

//// Painel 2 - CONFIGURAÇÕES DA CAPA
var painelCapa = janelaConfig.add("panel", undefined, "CONFIGURAÇÕES DA CAPA");
painelCapa.spacing = 10;
painelCapa.margins = 30;

var grupoPrefixo = painelCapa.add("group");
var staticTextPrefixo = grupoPrefixo.add(
  "statictext",
  label_size,
  "Prefixo do Código: (Ex.: X= P, P0901)"
);
var editTextPrefixo = grupoPrefixo.add("edittext", input_size);

var grupoPosfixo = painelCapa.add("group");
var statictextPosfixo = grupoPosfixo.add(
  "statictext",
  label_size,
  "Posfixo do Código: (Ex.: X=_APLIC, M0901_APLIC)"
);
var editTextPosfixo = grupoPosfixo.add("edittext", input_size);

var grupoCodificacao = painelCapa.add("group");
var statictextCodificacao = grupoCodificacao.add(
  "statictext",
  label_size,
  "Codificação: (Ex.: 0910)"
);
var selectCodificacao = grupoCodificacao.add("dropdownlist", input_size, [
  "SIM",
  "NÃO"
]);
selectCodificacao.selection = 0;

//// Painel 3 - CONFIGURAÇÕES DE CABEÇALHO
var painelCabecalho = janelaConfig.add(
  "panel",
  undefined,
  "CONFIGURAÇÕES DE CABEÇALHO"
);
painelCabecalho.spacing = 10;
painelCabecalho.margins = 30;

var grupoPrefixoCab = painelCabecalho.add("group");
var statictextPrefixoCab = grupoPrefixoCab.add(
  "statictext",
  label_size,
  "Prefixo do Código: (Ex.: X= P, P0901)"
);
var editTextPrefixoCab = grupoPrefixoCab.add("edittext", input_size);

var grupoPosfixoCab = painelCabecalho.add("group");
var statictextPosfixoCab = grupoPosfixoCab.add(
  "statictext",
  label_size,
  "Posfixo do Código: (Ex.: X=_APLIC, M0901_APLIC)"
);
var editTextPosfixoCab = grupoPosfixoCab.add("edittext", input_size);

var grupoCodificacaoCab = painelCabecalho.add("group");
var statictextCodificacaoCab = grupoCodificacaoCab.add(
  "statictext",
  label_size,
  "Codificação: (Ex.: 0910)"
);
var selectCodificacaoCab = grupoCodificacaoCab.add("dropdownlist", input_size, [
  "SIM",
  "NÃO"
]);
selectCodificacaoCab.selection = 0;

//// Painel 4 - CONFIGURAÇÕES DE PRESET DE PDF
var painelPDF = janelaConfig.add(
  "panel",
  undefined,
  "CONFIGURAÇÕES DO PDF PARA IMPRESSÃO"
);
painelPDF.spacing = 10;
painelPDF.margins = 30;

var grupoPDF = painelPDF.add("group");
var statictextPDF = grupoPDF.add(
  "statictext",
  label_size,
  "Selecione o preset de PDF"
);
var selectPDF = grupoPDF.add("dropdownlist", input_size, undefined, {
  items: pdfPresets
});

//// Painel 5 - INFORMAÇÕES DE TRALHAS E ITENS
var painelTralhas = janelaConfig.add(
  "panel",
  undefined,
  "INFORMAÇÕES SOBRE TRALHAS E ITENS"
);
painelTralhas.spacing = 10;
painelTralhas.margins = 30;

var grupoTralhas = painelTralhas.add("group");
var statictextTralhas = grupoTralhas.add(
  "statictext",
  total_size,
  "Substituímos '\d\d)' por '##' em " + tralhas_substituidas + " ocorrências"
);
var grupoItens = painelTralhas.add("group");
var statictextItens = grupoItens.add(
  "statictext",
  total_size,
  "Com base nas tralhas encontradas, identificamos um total de " + total_tralhas + " itens"
);

var ok = janelaConfig.add("button", undefined, "OK");

janelaConfig.show();

ok.onClick = function () {
  pdfPresetname = String(selectPDF.selection);
  if (pdfPresetname == "null") {
    alert("Você precisa selecionar um preset de PDF para continuar");
  } else {
    janelaConfig.close();
    TipoCaderno = selectTipoCaderno.selection;
    precapa = editTextPrefixo.text;
    poscapa = editTextPosfixo.text;
    codcapa = selectCodificacao.selection;
    precabecalho = editTextPrefixoCab.text;
    poscabecalho = editTextPosfixoCab.text;
    codcabecalho = selectCodificacaoCab.selection;
    casas_decimais = parseInt(editTextDigitos.text);
    tralhas = "";
    for (var p_digitos = 1; casas_decimais >= p_digitos; p_digitos++) {
      tralhas = tralhas + "#";
    }
    NumeroCadernoInicial = editTextNumeracao.text;
    pdfPreset = app.pdfExportPresets.item(String(selectPDF.selection));

    montarCaderno(ArrayPrincipal, myFolder);
  }
};

function carregaCSV() {
  selecionaCSV.open("r");
  var myarray = [];

  if (!selecionaCSV) {
    exit();
  }

  var linhas = selecionaCSV.read().split("\n");

  for (var i = 0; i < linhas.length; i++) {
    myarray.push([]);
    linha = linhas[i];

    var celulas = linha.split(";");

    for (var j = 0; j < celulas.length; j++) {
      celula = celulas[j];
      myarray[i].push(celula);
    }
  }

  selecionaCSV.close();
  return myarray;
}

function montarCaderno(ArrayPrincipal, myFolder) {
  var bookNew;
  var documentoNew;
  var documentoOld;
  var contCaderno = 0;
  var contBloco = 1;
  var contagem = 1;
  var p_inicio = 0;
  var p_fim = 0;

  for (contCaderno; contCaderno < ArrayPrincipal.length; contCaderno++) {
    continuar = false;
    var totalPaginas = 0;
    contagem = 1;
    var contBloco = 1;
    setsection = false;
    var pasta = new Folder(
      myFolder + "/Caderno " + ArrayPrincipal[contCaderno][0]
    );
    if (pasta.exists == false) pasta.create();

    bookNew = app.books.add(
      File(pasta + "/" + ArrayPrincipal[contCaderno][0] + ".indb"),
      false
    );
    //Definir o intervalo de númeração e cabeçalho da página
    for (
      var p_count = 3;
      ArrayPrincipal[contCaderno].length > p_count;
      p_count++
    ) {
      if (
        ArrayPrincipal[contCaderno][1] == ArrayPrincipal[contCaderno][p_count]
      ) {
        p_inicio = p_count;
      } else if (
        ArrayPrincipal[contCaderno][2] == ArrayPrincipal[contCaderno][p_count]
      ) {
        p_fim = p_count;
      }
    }

    for (
      contBloco;
      contBloco < ArrayPrincipal[contCaderno].length;
      contBloco++
    ) {
      //coletando informações para o progress bar
      var blocosProgress = ArrayPrincipal[contCaderno].length + 1;
      var tamanhoProgress = 500;

      if (contBloco == 1) {
        var myIncrement = tamanhoProgress / blocosProgress;

        function myCreateProgressPanel(
          myMaximumValue,
          myProgressBarWidth,
          nomeCad
        ) {
          myProgressPanel = new Window(
            "window",
            "Montando " + nomeCad + ".pdf ... - (Produz Caderno - InDesign)"
          );
          with (myProgressPanel) {
            myProgressPanel.myProgressBar = add(
              "progressbar",
              [12, 12, myProgressBarWidth, 24],
              0,
              myMaximumValue
            );
          }
        }

        myCreateProgressPanel(
          tamanhoProgress,
          tamanhoProgress,
          ArrayPrincipal[contCaderno][0]
        );
        myProgressPanel.myProgressBar.value = 0;
        myProgressPanel.show();
      }

      if (contBloco > 2) {
        documentoOld = app.open(
          File(
            myFolder + "/" + ArrayPrincipal[contCaderno][contBloco] + ".indd"
          ),
          false
        );

        if (contBloco >= p_inicio && contBloco <= p_fim) {
          contagem = mySnippet(
            documentoOld,
            contagem,
            parseInt(NumeroCadernoInicial) + contCaderno,
            true
          );
        } else {
          contagem = mySnippet(
            documentoOld,
            contagem,
            parseInt(NumeroCadernoInicial) + contCaderno,
            false
          );
        }

        totalPaginas = totalPaginas + documentoOld.pages.count();

        if (
          contBloco + 1 == ArrayPrincipal[contCaderno].length &&
          TipoCaderno == 1
        ) {
          if (totalPaginas % 4 != 0) {
            for (
              var mm = 1;
              mm <= ExibirNumDivisivel(totalPaginas, 4) - totalPaginas;
              mm++
            ) {
              documentoOld.pages.add(
                LocationOptions.AT_BEGINNING,
                documentoOld.masterSpreads.itemByName("A-Master"),
                documentoOld.pages.item(0).properties
              );
            }
          }
        } else if (
          contBloco + 1 == ArrayPrincipal[contCaderno].length &&
          TipoCaderno == 2
        ) {
          if (totalPaginas % 8 != 0) {
            for (
              var mm = 1;
              mm <= ExibirNumDivisivel(totalPaginas, 8) - totalPaginas;
              mm++
            ) {
              documentoOld.pages.add(
                LocationOptions.AT_BEGINNING,
                documentoOld.masterSpreads.itemByName("A-Master"),
                documentoOld.pages.item(0).properties
              );
            }
          }
        }

        documentoNew = documentoOld.save(
          File(pasta + "/" + ArrayPrincipal[contCaderno][contBloco] + ".indd")
        );

        bookNew.bookContents.add(
          File(pasta + "/" + ArrayPrincipal[contCaderno][contBloco] + ".indd")
        );

        fecharTodos();

        myProgressPanel.myProgressBar.value += myIncrement;
      }
    }

    myProgressPanel.myProgressBar.value += myIncrement;
    bookNew.updateAllNumbers();
    ExportarPDF(myFolder, bookNew, ArrayPrincipal[contCaderno][0]);
    myProgressPanel.myProgressBar.value += myIncrement;
    bookNew.save();
    fecharTodosBooks();
    myProgressPanel.hide();
  }
}

function mySnippet(documentoPrincipal, contagem, caderno, p_numeracao) {
  var myDocument = documentoPrincipal;

  var paginas = myDocument.pages.length;

  var j = 0,
    i = contagem;

  var inddcaderno = caderno;

  for (j; j < paginas; j++) {
    inddcaderno = caderno;
    var myPage = myDocument.pages.item(j);
    var frames = myPage.textFrames.length;

    //	var numerado=false;
    for (var gh = 0; frames > gh; gh++) {
      var myFrame = myPage.textFrames.item(gh);

      //var Chars = myFrame.characters.length-1;
      try {
        //PORQUE var paragrafos = myFrame.paragraphs.length-1; ?
        var paragrafos = myFrame.paragraphs.count();

        for (var k = 0; k < paragrafos; k++) {
          if (myFrame.paragraphs.item(k).tables.count() > 0) {
            var contadortabelaver1 = myFrame.paragraphs.item(k).tables.count();

            for (
              var contadorver1 = 0;
              contadortabelaver1 > contadorver1;
              contadorver1++
            ) {
              var tabelaver1 = myFrame.paragraphs
                .item(k)
                .tables.item(contadorver1);

              var contadorcolunasver1 = tabelaver1.columns.count();

              for (
                var contadorver2 = 0;
                contadorcolunasver1 > contadorver2;
                contadorver2++
              ) {
                var colunaver1 = tabelaver1.columns.item(contadorver2);

                var contadorcelulasver1 = colunaver1.cells.count();

                for (
                  var contadorver3 = 0;
                  contadorcelulasver1 > contadorver3;
                  contadorver3++
                ) {
                  var celulaver1 = colunaver1.cells.item(contadorver3);

                  var totalParagravover1 = celulaver1.paragraphs.count();

                  for (
                    var contadorver4 = 0;
                    contadorver4 < totalParagravover1;
                    contadorver4++
                  ) {
                    var palavraver1 =
                      celulaver1.paragraphs.item(contadorver4).contents;
                    var resultadover1 = palavraver1.search("##");
                    var codcapaVER1 = palavraver1.search(tralhas);
                    if (resultadover1 != -1 || codcapaVER1 != -1) {
                      var charsver1 =
                        celulaver1.paragraphs.item(contadorver4).characters
                          .length;
                      for (
                        var contadorver5 = 0;
                        contadorver5 < charsver1 - 1;
                        contadorver5++
                      ) {
                        var frasever1 =
                          celulaver1.paragraphs.item(contadorver4).contents;

                        if (contadorver5 != frasever1.length - 1) {
                          var caractervar1 =
                            frasever1[contadorver5] +
                            frasever1[contadorver5 + 1];
                          var wordCodCapaVER1 = frasever1[contadorver5];
                          if (casas_decimais > 0) {
                            for (
                              var p_casas_decimais = 1;
                              casas_decimais > p_casas_decimais;
                              p_casas_decimais++
                            ) {
                              wordCodCapaVER1 =
                                wordCodCapaVER1 +
                                frasever1[contadorver5 + p_casas_decimais];
                            }
                          }
                        }

                        if (
                          caractervar1 == "##" &&
                          wordCodCapaVER1 != tralhas
                        ) {
                          caractervar1 = "";
                          //    if (i<10)  myFrame.paragraphs.item(k).contents = myWord.replace("##", '0'+i+'');
                          // else myFrame.paragraphs.item(k).contents = myWord.replace("##", i+'');
                          padraoEncontrado = true;
                          numerado = true;
                          if (i < 10) {
                            celulaver1.paragraphs
                              .item(contadorver4)
                              .characters.item(contadorver5).contents = 0 + "";
                            celulaver1.paragraphs
                              .item(contadorver4)
                              .characters.item(contadorver5 + 1).contents =
                              i + "";
                          } else {
                            var decimal = i + "";
                            celulaver1.paragraphs
                              .item(contadorver4)
                              .characters.item(contadorver5).contents =
                              decimal[0];
                            celulaver1.paragraphs
                              .item(contadorver4)
                              .characters.item(contadorver5 + 1).contents =
                              decimal[1];
                          }
                          i++;
                        } else if (wordCodCapaVER1 == tralhas) {
                          var p_cadernoTexto = inddcaderno + "";
                          for (
                            var p_contindd = 1;
                            p_contindd <=
                            casas_decimais - 3 - p_cadernoTexto.length;
                            p_contindd++
                          )
                            inddcaderno = "0" + inddcaderno;
                          //  padraoEncontrado = true;
                          /*     if (inddcaderno < 10)
                                                         {
                                                             inddcaderno='0'+caderno;
                                                         }*/
                          if (codcapa == 0)
                            var formcod = precapa + inddcaderno + poscapa;
                          else var formcod = precapa + poscapa;
                          for (nn = 0; nn <= formcod.length - 1; nn++) {
                            celulaver1.paragraphs
                              .item(totalParagravover1)
                              .characters.item(frasever1 + nn).contents =
                              formcod[nn];
                          }
                        }
                      }
                    }
                  }
                }
              }
            }
          } else {
            var myWord = myFrame.paragraphs.item(k).contents;
            var resultado = myWord.search("##");
            var codcapa = myWord.search(tralhas);
            if (resultado != -1 || codcapa != -1) {
              var chars = myFrame.paragraphs.item(k).characters.length;
              for (var gg = 0; gg < chars - 1; gg++) {
                myWord = myFrame.paragraphs.item(k).contents;
                if (k != myWord.length - 1) {
                  var wordChar = myWord[gg] + myWord[gg + 1];
                  var wordCodCapa = myWord[gg];
                  if (casas_decimais > 0) {
                    for (
                      var p_casas_decimais = 1;
                      casas_decimais > p_casas_decimais;
                      p_casas_decimais++
                    ) {
                      wordCodCapa = wordCodCapa + myWord[gg + p_casas_decimais];
                    }
                  }
                }

                if (wordChar == "##" && wordCodCapa != tralhas) {
                  wordChar = "";
                  //    if (i<10)  myFrame.paragraphs.item(k).contents = myWord.replace("##", '0'+i+'');
                  // else myFrame.paragraphs.item(k).contents = myWord.replace("##", i+'');
                  padraoEncontrado = true;
                  //   numerado=true;
                  if (i < 10) {
                    myFrame.paragraphs.item(k).characters.item(gg).contents =
                      0 + "";
                    myFrame.paragraphs
                      .item(k)
                      .characters.item(gg + 1).contents = i + "";
                  } else {
                    var decimal = i + "";
                    myFrame.paragraphs.item(k).characters.item(gg).contents =
                      decimal[0];
                    myFrame.paragraphs
                      .item(k)
                      .characters.item(gg + 1).contents = decimal[1];
                  }
                  i++;
                } else if (wordCodCapa == tralhas) {
                  padraoEncontrado = true;
                  var p_cadernoTexto = inddcaderno + "";
                  for (
                    var p_contindd = 1;
                    p_contindd <= casas_decimais - 3 - p_cadernoTexto.length;
                    p_contindd++
                  )
                    inddcaderno = "0" + inddcaderno;
                  /*    if (inddcaderno < 10)
                                        {
                                            inddcaderno='0'+caderno;
                                        }*/
                  if (codcapa == 0)
                    var formcod = precapa + inddcaderno + poscapa;
                  else var formcod = precapa + poscapa;
                  for (nn = 0; nn <= formcod.length - 1; nn++) {
                    myFrame.paragraphs
                      .item(k)
                      .characters.item(gg + nn).contents = formcod[nn];
                  }
                }
              }
            }
          }
        }
      } catch (myError) {
        //try
        //faz nada, continua
      }
      try {
        if (padraoEncontrado == false) {
          var tabelas = myFrame.tables.count();

          for (var k = 0; k < tabelas; k++) {
            var myTable = myFrame.tables.item(k);
            //var colunas = myTable.columns.length;
            var myColumns = myTable.columns.item(0);
            //var celulas = myColumns.cells.length;
            var myCell = myColumns.cells.item(0);

            // var charsCell = myCell.characters.length;
            var paragrafos = myCell.paragraphs.length - 1;
            //var procurado = myCell.paragraphs.item(0).contents;
            //myCell.paragraphs.item(0).contents = "Questão 01";
            //for (var l=0;l<paragrafos; l++) {
            var procurado = myCell.paragraphs.item(0).contents;
            var answer = procurado.search("##");
            var codcapa1 = procurado.search(tralhas);
            if (answer != -1 || codcapa1 != -1) {
              var chars = myCell.paragraphs.item(0).characters.length;
              for (var xlz = 0; xlz < chars; xlz++) {
                procurado = myCell.paragraphs.item(0).contents;
                if (xlz != chars - 1) {
                  var wordChar = procurado[xlz] + procurado[xlz + 1];
                  if (casas_decimais > 0) {
                    for (
                      var p_casas_decimais = 1;
                      casas_decimais > p_casas_decimais;
                      p_casas_decimais++
                    ) {
                      wordcodcapa = procurado[xlz + p_casas_decimais];
                    }
                  }
                }

                if (wordChar == "##" && wordcodcapa != tralhas) {
                  //padraoEncontrado = true;
                  wordChar = "";
                  numerado = true;
                  if (i < 10) {
                    myCell.paragraphs.item(0).characters.item(xlz).contents =
                      0 + "";
                    myCell.paragraphs
                      .item(0)
                      .characters.item(xlz + 1).contents = i + "";
                    //myCell.paragraphs.item(0).contents = procurado.replace("altnum", '0'+i+'');
                  } else {
                    var decimal = i + "";
                    //   myCell.paragraphs.item(0).contents = procurado.replace("altnum", i+'');
                    myCell.paragraphs.item(0).characters.item(xlz).contents =
                      decimal[0];
                    myCell.paragraphs
                      .item(0)
                      .characters.item(xlz + 1).contents = decimal[1];
                  }
                  i++;
                } else if (wordcodcapa == tralhas) {
                  var p_cadernoTexto = inddcaderno + "";
                  for (
                    var p_contindd = 1;
                    p_contindd <= casas_decimais - p_cadernoTexto.length;
                    p_contindd++
                  )
                    inddcaderno = "0" + inddcaderno;
                  if (codcapa == 0)
                    var formcod = precapa + inddcaderno + poscapa;
                  else var formcod = precapa + poscapa;
                  myCell.paragraphs.item(0).characters.contents = formcod;
                }
              }
            }
          }
        }
      } catch (myError) {}

      if (p_numeracao == true) {
        var p_cadernoTexto = inddcaderno + "";
        for (
          var p_contindd = 1;
          p_contindd <= casas_decimais - 3 - p_cadernoTexto.length;
          p_contindd++
        )
          inddcaderno = "0" + inddcaderno;

        if (codcabecalho == 0)
          InserirQuadroTexto(myPage, precabecalho + inddcaderno + poscabecalho);
        else InserirQuadroTexto(myPage, precabecalho + poscabecalho);
      }
      if (p_numeracao == true) {
        InserirPagina(myPage, documentoPrincipal);
        if (setsection == false) {
          var objspread = myPage.parent;
          var objdocumento = objspread.parent;
          var newSection = objdocumento.sections.add({
            pageStart: objdocumento.pages[0],
            continueNumbering: false,
            pageNumberStart: 1
          });
          setsection = true;
        }
      }
    }
  }

  return i;
}

function fecharTodos() {
  for (myCounter = app.documents.length; myCounter > 0; myCounter--) {
    app.documents.item(myCounter - 1).close(SaveOptions.no);
  }
}

function fecharTodosBooks() {
  for (myCounter = app.books.length; myCounter > 0; myCounter--) {
    app.books.item(myCounter - 1).close(SaveOptions.no);
  }
}

function checarPreflight(myDoc) {
  // Assume there is a document.
  var myProfile = app.preflightProfiles.item(0);
  //Process the myDoc with the rule
  var myProcess = app.preflightProcesses.add(myDoc, myProfile);
  myProcess.waitForProcess();

  var myResults = myProcess.processResults;

  //<fragment>
  //If Errors were found length
  if (
    myResults.toUpperCase().search("NONE") == -1 &&
    myResults.toUpperCase().search("NO ERRORS") == -1
  ) {
    ShowDialog(myResults.toUpperCase());
    return 1;
  } else return 0;
  //</fragment>

  //Cleanup
  myProcess.remove();
}

function ShowDialog(mensagem) {
  //<fragment>
  var myDialog = app.dialogs.add({
    name: "Produz Caderno - InDesign",
    canCancel: false
  });
  //Add a dialog column.
  with (myDialog.dialogColumns.add()) {
    staticTexts.add({
      staticLabel: mensagem
    });
  }
  //Show the dialog box.
  var myResult = myDialog.show();

  return myDialog;
}

function ShowConfirm(mensagem) {
  //<fragment>
  var myDialog = app.dialogs.add({
    name: "Produz Caderno - InDesign",
    canCancel: true
  });
  //Add a dialog column.
  with (myDialog.dialogColumns.add()) {
    staticTexts.add({
      staticLabel: mensagem
    });
  }
  //Show the dialog box.
  var myResult = myDialog.show();

  return myResult;
}

function ExibirNumDivisivel(numero, div) {
  num = numero;
  while (num % div != 0) {
    num++;
  }

  return num;
}

function PreLoad() {
  if (app.documents.item(0) != null || app.books.item(0) != null) {
    dialogo = ShowConfirm(
      "O sistema encontrou documentos abertos durante o processo!" +
        " Clique 'OK' para o sistema fechar todos os documentos (sem salvar) automaticamente."
    );
    if (dialogo == true) {
      fecharTodos();
      fecharTodosBooks();
      return true;
    } else {
      return false;
    }
  } else return true;
}

function InserirQuadroTexto(myPage, codText) {
  var y1 = myPage.marginPreferences.top / 2.083;
  var x1 = myPage.marginPreferences.left;
  var y2 = y1 + 2.826;
  var x2 =
    myPage.textFrames.item(0).geometricBounds[3] -
    myPage.textFrames.item(0).geometricBounds[0] +
    x1;

  var quadroTexto = myPage.textFrames.add();
  // [(12.5-5.087)-1.413, (7.413+6.5)-1.413,((12.5+2.826)-5.087)-1.413, (180+7.413+6.5)-1.413]
  quadroTexto.geometricBounds = [y1, x1, y2, x2];
  quadroTexto.contents = codText;
  quadroTexto.parentStory.paragraphs.item(0).justification =
    Justification.CENTER_ALIGN;
  quadroTexto.textFramePreferences.verticalJustification =
    VerticalJustification.CENTER_ALIGN;
  quadroTexto.parentStory.paragraphs.item(0).pointSize = 8;
  quadroTexto.parentStory.paragraphs.item(0).appliedFont = "Arial";
  quadroTexto.paragraphs.everyItem().appliedParagraphStyle = "cabecalho-rodape";
}

function InserirPagina(objPagina, objdocumento) {
  var y1 = 1.01;
  var diferenca = 1.01;
  var altura = Math.round(objdocumento.documentPreferences.pageHeight);
  if (altura == 275) diferenca = 11.25;
  else diferenca = (objdocumento.documentPreferences.pageHeight * 11.25) / 275;
  y1 = objdocumento.documentPreferences.pageHeight - diferenca;
  var x1 = objPagina.marginPreferences.left;
  var y2 = y1 + 8;
  var x2 =
    objPagina.textFrames.item(1).geometricBounds[3] -
    objPagina.textFrames.item(1).geometricBounds[0] +
    x1;
  var quadroTexto = objPagina.textFrames.add();

  quadroTexto.geometricBounds = [y1, x1, y2, x2];
  quadroTexto.contents = SpecialCharacters.autoPageNumber;
  quadroTexto.parentStory.paragraphs.item(0).justification =
    Justification.CENTER_ALIGN;
  quadroTexto.textFramePreferences.verticalJustification =
    VerticalJustification.CENTER_ALIGN;
  quadroTexto.parentStory.paragraphs.item(0).pointSize = 8;
  quadroTexto.parentStory.paragraphs.item(0).appliedFont = "Arial";
  quadroTexto.paragraphs.everyItem().appliedParagraphStyle = "cabecalho-rodape";
}

function ExportarPDF(myPasta, documentoResultante, nomePdf) {
  var pasta = new Folder(myPasta + "/Cadernos");
  if (pasta.exists == false) pasta.create();

  documentoResultante.exportFile(
    ExportFormat.pdfType,
    File(pasta + "/" + nomePdf + ".pdf"),
    false,
    pdfPreset
  );
}
