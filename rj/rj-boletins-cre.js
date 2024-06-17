//Resetando as preferências de PDF do Indesign
app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;
app.pdfExportPreferences.viewPDF = false;

var doc = app.activeDocument;
var docPath = doc.filePath;
var docname = doc.name.split(".")[0];
var docGerando = docPath + "/" + docname + "_gerando.indd";
doc.save(new File(docGerando));

var page_itens = doc.pageItems;

// Caixas de diálogo para selelecionar pasta onde estão máscaras e pasta onde serão salvas os boletins
var pastaCartazes =
  Folder.selectDialog("Selecione a pasta onde deseja salvar os arquivos")
    .absoluteURI + "/";
var caminhoJSON = File.openDialog("Selecione um arquivo JSON");

caminhoJSON.open("r");

var infoObject = caminhoJSON.read();

// Trasnforma o obejto em objeto json interpretável pelo indeisgn
var newJson = eval("(" + infoObject + ")");

var jsonItens = newJson.itens;

// Pega todas as chaves do json para criar TextVariables no documento
var jsonKeys = newJson.itens[0];

var getJsonKeys = function (jsonObject) {
  var arrayWithKeys = [],
    jsonObject;
  for (key in jsonObject) {
    // Avoid returning these keys from the Associative Array that are stored in it for some reason
    if (key !== undefined && key !== "toJSONString" && key !== "parseJSON") {
      arrayWithKeys.push(key);
    }
  }
  return arrayWithKeys;
};

var keys = getJsonKeys(jsonKeys);

for (i = 0; i < jsonItens.length; i++) {
  var jsonItem = jsonItens[i];

  for (m = 0; m < keys.length; m++) {
    jsonKey = keys[m];

    if (
      jsonKey.match("ESCOLA") ||
      jsonKey.match("REGIONAL") ||
      jsonKey.match("ETAPA")
    ) {
      var page_item = doc.masterSpreads[0].pageItems.item(jsonKey);

      if (page_item.isValid) {
        page_item.contents = jsonItem[jsonKey];
      }
    } else {
      var page_item = doc.pageItems.item(jsonKey);

      if (page_item.isValid) {
        if (
          jsonKey.match("VL_PRT") ||
          jsonKey.match("VL_ACERTO") ||
          jsonKey.match("VL_N")
        ) {
          if (
            jsonItem[jsonKey] == "0" &&
            jsonItem[jsonKey] == "" &&
            jsonItem[jsonKey] == "-"
          ) {
            page_item.contents = jsonItem[jsonKey];
          } else {
            page_item.contents = jsonItem[jsonKey] + "%";
          }
        } else if (jsonKey.match("NU_N01")) {
          if (
            jsonItem[jsonKey] == "0" &&
            jsonItem[jsonKey] == "" &&
            jsonItem[jsonKey] == "-"
          ) {
            page_item.contents = "Abaixo do básico - Nenhum estudante";
          } else if (jsonItem[jsonKey] == "1") {
            page_item.contents =
              "Abaixo do básico - " + jsonItem[jsonKey] + " estudante";
          } else {
            page_item.contents =
              "Abaixo do básico - " + jsonItem[jsonKey] + " estudantes";
          }
        } else if (jsonKey.match("NU_N02")) {
          if (
            jsonItem[jsonKey] == "0" &&
            jsonItem[jsonKey] == "" &&
            jsonItem[jsonKey] == "-"
          ) {
            page_item.contents = "Básico - Nenhum estudante";
          } else if (jsonItem[jsonKey] == "1") {
            page_item.contents = "Básico - " + jsonItem[jsonKey] + " estudante";
          } else {
            page_item.contents =
              "Básico - " + jsonItem[jsonKey] + " estudantes";
          }
        } else if (jsonKey.match("NU_N03")) {
          if (
            jsonItem[jsonKey] == "0" &&
            jsonItem[jsonKey] == "" &&
            jsonItem[jsonKey] == "-"
          ) {
            page_item.contents = "Adequado - Nenhum estudante";
          } else if (jsonItem[jsonKey] == "1") {
            page_item.contents =
              "Adequado - " + jsonItem[jsonKey] + " estudante";
          } else {
            page_item.contents =
              "Adequado - " + jsonItem[jsonKey] + " estudantes";
          }
        } else if (jsonKey.match("NU_N04")) {
          if (
            jsonItem[jsonKey] == "0" &&
            jsonItem[jsonKey] == "" &&
            jsonItem[jsonKey] == "-"
          ) {
            page_item.contents = "Avançado - Nenhum estudante";
          } else if (jsonItem[jsonKey] == "1") {
            page_item.contents =
              "Avançado - " + jsonItem[jsonKey] + " estudante";
          } else {
            page_item.contents =
              "Avançado - " + jsonItem[jsonKey] + " estudantes";
          }
        } else {
          page_item.contents = jsonItem[jsonKey];
        }
      }
    }
  }

  var doc = app.activeDocument;

  app.findGrepPreferences = app.changeGrepPreferences = null;

  var find = "<table>(.+)</table>";

  app.findGrepPreferences.findWhat = find;

  var Founds = doc.findGrep();

  for (var j = 0; j < Founds.length; j++) {
    var curFound = Founds[j];

    curFound.select();

    app.changeGrepPreferences.changeTo = "$1";

    doc.changeGrep();
    app.selection[0].convertToTable("*", "|");
  }

  var myTable = app.activeDocument.stories
    .everyItem()
    .tables.everyItem()
    .getElements();

  myTable[0].appliedTableStyle = "matriz";
  myTable[0].cells[0].width = 100;

  myTable[0].cells[1].width = 480;
  myTable[0].cells[2].width = 65;
  myTable[0].cells[3].width = 65;
  myTable[0].cells[0].rowType = RowTypes.HEADER_ROW;

  myTable[0].cells[0].appliedCellStyle = "matriz-cab";
  myTable[0].cells[1].appliedCellStyle = "matriz-cab";
  myTable[0].cells[2].appliedCellStyle = "matriz-cab-center";
  myTable[0].cells[3].appliedCellStyle = "matriz-cab-center";

  myTable[0].cells[0].clearCellStyleOverrides(true);
  myTable[0].cells[1].clearCellStyleOverrides(true);
  myTable[0].cells[2].clearCellStyleOverrides(true);
  myTable[0].cells[3].clearCellStyleOverrides(true);

  var col_lp_dc = myTable[0].columns[0].cells;

  for (var j = 1; j < col_lp_dc.length; j++) {
    var cell = col_lp_dc[j];
    cell.appliedCellStyle = "matriz-descritores";
    cell.clearCellStyleOverrides(true);
  }

  var col_lp_txt = myTable[0].columns[1].cells;

  for (var j = 1; j < col_lp_txt.length; j++) {
    var cell = col_lp_txt[j];
    cell.appliedCellStyle = "matriz-corpo";
    cell.clearCellStyleOverrides(true);
  }

  var col_lp_reg = myTable[0].columns[2].cells;

  for (var j = 1; j < col_lp_reg.length; j++) {
    var cell = col_lp_reg[j];
    cell.appliedCellStyle = "matriz-acerto";
    cell.clearCellStyleOverrides(true);
  }

  var col_lp_mun = myTable[0].columns[3].cells;

  for (var j = 1; j < col_lp_mun.length; j++) {
    var cell = col_lp_mun[j];
    cell.appliedCellStyle = "matriz-acerto";
    cell.clearCellStyleOverrides(true);
  }

  myTable[1].appliedTableStyle = "matriz";
  myTable[1].cells[0].width = 100;
  myTable[1].cells[1].width = 480;
  myTable[1].cells[2].width = 65;
  myTable[1].cells[3].width = 65;
  myTable[1].cells[0].rowType = RowTypes.HEADER_ROW;

  myTable[1].cells[0].appliedCellStyle = "matriz-cab";
  myTable[1].cells[1].appliedCellStyle = "matriz-cab";
  myTable[1].cells[2].appliedCellStyle = "matriz-cab-center";
  myTable[1].cells[3].appliedCellStyle = "matriz-cab-center";

  myTable[1].cells[0].clearCellStyleOverrides(true);
  myTable[1].cells[1].clearCellStyleOverrides(true);
  myTable[1].cells[2].clearCellStyleOverrides(true);
  myTable[1].cells[3].clearCellStyleOverrides(true);

  var col_mt_dc = myTable[1].columns[0].cells;

  for (var j = 1; j < col_mt_dc.length; j++) {
    var cell = col_mt_dc[j];
    cell.appliedCellStyle = "matriz-descritores";
    cell.clearCellStyleOverrides(true);
  }

  var col_mt_txt = myTable[1].columns[1].cells;

  for (var j = 1; j < col_mt_txt.length; j++) {
    var cell = col_mt_txt[j];
    cell.appliedCellStyle = "matriz-corpo";
    cell.clearCellStyleOverrides(true);
  }

  var col_mt_reg = myTable[1].columns[2].cells;

  for (var j = 1; j < col_mt_reg.length; j++) {
    var cell = col_mt_reg[j];
    cell.appliedCellStyle = "matriz-acerto";
    cell.clearCellStyleOverrides(true);
  }

  var col_mt_mun = myTable[1].columns[3].cells;

  for (var j = 1; j < col_mt_mun.length; j++) {
    var cell = col_mt_mun[j];
    cell.appliedCellStyle = "matriz-acerto";
    cell.clearCellStyleOverrides(true);
  }

  for (m = 0; m < keys.length; m++) {
    jsonKey = keys[m];

    app.findTextPreferences = app.changeTextPreferences = null;
    var buscar = (app.findTextPreferences.findWhat = "<<" + jsonKey + ">>");

    var allFounds = doc.findText();

    for (var n = 0; n < allFounds.length; n++) {
      var curFound = allFounds[n];

      curFound.select();

      app.changeTextPreferences.changeTo = jsonItem[jsonKey].toString();
      doc.changeText();
    }
  }

  app.findTextPreferences = app.changeTextPreferences = null;
  var buscar = (app.findTextPreferences.findWhat = "\\r");

  var allFounds = doc.findText();

  for (var n = 0; n < allFounds.length; n++) {
    var curFound = allFounds[n];

    curFound.select();

    app.changeTextPreferences.changeTo = "^n";
    doc.changeText();
  }

  var esconde_escrita = doc.pageItems.item('esconde_escrita');
  var master_etapa = doc.masterSpreads[0].pageItems.item('ETAPA');

  if(master_etapa.contents == 'ENSINO FUNDAMENTAL DE 9 ANOS - 7º ANO' || master_etapa.contents == 'ENSINO FUNDAMENTAL DE 9 ANOS - 8º ANO' || master_etapa.contents == 'ENSINO FUNDAMENTAL DE 9 ANOS - 9º ANO'){
    esconde_escrita.visible = true;
  } else{
    esconde_escrita.visible = false;
  }

  redimensiona();
  coloreHabilidades();

  // salva o pdf na pasta selecionada anteriormente.
  var tipo_pdf = app.pdfExportPresets.item("[Smallest File Size]");
  var pdfname =
    pastaCartazes +
    "/" +
    jsonItem.REGIONAL + "_" + jsonItem.ETAPA_MINI +
    "_ADR1_2024" +
    ".pdf";
  doc.exportFile(ExportFormat.PDF_TYPE, new File(pdfname), false, tipo_pdf);
}

function redimensiona() {
  var pontos_graficos = [
    ["barra", "VL_N01_LP_REG_1B", "bh_lp_n1_1b", 644, 100],
    ["barra", "VL_N02_LP_REG_1B", "bh_lp_n2_1b", 644, 100],
    ["barra", "VL_N03_LP_REG_1B", "bh_lp_n3_1b", 644, 100],
    /* ["barra", "VL_N04_LP_REG_1B", "bh_lp_n4_1b", 644, 100],
     */
    ["barra", "VL_N01_MT_REG_1B", "bh_mt_n1_1b", 644, 100],
    ["barra", "VL_N02_MT_REG_1B", "bh_mt_n2_1b", 644, 100],
    ["barra", "VL_N03_MT_REG_1B", "bh_mt_n3_1b", 644, 100],
    /* ["barra", "VL_N04_MT_REG_1B", "bh_mt_n4_1b", 644, 100], */

    ["barra", "VL_N01_ESCRITA_REG_1B", "bh_escrita_n1_1b", 644, 100],
    ["barra", "VL_N02_ESCRITA_REG_1B", "bh_escrita_n2_1b", 644, 100],
    ["barra", "VL_N03_ESCRITA_REG_1B", "bh_escrita_n3_1b", 644, 100],
    /* ["barra", "VL_N04_ESCRITA_REG_1B", "bh_escrita_n4_1b", 200, 100], */
    
    /* ["barra", "VL_N01_LP_REG_2B", "bh_lp_n1_2b", 644, 100],
    ["barra", "VL_N02_LP_REG_2B", "bh_lp_n2_2b", 644, 100],
    ["barra", "VL_N03_LP_REG_2B", "bh_lp_n3_2b", 644, 100],
    ["barra", "VL_N04_LP_REG_2B", "bh_lp_n4_2b", 644, 100], */
    
    /* ["barra", "VL_N01_MT_REG_2B", "bh_mt_n1_2b", 644, 100],
    ["barra", "VL_N02_MT_REG_2B", "bh_mt_n2_2b", 644, 100],
    ["barra", "VL_N03_MT_REG_2B", "bh_mt_n3_2b", 644, 100],
    ["barra", "VL_N04_MT_REG_2B", "bh_mt_n4_2b", 644, 100], */
   
    /* ["barra", "VL_N01_LP_REG_3B", "bh_lp_n1_3b", 644, 100],
    ["barra", "VL_N02_LP_REG_3B", "bh_lp_n2_3b", 644, 100],
    ["barra", "VL_N03_LP_REG_3B", "bh_lp_n3_3b", 644, 100],
    ["barra", "VL_N04_LP_REG_3B", "bh_lp_n4_3b", 644, 100], */
    
    /* ["barra", "VL_N01_MT_REG_3B", "bh_mt_n1_3b", 644, 100],
    ["barra", "VL_N02_MT_REG_3B", "bh_mt_n2_3b", 644, 100],
    ["barra", "VL_N03_MT_REG_3B", "bh_mt_n3_3b", 644, 100],
    ["barra", "VL_N04_MT_REG_3B", "bh_mt_n4_3b", 644, 100], */

    /*  ["barra", "VL_N01_ESCRITA_REG_3B", "bh_escrita_n1_3b", 200, 100],
    ["barra", "VL_N02_ESCRITA_REG_3B", "bh_escrita_n2_3b", 200, 100],
    ["barra", "VL_N03_ESCRITA_REG_3B", "bh_escrita_n3_3b", 200, 100],
    ["barra", "VL_N04_ESCRITA_REG_3B", "bh_escrita_n4_3b", 200, 100], */
    
    /* ["barra", "VL_N01_LP_REG_4B", "bh_lp_n1_4b", 644, 100],
    ["barra", "VL_N02_LP_REG_4B", "bh_lp_n2_4b", 644, 100],
    ["barra", "VL_N03_LP_REG_4B", "bh_lp_n3_4b", 644, 100],
    ["barra", "VL_N04_LP_REG_4B", "bh_lp_n4_4b", 644, 100], */
    
    /* ["barra", "VL_N01_MT_REG_4B", "bh_mt_n1_4b", 644, 100],
    ["barra", "VL_N02_MT_REG_4B", "bh_mt_n2_4b", 644, 100],
    ["barra", "VL_N03_MT_REG_4B", "bh_mt_n3_4b", 644, 100],
    ["barra", "VL_N04_MT_REG_4B", "bh_mt_n4_4b", 644, 100], */
    
    /// PROVA RIO CEBRASPE

    /* ["barra", "VL_N01_CEBRASPE_LP_REG_2021", "bh_lp_n1_2021", 200, 100],
    ["barra", "VL_N02_CEBRASPE_LP_REG_2021", "bh_lp_n2_2021", 200, 100],
    ["barra", "VL_N03_CEBRASPE_LP_REG_2021", "bh_lp_n3_2021", 200, 100],
    ["barra", "VL_N04_CEBRASPE_LP_REG_2021", "bh_lp_n4_2021", 200, 100],

    ["barra", "VL_N01_CEBRASPE_LP_REG_2022", "bh_lp_n1_2022", 200, 100],
    ["barra", "VL_N02_CEBRASPE_LP_REG_2022", "bh_lp_n2_2022", 200, 100],
    ["barra", "VL_N03_CEBRASPE_LP_REG_2022", "bh_lp_n3_2022", 200, 100],
    ["barra", "VL_N04_CEBRASPE_LP_REG_2022", "bh_lp_n4_2022", 200, 100],

    ["barra", "VL_N01_CEBRASPE_LP_REG_2023", "bh_lp_n1_2023", 200, 100],
    ["barra", "VL_N02_CEBRASPE_LP_REG_2023", "bh_lp_n2_2023", 200, 100],
    ["barra", "VL_N03_CEBRASPE_LP_REG_2023", "bh_lp_n3_2023", 200, 100],
    ["barra", "VL_N04_CEBRASPE_LP_REG_2023", "bh_lp_n4_2023", 200, 100],

    ["barra", "VL_N01_CEBRASPE_MT_REG_2021", "bh_mt_n1_2021", 200, 100],
    ["barra", "VL_N02_CEBRASPE_MT_REG_2021", "bh_mt_n2_2021", 200, 100],
    ["barra", "VL_N03_CEBRASPE_MT_REG_2021", "bh_mt_n3_2021", 200, 100],
    ["barra", "VL_N04_CEBRASPE_MT_REG_2021", "bh_mt_n4_2021", 200, 100],

    ["barra", "VL_N01_CEBRASPE_MT_REG_2022", "bh_mt_n1_2022", 200, 100],
    ["barra", "VL_N02_CEBRASPE_MT_REG_2022", "bh_mt_n2_2022", 200, 100],
    ["barra", "VL_N03_CEBRASPE_MT_REG_2022", "bh_mt_n3_2022", 200, 100],
    ["barra", "VL_N04_CEBRASPE_MT_REG_2022", "bh_mt_n4_2022", 200, 100],

    ["barra", "VL_N01_CEBRASPE_MT_REG_2023", "bh_mt_n1_2023", 200, 100],
    ["barra", "VL_N02_CEBRASPE_MT_REG_2023", "bh_mt_n2_2023", 200, 100],
    ["barra", "VL_N03_CEBRASPE_MT_REG_2023", "bh_mt_n3_2023", 200, 100],
    ["barra", "VL_N04_CEBRASPE_MT_REG_2023", "bh_mt_n4_2023", 200, 100],

    ["barra", "VL_N01_CEBRASPE_ESCRITA_REG_2021", "bh_escrita_n1_2021", 200, 100],
    ["barra", "VL_N02_CEBRASPE_ESCRITA_REG_2021", "bh_escrita_n2_2021", 200, 100],
    ["barra", "VL_N03_CEBRASPE_ESCRITA_REG_2021", "bh_escrita_n3_2021", 200, 100],
    ["barra", "VL_N04_CEBRASPE_ESCRITA_REG_2021", "bh_escrita_n4_2021", 200, 100],

    ["barra", "VL_N01_CEBRASPE_ESCRITA_REG_2022", "bh_escrita_n1_2022", 200, 100],
    ["barra", "VL_N02_CEBRASPE_ESCRITA_REG_2022", "bh_escrita_n2_2022", 200, 100],
    ["barra", "VL_N03_CEBRASPE_ESCRITA_REG_2022", "bh_escrita_n3_2022", 200, 100],
    ["barra", "VL_N04_CEBRASPE_ESCRITA_REG_2022", "bh_escrita_n4_2022", 200, 100],
    
    ["barra", "VL_N01_CEBRASPE_ESCRITA_REG_2023", "bh_escrita_n1_2023", 200, 100],
    ["barra", "VL_N02_CEBRASPE_ESCRITA_REG_2023", "bh_escrita_n2_2023", 200, 100],
    ["barra", "VL_N03_CEBRASPE_ESCRITA_REG_2023", "bh_escrita_n3_2023", 200, 100],
    ["barra", "VL_N04_CEBRASPE_ESCRITA_REG_2023", "bh_escrita_n4_2023", 200, 100], */
  ];

  for (var m = 0; m < pontos_graficos.length; m++) {
    var ponto = pontos_graficos[m];
    var tipo = ponto[0];
    var fator = ponto[4];
    var tamanho_maximo = ponto[3];

    var redimensionar = doc.pageItems.item(ponto[2]);

    var redimensionar_valor = jsonItem[ponto[1]].toString();

    var redimensionar_calculo = calcula(redimensionar_valor, fator);

    if (tipo == "barra") {
      if (
        redimensionar_valor.match("-") ||
        redimensionar_valor == 0 ||
        redimensionar_valor == "" ||
        redimensionar_valor == null
      ) {
        redimensionar.visible = false;

        redimensionar.resize(
          CoordinateSpaces.INNER_COORDINATES,
          AnchorPoint.LEFT_CENTER_ANCHOR,
          ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
          [5, 10]
        );
      } else {
        redimensionar.visible = true;
        redimensionar.resize(
          CoordinateSpaces.INNER_COORDINATES,
          AnchorPoint.LEFT_CENTER_ANCHOR,
          ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
          [redimensionar_calculo, 10]
        );
      }
    }

    function calcula(texto, fator) {
      var novo = Math.abs(texto * tamanho_maximo) / fator;

      return novo;
    }
  }
}

function coloreHabilidades() {
  var doc = app.activeDocument;
  var docTables = doc.stories.everyItem().tables.everyItem().getElements();
  var acerto1 = doc.colors.item("acerto1");
  var acerto2 = doc.colors.item("acerto2");
  var acerto3 = doc.colors.item("acerto3");
  var acerto4 = doc.colors.item("acerto4");
  var acertobase = doc.colors.item("acertobase");

  for (var i = 0; i < docTables.length; i++) {
    var table = docTables[i];

    var celulas = table.cells;

    for (var j = 0; j < celulas.length; j++) {
      var celula = celulas[j];
      var cellStyle = celula.appliedCellStyle;

      if (cellStyle.name === "matriz-acerto") {
        var taxa = celula.contents;

        if (taxa > 80) {
          celula.fillColor = acerto4;
        } else if (taxa > 60 && taxa <= 80) {
          celula.fillColor = acerto3;
        } else if (taxa > 40 && taxa <= 60) {
          celula.fillColor = acerto2;
        } else if (taxa <= 40) {
          celula.fillColor = acerto1;
        }
      }
    }
  }
}
