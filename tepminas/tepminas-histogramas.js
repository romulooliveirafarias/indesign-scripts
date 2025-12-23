var doc = app.activeDocument;
var docPath = doc.filePath;
var docname = doc.name.split(".")[0];
var docGerando = docPath + "/" + docname + "_gerando.indd";
doc.save(new File(docGerando));

var nmInstituicao = doc.pageItems.item("NM_INSTITUICAO");
var nmEstudante = doc.pageItems.item("NM_ESTUDANTE");
var cpfEstudante = doc.pageItems.item("CPF_ESTUDANTE");
var periodoEstudante = doc.pageItems.item("PERIODO_ESTUDANTE");
var vlNotaEstudante = doc.pageItems.item("nota").pageItems.item("VL_NOTA_ESTUDANTE");
var vl6 = doc.pageItems.item("vl-6");
var vl12 = doc.pageItems.item("vl-12");
var vl18 = doc.pageItems.item("vl-18");
var vl24 = doc.pageItems.item("vl-24");
var vl30 = doc.pageItems.item("vl-30");
var vl36 = doc.pageItems.item("vl-36");
var vl42 = doc.pageItems.item("vl-42");
var vl48 = doc.pageItems.item("vl-48");
var vl54 = doc.pageItems.item("vl-54");
var vl60 = doc.pageItems.item("vl-60");
var vl66 = doc.pageItems.item("vl-66");
var vl72 = doc.pageItems.item("vl-72");
var vl78 = doc.pageItems.item("vl-78");
var vl84 = doc.pageItems.item("vl-84");
var vl90 = doc.pageItems.item("vl-90");
var vl96 = doc.pageItems.item("vl-96");
var vl102 = doc.pageItems.item("vl-102");
var vl108 = doc.pageItems.item("vl-108");
var vl114 = doc.pageItems.item("vl-114");
var vl120 = doc.pageItems.item("vl-120");
var col6 = doc.pageItems.item("col-6");
var col12 = doc.pageItems.item("col-12");
var col18 = doc.pageItems.item("col-18");
var col24 = doc.pageItems.item("col-24");
var col30 = doc.pageItems.item("col-30");
var col36 = doc.pageItems.item("col-36");
var col42 = doc.pageItems.item("col-42");
var col48 = doc.pageItems.item("col-48");
var col54 = doc.pageItems.item("col-54");
var col60 = doc.pageItems.item("col-60");
var col66 = doc.pageItems.item("col-66");
var col72 = doc.pageItems.item("col-72");
var col78 = doc.pageItems.item("col-78");
var col84 = doc.pageItems.item("col-84");
var col90 = doc.pageItems.item("col-90");
var col96 = doc.pageItems.item("col-96");
var col102 = doc.pageItems.item("col-102");
var col108 = doc.pageItems.item("col-108");
var col114 = doc.pageItems.item("col-114");
var col120 = doc.pageItems.item("col-120");
var nota = doc.textFrames.item("nota");

// Caixas de diálogo para selelecionar pasta onde estão máscaras e pasta onde serão salvas as capas

var caminhoHistogramas =
  Folder.selectDialog("Selecione a pasta onde deseja salvar os histogramas")
    .absoluteURI + "/";
var caminhoJSON = File.openDialog("Selecione um arquivo JSON");

caminhoJSON.open("r");

var infoObject = caminhoJSON.read();

// Transforma o obejto em objeto json interpretável pelo indeisgn
var newJson = eval("(" + infoObject + ")");

var jsonItens = newJson.itens;

alert(jsonItens.length);

for (i = 0; i < jsonItens.length; i++) {
  var jsonItem = jsonItens[i];

  nmInstituicao.contents = jsonItem.NM_INSTITUICAO;
  nmEstudante.contents = jsonItem.NM_ESTUDANTE;
  cpfEstudante.contents = jsonItem.CPF_ESTUDANTE;
  periodoEstudante.contents = jsonItem.PERIODO_ESTUDANTE;
  vlNotaEstudante.contents = jsonItem.VL_NOTA_ESTUDANTE.replace(',','.');
  vl6.contents = jsonItem.TX_I01.replace(',','.');
  vl12.contents = jsonItem.TX_I02.replace(',','.');
  vl18.contents = jsonItem.TX_I03.replace(',','.');
  vl24.contents = jsonItem.TX_I04.replace(',','.');
  vl30.contents = jsonItem.TX_I05.replace(',','.');
  vl36.contents = jsonItem.TX_I06.replace(',','.');
  vl42.contents = jsonItem.TX_I07.replace(',','.');
  vl48.contents = jsonItem.TX_I08.replace(',','.');
  vl54.contents = jsonItem.TX_I09.replace(',','.');
  vl60.contents = jsonItem.TX_I10.replace(',','.');
  vl66.contents = jsonItem.TX_I11.replace(',','.');
  vl72.contents = jsonItem.TX_I12.replace(',','.');
  vl78.contents = jsonItem.TX_I13.replace(',','.');
  vl84.contents = jsonItem.TX_I14.replace(',','.');
  vl90.contents = jsonItem.TX_I15.replace(',','.');
  vl96.contents = jsonItem.TX_I16.replace(',','.');
  vl102.contents = jsonItem.TX_I17.replace(',','.');
  vl108.contents = jsonItem.TX_I18.replace(',','.');
  vl114.contents = jsonItem.TX_I19.replace(',','.');
  vl120.contents = jsonItem.TX_I20.replace(',','.');
  col6.contents = jsonItem.TX_I01.replace(',','.');
  col12.contents = jsonItem.TX_I02.replace(',','.');
  col18.contents = jsonItem.TX_I03.replace(',','.');
  col24.contents = jsonItem.TX_I04.replace(',','.');
  col30.contents = jsonItem.TX_I05.replace(',','.');
  col36.contents = jsonItem.TX_I06.replace(',','.');
  col42.contents = jsonItem.TX_I07.replace(',','.');
  col48.contents = jsonItem.TX_I08.replace(',','.');
  col54.contents = jsonItem.TX_I09.replace(',','.');
  col60.contents = jsonItem.TX_I10.replace(',','.');
  col66.contents = jsonItem.TX_I11.replace(',','.');
  col72.contents = jsonItem.TX_I12.replace(',','.');
  col78.contents = jsonItem.TX_I13.replace(',','.');
  col84.contents = jsonItem.TX_I14.replace(',','.');
  col90.contents = jsonItem.TX_I15.replace(',','.');
  col96.contents = jsonItem.TX_I16.replace(',','.');
  col102.contents = jsonItem.TX_I17.replace(',','.');
  col108.contents = jsonItem.TX_I18.replace(',','.');
  col114.contents = jsonItem.TX_I19.replace(',','.');
  col120.contents = jsonItem.TX_I20.replace(',','.');

  redimensiona();

  var tipo_pdf = app.pdfExportPresets.item("[Smallest File Size]");
  var pdfname = caminhoHistogramas + jsonItem.CPF_ESTUDANTE + ".pdf";

  doc.exportFile(ExportFormat.PDF_TYPE, new File(pdfname), false, tipo_pdf);

  reset()

  function reset(){
    nmInstituicao.contents = "";
    nmEstudante.contents = "";
    cpfEstudante.contents = "";
    periodoEstudante.contents = "";
    nota.move([42.5, 301.028]); 
  
}

}

function redimensiona() {
  var itens = doc.textFrames;
  doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.PIXELS;
  doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.PIXELS;
  var tamanho_maximo = 263.15;
  
  var notaVl = Math.round(nota.textFrames.item("VL_NOTA_ESTUDANTE").contents);
  var posX = 9.038;
  var breaks = [
    0, 6, 12, 18, 24, 30, 36, 42, 48, 54, 60, 66, 72, 78, 84, 90, 96, 102, 108,
    114, 120,
  ];

  for (j = 0; j < itens.length; j++) {
    var item = itens[j];

    if (item.name.match("col")) {
      var resize_calc = calc(item.contents, 50);

      if(item.contents == 0 || item.contents == 0){
        item.fillColor = 'transparent';
      } else{
        item.fillColor = 'azul-claro';

      }

      item.resize(
        CoordinateSpaces.INNER_COORDINATES,
        AnchorPoint.BOTTOM_CENTER_ANCHOR,
        ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
        [24.18, resize_calc]
      );
    } else if (item.name.match("vl")) {
      var resize_calc = calc(item.contents, 50);
      item.resize(
        CoordinateSpaces.INNER_COORDINATES,
        AnchorPoint.BOTTOM_CENTER_ANCHOR,
        ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
        [24.18, Math.abs(resize_calc + 10)]
      );
      item.contents = item.contents + "%";
    }
  }

  function calc(texto, fator) {
    if (texto == "-" || texto == 0) {
      var newSize = 7.11;
      return newSize;
    } else {
      var newSize = Math.abs(texto * tamanho_maximo) / fator;
      return newSize;
    }
  }

  for (k = 0; k < breaks.length; k++) {
    if (notaVl > breaks[k] && notaVl <= breaks[k + 1]) {
      var col = doc.textFrames.item("col-" + breaks[k + 1]);
      var vl = doc.textFrames.item("vl-" + breaks[k + 1]);
      var posNota = vl.geometricBounds[2];
      col.fillColor = "azul-escuro";
      newX = Math.abs(42.5 + 25.6 * k);
      nota.move([newX, Math.abs(col.geometricBounds[0] - 50)]);
    }
  }
}
