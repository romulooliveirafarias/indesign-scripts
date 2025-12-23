//Resetando as preferências de PDF do Indesign
app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;
app.pdfExportPreferences.viewPDF = false;

var doc = app.activeDocument;

// Declaração dos elementos da capa
var escola = doc.pageItems.item("escola");
var municipio = doc.pageItems.item("municipio");
var crede = doc.pageItems.item("crede");
var alunos = doc.pageItems.item("alunos");

// Caixas de diálogo para selelecionar pasta onde estão máscaras e pasta onde serão salvas as capas
var caminhoLaudos =
  Folder.selectDialog("Selecione a pasta onde deseja salvar os arquivos")
    .absoluteURI + "/";
var caminhoJSON = File.openDialog("Selecione um arquivo JSON");

var logFile = new File("C:/Users/romulo.farias/Documents/LAUDOS CE/log.txt");
logFile.open("w");
logFile.write("Log de execução: " + new Date() + "\n");
logFile.close();

caminhoJSON.open("r");

var infoObject = caminhoJSON.read();

// Transforma o obejto em objeto json interpretável pelo indeisgn
var newJson = eval("(" + infoObject + ")");

var jsonItens = newJson.itens;

for (i = 0; i < jsonItens.length; i++) {
  var jsonItem = jsonItens[i];

  puxaDados1(jsonItem);
  criaTabela();
  /* criaUltimoCaracater(); */
  formataTabela();
  puxaDados2(jsonItem);
}

function puxaDados1(item) {
  alunos.contents = item.alunos;
}

function puxaDados2(item) {
  escola.contents = item.escola;
  municipio.contents = item.municipio;
  crede.contents = item.crede;

  logFile.open("a");
  logFile.write(item.escola + "\n");
  logFile.close();

  $.sleep(3000);

  removeBlankPages();
}

function criaTabela() {
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
}

function criaUltimoCaracater() {
  logFile.open("a");

  var myTable = app.activeDocument.stories
    .everyItem()
    .tables.everyItem()
    .getElements();

  var lastCell = myTable[0].cells[myTable[0].cells.length - 1];
  if (lastCell.contents.length > 0) {
    lastCell.contents = lastCell.contents
    logFile.write("Última célula: " + lastCell.contents + "\n");

  } else {
    lastCell.contents = " ";
    logFile.write("Nova última célula: " + lastCell.contents + "\n");
  }
  logFile.close();
}

function formataTabela() {
  var myTable = app.activeDocument.stories
    .everyItem()
    .tables.everyItem()
    .getElements();

  myTable[0].appliedTableStyle = "laudos";
  myTable[0].cells[4].width = 12;
  myTable[0].cells[5].width = 25;
  myTable[0].cells[6].width = 81;
  myTable[0].cells[0].rowType = RowTypes.HEADER_ROW;
}

function saveDoc(item) {
  var tipo_pdf = app.pdfExportPresets.item("[Smallest File Size]");
  var pdfname = caminhoLaudos + item.cd_escola + ".pdf";
  doc.exportFile(ExportFormat.PDF_TYPE, new File(pdfname), false, tipo_pdf);
}

function resetDoc() {
  // Reseta para os campos originais vazios
  escola.contents = "";
  municipio.contents = "";
  crede.contents = "";
  alunos.contents = "";
}

function removeBlankPages() {
  logFile.open("a");

  var myTable = app.activeDocument.stories
    .everyItem()
    .tables.everyItem()
    .getElements();

  var lastCell = myTable[0].cells[myTable[0].cells.length - 1];

  var lastPageContents =
    lastCell.insertionPoints[0].parentTextFrames[0].parentPage.name;

  logFile.write("Última célula: " + lastCell.contents + "\n");
  logFile.write("Última página com conteúdo: " + lastPageContents + "\n");

  if (doc.pages.length > 1) {
    var lastPage = doc.pages[doc.pages.length - 1];

    logFile.write("Última página: " + lastPage.name + "\n");

    if (lastPage.name > lastPageContents) {
      logFile.write(
        "Removendo última página: " + lastPage.name + "\n---------/\n"
      );
      lastPage.remove();

      logFile.close();

      removeBlankPages();
    } else {
      saveDoc(jsonItem);
      resetDoc();
    }
  } else {
    saveDoc(jsonItem);
    resetDoc();
  }
}
