var doc = app.activeDocument;
var paginas = doc.pages;

var mestras = doc.masterSpreads;
var sexta = [];

for (j = 0; j < paginas.length; j++) {
  var pagina = paginas[j];

  var caixas = pagina.textFrames;

  for (i = 0; i < caixas.length; i++) {
    conteudo = caixas[i].contents;

    if (conteudo.match("Segunda-feira")) {
      for (k = 0; k < conteudo.length; k++) {
        if (k % 2 == 0) {
          pagina.appliedMaster = doc.masterSpreads.item("SE1-PV");
        } else {
          pagina.appliedMaster = doc.masterSpreads.item("SE2-PN");
        }
      }
    } else if (conteudo.match("Terça-feira")) {
      pagina.appliedMaster = doc.masterSpreads.item("TER--");
    } else if (conteudo.match("Quarta-feira")) {
      pagina.appliedMaster = doc.masterSpreads.item("QUA--");
    } else if (conteudo.match("Quinta-feira")) {
      pagina.appliedMaster = doc.masterSpreads.item("QUI--");
    } else if (conteudo.match("Sexta-feira")) {
      for (k = 0; k < conteudo.length; k++) {
        if (k % 2 == 0) {
          pagina.appliedMaster = doc.masterSpreads.item("SX1-VD");
        } else {
          pagina.appliedMaster = doc.masterSpreads.item("SX2-GE");
        }
      }
    } else if (conteudo.match("Sábado")) {
      pagina.appliedMaster = doc.masterSpreads.item("SD--");
    }
  }
}
