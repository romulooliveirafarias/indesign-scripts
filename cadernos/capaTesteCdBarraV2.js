//Versão 1.3 - 24.01.24

#target indesign;
#targetengine "session";
//////////////

//Resetando as preferências de PDF do Indesign
app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;
app.pdfExportPreferences.viewPDF = false;

var doc = app.activeDocument;

// Declaração dos elementos da capa
var base = doc.pageItems.item("base")
var base45 = doc.pageItems.item("base45")
var baseTransc = doc.pageItems.item("baseTransc")
var base45Transc = doc.pageItems.item("base45Transc")

var artigoAvaliacaoBase = doc.pageItems.item("base").pageItems.item("artigoAvaliacao");
var artigoAvaliacaoBase45 = doc.pageItems.item("base45").pageItems.item("artigoAvaliacao");
var artigoAvaliacaoBaseTransc = doc.pageItems.item("baseTransc").pageItems.item("artigoAvaliacao");
var artigoAvaliacaoBase45Transc = doc.pageItems.item("base45Transc").pageItems.item("artigoAvaliacao");

var avaliacao = doc.pageItems.item("avaliacao");
var avaliacaoBase = doc.pageItems.item("base").pageItems.item("avaliacao");
var avaliacaoBase45 = doc.pageItems.item("base45").pageItems.item("avaliacao");
var avaliacaoBaseTransc = doc.pageItems.item("baseTransc").pageItems.item("avaliacao");
var avaliacaoBase45Transc = doc.pageItems.item("base45Transc").pageItems.item("avaliacao");

var ano = doc.pageItems.item("ano");
var anoBase = doc.pageItems.item("base").pageItems.item("ano");
var anoBase45 = doc.pageItems.item("base45").pageItems.item("ano");
var anoBaseTransc = doc.pageItems.item("baseTransc").pageItems.item("ano");
var anoBase45Transc = doc.pageItems.item("base45Transc").pageItems.item("ano");

var disciplina = doc.pageItems.item("disciplina");
var disciplinaBase = doc.pageItems.item("base").pageItems.item("disciplina");
var disciplinaBase45 = doc.pageItems.item("base45").pageItems.item("disciplina");
var disciplinaBaseTransc = doc.pageItems.item("baseTransc").pageItems.item("disciplina");
var disciplinaBase45Transc = doc.pageItems.item("base45Transc").pageItems.item("disciplina");

var questionarioBase = doc.pageItems.item("base").pageItems.item("questionario");
var questionarioBase45 = doc.pageItems.item("base45").pageItems.item("questionario");
var questionarioBaseTransc = doc.pageItems.item("baseTransc").pageItems.item("questionario");
var questionarioBase45Transc = doc.pageItems.item("base45Transc").pageItems.item("questionario");

var etapa = doc.pageItems.item("etapa");
var qr = doc.pageItems.item("qr");
var qrId = doc.textFrames.item("qrId");
var mascara = doc.pageItems.item("mascara");
var cdCaderno = doc.pageItems.item("cd_caderno");
var cdBarra = doc.pageItems.item("cd_barra");
var sequencial = doc.pageItems.item("sequencial128");
var caixaAplicador = doc.pageItems.item("caixa_aplicador");
var capaCor = doc.swatches.item("base");

// Caixas de diálogo para selelecionar pasta onde estão máscaras e pasta onde serão salvas as capas

var caminhoCapas = Folder.selectDialog("Selecione a pasta onde deseja salvar as capas").absoluteURI + "/";
var caminhoJSON = File.openDialog("Selecione um arquivo JSON");

  caminhoJSON.open('r');

  var infoObject = caminhoJSON.read();
  
  // Trasnforma o obejto em objeto json interpretável pelo indeisgn
  var newJson = eval("(" + infoObject + ")");
  var jsonMascara = newJson.mascara;

    if(jsonMascara !== "Não possui"){
      var caminhoMascaras = Folder.selectDialog("Selecione a pasta onde estão salvas as máscaras")
    }

  var jsonItens = newJson.itens;

  for ( i = 0; i < jsonItens.length; i++) {
    var jsonItem = jsonItens[i];

  // Para cada item do json informamos a chave pra retornar o valor. 
  // A partir do valor preenchemos os campos das capas.  
  
  if(artigoAvaliacaoBase.isValid){
    artigoAvaliacaoBase.contents = jsonItem.artigo_avaliacao;
    artigoAvaliacaoBase45.contents = jsonItem.artigo_avaliacao;
    artigoAvaliacaoBaseTransc.contents = jsonItem.artigo_avaliacao;
    artigoAvaliacaoBase45Transc.contents = jsonItem.artigo_avaliacao;
    }
  if(avaliacao.isValid){
  avaliacao.contents = jsonItem.avaliacao;
  }
  if(avaliacaoBase.isValid){
    avaliacaoBase.contents = jsonItem.avaliacao;
    avaliacaoBase45.contents = jsonItem.avaliacao;
    avaliacaoBaseTransc.contents = jsonItem.avaliacao;
    avaliacaoBase45Transc.contents = jsonItem.avaliacao;
    }
  if(ano.isValid){
    ano.contents = jsonItem.ano;}
    if(anoBase.isValid){
        anoBase.contents = jsonItem.ano;
        anoBase45.contents = jsonItem.ano;
        anoBaseTransc.contents = jsonItem.ano;
        anoBase45Transc.contents = jsonItem.ano;
      }
  if(disciplina){
    disciplina.contents = jsonItem.disciplina;
  }
  if(disciplinaBase){
    disciplinaBase.contents = jsonItem.disciplina;
    disciplinaBase45.contents = jsonItem.disciplina;
    disciplinaBaseTransc.contents = jsonItem.disciplina;
    disciplinaBase45Transc.contents = jsonItem.disciplina;
  }
  if(etapa.isValid){
    etapa.contents = jsonItem.etapa;
  }
  if(qr.isValid){
    qr.createPlainTextQRCode(jsonItem.qr);
  }
  if(qrId.isValid){
    qrId.contents = jsonItem.qr;
  }
  if(cdCaderno.isValid){
    cdCaderno.contents = jsonItem.cd_caderno;
  }
  if(mascara.isValid){
    mascara.place(new File(caminhoMascaras + "/" + jsonItem.mascara + ".pdf"));
  }
  
  if(capaCor.isValid){
    var cores = jsonItem.cor.split(",");
  parsecores = eval('['+ cores +']');
  capaCor.colorValue = parsecores;
  }
  
  if(caixaAplicador.isValid){
    if(jsonItem.tp_caderno == "_APLIC"){
      caixaAplicador.visible = true;
      if(qr.isValid){
        qr.visible = false;
        }
        if(qrId.isValid){
        qrId.visible = false;
        }
  
    }
    else{
      caixaAplicador.visible = false;
      if(qr.isValid){
      qr.visible = true;
      }
      if(qrId.isValid){
      qrId.visible = true;
      }
    }
  }

if(cdBarra.isValid){
    cdBarra.contents = jsonItem.codigo128;
    }

    if(sequencial.isValid){
        sequencial.contents = jsonItem.sequencial128;
        }
    if(questionarioBase.isValid){
        questionarioBase.contents = jsonItem.txt_questionario;
        questionarioBase45.contents = jsonItem.txt_questionario;
        questionarioBaseTransc.contents = jsonItem.txt_questionario;
        questionarioBase45Transc.contents = jsonItem.txt_questionario;
        }

        // Função para checar e substituir os ":" por tracinhos
function substituirPontosPorTracinhos(str) {
  if (str.match(":")) {
    // Substitui todos os ":" por tracinhos
    str = str.replace(/:/g, " - ");
  }
  return str;
}

var nm_avaliacao = substituirPontosPorTracinhos(jsonItem.avaliacao);

// salva o pdf da capa na pasta selecionada anteriormente.
  /* doc.saveACopy(File(caminhoCapas + jsonItem.avaliacao + " " + jsonItem.cd_caderno + ".indd")); */
  doc.saveACopy(File(caminhoCapas + nm_avaliacao + " " + jsonItem.nm_arquivo + ".indd"));

  if(jsonItem.caixa_base == "base"){
    base.visible = true;
    base45.visible = false;
    baseTransc.visible = false;
    base45Transc.visible = false;

  }else if (jsonItem.caixa_base == "base45"){
    base.visible = false;
    base45.visible = true;
    baseTransc.visible = false;
    base45Transc.visible = false;

  }else if (jsonItem.caixa_base == "baseTransc"){
    base.visible = false;
    base45.visible = false;
    baseTransc.visible = true;
    base45Transc.visible = false;
    
  }else if (jsonItem.caixa_base == "base45Transc"){
    base.visible = false;
    base45.visible = false;
    baseTransc.visible = false;
    base45Transc.visible = true;
    
  }

  // Reseta a capa para os campos originais vazios
  if(avaliacao.isValid){
    avaliacao.contents = "";
  }
  if(ano.isValid){
    ano.contents = "";
  }
  if(disciplina.isValid){
    disciplina.contents = "";
  }
  if(etapa.isValid){
    etapa.contents = "";
  }
 if(qr.isValid){
  qr.graphics[0].remove();

 }
  if(mascara.isValid){
    mascara.graphics[0].remove();
  }
  if(qrId.isValid){
    qrId.contents = "";
  }
  if(cdCaderno.isValid){
    cdCaderno.contents = "";
  }
 if(caixaAplicador.isValid){
  caixaAplicador.visible = false;

 }

 if(capaCor.isValid){
  
capaCor.colorValue = [0,0,0,50];
}

  }
