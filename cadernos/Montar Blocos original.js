#target indesign;

var w = new Window("dialog");
//~ w.preferredSize = [300, 250];
w.preferredSize.width = 300;
//~ w.preferredSize.height = 250;
w.DD = w.add('dropdownlist', undefined, ["Multiplos Blocos", "Blocão"]);
w.DD.selection = 0;
w.DD.alignment = ["fill", "top"];


var caminhoTxt;
var caminhoCSV;
var caminhoPasta;
var caminhoIndd;
var caminhoCSV2;
var caminhoIndd2;



// Adicionando botões para selecionar pastas na opção A
var selectFolderButton1 = w.add("button", undefined, "Selecionar TXT");
selectFolderButton1.alignment = ["fill", "top"];
selectFolderButton1.onClick = function () {
    caminhoTxt = File.openDialog("Selecione o arquivo de texto (*.txt)", "*.txt")
    
    if (caminhoTxt) {
        updateSelectedPathText1(caminhoTxt.fsName); // Atualizar o texto com o caminho selecionado
    }
};

var selectFolderButton2 = w.add("button", undefined, "Selecionar CSV");
selectFolderButton2.alignment = ["fill", "top"];
selectFolderButton2.onClick = function () {
    caminhoCSV = File.openDialog("Selecione o arquivo CSV (*.csv)", "*.csv");
      
    if (caminhoCSV) {
        updateSelectedPathText2(caminhoCSV.fsName); // Atualizar o texto com o caminho selecionado
    }

};

var selectFolderButton3 = w.add("button", undefined, "Selecionar INDT");
selectFolderButton3.alignment = ["fill", "top"];
selectFolderButton3.onClick = function () {
    caminhoPasta = Folder.selectDialog("Selecione a pasta contendo o arquivo Bloco.indt"); // Retorna String com caminho
    
    if (caminhoPasta) {
        updateSelectedPathText3(caminhoPasta.fsName); // Atualizar o texto com o caminho selecionado

    }
    
};

var selectFolderButton4 = w.add("button", undefined, "Selecionar INDD");
selectFolderButton4.alignment = ["fill", "top"];
selectFolderButton4.onClick = function () {
    caminhoIndd = Folder.selectDialog("Selecione a pasta com os arquivos em .indd");// Retorna String com caminho
    
    if (caminhoIndd) {
        updateSelectedPathText4(caminhoIndd.fsName); // Atualizar o texto com o caminho selecionado
    }
};

//------------------------------------------------------------------------------------------------------------------------------------------------------------------------
// Adicionando botões para selecionar pastas na opção B
var selectFolderButton5 = w.add("button", undefined, "Selecionar CSV");
selectFolderButton5.alignment = ["fill", "top"];
selectFolderButton5.onClick = function () {
    
    caminhoCSV2 = File.openDialog("Selecione o arquivo CSV (*.csv)", "*.csv");
      
    if (caminhoCSV2) {
        updateSelectedPathText5(caminhoCSV2.fsName); // Atualizar o texto com o caminho selecionado
    }

};

var selectFolderButton6 = w.add("button", undefined, "Selecionar INDD");
selectFolderButton6.alignment = ["fill", "top"];
selectFolderButton6.onClick = function () {
    
    caminhoIndd2 = Folder.selectDialog("Selecione a pasta com os arquivos em .indd");// Retorna String com caminho
    
    if (caminhoIndd2) {
        updateSelectedPathText6(caminhoIndd2.fsName); // Atualizar o texto com o caminho selecionado
    }
};
//--------------------------------------------------------------------------------------------------------------------------------------------------------------------------


// Adicionando botão "OK" para iniciar o loop de script na opção A
var okButton = w.add("button", undefined, "Gerar");
okButton.alignment = ["middle", "top"];
okButton.size = [100,40];
okButton.visible = true;
okButton.onClick = main;


// Adicionando botão "OK2" para a opção B
var okButton2 = w.add("button", undefined, "Gerar");
okButton2.alignment = ["middle", "top"];
okButton.size = [100,40];
okButton2.visible = true;
okButton2.onClick = main;



//----------------------------------------------------------------------------------------------------------------------///----------------------------------------------------------------------


//Setting a stacked group to items overlay
w.gp = w.add('group');
w.gp.orientation = 'stack';
w.gp.alignment = ["fill", "top"];
w.gp.alignChildren = ["fill", "top"];
w.gp.p1 = w.gp.add('panel', undefined, "Multiplos Blocos");
w.gp.p2 = w.gp.add('panel', undefined, "Blocão");
w.gp.p2.visible = false; // Inicia invisível
okButton2.visible = false;
selectFolderButton5.visible = false;
selectFolderButton6.visible = false;

var staticText1 = w.gp.p1.add('statictext', undefined, '')
staticText1.alignment = ["fill", "top"];

var staticText2 = w.gp.p1.add('statictext', undefined, '')
staticText2.alignment = ["fill", "top"];

var staticText3 = w.gp.p1.add('statictext', undefined, '')
staticText3.alignment = ["fill", "top"];

var staticText4 = w.gp.p1.add('statictext', undefined, '')
staticText4.alignment = ["fill", "top"];

var staticText7 = w.gp.p2.add('statictext', undefined, '')
staticText7.alignment = ["fill", "top"];

var staticText5 = w.gp.p2.add('statictext', undefined, '')
staticText5.alignment = ["fill", "top"];

var staticText6 = w.gp.p2.add('statictext', undefined, '')
staticText6.alignment = ["fill", "top"];




function updateSelectedPathText1(path) {
    staticText1.text = "Caminho TXT: OK ";
    }

    function updateSelectedPathText2(path) {
    staticText2.text = "Caminho CSV: OK ";
}


function updateSelectedPathText3(path) {
    staticText3.text = "Caminho INDT: OK ";
}


function updateSelectedPathText4(path) {
    staticText4.text = "Caminho INDDs: OK ";
}

    function updateSelectedPathText5(path) {
    staticText5.text = "Caminho CSV: OK ";
}

function updateSelectedPathText6(path) {
    staticText6.text = "Caminho INDDs: OK ";
    }
w.DD.onChange = function () {
    if (w.DD.selection == 0) { // Opção A selecionada
        selectFolderButton1.visible = true;
        selectFolderButton2.visible = true;
        selectFolderButton3.visible = true;
        selectFolderButton4.visible = true;
        selectFolderButton5.visible = false;
        selectFolderButton6.visible = false;
        w.gp.p1.visible = true;
        w.gp.p2.visible = false;
        okButton.visible = true;
        okButton2.visible = false;
          
         if (caminhoTxt) {
            updateSelectedPathText1(caminhoTxt.fsName);
        }
        if (caminhoCSV) {
            updateSelectedPathText2(caminhoCSV.fsName);
        }
        if (caminhoPasta) {
            updateSelectedPathText3(caminhoPasta.fsName);
        }
        if (caminhoIndd) {
            updateSelectedPathText4(caminhoIndd.fsName);            
        }

    } else { // Opção B selecionada
        selectFolderButton1.visible = false;
        selectFolderButton2.visible = false;
        selectFolderButton3.visible = false;
        selectFolderButton4.visible = false;
        selectFolderButton5.visible = true;
        selectFolderButton6.visible = true;
        w.gp.p1.visible = false;
        w.gp.p2.visible = true;
        okButton.visible = false;
        okButton2.visible = true;
        
        if (app.documents.length > 0) {
                            staticText7.text  = "Arquivo do Blocão Aberto"
                            } else {

                            staticText7.text  = "Não esquecer: Gerar com o .indd aberto"
                           
                            }
        
        if (caminhoCSV2) {
            updateSelectedPathText5(caminhoCSV2.fsName);
        }
        if (caminhoIndd2) {
            updateSelectedPathText6(caminhoIndd2.fsName);
        }
  
    }
   w.gp.p1.size = [270, 130];
   w.gp.p2.size = [270, 130];
//~    w.gp.size.height = w.gp.p1.visible ? 400 : 400
}

w.show();


function lerArquivoTexto() {
   
   // Se o usuário cancelar a seleção do arquivo, caminhoTxt será null
    if (!caminhoTxt) {
        
        // Encerrar o script, já que não há arquivo selecionado
        alert("Nenhum arquivo de texto selecionado.");
        exit();
    }    

    var arquivo = File(caminhoTxt);
    arquivo.open("r");
    var Blocos = [];
    while (!arquivo.eof) {
        Blocos.push(arquivo.readln());
    }
    arquivo.close();
    return Blocos;
   
}

// Função para ler o conteúdo de um arquivo CSV e armazenar as linhas
function lerArquivoCSV() {
    
    var arquivo = File(caminhoCSV);
    arquivo.open("r");
    var linhasbloco = [];
    
    if (!caminhoCSV) {
        // Encerrar o script, já que não há arquivo selecionado
        alert("Nenhum arquivo CSV selecionado.");
        exit();
    }
    else{
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

  // Função para ler o conteúdo de um arquivo CSV e armazenar as linhas
        function lerArqCSV() {
    
       var arquivo2 = File(caminhoCSV2);
              arquivo2.open("r");
       var linhasbloco2 = [];
    
        if (!caminhoCSV2) {
        // Encerrar o script, já que não há arquivo selecionado
        alert("Nenhum arquivo CSV selecionado.");
        exit();
        }
            else{
            var linhas2 = arquivo2.read().split("\n"); // Verificar melhor uso.
            for (var i = 0; i < linhas2.length; i++) {
            // Dividir a linha em elementos com base em uma vírgula (pode ser ajustado)
            var elementos2 = linhas2[i].split(",");
            // Adicionar os elementos ao array linhasbloco
            linhasbloco2.push(elementos2);
        }
        arquivo2.close();
        return linhasbloco2;
    }
}


//---------------------------------------------------------------escolha dos oks--------------------------

var valor;

    


function main() {
      
    if (okButton.visible === true) {
    
    valor = 90;
    
    }   

    if (okButton2.visible === true) {
    
    valor = 110;
    
    }   

    if (valor > 100) {
        
            menu2(caminhoCSV2, caminhoIndd2);

    
    } else {
            
            menu1(caminhoTxt, caminhoCSV, caminhoPasta, caminhoIndd);
    }
}
main();

//--------------------------------------------Montagem de vários blocos-------------------------------------------------------------------------------



function menu1(caminhoTxt, caminhoCSV, caminhoPasta, caminhoIndd){
                
               
                    // Função para ler o conteúdo de um arquivo de texto e armazenar os nomes dos arquivos


// Função para manipular o conteúdo do documento InDesign
function manipularDocumentoIndesign(caminhoBlocoIndt, listaLinha, blocos, caminhoIndd) {
    
    w.close();
    
   
                documento= app.open(File(caminhoBlocoIndt));
    
    // lista os arquivos de caminhoIndd
    var fileList = caminhoIndd.getFiles();

    //Ordena os arquivos
    fileList.sort();
    
    

    // Filtrando os arquivos existentes na pasta de origem (parte copiada)
    var ListaArqPast = [] // Retorna lista filtrada
    var naoindd = [] // Outros arquivos
    
    //adiciona a lista sortedFileList, arquivos indd e a naoindd arquivos diversos
    for (var h = 0; h <= fileList.length; h++) {
        var file = fileList[h];
        if (file instanceof File && file.name.match(/\.indd$/i)) {
            ListaArqPast.push(file);
        }
        else{
            naoindd.push(file)
        }
}

        
    // Verificado 20032024 12:37
    var verifyListCsv =[];
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
        var lastTextFrame = documento.pages.item(-1).textFrames[documento.pages.item(-1).textFrames.length - 1];
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
        var lastTextFrame = previousPage.textFrames[previousPage.textFrames.length - 1];
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

for (i = (masters.length -1); i > 1 ; i--) {

    masters[i].remove();
}
            }
        }
    }
}    

            
            // Salva o documento original com o nome baseado em 'blocos'
            var pasta = new File(caminhoBlocoIndt).parent;
            documento.save(File(pasta + "/" + blocos + ".indd"));
            documento.close(SaveOptions.YES); 
            
            
       
}

                    
                    function laco(listaLinhas, blocos, caminhoPasta, caminhoIndd){
                    //Etapa 4: Início do loop de criação de blocos
                    
                    if (File(caminhoPasta + "/Bloco.indt").exists) {
                     
                     caminhoBlocoIndt = caminhoPasta + "/Bloco.indt" //defindo caminho bloco
                       
                    } else {
                    
                    alert("O arquivo Bloco.indt não está presente na pasta.");
                    exit();
                    
                    }
                     
                    // asubtração do tamanho da lista se deve a linha em braco
                    for (var i = 0; i < listaLinhas.length-1; i++) {
                     //Etapa 5: Manipular o arquivo InDesign
                    manipularDocumentoIndesign(caminhoBlocoIndt, listaLinhas[i], blocos[i], caminhoIndd);
                    } 
                    // Alerta de fim de execução
        
                        alert("Blocos montados");
                        exit();

                    } 
                        var blocos = lerArquivoTexto(); //Recebe lista (X,Y,Z)
    
                        var listaLinhas = lerArquivoCSV(); //Recebe tupla [(a,b,c),(d,e,f),(g,h,i,j)]
    
                    laco(listaLinhas, blocos, caminhoPasta, caminhoIndd)
                    
}
menu1(caminhoTxt, caminhoCSV, caminhoPasta, caminhoIndd);


//-------------------------------------------------------------------------------  Bloção  -------------------------------------------------------------------------------



function menu2(caminhoCSV2, caminhoIndd2){
 
    
       w.close();
            


    // Função para manipular o conteúdo do documento InDesign
    
                function manipularDocumentoIndesign2(caminhoBlocoIndt2, listaLinha2, caminhoIndd2) {
                    
                w.close();
    
                documento2 = caminhoBlocoIndt2;
                   
                // lista os arquivos de caminhoIndd
                var fileList2 = caminhoIndd2.getFiles();
    
                //Ordena os arquivos
                fileList2.sort();
    
    

                // Filtrando os arquivos existentes na pasta de origem (parte copiada)
                var ListaArqPast2 = [] // Retorna lista filtrada
                var naoindd2 = [] // Outros arquivos
    
                //adiciona a lista sortedFileList, arquivos indd e a naoindd arquivos diversos
                for (var h = 0; h <= fileList2.length; h++) {
                        var file = fileList2[h];
                        if (file instanceof File && file.name.match(/\.indd$/i)) {
                        ListaArqPast2.push(file);
                        }
                            else {
                            naoindd2.push(file)
                       }
                }

        
               // Verificado 20032024 12:37
                var verifyListCsv2 =[];
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
              
                if (ListaArqPast2[f].name.match(new RegExp("^" + fileName2 + "$", "i"))) {
                    
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
                sourcePages2.duplicate(LocationOptions.AFTER, source_doc2.pages.item(0));
                sourcePages2.move(LocationOptions.AFTER, documento2.pages.item(-1));

                // Adicionar uma linha em branco separadora ao final do texto na última caixa de texto da página atual
                var lastTextFrame2 = documento2.pages.item(-1).textFrames[documento2.pages.item(-1).textFrames.length - 1];
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
                        var lastTextFrame = previousPage.textFrames[previousPage.textFrames.length - 1];
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

for (i = (masters.length -1); i > 1 ; i--) {

    masters[i].remove();
}

            }
        }
    }
}    
       
        exit();
    
}

      
                        function laco2(listaLinhas2, caminhoIndd2){

                            if (app.documents.length > 0) {
                            caminhoBlocoIndt2 = app.activeDocument;
                            } else {

                             alert("Nenhum documento aberto!")
                             exit();
                            }

                            
                           

                            // asubtração do tamanho da lista se deve a linha em braco
                            for (var i = 0; i < listaLinhas2.length-1; i++) {
                            //Etapa 5: Manipular o arquivo InDesign
                            manipularDocumentoIndesign2(caminhoBlocoIndt2, listaLinhas2[i], caminhoIndd2);
                
                            } 
                            // Alerta de fim de execução
                                alert("Bloco montado");
                                exit();
                            } 
                                                            
                                var listaLinhas2 = lerArqCSV(); //Recebe tupla [(a,b,c),(d,e,f),(g,h,i,j)]

                                laco2(listaLinhas2, caminhoIndd2)
                                
                      }
menu2(caminhoCSV2, caminhoIndd2);
