/*function doGet() {
  return HtmlService.createTemplateFromFile('Pagina').evaluate()
      .setTitle('Desempeño Ambiental')
      .setSandboxMode(HtmlService.SandboxMode.EMULATED);
}*/

/**
 * Get the URL for the Google Apps Script running as a WebApp.
 */
function getScriptUrl() {
 var url = ScriptApp.getService().getUrl();
 return url;
}

/**
 * Get "home page", or a requested page.
 * Expects a 'page' parameter in querystring.
 *
 * @param {event} e Event passed to doGet, with querystring
 * @returns {String/html} Html to be served
 */
 
function doGet(e) {
  Logger.log( Utilities.jsonStringify(e) );
  if (!e.parameter.page) {
    // When no specific page requested, return "home page"
    return HtmlService.createTemplateFromFile('Login').evaluate();
  }
  // else, use page parameter to pick an html file from the script
  var t=HtmlService.createTemplateFromFile(e.parameter['page']);
  t.dep=e.parameter.dep;
  var ss=SpreadsheetApp.openById('0AiSCXV9Wfr59dHNzRnNRTERJTEEyYU5GUWJia3FBMkE');
  t.general=ss.getSheetByName('General').getDataRange().getValues();
  t.energia_a=ss.getSheetByName('Energia A').getDataRange().getValues();
  t.energia_b=ss.getSheetByName('Energia B').getDataRange().getValues();
  t.agua=ss.getSheetByName('Agua').getDataRange().getValues();
  t.papel=ss.getSheetByName('Papel').getDataRange().getValues();
  t.residuos=ss.getSheetByName('Residuos').getDataRange().getValues();
  t.emisiones=ss.getSheetByName('Emisiones').getDataRange().getValues();
  return t.evaluate();

}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

//dependencia comienza en 1 y se resta 1, mes comienza en 1

function guardar(form){
   Logger.log('hoja: '+form.hoja);
   Logger.log('mes: '+form.mes);
   var a=new Array();
   var aux=new Array("a", "b", "c", "d","e", "f", "g", "h");
   for(var i=0; i<form.n; i++){
       a.push(form[aux[i]]);
   }
   Logger.log(a);
  escribir(form.hoja, form.dep, form.mes, a);
  
}

function guardarCombustibles(form){
  escribirCombustibles(form.hoja, form.dep, form.mes, form.tipo,form.consumo);
}


function escribirCombustibles(hoja, dependencia, mes, tipo, consumo){
  
  var ss=SpreadsheetApp.openById('0AiSCXV9Wfr59dHNzRnNRTERJTEEyYU5GUWJia3FBMkE');
  //var ss=SpreadsheetApp.openById('0AiSCXV9Wfr59dDNKcGhDc2tUTlFfVEJudlJYN3pfYXc');
  var sheet = ss.getSheetByName(hoja);

    var i=0;
    var fila=12*(dependencia-1)+parseInt(mes);
    sheet.getRange(fila+1, i+1).setValue(nombreDep(dependencia));
    sheet.getRange(fila+1, i+2).setValue(nombreMes(mes));
    
   
    sheet.getRange(fila+1, parseInt(tipo)+2).setValue(consumo);
  
}



function escribir(hoja, dependencia, mes, datos){
  
  var ss=SpreadsheetApp.openById('0AiSCXV9Wfr59dHNzRnNRTERJTEEyYU5GUWJia3FBMkE');
  //var ss=SpreadsheetApp.openById('0AiSCXV9Wfr59dDNKcGhDc2tUTlFfVEJudlJYN3pfYXc');
  var sheet = ss.getSheetByName(hoja);
  
  

  
  if(hoja!='Derrames'&&hoja!='Emisiones'&&hoja!='Efluentes'){

    var i=0;
    var fila=12*(dependencia-1)+parseInt(mes);
    sheet.getRange(fila+1, i+1).setValue(nombreDep(dependencia));
    sheet.getRange(fila+1, i+2).setValue(nombreMes(mes));
    for(i=0; i<datos.length; i++)
    sheet.getRange(fila+1, i+3).setValue(datos[i]);
  }
  else{
    datos.unshift(nombreDep(dependencia));
    sheet.appendRow(datos);
  }
 
  
}

function nombreDep(dep){
  var s='';
  switch(parseInt(dep)){
    case 1:
      s='Conchan';
      break;
    case 2:
      s='Comerciales';
      break;
    case 3:
      s='Oficina Principal';
      break;
    case 4:
      s='Oleoducto';
      break;
    case 5:
      s='Selva';
      break;
    case 6:
      s='Talara';
      break;
  }
  
  return s;
}

function nombreMes(mes){
  var nombres=new Array('Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Setiembre', 'Octubre', 'Noviembre', 'Diciembre');
  return nombres[parseInt(mes)-1];
}