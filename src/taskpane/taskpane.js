/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      // eslint-disable-next-line no-undef
      console.log("Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    // eslint-disable-next-line no-undef
    document.getElementById("create-content-control").onclick = createContentControl;
    document.getElementById("replace-content-in-control").onclick = replaceContentInControl;
    document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("reset-internal-count").onclick = resetInternalCount;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function insertParagraph() {
  Word.run(function (context) {
    var docBody = context.document.body.getRange("");
    var name = document.getElementById("name").value;
    var lastname = document.getElementById("lastname").value;
    var tax = document.getElementById("tax").value;
    var text = "Estimado contribuyente " 
      + name + " " 
      + lastname + " le informamos que debe impuestos " 
      + tax;

    var proemio = "Esta Administración Desconcentrada de Auditoría de Comercio Exterior del Norte Centro, con sede en Coahuila de Zaragoza, de la Administración General de Auditoría de Comercio Exterior del Servicio de Administración Tributaria, con fundamento en los artículos 16 de la Constitución Política de los Estados Unidos Mexicanos; 1, 7, fracciones VII y XVIII y 8, fracción III de la Ley del Servicio de Administración Tributaria publicada en el Diario Oficial de la Federación el 15 de diciembre de 1995, vigente a partir del 1 de julio de 1997, reformada mediante Decretos publicados en el Diario Oficial de la Federación en fechas 4 de enero de 1999, 12 de junio de 2003, 6 de mayo del 2009, 9 de abril del 2012, 17 de diciembre de 2015 y 4 de diciembre de 2018; 1, 2, primer párrafo, apartado C, y segundo párrafo, 5, tercer párrafo, 6, primer párrafo, apartado B, fracción II y último párrafo, 14, primer párrafo, fracción VI, 27, en relación con el artículo 25, párrafos primero, fracción LVI, último, numeral 8 del Reglamento Interior del Servicio de Administración Tributaria, publicado en el Diario Oficial de la Federación el 24 de agosto de 2015, vigente a partir del 22 de noviembre de 2015 de conformidad con lo dispuesto en el primer párrafo del artículo Primero Transitorio de dicho Reglamento; y reformado mediante Decreto por el que se reforman y adicionan diversas disposiciones del Reglamento Interior de la Secretaría de Hacienda y Crédito Público y del Reglamento Interior del Servicio de Administración Tributaria, y por el que se expide el Reglamento Interior de la Agencia Nacional de Aduanas de México publicado en el mismo órgano oficial el 21 de diciembre de 2021, vigente a partir del 01 de enero de 2022, de conformidad con lo dispuesto en el Artículo Primero Transitorio de dicho Decreto; así como en el artículo 144, párrafo primero, fracciones II, III, IV, VII, X y XXXIX y 155 de la Ley Aduanera, 33, último párrafo y 63 del Código Fiscal de la Federación, procede a determinar su situación fiscal en materia de comercio exterior de conformidad a lo siguiente:";
    // docBody.insertParagraph(proemio, "Start");
    docBody.insertText(proemio, "Start");

    return context.sync();
  }).catch(function (error) {
    // eslint-disable-next-line no-undef
    console.log("Error: " + error);
    // eslint-disable-next-line no-undef
    if (error instanceof OfficeExtension.Error) {
      // eslint-disable-next-line no-undef
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function replaceContentInControl() {
  Word.run(function (context) {
    for (let index = 1; index <= count; index++) {
    var serviceNameContentControl = context.document.contentControls.getByTag("tag"+index).getFirst();
    serviceNameContentControl.insertText(texts_area[index-1], "Replace");
    }
    //var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    //serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");///

    return context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

count = 0;
texts_area = [];


function createContentControl() {
  Word.run(function (context) {
    // eslint-disable-next-line no-undef
    count++;
    let etiqueta_txt = document.getElementById("text_area").value;
    texts_area.push(etiqueta_txt);
    var serviceNameRange = context.document.getSelection();
    
    var serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Tag";
    serviceNameContentControl.tag = "tag"+count;
    serviceNameContentControl.appearance = "None";
    serviceNameContentControl.color = "gray";

      return context.sync();
  }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function resetInternalCount() {
texts_area = []; 
  count = 0;
}