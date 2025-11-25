const SPREADSHEET_ID = '1-VvTkfdM2-N9Pf8i04yt2j1Aoz5V1mmF-DVYvQTaqEU'; // ton ID ici

function doPost(e) {
  try {
    Logger.log('Event brut: ' + JSON.stringify(e));

    const jsonStr = e.parameter.payload || (e.postData && e.postData.contents);
    if (!jsonStr) {
      throw new Error('Aucune donnée reçue (payload vide)');
    }

    Logger.log('JSON reçu: ' + jsonStr);

    const obj = JSON.parse(jsonStr);
    const charts = obj.charts || [];

    if (!Array.isArray(charts) || charts.length === 0) {
      throw new Error('Aucun graphique dans "charts".');
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const maxCharts = 14;
    const chartsToCreate = charts.slice(0, maxCharts);

    let createdCount = 0;

    // On stocke ici les EmbeddedChart (ceux *dans* le sheet)
    const embeddedChartsInfo = []; // { index, title, sheet, chart }

    chartsToCreate.forEach((chartObj, index) => {
      const title = chartObj.title || ('Graphique ' + (index + 1));
      const data = chartObj.data || [];

      if (!Array.isArray(data) || data.length === 0) {
        Logger.log('Graphique ' + (index + 1) + ' ignoré (pas de données).');
        return;
      }

      const sheetName = 'ChartData_' + (index + 1);
      let sheet = ss.getSheetByName(sheetName);

      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      } else {
        sheet.clear();
      }

      // En-têtes
      sheet.getRange(1, 1).setValue('Label');
      sheet.getRange(1, 2).setValue('Valeur');

      // Données
      const values = data.map(row => [row.label, row.value]);
      sheet.getRange(2, 1, values.length, 2).setValues(values);

      // Supprimer anciens graphiques dans cette feuille
      const chartsInSheetBefore = sheet.getCharts();
      chartsInSheetBefore.forEach(c => sheet.removeChart(c));

      // Créer le graphique en colonnes
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const chartBuilder = sheet.newChart()
          .asColumnChart()
          .addRange(sheet.getRange(1, 1, lastRow, 2))
          .setPosition(1, 4, 0, 0)
          .setOption('title', title);

        const builtChart = chartBuilder.build();
        sheet.insertChart(builtChart);
        createdCount++;

        // IMPORTANT : on relit le chart depuis la feuille (celui vraiment inséré)
        const chartsInSheetAfter = sheet.getCharts();
        const insertedChart = chartsInSheetAfter[chartsInSheetAfter.length - 1];

        embeddedChartsInfo.push({
          index: index + 1,
          title: title,
          sheet: sheet,
          chart: insertedChart
        });
      }
    });

    // === Création du Google Slides et insertion des graphiques linkés ===
    let presentationUrl = null;

    if (embeddedChartsInfo.length > 0) {
      const now = new Date();
      const presTitle = 'Rapport Graphiques - ' + now.toLocaleString();
      const presentation = SlidesApp.create(presTitle);
      presentationUrl = presentation.getUrl();
      Logger.log('Présentation créée: ' + presentationUrl);

      const slides = presentation.getSlides();

      embeddedChartsInfo.forEach((info, idx) => {
        // première slide = celle créée par défaut
        const slide = (idx === 0)
          ? slides[0]
          : presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

        // Insertion du graphique Sheets LINKÉ
        const sheetsChart = slide.insertSheetsChart(info.chart);

        // Redimensionnement / position
        sheetsChart
          .setWidth(400)
          .setHeight(300)
          .setLeft(50)
          .setTop(80);

        // Titre de la slide
        const titleShape = slide.insertShape(
          SlidesApp.ShapeType.TEXT_BOX,
          50, 20, 400, 40
        );
        titleShape.getText().setText(info.title);
        titleShape.getText().getTextStyle().setBold(true).setFontSize(18);
      });
    } else {
      Logger.log('Aucun graphique créé dans les sheets, donc pas de Slides.');
    }

    const output = {
      status: 'ok',
      message: 'Graphiques créés dans Sheets : ' + createdCount,
      requested: charts.length,
      maxHandled: maxCharts,
      slidesUrl: presentationUrl
    };

    return ContentService
      .createTextOutput(JSON.stringify(output))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Erreur: ' + err);

    const output = {
      status: 'error',
      message: err.toString()
    };

    return ContentService
      .createTextOutput(JSON.stringify(output))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
