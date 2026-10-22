Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("translateBtn").onclick = handleTranslate;
  }
});

// Helper HTTP Request menggunakan XMLHttpRequest murni (Universal)
function fetchTranslation(text, callback) {
  var url = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=id&tl=en&dt=t&q=" + encodeURIComponent(text);
  var xhr = new XMLHttpRequest();

  xhr.open("GET", url, true);
  xhr.onreadystatechange = function () {
    if (xhr.readyState === 4) {
      if (xhr.status === 200) {
        try {
          var data = JSON.parse(xhr.responseText);
          if (data && data[0]) {
            var translatedText = "";
            for (var i = 0; i < data[0].length; i++) {
              if (data[0][i][0]) {
                translatedText += data[0][i][0];
              }
            }
            callback(null, translatedText);
          } else {
            callback("Gagal membaca hasil terjemahan.", null);
          }
        } catch (e) {
          callback("Format respon tidak valid.", null);
        }
      } else {
        callback("Gagal terhubung ke API (Status: " + xhr.status + ")", null);
      }
    }
  };
  xhr.send();
}

// Trim teks tanpa bergantung pada String.prototype.trim (ES5 Polyfill)
function cleanTrim(str) {
  return String(str).replace(/^\s+|\s+$/g, '');
}

function handleTranslate() {
  var statusDiv = document.getElementById("status");
  statusDiv.className = "";
  statusDiv.innerText = "⏳ Sedang menerjemahkan...";

  var radios = document.getElementsByName("position");
  var position = "inline";
  for (var i = 0; i < radios.length; i++) {
    if (radios[i].checked) {
      position = radios[i].value;
      break;
    }
  }

  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var selectedRange = context.workbook.getSelectedRange();
    var shapes = sheet.shapes;

    selectedRange.load(["values", "rowCount", "columnCount"]);
    shapes.load("items");

    return context.sync().then(function () {
      var originalValues = selectedRange.values;
      var updatedValues = JSON.parse(JSON.stringify(originalValues));
      var cellsToUpdate = [];
      var totalProcessed = 0;

      // 1. Kumpulkan semua sel yang berisi teks
      for (var r = 0; r < selectedRange.rowCount; r++) {
        for (var c = 0; c < selectedRange.columnCount; c++) {
          var cellVal = originalValues[r][c];
          if (cellVal !== null && cellVal !== undefined && cleanTrim(cellVal) !== "") {
            cellsToUpdate.push({
              type: "cell",
              row: r,
              col: c,
              text: cleanTrim(cellVal)
            });
          }
        }
      }

      // 2. Kumpulkan Shape yang memiliki teks
      for (var s = 0; s < shapes.items.length; s++) {
        var shp = shapes.items[s];
        if (shp.textFrame) {
          shp.textFrame.load(["hasText", "textRange"]);
        }
      }

      return context.sync().then(function () {
        var shapeItems = [];
        for (var k = 0; k < shapes.items.length; k++) {
          var itemShape = shapes.items[k];
          if (itemShape.textFrame && itemShape.textFrame.hasText) {
            var rawText = itemShape.textFrame.textRange.text;
            if (rawText && cleanTrim(rawText) !== "") {
              shapeItems.push({
                type: "shape",
                shapeObj: itemShape,
                text: cleanTrim(rawText)
              });
            }
          }
        }

        var allItems = cellsToUpdate.concat(shapeItems);

        if (allItems.length === 0) {
          statusDiv.className = "error";
          statusDiv.innerText = "Pilih sel atau shape yang berisi teks!";
          return;
        }

        // 3. Proses terjemahan item satu per satu
        function processNext(index) {
          if (index >= allItems.length) {
            // Tulis kembali hasil ke Range Sel
            selectedRange.values = updatedValues;
            if (position === "newline") {
              selectedRange.format.wrapText = true;
            }

            return context.sync().then(function () {
              statusDiv.className = "success";
              statusDiv.innerText = "✓ Berhasil memperbarui " + totalProcessed + " elemen!";
            });
          }

          var current = allItems[index];
          fetchTranslation(current.text, function (err, translated) {
            if (!err && translated) {
              var cleanTrans = cleanTrim(translated);
              var result = "";

              if (position === "newline") {
                result = current.text + "\n" + cleanTrans;
              } else {
                result = current.text + " " + cleanTrans;
              }

              if (current.type === "cell") {
                updatedValues[current.row][current.col] = result;
              } else if (current.type === "shape") {
                current.shapeObj.textFrame.textRange.text = result;
              }
              totalProcessed++;
            }
            processNext(index + 1);
          });
        }

        processNext(0);
      });
    });
  }).catch(function (error) {
    console.log(error);
    statusDiv.className = "error";
    statusDiv.innerText = "Gagal: " + (error.message || error);
  });
}