Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("translateBtn").onclick = handleTranslate;
  }
});

// Helper HTTP Request menggunakan XMLHttpRequest murni (Kompatibel IE11 / Excel 2016)
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
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load(["values", "rowCount", "columnCount"]);

    return context.sync().then(function () {
      var originalValues = selectedRange.values;
      var updatedValues = JSON.parse(JSON.stringify(originalValues));
      var cellsToUpdate = [];

      // 1. Kumpulkan semua sel yang berisi teks
      for (var r = 0; r < selectedRange.rowCount; r++) {
        for (var c = 0; c < selectedRange.columnCount; c++) {
          var cellVal = originalValues[r][c];
          if (cellVal !== null && cellVal !== undefined && cleanTrim(cellVal) !== "") {
            cellsToUpdate.push({
              row: r,
              col: c,
              text: cleanTrim(cellVal)
            });
          }
        }
      }

      // Cek apakah ada sel yang diproses
      if (cellsToUpdate.length === 0) {
        statusDiv.className = "error";
        statusDiv.innerText = "Pilih sel yang berisi teks terlebih dahulu!";
        return;
      }

      var totalProcessed = 0;

      // 2. Proses terjemahan sel satu per satu secara berurutan
      function processNextCell(index) {
        if (index >= cellsToUpdate.length) {
          // Setelah semua sel selesai diterjemahkan, tulis kembali ke Excel
          selectedRange.values = updatedValues;
          if (position === "newline") {
            selectedRange.format.wrapText = true;
          }

          return context.sync().then(function () {
            statusDiv.className = "success";
            statusDiv.innerText = "✓ Berhasil memperbarui " + totalProcessed + " sel!";
          });
        }

        var item = cellsToUpdate[index];
        fetchTranslation(item.text, function (err, translated) {
          if (!err && translated) {
            var cleanTrans = cleanTrim(translated);
            if (position === "newline") {
              updatedValues[item.row][item.col] = item.text + "\n" + cleanTrans;
            } else {
              updatedValues[item.row][item.col] = item.text + " " + cleanTrans;
            }
            totalProcessed++;
          }
          processNextCell(index + 1);
        });
      }

      processNextCell(0);
    });
  }).catch(function (error) {
    console.log(error);
    statusDiv.className = "error";
    statusDiv.innerText = "Gagal: " + (error.message || error);
  });
}
