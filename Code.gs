function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Driver Portal - KPK')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getEmployeeData(targetId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('Driver_detail');
  const jadwalSheet = ss.getSheetByName('Driver_sch');
  
  const masterData = masterSheet.getDataRange().getValues();
  const jadwalData = jadwalSheet.getDataRange().getValues();
  const cleanTargetId = targetId.toString().trim();

  let result = { 
    found: false, 
    biodata: {}, 
    listJadwal: [], 
    rules: [], 
    monthName: "",
    firstDayOffset: 0,
    linkAbsen: "https://docs.google.com/forms/d/e/1FAIpQLSdeld9zz__PtknsB7OfXrQQn567ei5lFFx8OIvo8TBwVq8ktw/viewform"
  };

  let filteredJadwal = jadwalData.filter(row => row[0].toString().trim() === cleanTargetId);
  if (filteredJadwal.length === 0) return { found: false, msg: "SAP ID salah atau tidak ditemukan" };

  let userFoundInMaster = false;
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][0].toString().trim() === cleanTargetId) {
      result.found = true;
      userFoundInMaster = true;

      const rawStatus = masterData[i][2] ? masterData[i][2].toString().toUpperCase() : "";
      
      result.biodata = { 
        sapId: masterData[i][0],
        nama: masterData[i][1], 
        status: rawStatus, 
        telp: masterData[i][3] 
      };

      if (rawStatus.includes("FT") || rawStatus.includes("FULL TIME") || rawStatus.includes("FULLTIME")) {
          result.rules.push("FT 6 Libur sesuai jadwal → tidak perlu isi link izin & tidak perlu kembalikan unit & kunci (KHUSUS BCH)");
          result.rules.push("FT 5 Libur sesuai jadwal → tidak perlu isi link izin & wajib kembalikan unit & kunci");
        // Rules umum Fulltime
        result.rules.push("Libur 1 hari di luar jadwal → wajib isi link izin dan masuk di hari libur pengganti");
        result.rules.push("Libur > 1 hari berturut-turut → wajib kembalikan unit & kunci (baterai ≥ 95%)");
      } 
      else if (rawStatus.includes("PT") || rawStatus.includes("PART TIME") || rawStatus.includes("PARTTIME")) {
        result.rules.push("PT → Wajib mengembalikan unit & kunci setiap hari (baterai ≥ 95%)");
      }
      break;
    }
  }

  if (!userFoundInMaster) return { found: false, msg: "SAP ID Anda sedang di Suspend, segera melapor ke Depot / OS" };

  // 3. PROSES DATA KALENDER
  const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  let sampleDate = new Date(filteredJadwal[0][1]); 
  result.monthName = months[sampleDate.getMonth()] + " " + sampleDate.getFullYear();
  let firstDate = new Date(sampleDate.getFullYear(), sampleDate.getMonth(), 1);
  result.firstDayOffset = firstDate.getDay(); 

  result.listJadwal = filteredJadwal.map(row => {
    let tglAsli = new Date(row[1]);
    return {
      isoDate: Utilities.formatDate(tglAsli, "GMT+7", "yyyy-MM-dd"),
      tglAngka: Utilities.formatDate(tglAsli, "GMT+7", "d"),
      shift: row[2] ? row[2].toString().toUpperCase() : "OFF", 
      keterangan: row[3] ? row[3].toString().toUpperCase() : "" 
    };
  });

  return result;
}

function getInspectionData(plateNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const insheet = ss.getSheetByName('Inspection_Schedule');
  if (!insheet) return { found: false, msg: "Sheet Inspeksi tidak ditemukan" };

  const data = insheet.getDataRange().getValues();
  const targetPlate = plateNumber.toString().toUpperCase().replace(/\s/g, "");
  
  let result = { found: false, data: {} };

  for (let i = 1; i < data.length; i++) {
    let currentPlate = data[i][0].toString().toUpperCase().replace(/\s/g, "");
    
    if (currentPlate === targetPlate) {
      result.found = true;
      result.data = {
        plate: data[i][0],      // Plate Number (Col A)
        hull: data[i][1],       // Hull Number (Col B)
        team: data[i][3],       // Team (Col G)
        day: data[i][2]         // Day (Col H)
      };
      break;
    }
  }
  return result;
}

function getDriversList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('Driver_detail');
  const masterData = masterSheet.getDataRange().getValues();

  let drivers = [];
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][0]) { // SAP ID ada
      drivers.push({
        id: masterData[i][0].toString().trim(),
        nama: masterData[i][1] ? masterData[i][1].toString().trim() : "",
        status: masterData[i][2] ? masterData[i][2].toString().trim() : "",
        telp: masterData[i][3] ? masterData[i][3].toString().trim() : ""
      });
    }
  }
  return drivers;
}
// ===== TAMBAHAN: tempel di PALING BAWAH Code.gs (kode lama tidak diubah) =====

// Daftar nopol + nomor lambung dari sheet Inspection_Schedule untuk dropdown pencarian
function getPlatesList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const insheet = ss.getSheetByName('Inspection_Schedule');
  if (!insheet) return [];

  const data = insheet.getDataRange().getValues();
  const plates = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {                       // Plate Number (Col A)
      plates.push({
        plate: data[i][0].toString().trim(),
        hull: data[i][1] ? data[i][1].toString().trim() : ""   // Hull Number (Col B)
      });
    }
  }
  return plates;
}