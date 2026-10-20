// ==========================================
// BACKEND: Code.gs - Siomay Langkat Order (Revised with Dynamic Social Links)
// ==========================================

function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Siomay Langkat Order')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (error) {
    return HtmlService.createHtmlOutput(
      '<h2 style="font-family:sans-serif; text-align:center; margin-top:50px; color:#e74c3c;">' +
      '⚠️ Error: File "index.html" belum dibuat atau namanya salah.<br><br>' +
      '<span style="font-size:16px; color:#555;">Pastikan file HTML diberi nama persis <b>index</b> (huruf kecil semua).</span>' +
      '</h2>'
    );
  }
}

// Otomatis membuat database sheet & header termasuk SocialMedia
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const targetSheets = [
    { name: 'Menu', headers: ['ID_Menu', 'Nama_Menu', 'Kategori', 'Harga', 'Gambar_URL'] },
    { name: 'Orders', headers: ['ID_Order', 'Tanggal', 'Meja', 'Pesanan', 'Total', 'Status', 'Bukti_Bayar'] },
    { name: 'Admin', headers: ['Username', 'Password'] },
    { name: 'KritikSaran', headers: ['ID', 'Tanggal', 'Nama', 'Kritik', 'Saran'] },
    { name: 'SocialMedia', headers: ['Platform', 'Link'] }
  ];

  targetSheets.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);
      if (s.name === 'Admin') {
        sheet.appendRow(['admin', 'admin123']);
      }
      if (s.name === 'SocialMedia') {
        sheet.appendRow(['WhatsApp', 'https://wa.me/6281234567890']);
        sheet.appendRow(['Instagram', 'https://instagram.com']);
        sheet.appendRow(['TikTok', 'https://tiktok.com']);
      }
    } else {
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(s.headers);
        if (s.name === 'SocialMedia') {
          sheet.appendRow(['WhatsApp', 'https://wa.me/6281234567890']);
          sheet.appendRow(['Instagram', 'https://instagram.com']);
          sheet.appendRow(['TikTok', 'https://tiktok.com']);
        }
      }
    }
  });
  
  return "Setup Database Berhasil! Semua sheet termasuk SocialMedia telah siap.";
}

function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data.shift();
  
  return data.map(row => {
    let obj = {};
    headers.forEach((header, i) => obj[header] = row[i]);
    return obj;
  });
}

// --- MENU ---
function getMenu() {
  return getSheetData('Menu');
}

function addMenu(menuData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Menu');
  const id = "M" + new Date().getTime();
  sheet.appendRow([id, menuData.nama, menuData.kategori, menuData.harga, menuData.gambar]);
  return "Menu berhasil ditambahkan!";
}

function deleteMenu(idMenu) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Menu');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == idMenu) {
      sheet.deleteRow(i + 1);
      return "Menu berhasil dihapus!";
    }
  }
  return "Menu tidak ditemukan.";
}

// --- ORDERS ---
function getOrders() {
  return getSheetData('Orders');
}

function addOrder(orderData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Orders');
    const id = "ORD" + new Date().getTime();
    
    const date = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
    const pesananString = orderData.items.map(item => `${item.nama} (${item.qty}x)`).join(", ");
    
    sheet.appendRow([id, date, orderData.meja, pesananString, orderData.total, "Belum Bayar", ""]);
    return "Pesanan berhasil dibuat! Silakan tunggu di meja Anda.";
  } catch(e) {
    return "Error: " + e.message;
  }
}

function updateOrderStatus(idOrder, status, buktiBayar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Orders');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == idOrder) {
      sheet.getRange(i + 1, 6).setValue(status);
      if (buktiBayar !== undefined && buktiBayar !== null) {
        sheet.getRange(i + 1, 7).setValue(buktiBayar);
      }
      return "Status pesanan & bukti bayar diperbarui!";
    }
  }
  return "Pesanan tidak ditemukan.";
}

// --- KRITIK & SARAN ---
function getKritikSaran() {
  return getSheetData('KritikSaran');
}

function addKritikSaran(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('KritikSaran');
  if (!sheet) {
    sheet = ss.insertSheet('KritikSaran');
    sheet.appendRow(['ID', 'Tanggal', 'Nama', 'Kritik', 'Saran']);
  }
  const id = "KS" + new Date().getTime();
  const date = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
  sheet.appendRow([id, date, data.nama || 'Anonim', data.kritik || '-', data.saran || '-']);
  return "Terima kasih atas kritik & saran Anda!";
}

// --- SOCIAL MEDIA SETTINGS ---
function getSocialMedia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('SocialMedia');
  if (!sheet) {
    sheet = ss.insertSheet('SocialMedia');
    sheet.appendRow(['Platform', 'Link']);
    sheet.appendRow(['WhatsApp', 'https://wa.me/6281234567890']);
    sheet.appendRow(['Instagram', 'https://instagram.com']);
    sheet.appendRow(['TikTok', 'https://tiktok.com']);
  }
  return getSheetData('SocialMedia');
}

function updateSocialMedia(platform, link) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SocialMedia');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(platform).trim().toLowerCase()) {
      sheet.getRange(i + 1, 2).setValue(link);
      return `Link ${platform} berhasil diperbarui!`;
    }
  }
  sheet.appendRow([platform, link]);
  return `Link ${platform} berhasil disimpan!`;
}

// --- ADMIN AUTH ---
function loginAdmin(username, password) {
  const admins = getSheetData('Admin');
  const user = admins.find(a => String(a.Username).trim() === String(username).trim() && String(a.Password).trim() === String(password).trim());
  if (user) {
    return { success: true, message: "Login berhasil!" };
  } else {
    return { success: false, message: "Username atau Password salah." };
  }
}
