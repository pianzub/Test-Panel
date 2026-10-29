/**
 * FOOD TEST PANEL APP
 * Backend Logic - Version: 3.8.0
 * Performance update: smaller payloads, indexed statistics, cache, batch writes,
 * UUID identifiers, retry-safe submissions, and duplicate-response protection.
 */

const FOOD_TEST_CACHE_TTL = 120;
const FOOD_TEST_CACHE_KEYS = {
  ADMIN_MASTER: 'FOOD_TEST_ADMIN_MASTER_V2',
  ACTIVE_PROJECTS: 'FOOD_TEST_ACTIVE_PROJECTS_V2'
};

function doGet(e) {
  // Health check untuk memastikan URL /exec benar-benar hidup.
  // Buka .../exec?health=1 di browser; hasil yang benar adalah JSON.
  if (e && e.parameter && String(e.parameter.health) === '1') {
    return jsonResponse_({
      ok: true,
      service: 'Food Test Panel API',
      status: 'online',
      timestamp: new Date().toISOString()
    });
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Food Test Panel Ultimate')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/*
 * Endpoint API untuk frontend GitHub Pages.
 * Payload dikirim sebagai text/plain agar request browser tetap sederhana.
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Request body kosong.');
    }

    const request = JSON.parse(e.postData.contents);
    const action = String(request.action || '');
    const args = Array.isArray(request.args) ? request.args : [];
    const authToken = String(request.authToken || '');

    const handlers = {
      loginAdmin: loginAdmin,
      createProject: createProject,
      editProject: editProject,
      addMenu: addMenu,
      editMenu: editMenu,
      reorderMenus: reorderMenus,
      getDashboardData: getDashboardData,
      getAdminMasterData: getAdminMasterData,
      getProjectExportData: getProjectExportData,
      getMenuNpsDetails: getMenuNpsDetails,
      calculateDetailedStats: calculateDetailedStats,
      generatePDFReport: generatePDFReport,
      getActiveProjects: getActiveProjects,
      getMenusForProject: getMenusForProject,
      submitResponse: submitResponse
    };

    if (!Object.prototype.hasOwnProperty.call(handlers, action)) {
      throw new Error('Action tidak diizinkan: ' + action);
    }

    const adminActions = {
      getDashboardData: true,
      getAdminMasterData: true,
      getProjectExportData: true,
      getMenuNpsDetails: true,
      calculateDetailedStats: true,
      generatePDFReport: true,
      createProject: true,
      editProject: true,
      addMenu: true,
      editMenu: true,
      reorderMenus: true
    };
    const superAdminActions = {
      createProject: true,
      editProject: true,
      addMenu: true,
      editMenu: true,
      reorderMenus: true
    };

    if (adminActions[action]) {
      const session = validateApiSession_(authToken);
      if (!session) throw new Error('Sesi admin tidak valid atau sudah berakhir. Silakan login kembali.');
      if (superAdminActions[action] && session.role === 'Admin') {
        throw new Error('Akses ditolak. Fungsi ini hanya tersedia untuk Super Admin.');
      }
    }

    const result = handlers[action].apply(null, args);
    return jsonResponse_({
      ok: true,
      result: result,
      requestId: String(request.requestId || (args[0] && args[0].requestId) || '')
    });
  } catch (error) {
    return jsonResponse_({
      ok: false,
      error: error && error.message ? error.message : String(error),
      requestId: (typeof request !== 'undefined' && request)
        ? String(request.requestId || (request.args && request.args[0] && request.args[0].requestId) || '')
        : ''
    });
  }
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function createApiSession_(name, role) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  const session = JSON.stringify({ name: String(name || ''), role: String(role || 'Admin') });
  CacheService.getScriptCache().put('FOOD_TEST_SESSION_' + token, session, 21600);
  return token;
}

function validateApiSession_(token) {
  if (!token) return null;
  const value = CacheService.getScriptCache().get('FOOD_TEST_SESSION_' + token);
  if (!value) return null;
  try {
    return JSON.parse(value);
  } catch (error) {
    return null;
  }
}

function getCachedJson_(key) {
  try {
    const value = CacheService.getScriptCache().get(key);
    return value ? JSON.parse(value) : null;
  } catch (error) {
    return null;
  }
}

function putCachedJson_(key, value, ttl) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), ttl || FOOD_TEST_CACHE_TTL);
  } catch (error) {
    // Cache is an optimization only. Large payloads or temporary cache failures
    // must never block the application.
  }
}

function invalidateMasterCache_() {
  CacheService.getScriptCache().removeAll([
    FOOD_TEST_CACHE_KEYS.ADMIN_MASTER,
    FOOD_TEST_CACHE_KEYS.ACTIVE_PROJECTS
  ]);
}

function invalidateProjectCache_(projectId) {
  const id = String(projectId || '');
  CacheService.getScriptCache().removeAll([
    'FOOD_TEST_MENUS_' + id,
    'FOOD_TEST_STATS_' + id + '_0',
    'FOOD_TEST_STATS_' + id + '_1'
  ]);
}

function getSheetRows_(sheetName, displayValues) {
  const sheet = getSheetOrSetup(sheetName);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) return [];
  const range = sheet.getRange(1, 1, lastRow, lastColumn);
  return displayValues ? range.getDisplayValues() : range.getValues();
}

function ensureMenuOrderColumn_() {
  const sheet = getSheetOrSetup('Menus');
  const headerCell = sheet.getRange(1, 8);
  const currentHeader = String(headerCell.getValue() || '').trim();
  if (!currentHeader) headerCell.setValue('Urutan');
  else if (currentHeader !== 'Urutan') throw new Error('Kolom H sheet Menus harus digunakan untuk Urutan.');
  return sheet;
}

function getNextMenuOrder_(sheet, projectId) {
  const data = sheet.getDataRange().getValues();
  let maxOrder = 0;
  let projectMenuCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) !== String(projectId)) continue;
    projectMenuCount++;
    const order = Number(data[i][7]);
    if (Number.isFinite(order) && order > maxOrder) maxOrder = order;
  }
  return Math.max(maxOrder, projectMenuCount) + 1;
}

// --- DATABASE HELPER ---
function getSheetOrSetup(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    initialSetup(); 
    sheet = ss.getSheetByName(sheetName);
  }
  return sheet;
}

function initialSetup() {
  try {
    const root = DriveApp.getRootFolder(); 
  } catch (e) {
    console.error("Izin Drive belum lengkap: " + e.toString());
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ss.getSheetByName('Admins')) {
    const s = ss.insertSheet('Admins');
    s.appendRow(['NIK', 'Password', 'Nama', 'Role']); 
    s.appendRow(['12345', 'admin', 'Super Admin', 'Super Admin']); 
  }
  if (!ss.getSheetByName('Projects')) {
    const s = ss.insertSheet('Projects');
    s.appendRow(['ID', 'Nama_Project', 'Tanggal', 'Status']);
  }
  if (!ss.getSheetByName('Menus')) {
    const s = ss.insertSheet('Menus');
    s.appendRow(['ID', 'Project_ID', 'Nama_Menu', 'Deskripsi', 'Foto_ID', 'Pertanyaan_JSON', 'Status', 'Urutan']);
  }
  if (!ss.getSheetByName('Responses')) {
    const s = ss.insertSheet('Responses');
    s.appendRow(['ID', 'Project_ID', 'Menu_ID', 'Nama_Panelis', 'Jawaban_JSON', 'Avg_Raw_Score', 'Timestamp', 'Panelist_ID']);
  }
}

function getImagesFolder() {
  const SPECIFIC_FOLDER_ID = '1j3RVdhYTROAo_xbJZl8pvxNSp8XZHfPb'; // Ganti jika perlu
  try {
    return DriveApp.getFolderById(SPECIFIC_FOLDER_ID);
  } catch (e) {
    const FOLDER_NAME = "FoodTestPanel_Images";
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    return folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
  }
}

// --- FUNGSI ADMIN ---

function loginAdmin(nik, password) {
  const sheet = getSheetOrSetup('Admins'); 
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) == String(nik) && String(data[i][1]) == String(password)) {
      const role = data[i][3] ? String(data[i][3]) : 'Admin';
      const name = data[i][2];
      return { success: true, name: name, role: role, token: createApiSession_(name, role) };
    }
  }
  return { success: false };
}

function createProject(namaProject) {
  const sheet = getSheetOrSetup('Projects');
  const id = 'PROJ-' + Utilities.getUuid();
  const now = new Date();
  const dateStr = [now.getFullYear(), (now.getMonth() + 1).toString().padStart(2,'0'), now.getDate().toString().padStart(2,'0')].join('-');
  sheet.appendRow([id, namaProject, dateStr, 'Active']);
  invalidateMasterCache_();
  return { success: true, id: id };
}

function editProject(id, newName, newStatus) {
  const sheet = getSheetOrSetup('Projects');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 2).setValue(newName);
      sheet.getRange(i + 1, 4).setValue(newStatus);
      invalidateMasterCache_();
      return { success: true };
    }
  }
  return { success: false, message: "ID tidak ditemukan" };
}

function addMenu(projectId, namaMenu, deskripsi, fotoData, pertanyaanArray, status) {
  const sheet = ensureMenuOrderColumn_();
  const id = 'MENU-' + Utilities.getUuid();
  let fotoId = '';
  
  const menuStatus = status || 'Active';

  if (fotoData && fotoData.data) {
    try {
      const folder = getImagesFolder();
      const blob = Utilities.newBlob(Utilities.base64Decode(fotoData.data), fotoData.mimeType, `menu_${id}.jpg`);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fotoId = file.getId();
    } catch (e) {
      return { success: false, message: "Gagal Upload: " + e.message };
    }
  }

  const menuOrder = getNextMenuOrder_(sheet, projectId);
  sheet.appendRow([id, projectId, namaMenu, deskripsi, fotoId, JSON.stringify(pertanyaanArray), menuStatus, menuOrder]);
  invalidateMasterCache_();
  invalidateProjectCache_(projectId);
  return { success: true };
}

function editMenu(menuId, projectId, namaMenu, deskripsi, fotoData, pertanyaanArray, status) {
  const sheet = ensureMenuOrderColumn_();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(menuId)) {
      const row = i + 1;
      const previousProjectId = String(data[i][1] || '');
      let fotoId = String(data[i][4] || '');
      const menuStatus = status || 'Active';
      let menuOrder = Number(data[i][7]) > 0 ? Number(data[i][7]) : i;
      if (previousProjectId !== String(projectId)) menuOrder = getNextMenuOrder_(sheet, projectId);

      if (fotoData && fotoData.data) {
        try {
          const folder = getImagesFolder();
          const blob = Utilities.newBlob(Utilities.base64Decode(fotoData.data), fotoData.mimeType, `menu_${menuId}_v2.jpg`);
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          fotoId = file.getId();
        } catch (e) { return { success: false, message: "Gagal update foto: " + e.message }; }
      }

      sheet.getRange(row, 2, 1, 7).setValues([[
        projectId,
        namaMenu,
        deskripsi,
        fotoId,
        JSON.stringify(pertanyaanArray),
        menuStatus,
        menuOrder
      ]]);
      invalidateMasterCache_();
      invalidateProjectCache_(previousProjectId);
      invalidateProjectCache_(projectId);
      return { success: true };
    }
  }
  return { success: false, message: "Menu tidak ditemukan" };
}

function reorderMenus(projectId, orderedMenuIds) {
  if (!projectId || !Array.isArray(orderedMenuIds) || !orderedMenuIds.length) {
    throw new Error('Urutan menu tidak valid.');
  }
  const ids = orderedMenuIds.map(id => String(id));
  if (new Set(ids).size !== ids.length) throw new Error('Terdapat ID menu duplikat pada urutan.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = ensureMenuOrderColumn_();
    const data = sheet.getDataRange().getValues();
    const projectRows = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][1]) === String(projectId)) {
        projectRows.push({ row: i + 1, id: String(data[i][0]) });
      }
    }

    const expectedIds = new Set(projectRows.map(item => item.id));
    if (ids.length !== projectRows.length || ids.some(id => !expectedIds.has(id))) {
      throw new Error('Daftar menu berubah. Refresh database lalu coba kembali.');
    }

    const orderById = new Map(ids.map((id, index) => [id, index + 1]));
    const orderValues = data.slice(1).map(row => [row[7] || '']);
    projectRows.forEach(item => {
      orderValues[item.row - 2][0] = orderById.get(item.id);
    });
    if (orderValues.length) sheet.getRange(2, 8, orderValues.length, 1).setValues(orderValues);

    invalidateMasterCache_();
    invalidateProjectCache_(projectId);
    return { success: true, count: ids.length };
  } finally {
    lock.releaseLock();
  }
}

function getDashboardData() {
  const projSheet = getSheetOrSetup('Projects');
  const menuSheet = getSheetOrSetup('Menus');
  const respSheet = getSheetOrSetup('Responses'); 
  
  const projData = projSheet.getDataRange().getDisplayValues(); 
  const menuRaw = menuSheet.getDataRange().getDisplayValues();
  const respRaw = respSheet.getDataRange().getDisplayValues();

  const projects = [];
  if (projData.length > 1) {
    for (let i = 1; i < projData.length; i++) {
      const r = projData[i];
      if (r[0] && r[0].toString().trim() !== "") {
        projects.push({
          id: String(r[0]),
          name: String(r[1]),
          date: String(r[2]), 
          status: String(r[3] || 'Active')
        });
      }
    }
  }

  const menus = [];
  if (menuRaw.length > 1) {
    for (let i = 1; i < menuRaw.length; i++) {
      const r = menuRaw[i];
      if (r[0]) {
        const st = r[6] ? String(r[6]) : 'Active';
        menus.push({ 
          id: String(r[0]), 
          projectId: String(r[1]), 
          name: String(r[2]),
          desc: String(r[3]),
          fotoId: String(r[4]), 
          questions: String(r[5]),
          status: st,
          order: Number(r[7]) > 0 ? Number(r[7]) : i
        });
      }
    }
  }

  menus.sort((a, b) => String(a.projectId) === String(b.projectId) ? a.order - b.order : 0);
  const rawResponses = respRaw.length > 1 ? respRaw.slice(1) : [];
  
  return { projects: projects, menus: menus, rawResponses: rawResponses };
}

// Lightweight payload for normal admin navigation. Responses are intentionally
// excluded and requested only for a selected project/export.
function getAdminMasterData() {
  const cached = getCachedJson_(FOOD_TEST_CACHE_KEYS.ADMIN_MASTER);
  if (cached) return cached;

  const projData = getSheetRows_('Projects', true);
  const menuRaw = getSheetRows_('Menus', true);
  const projects = projData.slice(1).filter(r => r[0]).map(r => ({
    id: String(r[0]),
    name: String(r[1]),
    date: String(r[2]),
    status: String(r[3] || 'Active')
  }));
  const menus = menuRaw.slice(1).filter(r => r[0]).map((r, index) => ({
    id: String(r[0]),
    projectId: String(r[1]),
    name: String(r[2]),
    desc: String(r[3]),
    fotoId: String(r[4]),
    questions: String(r[5]),
    status: String(r[6] || 'Active'),
    order: Number(r[7]) > 0 ? Number(r[7]) : index + 1
  }));
  menus.sort((a, b) => String(a.projectId) === String(b.projectId) ? a.order - b.order : 0);
  const result = { projects: projects, menus: menus, rawResponses: [] };
  putCachedJson_(FOOD_TEST_CACHE_KEYS.ADMIN_MASTER, result);
  return result;
}

function getProjectExportData(projectId, includeHidden) {
  const id = String(projectId || '').trim();
  const master = getAdminMasterData();
  const menus = master.menus.filter(m => {
    if (String(m.projectId) !== id) return false;
    return includeHidden || m.status !== 'Hidden';
  });
  const menuIds = new Set(menus.map(m => String(m.id)));
  const responseRows = getSheetRows_('Responses', true).slice(1).filter(r => {
    return String(r[1]).trim() === id && menuIds.has(String(r[2]));
  });
  return { projects: master.projects.filter(p => String(p.id) === id), menus: menus, rawResponses: responseRows };
}

function getMenuNpsDetails(menuId) {
  const responses = getSheetRows_('Responses', false).slice(1)
    .filter(r => String(r[2]) === String(menuId));
  const categoryStats = {};
  responses.forEach(row => {
    try {
      const answers = JSON.parse(row[4]);
      answers.forEach(answer => {
        if ((answer.type === 'rating' || (!answer.type && answer.answer && !isNaN(answer.answer) && Number(answer.answer) >= 1 && Number(answer.answer) <= 10)) && answer.answer !== '') {
          const label = answer.question || answer.label || 'Rating Umum';
          if (!categoryStats[label]) categoryStats[label] = { detractors: 0, passives: 0, promoters: 0, total: 0 };
          const value = Number(answer.answer);
          if (value >= 9) categoryStats[label].promoters++;
          else if (value >= 7) categoryStats[label].passives++;
          else categoryStats[label].detractors++;
          categoryStats[label].total++;
        }
      });
    } catch (error) {}
  });

  let totalQuestionNps = 0;
  let questionCount = 0;
  let promoters = 0;
  let passives = 0;
  let detractors = 0;
  Object.keys(categoryStats).forEach(label => {
    const stats = categoryStats[label];
    if (!stats.total) return;
    const nps = ((stats.promoters - stats.detractors) / stats.total) * 100;
    if (nps !== 0) {
      totalQuestionNps += nps;
      questionCount++;
    }
    promoters += stats.promoters;
    passives += stats.passives;
    detractors += stats.detractors;
  });
  const rawNps = questionCount ? totalQuestionNps / questionCount : 0;
  return {
    nps: rawNps >= 0 ? Math.ceil(rawNps) : Math.floor(rawNps),
    promoters: promoters,
    passives: passives,
    detractors: detractors,
    responseCount: responses.length
  };
}

function calculateDetailedStats(projectId, includeHidden) {
  const cacheKey = 'FOOD_TEST_STATS_' + String(projectId) + '_' + (includeHidden ? '1' : '0');
  const cached = getCachedJson_(cacheKey);
  if (cached) return cached;

  const menuData = getSheetRows_('Menus', false).slice(1);
  const respData = getSheetRows_('Responses', false).slice(1);

  // 1. Dapatkan daftar menu yang valid berdasarkan status (Active/Hidden)
  const projectMenus = menuData.filter(r => {
     const isProjectMatch = String(r[1]).trim() === String(projectId).trim();
     const status = r[6] ? String(r[6]) : 'Active';
     let isStatusMatch = true;
     if (!includeHidden) {
        if (status === 'Hidden') isStatusMatch = false;
     }
     return isProjectMatch && isStatusMatch;
  }).map(r => ({
    id: r[0], name: r[2]
  }));

  const projectMenuIds = new Set(projectMenus.map(menu => String(menu.id).trim()));
  const responsesByMenu = new Map();
  respData.forEach(row => {
    const menuId = String(row[2] || '').trim();
    if (!projectMenuIds.has(menuId)) return;
    if (!responsesByMenu.has(menuId)) responsesByMenu.set(menuId, []);
    responsesByMenu.get(menuId).push(row);
  });

  let rankings = [];
  let uniquePanelistsSet = new Set(); 
  let totalValidSubmissions = 0;      
  
  const stopWords = [
    'dan', 'yang', 'di', 'ini', 'itu', 'sangat', 'kurang', 'agak', 'cukup', 
    'terlalu', 'sedikit', 'banyak', 'rasa', 'menu', 'makan', 'saya', 'tidak', 
    'yg', 'gak', 'karena', 'kalau', 'pas', 'juga', 'nya', 'untuk', 'aja', 
    'bgt', 'banget', 'tapi', 'cuma', 'masih', 'bisa', 'ada', 'lagi', 'sudah', 
    'udah', 'dengan', 'dari', 'ke', 'saja', 'lebih', 'lumayan', 'seperti', 'kayak',
    'enak', 'mantap', 'oke', 'ok' 
  ];

  // 2. Loop hanya pada menu yang valid untuk menghitung Partisipan dan Statistik
  projectMenus.forEach(menu => {
    const responses = responsesByMenu.get(String(menu.id).trim()) || [];
    const totalResp = responses.length;
    
    // Tambahkan ke total submission valid
    totalValidSubmissions += totalResp;

    if (totalResp === 0) {
      rankings.push({ ...menu, nps: 0, count: 0, comments: [], keywords: [], avgPrice: 0, suggestedPrice: 0, summaryText: "Belum ada data untuk dianalisis.", rawScores: [], promoters: 0, passives: 0, detractors: 0 });
      return;
    }

    let categoryStats = {}; 
    let commentList = [];
    let allCommentText = "";
    let totalPrice = 0;
    let priceCount = 0;
    let rawScores = [];
    const panelistGroupCounts = {};

    responses.forEach(r => {
      // Masukkan Panelis ke dalam Unique Set (Hanya Panelis dari menu valid yang terhitung)
      if (r[7] && String(r[7]).trim() !== "") {
        uniquePanelistsSet.add(String(r[7])); 
      } else {
        uniquePanelistsSet.add(String(r[3]).trim().toLowerCase());
      }

      rawScores.push(parseFloat(r[5] || 0)); 
      const panelistGroup = String(r[3] || '').trim();
      if (panelistGroup) panelistGroupCounts[panelistGroup] = (panelistGroupCounts[panelistGroup] || 0) + 1;

      try {
        const answers = JSON.parse(r[4]); 
        if (Array.isArray(answers)) {
          answers.forEach(a => {
            if (a.type === 'text' && a.answer && String(a.answer).trim() !== "") {
              const cleanComment = String(a.answer).trim();
              commentList.push(cleanComment);
              allCommentText += " " + cleanComment.toLowerCase();
            }
            if (a.type === 'currency' && a.answer) {
              const val = parseInt(String(a.answer).replace(/[^0-9]/g, ''));
              if (!isNaN(val) && val > 0) {
                totalPrice += val;
                priceCount++;
              }
            }
            
            // FIX: Hanya menghitung jika tipe secara spesifik adalah 'rating'. 
            // Fallback (jika type tidak ada untuk data sangat lama) pastikan hanya hitung angka 1-10.
            if ((a.type === 'rating' || (!a.type && a.answer && !isNaN(a.answer) && Number(a.answer) <= 10 && Number(a.answer) >= 1)) && a.answer !== "") {
               const qLabel = a.question || a.label || 'Rating Umum';
               if (!categoryStats[qLabel]) {
                   categoryStats[qLabel] = { promoters: 0, passives: 0, detractors: 0, total: 0 };
               }
               const val = Number(a.answer);
               if (val >= 9) categoryStats[qLabel].promoters++;
               else if (val >= 7) categoryStats[qLabel].passives++;
               else categoryStats[qLabel].detractors++;

               categoryStats[qLabel].total++;
            }
          });
        }
      } catch (e) {}
    });

    let totalQuestionNPS = 0;
    let questionCount = 0;
    
    let totalP = 0, totalPa = 0, totalD = 0;

    for (let q in categoryStats) {
        const stats = categoryStats[q];
        if (stats.total > 0) {
            const pctPro = stats.promoters / stats.total;
            const pctDet = stats.detractors / stats.total;
            const qNPS = (pctPro - pctDet) * 100;
            
            if (qNPS !== 0) {
                totalQuestionNPS += qNPS;
                questionCount++;
            }

            totalP += stats.promoters;
            totalPa += stats.passives;
            totalD += stats.detractors;
        }
    }

    let avgNpsScore = questionCount > 0 ? (totalQuestionNPS / questionCount) : 0;
    avgNpsScore = avgNpsScore >= 0 ? Math.ceil(avgNpsScore) : Math.floor(avgNpsScore);

    const words = allCommentText.replace(/[^\w\s]/gi, '').split(/\s+/);
    const wordCounts = {};
    words.forEach(w => {
      if (w.length > 3 && !stopWords.includes(w)) {
        wordCounts[w] = (wordCounts[w] || 0) + 1;
      }
    });
    
    const topKeywords = Object.keys(wordCounts)
      .sort((a,b) => wordCounts[b] - wordCounts[a])
      .slice(0, 5) 
      .map(k => ({ word: k, count: wordCounts[k] }));

    let sentimentLabel = "";
    if (avgNpsScore >= 40) sentimentLabel = "Luar Biasa";
    else if (avgNpsScore >= 35) sentimentLabel = "Sangat Disukai";
    else if (avgNpsScore >= 0) sentimentLabel = "Cukup Baik";
    else sentimentLabel = "Perlu Evaluasi";

    let summaryText = "";
    if (topKeywords.length > 0) {
       const keyStr = topKeywords.slice(0, 3).map(k => k.word).join(", ");
       summaryText = `Menu ini dinilai ${sentimentLabel}. Panelis menyoroti aspek: ${keyStr}.`;
    } else {
       summaryText = `Menu ini dinilai ${sentimentLabel}, namun belum ada kata kunci spesifik yang dominan.`;
    }

    const avgPrice = priceCount > 0 ? (totalPrice / priceCount) : 0;
    const suggestedPrice = avgPrice * 1.20;
    const dominantPanelistGroup = Object.keys(panelistGroupCounts).sort((a, b) => panelistGroupCounts[b] - panelistGroupCounts[a])[0] || '';
    
    rankings.push({ 
      ...menu, 
      nps: avgNpsScore,
      promoters: totalP, 
      passives: totalPa, 
      detractors: totalD,
      count: totalResp,
      comments: commentList.reverse().slice(0, 20), 
      keywords: topKeywords,
      avgPrice: avgPrice,
      suggestedPrice: suggestedPrice,
      dominantPanelistGroup: dominantPanelistGroup,
      summaryText: summaryText, 
      rawScores: rawScores.sort((a,b)=>b-a)
    });
  });

  rankings.sort((a, b) => b.nps - a.nps);
  
  const result = {
    rankings: rankings,
    totalUniquePanelists: uniquePanelistsSet.size, // Size diambil setelah difilter
    totalSubmissions: totalValidSubmissions
  };
  putCachedJson_(cacheKey, result);
  return result;
}

// --- FUNGSI UPGRADE PDF REPORT ---
function generatePDFReport(projectId, includeHidden) {
  const projSheet = getSheetOrSetup('Projects');
  const menuSheet = getSheetOrSetup('Menus');
  const respSheet = getSheetOrSetup('Responses');

  const projData = projSheet.getDataRange().getValues();
  let projectName = "Project Report";
  for (let i = 1; i < projData.length; i++) {
    if (String(projData[i][0]) === String(projectId)) {
      projectName = projData[i][1]; break;
    }
  }

  const menus = menuSheet.getDataRange().getValues().slice(1);
  menus.sort((a, b) => {
    if (String(a[1]) !== String(b[1])) return 0;
    const orderA = Number(a[7]) > 0 ? Number(a[7]) : Number.MAX_SAFE_INTEGER;
    const orderB = Number(b[7]) > 0 ? Number(b[7]) : Number.MAX_SAFE_INTEGER;
    return orderA - orderB;
  });
  const responses = respSheet.getDataRange().getValues().slice(1);

  const today = new Date();
  const dateOptions = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', timeZone: 'Asia/Jakarta' };
  const dateStr = today.toLocaleDateString('id-ID', dateOptions);

  // Palet Warna & Style Dasar
  const styleLaunch = "background-color: #dcfce7; color: #166534; border: 1px solid #bbf7d0;";
  const styleRevisi = "background-color: #fef08a; color: #854d0e; border: 1px solid #fde047;";
  const styleDrop = "background-color: #fee2e2; color: #991b1b; border: 1px solid #fecaca;";
  const badgeBase = "display: inline-block; padding: 6px 14px; border-radius: 20px; font-weight: bold; font-size: 10px; text-transform: uppercase; letter-spacing: 0.5px;";

  let html = `
  <html>
  <head>
    <style>
      body { font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; font-size: 12px; color: #334155; margin: 0; padding: 0; background-color: #f8fafc; }
      .page-container { background-color: #ffffff; width: 100%; padding: 40px; box-sizing: border-box; }
      
      .header-table { width: 100%; border-bottom: 4px solid #34ae70; padding-bottom: 25px; margin-bottom: 35px; }
      .brand-title { color: #0f172a; font-size: 32px; font-weight: 800; margin: 0; letter-spacing: -1px; text-transform: uppercase; }
      .brand-subtitle { color: #34ae70; font-size: 12px; font-weight: bold; letter-spacing: 2px; text-transform: uppercase; margin-top: 5px; }
      .project-card { background-color: #f1f5f9; padding: 15px 20px; border-radius: 12px; border-left: 4px solid #cbd5e1; }
      .project-label { font-size: 10px; color: #64748b; text-transform: uppercase; font-weight: bold; letter-spacing: 1px; margin: 0 0 4px 0; }
      .project-name { font-size: 16px; color: #0f172a; font-weight: 800; margin: 0 0 6px 0; }
      .project-date { font-size: 11px; color: #94a3b8; margin: 0; }

      .legend-container { margin-bottom: 30px; }
      .legend-title { font-size: 11px; font-weight: bold; color: #64748b; text-transform: uppercase; margin-bottom: 10px; letter-spacing: 1px; border-bottom: 1px solid #e2e8f0; padding-bottom: 5px;}
      .legend-table { width: 100%; border-collapse: collapse; background: #ffffff; }
      .legend-table td { padding: 10px; font-size: 11px; color: #475569; }
      
      /* Pembaruan Layout Card agar tabel tidak terpotong */
      .menu-card { margin-bottom: 40px; page-break-inside: avoid; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 10px -3px rgba(0,0,0,0.05); }
      .menu-header-bg { background-color: #f8fafc; border-bottom: 1px solid #e2e8f0; padding: 20px; }
      .menu-title { font-size: 20px; font-weight: bold; color: #0f172a; margin: 0 0 6px 0; }
      .menu-respondents { font-size: 11px; color: #64748b; font-weight: bold; text-transform: uppercase; letter-spacing: 0.5px; }
      
      .img-box { width: 75px; height: 75px; object-fit: cover; border-radius: 10px; border: 2px solid #ffffff; box-shadow: 0 2px 5px rgba(0,0,0,0.05); display: inline-block; }
      
      .data-table { width: 100%; border-collapse: collapse; }
      .data-table th { padding: 12px 10px; border-bottom: 2px solid #cbd5e1; color: #64748b; font-weight: bold; font-size: 10px; text-transform: uppercase; letter-spacing: 0.5px; text-align: center; background-color: #f8fafc; }
      .data-table td { padding: 14px 10px; border-bottom: 1px solid #f1f5f9; text-align: center; font-size: 13px; color: #334155; }
      .data-table tr.alt-row td { background-color: #fcfcfd; }
      .text-left { text-align: left !important; }
      
      .score-box { font-size: 38px; font-weight: 800; color: #0f172a; line-height: 1; }
      .score-label { font-size: 10px; color: #64748b; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; }
      
      .total-row td { background-color: #f1f5f9 !important; font-weight: 800; font-size: 13px; color: #0f172a; border-top: 2px solid #cbd5e1; padding: 16px 10px; }
      
      .footer { margin-top: 40px; text-align: center; font-size: 10px; color: #94a3b8; border-top: 1px solid #e2e8f0; padding-top: 20px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; }
    </style>
  </head>
  <body>
    <div class="page-container">
      
      <!-- Header Area -->
      <table class="header-table" cellpadding="0" cellspacing="0">
        <tr>
          <td valign="middle">
            <h1 class="brand-title">Product Intelligence</h1>
            <p class="brand-subtitle">Sensory Evaluation Report</p>
          </td>
          <td valign="middle" align="right" width="300">
            <div class="project-card" align="left">
              <p class="project-label">Analisis Project</p>
              <p class="project-name">${projectName}</p>
              <p class="project-date">Dihasilkan: ${dateStr}</p>
            </div>
          </td>
        </tr>
      </table>
      
      <!-- Legend Area -->
      <div class="legend-container">
        <div class="legend-title">Indikator Penilaian (NPS)</div>
        <table class="legend-table">
          <tr>
            <td width="33%"><span style="${badgeBase} ${styleLaunch}">Launch (&ge; 40)</span><br><span style="font-size:10px; color:#64748b; margin-top:5px; display:block;">Produk sangat direkomendasikan dan siap dipasarkan.</span></td>
            <td width="33%" style="border-left: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0;"><span style="${badgeBase} ${styleRevisi}">Revisi Minor (35 - 39)</span><br><span style="font-size:10px; color:#64748b; margin-top:5px; display:block;">Produk butuh penyesuaian/perbaikan skala kecil.</span></td>
            <td width="33%"><span style="${badgeBase} ${styleDrop}">Drop (&lt; 35)</span><br><span style="font-size:10px; color:#64748b; margin-top:5px; display:block;">Produk tidak memenuhi standar pasar (Dihentikan).</span></td>
          </tr>
        </table>
      </div>
  `;

  menus.forEach(m => {
    const mId = m[0];
    const mProjId = m[1];
    const mName = m[2];
    const mFotoId = m[4];
    const mStatus = m[6] ? String(m[6]) : 'Active';

    if (String(mProjId) !== String(projectId)) return;
    if (!includeHidden && mStatus === 'Hidden') return;

    const menuRaw = responses.filter(r => String(r[2]) === String(mId));
    const totalResp = menuRaw.length;
    if (totalResp === 0) return;

    const categoryStats = {};
    menuRaw.forEach(r => {
      try {
        const answers = JSON.parse(r[4]);
        answers.forEach(a => {
          // FIX: Hanya menghitung NPS jika tipe secara spesifik adalah rating. 
          if((a.type === 'rating' || (!a.type && a.answer && !isNaN(a.answer) && Number(a.answer) <= 10 && Number(a.answer) >= 1)) && a.answer !== "") {
            const qLabel = a.question || a.label;
            if(!categoryStats[qLabel]) categoryStats[qLabel] = { detractors: 0, passives: 0, promoters: 0, total: 0 };
            const val = Number(a.answer);
            if(val >= 9) categoryStats[qLabel].promoters++;
            else if(val >= 7) categoryStats[qLabel].passives++;
            else categoryStats[qLabel].detractors++;
            categoryStats[qLabel].total++;
          }
        });
      } catch(e) {}
    });

    let totalNpsScore = 0;
    let categoryCount = 0;

    let tableRows = '';
    let rowIndex = 0;
    for (const [category, stats] of Object.entries(categoryStats)) {
      if (stats.total === 0) continue;
      const pctDet = (stats.detractors / stats.total);
      const pctPas = (stats.passives / stats.total);
      const pctPro = (stats.promoters / stats.total);
      const nps = (pctPro - pctDet) * 100;
      
      if (nps !== 0) {
          totalNpsScore += nps;
          categoryCount++;
      }

      let statusStyle = styleDrop; let statusText = "Drop";
      if (nps >= 40) { statusStyle = styleLaunch; statusText = "Launch"; } 
      else if (nps >= 35) { statusStyle = styleRevisi; statusText = "Revisi Minor"; }

      const rowClass = rowIndex % 2 === 0 ? "" : "alt-row";
      
      tableRows += `<tr class="${rowClass}">
        <td class="text-left font-bold" style="color: #1e293b;">${category}</td>
        <td style="color: #ef4444;">${(pctDet * 100).toFixed(2)}%</td>
        <td style="color: #f59e0b;">${(pctPas * 100).toFixed(2)}%</td>
        <td style="color: #10b981;">${(pctPro * 100).toFixed(2)}%</td>
        <td><span style="${badgeBase} ${statusStyle}">${statusText}</span></td>
        <td style="font-weight: 800; font-size: 14px;">${nps.toFixed(2)}</td>
      </tr>`;
      rowIndex++;
    }

    let avgNpsRaw = categoryCount > 0 ? (totalNpsScore / categoryCount) : 0;
    let avgNps = avgNpsRaw >= 0 ? Math.ceil(avgNpsRaw) : Math.floor(avgNpsRaw);

    let statusAvgStyle = styleDrop; let statusAvgText = "DROP (TIDAK LANJUT)";
    if (avgNps >= 40) { statusAvgStyle = styleLaunch; statusAvgText = "LAUNCH (SIAP JUAL)"; }
    else if (avgNps >= 35) { statusAvgStyle = styleRevisi; statusAvgText = "REVISI MINOR"; }

    let imgTag = '<div style="width:75px;height:75px;background:#e2e8f0;text-align:center;line-height:75px;color:#94a3b8;border-radius:10px;display:inline-block;font-size:10px;font-weight:bold;text-transform:uppercase;">No Pic</div>';
    if (mFotoId) {
      try {
        const file = DriveApp.getFileById(mFotoId);
        const base64 = Utilities.base64Encode(file.getBlob().getBytes());
        const mime = file.getMimeType();
        imgTag = `<img src="data:${mime};base64,${base64}" class="img-box">`;
      } catch(e) {}
    }

    html += `
    <div class="menu-card">
      <div class="menu-header-bg">
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td width="90" valign="middle">
              ${imgTag}
            </td>
            <td valign="middle">
               <h2 class="menu-title">${mName}</h2>
               <div class="menu-respondents"><span style="color:#34ae70;">&bull;</span> Total Responden: ${totalResp} Panelis</div>
            </td>
            <td valign="middle" align="right">
               <div style="display:inline-block; vertical-align:middle; text-align:right; margin-right:15px;">
                   <div class="score-label" style="margin-bottom:6px;">Rata-rata NPS</div>
                   <span style="${badgeBase} ${statusAvgStyle}">${statusAvgText}</span>
               </div>
               <div class="score-box" style="color: ${avgNps >= 40 ? '#10b981' : (avgNps >= 35 ? '#f59e0b' : '#ef4444')}; display:inline-block; vertical-align:middle;">
                   ${avgNps}
               </div>
            </td>
          </tr>
        </table>
      </div>
      <div style="padding: 25px;">
        <table class="data-table">
          <tr>
            <th class="text-left" width="30%">Kategori / Atribut</th>
            <th width="12%">Detractors<br><span style="font-size:8px;font-weight:normal;">(Skor 1-6)</span></th>
            <th width="12%">Passives<br><span style="font-size:8px;font-weight:normal;">(Skor 7-8)</span></th>
            <th width="12%">Promoters<br><span style="font-size:8px;font-weight:normal;">(Skor 9-10)</span></th>
            <th width="20%">Status Evaluasi</th>
            <th width="14%">NPS Skor</th>
          </tr>
          ${tableRows}
          <tr class="total-row">
            <td class="text-left" colspan="4">TOTAL KESELURUHAN (Rata-rata Atribut)</td>
            <td><span style="${badgeBase} ${statusAvgStyle}">${statusAvgText}</span></td>
            <td style="font-size: 16px;">${avgNps}</td>
          </tr>
        </table>
      </div>
    </div>
    `;
  });

  html += `
      <div class="footer">
        Generated automatically by ProductPanel R&D Intelligence System &nbsp;&bull;&nbsp; Confidential & Internal Use Only
      </div>
    </div>
  </body></html>`;

  const blob = Utilities.newBlob(html, MimeType.HTML).setName("Report_" + projectName + ".pdf").getAs(MimeType.PDF);
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    filename: "Report_" + projectName.replace(/[^a-zA-Z0-9]/g, '_') + ".pdf"
  };
}

// --- FUNGSI PANELIS ---

function getActiveProjects() {
  const cached = getCachedJson_(FOOD_TEST_CACHE_KEYS.ACTIVE_PROJECTS);
  if (cached) return cached;
  const data = getSheetRows_('Projects', false);
  let projects = [];
  
  if (data.length > 1) {
    for (let i = 1; i < data.length; i++) {
      const status = data[i][3] ? String(data[i][3]) : 'Active';
      if (status === 'Active' && data[i][0]) {
        projects.push({ id: data[i][0], name: data[i][1] });
      }
    }
  }
  putCachedJson_(FOOD_TEST_CACHE_KEYS.ACTIVE_PROJECTS, projects);
  return projects;
}

function getMenusForProject(projectId, panelistId) {
  const projectKey = String(projectId || '');
  const cacheKey = 'FOOD_TEST_MENUS_' + projectKey;
  let menus = getCachedJson_(cacheKey);
  const data = menus ? null : getSheetRows_('Menus', false);
  if (!menus) menus = [];
  
  if (data && data.length > 1) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(projectId)) {
        const status = data[i][6] ? String(data[i][6]) : 'Active';
        if (status !== 'Hidden') {
            const fotoId = data[i][4];
            const imageUrl = fotoId ? `https://drive.google.com/thumbnail?id=${fotoId}&sz=w400` : '';
            menus.push({
              id: data[i][0],
              name: data[i][2],
              desc: data[i][3],
              image: imageUrl,
              questions: JSON.parse(data[i][5] || '[]'),
              order: Number(data[i][7]) > 0 ? Number(data[i][7]) : i
            });
        }
      }
    }
    menus.sort((a, b) => a.order - b.order);
    putCachedJson_(cacheKey, menus);
  }

  if (panelistId) {
    const submittedMenuIds = new Set(
      getSheetRows_('Responses', false).slice(1)
        .filter(r => String(r[1]) === projectKey && String(r[7]) === String(panelistId))
        .map(r => String(r[2]))
    );
    return menus.map(menu => Object.assign({}, menu, { submitted: submittedMenuIds.has(String(menu.id)) }));
  }
  return menus;
}

function ensureResponseRequestIdColumn_() {
  const sheet = getSheetOrSetup('Responses');
  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 9) {
    sheet.getRange(1, 9).setValue('Request_ID');
  } else {
    const header = String(sheet.getRange(1, 9).getValue() || '').trim();
    if (!header) sheet.getRange(1, 9).setValue('Request_ID');
  }

  return sheet;
}

/**
 * Cari response existing dengan dua lapis proteksi:
 * 1) requestId yang sama -> retry dari browser tidak membuat row baru.
 * 2) project + menu + panelist -> tetap menjaga aturan duplicate lama.
 *
 * Hanya kolom yang diperlukan yang dibaca, bukan seluruh sheet.
 */
function findExistingResponse_(sheet, projectId, menuId, panelistId, requestId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const rowCount = lastRow - 1;

  // B=Project_ID, C=Menu_ID
  const projectMenu = sheet.getRange(2, 2, rowCount, 2).getValues();

  // H=Panelist_ID, I=Request_ID
  const panelistRequest = sheet.getRange(2, 8, rowCount, 2).getValues();

  const pid = String(projectId || '');
  const mid = String(menuId || '');
  const panel = String(panelistId || '');
  const req = String(requestId || '');

  for (let i = 0; i < rowCount; i++) {
    const existingProject = String(projectMenu[i][0] || '');
    const existingMenu = String(projectMenu[i][1] || '');
    const existingPanel = String(panelistRequest[i][0] || '');
    const existingRequest = String(panelistRequest[i][1] || '');

    if (req && existingRequest === req) {
      return { duplicate: true, reason: 'requestId', row: i + 2 };
    }

    if (panel && existingProject === pid && existingMenu === mid && existingPanel === panel) {
      return { duplicate: true, reason: 'panelist', row: i + 2 };
    }
  }

  return null;
}

function submitResponse(data) {
  if (!data || !data.projectId || !data.menuId || !Array.isArray(data.answers)) {
    throw new Error('Data penilaian tidak lengkap.');
  }

  const requestId = String(data.requestId || '').trim();

  // Request ID wajib untuk submission dari versi baru.
  // Untuk kompatibilitas dengan versi lama, jika kosong sistem tetap bekerja
  // dengan duplicate protection project+menu+panelist.
  const lock = LockService.getScriptLock();

  // Tidak menunggu terlalu lama. Browser akan melakukan retry untuk submit
  // karena submission sudah memiliki requestId/idempotency key.
  if (!lock.tryLock(10000)) {
    throw new Error('Server sedang memproses penilaian lain. Silakan tunggu beberapa detik lalu coba kembali.');
  }

  try {
    const sheet = ensureResponseRequestIdColumn_();
    const panelistId = String(data.panelistId || '');

    const existing = findExistingResponse_(
      sheet,
      data.projectId,
      data.menuId,
      panelistId,
      requestId
    );

    if (existing) {
      return {
        success: true,
        duplicate: true,
        alreadyProcessed: true,
        duplicateReason: existing.reason
      };
    }

    const id = 'RESP-' + Utilities.getUuid();
    const timestamp = new Date();

    let npsScore = 0;
    let totalRating = 0;
    let countRating = 0;

    data.answers.forEach(function(a) {
      // Avg_Raw_Score hanya dihitung dari pertanyaan tipe rating.
      if (a && a.type === 'rating' && a.answer !== '' && a.answer != null) {
        const value = parseFloat(a.answer);
        if (!isNaN(value)) {
          totalRating += value;
          countRating++;
        }
      }
    });

    if (countRating > 0) {
      npsScore = parseFloat((totalRating / countRating).toFixed(2));
    }

    sheet.appendRow([
      id,
      data.projectId,
      data.menuId,
      data.panelistName,
      JSON.stringify(data.answers),
      npsScore,
      timestamp,
      panelistId,
      requestId
    ]);

    SpreadsheetApp.flush();
    invalidateProjectCache_(data.projectId);

    return {
      success: true,
      duplicate: false,
      alreadyProcessed: false,
      requestId: requestId
    };
  } finally {
    lock.releaseLock();
  }
}
