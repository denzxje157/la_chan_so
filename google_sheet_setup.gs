/**
 * GOOGLE APPS SCRIPT - HỆ THỐNG LÁ CHẮN SỐ ĐỒNG NAI (CHỐNG LỪA 101)
 * Phiên bản Production:
 * 1. Lưu trữ 7 bảng (Citizens, Scenarios, Reports, Blacklist, Broadcasts, TaiKhoan, TinTucCanhBao)
 * 2. Bảo mật GEMINI_API_KEY trong Script Properties (PropertiesService) + tự fallback model
 * 3. AI Phân tích 4 Trụ cột A05 & NCSC (action: analyze_scam)
 * 4. AI Bóc tách Link bài báo tạo Kịch bản NHẬP VAI chuẩn 3 tầng x ĐÚNG 3 phương án (action: generate_scenario_from_url)
 * 5. Quét tin tức thật từ RSS VnExpress Pháp luật / Báo Đồng Nai / Công an Đồng Nai (fetchDailyScamNews)
 * 6. Chống ghi dòng trống timestamp & cập nhật thông minh theo certId
 */

function initAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = [
    { name: 'Citizens', headers: ['Mã Chứng chỉ', 'Họ và Tên', 'SĐT', 'Quận/Huyện', 'Xã/Phường', 'Đối tượng', 'Điểm số', 'Xếp loại', 'Thời gian cấp'] },
    { name: 'Scenarios', headers: ['ID Kịch bản', 'Tiêu đề Tình huống', 'Dữ liệu JSON 3 Tầng'] },
    { name: 'Reports', headers: ['Mã Hồ sơ', 'Thời gian', 'Địa bàn', 'Đối tượng Bị Tố giác', 'Mô tả sự việc', 'Trạng thái'] },
    { name: 'Blacklist', headers: ['Loại', 'Giá trị (SĐT/STK/Link)', 'Đối tượng Mạo danh', 'Lượt báo cáo', 'Mức độ'] },
    { name: 'Broadcasts', headers: ['Nội dung Cảnh báo', 'Mức độ', 'Thời gian phát'] },
    { name: 'TaiKhoan', headers: ['Tài khoản (SĐT/CCCD)', 'Mật khẩu (SHA-256)', 'Vai trò (user/admin)', 'Nhóm đối tượng', 'Họ và Tên', 'CCCD', 'Năm sinh', 'Quận/Huyện', 'Xã/Phường', 'Ngày tạo'] },
    { name: 'TinTucCanhBao', headers: ['Tiêu đề', 'Tóm tắt', 'Ảnh minh họa (og:image)', 'Link gốc', 'Nguồn báo', 'Ngày đăng'] },
    { name: 'LichSuQuetAI', headers: ['Thời gian', 'Nội dung tin nhắn', 'Điểm rủi ro', 'Mức độ', 'Trụ cột 1', 'Trụ cột 2', 'Trụ cột 3', 'Trụ cột 4', 'Kết luận'] },
    { name: 'CamNangTuyenTruyen', headers: ['ID', 'Tiêu đề Cẩm nang', 'Địa bàn / Đơn vị', 'Nội dung chi tiết', 'Link PDF / Tài liệu', 'Ngày đăng'] }
  ];
  tabs.forEach(function(t) {
    var sh = ss.getSheetByName(t.name);
    if (!sh) {
      sh = ss.insertSheet(t.name);
      sh.appendRow(t.headers);
    }
  });
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getCitizens';
  
  if (action === 'getScenarios') {
    var sh = ss.getSheetByName('Scenarios');
    if (!sh) return jsonOut([]);
    var rows = sh.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][2]) {
        try { out.push(JSON.parse(rows[i][2])); } catch(err){}
      }
    }
    return jsonOut(out);
  }
  
  if (action === 'getBlacklist') {
    var sh = ss.getSheetByName('Blacklist');
    if (!sh) return jsonOut([]);
    var rows = sh.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][1]) continue;
      out.push({ type: rows[i][0], val: rows[i][1], title: rows[i][2], reports: rows[i][3] || 1, level: rows[i][4] || 'NGUY HIỂM' });
    }
    return jsonOut(out);
  }
  
  if (action === 'getReports') {
    var sh = ss.getSheetByName('Reports');
    if (!sh) return jsonOut([]);
    var rows = sh.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      out.push({ reportId: rows[i][0], timestamp: rows[i][1], district: rows[i][2], target: rows[i][3], desc: rows[i][4], status: rows[i][5] || 'Mới tiếp nhận' });
    }
    return jsonOut(out.reverse());
  }
  
  if (action === 'getBroadcasts') {
    var sh = ss.getSheetByName('Broadcasts');
    if (!sh) return jsonOut([]);
    var rows = sh.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      out.push({ message: rows[i][0], level: rows[i][1], timestamp: rows[i][2] });
    }
    return jsonOut(out.reverse());
  }

  if (action === 'get_news') {
    var sh = ss.getSheetByName('TinTucCanhBao');
    if (!sh) return jsonOut([]);
    var rows = sh.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      out.push({ title: rows[i][0], summary: rows[i][1], imageUrl: rows[i][2], url: rows[i][3], source: rows[i][4], publishedDate: rows[i][5] });
    }
    return jsonOut(out.reverse().slice(0, 20));
  }

  if (action === 'fetch_news_now') {
    var added = fetchDailyScamNews();
    return jsonOut({ ok: true, added: added });
  }

  if (action === 'get_handbooks') {
    var sh = ss.getSheetByName('CamNangTuyenTruyen');
    if (!sh) return jsonOut([]);
    var rows = sh.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][1]) continue;
      out.push({ id: rows[i][0], title: rows[i][1], unit: rows[i][2], content: rows[i][3], fileUrl: rows[i][4] || '', date: rows[i][5] || rows[i][4] || '' });
    }
    return jsonOut(out.reverse());
  }

  if (action === 'admin_login') {
    var sh = ensureAdminAccounts_(ss);
    var acc = (e.parameter.account || '').trim().toLowerCase();
    var pwh = (e.parameter.pwHash || '').trim().toLowerCase();
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var rowAcc = String(rows[i][0] || '').trim().toLowerCase();
      if (rowAcc === acc) {
        var role = String(rows[i][2] || 'user').trim().toLowerCase();
        if (role !== 'admin') {
          logAdminLogin_(ss, acc, rows[i][4] || '', 'TU_CHOI_KHONG_PHAI_ADMIN');
          return jsonOut({ handled: true, ok: false, reason: 'not_admin' });
        }
        var stored = String(rows[i][1] || '').trim();
        var storedHash = stored.length === 64 ? stored.toLowerCase() : sha256Hex_(stored);
        if (storedHash === pwh) {
          logAdminLogin_(ss, acc, rows[i][4] || 'Cán bộ Quản trị', 'THANH_CONG');
          return jsonOut({
            handled: true,
            ok: true,
            role: 'admin',
            profile: { account: rows[i][0], name: rows[i][4] || 'Cán bộ Quản trị', unit: rows[i][7] || 'Tỉnh Đồng Nai' }
          });
        } else {
          logAdminLogin_(ss, acc, rows[i][4] || '', 'SAI_MAT_KHAU');
          return jsonOut({ handled: true, ok: false, reason: 'wrong_password' });
        }
      }
    }
    logAdminLogin_(ss, acc, '', 'KHONG_TON_TAI');
    return jsonOut({ handled: true, ok: false, reason: 'not_found' });
  }

  if (action === 'login') {
    var sh = ss.getSheetByName('TaiKhoan');
    if (!sh) return jsonOut({ handled: false });
    var acc = (e.parameter.account || '').trim();
    var pwh = (e.parameter.pwHash || '').trim();
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if ((String(rows[i][0]) === acc || String(rows[i][5]) === acc)) {
        if (String(rows[i][1]) === pwh) {
          return jsonOut({
            handled: true, ok: true, role: rows[i][2] || 'user', group: rows[i][3] || 'student',
            profile: { account: rows[i][0], name: rows[i][4], cccd: rows[i][5], birthYear: rows[i][6], district: rows[i][7], ward: rows[i][8] }
          });
        } else {
          return jsonOut({ handled: true, ok: false });
        }
      }
    }
    return jsonOut({ handled: false });
  }

  // Mặc định: Lấy danh sách Citizens (bỏ qua mọi dòng trống)
  var sheet = ss.getSheets()[0];
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0] || !rows[i][1]) continue;
    result.push({
      certId: rows[i][0], name: rows[i][1], phone: rows[i][2],
      district: rows[i][3], ward: rows[i][4], role: rows[i][5],
      score: rows[i][6], rank: rows[i][7], timestamp: rows[i][8]
    });
  }
  return jsonOut(result);
}

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return jsonOut({ status: 'ignored', message: 'doPost chỉ nhận HTTP POST từ Web App.' });
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = {};
  try {
    data = JSON.parse(e.postData.contents);
  } catch(err) {
    return jsonOut({ status: 'error', message: 'Invalid JSON' });
  }

  // 1. AI Phân tích 4 Trụ cột A05 & NCSC (Bảo mật GEMINI_API_KEY trên Server) + Tự lưu LichSuQuetAI
  if (data.action === 'analyze_scam') {
    var aiOut = runGeminiScamAnalysis(data.content || '');
    if (aiOut && aiOut.ok && aiOut.data) {
      saveScamScanToSheet(ss, data.content || '', aiOut.data);
    }
    return jsonOut(aiOut);
  }

  // 1B. Lưu lịch sử quét tin nhắn từ Frontend lên Sheet LichSuQuetAI
  if (data.action === 'save_scam_scan') {
    saveScamScanToSheet(ss, data.content || '', data);
    return jsonOut({ status: 'ok' });
  }

  // 2B. Đọc toàn văn bài báo thật từ link gốc để hiển thị ngay trong app
  if (data.action === 'fetch_article_content') {
    return jsonOut(fetchArticleFullContent(data.url || ''));
  }

  // 2C. Đăng mẫu Cẩm nang Tuyên truyền mới (hỗ trợ upload file PDF từ máy Admin lên Google Drive -> đẩy xuống App Dân)
  if (data.action === 'save_handbook') {
    if (!data.title) return jsonOut({ status: 'ignored' });
    var sh = ss.getSheetByName('CamNangTuyenTruyen') || ss.insertSheet('CamNangTuyenTruyen');
    if (sh.getLastRow() === 0) sh.appendRow(['ID', 'Tiêu đề Cẩm nang', 'Địa bàn / Đơn vị', 'Nội dung chi tiết', 'Link PDF / Tài liệu', 'Ngày đăng']);
    var fileUrl = data.fileUrl || '';
    if (data.pdfBase64 && data.pdfName) {
      try {
        var bytes = Utilities.base64Decode(data.pdfBase64);
        var blob = Utilities.newBlob(bytes, 'application/pdf', data.pdfName);
        var file = DriveApp.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl = 'https://drive.google.com/uc?export=download&id=' + file.getId();
      } catch (err) {
        return jsonOut({ status: 'error', message: 'Lỗi lưu PDF lên Drive: ' + err });
      }
    }
    sh.appendRow([data.id || ('CN-' + Date.now()), data.title, data.unit || 'Công an & Tỉnh Đoàn Đồng Nai', data.content || 'Tài liệu tuyên truyền PDF đính kèm', fileUrl, Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy')]);
    return jsonOut({ status: 'ok', fileUrl: fileUrl });
  }

  // 2. AI Bóc tách Link bài báo tạo Kịch bản 3 tầng x ĐÚNG 4 phương án vượt CORS
  if (data.action === 'generate_scenario_from_url') {
    return jsonOut(runGeminiScenarioFromUrl(data.url || ''));
  }

  // 2D. AI Chuyển thể đoạn tóm tắt thành Kịch bản 3 tầng x ĐÚNG 4 phương án
  if (data.action === 'generate_scenario_from_text') {
    return jsonOut(runGeminiScenarioFromText(data.text || '', data.sourceUrl || ''));
  }

  // 10B. Xóa hồ sơ công dân khỏi Sheet Citizens theo certId
  if (data.action === 'delete_citizen') {
    if (!data.certId) return jsonOut({ status: 'ignored' });
    var shDel = ss.getSheets()[0];
    var rowsDel = shDel.getDataRange().getValues();
    for (var d = rowsDel.length - 1; d >= 1; d--) {
      if (String(rowsDel[d][0]) === String(data.certId)) {
        shDel.deleteRow(d + 1);
      }
    }
    return jsonOut({ status: 'deleted' });
  }

  // 3. Quét tin tức thật ngay lập tức
  if (data.action === 'fetch_news_now') {
    var added = fetchDailyScamNews();
    return jsonOut({ ok: true, added: added });
  }

  // 4. Lưu kịch bản (mỗi kịch bản 1 dòng)
  if (data.action === 'saveScenarios') {
    var sh = ss.getSheetByName('Scenarios') || ss.insertSheet('Scenarios');
    sh.clearContents();
    sh.appendRow(['ID Kịch bản', 'Tiêu đề Tình huống', 'Dữ liệu JSON 3 Tầng']);
    var list = data.scenarios || [];
    list.forEach(function(sc) {
      sh.appendRow([sc.id, sc.title, JSON.stringify(sc)]);
    });
    return jsonOut({status: 'ok', count: list.length});
  }

  // 5. Lưu Blacklist
  if (data.action === 'saveBlacklist') {
    if (!data.val) return jsonOut({status: 'ignored'});
    var sh = ss.getSheetByName('Blacklist') || ss.insertSheet('Blacklist');
    if (sh.getLastRow() === 0) sh.appendRow(['Loại', 'Giá trị (SĐT/STK/Link)', 'Đối tượng Mạo danh', 'Lượt báo cáo', 'Mức độ']);
    sh.appendRow([data.type || 'SĐT/Link', data.val, data.title || 'Lừa đảo công nghệ cao', 1, 'NGUY HIỂM']);
    return jsonOut({status: 'ok'});
  }

  // 6. Lưu Tố giác
  if (data.action === 'submitReport') {
    if (!data.target) return jsonOut({status: 'ignored'});
    var sh = ss.getSheetByName('Reports') || ss.insertSheet('Reports');
    if (sh.getLastRow() === 0) sh.appendRow(['Mã Hồ sơ', 'Thời gian', 'Địa bàn', 'Đối tượng Bị Tố giác', 'Mô tả sự việc', 'Trạng thái']);
    sh.appendRow([data.reportId || ('PA05-' + Date.now()), data.timestamp || new Date().toLocaleString('vi-VN'), data.district || 'Đồng Nai', data.target, data.desc || '', data.status || 'Đang xác minh']);
    return jsonOut({status: 'ok'});
  }

  // 7. Lưu Broadcast khẩn cấp
  if (data.action === 'saveBroadcast') {
    if (!data.message) return jsonOut({status: 'ignored'});
    var sh = ss.getSheetByName('Broadcasts') || ss.insertSheet('Broadcasts');
    if (sh.getLastRow() === 0) sh.appendRow(['Nội dung Cảnh báo', 'Mức độ', 'Thời gian phát']);
    sh.appendRow([data.message, data.level || 'KHẨN CẤP', new Date().toLocaleString('vi-VN')]);
    return jsonOut({status: 'ok'});
  }

  // 8. Đăng ký tài khoản (luôn role = 'user')
  if (data.action === 'register') {
    if (!data.account || !data.pwHash) return jsonOut({status: 'ignored'});
    var sh = ss.getSheetByName('TaiKhoan') || ss.insertSheet('TaiKhoan');
    if (sh.getLastRow() === 0) sh.appendRow(['Tài khoản (SĐT/CCCD)', 'Mật khẩu (SHA-256)', 'Vai trò (user/admin)', 'Nhóm đối tượng', 'Họ và Tên', 'CCCD', 'Năm sinh', 'Quận/Huyện', 'Xã/Phường', 'Ngày tạo']);
    sh.appendRow([data.account, data.pwHash, 'user', data.group || 'student', data.name || '', data.cccd || '', data.birthYear || '', data.district || '', data.ward || '', new Date().toLocaleString('vi-VN')]);
    return jsonOut({status: 'ok'});
  }

  // 9. Cập nhật nhóm đối tượng (senior / student / member)
  if (data.action === 'update_group') {
    var sh = ss.getSheetByName('TaiKhoan');
    if (sh && data.account) {
      var rows = sh.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.account)) {
          sh.getRange(i + 1, 4).setValue(data.group);
          break;
        }
      }
    }
    return jsonOut({status: 'ok'});
  }

  // 10. Mặc định: Lưu chứng nhận công dân (CHẶN TUYỆT ĐỐI GHI DÒNG TRỐNG & CẬP NHẬT NẾU TRÙNG certId)
  if (!data.certId || !data.name) {
    return jsonOut({ status: 'ignored', message: 'Bỏ qua request thiếu certId hoặc name (tránh tạo dòng trống).' });
  }
  var sheet = ss.getSheets()[0];
  var rows = sheet.getDataRange().getValues();
  var nowStr = new Date().toLocaleString('vi-VN');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.certId)) {
      sheet.getRange(i + 1, 2, 1, 8).setValues([[data.name, data.phone || '', data.district || '', data.ward || '', data.role || 'student', data.score || 0, data.rank || '', nowStr]]);
      return jsonOut({status: 'updated'});
    }
  }
  sheet.appendRow([data.certId, data.name, data.phone || '', data.district || '', data.ward || '', data.role || 'student', data.score || 0, data.rank || '', nowStr]);
  return jsonOut({status: 'ok'});
}

function sha256Hex_(str) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(str || ''), Utilities.Charset.UTF_8);
  return raw.map(function(b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function ensureAdminAccounts_(ss) {
  var sh = ss.getSheetByName('TaiKhoan') || ss.insertSheet('TaiKhoan');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Tài khoản (SĐT/CCCD)', 'Mật khẩu (SHA-256)', 'Vai trò (user/admin)', 'Nhóm đối tượng', 'Họ và Tên', 'CCCD', 'Năm sinh', 'Quận/Huyện', 'Xã/Phường', 'Ngày tạo']);
  }
  var rows = sh.getDataRange().getValues();
  var hasAdmin = false;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][2] || '').trim().toLowerCase() === 'admin') {
      hasAdmin = true;
      break;
    }
  }
  if (!hasAdmin) {
    var defaultHash = sha256Hex_('dongnai2026');
    var nowStr = new Date().toLocaleString('vi-VN');
    sh.appendRow(['admin', defaultHash, 'admin', 'admin', 'Cán bộ Quản trị SOC', '', '', 'Tỉnh Đồng Nai', 'Trực ban PA05 / Tỉnh Đoàn', nowStr]);
    sh.appendRow(['admin_dongnai', defaultHash, 'admin', 'admin', 'Quản trị viên Tỉnh Đoàn Đồng Nai', '', '', 'TP. Biên Hòa', 'Tỉnh Đoàn Đồng Nai', nowStr]);
    sh.appendRow(['pa05_dongnai', defaultHash, 'admin', 'admin', 'Cán bộ Phòng PA05 Công an Đồng Nai', '', '', 'TP. Biên Hòa', 'Công an Tỉnh Đồng Nai', nowStr]);
  }
  return sh;
}

function logAdminLogin_(ss, account, name, status) {
  try {
    var logSh = ss.getSheetByName('NhatKyDangNhap') || ss.insertSheet('NhatKyDangNhap');
    if (logSh.getLastRow() === 0) {
      logSh.appendRow(['Thời gian', 'Tài khoản', 'Họ tên / Đơn vị', 'Trạng thái']);
    }
    logSh.appendRow([new Date().toLocaleString('vi-VN'), account || '', name || '', status || '']);
  } catch (err) {}
}

// ==================== HÀM GỌI GEMINI AI TRÊN SERVER (TỰ FALLBACK MODEL & BẮT LỖI CHI TIẾT) ====================
function getGeminiApiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
}

function callGeminiJSON(prompt, temp) {
  var key = getGeminiApiKey();
  if (!key) {
    return { ok: false, fallback: true, message: 'Chưa cấu hình GEMINI_API_KEY trong Script Properties.' };
  }
  var models = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash'];
  var lastErr = 'Unknown error';
  for (var i = 0; i < models.length; i++) {
    try {
      var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent?key=' + key;
      var res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: temp || 0.2, responseMimeType: 'application/json' }
        }),
        muteHttpExceptions: true
      });
      var raw = res.getContentText();
      var j = JSON.parse(raw);
      if (j && j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts && j.candidates[0].content.parts[0]) {
        var txt = j.candidates[0].content.parts[0].text;
        return { ok: true, model: models[i], data: JSON.parse(txt) };
      }
      lastErr = (j && j.error && j.error.message) ? (models[i] + ': ' + j.error.message) : (models[i] + ': ' + raw.slice(0, 180));
    } catch (e) {
      lastErr = models[i] + ': ' + String(e);
    }
  }
  return { ok: false, fallback: true, message: lastErr };
}

function saveScamScanToSheet(ss, content, d) {
  try {
    if (!content) return;
    var sh = ss.getSheetByName('LichSuQuetAI') || ss.insertSheet('LichSuQuetAI');
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Thời gian', 'Nội dung tin nhắn', 'Điểm rủi ro', 'Mức độ', 'Trụ cột 1', 'Trụ cột 2', 'Trụ cột 3', 'Trụ cột 4', 'Kết luận']);
    }
    var p = Array.isArray(d.pillars) ? d.pillars : [];
    sh.appendRow([
      new Date().toLocaleString('vi-VN'),
      String(content).slice(0, 500),
      d.riskScore || 0,
      d.level || '',
      p[0] ? (p[0].tag + ' - ' + p[0].risk) : '',
      p[1] ? (p[1].tag + ' - ' + p[1].risk) : '',
      p[2] ? (p[2].tag + ' - ' + p[2].risk) : '',
      p[3] ? (p[3].tag + ' - ' + p[3].risk) : '',
      d.recommendation || d.summary || ''
    ]);
  } catch (e) {}
}

function runGeminiScamAnalysis(content) {
  var prompt = 'Bạn là chuyên gia An ninh mạng PA05 Công an Đồng Nai. Hãy phân tích đoạn tin nhắn sau theo 2 phần ngắn gọn, đi thẳng vào trọng tâm:\n' +
    '1. 4 Trụ cột A05 & NCSC: Mỗi trụ cột ghi NHÃN NGẮN GỌN (ví dụ: "Công an / Cơ quan điều tra", "Dọa án / Bắt cóc online", "Ảnh hưởng đến tiền & tài sản", "SĐT lạ 0899..., Link .vip, Tệp .APK") và bên dưới giải thích NGẮN GỌN nguy cơ tiềm ẩn của từng thông tin trong trụ cột đó.\n' +
    '2. Bóc tách trực tiếp các cụm từ đáng ngờ CÓ TRONG CHÍNH TIN NHẮN NGƯỜI DÙNG (highlights).\n' +
    '3. Kết luận cảnh báo ngắn gọn 1 câu.\n' +
    'Trả về duy nhất 1 JSON hợp lệ:\n' +
    '{"riskScore": 0-100, "level": "An toàn|Cảnh giác|Nguy hiểm cao", "summary": "Kết luận ngắn 1 câu", "pillars": [' +
    '{"name": "Trụ cột 1: Giả mạo danh tính", "tag": "nhãn ngắn gọn", "status": "An toàn|Cảnh giác|Nguy hiểm", "risk": "Giải thích nguy cơ tiềm ẩn của trụ cột 1"},' +
    '{"name": "Trụ cột 2: Thao túng tâm lý", "tag": "nhãn ngắn gọn", "status": "An toàn|Cảnh giác|Nguy hiểm", "risk": "Giải thích nguy cơ tiềm ẩn của trụ cột 2"},' +
    '{"name": "Trụ cột 3: Hành vi rủi ro", "tag": "nhãn ngắn gọn", "status": "An toàn|Cảnh giác|Nguy hiểm", "risk": "Giải thích nguy cơ tiềm ẩn của trụ cột 3"},' +
    '{"name": "Trụ cột 4: Chỉ báo kỹ thuật IoCs", "tag": "liệt kê SĐT/STK/Link/File kỹ thuật tìm thấy hoặc Không có", "status": "An toàn|Cảnh giác|Nguy hiểm", "risk": "Giải thích nguy cơ tiềm ẩn của chỉ báo kỹ thuật"}' +
    '], "highlights": [{"phrase": "trích nguyên văn cụm từ đáng ngờ trong tin nhắn", "reason": "giải thích ngắn gọn 1 câu"}], "recommendation": "Kết luận cảnh báo & khuyến nghị hành động ngắn gọn"}\n\nNội dung cần phân tích:\n' + content;
  return callGeminiJSON(prompt, 0.15);
}

function sanitizeRoleplayMsgs(msgs, fallbackLine) {
  if (!Array.isArray(msgs) || msgs.length === 0) {
    return [{ f: 'them', text: fallbackLine }];
  }
  return msgs.map(function(m) {
    var txt = String((m && m.text) || '').trim();
    // Chặn tuyệt đối trường hợp văn phong báo chí dài dòng lọt vào khung chat
    if (m && m.f === 'them' && (txt.length > 260 || (/theo cơ quan|bị cáo buộc|tòa án nhân dân|công an tỉnh cho biết|vừa qua/i.test(txt) && !/tôi là|chào anh|chào chị|alo/i.test(txt)))) {
      txt = fallbackLine;
    }
    return { f: (m && m.f) ? m.f : 'them', text: txt || fallbackLine };
  });
}

function normalizeScenarioSteps(scen) {
  if (!scen || !Array.isArray(scen.steps)) return scen;
  var defaultScores = [100, 70, 45];
  var fallbackLines = [
    'Chào anh/chị, tôi liên hệ để thông báo hồ sơ của anh/chị đang có [[vấn đề khẩn cấp cần xác minh ngay]] qua kênh riêng này.',
    'Thời gian sắp hết rồi, anh/chị phải [[bấm vào đường link hoặc cung cấp thông tin]] ngay trong 15 phút nếu không sẽ gặp rắc rối lớn!',
    'Bước cuối cùng: anh/chị hãy [[chuyển khoản phí xác thực / gửi mã OTP]] ngay để hệ thống mở khóa hồ sơ!'
  ];
  scen.steps = scen.steps.slice(0, 3).map(function(st, idx) {
    st.msgs = sanitizeRoleplayMsgs(st.msgs, fallbackLines[idx] || fallbackLines[0]);
    var opts = Array.isArray(st.opts) ? st.opts : [];
    var safeScore = defaultScores[idx] || 45;
    var hasSafe = opts.some(function(o) { return (o.pts || 0) >= 40; });
    if (!hasSafe) {
      opts.unshift({ t: 'Từ chối ngay, chặn liên hệ & kiểm tra qua kênh chính thức', say: 'Tôi sẽ tự xác minh trực tiếp qua tổng đài chính thức, không giao dịch qua mạng.', pts: safeScore, fb: 'Chính xác! Không làm theo yêu cầu qua mạng và luôn xác minh qua kênh chính thống.' });
    }
    var trapCount = opts.filter(function(o) { return (o.pts || 0) === 0; }).length;
    if (trapCount < 1) {
      opts.push({ t: 'Làm theo hướng dẫn chuyển tiền / bấm link ngay kẻo lỡ', say: 'Vâng, để tôi làm theo hướng dẫn ngay bây giờ.', pts: 0, fb: 'Nguy hiểm! Bạn vừa rơi vào bẫy thao túng của đối tượng lừa đảo.', next: idx < 2 ? idx + 1 : undefined });
    }
    if (trapCount < 2) {
      opts.push({ t: 'Cung cấp ảnh CCCD / thông tin cá nhân để đối tượng kiểm tra hộ', say: 'Đây là thông tin của tôi, anh/chị kiểm tra giúp tôi nhé.', pts: 0, fb: 'Rất nguy hiểm! Lộ thông tin CCCD/tài khoản giúp kẻ gian tiếp tục leo thang lừa đảo.', next: idx < 2 ? idx + 1 : undefined });
    }
    while (opts.length < 4) {
      opts.push({ t: 'Chần chừ hỏi thêm thông tin / xin xem giấy tờ qua mạng', say: 'Anh/chị gửi hình ảnh giấy tờ hoặc mã hồ sơ qua đây cho tôi xem trước được không?', pts: 15, fb: 'Cẩn trọng! Kẻ gian thường dùng giấy tờ làm giả để tạo lòng tin, hãy dừng đối thoại ngay.' });
    }
    st.opts = opts.slice(0, 4);
    return st;
  });
  return scen;
}

function buildRoleplayPrompt(title, bodyText) {
  return 'Bạn là chuyên gia thiết kế bài tập tình huống của Phòng PA05 Công an Đồng Nai.\n' +
    'NHIỆM VỤ: Đọc bài báo/tóm tắt bên dưới và CHUYỂN THỂ thủ đoạn đó thành KỊCH BẢN NHẬP VAI HỘI THOẠI TRỰC TIẾP (giữa kẻ lừa đảo và người dân qua Zalo/SMS/Cuộc gọi).\n' +
    'QUY TẮC BẮT BUỘC:\n' +
    '1. TUYỆT ĐỐI KHÔNG copy nguyên văn bài báo hay văn phong tin tức ("Theo cơ quan công an...", "Bị cáo buộc...", "Ngày ... đối tượng..."). Mỗi tin nhắn trong msgs chỉ ngắn 1-3 câu, viết lời thoại trực tiếp ngôi thứ nhất ("Chào anh/chị...", "Tôi là...").\n' +
    '2. Đánh dấu [[cụm từ dấu hiệu lừa đảo]] trong lời thoại của đối tượng.\n' +
    '3. MỖI BƯỚC (steps[0], steps[1], steps[2]) BẮT BUỘC PHẢI CÓ ĐÚNG 4 PHƯƠNG ÁN TRONG MẢNG opts:\n' +
    '   - Phương án 1: An toàn tuyệt đối (pts: 100 ở bước 0, 70 ở bước 1, 45 ở bước 2)\n' +
    '   - Phương án 2: Nguy hiểm / Sập bẫy trực tiếp (pts: 0, có "next": 1 hoặc 2 nếu chưa phải bước cuối)\n' +
    '   - Phương án 3: Nguy hiểm / Lộ thông tin cá nhân hoặc làm theo một phần (pts: 0, có "next": 1 hoặc 2 nếu chưa phải bước cuối)\n' +
    '   - Phương án 4: Do dự / Chưa dứt khoát (pts: 15)\n' +
    '4. Trả về duy nhất 1 JSON hợp lệ theo schema:\n' +
    '{"id":"scen-ai-1","icon":"🚨","color":"#1E40AF","bg":"#EFF6FF","diff":"Khó","tag":"...","title":"...","plat":"Zalo/SMS","persona":{"name":"...","sub":"...","av":"🚨"},"flags":["..."],"rule":"...","fact":"...",' +
    '"steps":[' +
    '{"msgs":[{"f":"sys","text":"Tình huống mở đầu..."},{"f":"them","text":"Lời thoại mồi ngắn gọn của kẻ lừa đảo có [[dấu hiệu]]..."}],"q":"Bạn sẽ xử lý thế nào?","opts":[{"t":"...","say":"...","pts":100,"fb":"..."},{"t":"...","say":"...","pts":0,"fb":"...","next":1},{"t":"...","say":"...","pts":0,"fb":"...","next":1},{"t":"...","say":"...","pts":15,"fb":"..."}]},' +
    '{"msgs":[{"f":"them","text":"Lời thoại ép tâm lý tầng 2..."}],"q":"Đối tượng đang thúc ép, bạn chọn gì?","opts":[{"t":"...","say":"...","pts":70,"fb":"..."},{"t":"...","say":"...","pts":0,"fb":"...","next":2},{"t":"...","say":"...","pts":0,"fb":"...","next":2},{"t":"...","say":"...","pts":15,"fb":"..."}]},' +
    '{"msgs":[{"f":"them","text":"Lời thoại chốt hạ yêu cầu chuyển tiền/OTP..."}],"q":"Quyết định cuối cùng của bạn?","opts":[{"t":"...","say":"...","pts":45,"fb":"..."},{"t":"...","say":"...","pts":0,"fb":"..."},{"t":"...","say":"...","pts":0,"fb":"..."},{"t":"...","say":"...","pts":15,"fb":"..."}]}' +
    ']}\n\nTiêu đề: ' + (title || 'Tình huống thực tế') + '\nNội dung:\n' + bodyText;
}

function runGeminiScenarioFromText(text, sourceUrl) {
  if (!text) return { ok: false, message: 'Thiếu nội dung tóm tắt' };
  var prompt = buildRoleplayPrompt('Tình huống từ tóm tắt', String(text).slice(0, 4500));
  var aiRes = callGeminiJSON(prompt, 0.25);
  if (aiRes.ok && aiRes.data) {
    var scen = normalizeScenarioSteps(aiRes.data);
    if (sourceUrl) scen.sourceUrl = sourceUrl;
    return { ok: true, scenario: scen, model: aiRes.model };
  }
  return { ok: false, fallback: true, aiError: aiRes.message };
}

function runGeminiScenarioFromUrl(url) {
  if (!url) return { ok: false, message: 'Thiếu URL' };
  var title = '', bodyText = '', imageUrl = '';
  try {
    var resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; DongNaiCyberShield/1.0)' }
    });
    var html = resp.getContentText();
    var mTitle = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
    if (mTitle) title = mTitle[1].replace(/\s+/g, ' ').trim();
    var mImg = html.match(/<meta[^>]*property=["']og:image["'][^>]*content=["']([^"']+)["']/i);
    if (mImg) imageUrl = mImg[1];
    bodyText = html.replace(/<script[\s\S]*?<\/script>/gi, ' ').replace(/<style[\s\S]*?<\/style>/gi, ' ').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 5000);
  } catch(e) {
    return { ok: false, message: 'Không thể fetch URL: ' + e };
  }

  var prompt = 'Bạn là chuyên gia thiết kế bài tập tình huống của Phòng PA05 Công an Đồng Nai.\n' +
    'NHIỆM VỤ: Đọc bài báo bên dưới và CHUYỂN THỂ thủ đoạn trong bài báo thành KỊCH BẢN NHẬP VAI HỘI THOẠI TRỰC TIẾP (giữa kẻ lừa đảo và người dân qua Zalo/SMS/Cuộc gọi).\n' +
    'QUY TẮC BẮT BUỘC:\n' +
    '1. TUYỆT ĐỐI KHÔNG copy nguyên văn bài báo hay văn phong tin tức ("Theo cơ quan công an...", "Ngày ... đối tượng..."). Phải viết lời thoại trực tiếp mà kẻ lừa đảo nhắn/nói với nạn nhân (ngôi thứ nhất: "Chào anh/chị...", "Tôi là...").\n' +
    '2. Đánh dấu [[cụm từ dấu hiệu lừa đảo]] trong lời thoại của đối tượng.\n' +
    '3. MỖI BƯỚC (steps[0], steps[1], steps[2]) BẮT BUỘC PHẢI CÓ ĐÚNG 4 PHƯƠNG ÁN TRONG MẢNG opts:\n' +
    '   - Phương án 1: An toàn tuyệt đối (pts: 100 ở bước 0, 70 ở bước 1, 45 ở bước 2)\n' +
    '   - Phương án 2: Nguy hiểm / Sập bẫy trực tiếp (pts: 0, có "next": 1 hoặc 2 nếu chưa phải bước cuối)\n' +
    '   - Phương án 3: Nguy hiểm / Lộ thông tin cá nhân hoặc làm theo một phần (pts: 0, có "next": 1 hoặc 2 nếu chưa phải bước cuối)\n' +
    '   - Phương án 4: Do dự / Chưa dứt khoát (pts: 15)\n' +
    '4. Trả về duy nhất 1 JSON hợp lệ theo schema:\n' +
    '{"id":"scen-ai-1","icon":"🚨","color":"#1E40AF","bg":"#EFF6FF","diff":"Khó","tag":"...","title":"...","plat":"Zalo/SMS","persona":{"name":"...","sub":"...","av":"🚨"},"flags":["..."],"rule":"...","fact":"...",' +
    '"steps":[' +
    '{"msgs":[{"f":"sys","text":"Tình huống mở đầu..."},{"f":"them","text":"Lời thoại mồi của kẻ lừa đảo có [[dấu hiệu]]..."}],"q":"Bạn sẽ xử lý thế nào?","opts":[{"t":"...","say":"...","pts":100,"fb":"..."},{"t":"...","say":"...","pts":0,"fb":"...","next":1},{"t":"...","say":"...","pts":0,"fb":"...","next":1},{"t":"...","say":"...","pts":15,"fb":"..."}]},' +
    '{"msgs":[{"f":"them","text":"Lời thoại ép tâm lý tầng 2..."}],"q":"Đối tượng đang thúc ép, bạn chọn gì?","opts":[{"t":"...","say":"...","pts":70,"fb":"..."},{"t":"...","say":"...","pts":0,"fb":"...","next":2},{"t":"...","say":"...","pts":0,"fb":"...","next":2},{"t":"...","say":"...","pts":15,"fb":"..."}]},' +
    '{"msgs":[{"f":"them","text":"Lời thoại chốt hạ yêu cầu chuyển tiền/OTP..."}],"q":"Quyết định cuối cùng của bạn?","opts":[{"t":"...","say":"...","pts":45,"fb":"..."},{"t":"...","say":"...","pts":0,"fb":"..."},{"t":"...","say":"...","pts":0,"fb":"..."},{"t":"...","say":"...","pts":15,"fb":"..."}]}' +
    ']}\n\nTiêu đề bài báo: ' + title + '\nNội dung bài báo:\n' + bodyText;

  var aiRes = callGeminiJSON(prompt, 0.25);
  if (aiRes.ok && aiRes.data) {
    var scen = normalizeScenarioSteps(aiRes.data);
    return { ok: true, scenario: scen, title: title, content: bodyText, imageUrl: imageUrl, model: aiRes.model };
  }
  return { ok: true, fallback: true, aiError: aiRes.message, title: title, content: bodyText, imageUrl: imageUrl };
}

// ==================== TỰ ĐỘNG QUÉT TIN TỨC THẬT (RSS + CỔNG TTĐT ĐỒNG NAI) ====================
function fetchDailyScamNews() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('TinTucCanhBao') || ss.insertSheet('TinTucCanhBao');
  if (sh.getLastRow() === 0) sh.appendRow(['Tiêu đề', 'Tóm tắt', 'Ảnh minh họa (og:image)', 'Link gốc', 'Nguồn báo', 'Ngày đăng']);

  var existingUrls = {};
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][3]) existingUrls[String(rows[i][3]).trim()] = true;
  }

  var keywords = /lừa đảo|tội phạm mạng|pa05|chiếm đoạt|bắt cóc|giả danh|mạo danh|vneid|deepfake|an ninh mạng|chuyển tiền|chiêu trò/i;
  var todayStr = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy');
  var totalAdded = 0;

  // Nguồn 1: RSS Pháp luật VnExpress (XML chuẩn, ổn định, không bị Cloudflare chặn)
  try {
    var xmlText = UrlFetchApp.fetch('https://vnexpress.net/rss/phap-luat.rss', { muteHttpExceptions: true }).getContentText();
    var doc = XmlService.parse(xmlText);
    var items = doc.getRootElement().getChild('channel').getChildren('item');
    for (var k = 0; k < items.length && totalAdded < 6; k++) {
      var itemTitle = items[k].getChildText('title') || '';
      var itemLink = (items[k].getChildText('link') || '').trim();
      var itemDescRaw = items[k].getChildText('description') || '';
      var imgMatch = itemDescRaw.match(/src=["']([^"']+)["']/i);
      var itemImg = imgMatch ? imgMatch[1] : '';
      var itemSummary = itemDescRaw.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
      if (itemLink && !existingUrls[itemLink] && keywords.test(itemTitle + ' ' + itemSummary)) {
        existingUrls[itemLink] = true;
        sh.appendRow([itemTitle, itemSummary.slice(0, 260), itemImg, itemLink, 'VnExpress Pháp luật', todayStr]);
        totalAdded++;
      }
    }
  } catch (e) {}

  // Nguồn 2: Cổng TTĐT Công an Đồng Nai & Báo Đồng Nai
  var webSources = [
    { name: 'Công an Đồng Nai', url: 'https://congan.dongnai.gov.vn' },
    { name: 'Báo Đồng Nai', url: 'https://baodongnai.com.vn/phap-luat/' }
  ];
  webSources.forEach(function(src) {
    try {
      var html = UrlFetchApp.fetch(src.url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } }).getContentText();
      var linkRegex = /<a[^>]+href=["'](https?:\/\/[^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
      var match, added = 0;
      while ((match = linkRegex.exec(html)) !== null && added < 3) {
        var href = match[1].trim();
        var linkText = match[2].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
        if (linkText.length > 25 && keywords.test(linkText) && !existingUrls[href]) {
          existingUrls[href] = true;
          var imgUrl = '';
          try {
            var artHtml = UrlFetchApp.fetch(href, { muteHttpExceptions: true }).getContentText();
            var mImg = artHtml.match(/<meta[^>]*property=["']og:image["'][^>]*content=["']([^"']+)["']/i);
            if (mImg) imgUrl = mImg[1];
          } catch(err){}
          sh.appendRow([linkText, 'Cảnh báo thủ đoạn tội phạm mạng mới nhất từ ' + src.name, imgUrl, href, src.name, todayStr]);
          added++;
          totalAdded++;
        }
      }
    } catch(e) {}
  });

  return totalAdded;
}

function fetchArticleFullContent(url) {
  if (!url) return { ok: false, message: 'Thiếu URL' };
  try {
    var resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; DongNaiCyberShield/1.0)' }
    });
    var html = resp.getContentText();
    var mTitle = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
    var title = mTitle ? mTitle[1].replace(/\s+/g, ' ').trim() : '';
    var mImg = html.match(/<meta[^>]*property=["']og:image["'][^>]*content=["']([^"']+)["']/i);
    var img = mImg ? mImg[1] : '';
    var cleaned = html.replace(/<script[\s\S]*?<\/script>/gi, ' ').replace(/<style[\s\S]*?<\/style>/gi, ' ');
    var pMatches = cleaned.match(/<p[^>]*>([\s\S]*?)<\/p>/gi) || [];
    var paragraphs = [];
    for (var i = 0; i < pMatches.length && paragraphs.length < 18; i++) {
      var txt = pMatches[i].replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
      if (txt.length > 35 && !/copyright|vnexpress|đăng nhập|bình luận/i.test(txt.slice(0, 30))) {
        paragraphs.push(txt);
      }
    }
    return { ok: true, title: title, imageUrl: img, paragraphs: paragraphs };
  } catch (e) {
    return { ok: false, message: String(e) };
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
