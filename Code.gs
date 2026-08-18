// ==========================================
// CẤU HÌNH GOOGLE SHEETS - APP NHÂN SỰ ĐI XANH
// Phiên bản chỉ sử dụng đúng 4 sheet:
//   MAIN, PHAN_QUYEN, HOP_DONG_MOI, PHU_LUC_HOP_DONG
// ==========================================
// Trong Apps Script, tạo Script Property:
//   Key: SPREADSHEET_ID
//   Value: ID của file Google Sheets cần kết nối.
const SPREADSHEET_ID_PROPERTY = 'SPREADSHEET_ID';

const SHEETS = Object.freeze({
  MAIN: 'MAIN',
  PHAN_QUYEN: 'PHAN_QUYEN',
  HOP_DONG_MOI: 'HOP_DONG_MOI',
  PHU_LUC_HOP_DONG: 'PHU_LUC_HOP_DONG'
});

// Những cột này luôn lấy từ MAIN, bảng hợp đồng/phụ lục không được ghi đè.
const PROTECTED_MAIN_FIELDS = Object.freeze([
  'id', 'ID', 'ID_NV', 'HO_TEN_MA_NV', 'id_main', 'idNhanSu'
]);

// ==========================================
// API ENDPOINT
// GET:
// /exec?id=<MAIN.ID_NV>&hopDongId=<HOP_DONG_MOI.id>
//      &template=<ten_template>&username=<user>&password=<pass>
//
// Nên ưu tiên POST để không đưa mật khẩu lên URL.
// ==========================================
function doGet(e) {
  return handleRequest_((e && e.parameter) || {});
}

function doPost(e) {
  let params = (e && e.parameter) || {};

  try {
    if (e && e.postData && e.postData.contents) {
      const body = JSON.parse(e.postData.contents);
      params = Object.assign({}, params, body || {});
    }
  } catch (error) {
    return jsonOutput_({
      success: false,
      message: 'Nội dung POST không phải JSON hợp lệ: ' + error.message
    });
  }

  return handleRequest_(params);
}

function handleRequest_(params) {
  try {
    params = params || {};

    const id = clean_(params.id);
    const requestedHopDongId = clean_(params.hopDongId);
    const template = clean_(params.template) || 'ho_so_nhan_su';
    const username = clean_(params.username || params.user);
    const password = clean_(params.password || params.pass);

    if (!username || !password) {
      return jsonOutput_({
        success: false,
        auth_failed: true,
        message: 'Yêu cầu đăng nhập.'
      });
    }

    if (!id) {
      return jsonOutput_({
        success: false,
        message: "Thiếu tham số 'id'. Giá trị này phải là MAIN.ID_NV."
      });
    }

    const ss = openConfiguredSpreadsheet_();
    const auth = authenticate_(ss, username, password);

    if (!auth.valid) {
      return jsonOutput_({
        success: false,
        auth_failed: true,
        message: auth.message
      });
    }

    const main = findOneInSheetByField_(ss, SHEETS.MAIN, 'ID_NV', id);

    if (!main) {
      return jsonOutput_({
        success: false,
        message: 'Không tìm thấy nhân sự có MAIN.ID_NV: ' + id
      });
    }

    const related = loadRelatedData_(ss, main, requestedHopDongId);
    const mergedData = buildEmployeeData_(main, related);

    const response = {
      success: true,
      template: template,
      id: id,
      hopDongId: clean_(getFieldValue_(related.hopDongHienTai, 'id')),
      requested_hop_dong_id: requestedHopDongId,
      key_source: 'MAIN.ID_NV',
      data: mergedData
    };

    const warnings = buildWarnings_(main, related);
    if (warnings.length) response.warnings = warnings;

    return jsonOutput_(response);
  } catch (error) {
    console.error(error);

    if (error && error.apiCode) {
      return jsonOutput_({
        success: false,
        error_code: error.apiCode,
        message: error.message
      });
    }

    return jsonOutput_({
      success: false,
      message: 'Lỗi hệ thống: ' + error.message
    });
  }
}

// ==========================================
// XÁC THỰC TRÊN SHEET PHAN_QUYEN
// Các cột bắt buộc: user, pass
// Cột tùy chọn: trangThai
// ==========================================
function authenticate_(ss, username, password) {
  const inputUser = clean_(username);
  const inputPass = clean_(password);
  const accounts = findRowsInSheetByField_(
    ss,
    SHEETS.PHAN_QUYEN,
    'user',
    inputUser
  );

  for (let i = 0; i < accounts.length; i++) {
    const account = accounts[i];
    const sheetUser = clean_(getFieldValue_(account, 'user'));
    const sheetPass = clean_(getFieldValue_(account, 'pass'));
    const status = clean_(
      getFieldValue_(account, ['trangThai', 'trang_thai'])
    ).toLowerCase();

    if (sheetUser === inputUser && sheetPass === inputPass) {
      if (status && status !== 'còn hoạt động') {
        return {
          valid: false,
          message: 'Tài khoản đã ngừng hoạt động.'
        };
      }

      return {
        valid: true,
        account: account
      };
    }
  }

  return {
    valid: false,
    message: 'Sai tên đăng nhập hoặc mật khẩu!'
  };
}

// ==========================================
// ĐỌC VÀ LIÊN KẾT DỮ LIỆU
// Liên kết dữ liệu:
//   MAIN.ID_NV = HOP_DONG_MOI.ID_MAIN
//   MAIN.ID_NV = PHU_LUC_HOP_DONG.PLHD_CUA_NV
// ==========================================
function loadRelatedData_(ss, main, requestedHopDongId) {
  const mainId = clean_(getFieldValue_(main, ['ID_NV', 'id_nv']));

  const allHopDongList = findRowsInSheetByField_(
    ss,
    SHEETS.HOP_DONG_MOI,
    'id_main',
    mainId
  );

  const hopDongList = allHopDongList.filter(isUsableRow_);
  const requestedHopDong = validateRequestedContract_(
    allHopDongList,
    requestedHopDongId
  );

  const hopDongHienTai = requestedHopDong ||
    selectCurrentContract_(
      hopDongList,
      getFieldValue_(main, ['idHopDong', 'id_hop_dong'])
    );

  const phuLucHopDongList = findRowsInSheetByField_(
    ss,
    SHEETS.PHU_LUC_HOP_DONG,
    'PLHD_CUA_NV',
    mainId
  ).filter(isUsableRow_);

  return {
    hopDongList: sortRowsByDateDesc_(hopDongList, 'ngayBatDau'),
    hopDongHienTai: hopDongHienTai,
    phuLucHopDongList: sortRowsByDateDesc_(
      phuLucHopDongList,
      'ngayHieuLucPLHD'
    ),
    phuLucHopDongMoiNhat: selectLatestByDate_(
      phuLucHopDongList,
      'ngayHieuLucPLHD'
    )
  };
}

function validateRequestedContract_(allHopDongList, requestedHopDongId) {
  const contractId = clean_(requestedHopDongId);
  if (!contractId) return null;

  const contract = findByField_(allHopDongList, 'id', contractId);

  if (!contract) {
    throw apiError_(
      'HOP_DONG_KHONG_HOP_LE',
      'Hợp đồng được chọn không tồn tại hoặc không thuộc nhân sự này.'
    );
  }

  if (!isUsableRow_(contract)) {
    const contractLabel = clean_(
      getFieldValue_(contract, ['soHD', 'so_hd'])
    ) || contractId;
    throw apiError_(
      'HOP_DONG_DA_XOA',
      'Hợp đồng ' + contractLabel + ' đã bị xóa, không thể xuất biểu mẫu.'
    );
  }

  return contract;
}

function selectCurrentContract_(rows, idHopDongFromMain) {
  if (!rows || !rows.length) return {};

  const linkedId = clean_(idHopDongFromMain);
  if (linkedId) {
    const linked = findByField_(rows, 'id', linkedId);
    if (linked) return linked;
  }

  const activeRows = rows.filter(function(row) {
    return clean_(
      getFieldValue_(row, ['trangThai', 'trang_thai'])
    ).toLowerCase() === 'còn hiệu lực';
  });

  return selectLatestByDate_(
    activeRows.length ? activeRows : rows,
    'ngayBatDau'
  );
}

// ==========================================
// GỘP DỮ LIỆU DÙNG CHO DOCXTEMPLATER
// Nguồn dữ liệu chỉ gồm MAIN, HOP_DONG_MOI và PHU_LUC_HOP_DONG.
// ==========================================
function buildEmployeeData_(main, related) {
  const data = Object.assign({}, main);

  const mainId = clean_(getFieldValue_(main, ['ID_NV', 'id_nv']));
  data.id = mainId;
  data.id_main = mainId;
  data.idNhanSu = mainId;

  // Đưa field hợp đồng và phụ lục mới nhất ra cấp ngoài.
  mergeNonBlank_(data, related.hopDongHienTai, PROTECTED_MAIN_FIELDS);
  mergeNonBlank_(data, related.phuLucHopDongMoiNhat, PROTECTED_MAIN_FIELDS);

  // Sinh thêm field có tiền tố để template phân biệt nguồn dữ liệu.
  addPrefixedFields_(data, 'main_', main);
  addPrefixedFields_(data, 'hopDong_', related.hopDongHienTai);
  addPrefixedFields_(data, 'phuLucHopDong_', related.phuLucHopDongMoiNhat);

  // Các mảng dùng cho vòng lặp Docxtemplater.
  data.danhSachHopDong = related.hopDongList;
  data.danhSachPhuLucHopDong = related.phuLucHopDongList;

  data.soLuongHopDong = related.hopDongList.length;
  data.soLuongPhuLucHopDong = related.phuLucHopDongList.length;

  return normalizeForJson_(data);
}

function mergeNonBlank_(target, source, protectedFields) {
  if (!source) return target;

  const protectedMap = {};
  (protectedFields || []).forEach(function(key) {
    protectedMap[normalizeFieldName_(key)] = true;
  });

  Object.keys(source).forEach(function(key) {
    if (
      !protectedMap[normalizeFieldName_(key)] &&
      !isBlank_(source[key])
    ) {
      target[key] = source[key];
    }
  });

  return target;
}

function addPrefixedFields_(target, prefix, source) {
  if (!source) return;

  Object.keys(source).forEach(function(key) {
    target[prefix + key] = source[key];
  });
}

function buildWarnings_(main, related) {
  const warnings = [];

  if (
    !related.hopDongHienTai ||
    !clean_(getFieldValue_(related.hopDongHienTai, 'id'))
  ) {
    warnings.push(
      'Không tìm thấy hợp đồng liên kết cho nhân sự ' +
      (clean_(getFieldValue_(main, ['maNV', 'ma_nv'])) ||
        clean_(getFieldValue_(main, ['ID_NV', 'id_nv'])))
    );
  }

  return warnings;
}

// ==========================================
// HÀM ĐỌC GOOGLE SHEETS
// ==========================================
function openConfiguredSpreadsheet_() {
  const spreadsheetId = clean_(
    PropertiesService.getScriptProperties().getProperty(
      SPREADSHEET_ID_PROPERTY
    )
  );

  if (!spreadsheetId) {
    throw new Error(
      'Chưa cấu hình Script Property ' + SPREADSHEET_ID_PROPERTY
    );
  }

  return SpreadsheetApp.openById(spreadsheetId);
}

function findOneInSheetByField_(ss, sheetName, fieldName, value) {
  const rows = findRowsInSheetByField_(ss, sheetName, fieldName, value);
  return rows[0] || null;
}

function findRowsInSheetByField_(ss, sheetName, fieldName, value) {
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Không tìm thấy sheet: ' + sheetName);
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(function(header) {
      return clean_(header);
    });

  const normalizedFieldName = normalizeFieldName_(fieldName);
  const fieldIndex = headers.findIndex(function(header) {
    return normalizeFieldName_(header) === normalizedFieldName;
  });
  if (fieldIndex === -1) {
    throw new Error('Sheet ' + sheetName + ' không có cột ' + fieldName);
  }

  const target = clean_(value);
  if (!target) return [];

  const keyValues = sheet
    .getRange(2, fieldIndex + 1, lastRow - 1, 1)
    .getDisplayValues();

  const matchedRowNumbers = [];
  keyValues.forEach(function(row, index) {
    if (clean_(row[0]) === target) {
      matchedRowNumbers.push(index + 2);
    }
  });

  return matchedRowNumbers.map(function(rowNumber) {
    // getDisplayValues giữ nguyên cách ngày, giờ và tiền đang hiển thị.
    const row = sheet
      .getRange(rowNumber, 1, 1, lastColumn)
      .getDisplayValues()[0];

    const obj = {};
    headers.forEach(function(header, index) {
      if (header) obj[header] = row[index];
    });

    return obj;
  });
}

function findByField_(rows, fieldName, value) {
  const target = clean_(value);
  if (!target) return null;

  for (let i = 0; i < (rows || []).length; i++) {
    if (clean_(getFieldValue_(rows[i], fieldName)) === target) return rows[i];
  }

  return null;
}

function isUsableRow_(row) {
  const deletedState = clean_(
    getFieldValue_(row, [
      'trang_thai',
      'trang_thai_xoa',
      'trangThaiXoa',
      'xoa_row'
    ])
  ).toLowerCase();

  return deletedState !== 'delete' &&
    deletedState !== 'đã xóa' &&
    deletedState !== 'đã xoá';
}

function apiError_(code, message) {
  const error = new Error(message);
  error.apiCode = code;
  return error;
}

// ==========================================
// NGÀY THÁNG VÀ JSON
// ==========================================
function selectLatestByDate_(rows, fieldName) {
  if (!rows || !rows.length) return {};
  return sortRowsByDateDesc_(rows, fieldName)[0] || {};
}

function sortRowsByDateDesc_(rows, fieldName) {
  return (rows || []).slice().sort(function(a, b) {
    return parseDateValue_(getFieldValue_(b, fieldName)) -
      parseDateValue_(getFieldValue_(a, fieldName));
  });
}

function parseDateValue_(value) {
  if (value instanceof Date) return value.getTime();

  const text = clean_(value);
  let match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  if (match) {
    return new Date(
      Number(match[3]),
      Number(match[2]) - 1,
      Number(match[1])
    ).getTime();
  }

  match = text.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
  if (match) {
    return new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3])
    ).getTime();
  }

  return 0;
}

function normalizeForJson_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      'Asia/Ho_Chi_Minh',
      'dd/MM/yyyy'
    );
  }

  if (Array.isArray(value)) {
    return value.map(normalizeForJson_);
  }

  if (value && typeof value === 'object') {
    const output = {};
    Object.keys(value).forEach(function(key) {
      output[key] = normalizeForJson_(value[key]);
    });
    return output;
  }

  return value;
}

function clean_(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .normalize('NFC')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ')
    .trim();
}

/**
 * Chuẩn hóa tên cột để các dạng USER/user, ID_MAIN/id_main,
 * TRANG_THAI/trangThai đều được xem là cùng một field.
 */
function normalizeFieldName_(value) {
  return clean_(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

/**
 * Lấy giá trị object theo tên cột nhưng không phân biệt hoa/thường,
 * dấu cách, dấu gạch dưới hoặc kiểu camelCase.
 */
function getFieldValue_(row, fieldNames) {
  if (!row || typeof row !== 'object') return undefined;

  const names = Array.isArray(fieldNames) ? fieldNames : [fieldNames];
  const normalizedKeys = {};

  Object.keys(row).forEach(function(key) {
    const normalizedKey = normalizeFieldName_(key);
    if (!Object.prototype.hasOwnProperty.call(normalizedKeys, normalizedKey)) {
      normalizedKeys[normalizedKey] = key;
    }
  });

  for (let i = 0; i < names.length; i++) {
    const matchedKey = normalizedKeys[normalizeFieldName_(names[i])];
    if (matchedKey !== undefined) return row[matchedKey];
  }

  return undefined;
}

function isBlank_(value) {
  return value === null || value === undefined || clean_(value) === '';
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// HÀM TEST - CHỈ ĐỌC DỮ LIỆU, KHÔNG GHI/XÓA SHEET
// ==========================================

/**
 * Kiểm tra 4 sheet và các cột tối thiểu mà API cần sử dụng.
 */
function testCauHinh4Sheet() {
  const ss = openConfiguredSpreadsheet_();
  const requirements = {};
  requirements[SHEETS.MAIN] = ['ID_NV'];
  requirements[SHEETS.PHAN_QUYEN] = ['user', 'pass'];
  requirements[SHEETS.HOP_DONG_MOI] = ['id', 'id_main'];
  requirements[SHEETS.PHU_LUC_HOP_DONG] = ['PLHD_CUA_NV'];

  const result = [];

  Object.keys(requirements).forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      result.push({
        sheet: sheetName,
        success: false,
        message: 'Không tìm thấy sheet'
      });
      return;
    }

    const lastColumn = sheet.getLastColumn();
    const headers = lastColumn > 0
      ? sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(clean_)
      : [];

    const missingColumns = requirements[sheetName].filter(function(column) {
      const normalizedColumn = normalizeFieldName_(column);
      return !headers.some(function(header) {
        return normalizeFieldName_(header) === normalizedColumn;
      });
    });

    result.push({
      sheet: sheetName,
      success: missingColumns.length === 0,
      missingColumns: missingColumns
    });
  });

  const success = result.every(function(item) {
    return item.success;
  });

  const summary = {
    success: success,
    spreadsheetIdConfigured: true,
    sheets: result
  };

  Logger.log(JSON.stringify(summary, null, 2));

  if (!success) {
    throw new Error('Cấu hình 4 sheet chưa hợp lệ. Xem Execution log.');
  }

  return summary;
}

/**
 * Smoke test API bằng dữ liệu thật.
 * Tạo các Script Properties trước khi chạy:
 *   TEST_MAIN_ID       : MAIN.ID_NV dùng để test
 *   TEST_USERNAME      : PHAN_QUYEN.user dùng để test
 *   TEST_PASSWORD      : PHAN_QUYEN.pass dùng để test
 *   TEST_HOP_DONG_ID   : không bắt buộc
 *   TEST_TEMPLATE      : không bắt buộc
 */
function testApiThucTe() {
  const properties = PropertiesService.getScriptProperties();
  const params = {
    id: clean_(properties.getProperty('TEST_MAIN_ID')),
    username: clean_(properties.getProperty('TEST_USERNAME')),
    password: clean_(properties.getProperty('TEST_PASSWORD')),
    hopDongId: clean_(properties.getProperty('TEST_HOP_DONG_ID')),
    template: clean_(properties.getProperty('TEST_TEMPLATE')) ||
      'ho_so_nhan_su'
  };

  const missing = [];
  if (!params.id) missing.push('TEST_MAIN_ID');
  if (!params.username) missing.push('TEST_USERNAME');
  if (!params.password) missing.push('TEST_PASSWORD');

  if (missing.length) {
    throw new Error(
      'Thiếu Script Properties phục vụ test: ' + missing.join(', ')
    );
  }

  const output = handleRequest_(params);
  const response = JSON.parse(output.getContent());

  // Không ghi username/password vào log.
  Logger.log(JSON.stringify(response, null, 2));

  if (!response.success) {
    throw new Error('API test thất bại: ' + response.message);
  }

  return response;
}

/**
 * Test nhánh chặn request thiếu thông tin đăng nhập.
 * Hàm này không cần kết nối Google Sheet.
 */
function testApiThieuDangNhap() {
  const output = handleRequest_({ id: 'TEST_ID' });
  const response = JSON.parse(output.getContent());

  const passed = response.success === false &&
    response.auth_failed === true &&
    response.message === 'Yêu cầu đăng nhập.';

  Logger.log(JSON.stringify({
    success: passed,
    response: response
  }, null, 2));

  if (!passed) {
    throw new Error('Test thiếu đăng nhập không đạt.');
  }

  return response;
}

/**
 * Chạy lần lượt các test an toàn.
 * testApiThucTe chỉ chạy khi đã có đủ TEST_* Script Properties.
 */
function testTatCa() {
  const tests = [
    { name: 'testApiThieuDangNhap', fn: testApiThieuDangNhap },
    { name: 'testCauHinh4Sheet', fn: testCauHinh4Sheet }
  ];

  const properties = PropertiesService.getScriptProperties();
  if (
    clean_(properties.getProperty('TEST_MAIN_ID')) &&
    clean_(properties.getProperty('TEST_USERNAME')) &&
    clean_(properties.getProperty('TEST_PASSWORD'))
  ) {
    tests.push({ name: 'testApiThucTe', fn: testApiThucTe });
  }

  const results = [];

  tests.forEach(function(test) {
    try {
      test.fn();
      results.push({ name: test.name, success: true });
    } catch (error) {
      results.push({
        name: test.name,
        success: false,
        message: error.message
      });
    }
  });

  const failed = results.filter(function(item) {
    return !item.success;
  });

  const summary = {
    success: failed.length === 0,
    total: results.length,
    passed: results.length - failed.length,
    failed: failed.length,
    results: results
  };

  Logger.log(JSON.stringify(summary, null, 2));

  if (failed.length) {
    throw new Error(
      'Có ' + failed.length + ' test thất bại. Xem Execution log.'
    );
  }

  return summary;
}
