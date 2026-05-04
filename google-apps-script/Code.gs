const APP_VERSION = 1;

const TABLES = {
  products: [
    "id",
    "code",
    "name",
    "category",
    "size",
    "unit",
    "piecesPerBox",
    "sqmPerBox",
    "importPrice",
    "salePrice",
    "supplierId",
    "minStock",
    "isActive",
    "note",
    "createdAt",
    "imageUrl",
  ],
  customers: ["id", "name", "phone", "address", "customerGroup", "note", "createdAt"],
  suppliers: ["id", "name", "phone", "address", "contactPerson", "note", "createdAt"],
  purchase_orders: [
    "id",
    "code",
    "supplierId",
    "orderDate",
    "productId",
    "quantity",
    "unitPrice",
    "totalAmount",
    "paidAmount",
    "debtAmount",
    "paymentMethod",
    "status",
    "note",
    "createdAt",
  ],
  sales_orders: [
    "id",
    "code",
    "customerId",
    "orderDate",
    "productId",
    "quantity",
    "unitPrice",
    "discountAmount",
    "shippingFee",
    "subtotalAmount",
    "totalAmount",
    "paidAmount",
    "debtAmount",
    "paymentMethod",
    "deliveryStatus",
    "paymentStatus",
    "status",
    "note",
    "createdAt",
  ],
  stock_movements: [
    "id",
    "productId",
    "movementDate",
    "quantity",
    "movementType",
    "referenceType",
    "referenceId",
    "note",
    "createdAt",
  ],
  payments: [
    "id",
    "paymentDate",
    "direction",
    "partyType",
    "customerId",
    "supplierId",
    "referenceType",
    "referenceId",
    "amount",
    "method",
    "note",
    "createdAt",
  ],
  expenses: ["id", "category", "amount", "expenseDate", "paymentMethod", "note", "createdAt"],
  audit_logs: ["id", "time", "user", "action", "tableName", "recordCode", "note"],
  settings: ["key", "value", "updatedAt"],
};

const BACKUP_TABLES = {
  products: "products",
  customers: "customers",
  suppliers: "suppliers",
  purchases: "purchase_orders",
  sales: "sales_orders",
  stockMovements: "stock_movements",
  payments: "payments",
  expenses: "expenses",
  auditLogs: "audit_logs",
};

const DRIVE_BACKUP_FOLDER_NAME = "KhoGachKonTum-Backups";
const DRIVE_BACKUP_LATEST_FILE = "kho-gach-backup-latest.json";
const DRIVE_BACKUP_FILE_PREFIX = "kho-gach-backup-";
const DRIVE_BACKUP_KEEP_FILES = 100;
const DRIVE_PRODUCT_IMAGES_FOLDER_NAME = "KhoGachKonTum-ProductImages";

const NUMBER_FIELDS = new Set([
  "piecesPerBox",
  "sqmPerBox",
  "importPrice",
  "salePrice",
  "minStock",
  "quantity",
  "unitPrice",
  "totalAmount",
  "paidAmount",
  "debtAmount",
  "discountAmount",
  "shippingFee",
  "subtotalAmount",
  "amount",
]);

const TEXT_FIELDS = new Set([
  "phone",
]);

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABLES).forEach((tableName) => ensureSheet_(ss, tableName));
  seedIfEmpty_();
  getRevision_();
}

function setAppKey(appKey) {
  PropertiesService.getScriptProperties().setProperty("APP_KEY", String(appKey || ""));
}

function setupDriveBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === "backupNowToDrive") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("backupNowToDrive").timeBased().everyMinutes(10).create();
  backupNowToDrive(true);
  return "Đã bật backup Google Drive mỗi 10 phút.";
}

function stopDriveBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === "backupNowToDrive") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  return "Đã tắt backup Google Drive tự động.";
}

function backupNowToDrive(force) {
  const data = snapshot_();
  const dataHash = makeHash_(JSON.stringify(data));
  const previousHash = getSetting_("drive_backup_hash");

  if (!force && previousHash === dataHash) {
    return "Dữ liệu chưa thay đổi, không cần tạo backup mới.";
  }

  const exportedAt = now_();
  const backup = {
    app: "tile-warehouse-app",
    version: APP_VERSION,
    exportedAt,
    data,
  };
  const content = JSON.stringify(backup, null, 2);
  const folder = getDriveBackupFolder_();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  const filename = `${DRIVE_BACKUP_FILE_PREFIX}${timestamp}.json`;

  upsertDriveTextFile_(folder, DRIVE_BACKUP_LATEST_FILE, content);
  folder.createFile(filename, content, "application/json");
  pruneDriveBackups_(folder);
  setSetting_("drive_backup_hash", dataHash);
  setSetting_("drive_backup_at", exportedAt);

  return `Đã backup Google Drive: ${filename}`;
}

function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

function onEdit(e) {
  const range = e && e.range;
  const sheet = range && range.getSheet();
  const sheetName = sheet && sheet.getName();

  if (sheetName && TABLES[sheetName] && sheetName !== "settings") {
    touchRevision_();
  }
}

function handleRequest_(e) {
  const params = (e && e.parameter) || {};
  const callback = params.callback || "";

  try {
    authorize_(params);
    const action = params.action || "snapshot";

    if (action === "snapshot") {
      return respond_(callback, {
        ok: true,
        version: APP_VERSION,
        revision: getRevision_(),
        data: snapshot_(),
      });
    }

    if (action === "meta") {
      return respond_(callback, {
        ok: true,
        version: APP_VERSION,
        revision: getRevision_(),
      });
    }

    if (action === "mutate") {
      const type = params.type;
      const payload = JSON.parse(params.payload || "{}");
      const result = mutate_(type, payload);
      return respond_(callback, result);
    }

    throw new Error("Action không hợp lệ.");
  } catch (error) {
    return respond_(callback, {
      ok: false,
      error: error && error.message ? error.message : String(error),
    });
  }
}

function authorize_(params) {
  const expectedKey = PropertiesService.getScriptProperties().getProperty("APP_KEY");

  if (!expectedKey) {
    return;
  }

  if (params.key !== expectedKey) {
    throw new Error("Sai app key.");
  }
}

function respond_(callback, payload) {
  const json = JSON.stringify(payload);

  if (callback) {
    return ContentService.createTextOutput(`${callback}(${json});`).setMimeType(
      ContentService.MimeType.JAVASCRIPT,
    );
  }

  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function mutate_(type, payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (type === "addProduct") addProduct_(payload);
    else if (type === "updateProductImage") updateProductImage_(payload);
    else if (type === "addCustomer") addCustomer_(payload);
    else if (type === "addSupplier") addSupplier_(payload);
    else if (type === "addPurchase") addPurchase_(payload);
    else if (type === "addSale") addSale_(payload);
    else if (type === "recordCustomerPayment") recordCustomerPayment_(payload);
    else if (type === "recordSupplierPayment") recordSupplierPayment_(payload);
    else if (type === "addExpense") addExpense_(payload);
    else if (type === "restoreBackup") restoreBackup_(payload);
    else throw new Error("Loại thao tác không hợp lệ.");

    const revision = touchRevision_();
    SpreadsheetApp.flush();
    return { ok: true, version: APP_VERSION, revision, data: snapshot_() };
  } finally {
    lock.releaseLock();
  }
}

function snapshot_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABLES).forEach((tableName) => ensureSheet_(ss, tableName));

  return {
    products: readTable_("products"),
    customers: readTable_("customers"),
    suppliers: readTable_("suppliers"),
    purchases: readTable_("purchase_orders"),
    sales: readTable_("sales_orders"),
    stockMovements: readTable_("stock_movements"),
    payments: readTable_("payments"),
    expenses: readTable_("expenses"),
    auditLogs: readTable_("audit_logs"),
  };
}

function restoreBackup_(payload) {
  const data = payload && payload.data;

  if (!data || typeof data !== "object") {
    throw new Error("File backup không có dữ liệu để phục hồi.");
  }

  Object.keys(BACKUP_TABLES).forEach((dataKey) => {
    const tableName = BACKUP_TABLES[dataKey];
    const rows = Array.isArray(data[dataKey]) ? data[dataKey] : [];
    writeTable_(tableName, rows);
  });

  appendRow_("audit_logs", {
    id: makeId_("audit"),
    time: now_(),
    user: "Google Sheet API",
    action: "restore",
    tableName: "all",
    recordCode: "backup",
    note: `Phục hồi dữ liệu từ backup ${text_(payload.exportedAt) || ""}`.trim(),
  });
}

function ensureSheet_(ss, tableName) {
  let sheet = ss.getSheetByName(tableName);

  if (!sheet) {
    sheet = ss.insertSheet(tableName);
  }

  const headers = TABLES[tableName];
  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeader = headers.some((header, index) => currentHeaders[index] !== header);

  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  applyTextColumnFormats_(sheet, headers);

  return sheet;
}

function applyTextColumnFormats_(sheet, headers) {
  headers.forEach((header, index) => {
    if (TEXT_FIELDS.has(header)) {
      const column = index + 1;
      sheet.getRange(1, column, sheet.getMaxRows(), 1).setNumberFormat("@");

      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const range = sheet.getRange(2, column, lastRow - 1, 1);
        const values = range.getValues();
        const fixedValues = values.map((row) => [normalizeTextValue_(header, row[0])]);
        const changed = fixedValues.some((row, rowIndex) => row[0] !== String(values[rowIndex][0] || "").trim());

        if (changed) {
          range.setValues(fixedValues);
        }
      }
    }
  });
}

function readTable_(tableName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheet_(ss, tableName);
  const lastRow = sheet.getLastRow();
  const headers = TABLES[tableName];

  if (lastRow < 2) {
    return [];
  }

  return sheet
    .getRange(2, 1, lastRow - 1, headers.length)
    .getValues()
    .filter((row) => row.some((cell) => cell !== ""))
    .map((row) => rowToObject_(headers, row));
}

function writeTable_(tableName, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheet_(ss, tableName);
  const headers = TABLES[tableName];
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows.map((row) => objectToRow_(headers, row)));
    applyTextColumnFormats_(sheet, headers);
  }
}

function getSetting_(key) {
  const rows = readTable_("settings");
  const row = rows.find((item) => String(item.key) === String(key));
  return row ? String(row.value || "") : "";
}

function setSetting_(key, value) {
  const rows = readTable_("settings");
  const index = rows.findIndex((item) => String(item.key) === String(key));
  const nextRow = {
    key: String(key),
    value: String(value),
    updatedAt: now_(),
  };

  if (index >= 0) {
    rows[index] = nextRow;
  } else {
    rows.push(nextRow);
  }

  writeTable_("settings", rows);
}

function getRevision_() {
  const revision = getSetting_("sheet_revision");

  if (revision) {
    return revision;
  }

  return touchRevision_();
}

function touchRevision_() {
  const revision = `${Date.now()}-${Utilities.getUuid()}`;
  setSetting_("sheet_revision", revision);
  return revision;
}

function getDriveBackupFolder_() {
  const folders = DriveApp.getFoldersByName(DRIVE_BACKUP_FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(DRIVE_BACKUP_FOLDER_NAME);
}

function getProductImagesFolder_() {
  const folders = DriveApp.getFoldersByName(DRIVE_PRODUCT_IMAGES_FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  const folder = DriveApp.createFolder(DRIVE_PRODUCT_IMAGES_FOLDER_NAME);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return folder;
}

function saveProductImage_(payload, productCode) {
  const imageData = text_(payload.imageData);

  if (!imageData) {
    return text_(payload.imageUrl);
  }

  const match = imageData.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);

  if (!match) {
    throw new Error("Dữ liệu ảnh không hợp lệ.");
  }

  const mimeType = match[1] || "image/jpeg";
  const bytes = Utilities.base64Decode(match[2]);

  if (bytes.length > 600 * 1024) {
    throw new Error("Ảnh sau khi nén vẫn quá lớn. Hãy chọn ảnh nhỏ hơn.");
  }

  const extension = mimeType.indexOf("png") >= 0 ? "png" : "jpg";
  const safeName = sanitizeFileName_(text_(payload.imageFileName) || productCode || "anh-san-pham");
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  const blob = Utilities.newBlob(bytes, mimeType, `${safeName}-${timestamp}.${extension}`);
  const file = getProductImagesFolder_().createFile(blob);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w800`;
}

function sanitizeFileName_(value) {
  return String(value || "anh-san-pham")
    .replace(/\.[^.]+$/, "")
    .replace(/[^a-zA-Z0-9._-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80) || "anh-san-pham";
}

function extractDriveFileId_(url) {
  const text = String(url || "");
  const fileMatch = text.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  const idMatch = text.match(/[?&]id=([a-zA-Z0-9_-]+)/);

  return fileMatch ? fileMatch[1] : idMatch ? idMatch[1] : "";
}

function trashDriveFileFromUrl_(url) {
  const fileId = extractDriveFileId_(url);

  if (!fileId) {
    return;
  }

  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (error) {
    // Ignore old links the script cannot manage.
  }
}

function upsertDriveTextFile_(folder, filename, content) {
  const files = folder.getFilesByName(filename);

  if (files.hasNext()) {
    const file = files.next();
    file.setContent(content);
    return file;
  }

  return folder.createFile(filename, content, "application/json");
}

function pruneDriveBackups_(folder) {
  const files = [];
  const iterator = folder.getFiles();

  while (iterator.hasNext()) {
    const file = iterator.next();
    const name = file.getName();

    if (name.indexOf(DRIVE_BACKUP_FILE_PREFIX) === 0 && name !== DRIVE_BACKUP_LATEST_FILE) {
      files.push(file);
    }
  }

  files
    .sort((a, b) => b.getDateCreated().getTime() - a.getDateCreated().getTime())
    .slice(DRIVE_BACKUP_KEEP_FILES)
    .forEach((file) => file.setTrashed(true));
}

function makeHash_(content) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    content,
    Utilities.Charset.UTF_8,
  );

  return Utilities.base64EncodeWebSafe(digest);
}

function appendRow_(tableName, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheet_(ss, tableName);
  const headers = TABLES[tableName];
  const nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, headers.length).setValues([objectToRow_(headers, row)]);
  applyTextColumnFormats_(sheet, headers);
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((header, index) => {
    const value = row[index];
    obj[header] = NUMBER_FIELDS.has(header) ? Number(value || 0) : normalizeTextValue_(header, value);
  });
  return obj;
}

function objectToRow_(headers, obj) {
  return headers.map((header) => {
    const value = obj[header] === undefined ? "" : obj[header];

    if (TEXT_FIELDS.has(header)) {
      return normalizeTextValue_(header, value);
    }

    return value;
  });
}

function normalizeTextValue_(header, value) {
  const text = String(value || "").trim();

  if (header === "phone") {
    return normalizePhone_(text);
  }

  return text;
}

function normalizePhone_(value) {
  const text = String(value || "").trim();

  if (!text || text[0] === "+" || text[0] === "0") {
    return text;
  }

  if (/^\d{9}$/.test(text)) {
    return `0${text}`;
  }

  return text;
}

function addProduct_(payload) {
  const code = text_(payload.code);
  if (!code || !text_(payload.name)) throw new Error("Thiếu mã hàng hoặc tên hàng.");
  if (readTable_("products").some((product) => product.code === code)) throw new Error("Mã hàng đã tồn tại.");
  const id = makeId_("product");
  const imageUrl = saveProductImage_(payload, code);

  appendRow_("products", {
    id,
    code,
    name: text_(payload.name),
    category: text_(payload.category),
    size: text_(payload.size),
    unit: "thùng",
    piecesPerBox: number_(payload.piecesPerBox),
    sqmPerBox: number_(payload.sqmPerBox),
    importPrice: number_(payload.importPrice),
    salePrice: number_(payload.salePrice),
    supplierId: text_(payload.supplierId),
    minStock: number_(payload.minStock),
    isActive: true,
    note: text_(payload.note),
    createdAt: now_(),
    imageUrl,
  });
  audit_("create", "products", code, `Tạo sản phẩm ${text_(payload.name)}`);
}

function updateProductImage_(payload) {
  const productId = text_(payload.productId);
  const products = readTable_("products");
  const index = products.findIndex((product) => product.id === productId);

  if (index < 0) {
    throw new Error("Không tìm thấy sản phẩm cần cập nhật ảnh.");
  }

  const product = products[index];
  const oldImageUrl = text_(product.imageUrl);
  const imageUrl = text_(payload.clearImage) ? "" : saveProductImage_(payload, product.code);

  if (oldImageUrl && oldImageUrl !== imageUrl) {
    trashDriveFileFromUrl_(oldImageUrl);
  }

  products[index] = Object.assign({}, product, { imageUrl });
  writeTable_("products", products);
  audit_("update", "products", product.code, imageUrl ? "Cập nhật ảnh sản phẩm" : "Gỡ ảnh sản phẩm");
}

function addCustomer_(payload) {
  const name = text_(payload.name);
  if (!name) throw new Error("Thiếu tên khách hàng.");

  appendRow_("customers", {
    id: makeId_("customer"),
    name,
    phone: text_(payload.phone),
    address: text_(payload.address),
    customerGroup: text_(payload.customerGroup),
    note: text_(payload.note),
    createdAt: now_(),
  });
  audit_("create", "customers", name, `Tạo khách hàng ${name}`);
}

function addSupplier_(payload) {
  const name = text_(payload.name);
  if (!name) throw new Error("Thiếu tên nhà cung cấp.");

  appendRow_("suppliers", {
    id: makeId_("supplier"),
    name,
    phone: text_(payload.phone),
    address: text_(payload.address),
    contactPerson: text_(payload.contactPerson),
    note: text_(payload.note),
    createdAt: now_(),
  });
  audit_("create", "suppliers", name, `Tạo nhà cung cấp ${name}`);
}

function addPurchase_(payload) {
  const purchases = readTable_("purchase_orders");
  const code = nextCode_("PO", purchases.map((order) => order.code));
  const items = getOrderItems_(payload);
  const lineTotals = items.map((item) => number_(item.quantity) * number_(item.unitPrice));
  const totalAmount = lineTotals.reduce((total, value) => total + value, 0);
  const paidAmount = Math.min(number_(payload.paidAmount), totalAmount);
  const paidShares = allocateAmount_(paidAmount, lineTotals);
  const orderDate = text_(payload.orderDate) || today_();
  const supplierId = text_(payload.supplierId);
  const paymentMethod = text_(payload.paymentMethod);
  const note = text_(payload.note);
  const createdAt = now_();

  if (!supplierId) throw new Error("Thiếu nhà cung cấp.");
  if (items.length === 0) throw new Error("Thêm ít nhất một dòng sản phẩm để nhập.");

  items.forEach((item, index) => {
    const id = makeId_("purchase");
    const lineTotal = lineTotals[index] || 0;
    const linePaid = paidShares[index] || 0;

    appendRow_("purchase_orders", {
      id,
      code,
      supplierId,
      orderDate,
      productId: item.productId,
      quantity: item.quantity,
      unitPrice: item.unitPrice,
      totalAmount: lineTotal,
      paidAmount: linePaid,
      debtAmount: Math.max(lineTotal - linePaid, 0),
      paymentMethod,
      status: "confirmed",
      note,
      createdAt,
    });
    appendRow_("stock_movements", {
      id: makeId_("movement"),
      productId: item.productId,
      movementDate: orderDate,
      quantity: item.quantity,
      movementType: "purchase",
      referenceType: "purchase_order",
      referenceId: id,
      note: `Nhập ${code}`,
      createdAt,
    });
  });

  if (paidAmount > 0) {
    appendRow_("payments", {
      id: makeId_("payment"),
      paymentDate: orderDate,
      direction: "out",
      partyType: "supplier",
      customerId: "",
      supplierId,
      referenceType: "purchase_order",
      referenceId: code,
      amount: paidAmount,
      method: paymentMethod,
      note: `Trả tiền ${code}`,
      createdAt,
    });
  }

  const totalQuantity = items.reduce((total, item) => total + number_(item.quantity), 0);
  audit_("confirm", "purchase_orders", code, `Xác nhận nhập ${items.length} dòng / ${totalQuantity} thùng`);
}

function addSale_(payload) {
  const sales = readTable_("sales_orders");
  const customerId = text_(payload.customerId);
  const items = getOrderItems_(payload);

  if (!customerId) throw new Error("Thiếu khách hàng.");
  if (items.length === 0) throw new Error("Thêm ít nhất một dòng sản phẩm để bán.");

  const requirements = getProductRequirements_(items);
  Object.keys(requirements).forEach((productId) => {
    const availableStock = getAvailableStock_(productId);
    const requiredQuantity = number_(requirements[productId]);
    if (availableStock < requiredQuantity) {
      throw new Error(`Không đủ tồn kho ${getProductNameById_(productId)}. Hiện còn ${availableStock} thùng, cần ${requiredQuantity} thùng.`);
    }
  });

  const code = nextCode_("SO", sales.map((order) => order.code));
  const orderDate = text_(payload.orderDate) || today_();
  const discountAmount = number_(payload.discountAmount);
  const shippingFee = number_(payload.shippingFee);
  const lineSubtotals = items.map((item) => number_(item.quantity) * number_(item.unitPrice));
  const subtotalAmount = lineSubtotals.reduce((total, value) => total + value, 0);
  const totalAmount = Math.max(subtotalAmount - discountAmount + shippingFee, 0);
  const paidAmount = Math.min(number_(payload.paidAmount), totalAmount);
  const totalShares = allocateAmount_(totalAmount, lineSubtotals);
  const paidShares = allocateAmount_(paidAmount, totalShares);
  const discountShares = allocateAmount_(Math.min(discountAmount, subtotalAmount), lineSubtotals);
  const shippingShares = allocateAmount_(shippingFee, lineSubtotals);
  const paymentMethod = text_(payload.paymentMethod);
  const note = text_(payload.note);
  const createdAt = now_();

  items.forEach((item, index) => {
    const id = makeId_("sale");
    const lineTotal = totalShares[index] || 0;
    const linePaid = paidShares[index] || 0;
    const debtAmount = Math.max(lineTotal - linePaid, 0);

    appendRow_("sales_orders", {
      id,
      code,
      customerId,
      orderDate,
      productId: item.productId,
      quantity: item.quantity,
      unitPrice: item.unitPrice,
      discountAmount: discountShares[index] || 0,
      shippingFee: shippingShares[index] || 0,
      subtotalAmount: lineSubtotals[index] || 0,
      totalAmount: lineTotal,
      paidAmount: linePaid,
      debtAmount,
      paymentMethod,
      deliveryStatus: "pending",
      paymentStatus: debtAmount <= 0 ? "paid" : linePaid > 0 ? "partial" : "unpaid",
      status: "confirmed",
      note,
      createdAt,
    });
    appendRow_("stock_movements", {
      id: makeId_("movement"),
      productId: item.productId,
      movementDate: orderDate,
      quantity: -number_(item.quantity),
      movementType: "sale",
      referenceType: "sales_order",
      referenceId: id,
      note: `Bán ${code}`,
      createdAt,
    });
  });

  if (paidAmount > 0) {
    appendRow_("payments", {
      id: makeId_("payment"),
      paymentDate: orderDate,
      direction: "in",
      partyType: "customer",
      customerId,
      supplierId: "",
      referenceType: "sales_order",
      referenceId: code,
      amount: paidAmount,
      method: paymentMethod,
      note: `Thu tiền ${code}`,
      createdAt,
    });
  }

  const totalQuantity = items.reduce((total, item) => total + number_(item.quantity), 0);
  audit_("confirm", "sales_orders", code, `Xác nhận bán ${items.length} dòng / ${totalQuantity} thùng`);
}

function recordCustomerPayment_(payload) {
  const customerId = text_(payload.customerId);
  const amount = number_(payload.amount);
  const paymentDate = text_(payload.paymentDate) || today_();
  const sales = readTable_("sales_orders");
  const outstandingOrders = sales
    .filter((order) => order.customerId === customerId && order.status !== "canceled" && number_(order.debtAmount) > 0)
    .sort((a, b) => String(a.orderDate).localeCompare(String(b.orderDate)));
  const totalDebt = outstandingOrders.reduce((total, order) => total + number_(order.debtAmount), 0);

  if (!customerId) throw new Error("Chọn khách hàng cần trả nợ.");
  if (amount <= 0) throw new Error("Số tiền trả nợ phải lớn hơn 0.");
  if (totalDebt <= 0) throw new Error("Khách hàng này không còn công nợ.");
  if (amount > totalDebt) throw new Error(`Số tiền vượt công nợ hiện tại. Khách còn nợ ${formatVnd_(totalDebt)}.`);

  let remaining = amount;
  const affectedCodes = [];
  const updatedSales = sales.map((order) => {
    if (order.customerId !== customerId || order.status === "canceled" || number_(order.debtAmount) <= 0 || remaining <= 0) {
      return order;
    }

    const paidForOrder = Math.min(number_(order.debtAmount), remaining);
    remaining -= paidForOrder;
    affectedCodes.push(order.code);
    const nextDebt = Math.max(number_(order.debtAmount) - paidForOrder, 0);
    return Object.assign({}, order, {
      paidAmount: number_(order.paidAmount) + paidForOrder,
      debtAmount: nextDebt,
      paymentStatus: nextDebt <= 0 ? "paid" : "partial",
    });
  });

  writeTable_("sales_orders", updatedSales);
  appendRow_("payments", {
    id: makeId_("payment"),
    paymentDate,
    direction: "in",
    partyType: "customer",
    customerId,
    supplierId: "",
    referenceType: "manual",
    referenceId: affectedCodes.join(", "),
    amount,
    method: text_(payload.method) || "Tiền mặt",
    note: text_(payload.note) || `Khách trả nợ: ${affectedCodes.join(", ")}`,
    createdAt: now_(),
  });
  audit_("payment", "customer_debts", affectedCodes.join(", "), `Ghi nhận khách trả ${formatVnd_(amount)}`);
}

function recordSupplierPayment_(payload) {
  const supplierId = text_(payload.supplierId);
  const amount = number_(payload.amount);
  const paymentDate = text_(payload.paymentDate) || today_();
  const purchases = readTable_("purchase_orders");
  const outstandingOrders = purchases
    .filter((order) => order.supplierId === supplierId && order.status !== "canceled" && number_(order.debtAmount) > 0)
    .sort((a, b) => String(a.orderDate).localeCompare(String(b.orderDate)));
  const totalDebt = outstandingOrders.reduce((total, order) => total + number_(order.debtAmount), 0);

  if (!supplierId) throw new Error("Chon nha cung cap can tra no.");
  if (amount <= 0) throw new Error("So tien tra no phai lon hon 0.");
  if (totalDebt <= 0) throw new Error("Nha cung cap nay khong con cong no.");
  if (amount > totalDebt) throw new Error(`So tien vuot cong no hien tai. Nha cung cap con no ${formatVnd_(totalDebt)}.`);

  let remaining = amount;
  const affectedCodes = [];
  const updatedPurchases = purchases.map((order) => {
    if (order.supplierId !== supplierId || order.status === "canceled" || number_(order.debtAmount) <= 0 || remaining <= 0) {
      return order;
    }

    const paidForOrder = Math.min(number_(order.debtAmount), remaining);
    remaining -= paidForOrder;
    affectedCodes.push(order.code);
    return Object.assign({}, order, {
      paidAmount: number_(order.paidAmount) + paidForOrder,
      debtAmount: Math.max(number_(order.debtAmount) - paidForOrder, 0),
    });
  });

  writeTable_("purchase_orders", updatedPurchases);
  appendRow_("payments", {
    id: makeId_("payment"),
    paymentDate,
    direction: "out",
    partyType: "supplier",
    customerId: "",
    supplierId,
    referenceType: "manual",
    referenceId: affectedCodes.join(", "),
    amount,
    method: text_(payload.method) || "Tien mat",
    note: text_(payload.note) || `Tra no nha cung cap: ${affectedCodes.join(", ")}`,
    createdAt: now_(),
  });
  audit_("payment", "supplier_debts", affectedCodes.join(", "), `Ghi nhan tra NCC ${formatVnd_(amount)}`);
}

function addExpense_(payload) {
  const amount = number_(payload.amount);
  const expenseDate = text_(payload.expenseDate) || today_();
  const id = makeId_("expense");
  if (amount <= 0) throw new Error("Số tiền chi phải lớn hơn 0.");

  appendRow_("expenses", {
    id,
    category: text_(payload.category),
    amount,
    expenseDate,
    paymentMethod: text_(payload.paymentMethod),
    note: text_(payload.note),
    createdAt: now_(),
  });
  appendRow_("payments", {
    id: makeId_("payment"),
    paymentDate: expenseDate,
    direction: "out",
    partyType: "other",
    customerId: "",
    supplierId: "",
    referenceType: "expense",
    referenceId: id,
    amount,
    method: text_(payload.paymentMethod),
    note: text_(payload.note),
    createdAt: now_(),
  });
  audit_("create", "expenses", text_(payload.category), `Tạo chi phí ${text_(payload.category)}`);
}

function computeInventory_() {
  const products = readTable_("products");
  const movements = readTable_("stock_movements");
  return products.map((product) => {
    const currentStock = movements
      .filter((movement) => movement.productId === product.id)
      .reduce((total, movement) => total + number_(movement.quantity), 0);
    return Object.assign({}, product, { currentStock });
  });
}

function getAvailableStock_(productId) {
  const item = computeInventory_().find((product) => product.id === productId);
  return item ? number_(item.currentStock) : 0;
}

function audit_(action, tableName, recordCode, note) {
  appendRow_("audit_logs", {
    id: makeId_("audit"),
    time: now_(),
    user: "Google Sheet API",
    action,
    tableName,
    recordCode,
    note,
  });
}

function seedIfEmpty_() {
  if (readTable_("products").length > 0) return;

  appendRow_("suppliers", {
    id: "supplier-abc",
    name: "Nhà cung cấp Gạch ABC",
    phone: "0901000001",
    address: "KCN Tân Tạo, TP.HCM",
    contactPerson: "Anh Minh",
    note: "Nguồn hàng chính",
    createdAt: now_(),
  });
  appendRow_("customers", {
    id: "customer-a",
    name: "Khách Nguyễn Văn A",
    phone: "0912000001",
    address: "Quận 8, TP.HCM",
    customerGroup: "Bán lẻ",
    note: "Khách mẫu",
    createdAt: now_(),
  });
  appendRow_("products", {
    id: "product-6060-a01",
    code: "G6060-A01",
    name: "Gạch bóng kính 60x60 A01",
    category: "Gạch lát nền",
    size: "60x60",
    unit: "thùng",
    piecesPerBox: 4,
    sqmPerBox: 1.44,
    importPrice: 120000,
    salePrice: 150000,
    supplierId: "supplier-abc",
    minStock: 20,
    isActive: true,
    imageUrl: "",
    note: "Hàng mẫu",
    createdAt: now_(),
  });
  audit_("seed", "settings", "INIT", "Tạo dữ liệu mẫu ban đầu");
}

function makeId_(prefix) {
  return `${prefix}-${Utilities.getUuid()}`;
}

function nextCode_(prefix, codes) {
  const nextNumber =
    codes.reduce((max, code) => {
      const number = Number(String(code || "").replace(`${prefix}-`, ""));
      return isFinite(number) ? Math.max(max, number) : max;
    }, 0) + 1;
  return `${prefix}-${String(nextNumber).padStart(4, "0")}`;
}

function text_(value) {
  return String(value === undefined || value === null ? "" : value).trim();
}

function number_(value) {
  const parsed = Number(String(value === undefined || value === null ? 0 : value).replace(/[^\d.-]/g, ""));
  return isFinite(parsed) ? parsed : 0;
}

function array_(value) {
  if (Array.isArray(value)) return value;
  if (value === undefined || value === null) return [];
  return [value];
}

function getOrderItems_(payload) {
  const productIds = array_(payload.itemProductId);
  const quantities = array_(payload.itemQuantity);
  const unitPrices = array_(payload.itemUnitPrice);

  if (productIds.length > 0) {
    return productIds
      .map((productId, index) => ({
        productId: text_(productId),
        quantity: number_(quantities[index]),
        unitPrice: number_(unitPrices[index]),
      }))
      .filter((item) => item.productId && item.quantity > 0);
  }

  const legacyItem = {
    productId: text_(payload.productId),
    quantity: number_(payload.quantity),
    unitPrice: number_(payload.unitPrice),
  };

  return legacyItem.productId && legacyItem.quantity > 0 ? [legacyItem] : [];
}

function allocateAmount_(total, weights) {
  const weightTotal = weights.reduce((sum, weight) => sum + number_(weight), 0);

  if (weights.length === 0) return [];
  if (weightTotal <= 0) return weights.map(() => 0);

  let allocated = 0;
  return weights.map((weight, index) => {
    if (index === weights.length - 1) {
      return Math.max(total - allocated, 0);
    }

    const share = Math.min(Math.round((total * number_(weight)) / weightTotal), total - allocated);
    allocated += share;
    return Math.max(share, 0);
  });
}

function getProductRequirements_(items) {
  return items.reduce((result, item) => {
    result[item.productId] = number_(result[item.productId]) + number_(item.quantity);
    return result;
  }, {});
}

function getProductNameById_(productId) {
  const product = readTable_("products").find((item) => item.id === productId);
  return product ? `${product.code} - ${product.name}` : productId;
}

function now_() {
  return new Date().toISOString();
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatVnd_(value) {
  return `${Number(value || 0).toLocaleString("vi-VN")}đ`;
}
