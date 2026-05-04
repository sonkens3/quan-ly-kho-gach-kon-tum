"use client";

import { useEffect, useMemo, useState } from "react";
import {
  fetchSheetMeta,
  fetchSheetSnapshotWithMeta,
  formDataToPayload,
  isSheetConfigured,
  loadSheetConfig,
  mutateSheetWithMeta,
  postMutateSheetWithMeta,
  type SheetConfig,
} from "@/lib/google-sheets/client";
import { createBackup, createSeedLocalBackup, localDataKey, type LocalBackup } from "@/lib/local/free-mode";
import {
  computeInventory,
  getAvailableStock,
  getProductName,
  getToday,
  makeId,
  nextCode,
  sum,
} from "@/lib/local/calculations";
import type {
  AuditLog,
  Customer,
  Expense,
  Payment,
  Product,
  PurchaseOrder,
  SalesOrder,
  StockMovement,
  Supplier,
  WarehouseData,
} from "@/lib/local/types";

type ActionResult = {
  ok: boolean;
  message: string;
};

type SyncMode = "local" | "google-sheet";

const sheetCacheKey = "tile_warehouse_sheet_cache";
const sheetMetaPollMs = 15000;

type SheetCache = {
  app: "tile-warehouse-app";
  source: "google-sheet";
  version: 1;
  endpoint: string;
  revision: string;
  syncedAt: string;
  data: WarehouseData;
};

function parseBackup(raw: string | null): LocalBackup | null {
  if (!raw) {
    return null;
  }

  try {
    const parsed = JSON.parse(raw) as LocalBackup;

    if (parsed.app !== "tile-warehouse-app" || parsed.version !== 1 || !parsed.data) {
      return null;
    }

    return parsed;
  } catch {
    return null;
  }
}

function parseSheetCache(raw: string | null, config: SheetConfig): SheetCache | null {
  if (!raw) {
    return null;
  }

  try {
    const parsed = JSON.parse(raw) as Partial<SheetCache>;

    if (
      parsed.app !== "tile-warehouse-app" ||
      parsed.source !== "google-sheet" ||
      parsed.version !== 1 ||
      parsed.endpoint !== config.endpoint ||
      !parsed.data
    ) {
      return null;
    }

    return parsed as SheetCache;
  } catch {
    return null;
  }
}

function readSheetCache(config: SheetConfig) {
  return parseSheetCache(window.localStorage.getItem(sheetCacheKey), config);
}

function writeData(data: WarehouseData) {
  window.localStorage.setItem(localDataKey, JSON.stringify(createBackup(data)));
}

function writeSheetCache(config: SheetConfig, data: WarehouseData, revision: string) {
  const cache: SheetCache = {
    app: "tile-warehouse-app",
    source: "google-sheet",
    version: 1,
    endpoint: config.endpoint,
    revision,
    syncedAt: new Date().toISOString(),
    data,
  };

  window.localStorage.setItem(sheetCacheKey, JSON.stringify(cache));
  writeData(data);
}

function formatSyncTime(value: string | Date) {
  return new Date(value).toLocaleTimeString("vi-VN");
}

function makeAudit(action: string, tableName: string, recordCode: string, note: string): AuditLog {
  return {
    id: makeId("audit"),
    time: new Date().toISOString(),
    user: "Chủ kho local",
    action,
    tableName,
    recordCode,
    note,
  };
}

function toNumber(value: FormDataEntryValue | null) {
  const parsed = Number(value ?? 0);
  return Number.isFinite(parsed) ? parsed : 0;
}

function toText(value: FormDataEntryValue | null) {
  return String(value ?? "").trim();
}

function hasProductImageUpload(formData: FormData) {
  return Boolean(toText(formData.get("imageData")));
}

type OrderItemInput = {
  productId: string;
  quantity: number;
  unitPrice: number;
};

function getOrderItems(formData: FormData): OrderItemInput[] {
  const productIds = formData.getAll("itemProductId").map((value) => toText(value));
  const quantities = formData.getAll("itemQuantity").map((value) => toNumber(value));
  const unitPrices = formData.getAll("itemUnitPrice").map((value) => toNumber(value));

  if (productIds.length > 0) {
    return productIds
      .map((productId, index) => ({
        productId,
        quantity: quantities[index] ?? 0,
        unitPrice: unitPrices[index] ?? 0,
      }))
      .filter((item) => item.productId && item.quantity > 0);
  }

  return [
    {
      productId: toText(formData.get("productId")),
      quantity: toNumber(formData.get("quantity")),
      unitPrice: toNumber(formData.get("unitPrice")),
    },
  ].filter((item) => item.productId && item.quantity > 0);
}

function allocateAmount(total: number, weights: number[]) {
  const weightTotal = sum(weights);

  if (weights.length === 0) {
    return [];
  }

  if (weightTotal <= 0) {
    return weights.map(() => 0);
  }

  let allocated = 0;

  return weights.map((weight, index) => {
    if (index === weights.length - 1) {
      return Math.max(total - allocated, 0);
    }

    const share = Math.min(Math.round((total * weight) / weightTotal), total - allocated);
    allocated += share;
    return Math.max(share, 0);
  });
}

function getProductRequirements(items: OrderItemInput[]) {
  return items.reduce<Record<string, number>>((result, item) => {
    result[item.productId] = (result[item.productId] ?? 0) + item.quantity;
    return result;
  }, {});
}

export function useWarehouseStore() {
  const [data, setData] = useState<WarehouseData | null>(null);
  const [lastMessage, setLastMessage] = useState<string | null>(null);
  const [syncMode, setSyncMode] = useState<SyncMode>("local");
  const [syncStatus, setSyncStatus] = useState("Đang kiểm tra nguồn dữ liệu...");

  useEffect(() => {
    let intervalId: number | undefined;
    let canceled = false;

    void loadSheetConfig().then(async (sheetConfig) => {
      if (canceled) {
        return;
      }

      if (!isSheetConfigured(sheetConfig)) {
        const backup = parseBackup(window.localStorage.getItem(localDataKey)) ?? createSeedLocalBackup();
        window.localStorage.setItem(localDataKey, JSON.stringify(backup));
        setData(backup.data);
        setSyncMode("local");
        setSyncStatus("Chưa cấu hình Google Sheet, đang dùng dữ liệu local trong trình duyệt.");
        return;
      }

      setSyncMode("google-sheet");

      const cachedSheet = readSheetCache(sheetConfig);

      if (cachedSheet) {
        setData(cachedSheet.data);
        setSyncStatus(
          `Đang dùng cache Google Sheet lúc ${formatSyncTime(cachedSheet.syncedAt)}, đang kiểm tra thay đổi...`,
        );
      } else {
        setSyncStatus("Lần đầu tải dữ liệu từ Google Sheet...");
      }

      async function pullIfChanged(isFirstCheck: boolean) {
        const cached = readSheetCache(sheetConfig);

        try {
          const meta = await fetchSheetMeta(sheetConfig);

          if (canceled) {
            return;
          }

          if (cached?.data && cached.revision === meta.revision) {
            setData(cached.data);
            setSyncStatus(
              `Google Sheet chưa đổi, đang dùng dữ liệu đã lưu. Kiểm tra lại mỗi ${sheetMetaPollMs / 1000} giây.`,
            );
            return;
          }

          const snapshot = await fetchSheetSnapshotWithMeta(sheetConfig);

          if (canceled) {
            return;
          }

          setData(snapshot.data);
          writeSheetCache(sheetConfig, snapshot.data, snapshot.revision || meta.revision);
          setSyncStatus(
            isFirstCheck
              ? "Đã tải dữ liệu mới từ Google Sheet."
              : `Google Sheet có thay đổi, đã cập nhật lúc ${formatSyncTime(new Date())}.`,
          );
        } catch (error) {
          if (canceled) {
            return;
          }

          if (cached?.data) {
            setData(cached.data);
            setSyncStatus(
              error instanceof Error
                ? `Đang dùng cache Google Sheet. ${error.message}`
                : "Đang dùng cache Google Sheet, chưa kiểm tra được thay đổi.",
            );
            return;
          }

          try {
            const snapshot = await fetchSheetSnapshotWithMeta(sheetConfig);

            if (canceled) {
              return;
            }

            setData(snapshot.data);
            writeSheetCache(sheetConfig, snapshot.data, snapshot.revision || `snapshot-${Date.now()}`);
            setSyncStatus("Đã tải dữ liệu Google Sheet.");
          } catch (snapshotError) {
            setSyncStatus(
              snapshotError instanceof Error
                ? snapshotError.message
                : "Không tải được dữ liệu Google Sheet.",
            );
          }
        }
      }

      void pullIfChanged(true);
      intervalId = window.setInterval(() => {
        void pullIfChanged(false);
      }, sheetMetaPollMs);
    });

    return () => {
      canceled = true;
      if (intervalId) {
        window.clearInterval(intervalId);
      }
    };
  }, []);

  function commit(updater: (current: WarehouseData) => WarehouseData, message: string): ActionResult {
    if (!data) {
      return { ok: false, message: "Dữ liệu chưa sẵn sàng." };
    }

    const nextData = updater(data);
    setData(nextData);
    writeData(nextData);
    setLastMessage(message);
    return { ok: true, message };
  }

  const inventory = useMemo(() => (data ? computeInventory(data) : []), [data]);

  async function commitRemoteOrLocal(
    type: string,
    formData: FormData,
    localUpdater: (current: WarehouseData) => WarehouseData,
    message: string,
    options: { usePost?: boolean } = {},
  ): Promise<ActionResult> {
    const sheetConfig = await loadSheetConfig();

    if (isSheetConfigured(sheetConfig)) {
      try {
        setSyncMode("google-sheet");
        setSyncStatus("Đang ghi Google Sheet...");
        const payload = formDataToPayload(formData);
        const result = options.usePost
          ? await postMutateSheetWithMeta(type, payload, sheetConfig)
          : await mutateSheetWithMeta(type, payload, sheetConfig);
        setData(result.data);
        writeSheetCache(sheetConfig, result.data, result.revision || `mutate-${Date.now()}`);
        setLastMessage(message);
        setSyncStatus("Đã ghi Google Sheet và cập nhật cache.");
        return { ok: true, message };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Không ghi được Google Sheet.";
        setLastMessage(errorMessage);
        setSyncStatus(errorMessage);
        return { ok: false, message: errorMessage };
      }
    }

    return commit(localUpdater, message);
  }

  function addProduct(formData: FormData) {
    return commitRemoteOrLocal(
      "addProduct",
      formData,
      (current) => {
        const code = toText(formData.get("code"));
        const product: Product = {
          id: makeId("product"),
          code,
          name: toText(formData.get("name")),
          category: toText(formData.get("category")),
          size: toText(formData.get("size")),
          unit: "thùng",
          piecesPerBox: toNumber(formData.get("piecesPerBox")),
          sqmPerBox: toNumber(formData.get("sqmPerBox")),
          importPrice: toNumber(formData.get("importPrice")),
          salePrice: toNumber(formData.get("salePrice")),
          supplierId: toText(formData.get("supplierId")),
          minStock: toNumber(formData.get("minStock")),
          isActive: true,
          note: toText(formData.get("note")),
          createdAt: new Date().toISOString(),
          imageUrl: toText(formData.get("imageData")) || toText(formData.get("imageUrl")),
        };

        return {
          ...current,
          products: [product, ...current.products],
          auditLogs: [
            makeAudit("create", "products", code, `Tạo sản phẩm ${product.name}`),
            ...current.auditLogs,
          ],
        };
      },
      "Đã thêm sản phẩm.",
      { usePost: hasProductImageUpload(formData) },
    );
  }

  function updateProductImage(formData: FormData) {
    return commitRemoteOrLocal(
      "updateProductImage",
      formData,
      (current) => {
        const productId = toText(formData.get("productId"));
        const imageUrl = toText(formData.get("clearImage")) ? "" : toText(formData.get("imageData")) || toText(formData.get("imageUrl"));
        const product = current.products.find((item) => item.id === productId);

        if (!product) {
          return current;
        }

        return {
          ...current,
          products: current.products.map((item) =>
            item.id === productId ? { ...item, imageUrl } : item,
          ),
          auditLogs: [
            makeAudit("update", "products", product.code, imageUrl ? "Cập nhật ảnh sản phẩm" : "Gỡ ảnh sản phẩm"),
            ...current.auditLogs,
          ],
        };
      },
      "Đã cập nhật ảnh sản phẩm.",
      { usePost: true },
    );
  }

  function addCustomer(formData: FormData) {
    return commitRemoteOrLocal(
      "addCustomer",
      formData,
      (current) => {
        const customer: Customer = {
          id: makeId("customer"),
          name: toText(formData.get("name")),
          phone: toText(formData.get("phone")),
          address: toText(formData.get("address")),
          customerGroup: toText(formData.get("customerGroup")),
          note: toText(formData.get("note")),
          createdAt: new Date().toISOString(),
        };

        return {
          ...current,
          customers: [customer, ...current.customers],
          auditLogs: [
            makeAudit("create", "customers", customer.name, `Tạo khách hàng ${customer.name}`),
            ...current.auditLogs,
          ],
        };
      },
      "Đã thêm khách hàng.",
    );
  }

  function addSupplier(formData: FormData) {
    return commitRemoteOrLocal(
      "addSupplier",
      formData,
      (current) => {
        const supplier: Supplier = {
          id: makeId("supplier"),
          name: toText(formData.get("name")),
          phone: toText(formData.get("phone")),
          address: toText(formData.get("address")),
          contactPerson: toText(formData.get("contactPerson")),
          note: toText(formData.get("note")),
          createdAt: new Date().toISOString(),
        };

        return {
          ...current,
          suppliers: [supplier, ...current.suppliers],
          auditLogs: [
            makeAudit("create", "suppliers", supplier.name, `Tạo nhà cung cấp ${supplier.name}`),
            ...current.auditLogs,
          ],
        };
      },
      "Đã thêm nhà cung cấp.",
    );
  }

  function addPurchase(formData: FormData) {
    return commitRemoteOrLocal(
      "addPurchase",
      formData,
      (current) => {
        const items = getOrderItems(formData);
        const lineTotals = items.map((item) => item.quantity * item.unitPrice);
        const totalAmount = sum(lineTotals);
        const paidInput = toNumber(formData.get("paidAmount"));
        const paidAmount = Math.min(paidInput, totalAmount);
        const paidShares = allocateAmount(paidAmount, lineTotals);
        const code = nextCode("PO", current.purchases.map((order) => order.code));
        const orderDate = toText(formData.get("orderDate")) || getToday();
        const supplierId = toText(formData.get("supplierId"));
        const paymentMethod = toText(formData.get("paymentMethod"));
        const note = toText(formData.get("note"));
        const createdAt = new Date().toISOString();
        const purchases: PurchaseOrder[] = items.map((item, index) => {
          const lineTotal = lineTotals[index] ?? 0;
          const linePaid = paidShares[index] ?? 0;

          return {
            id: makeId("purchase"),
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
          };
        });
        const movements: StockMovement[] = purchases.map((purchase) => ({
          id: makeId("movement"),
          productId: purchase.productId,
          movementDate: orderDate,
          quantity: purchase.quantity,
          movementType: "purchase",
          referenceType: "purchase_order",
          referenceId: purchase.id,
          note: `Nhập ${code}`,
          createdAt,
        }));

        const payment: Payment | null =
          paidAmount > 0
            ? {
                id: makeId("payment"),
                paymentDate: orderDate,
                direction: "out",
                partyType: "supplier",
                supplierId,
                referenceType: "purchase_order",
                referenceId: code,
                amount: paidAmount,
                method: paymentMethod,
                note: `Trả tiền ${code}`,
                createdAt,
              }
            : null;

        return {
          ...current,
          purchases: [...purchases, ...current.purchases],
          stockMovements: [...movements, ...current.stockMovements],
          payments: payment ? [payment, ...current.payments] : current.payments,
          auditLogs: [
            makeAudit(
              "confirm",
              "purchase_orders",
              code,
              `Xác nhận nhập ${items.length} dòng / ${sum(items.map((item) => item.quantity))} thùng`,
            ),
            ...current.auditLogs,
          ],
        };
      },
      "Đã nhập kho và cộng tồn.",
    );
  }

  async function addSale(formData: FormData): Promise<ActionResult> {
    const sheetConfig = await loadSheetConfig();

    if (isSheetConfigured(sheetConfig)) {
      return commitRemoteOrLocal("addSale", formData, (current) => current, "Đã tạo đơn bán và trừ tồn.");
    }

    if (!data) {
      return { ok: false, message: "Dữ liệu chưa sẵn sàng." };
    }

    const items = getOrderItems(formData);

    if (items.length === 0) {
      const message = "Thêm ít nhất một dòng sản phẩm để bán.";
      setLastMessage(message);
      return { ok: false, message };
    }

    const requirements = getProductRequirements(items);

    for (const [productId, requiredQuantity] of Object.entries(requirements)) {
      const availableStock = getAvailableStock(data, productId);

      if (availableStock < requiredQuantity) {
        const message = `Không đủ tồn kho ${getProductName(data.products, productId)}. Hiện còn ${availableStock} thùng, cần ${requiredQuantity} thùng.`;
        setLastMessage(message);
        return { ok: false, message };
      }
    }

    return commit((current) => {
      const discountAmount = toNumber(formData.get("discountAmount"));
      const shippingFee = toNumber(formData.get("shippingFee"));
      const paidAmountInput = toNumber(formData.get("paidAmount"));
      const lineSubtotals = items.map((item) => item.quantity * item.unitPrice);
      const subtotalAmount = sum(lineSubtotals);
      const totalAmount = Math.max(subtotalAmount - discountAmount + shippingFee, 0);
      const paidAmount = Math.min(paidAmountInput, totalAmount);
      const totalShares = allocateAmount(totalAmount, lineSubtotals);
      const paidShares = allocateAmount(paidAmount, totalShares);
      const discountShares = allocateAmount(Math.min(discountAmount, subtotalAmount), lineSubtotals);
      const shippingShares = allocateAmount(shippingFee, lineSubtotals);
      const code = nextCode("SO", current.sales.map((order) => order.code));
      const orderDate = toText(formData.get("orderDate")) || getToday();
      const customerId = toText(formData.get("customerId"));
      const paymentMethod = toText(formData.get("paymentMethod"));
      const note = toText(formData.get("note"));
      const createdAt = new Date().toISOString();
      const sales: SalesOrder[] = items.map((item, index) => {
        const lineTotal = totalShares[index] ?? 0;
        const linePaid = paidShares[index] ?? 0;
        const debtAmount = Math.max(lineTotal - linePaid, 0);

        return {
          id: makeId("sale"),
          code,
          customerId,
          orderDate,
          productId: item.productId,
          quantity: item.quantity,
          unitPrice: item.unitPrice,
          discountAmount: discountShares[index] ?? 0,
          shippingFee: shippingShares[index] ?? 0,
          subtotalAmount: lineSubtotals[index] ?? 0,
          totalAmount: lineTotal,
          paidAmount: linePaid,
          debtAmount,
          paymentMethod,
          deliveryStatus: "pending",
          paymentStatus: debtAmount <= 0 ? "paid" : linePaid > 0 ? "partial" : "unpaid",
          status: "confirmed",
          note,
          createdAt,
        };
      });
      const movements: StockMovement[] = sales.map((sale) => ({
        id: makeId("movement"),
        productId: sale.productId,
        movementDate: orderDate,
        quantity: -sale.quantity,
        movementType: "sale",
        referenceType: "sales_order",
        referenceId: sale.id,
        note: `Bán ${code}`,
        createdAt,
      }));

      const payment: Payment | null =
        paidAmount > 0
          ? {
              id: makeId("payment"),
              paymentDate: orderDate,
              direction: "in",
              partyType: "customer",
              customerId,
              referenceType: "sales_order",
              referenceId: code,
              amount: paidAmount,
              method: paymentMethod,
              note: `Thu tiền ${code}`,
              createdAt,
            }
          : null;

      return {
        ...current,
        sales: [...sales, ...current.sales],
        stockMovements: [...movements, ...current.stockMovements],
        payments: payment ? [payment, ...current.payments] : current.payments,
        auditLogs: [
          makeAudit(
            "confirm",
            "sales_orders",
            code,
            `Xác nhận bán ${items.length} dòng / ${sum(items.map((item) => item.quantity))} thùng`,
          ),
          ...current.auditLogs,
        ],
      };
    }, "Đã tạo đơn bán và trừ tồn.");
  }

  async function recordCustomerPayment(formData: FormData): Promise<ActionResult> {
    const sheetConfig = await loadSheetConfig();

    if (isSheetConfigured(sheetConfig)) {
      return commitRemoteOrLocal(
        "recordCustomerPayment",
        formData,
        (current) => current,
        "Đã ghi nhận khách trả nợ và cập nhật công nợ.",
      );
    }

    if (!data) {
      return { ok: false, message: "Dữ liệu chưa sẵn sàng." };
    }

    const customerId = toText(formData.get("customerId"));
    const amount = toNumber(formData.get("amount"));
    const paymentDate = toText(formData.get("paymentDate")) || getToday();
    const method = toText(formData.get("method")) || "Tiền mặt";
    const note = toText(formData.get("note"));
    const outstandingOrders = data.sales
      .filter((order) => order.customerId === customerId && order.status !== "canceled" && order.debtAmount > 0)
      .sort((a, b) => a.orderDate.localeCompare(b.orderDate));
    const totalDebt = outstandingOrders.reduce((total, order) => total + order.debtAmount, 0);

    if (!customerId) {
      const message = "Chọn khách hàng cần trả nợ.";
      setLastMessage(message);
      return { ok: false, message };
    }

    if (amount <= 0) {
      const message = "Số tiền trả nợ phải lớn hơn 0.";
      setLastMessage(message);
      return { ok: false, message };
    }

    if (totalDebt <= 0) {
      const message = "Khách hàng này không còn công nợ.";
      setLastMessage(message);
      return { ok: false, message };
    }

    if (amount > totalDebt) {
      const message = `Số tiền vượt công nợ hiện tại. Khách còn nợ ${totalDebt.toLocaleString("vi-VN")}đ.`;
      setLastMessage(message);
      return { ok: false, message };
    }

    return commit((current) => {
      let remaining = amount;
      const affectedCodes: string[] = [];
      const updatedSales = current.sales.map((order) => {
        if (order.customerId !== customerId || order.status === "canceled" || order.debtAmount <= 0 || remaining <= 0) {
          return order;
        }

        const paidForOrder = Math.min(order.debtAmount, remaining);
        remaining -= paidForOrder;
        affectedCodes.push(order.code);
        const nextDebt = Math.max(order.debtAmount - paidForOrder, 0);
        const nextPaid = order.paidAmount + paidForOrder;

        return {
          ...order,
          paidAmount: nextPaid,
          debtAmount: nextDebt,
          paymentStatus: nextDebt <= 0 ? ("paid" as const) : ("partial" as const),
        };
      });

      const payment: Payment = {
        id: makeId("payment"),
        paymentDate,
        direction: "in",
        partyType: "customer",
        customerId,
        referenceType: "manual",
        referenceId: affectedCodes.join(", "),
        amount,
        method,
        note: note || `Khách trả nợ: ${affectedCodes.join(", ")}`,
        createdAt: new Date().toISOString(),
      };

      return {
        ...current,
        sales: updatedSales,
        payments: [payment, ...current.payments],
        auditLogs: [
          makeAudit(
            "payment",
            "customer_debts",
            affectedCodes.join(", "),
            `Ghi nhận khách trả ${amount.toLocaleString("vi-VN")}đ`,
          ),
          ...current.auditLogs,
        ],
      };
    }, "Đã ghi nhận khách trả nợ và cập nhật công nợ.");
  }

  async function recordSupplierPayment(formData: FormData): Promise<ActionResult> {
    const sheetConfig = await loadSheetConfig();

    if (isSheetConfigured(sheetConfig)) {
      return commitRemoteOrLocal(
        "recordSupplierPayment",
        formData,
        (current) => current,
        "Đã ghi nhận trả nợ nhà cung cấp và cập nhật công nợ.",
      );
    }

    if (!data) {
      return { ok: false, message: "Dữ liệu chưa sẵn sàng." };
    }

    const supplierId = toText(formData.get("supplierId"));
    const amount = toNumber(formData.get("amount"));
    const paymentDate = toText(formData.get("paymentDate")) || getToday();
    const method = toText(formData.get("method")) || "Tiền mặt";
    const note = toText(formData.get("note"));
    const outstandingOrders = data.purchases
      .filter((order) => order.supplierId === supplierId && order.status !== "canceled" && order.debtAmount > 0)
      .sort((a, b) => a.orderDate.localeCompare(b.orderDate));
    const totalDebt = outstandingOrders.reduce((total, order) => total + order.debtAmount, 0);

    if (!supplierId) {
      const message = "Chọn nhà cung cấp cần trả nợ.";
      setLastMessage(message);
      return { ok: false, message };
    }

    if (amount <= 0) {
      const message = "Số tiền trả nợ phải lớn hơn 0.";
      setLastMessage(message);
      return { ok: false, message };
    }

    if (totalDebt <= 0) {
      const message = "Nhà cung cấp này không còn công nợ.";
      setLastMessage(message);
      return { ok: false, message };
    }

    if (amount > totalDebt) {
      const message = `Số tiền vượt công nợ hiện tại. Nhà cung cấp còn nợ ${totalDebt.toLocaleString("vi-VN")}đ.`;
      setLastMessage(message);
      return { ok: false, message };
    }

    return commit((current) => {
      let remaining = amount;
      const affectedCodes: string[] = [];
      const updatedPurchases = current.purchases.map((order) => {
        if (order.supplierId !== supplierId || order.status === "canceled" || order.debtAmount <= 0 || remaining <= 0) {
          return order;
        }

        const paidForOrder = Math.min(order.debtAmount, remaining);
        remaining -= paidForOrder;
        affectedCodes.push(order.code);
        const nextDebt = Math.max(order.debtAmount - paidForOrder, 0);

        return {
          ...order,
          paidAmount: order.paidAmount + paidForOrder,
          debtAmount: nextDebt,
        };
      });

      const payment: Payment = {
        id: makeId("payment"),
        paymentDate,
        direction: "out",
        partyType: "supplier",
        supplierId,
        referenceType: "manual",
        referenceId: affectedCodes.join(", "),
        amount,
        method,
        note: note || `Trả nợ nhà cung cấp: ${affectedCodes.join(", ")}`,
        createdAt: new Date().toISOString(),
      };

      return {
        ...current,
        purchases: updatedPurchases,
        payments: [payment, ...current.payments],
        auditLogs: [
          makeAudit(
            "payment",
            "supplier_debts",
            affectedCodes.join(", "),
            `Ghi nhận trả NCC ${amount.toLocaleString("vi-VN")}đ`,
          ),
          ...current.auditLogs,
        ],
      };
    }, "Đã ghi nhận trả nợ nhà cung cấp và cập nhật công nợ.");
  }

  function addExpense(formData: FormData) {
    return commitRemoteOrLocal(
      "addExpense",
      formData,
      (current) => {
        const amount = toNumber(formData.get("amount"));
        const expenseDate = toText(formData.get("expenseDate")) || getToday();
        const expense: Expense = {
          id: makeId("expense"),
          category: toText(formData.get("category")),
          amount,
          expenseDate,
          paymentMethod: toText(formData.get("paymentMethod")),
          note: toText(formData.get("note")),
          createdAt: new Date().toISOString(),
        };

        const payment: Payment = {
          id: makeId("payment"),
          paymentDate: expenseDate,
          direction: "out",
          partyType: "other",
          referenceType: "expense",
          referenceId: expense.id,
          amount,
          method: expense.paymentMethod,
          note: expense.note,
          createdAt: new Date().toISOString(),
        };

        return {
          ...current,
          expenses: [expense, ...current.expenses],
          payments: [payment, ...current.payments],
          auditLogs: [
            makeAudit("create", "expenses", expense.category, `Tạo chi phí ${expense.category}`),
            ...current.auditLogs,
          ],
        };
      },
      "Đã thêm khoản chi.",
    );
  }

  async function restoreBackup(backup: LocalBackup): Promise<ActionResult> {
    if (backup.app !== "tile-warehouse-app" || backup.version !== 1 || !backup.data) {
      const message = "File backup không đúng định dạng.";
      setLastMessage(message);
      return { ok: false, message };
    }

    const sheetConfig = await loadSheetConfig();

    if (isSheetConfigured(sheetConfig)) {
      try {
        setSyncMode("google-sheet");
        setSyncStatus("Đang phục hồi backup vào Google Sheet...");
        const result = await postMutateSheetWithMeta(
          "restoreBackup",
          {
            app: backup.app,
            version: backup.version,
            exportedAt: backup.exportedAt,
            data: backup.data,
          },
          sheetConfig,
        );

        setData(result.data);
        writeSheetCache(sheetConfig, result.data, result.revision || `restore-${Date.now()}`);
        setLastMessage("Đã phục hồi backup vào Google Sheet.");
        setSyncStatus("Đã phục hồi backup vào Google Sheet và cập nhật cache.");
        return { ok: true, message: "Đã phục hồi backup vào Google Sheet." };
      } catch (error) {
        const message = error instanceof Error ? error.message : "Không phục hồi được backup vào Google Sheet.";
        setLastMessage(message);
        setSyncStatus(message);
        return { ok: false, message };
      }
    }

    setData(backup.data);
    writeData(backup.data);
    setLastMessage("Đã phục hồi backup vào dữ liệu local.");
    setSyncStatus("Đã phục hồi backup vào dữ liệu local trong trình duyệt.");
    return { ok: true, message: "Đã phục hồi backup vào dữ liệu local." };
  }

  return {
    data,
    inventory,
    lastMessage,
    syncMode,
    syncStatus,
    actions: {
      addProduct,
      updateProductImage,
      addCustomer,
      addSupplier,
      addPurchase,
      addSale,
      recordCustomerPayment,
      recordSupplierPayment,
      addExpense,
      restoreBackup,
    },
  };
}
