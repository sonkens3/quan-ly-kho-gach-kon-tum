"use client";

/* eslint-disable @next/next/no-img-element */

import Link from "next/link";
import { useMemo, useRef, useState } from "react";
import { CalendarDays, Download, Filter, ImageIcon, Loader2, Pencil, Plus, Search, Trash2, Upload, X } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { AutoFileBackupAgent } from "@/components/local/auto-file-backup-agent";
import { LocalModeTools } from "@/components/local/local-mode-tools";
import { MoneyInput } from "@/components/ui/money-input";
import { formatCurrency, formatDate, formatNumber, toDateKey } from "@/lib/formatters";
import {
  getCustomerBought,
  getCustomerDebt,
  getCustomerName,
  getCustomerPaid,
  getProductName,
  getToday,
  getSupplierDebt,
  getSupplierImported,
  getSupplierName,
  getSupplierPaid,
  sum,
} from "@/lib/local/calculations";
import { useWarehouseStore } from "@/lib/local/store";
import type { InventoryRow, Product, WarehouseData } from "@/lib/local/types";
import { modulePages, type ModulePageConfig } from "@/lib/module-pages";
import { cn } from "@/lib/utils";

type PageKey = keyof typeof modulePages;

type Row = Record<string, string>;

type ProductImageTarget = {
  id: string;
  code: string;
  name: string;
  imageUrl: string;
};

type CompressedProductImage = {
  dataUrl: string;
  fileName: string;
  size: number;
};

const maxProductImageSide = 900;
const maxProductImageBytes = 360 * 1024;
const productImageSides = [maxProductImageSide, 760, 640];
const productImageQualities = [0.72, 0.64, 0.56, 0.48];

function getImageDataSize(dataUrl: string) {
  const base64 = dataUrl.split(",")[1] ?? "";
  return Math.ceil((base64.length * 3) / 4);
}

function loadImageFile(file: File) {
  return new Promise<HTMLImageElement>((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const image = new Image();

    image.onload = () => {
      URL.revokeObjectURL(url);
      resolve(image);
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Không đọc được file ảnh."));
    };
    image.src = url;
  });
}

function makeCompressedFileName(fileName: string) {
  const baseName = fileName
    .replace(/\.[^.]+$/, "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\w.-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase();

  return `${baseName || "anh-san-pham"}.jpg`;
}

async function compressProductImage(file: File): Promise<CompressedProductImage> {
  if (!file.type.startsWith("image/")) {
    throw new Error("Chỉ chọn file ảnh sản phẩm.");
  }

  const image = await loadImageFile(file);
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");

  if (!context) {
    throw new Error("Trình duyệt không hỗ trợ nén ảnh.");
  }

  let bestDataUrl = "";
  let bestSize = Number.POSITIVE_INFINITY;

  for (const maxSide of productImageSides) {
    const scale = Math.min(1, maxSide / Math.max(image.naturalWidth, image.naturalHeight));
    const width = Math.max(1, Math.round(image.naturalWidth * scale));
    const height = Math.max(1, Math.round(image.naturalHeight * scale));

    canvas.width = width;
    canvas.height = height;
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, width, height);
    context.drawImage(image, 0, 0, width, height);

    for (const quality of productImageQualities) {
      const dataUrl = canvas.toDataURL("image/jpeg", quality);
      const size = getImageDataSize(dataUrl);

      if (size < bestSize) {
        bestDataUrl = dataUrl;
        bestSize = size;
      }

      if (size <= maxProductImageBytes) {
        return {
          dataUrl,
          fileName: makeCompressedFileName(file.name),
          size,
        };
      }
    }
  }

  return {
    dataUrl: bestDataUrl,
    fileName: makeCompressedFileName(file.name),
    size: bestSize,
  };
}

async function attachCompressedImage(formData: FormData) {
  const file = formData.get("imageFile");

  formData.delete("imageFile");

  if (!(file instanceof File) || file.size <= 0) {
    return formData;
  }

  const compressed = await compressProductImage(file);
  formData.set("imageData", compressed.dataUrl);
  formData.set("imageFileName", compressed.fileName);
  formData.set("imageSize", String(compressed.size));
  return formData;
}

function normalizeFilterValue(value: string) {
  return value.trim().toLowerCase();
}

function formatPhone(value: string | number | null | undefined) {
  const phone = String(value ?? "").trim();

  if (!phone || phone.startsWith("+") || phone.startsWith("0")) {
    return phone;
  }

  return /^\d{9}$/.test(phone) ? `0${phone}` : phone;
}

function formatOrderItems(
  rows: Array<{ productId: string; quantity: number }>,
  products: Product[],
) {
  return rows
    .map((row) => `${getProductName(products, row.productId)} x${formatNumber(row.quantity)}`)
    .join("; ");
}

function escapeExcelValue(value: unknown) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function makeExportSlug(value: string) {
  return (
    value
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/đ/g, "d")
      .replace(/Đ/g, "D")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "") || "du-lieu"
  );
}

function downloadExcelFile(
  title: string,
  columns: ModulePageConfig["columns"],
  rows: Row[],
) {
  const today = new Date().toISOString().slice(0, 10);
  const filename = `kho-gach-${makeExportSlug(title)}-${today}.xls`;
  const headerCells = columns
    .map((column) => `<th style="background:#e2e8f0;font-weight:bold">${escapeExcelValue(column.label)}</th>`)
    .join("");
  const bodyRows =
    rows.length > 0
      ? rows
          .map(
            (row) =>
              `<tr>${columns
                .map((column) => `<td style="mso-number-format:'\\@';">${escapeExcelValue(row[column.key] ?? "")}</td>`)
                .join("")}</tr>`,
          )
          .join("")
      : `<tr><td colspan="${columns.length}">Không có dữ liệu</td></tr>`;
  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <style>
    table { border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; }
    th, td { border: 1px solid #94a3b8; padding: 6px 8px; vertical-align: top; }
    h1 { font-family: Arial, sans-serif; font-size: 16px; }
  </style>
</head>
<body>
  <h1>${escapeExcelValue(title)}</h1>
  <table>
    <thead><tr>${headerCells}</tr></thead>
    <tbody>${bodyRows}</tbody>
  </table>
</body>
</html>`;
  const blob = new Blob(["\ufeff", html], {
    type: "application/vnd.ms-excel;charset=utf-8",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

const usageSteps = [
  {
    title: "1. Tạo nhà cung cấp",
    href: "/suppliers",
    body:
      "Vào Nhà cung cấp, bấm Thêm NCC, nhập tên nhà cung cấp, số điện thoại, người liên hệ, địa chỉ rồi bấm Lưu. Bước này nên làm trước để sản phẩm và phiếu nhập có nơi liên kết.",
  },
  {
    title: "2. Tạo sản phẩm gạch",
    href: "/products",
    body:
      "Vào Sản phẩm gạch, bấm Thêm sản phẩm. Nhập mã hàng, tên hàng, loại, kích thước, quy cách viên/thùng, m2/thùng, giá nhập, giá bán, nhà cung cấp và tồn tối thiểu. Mã hàng nên đặt rõ như G6060-A01 để dễ tìm.",
  },
  {
    title: "3. Nhập hàng vào kho",
    href: "/purchases",
    body:
      "Vào Nhập hàng, bấm Tạo phiếu nhập. Chọn ngày nhập, nhà cung cấp, sản phẩm, số lượng thùng, đơn giá nhập và số tiền đã trả. Khi lưu, hệ thống tự cộng tồn kho, ghi lịch sử nhập và tạo công nợ nhà cung cấp nếu chưa trả đủ.",
  },
  {
    title: "4. Tạo khách hàng",
    href: "/customers",
    body:
      "Vào Khách hàng, bấm Thêm khách. Nhập tên khách, số điện thoại, địa chỉ, nhóm khách và ghi chú nếu có. Nên tạo khách trước khi bán để công nợ và lịch sử mua hàng theo đúng người.",
  },
  {
    title: "5. Tạo đơn bán hàng",
    href: "/sales",
    body:
      "Vào Bán hàng, bấm Tạo đơn. Chọn ngày bán, khách hàng, sản phẩm, số lượng, đơn giá bán, chiết khấu, phí giao hàng và số tiền khách đã trả. Khi lưu, hệ thống kiểm tra tồn kho; nếu đủ hàng thì trừ tồn, ghi doanh thu và tạo công nợ nếu khách chưa trả đủ.",
  },
  {
    title: "6. Ghi nhận khách trả nợ",
    href: "/customer-debts",
    body:
      "Vào Nợ khách hàng, bấm Ghi nhận thu. Chọn ngày trả, khách hàng, nhập số tiền khách trả và phương thức thanh toán. Hệ thống tự trừ vào các đơn còn nợ cũ nhất trước, cập nhật đã trả, còn nợ, thu chi và nhật ký thao tác.",
  },
  {
    title: "7. Kiểm tra tồn kho và công nợ",
    href: "/inventory",
    body:
      "Vào Tồn kho để xem số thùng còn lại, giá trị tồn, hàng sắp hết hoặc hết hàng. Vào Nợ khách hàng và Nợ nhà cung cấp để xem ngày phát sinh nợ, lần trả gần nhất, tổng đã mua/nhập, đã trả và còn nợ.",
  },
  {
    title: "8. Theo dõi thu chi và báo cáo",
    href: "/cashflow",
    body:
      "Vào Thu chi để xem tiền thu từ bán hàng, khách trả nợ, tiền chi nhập hàng và các khoản chi phí. Vào Báo cáo để xem tổng hợp doanh thu, chi phí, giá vốn tạm tính và lợi nhuận tạm tính.",
  },
  {
    title: "9. Sao lưu dữ liệu",
    href: "/settings",
    body:
      "Ở khung Backup dữ liệu, bấm Xuất backup để tải file JSON về máy. Khi dùng máy khác hoặc cần phục hồi, bấm Nhập backup và chọn file đã xuất. Nếu đang dùng Google Sheet, dữ liệu chính vẫn nằm trong Sheet và app tự đọc lại định kỳ.",
  },
];

function Field({
  label,
  children,
}: {
  label: string;
  children: React.ReactNode;
}) {
  return (
    <label className="space-y-1 text-sm">
      <span className="font-medium text-slate-700">{label}</span>
      {children}
    </label>
  );
}

function SelectField({
  name,
  children,
  defaultValue,
}: {
  name: string;
  children: React.ReactNode;
  defaultValue?: string;
}) {
  return (
    <select
      name={name}
      defaultValue={defaultValue}
      className="h-10 w-full rounded-md border border-input bg-white px-3 text-sm text-slate-800 outline-none focus:border-cyan-600 focus:ring-2 focus:ring-cyan-600/15"
      required
    >
      {children}
    </select>
  );
}

function ProductImageCell({
  row,
  onEdit,
  onPreview,
}: {
  row: Row;
  onEdit: (target: ProductImageTarget) => void;
  onPreview: (target: ProductImageTarget) => void;
}) {
  const target: ProductImageTarget = {
    id: row._productId,
    code: row._productCode,
    name: row._productName,
    imageUrl: row._imageUrl,
  };

  if (!target.id) {
    return <span className="text-slate-400">-</span>;
  }

  return (
    <div className="flex items-center justify-center gap-2">
      {target.imageUrl ? (
        <button
          type="button"
          onClick={() => onPreview(target)}
          className="h-12 w-12 overflow-hidden rounded-md border border-slate-200 bg-slate-50"
          title="Xem ảnh lớn"
        >
          <img src={target.imageUrl} alt={target.name} className="h-full w-full object-cover" loading="lazy" />
        </button>
      ) : (
        <span className="flex h-12 w-12 items-center justify-center rounded-md border border-dashed border-slate-300 bg-slate-50 text-slate-400">
          <ImageIcon className="h-5 w-5" aria-hidden="true" />
        </span>
      )}
      <Button type="button" variant="secondary" size="sm" onClick={() => onEdit(target)}>
        <Pencil className="h-4 w-4" aria-hidden="true" />
        {target.imageUrl ? "Đổi" : "Thêm"}
      </Button>
    </div>
  );
}

type OrderLineInput = {
  id: string;
  productId: string;
};

function getLineDefaultPrice(products: Product[], productId: string, mode: "purchase" | "sale") {
  const product = products.find((item) => item.id === productId) ?? products[0];

  return mode === "purchase" ? product?.importPrice ?? 0 : product?.salePrice ?? 0;
}

function MultiProductItems({
  products,
  mode,
}: {
  products: Product[];
  mode: "purchase" | "sale";
}) {
  const firstProductId = products[0]?.id ?? "";
  const [lines, setLines] = useState<OrderLineInput[]>([
    { id: `line-${Date.now()}`, productId: firstProductId },
  ]);

  function updateProduct(lineId: string, productId: string) {
    setLines((current) =>
      current.map((line) => (line.id === lineId ? { ...line, productId } : line)),
    );
  }

  function addLine() {
    setLines((current) => [
      ...current,
      { id: `line-${Date.now()}-${current.length}`, productId: firstProductId },
    ]);
  }

  function removeLine(lineId: string) {
    setLines((current) => (current.length <= 1 ? current : current.filter((line) => line.id !== lineId)));
  }

  return (
    <div className="sm:col-span-2 xl:col-span-4">
      <div className="mb-2 flex items-center justify-between gap-3">
        <p className="text-sm font-semibold text-slate-800">Danh sách sản phẩm</p>
        <Button type="button" variant="secondary" size="sm" onClick={addLine}>
          <Plus className="h-4 w-4" aria-hidden="true" />
          Thêm dòng
        </Button>
      </div>

      <div className="space-y-2">
        {lines.map((line, index) => (
          <div
            key={line.id}
            className="grid gap-2 rounded-md border border-cyan-100 bg-white p-3 md:grid-cols-[minmax(220px,1fr)_120px_180px_42px]"
          >
            <label className="space-y-1 text-sm">
              <span className="font-medium text-slate-700">Sản phẩm {index + 1}</span>
              <select
                name="itemProductId"
                value={line.productId}
                onChange={(event) => updateProduct(line.id, event.target.value)}
                className="h-10 w-full rounded-md border border-input bg-white px-3 text-sm text-slate-800 outline-none focus:border-cyan-600 focus:ring-2 focus:ring-cyan-600/15"
                required
              >
                {products.map((product) => (
                  <option key={product.id} value={product.id}>
                    {product.code} - {product.name}
                  </option>
                ))}
              </select>
            </label>

            <Field label="Số thùng">
              <Input name="itemQuantity" type="number" min="1" step="1" required defaultValue="1" />
            </Field>

            <Field label={mode === "purchase" ? "Đơn giá nhập" : "Đơn giá bán"}>
              <MoneyInput
                key={`${line.id}-${line.productId}-${mode}`}
                name="itemUnitPrice"
                required
                defaultValue={getLineDefaultPrice(products, line.productId, mode)}
              />
            </Field>

            <div className="flex items-end">
              <Button
                type="button"
                variant="secondary"
                size="icon"
                onClick={() => removeLine(line.id)}
                disabled={lines.length <= 1}
                aria-label="Xóa dòng sản phẩm"
              >
                <Trash2 className="h-4 w-4" aria-hidden="true" />
              </Button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

export function SettingsUsageGuide() {
  return (
    <section className="rounded-md border border-slate-200 bg-white shadow-soft">
      <div className="border-b border-slate-200 px-4 py-3">
        <p className="text-sm font-medium text-cyan-700">Hướng dẫn sử dụng</p>
        <h2 className="mt-1 text-lg font-semibold text-slate-950">
          Quy trình dùng phần mềm kho gạch
        </h2>
        <p className="mt-1 text-sm text-slate-500">
          Làm theo thứ tự dưới đây để dữ liệu tồn kho, công nợ và thu chi tự khớp với nhau.
        </p>
      </div>

      <div className="grid gap-3 p-4 lg:grid-cols-2">
        {usageSteps.map((step) => (
          <div key={step.title} className="rounded-md border border-slate-200 bg-slate-50 p-4">
            <div className="flex flex-col gap-2 sm:flex-row sm:items-start sm:justify-between">
              <h3 className="text-sm font-semibold text-slate-950">{step.title}</h3>
              <Link
                href={step.href}
                className="text-sm font-medium text-cyan-700 hover:text-cyan-900"
              >
                Mở trang
              </Link>
            </div>
            <p className="mt-2 text-sm leading-6 text-slate-600">{step.body}</p>
          </div>
        ))}
      </div>

      <div className="border-t border-slate-200 bg-amber-50 px-4 py-3 text-sm leading-6 text-amber-900">
        Lưu ý: trước khi bán hàng phải có sản phẩm và tồn kho. Nếu hệ thống báo không đủ tồn,
        hãy nhập hàng trước hoặc kiểm tra lại lịch sử nhập/bán của mã hàng đó.
      </div>
    </section>
  );
}

function getRows(pageKey: PageKey, data: WarehouseData, inventory: InventoryRow[] = []): Row[] {
  if (pageKey === "products") {
    return data.products.map((product) => ({
      _productId: product.id,
      _productCode: product.code,
      _productName: product.name,
      _imageUrl: product.imageUrl ?? "",
      _date: toDateKey(product.createdAt),
      _status: product.isActive ? "Đang bán" : "Ngừng bán",
      image: product.imageUrl ? "Có ảnh" : "Chưa có ảnh",
      code: product.code,
      name: product.name,
      category: product.category,
      size: product.size,
      unit: product.unit,
      salePrice: formatCurrency(product.salePrice),
      status: product.isActive ? "Đang bán" : "Ngừng bán",
    }));
  }

  if (pageKey === "customers") {
    return data.customers.map((customer) => ({
      _date: toDateKey(customer.createdAt),
      _status: getCustomerDebt(data, customer.id) > 0 ? "Có công nợ Còn nợ" : "Không công nợ Đã tất toán",
      name: customer.name,
      phone: formatPhone(customer.phone),
      address: customer.address,
      group: customer.customerGroup,
      total: formatCurrency(getCustomerBought(data, customer.id)),
      debt: formatCurrency(getCustomerDebt(data, customer.id)),
    }));
  }

  if (pageKey === "suppliers") {
    return data.suppliers.map((supplier) => ({
      _date: toDateKey(supplier.createdAt),
      _status: getSupplierDebt(data, supplier.id) > 0 ? "Có công nợ Còn nợ" : "Không công nợ Đã tất toán",
      name: supplier.name,
      phone: formatPhone(supplier.phone),
      contact: supplier.contactPerson,
      address: supplier.address,
      imported: formatCurrency(getSupplierImported(data, supplier.id)),
      debt: formatCurrency(getSupplierDebt(data, supplier.id)),
    }));
  }

  if (pageKey === "purchases") {
    const groups = new Map<string, typeof data.purchases>();

    data.purchases.forEach((order) => {
      groups.set(order.code, [...(groups.get(order.code) ?? []), order]);
    });

    return Array.from(groups.values()).map((orders) => {
      const first = orders[0];
      const totalAmount = sum(orders.map((order) => order.totalAmount));
      const paidAmount = sum(orders.map((order) => order.paidAmount));
      const debtAmount = sum(orders.map((order) => order.debtAmount));
      const quantity = sum(orders.map((order) => order.quantity));

      return {
        _date: toDateKey(first.orderDate),
        _status: `${first.status === "confirmed" ? "Đã xác nhận" : "Đã hủy"} ${debtAmount > 0 ? "Còn nợ" : "Đã trả đủ"}`,
        code: first.code,
        date: formatDate(first.orderDate),
        supplier: getSupplierName(data.suppliers, first.supplierId),
        items: `${orders.length} dòng / ${formatNumber(quantity)} thùng: ${formatOrderItems(orders, data.products)}`,
        total: formatCurrency(totalAmount),
        paid: formatCurrency(paidAmount),
        status: debtAmount > 0 ? "Còn nợ" : "Đã trả đủ",
      };
    });
  }

  if (pageKey === "sales") {
    const groups = new Map<string, typeof data.sales>();

    data.sales.forEach((order) => {
      groups.set(order.code, [...(groups.get(order.code) ?? []), order]);
    });

    return Array.from(groups.values()).map((orders) => {
      const first = orders[0];
      const totalAmount = sum(orders.map((order) => order.totalAmount));
      const paidAmount = sum(orders.map((order) => order.paidAmount));
      const debtAmount = sum(orders.map((order) => order.debtAmount));
      const quantity = sum(orders.map((order) => order.quantity));

      return {
        _date: toDateKey(first.orderDate),
        _status: [
          first.status === "confirmed" ? "Đã xác nhận" : first.status === "completed" ? "Hoàn tất" : "Đã hủy",
          first.deliveryStatus === "delivering"
            ? "Đang giao"
            : first.deliveryStatus === "delivered"
              ? "Đã giao"
              : "Chờ giao",
          debtAmount > 0 ? "Còn nợ" : "Đã trả đủ",
        ].join(" "),
        code: first.code,
        date: formatDate(first.orderDate),
        customer: getCustomerName(data.customers, first.customerId),
        items: `${orders.length} dòng / ${formatNumber(quantity)} thùng: ${formatOrderItems(orders, data.products)}`,
        total: formatCurrency(totalAmount),
        paid: formatCurrency(paidAmount),
        status: debtAmount > 0 ? "Còn nợ" : "Đã trả đủ",
      };
    });
  }

  if (pageKey === "inventory") {
    return inventory.map((item) => ({
      _productId: item.id,
      _productCode: item.code,
      _productName: item.name,
      _imageUrl: item.imageUrl ?? "",
      _date: toDateKey(item.createdAt),
      _status:
        item.stockStatus === "out_of_stock"
          ? "Hết hàng"
          : item.stockStatus === "low_stock"
            ? "Sắp hết"
            : item.stockStatus === "inactive"
              ? "Ngừng bán"
              : "Còn hàng Đang bán",
      image: item.imageUrl ? "Có ảnh" : "Chưa có ảnh",
      code: item.code,
      name: item.name,
      size: item.size,
      stock: formatNumber(item.currentStock),
      sqm: formatNumber(item.currentSqm),
      value: formatCurrency(item.stockValue),
      status:
        item.stockStatus === "out_of_stock"
          ? "Hết hàng"
          : item.stockStatus === "low_stock"
            ? "Sắp hết"
            : item.stockStatus === "inactive"
              ? "Ngừng bán"
              : "Còn hàng",
    }));
  }

  if (pageKey === "customer-debts") {
    return data.customers
      .map((customer) => {
        const outstandingOrders = data.sales
          .filter((order) => order.customerId === customer.id && order.status !== "canceled" && order.debtAmount > 0)
          .sort((a, b) => a.orderDate.localeCompare(b.orderDate));
        const customerPayments = data.payments
          .filter((payment) => payment.customerId === customer.id && payment.direction === "in")
          .sort((a, b) => b.paymentDate.localeCompare(a.paymentDate));

        return {
          _date: toDateKey(outstandingOrders[0]?.orderDate),
          _status: getCustomerDebt(data, customer.id) > 0 ? "Còn nợ" : "Đã tất toán",
          customer: customer.name,
          phone: formatPhone(customer.phone),
          debtDate: formatDate(outstandingOrders[0]?.orderDate),
          lastPaymentDate: formatDate(customerPayments[0]?.paymentDate),
          bought: formatCurrency(getCustomerBought(data, customer.id)),
          paid: formatCurrency(getCustomerPaid(data, customer.id)),
          debt: formatCurrency(getCustomerDebt(data, customer.id)),
          status: getCustomerDebt(data, customer.id) > 0 ? "Còn nợ" : "Đã tất toán",
        };
      })
      .filter((row) => row.debt !== formatCurrency(0));
  }

  if (pageKey === "supplier-debts") {
    return data.suppliers
      .map((supplier) => {
        const outstandingOrders = data.purchases
          .filter((order) => order.supplierId === supplier.id && order.status !== "canceled" && order.debtAmount > 0)
          .sort((a, b) => a.orderDate.localeCompare(b.orderDate));
        const supplierPayments = data.payments
          .filter((payment) => payment.supplierId === supplier.id && payment.direction === "out")
          .sort((a, b) => b.paymentDate.localeCompare(a.paymentDate));

        return {
          _date: toDateKey(outstandingOrders[0]?.orderDate),
          _status: getSupplierDebt(data, supplier.id) > 0 ? "Còn nợ" : "Đã tất toán",
          supplier: supplier.name,
          contact: supplier.contactPerson,
          debtDate: formatDate(outstandingOrders[0]?.orderDate),
          lastPaymentDate: formatDate(supplierPayments[0]?.paymentDate),
          imported: formatCurrency(getSupplierImported(data, supplier.id)),
          paid: formatCurrency(getSupplierPaid(data, supplier.id)),
          debt: formatCurrency(getSupplierDebt(data, supplier.id)),
          status: getSupplierDebt(data, supplier.id) > 0 ? "Còn nợ" : "Đã tất toán",
        };
      })
      .filter((row) => row.debt !== formatCurrency(0));
  }

  if (pageKey === "cashflow") {
    const payments = data.payments.map((payment) => ({
      _date: toDateKey(payment.paymentDate),
      _status: `${payment.direction === "in" ? "Thu" : "Chi"} ${payment.method}`,
      date: formatDate(payment.paymentDate),
      type: payment.direction === "in" ? "Thu" : "Chi",
      category:
        payment.partyType === "customer"
          ? "Thu khách hàng"
          : payment.partyType === "supplier"
            ? "Trả nhà cung cấp"
            : "Chi khác",
      amount: formatCurrency(payment.amount),
      method: payment.method,
      createdBy: "Local",
    }));

    return payments;
  }

  if (pageKey === "audit-logs") {
    return data.auditLogs.map((log) => ({
      _date: toDateKey(log.time),
      _status: log.action,
      time: formatDate(log.time),
      user: log.user,
      action: log.action,
      table: log.tableName,
      record: log.recordCode,
      note: log.note,
    }));
  }

  if (pageKey === "reports") {
    const revenue = sum(data.sales.map((order) => order.totalAmount));
    const purchaseCost = sum(data.purchases.map((order) => order.totalAmount));
    const expenses = sum(data.expenses.map((expense) => expense.amount));
    return [
      {
        _date: getToday(),
        _status: "Tạm tính",
        name: "Tổng hợp local",
        period: "Toàn bộ dữ liệu",
        revenue: formatCurrency(revenue),
        cost: formatCurrency(purchaseCost + expenses),
        profit: formatCurrency(revenue - purchaseCost - expenses),
        status: "Tạm tính",
      },
    ];
  }

  if (pageKey === "users") {
    return [
      {
        _date: getToday(),
        _status: "Hoạt động",
        name: "Chủ kho local",
        email: "admin@local.vn",
        phone: "-",
        role: "Admin",
        status: "Hoạt động",
        createdAt: formatDate(new Date()),
      },
    ];
  }

  if (pageKey === "settings") {
    return [
      {
        _date: getToday(),
        _status: "Miễn phí",
        name: "Chế độ dữ liệu",
        value: "localStorage",
        scope: "Trình duyệt hiện tại",
        updatedBy: "Hệ thống",
        updatedAt: formatDate(new Date()),
        status: "Miễn phí",
      },
    ];
  }

  return [];
}

function getStats(
  pageKey: PageKey,
  config: ModulePageConfig,
  data: WarehouseData,
  inventory: InventoryRow[],
) {
  const totalDebtCustomers = sum(data.customers.map((customer) => getCustomerDebt(data, customer.id)));
  const totalDebtSuppliers = sum(data.suppliers.map((supplier) => getSupplierDebt(data, supplier.id)));
  const stockValue = sum(inventory.map((item) => item.stockValue));
  const revenue = sum(data.sales.map((order) => order.totalAmount));
  const expenses = sum(data.expenses.map((expense) => expense.amount));
  const saleCodes = new Set(data.sales.map((order) => order.code));
  const purchaseCodes = new Set(data.purchases.map((order) => order.code));
  const debtSaleCodes = new Set(
    data.sales.filter((order) => order.debtAmount > 0).map((order) => order.code),
  );
  const debtPurchaseCodes = new Set(
    data.purchases.filter((order) => order.debtAmount > 0).map((order) => order.code),
  );

  if (pageKey === "sales") {
    return [
      { label: "Tổng đơn bán", value: String(saleCodes.size), tone: "cyan" as const },
      { label: "Doanh thu", value: formatCurrency(revenue), tone: "green" as const },
      { label: "Khách còn nợ", value: formatCurrency(totalDebtCustomers), tone: "amber" as const },
      { label: "Đơn còn nợ", value: String(debtSaleCodes.size), tone: "slate" as const },
    ];
  }

  if (pageKey === "purchases") {
    return [
      { label: "Tổng phiếu nhập", value: String(purchaseCodes.size), tone: "cyan" as const },
      { label: "Giá trị nhập", value: formatCurrency(sum(data.purchases.map((order) => order.totalAmount))), tone: "green" as const },
      { label: "Nợ nhà cung cấp", value: formatCurrency(totalDebtSuppliers), tone: "amber" as const },
      { label: "Phiếu còn nợ", value: String(debtPurchaseCodes.size), tone: "slate" as const },
    ];
  }

  if (pageKey === "inventory" || pageKey === "products") {
    return [
      { label: "Mã hàng", value: String(data.products.length), tone: "cyan" as const },
      { label: "Giá trị tồn", value: formatCurrency(stockValue), tone: "green" as const },
      { label: "Sắp hết", value: String(inventory.filter((item) => item.stockStatus === "low_stock").length), tone: "amber" as const },
      { label: "Hết hàng", value: String(inventory.filter((item) => item.stockStatus === "out_of_stock").length), tone: "red" as const },
    ];
  }

  if (pageKey === "customers" || pageKey === "customer-debts") {
    return [
      { label: "Tổng khách", value: String(data.customers.length), tone: "cyan" as const },
      { label: "Tổng đã mua", value: formatCurrency(sum(data.customers.map((customer) => getCustomerBought(data, customer.id)))), tone: "green" as const },
      { label: "Còn nợ", value: formatCurrency(totalDebtCustomers), tone: "amber" as const },
      { label: "Khách còn nợ", value: String(data.customers.filter((customer) => getCustomerDebt(data, customer.id) > 0).length), tone: "red" as const },
    ];
  }

  if (pageKey === "suppliers" || pageKey === "supplier-debts") {
    return [
      { label: "Tổng NCC", value: String(data.suppliers.length), tone: "cyan" as const },
      { label: "Tổng đã nhập", value: formatCurrency(sum(data.suppliers.map((supplier) => getSupplierImported(data, supplier.id)))), tone: "green" as const },
      { label: "Còn nợ", value: formatCurrency(totalDebtSuppliers), tone: "amber" as const },
      { label: "NCC còn nợ", value: String(data.suppliers.filter((supplier) => getSupplierDebt(data, supplier.id) > 0).length), tone: "red" as const },
    ];
  }

  if (pageKey === "cashflow") {
    const totalIn = sum(data.payments.filter((payment) => payment.direction === "in").map((payment) => payment.amount));
    const totalOut = sum(data.payments.filter((payment) => payment.direction === "out").map((payment) => payment.amount));
    return [
      { label: "Tổng thu", value: formatCurrency(totalIn), tone: "green" as const },
      { label: "Tổng chi", value: formatCurrency(totalOut), tone: "red" as const },
      { label: "Chênh lệch", value: formatCurrency(totalIn - totalOut), tone: "cyan" as const },
      { label: "Chi phí", value: formatCurrency(expenses), tone: "slate" as const },
    ];
  }

  return config.stats;
}

function EntryForm({
  pageKey,
  data,
  onSubmit,
  isSubmitting,
}: {
  pageKey: PageKey;
  data: WarehouseData;
  onSubmit: (formData: FormData) => void | Promise<void>;
  isSubmitting: boolean;
}) {
  const today = new Date().toISOString().slice(0, 10);
  const supplierDebtOptions = data.suppliers.filter((supplier) => getSupplierDebt(data, supplier.id) > 0);
  const supplierPaymentOptions = supplierDebtOptions.length > 0 ? supplierDebtOptions : data.suppliers;

  if (
    ![
      "products",
      "customers",
      "suppliers",
      "purchases",
      "sales",
      "cashflow",
      "customer-debts",
      "supplier-debts",
    ].includes(pageKey)
  ) {
    return null;
  }

  return (
    <form
      onSubmit={(event) => {
        event.preventDefault();

        if (isSubmitting) {
          return;
        }

        void onSubmit(new FormData(event.currentTarget));
      }}
      aria-busy={isSubmitting}
      className="grid gap-3 rounded-md border border-cyan-100 bg-cyan-50 p-4 sm:grid-cols-2 xl:grid-cols-4"
    >
      {pageKey === "products" ? (
        <>
          <Field label="Mã hàng"><Input name="code" required placeholder="G6060-A02" /></Field>
          <Field label="Tên hàng"><Input name="name" required placeholder="Gạch bóng kính..." /></Field>
          <Field label="Loại"><Input name="category" placeholder="Gạch lát nền" /></Field>
          <Field label="Kích thước"><Input name="size" placeholder="60x60" /></Field>
          <Field label="Viên/thùng"><Input name="piecesPerBox" type="number" min="0" step="1" defaultValue="4" /></Field>
          <Field label="m2/thùng"><Input name="sqmPerBox" type="number" min="0" step="0.01" defaultValue="1.44" /></Field>
          <Field label="Giá nhập"><MoneyInput name="importPrice" defaultValue="0" /></Field>
          <Field label="Giá bán"><MoneyInput name="salePrice" defaultValue="0" /></Field>
          <Field label="Nhà cung cấp">
            <SelectField name="supplierId" defaultValue={data.suppliers[0]?.id}>
              {data.suppliers.map((supplier) => <option key={supplier.id} value={supplier.id}>{supplier.name}</option>)}
            </SelectField>
          </Field>
          <Field label="Tồn tối thiểu"><Input name="minStock" type="number" min="0" defaultValue="0" /></Field>
          <Field label="Ghi chú"><Input name="note" placeholder="Ghi chú" /></Field>
          <Field label="Ảnh mẫu">
            <Input name="imageFile" type="file" accept="image/*" />
          </Field>
        </>
      ) : null}

      {pageKey === "customers" ? (
        <>
          <Field label="Tên khách"><Input name="name" required /></Field>
          <Field label="Số điện thoại"><Input name="phone" type="tel" inputMode="tel" /></Field>
          <Field label="Địa chỉ"><Input name="address" /></Field>
          <Field label="Nhóm khách"><Input name="customerGroup" defaultValue="Bán lẻ" /></Field>
          <Field label="Ghi chú"><Input name="note" /></Field>
        </>
      ) : null}

      {pageKey === "suppliers" ? (
        <>
          <Field label="Tên NCC"><Input name="name" required /></Field>
          <Field label="Số điện thoại"><Input name="phone" type="tel" inputMode="tel" /></Field>
          <Field label="Người liên hệ"><Input name="contactPerson" /></Field>
          <Field label="Địa chỉ"><Input name="address" /></Field>
          <Field label="Ghi chú"><Input name="note" /></Field>
        </>
      ) : null}

      {pageKey === "purchases" ? (
        <>
          <Field label="Ngày nhập"><Input name="orderDate" type="date" defaultValue={today} /></Field>
          <Field label="Nhà cung cấp">
            <SelectField name="supplierId" defaultValue={data.suppliers[0]?.id}>
              {data.suppliers.map((supplier) => <option key={supplier.id} value={supplier.id}>{supplier.name}</option>)}
            </SelectField>
          </Field>
          <MultiProductItems products={data.products} mode="purchase" />
          <Field label="Đã trả"><MoneyInput name="paidAmount" defaultValue="0" /></Field>
          <Field label="Thanh toán"><Input name="paymentMethod" defaultValue="Tiền mặt" /></Field>
          <Field label="Ghi chú"><Input name="note" /></Field>
        </>
      ) : null}

      {pageKey === "sales" ? (
        <>
          <Field label="Ngày bán"><Input name="orderDate" type="date" defaultValue={today} /></Field>
          <Field label="Khách hàng">
            <SelectField name="customerId" defaultValue={data.customers[0]?.id}>
              {data.customers.map((customer) => <option key={customer.id} value={customer.id}>{customer.name}</option>)}
            </SelectField>
          </Field>
          <MultiProductItems products={data.products} mode="sale" />
          <Field label="Chiết khấu"><MoneyInput name="discountAmount" defaultValue="0" /></Field>
          <Field label="Phí giao"><MoneyInput name="shippingFee" defaultValue="0" /></Field>
          <Field label="Khách đã trả"><MoneyInput name="paidAmount" defaultValue="0" /></Field>
          <Field label="Thanh toán"><Input name="paymentMethod" defaultValue="Tiền mặt" /></Field>
          <Field label="Ghi chú"><Input name="note" /></Field>
        </>
      ) : null}

      {pageKey === "cashflow" ? (
        <>
          <Field label="Ngày chi"><Input name="expenseDate" type="date" defaultValue={today} /></Field>
          <Field label="Danh mục"><Input name="category" required placeholder="Vận chuyển, bốc xếp..." /></Field>
          <Field label="Số tiền"><MoneyInput name="amount" required /></Field>
          <Field label="Thanh toán"><Input name="paymentMethod" defaultValue="Tiền mặt" /></Field>
          <Field label="Ghi chú"><Input name="note" /></Field>
        </>
      ) : null}

      {pageKey === "customer-debts" ? (
        <>
          <Field label="Ngày trả"><Input name="paymentDate" type="date" defaultValue={today} /></Field>
          <Field label="Khách hàng">
            <SelectField name="customerId" defaultValue={data.customers[0]?.id}>
              {data.customers.map((customer) => (
                <option key={customer.id} value={customer.id}>
                  {customer.name} - nợ {formatCurrency(getCustomerDebt(data, customer.id))}
                </option>
              ))}
            </SelectField>
          </Field>
          <Field label="Số tiền khách trả"><MoneyInput name="amount" required placeholder="100.000" /></Field>
          <Field label="Phương thức"><Input name="method" defaultValue="Tiền mặt" /></Field>
          <Field label="Ghi chú"><Input name="note" placeholder="VD: khách chuyển khoản thêm" /></Field>
        </>
      ) : null}

      {pageKey === "supplier-debts" ? (
        <>
          <Field label="Ngày trả"><Input name="paymentDate" type="date" defaultValue={today} /></Field>
          <Field label="Nhà cung cấp">
            <SelectField name="supplierId" defaultValue={supplierPaymentOptions[0]?.id}>
              {supplierPaymentOptions.map((supplier) => (
                <option key={supplier.id} value={supplier.id}>
                  {supplier.name} - nợ {formatCurrency(getSupplierDebt(data, supplier.id))}
                </option>
              ))}
            </SelectField>
          </Field>
          <Field label="Số tiền trả NCC"><MoneyInput name="amount" required placeholder="1.000.000" /></Field>
          <Field label="Phương thức"><Input name="method" defaultValue="Tiền mặt" /></Field>
          <Field label="Ghi chú"><Input name="note" placeholder="VD: chuyển khoản trả NCC" /></Field>
        </>
      ) : null}

      <div className="flex items-end">
        <Button type="submit" className="w-full" disabled={isSubmitting}>
          {isSubmitting ? (
            <Loader2 className="h-4 w-4 animate-spin" aria-hidden="true" />
          ) : (
            <Plus className="h-4 w-4" aria-hidden="true" />
          )}
          {isSubmitting ? "Đang lưu..." : "Lưu"}
        </Button>
      </div>
    </form>
  );
}

export function LocalWorkspacePage({ pageKey }: { pageKey: PageKey }) {
  const config = modulePages[pageKey];
  const { data, inventory, lastMessage, syncMode, syncStatus, actions } = useWarehouseStore();
  const [showForm, setShowForm] = useState(false);
  const [query, setQuery] = useState("");
  const [dateFilter, setDateFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("Tất cả");
  const [isSaving, setIsSaving] = useState(false);
  const [imageEditor, setImageEditor] = useState<ProductImageTarget | null>(null);
  const [imagePreview, setImagePreview] = useState<ProductImageTarget | null>(null);
  const [imageError, setImageError] = useState<string | null>(null);
  const savingRef = useRef(false);

  const rows = useMemo<Row[]>(() => {
    if (!data) {
      return [];
    }

    const pageRows = getRows(pageKey, data, inventory);
    const normalizedQuery = query.trim().toLowerCase();
    const normalizedStatus = normalizeFilterValue(statusFilter);

    return pageRows.filter((row) => {
      if (dateFilter && row._date !== dateFilter) {
        return false;
      }

      if (normalizedStatus && normalizedStatus !== "tất cả") {
        const statusText = normalizeFilterValue(
          [row._status, row.status, row.type, row.category].filter(Boolean).join(" "),
        );

        if (!statusText.includes(normalizedStatus)) {
          return false;
        }
      }

      if (!normalizedQuery) {
        return true;
      }

      return Object.entries(row)
        .filter(([key]) => !key.startsWith("_"))
        .some(([, value]) => value.toLowerCase().includes(normalizedQuery));
    });
  }, [data, dateFilter, inventory, pageKey, query, statusFilter]);

  if (!data) {
    return (
      <div className="rounded-md border border-slate-200 bg-white p-6 text-sm text-slate-600">
        {syncStatus}
      </div>
    );
  }

  const stats = getStats(pageKey, config, data, inventory);

  async function handleSubmit(formData: FormData) {
    if (savingRef.current) {
      return;
    }

    savingRef.current = true;
    setIsSaving(true);
    setImageError(null);

    try {
      await new Promise<void>((resolve) => {
        window.requestAnimationFrame(() => resolve());
      });

      const preparedFormData = pageKey === "products" ? await attachCompressedImage(formData) : formData;
      const result =
        pageKey === "products"
          ? await actions.addProduct(preparedFormData)
          : pageKey === "customers"
            ? await actions.addCustomer(formData)
            : pageKey === "suppliers"
              ? await actions.addSupplier(formData)
              : pageKey === "purchases"
                ? await actions.addPurchase(formData)
                : pageKey === "sales"
                  ? await actions.addSale(formData)
                  : pageKey === "cashflow"
                    ? await actions.addExpense(formData)
                    : pageKey === "customer-debts"
                      ? await actions.recordCustomerPayment(formData)
                      : pageKey === "supplier-debts"
                        ? await actions.recordSupplierPayment(formData)
                      : { ok: false, message: "Trang này chưa có form nhập liệu." };

      if (result.ok) {
        setShowForm(false);
      }
    } catch (error) {
      setImageError(error instanceof Error ? error.message : "Không xử lý được ảnh sản phẩm.");
    } finally {
      savingRef.current = false;
      setIsSaving(false);
    }
  }

  async function handleProductImageSubmit(formData: FormData) {
    if (savingRef.current) {
      return;
    }

    savingRef.current = true;
    setIsSaving(true);
    setImageError(null);

    try {
      const preparedFormData = await attachCompressedImage(formData);

      if (!preparedFormData.get("imageData") && !preparedFormData.get("clearImage")) {
        setImageError("Chọn một ảnh sản phẩm trước khi lưu.");
        return;
      }

      const result = await actions.updateProductImage(preparedFormData);

      if (result.ok) {
        setImageEditor(null);
      }
    } catch (error) {
      setImageError(error instanceof Error ? error.message : "Không cập nhật được ảnh sản phẩm.");
    } finally {
      savingRef.current = false;
      setIsSaving(false);
    }
  }

  return (
    <div className="space-y-5">
      <AutoFileBackupAgent data={data} />

      {isSaving ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/20 px-4 backdrop-blur-[1px]">
          <div className="flex w-full max-w-sm items-center gap-3 rounded-md border border-slate-200 bg-white p-4 shadow-soft">
            <span className="flex h-10 w-10 shrink-0 items-center justify-center rounded-md bg-cyan-50 text-cyan-700 ring-1 ring-cyan-100">
              <Loader2 className="h-5 w-5 animate-spin" aria-hidden="true" />
            </span>
            <div>
              <p className="text-sm font-semibold text-slate-950">
                {syncMode === "google-sheet" ? "Đang ghi Google Sheet..." : "Đang lưu dữ liệu..."}
              </p>
              <p className="mt-1 text-sm text-slate-500">
                Vui lòng chờ, hệ thống sẽ tự cập nhật khi xử lý xong.
              </p>
            </div>
          </div>
        </div>
      ) : null}

      <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <p className="text-sm font-medium text-cyan-700">{config.eyebrow}</p>
          <h1 className="mt-1 text-2xl font-semibold text-slate-950">{config.title}</h1>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <Button
            variant="secondary"
            size="sm"
            type="button"
            onClick={() => downloadExcelFile(config.title, config.columns, rows)}
          >
            <Download className="h-4 w-4" aria-hidden="true" />
            {config.exportLabel ?? "Xuất dữ liệu"}
          </Button>
          {[
            "products",
            "customers",
            "suppliers",
            "purchases",
            "sales",
            "cashflow",
            "customer-debts",
            "supplier-debts",
          ].includes(pageKey) ? (
            <Button size="sm" type="button" onClick={() => setShowForm((value) => !value)} disabled={isSaving}>
              <Plus className="h-4 w-4" aria-hidden="true" />
              {config.actionLabel}
            </Button>
          ) : null}
        </div>
      </div>

      {lastMessage ? (
        <div className="rounded-md border border-cyan-200 bg-cyan-50 px-3 py-2 text-sm font-medium text-cyan-800">
          {lastMessage}
        </div>
      ) : null}

      {imageError ? (
        <div className="rounded-md border border-red-200 bg-red-50 px-3 py-2 text-sm font-medium text-red-700">
          {imageError}
        </div>
      ) : null}

      {showForm ? (
        <EntryForm pageKey={pageKey} data={data} onSubmit={handleSubmit} isSubmitting={isSaving} />
      ) : null}

      {pageKey === "settings" ? (
        <LocalModeTools
          data={data}
          disabled={isSaving}
          onRestore={actions.restoreBackup}
          syncMode={syncMode}
        />
      ) : null}

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        {stats.map((stat) => (
          <div key={stat.label} className="rounded-md border border-slate-200 bg-white p-4">
            <p className="text-sm text-slate-500">{stat.label}</p>
            <div className="mt-2 flex items-center justify-between gap-3">
              <p className="text-xl font-semibold text-slate-950">{stat.value}</p>
              <Badge tone={stat.tone ?? "slate"}>{syncMode === "google-sheet" ? "Sheet" : "Local"}</Badge>
            </div>
          </div>
        ))}
      </div>

      <section className="rounded-md border border-slate-200 bg-white shadow-soft">
        <div className="grid gap-3 border-b border-slate-200 p-4 lg:grid-cols-[minmax(220px,1fr)_180px_180px_auto]">
          <div className="relative">
            <Search className="pointer-events-none absolute left-3 top-3 h-4 w-4 text-slate-400" />
            <Input
              className="pl-9"
              placeholder={config.searchPlaceholder}
              value={query}
              onChange={(event) => setQuery(event.target.value)}
            />
          </div>
          <div className="relative">
            <CalendarDays className="pointer-events-none absolute left-3 top-3 h-4 w-4 text-slate-400" />
            <Input
              className="pl-9"
              type="date"
              value={dateFilter}
              onChange={(event) => setDateFilter(event.target.value)}
            />
          </div>
          <label className="relative block">
            <Filter className="pointer-events-none absolute left-3 top-3 h-4 w-4 text-slate-400" />
            <select
              className="h-10 w-full rounded-md border border-input bg-white pl-9 pr-3 text-sm text-slate-800 outline-none focus:border-cyan-600 focus:ring-2 focus:ring-cyan-600/15"
              value={statusFilter}
              onChange={(event) => setStatusFilter(event.target.value)}
            >
              {(config.statusFilters ?? ["Tất cả"]).map((filter) => (
                <option key={filter} value={filter}>{filter}</option>
              ))}
            </select>
          </label>
          <Button
            variant="secondary"
            type="button"
            onClick={() => {
              setQuery("");
              setDateFilter("");
              setStatusFilter("Tất cả");
            }}
          >
            Xóa lọc
          </Button>
        </div>

        <div className="table-scroll overflow-x-auto">
          <table className="w-full min-w-[760px] border-separate border-spacing-0 text-sm">
            <thead>
              <tr className="bg-slate-50 text-left text-xs uppercase text-slate-500">
                {config.columns.map((column) => (
                  <th
                    key={column.key}
                    className={cn(
                      "border-b border-slate-200 px-4 py-3 font-semibold",
                      column.align === "right" && "text-right",
                      column.align === "center" && "text-center",
                    )}
                  >
                    {column.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.length > 0 ? (
                rows.map((row, index) => (
                  <tr key={index}>
                    {config.columns.map((column) => (
                      <td
                        key={column.key}
                        className={cn(
                          "border-b border-slate-100 px-4 py-3 text-slate-700",
                          column.align === "right" && "text-right tabular-nums",
                          column.align === "center" && "text-center",
                        )}
                      >
                        {column.key === "image" ? (
                          <ProductImageCell row={row} onEdit={setImageEditor} onPreview={setImagePreview} />
                        ) : (
                          row[column.key] ?? "-"
                        )}
                      </td>
                    ))}
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={config.columns.length} className="px-4 py-12 text-center">
                    <p className="text-sm font-medium text-slate-800">{config.emptyLabel}</p>
                    <p className="mt-1 text-sm text-slate-500">
                      Bấm nút thêm nhanh để tạo dữ liệu local đầu tiên.
                    </p>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>

      {imageEditor ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/30 px-4 backdrop-blur-[1px]">
          <form
            onSubmit={(event) => {
              event.preventDefault();
              void handleProductImageSubmit(new FormData(event.currentTarget));
            }}
            className="w-full max-w-lg rounded-md border border-slate-200 bg-white p-4 shadow-soft"
          >
            <input type="hidden" name="productId" value={imageEditor.id} />
            <div className="flex items-start justify-between gap-3">
              <div>
                <p className="text-sm font-medium text-cyan-700">Ảnh sản phẩm</p>
                <h2 className="mt-1 text-lg font-semibold text-slate-950">
                  {imageEditor.code} - {imageEditor.name}
                </h2>
              </div>
              <Button type="button" variant="secondary" size="icon" onClick={() => setImageEditor(null)}>
                <X className="h-4 w-4" aria-hidden="true" />
              </Button>
            </div>

            {imageEditor.imageUrl ? (
              <button
                type="button"
                onClick={() => setImagePreview(imageEditor)}
                className="mt-4 block aspect-[4/3] w-full overflow-hidden rounded-md border border-slate-200 bg-slate-50"
              >
                <img src={imageEditor.imageUrl} alt={imageEditor.name} className="h-full w-full object-contain" />
              </button>
            ) : null}

            <div className="mt-4">
              <Field label="Chọn ảnh từ máy">
                <Input name="imageFile" type="file" accept="image/*" />
              </Field>
            </div>

            <div className="mt-4 flex flex-wrap justify-end gap-2">
              {imageEditor.imageUrl ? (
                <Button
                  type="button"
                  variant="secondary"
                  disabled={isSaving}
                  onClick={() => {
                    const formData = new FormData();
                    formData.set("productId", imageEditor.id);
                    formData.set("clearImage", "1");
                    void handleProductImageSubmit(formData);
                  }}
                >
                  <Trash2 className="h-4 w-4" aria-hidden="true" />
                  Gỡ ảnh
                </Button>
              ) : null}
              <Button type="submit" disabled={isSaving}>
                {isSaving ? (
                  <Loader2 className="h-4 w-4 animate-spin" aria-hidden="true" />
                ) : (
                  <Upload className="h-4 w-4" aria-hidden="true" />
                )}
                {isSaving ? "Đang lưu..." : "Lưu ảnh"}
              </Button>
            </div>
          </form>
        </div>
      ) : null}

      {imagePreview ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/70 px-4 py-6">
          <div className="w-full max-w-4xl">
            <div className="mb-3 flex items-center justify-between gap-3 text-white">
              <p className="text-sm font-medium">
                {imagePreview.code} - {imagePreview.name}
              </p>
              <Button type="button" variant="secondary" size="icon" onClick={() => setImagePreview(null)}>
                <X className="h-4 w-4" aria-hidden="true" />
              </Button>
            </div>
            <div className="max-h-[78vh] overflow-hidden rounded-md bg-white p-2">
              <img src={imagePreview.imageUrl} alt={imagePreview.name} className="mx-auto max-h-[74vh] w-auto max-w-full object-contain" />
            </div>
          </div>
        </div>
      ) : null}

    </div>
  );
}
