"use client";

import {
  Banknote,
  Boxes,
  CircleDollarSign,
  ClipboardList,
  ShieldCheck,
  ShoppingCart,
  Truck,
  WalletCards,
} from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { MetricCard } from "@/components/dashboard/metric-card";
import { AutoFileBackupAgent } from "@/components/local/auto-file-backup-agent";
import { formatCurrency, formatDate, formatNumber } from "@/lib/formatters";
import {
  getCustomerDebt,
  getSupplierDebt,
  getToday,
  sum,
} from "@/lib/local/calculations";
import { useWarehouseStore } from "@/lib/local/store";

export function LocalDashboard() {
  const { data, inventory, syncMode, syncStatus } = useWarehouseStore();

  if (!data) {
    return (
      <div className="rounded-md border border-slate-200 bg-white p-6 text-sm text-slate-600">
        {syncStatus}
      </div>
    );
  }

  const today = getToday();
  const todaySales = data.sales.filter((order) => order.orderDate === today);
  const todayPurchases = data.purchases.filter((order) => order.orderDate === today);
  const todaySaleCodes = new Set(todaySales.map((order) => order.code));
  const todayPurchaseCodes = new Set(todayPurchases.map((order) => order.code));
  const todayPaymentsIn = data.payments.filter(
    (payment) => payment.paymentDate === today && payment.direction === "in",
  );
  const revenueToday = sum(todaySales.map((order) => order.totalAmount));
  const revenueTotal = sum(data.sales.map((order) => order.totalAmount));
  const receivedToday = sum(todayPaymentsIn.map((payment) => payment.amount));
  const customerDebt = sum(data.customers.map((customer) => getCustomerDebt(data, customer.id)));
  const supplierDebt = sum(data.suppliers.map((supplier) => getSupplierDebt(data, supplier.id)));
  const stockValue = sum(inventory.map((item) => item.stockValue));
  const lowStock = inventory.filter(
    (item) => item.stockStatus === "low_stock" || item.stockStatus === "out_of_stock",
  );
  const latestPayments = data.payments.slice(0, 5);

  return (
    <div className="space-y-5">
      <AutoFileBackupAgent data={data} />

      <div>
        <div>
          <p className="text-sm font-medium text-cyan-700">Tổng quan</p>
          <h1 className="mt-1 text-2xl font-semibold text-slate-950">Bảng điều khiển</h1>
        </div>
      </div>

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        <MetricCard
          label="Doanh thu hôm nay"
          value={formatCurrency(revenueToday)}
          hint={`Tổng doanh thu: ${formatCurrency(revenueTotal)}`}
          icon={<CircleDollarSign className="h-5 w-5" aria-hidden="true" />}
          tone="green"
        />
        <MetricCard
          label="Tiền đã thu hôm nay"
          value={formatCurrency(receivedToday)}
          hint={`${todayPaymentsIn.length} phiếu thu`}
          icon={<WalletCards className="h-5 w-5" aria-hidden="true" />}
          tone="cyan"
        />
        <MetricCard
          label="Công nợ khách hàng"
          value={formatCurrency(customerDebt)}
          hint={`Nợ NCC: ${formatCurrency(supplierDebt)}`}
          icon={<Banknote className="h-5 w-5" aria-hidden="true" />}
          tone="amber"
        />
        <MetricCard
          label="Giá trị hàng tồn"
          value={formatCurrency(stockValue)}
          hint={`${inventory.length} mã hàng`}
          icon={<Boxes className="h-5 w-5" aria-hidden="true" />}
          tone="slate"
        />
      </div>

      <div className="grid gap-5 xl:grid-cols-[1.15fr_0.85fr]">
        <section className="rounded-md border border-slate-200 bg-white shadow-soft">
          <div className="flex items-center justify-between border-b border-slate-200 px-4 py-3">
            <div>
              <h2 className="font-semibold text-slate-950">Giao dịch mới nhất</h2>
              <p className="text-sm text-slate-500">Thu chi, nhập kho, bán hàng</p>
            </div>
            <Badge tone="cyan">{syncMode === "google-sheet" ? "Sheet" : "Local"}</Badge>
          </div>
          <div className="divide-y divide-slate-100">
            {latestPayments.length > 0 ? (
              latestPayments.map((payment) => (
                <div key={payment.id} className="grid gap-2 px-4 py-3 sm:grid-cols-[120px_1fr_auto]">
                  <p className="text-sm font-medium text-slate-500">{formatDate(payment.paymentDate)}</p>
                  <p className="text-sm text-slate-800">
                    {payment.direction === "in" ? "Thu" : "Chi"} {formatCurrency(payment.amount)} -{" "}
                    {payment.note || payment.method}
                  </p>
                  <Badge tone={payment.direction === "in" ? "green" : "red"}>
                    {payment.direction === "in" ? "Thu" : "Chi"}
                  </Badge>
                </div>
              ))
            ) : (
              <div className="px-4 py-8 text-sm text-slate-500">Chưa có giao dịch.</div>
            )}
          </div>
        </section>

        <section className="rounded-md border border-slate-200 bg-white shadow-soft">
          <div className="flex items-center justify-between border-b border-slate-200 px-4 py-3">
            <div>
              <h2 className="font-semibold text-slate-950">Hàng sắp hết</h2>
              <p className="text-sm text-slate-500">Theo tồn tối thiểu</p>
            </div>
            <Badge tone="amber">Cảnh báo</Badge>
          </div>
          <div className="divide-y divide-slate-100">
            {lowStock.length > 0 ? (
              lowStock.slice(0, 6).map((item) => (
                <div key={item.id} className="grid gap-2 px-4 py-3 sm:grid-cols-[1fr_auto]">
                  <div>
                    <p className="text-sm font-semibold text-slate-950">{item.code}</p>
                    <p className="mt-1 text-sm text-slate-500">{item.name}</p>
                  </div>
                  <div className="text-left sm:text-right">
                    <p className="text-sm font-semibold text-red-600">
                      {formatNumber(item.currentStock)} thùng
                    </p>
                    <p className="mt-1 text-xs text-slate-500">Tối thiểu {formatNumber(item.minStock)}</p>
                  </div>
                </div>
              ))
            ) : (
              <div className="px-4 py-8 text-sm text-slate-500">Không có hàng dưới tồn tối thiểu.</div>
            )}
          </div>
        </section>
      </div>

      <div className="grid gap-5 xl:grid-cols-3">
        {[
          {
            title: "Đơn bán hôm nay",
            value: String(todaySaleCodes.size),
            icon: ShoppingCart,
          },
          {
            title: "Phiếu nhập hôm nay",
            value: String(todayPurchaseCodes.size),
            icon: Truck,
          },
          {
            title: "Nhật ký thao tác",
            value: String(data.auditLogs.length),
            icon: ShieldCheck,
          },
        ].map((item) => {
          const Icon = item.icon;

          return (
            <section key={item.title} className="rounded-md border border-slate-200 bg-white p-4">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm text-slate-500">{item.title}</p>
                  <p className="mt-2 text-2xl font-semibold text-slate-950">{item.value}</p>
                </div>
                <span className="flex h-10 w-10 items-center justify-center rounded-md bg-slate-50 text-slate-700 ring-1 ring-slate-200">
                  <Icon className="h-5 w-5" aria-hidden="true" />
                </span>
              </div>
            </section>
          );
        })}
      </div>

      <section className="rounded-md border border-slate-200 bg-white shadow-soft">
        <div className="flex items-center gap-3 border-b border-slate-200 px-4 py-3">
          <ClipboardList className="h-5 w-5 text-cyan-700" aria-hidden="true" />
          <h2 className="font-semibold text-slate-950">Nhật ký mới nhất</h2>
        </div>
        <div className="divide-y divide-slate-100">
          {data.auditLogs.slice(0, 6).map((log) => (
            <div key={log.id} className="grid gap-2 px-4 py-3 sm:grid-cols-[130px_1fr_auto]">
              <p className="text-sm text-slate-500">{formatDate(log.time)}</p>
              <p className="text-sm text-slate-800">{log.note}</p>
              <Badge tone="slate">{log.action}</Badge>
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}
