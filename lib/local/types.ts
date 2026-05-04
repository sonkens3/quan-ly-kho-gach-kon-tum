export type Product = {
  id: string;
  code: string;
  name: string;
  category: string;
  size: string;
  unit: string;
  piecesPerBox: number;
  sqmPerBox: number;
  importPrice: number;
  salePrice: number;
  supplierId: string;
  minStock: number;
  isActive: boolean;
  note: string;
  createdAt: string;
  imageUrl?: string;
};

export type Customer = {
  id: string;
  name: string;
  phone: string;
  address: string;
  customerGroup: string;
  note: string;
  createdAt: string;
};

export type Supplier = {
  id: string;
  name: string;
  phone: string;
  address: string;
  contactPerson: string;
  note: string;
  createdAt: string;
};

export type PurchaseOrder = {
  id: string;
  code: string;
  supplierId: string;
  orderDate: string;
  productId: string;
  quantity: number;
  unitPrice: number;
  totalAmount: number;
  paidAmount: number;
  debtAmount: number;
  paymentMethod: string;
  status: "confirmed" | "canceled";
  note: string;
  createdAt: string;
};

export type SalesOrder = {
  id: string;
  code: string;
  customerId: string;
  orderDate: string;
  productId: string;
  quantity: number;
  unitPrice: number;
  discountAmount: number;
  shippingFee: number;
  subtotalAmount: number;
  totalAmount: number;
  paidAmount: number;
  debtAmount: number;
  paymentMethod: string;
  deliveryStatus: "pending" | "delivering" | "delivered";
  paymentStatus: "unpaid" | "partial" | "paid";
  status: "confirmed" | "completed" | "canceled";
  note: string;
  createdAt: string;
};

export type StockMovement = {
  id: string;
  productId: string;
  movementDate: string;
  quantity: number;
  movementType: "purchase" | "sale" | "sale_cancel" | "purchase_cancel" | "adjustment";
  referenceType: "purchase_order" | "sales_order" | "manual";
  referenceId: string;
  note: string;
  createdAt: string;
};

export type Payment = {
  id: string;
  paymentDate: string;
  direction: "in" | "out";
  partyType: "customer" | "supplier" | "other";
  customerId?: string;
  supplierId?: string;
  referenceType: "purchase_order" | "sales_order" | "expense" | "manual";
  referenceId: string;
  amount: number;
  method: string;
  note: string;
  createdAt: string;
};

export type Expense = {
  id: string;
  category: string;
  amount: number;
  expenseDate: string;
  paymentMethod: string;
  note: string;
  createdAt: string;
};

export type AuditLog = {
  id: string;
  time: string;
  user: string;
  action: string;
  tableName: string;
  recordCode: string;
  note: string;
};

export type WarehouseData = {
  products: Product[];
  customers: Customer[];
  suppliers: Supplier[];
  purchases: PurchaseOrder[];
  sales: SalesOrder[];
  stockMovements: StockMovement[];
  payments: Payment[];
  expenses: Expense[];
  auditLogs: AuditLog[];
};

export type InventoryRow = Product & {
  currentStock: number;
  currentPieces: number;
  currentSqm: number;
  stockValue: number;
  stockStatus: "in_stock" | "low_stock" | "out_of_stock" | "inactive";
};
