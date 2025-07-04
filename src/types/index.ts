export interface Client {
  id: string;
  name: string;
  type: 'owner' | 'main_contractor' | 'consultant'; // نوع العميل
  phone?: string;
  email?: string;
  address?: string;
  contacts: ClientContact[]; // جهات الاتصال
  createdAt: Date;
  updatedAt: Date;
}

export interface ClientContact {
  id: string;
  name: string;
  position: string; // المنصب
  department: 'accounts' | 'engineering' | 'management' | 'other'; // القسم
  phone?: string;
  email?: string;
  isMainContact: boolean; // جهة الاتصال الرئيسية
}

export interface OperationItem {
  id: string;
  code: string;
  description: string;
  amount: number;
  contractNumber?: string;
  contractDate?: Date;
  executionPercentage: number;
  addedAt: Date;
}

export interface Deduction {
  id: string;
  name: string;
  type: 'percentage' | 'fixed';
  value: number;
  isActive: boolean;
}

export interface GuaranteeCheck {
  id: string;
  checkNumber: string;
  amount: number;
  checkDate: Date;
  deliveryDate: Date;
  expiryDate: Date;
  bank: string;
  isReturned: boolean;
  returnDate?: Date;
  relatedTo: 'operation' | 'item'; // مرتبط بالعملية أم ببند محدد
  relatedItemId?: string; // معرف البند إذا كان مرتبط ببند
}

export interface GuaranteeLetter {
  id: string;
  bank: string;
  letterDate: Date;
  letterNumber: string;
  amount: number;
  dueDate: Date;
  renewals: Array<{
    id: string;
    renewalDate: Date;
    newDueDate: Date;
    notes?: string;
  }>;
  relatedTo: 'operation' | 'item';
  relatedItemId?: string;
  isReturned: boolean;
  returnDate?: Date;
  notes?: string;
}

export interface WarrantyCertificate {
  id: string;
  certificateNumber: string;
  issueDate: Date;
  startDate: Date;
  endDate: Date;
  warrantyPeriodMonths: number;
  description: string;
  relatedTo: 'operation' | 'item';
  relatedItemId?: string;
  isActive: boolean;
  notes?: string;
}

export interface ReceivedPayment {
  id: string;
  type: 'check' | 'cash';
  amount: number;
  date: Date;
  checkNumber?: string;
  bank?: string;
  receiptDate?: Date;
  notes?: string;
}

export interface Operation {
  id: string;
  code: string;
  name: string;
  clientId: string;
  items: OperationItem[];
  deductions: Deduction[];
  guaranteeChecks: GuaranteeCheck[];
  guaranteeLetters: GuaranteeLetter[];
  warrantyCertificates: WarrantyCertificate[];
  receivedPayments: ReceivedPayment[];
  totalAmount: number;
  totalReceived: number;
  overallExecutionPercentage: number;
  status: 'in_progress' | 'completed' | 'completed_partial_payment' | 'completed_full_payment';
  createdAt: Date;
  updatedAt: Date;
}

export interface DashboardStats {
  totalOperations: number;
  completedOperations: number;
  inProgressOperations: number;
  totalAmount: number;
  totalReceived: number;
  outstandingGuarantees: number;
  activeWarranties: number;
  totalDeductions?: number;
  totalNetAmount?: number;
}