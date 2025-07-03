import React, { useState } from 'react';
import { CheckSquare, Search, Filter, Calendar, DollarSign, AlertTriangle, CheckCircle, Download, FileSpreadsheet } from 'lucide-react';
import { Operation, Client, ReceivedPayment } from '../../types';
import { formatCurrency, formatDate } from '../../utils/calculations';
import { exportChecksAndPaymentsToPDF, exportChecksAndPaymentsToExcel } from '../../utils/exportUtils';

interface ChecksManagerProps {
  operations: Operation[];
  clients: Client[];
}

const ChecksManager: React.FC<ChecksManagerProps> = ({ operations, clients }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<'all' | 'received' | 'pending'>('all');
  const [typeFilter, setTypeFilter] = useState<'all' | 'check' | 'cash'>('all');

  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  // Collect all received payments from all operations
  const allReceivedPayments = operations.flatMap(operation => 
    operation.receivedPayments.map(payment => ({
      ...payment,
      operationId: operation.id,
      operationName: operation.name,
      operationCode: operation.code,
      clientId: operation.clientId,
      clientName: getClientName(operation.clientId)
    }))
  );

  const filteredPayments = allReceivedPayments.filter(payment => {
    const matchesSearch = (payment.checkNumber || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
                         (payment.bank || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
                         payment.clientName.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         payment.operationName.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         payment.operationCode.toLowerCase().includes(searchTerm.toLowerCase());
    
    const matchesType = typeFilter === 'all' || payment.type === typeFilter;
    
    const matchesStatus = statusFilter === 'all' || 
                         (statusFilter === 'received' && payment.receiptDate) ||
                         (statusFilter === 'pending' && !payment.receiptDate && payment.type === 'check');
    
    return matchesSearch && matchesType && matchesStatus;
  });

  const totalAmount = allReceivedPayments.reduce((sum, payment) => sum + payment.amount, 0);
  const totalChecks = allReceivedPayments.filter(p => p.type === 'check').length;
  const totalCash = allReceivedPayments.filter(p => p.type === 'cash').length;
  const pendingChecks = allReceivedPayments.filter(p => p.type === 'check' && !p.receiptDate).length;

  const handleExportPDF = () => {
    try {
      exportChecksAndPaymentsToPDF(operations, clients);
    } catch (error) {
      console.error('خطأ في تصدير PDF:', error);
      alert('حدث خطأ أثناء تصدير ملف PDF');
    }
  };

  const handleExportExcel = () => {
    try {
      exportChecksAndPaymentsToExcel(operations, clients);
    } catch (error) {
      console.error('خطأ في تصدير Excel:', error);
      alert('حدث خطأ أثناء تصدير ملف Excel');
    }
  };

  return (
    <div className="p-6">
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-gray-900 mb-2">إدارة الشيكات والمدفوعات</h2>
        <p className="text-gray-600">متابعة وإدارة جميع المدفوعات المستلمة</p>
      </div>

      {/* Summary Cards */}
      <div className="mb-8 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
        <div className="bg-blue-50 p-6 rounded-lg border border-blue-200">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-blue-600">إجمالي المدفوعات</p>
              <p className="text-2xl font-bold text-blue-900">{formatCurrency(totalAmount)}</p>
            </div>
            <DollarSign className="w-8 h-8 text-blue-600" />
          </div>
        </div>

        <div className="bg-green-50 p-6 rounded-lg border border-green-200">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-green-600">الشيكات</p>
              <p className="text-2xl font-bold text-green-900">{totalChecks}</p>
            </div>
            <CheckSquare className="w-8 h-8 text-green-600" />
          </div>
        </div>

        <div className="bg-purple-50 p-6 rounded-lg border border-purple-200">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-purple-600">المدفوعات النقدية</p>
              <p className="text-2xl font-bold text-purple-900">{totalCash}</p>
            </div>
            <DollarSign className="w-8 h-8 text-purple-600" />
          </div>
        </div>

        <div className="bg-yellow-50 p-6 rounded-lg border border-yellow-200">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-yellow-600">شيكات معلقة</p>
              <p className="text-2xl font-bold text-yellow-900">{pendingChecks}</p>
            </div>
            <AlertTriangle className="w-8 h-8 text-yellow-600" />
          </div>
        </div>
      </div>

      {/* Export Buttons */}
      <div className="mb-6 flex gap-4">
        <button
          onClick={handleExportPDF}
          className="flex items-center gap-2 px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700 transition-colors"
        >
          <Download className="w-4 h-4" />
          تصدير PDF
        </button>
        <button
          onClick={handleExportExcel}
          className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 transition-colors"
        >
          <FileSpreadsheet className="w-4 h-4" />
          تصدير Excel
        </button>
      </div>

      {/* Filters */}
      <div className="mb-6 bg-white p-4 rounded-lg border border-gray-200">
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 w-4 h-4" />
            <input
              type="text"
              placeholder="البحث..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-10 pr-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>

          <select
            value={typeFilter}
            onChange={(e) => setTypeFilter(e.target.value as 'all' | 'check' | 'cash')}
            className="px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
          >
            <option value="all">جميع الأنواع</option>
            <option value="check">شيكات</option>
            <option value="cash">نقدي</option>
          </select>

          <select
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value as 'all' | 'received' | 'pending')}
            className="px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
          >
            <option value="all">جميع الحالات</option>
            <option value="received">مستلمة</option>
            <option value="pending">معلقة</option>
          </select>

          <button className="flex items-center gap-2 px-4 py-2 bg-gray-100 text-gray-700 rounded-md hover:bg-gray-200 transition-colors">
            <Filter className="w-4 h-4" />
            مرشحات متقدمة
          </button>
        </div>
      </div>

      {/* Payments Table */}
      <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead className="bg-gray-50">
              <tr>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  النوع
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  المبلغ
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  التاريخ
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  العميل
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  العملية
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  التفاصيل
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  الحالة
                </th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {filteredPayments.map((payment) => (
                <tr key={payment.id} className="hover:bg-gray-50">
                  <td className="px-6 py-4 whitespace-nowrap">
                    <div className="flex items-center gap-2">
                      {payment.type === 'check' ? (
                        <CheckSquare className="w-4 h-4 text-blue-600" />
                      ) : (
                        <DollarSign className="w-4 h-4 text-green-600" />
                      )}
                      <span className="text-sm font-medium text-gray-900">
                        {payment.type === 'check' ? 'شيك' : 'نقدي'}
                      </span>
                    </div>
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">
                    {formatCurrency(payment.amount)}
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                    {formatDate(payment.date)}
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                    {payment.clientName}
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                    <div>
                      <div className="font-medium">{payment.operationName}</div>
                      <div className="text-xs text-gray-500">{payment.operationCode}</div>
                    </div>
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                    {payment.type === 'check' ? (
                      <div>
                        {payment.checkNumber && (
                          <div className="font-medium">رقم: {payment.checkNumber}</div>
                        )}
                        {payment.bank && (
                          <div className="text-xs text-gray-500">{payment.bank}</div>
                        )}
                      </div>
                    ) : (
                      <span className="text-gray-500">-</span>
                    )}
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap">
                    {payment.type === 'check' ? (
                      <span className={`inline-flex items-center gap-1 px-2 py-1 text-xs font-semibold rounded-full ${
                        payment.receiptDate 
                          ? 'bg-green-100 text-green-800' 
                          : 'bg-yellow-100 text-yellow-800'
                      }`}>
                        {payment.receiptDate ? (
                          <>
                            <CheckCircle className="w-3 h-3" />
                            مستلم
                          </>
                        ) : (
                          <>
                            <AlertTriangle className="w-3 h-3" />
                            معلق
                          </>
                        )}
                      </span>
                    ) : (
                      <span className="inline-flex items-center gap-1 px-2 py-1 text-xs font-semibold rounded-full bg-green-100 text-green-800">
                        <CheckCircle className="w-3 h-3" />
                        مستلم
                      </span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {filteredPayments.length === 0 && (
          <div className="text-center py-12">
            <p className="text-gray-500">لا توجد مدفوعات تطابق المعايير المحددة</p>
          </div>
        )}
      </div>
    </div>
  );
};

export default ChecksManager;