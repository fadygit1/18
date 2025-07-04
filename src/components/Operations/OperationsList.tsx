import React, { useState } from 'react';
import { Eye, Edit, Trash2, Filter, Search, X, Calendar, DollarSign } from 'lucide-react';
import { Operation, Client } from '../../types';
import { formatCurrency, formatDate, calculateNetAmount, calculateTotalDeductions } from '../../utils/calculations';

interface OperationsListProps {
  operations: Operation[];
  clients: Client[];
  onEdit: (operation: Operation) => void;
  onDelete: (operationId: string) => void;
  onView: (operation: Operation) => void;
}

interface AdvancedFilters {
  clientId: string;
  status: string;
  startDate: string;
  endDate: string;
  minAmount: string;
  maxAmount: string;
  searchTerm: string;
}

const OperationsList: React.FC<OperationsListProps> = ({
  operations,
  clients,
  onEdit,
  onDelete,
  onView
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [clientFilter, setClientFilter] = useState<string>('all');
  const [showAdvancedFilters, setShowAdvancedFilters] = useState(false);
  
  const [advancedFilters, setAdvancedFilters] = useState<AdvancedFilters>({
    clientId: 'all',
    status: 'all',
    startDate: '',
    endDate: '',
    minAmount: '',
    maxAmount: '',
    searchTerm: ''
  });

  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  const getStatusLabel = (status: Operation['status']) => {
    const statusLabels = {
      'in_progress': 'قيد التنفيذ',
      'completed': 'مكتملة',
      'completed_partial_payment': 'مكتملة - دفع جزئي',
      'completed_full_payment': 'مكتملة ومدفوعة بالكامل'
    };
    return statusLabels[status];
  };

  const getStatusColor = (status: Operation['status']) => {
    const statusColors = {
      'in_progress': 'bg-yellow-100 text-yellow-800',
      'completed': 'bg-gray-100 text-gray-800',
      'completed_partial_payment': 'bg-orange-100 text-orange-800',
      'completed_full_payment': 'bg-green-100 text-green-800'
    };
    return statusColors[status];
  };

  const applyAdvancedFilters = (operation: Operation): boolean => {
    // البحث النصي
    const searchMatch = !advancedFilters.searchTerm || 
      operation.name.toLowerCase().includes(advancedFilters.searchTerm.toLowerCase()) ||
      operation.code.toLowerCase().includes(advancedFilters.searchTerm.toLowerCase()) ||
      getClientName(operation.clientId).toLowerCase().includes(advancedFilters.searchTerm.toLowerCase());

    // فلتر العميل
    const clientMatch = advancedFilters.clientId === 'all' || operation.clientId === advancedFilters.clientId;

    // فلتر الحالة
    const statusMatch = advancedFilters.status === 'all' || operation.status === advancedFilters.status;

    // فلتر التاريخ
    let dateMatch = true;
    if (advancedFilters.startDate) {
      const startDate = new Date(advancedFilters.startDate);
      dateMatch = dateMatch && new Date(operation.createdAt) >= startDate;
    }
    if (advancedFilters.endDate) {
      const endDate = new Date(advancedFilters.endDate);
      endDate.setHours(23, 59, 59, 999); // نهاية اليوم
      dateMatch = dateMatch && new Date(operation.createdAt) <= endDate;
    }

    // فلتر المبلغ
    let amountMatch = true;
    if (advancedFilters.minAmount) {
      const minAmount = parseFloat(advancedFilters.minAmount);
      amountMatch = amountMatch && operation.totalAmount >= minAmount;
    }
    if (advancedFilters.maxAmount) {
      const maxAmount = parseFloat(advancedFilters.maxAmount);
      amountMatch = amountMatch && operation.totalAmount <= maxAmount;
    }

    return searchMatch && clientMatch && statusMatch && dateMatch && amountMatch;
  };

  const filteredOperations = showAdvancedFilters 
    ? operations.filter(applyAdvancedFilters)
    : operations.filter(operation => {
        const matchesSearch = operation.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
                             operation.code.toLowerCase().includes(searchTerm.toLowerCase()) ||
                             getClientName(operation.clientId).toLowerCase().includes(searchTerm.toLowerCase());
        
        const matchesStatus = statusFilter === 'all' || operation.status === statusFilter;
        const matchesClient = clientFilter === 'all' || operation.clientId === clientFilter;
        
        return matchesSearch && matchesStatus && matchesClient;
      });

  const resetAdvancedFilters = () => {
    setAdvancedFilters({
      clientId: 'all',
      status: 'all',
      startDate: '',
      endDate: '',
      minAmount: '',
      maxAmount: '',
      searchTerm: ''
    });
  };

  const handleAdvancedFilterChange = (field: keyof AdvancedFilters, value: string) => {
    setAdvancedFilters(prev => ({
      ...prev,
      [field]: value
    }));
  };

  return (
    <div className="p-6">
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-gray-900 mb-2">استعراض العمليات</h2>
        <p className="text-gray-600">عرض وإدارة جميع العمليات المسجلة في النظام</p>
      </div>

      {/* Basic Filters */}
      {!showAdvancedFilters && (
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
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
              className="px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              <option value="all">جميع الحالات</option>
              <option value="in_progress">قيد التنفيذ</option>
              <option value="completed">مكتملة</option>
              <option value="completed_partial_payment">مكتملة - دفع جزئي</option>
              <option value="completed_full_payment">مكتملة ومدفوعة بالكامل</option>
            </select>

            <select
              value={clientFilter}
              onChange={(e) => setClientFilter(e.target.value)}
              className="px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              <option value="all">جميع العملاء</option>
              {clients.map((client) => (
                <option key={client.id} value={client.id}>
                  {client.name}
                </option>
              ))}
            </select>

            <button 
              onClick={() => setShowAdvancedFilters(true)}
              className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors"
            >
              <Filter className="w-4 h-4" />
              فلاتر متقدمة
            </button>
          </div>
        </div>
      )}

      {/* Advanced Filters */}
      {showAdvancedFilters && (
        <div className="mb-6 bg-white p-6 rounded-lg border border-gray-200 shadow-sm">
          <div className="flex justify-between items-center mb-4">
            <h3 className="text-lg font-semibold text-gray-900 flex items-center gap-2">
              <Filter className="w-5 h-5" />
              الفلاتر المتقدمة
            </h3>
            <button
              onClick={() => setShowAdvancedFilters(false)}
              className="text-gray-400 hover:text-gray-600 p-1 rounded"
            >
              <X className="w-5 h-5" />
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 mb-4">
            {/* البحث النصي */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                البحث في النص
              </label>
              <div className="relative">
                <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 w-4 h-4" />
                <input
                  type="text"
                  placeholder="اسم العملية، الكود، أو العميل..."
                  value={advancedFilters.searchTerm}
                  onChange={(e) => handleAdvancedFilterChange('searchTerm', e.target.value)}
                  className="w-full pl-10 pr-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
              </div>
            </div>

            {/* العميل */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                العميل
              </label>
              <select
                value={advancedFilters.clientId}
                onChange={(e) => handleAdvancedFilterChange('clientId', e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
              >
                <option value="all">جميع العملاء</option>
                {clients.map((client) => (
                  <option key={client.id} value={client.id}>
                    {client.name}
                  </option>
                ))}
              </select>
            </div>

            {/* الحالة */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                حالة العملية
              </label>
              <select
                value={advancedFilters.status}
                onChange={(e) => handleAdvancedFilterChange('status', e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
              >
                <option value="all">جميع الحالات</option>
                <option value="in_progress">قيد التنفيذ</option>
                <option value="completed">مكتملة</option>
                <option value="completed_partial_payment">مكتملة - دفع جزئي</option>
                <option value="completed_full_payment">مكتملة ومدفوعة بالكامل</option>
              </select>
            </div>

            {/* تاريخ البداية */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                من تاريخ
              </label>
              <div className="relative">
                <Calendar className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 w-4 h-4" />
                <input
                  type="date"
                  value={advancedFilters.startDate}
                  onChange={(e) => handleAdvancedFilterChange('startDate', e.target.value)}
                  className="w-full pl-10 pr-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
              </div>
            </div>

            {/* تاريخ النهاية */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                إلى تاريخ
              </label>
              <div className="relative">
                <Calendar className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 w-4 h-4" />
                <input
                  type="date"
                  value={advancedFilters.endDate}
                  onChange={(e) => handleAdvancedFilterChange('endDate', e.target.value)}
                  className="w-full pl-10 pr-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
              </div>
            </div>

            {/* الحد الأدنى للمبلغ */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                الحد الأدنى للمبلغ (ج.م)
              </label>
              <div className="relative">
                <DollarSign className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 w-4 h-4" />
                <input
                  type="number"
                  placeholder="0"
                  value={advancedFilters.minAmount}
                  onChange={(e) => handleAdvancedFilterChange('minAmount', e.target.value)}
                  className="w-full pl-10 pr-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  min="0"
                  step="0.01"
                />
              </div>
            </div>

            {/* الحد الأقصى للمبلغ */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                الحد الأقصى للمبلغ (ج.م)
              </label>
              <div className="relative">
                <DollarSign className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 w-4 h-4" />
                <input
                  type="number"
                  placeholder="∞"
                  value={advancedFilters.maxAmount}
                  onChange={(e) => handleAdvancedFilterChange('maxAmount', e.target.value)}
                  className="w-full pl-10 pr-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  min="0"
                  step="0.01"
                />
              </div>
            </div>
          </div>

          {/* أزرار التحكم */}
          <div className="flex gap-3">
            <button
              onClick={resetAdvancedFilters}
              className="px-4 py-2 bg-gray-100 text-gray-700 rounded-md hover:bg-gray-200 transition-colors"
            >
              مسح الفلاتر
            </button>
            <button
              onClick={() => setShowAdvancedFilters(false)}
              className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors"
            >
              إخفاء الفلاتر المتقدمة
            </button>
          </div>

          {/* عرض عدد النتائج */}
          <div className="mt-4 p-3 bg-blue-50 rounded-lg border border-blue-200">
            <p className="text-sm text-blue-800">
              <span className="font-semibold">{filteredOperations.length}</span> عملية من أصل {operations.length}
            </p>
          </div>
        </div>
      )}

      {/* Operations Table */}
      <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead className="bg-gray-50">
              <tr>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  كود العملية
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  اسم العملية
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  العميل
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  القيمة الإجمالية
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  الخصومات
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  الصافي المستحق
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  المبلغ المحصل
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  المتبقي
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  نسبة الإنجاز
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  الحالة
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  تاريخ الإنشاء
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  الإجراءات
                </th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {filteredOperations.map((operation) => {
                const totalDeductions = calculateTotalDeductions(operation);
                const netAmount = calculateNetAmount(operation);
                const remainingAmount = netAmount - operation.totalReceived;
                
                return (
                  <tr key={operation.id} className="hover:bg-gray-50">
                    <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">
                      {operation.code}
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                      {operation.name}
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                      {getClientName(operation.clientId)}
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                      {formatCurrency(operation.totalAmount)}
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-red-600">
                      {totalDeductions > 0 ? `- ${formatCurrency(totalDeductions)}` : '-'}
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-purple-600">
                      {formatCurrency(netAmount)}
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-green-600">
                      {formatCurrency(operation.totalReceived)}
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-orange-600">
                      {formatCurrency(remainingAmount)}
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                      <div className="flex items-center">
                        <div className="w-16 bg-gray-200 rounded-full h-2 mr-2">
                          <div
                            className="bg-blue-600 h-2 rounded-full"
                            style={{ width: `${operation.overallExecutionPercentage}%` }}
                          ></div>
                        </div>
                        <span className="text-xs font-medium">
                          {operation.overallExecutionPercentage.toFixed(1)}%
                        </span>
                      </div>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap">
                      <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${getStatusColor(operation.status)}`}>
                        {getStatusLabel(operation.status)}
                      </span>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                      {formatDate(operation.createdAt)}
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm font-medium">
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => onView(operation)}
                          className="text-blue-600 hover:text-blue-900"
                          title="عرض التفاصيل"
                        >
                          <Eye className="w-4 h-4" />
                        </button>
                        <button
                          onClick={() => onEdit(operation)}
                          className="text-green-600 hover:text-green-900"
                          title="تعديل"
                        >
                          <Edit className="w-4 h-4" />
                        </button>
                        <button
                          onClick={() => onDelete(operation.id)}
                          className="text-red-600 hover:text-red-900"
                          title="حذف"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {filteredOperations.length === 0 && (
          <div className="text-center py-12">
            <p className="text-gray-500">لا توجد عمليات تطابق المعايير المحددة</p>
          </div>
        )}
      </div>
    </div>
  );
};

export default OperationsList;