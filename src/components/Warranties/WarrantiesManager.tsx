import React, { useState } from 'react';
import { Award, Search, Filter, Calendar, AlertTriangle, CheckCircle, Plus, Edit, Trash2, Download, FileSpreadsheet, FileEdit } from 'lucide-react';
import { Operation, Client, WarrantyCertificate } from '../../types';
import { formatDate } from '../../utils/calculations';
import { exportWarrantyCertificatesReportToPDF, exportWarrantyCertificatesToExcel, exportWarrantyCertificatesReportToWord } from '../../utils/exportUtils';

interface WarrantiesManagerProps {
  operations: Operation[];
  clients: Client[];
  onUpdateOperation: (operation: Operation) => void;
}

const WarrantiesManager: React.FC<WarrantiesManagerProps> = ({ operations, clients, onUpdateOperation }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<'all' | 'active' | 'expired'>('all');
  const [showAddForm, setShowAddForm] = useState(false);
  const [editingWarranty, setEditingWarranty] = useState<{ warranty: WarrantyCertificate; operationId: string } | null>(null);
  const [selectedOperation, setSelectedOperation] = useState<string>('');

  const [formData, setFormData] = useState({
    certificateNumber: '',
    issueDate: new Date().toISOString().split('T')[0],
    startDate: new Date().toISOString().split('T')[0],
    warrantyPeriodMonths: 12,
    description: '',
    relatedTo: 'operation' as 'operation' | 'item',
    relatedItemId: '',
    notes: ''
  });

  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  const getOperationName = (operationId: string) => {
    const operation = operations.find(op => op.id === operationId);
    return operation ? operation.name : 'عملية غير معروفة';
  };

  const getItemName = (operationId: string, itemId: string) => {
    const operation = operations.find(op => op.id === operationId);
    if (!operation) return 'بند غير معروف';
    const item = operation.items.find(item => item.id === itemId);
    return item ? item.description : 'بند غير معروف';
  };

  // جمع جميع شهادات الضمان من جميع العمليات
  const allWarranties = operations.flatMap(operation => 
    (operation.warrantyCertificates || []).map(warranty => ({
      ...warranty,
      operationId: operation.id,
      operationName: operation.name,
      operationCode: operation.code,
      clientId: operation.clientId,
      clientName: getClientName(operation.clientId)
    }))
  );

  const isExpired = (endDate: Date) => {
    const today = new Date();
    return new Date(endDate) < today;
  };

  const isExpiringSoon = (endDate: Date) => {
    const today = new Date();
    const diffTime = new Date(endDate).getTime() - today.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return diffDays <= 30 && diffDays > 0;
  };

  const filteredWarranties = allWarranties.filter(warranty => {
    const matchesSearch = warranty.certificateNumber.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         warranty.description.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         warranty.clientName.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         warranty.operationName.toLowerCase().includes(searchTerm.toLowerCase());
    
    const matchesStatus = statusFilter === 'all' || 
                         (statusFilter === 'active' && warranty.isActive && !isExpired(warranty.endDate)) ||
                         (statusFilter === 'expired' && (!warranty.isActive || isExpired(warranty.endDate)));
    
    return matchesSearch && matchesStatus;
  });

  const resetForm = () => {
    setFormData({
      certificateNumber: '',
      issueDate: new Date().toISOString().split('T')[0],
      startDate: new Date().toISOString().split('T')[0],
      warrantyPeriodMonths: 12,
      description: '',
      relatedTo: 'operation',
      relatedItemId: '',
      notes: ''
    });
    setSelectedOperation('');
    setShowAddForm(false);
    setEditingWarranty(null);
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    
    if (!selectedOperation) {
      alert('يرجى اختيار العملية');
      return;
    }

    const operation = operations.find(op => op.id === selectedOperation);
    if (!operation) return;

    const startDate = new Date(formData.startDate);
    const endDate = new Date(startDate);
    endDate.setMonth(endDate.getMonth() + formData.warrantyPeriodMonths);

    const warrantyData: WarrantyCertificate = {
      id: editingWarranty?.warranty.id || crypto.randomUUID(),
      certificateNumber: formData.certificateNumber,
      issueDate: new Date(formData.issueDate),
      startDate: startDate,
      endDate: endDate,
      warrantyPeriodMonths: formData.warrantyPeriodMonths,
      description: formData.description,
      relatedTo: formData.relatedTo,
      relatedItemId: formData.relatedTo === 'item' ? formData.relatedItemId : undefined,
      isActive: true,
      notes: formData.notes
    };

    let updatedWarranties;
    if (editingWarranty) {
      // تحديث شهادة موجودة
      updatedWarranties = (operation.warrantyCertificates || []).map(w => 
        w.id === editingWarranty.warranty.id ? warrantyData : w
      );
    } else {
      // إضافة شهادة جديدة
      updatedWarranties = [...(operation.warrantyCertificates || []), warrantyData];
    }

    const updatedOperation = {
      ...operation,
      warrantyCertificates: updatedWarranties,
      updatedAt: new Date()
    };

    onUpdateOperation(updatedOperation);
    resetForm();
  };

  const handleEdit = (warranty: any) => {
    setEditingWarranty(warranty);
    setSelectedOperation(warranty.operationId);
    setFormData({
      certificateNumber: warranty.certificateNumber,
      issueDate: warranty.issueDate.toISOString().split('T')[0],
      startDate: warranty.startDate.toISOString().split('T')[0],
      warrantyPeriodMonths: warranty.warrantyPeriodMonths,
      description: warranty.description,
      relatedTo: warranty.relatedTo,
      relatedItemId: warranty.relatedItemId || '',
      notes: warranty.notes || ''
    });
    setShowAddForm(true);
  };

  const handleDelete = (warranty: any) => {
    if (window.confirm('هل أنت متأكد من حذف شهادة الضمان؟')) {
      const operation = operations.find(op => op.id === warranty.operationId);
      if (!operation) return;

      const updatedWarranties = (operation.warrantyCertificates || []).filter(w => w.id !== warranty.id);
      const updatedOperation = {
        ...operation,
        warrantyCertificates: updatedWarranties,
        updatedAt: new Date()
      };

      onUpdateOperation(updatedOperation);
    }
  };

  const handleExportPDF = () => {
    try {
      exportWarrantyCertificatesReportToPDF(operations, clients);
    } catch (error) {
      console.error('خطأ في تصدير PDF:', error);
      alert('حدث خطأ أثناء تصدير ملف PDF');
    }
  };

  const handleExportExcel = () => {
    try {
      exportWarrantyCertificatesToExcel(operations, clients);
    } catch (error) {
      console.error('خطأ في تصدير Excel:', error);
      alert('حدث خطأ أثناء تصدير ملف Excel');
    }
  };

  const handleExportWord = () => {
    try {
      exportWarrantyCertificatesReportToWord(operations, clients);
    } catch (error) {
      console.error('خطأ في تصدير Word:', error);
      alert('حدث خطأ أثناء تصدير ملف Word');
    }
  };

  const selectedOperationData = operations.find(op => op.id === selectedOperation);

  return (
    <div className="p-6">
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-gray-900 mb-2">إدارة شهادات الضمان</h2>
        <p className="text-gray-600">متابعة وإدارة شهادات الضمان للعمليات المكتملة</p>
      </div>

      {/* Summary Cards */}
      <div className="mb-8 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
        <div className="bg-green-50 p-6 rounded-lg border border-green-200">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-green-600">إجمالي الشهادات</p>
              <p className="text-2xl font-bold text-green-900">{allWarranties.length}</p>
            </div>
            <Award className="w-8 h-8 text-green-600" />
          </div>
        </div>

        <div className="bg-blue-50 p-6 rounded-lg border border-blue-200">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-blue-600">الشهادات النشطة</p>
              <p className="text-2xl font-bold text-blue-900">
                {allWarranties.filter(w => w.isActive && !isExpired(w.endDate)).length}
              </p>
            </div>
            <CheckCircle className="w-8 h-8 text-blue-600" />
          </div>
        </div>

        <div className="bg-yellow-50 p-6 rounded-lg border border-yellow-200">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-yellow-600">تنتهي قريباً</p>
              <p className="text-2xl font-bold text-yellow-900">
                {allWarranties.filter(w => w.isActive && isExpiringSoon(w.endDate)).length}
              </p>
            </div>
            <AlertTriangle className="w-8 h-8 text-yellow-600" />
          </div>
        </div>

        <div className="bg-red-50 p-6 rounded-lg border border-red-200">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-red-600">منتهية الصلاحية</p>
              <p className="text-2xl font-bold text-red-900">
                {allWarranties.filter(w => !w.isActive || isExpired(w.endDate)).length}
              </p>
            </div>
            <AlertTriangle className="w-8 h-8 text-red-600" />
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
        <button
          onClick={handleExportWord}
          className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors"
        >
          <FileEdit className="w-4 h-4" />
          تصدير Word
        </button>
      </div>

      {/* Header Actions */}
      <div className="mb-6 flex flex-col sm:flex-row gap-4 justify-between">
        <div className="flex gap-4">
          <div className="relative flex-1 max-w-md">
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
            onChange={(e) => setStatusFilter(e.target.value as 'all' | 'active' | 'expired')}
            className="px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
          >
            <option value="all">جميع الشهادات</option>
            <option value="active">الشهادات النشطة</option>
            <option value="expired">الشهادات المنتهية</option>
          </select>
        </div>
        
        <button
          onClick={() => setShowAddForm(true)}
          className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 transition-colors"
        >
          <Plus className="w-4 h-4" />
          إضافة شهادة ضمان
        </button>
      </div>

      {/* Add/Edit Form */}
      {showAddForm && (
        <div className="mb-6 bg-white p-6 rounded-lg border border-gray-200">
          <h3 className="text-lg font-semibold text-gray-900 mb-4">
            {editingWarranty ? 'تعديل شهادة الضمان' : 'إضافة شهادة ضمان جديدة'}
          </h3>
          
          <form onSubmit={handleSubmit} className="space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  العملية *
                </label>
                <select
                  value={selectedOperation}
                  onChange={(e) => setSelectedOperation(e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  required
                  disabled={!!editingWarranty}
                >
                  <option value="">اختر العملية</option>
                  {operations.map((operation) => (
                    <option key={operation.id} value={operation.id}>
                      {operation.name} - {getClientName(operation.clientId)}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  رقم الشهادة *
                </label>
                <input
                  type="text"
                  value={formData.certificateNumber}
                  onChange={(e) => setFormData({ ...formData, certificateNumber: e.target.value })}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  required
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  تاريخ الإصدار *
                </label>
                <input
                  type="date"
                  value={formData.issueDate}
                  onChange={(e) => setFormData({ ...formData, issueDate: e.target.value })}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  required
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  تاريخ بداية الضمان *
                </label>
                <input
                  type="date"
                  value={formData.startDate}
                  onChange={(e) => setFormData({ ...formData, startDate: e.target.value })}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  required
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  مدة الضمان (بالأشهر) *
                </label>
                <input
                  type="number"
                  value={formData.warrantyPeriodMonths}
                  onChange={(e) => setFormData({ ...formData, warrantyPeriodMonths: parseInt(e.target.value) || 0 })}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  min="1"
                  required
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  مرتبط بـ
                </label>
                <select
                  value={formData.relatedTo}
                  onChange={(e) => setFormData({ ...formData, relatedTo: e.target.value as 'operation' | 'item', relatedItemId: '' })}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="operation">العملية كاملة</option>
                  <option value="item">بند محدد</option>
                </select>
              </div>

              {formData.relatedTo === 'item' && selectedOperationData && (
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    البند
                  </label>
                  <select
                    value={formData.relatedItemId}
                    onChange={(e) => setFormData({ ...formData, relatedItemId: e.target.value })}
                    className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  >
                    <option value="">اختر البند</option>
                    {selectedOperationData.items.map((item) => (
                      <option key={item.id} value={item.id}>
                        {item.description}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              <div className="md:col-span-2">
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  وصف الضمان *
                </label>
                <textarea
                  value={formData.description}
                  onChange={(e) => setFormData({ ...formData, description: e.target.value })}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  rows={3}
                  required
                />
              </div>

              <div className="md:col-span-2">
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  ملاحظات
                </label>
                <textarea
                  value={formData.notes}
                  onChange={(e) => setFormData({ ...formData, notes: e.target.value })}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  rows={2}
                />
              </div>
            </div>

            <div className="flex gap-3">
              <button
                type="submit"
                className="px-4 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 transition-colors"
              >
                {editingWarranty ? 'تحديث' : 'إضافة'}
              </button>
              <button
                type="button"
                onClick={resetForm}
                className="px-4 py-2 bg-gray-300 text-gray-700 rounded-md hover:bg-gray-400 transition-colors"
              >
                إلغاء
              </button>
            </div>
          </form>
        </div>
      )}

      {/* Warranties Table */}
      <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead className="bg-gray-50">
              <tr>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  رقم الشهادة
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  العميل
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  العملية
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  البند
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  الوصف
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  تاريخ البداية
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  تاريخ النهاية
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  مدة الضمان
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  الحالة
                </th>
                <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">
                  الإجراءات
                </th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {filteredWarranties.map((warranty) => (
                <tr key={warranty.id} className="hover:bg-gray-50">
                  <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">
                    {warranty.certificateNumber}
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                    {warranty.clientName}
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                    <div>
                      <div className="font-medium">{warranty.operationName}</div>
                      <div className="text-xs text-gray-500">{warranty.operationCode}</div>
                    </div>
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                    {warranty.relatedTo === 'item' && warranty.relatedItemId 
                      ? getItemName(warranty.operationId, warranty.relatedItemId)
                      : 'العملية كاملة'
                    }
                  </td>
                  <td className="px-6 py-4 text-sm text-gray-900 max-w-xs truncate">
                    {warranty.description}
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                    {formatDate(warranty.startDate)}
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                    <div className="flex items-center gap-2">
                      {isExpired(warranty.endDate) && (
                        <AlertTriangle className="w-4 h-4 text-red-500" />
                      )}
                      {isExpiringSoon(warranty.endDate) && !isExpired(warranty.endDate) && (
                        <AlertTriangle className="w-4 h-4 text-yellow-500" />
                      )}
                      <span className={
                        isExpired(warranty.endDate) ? 'text-red-600' :
                        isExpiringSoon(warranty.endDate) ? 'text-yellow-600' :
                        'text-gray-900'
                      }>
                        {formatDate(warranty.endDate)}
                      </span>
                    </div>
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                    {warranty.warrantyPeriodMonths} شهر
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap">
                    <span className={`inline-flex items-center gap-1 px-2 py-1 text-xs font-semibold rounded-full ${
                      warranty.isActive && !isExpired(warranty.endDate)
                        ? isExpiringSoon(warranty.endDate)
                          ? 'bg-yellow-100 text-yellow-800'
                          : 'bg-green-100 text-green-800'
                        : 'bg-red-100 text-red-800'
                    }`}>
                      {warranty.isActive && !isExpired(warranty.endDate) ? (
                        isExpiringSoon(warranty.endDate) ? (
                          <>
                            <AlertTriangle className="w-3 h-3" />
                            ينتهي قريباً
                          </>
                        ) : (
                          <>
                            <CheckCircle className="w-3 h-3" />
                            نشط
                          </>
                        )
                      ) : (
                        <>
                          <AlertTriangle className="w-3 h-3" />
                          منتهي
                        </>
                      )}
                    </span>
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm font-medium">
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => handleEdit(warranty)}
                        className="text-blue-600 hover:text-blue-900"
                        title="تعديل"
                      >
                        <Edit className="w-4 h-4" />
                      </button>
                      <button
                        onClick={() => handleDelete(warranty)}
                        className="text-red-600 hover:text-red-900"
                        title="حذف"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {filteredWarranties.length === 0 && (
          <div className="text-center py-12">
            <p className="text-gray-500">لا توجد شهادات ضمان تطابق المعايير المحددة</p>
          </div>
        )}
      </div>
    </div>
  );
};

export default WarrantiesManager;