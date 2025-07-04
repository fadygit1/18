import React from 'react';
import { X, Calendar, DollarSign, User, FileText, CreditCard, CheckCircle, AlertTriangle, Download, Minus, FileEdit } from 'lucide-react';
import { Operation, Client } from '../../types';
import { formatCurrency, formatDate, calculateNetAmount, calculateTotalDeductions, calculateExecutedTotal } from '../../utils/calculations';
import { exportOperationDetailsToPDF, exportOperationDetailsToWord } from '../../utils/exportUtils';

interface OperationDetailsProps {
  operation: Operation;
  client: Client;
  onClose: () => void;
}

const OperationDetails: React.FC<OperationDetailsProps> = ({ operation, client, onClose }) => {
  const executedAmount = calculateExecutedTotal(operation.items);
  const totalDeductions = calculateTotalDeductions(operation);
  const netAmount = calculateNetAmount(operation);
  const remainingAmount = netAmount - operation.totalReceived;

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

  const handleExportPDF = () => {
    try {
      exportOperationDetailsToPDF(operation, client);
    } catch (error) {
      console.error('خطأ في تصدير PDF:', error);
      alert('حدث خطأ أثناء تصدير ملف PDF');
    }
  };

  const handleExportWord = () => {
    try {
      exportOperationDetailsToWord(operation, client);
    } catch (error) {
      console.error('خطأ في تصدير Word:', error);
      alert('حدث خطأ أثناء تصدير ملف Word');
    }
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-lg max-w-6xl w-full max-h-[90vh] overflow-y-auto">
        <div className="sticky top-0 bg-white border-b border-gray-200 p-6 flex justify-between items-center">
          <div>
            <h2 className="text-2xl font-bold text-gray-900">{operation.name}</h2>
            <div className="flex items-center gap-4 mt-2">
              <p className="text-gray-600">كود العملية: {operation.code}</p>
              <span className={`inline-flex px-3 py-1 text-sm font-semibold rounded-full ${getStatusColor(operation.status)}`}>
                {getStatusLabel(operation.status)}
              </span>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <button
              onClick={handleExportPDF}
              className="flex items-center gap-2 px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 transition-colors"
            >
              <Download className="w-4 h-4" />
              تصدير HTML
            </button>
            <button
              onClick={handleExportWord}
              className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
            >
              <FileEdit className="w-4 h-4" />
              تصدير Word
            </button>
            <button
              onClick={onClose}
              className="text-gray-400 hover:text-gray-600 p-2 rounded-lg hover:bg-gray-100"
            >
              <X className="w-6 h-6" />
            </button>
          </div>
        </div>

        <div className="p-6 space-y-8">
          {/* Basic Information */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
            <div className="bg-blue-50 p-4 rounded-lg border border-blue-200">
              <div className="flex items-center gap-3">
                <User className="w-8 h-8 text-blue-600" />
                <div>
                  <p className="text-sm text-blue-600 font-medium">العميل</p>
                  <p className="text-lg font-bold text-blue-900">{client.name}</p>
                </div>
              </div>
            </div>

            <div className="bg-green-50 p-4 rounded-lg border border-green-200">
              <div className="flex items-center gap-3">
                <DollarSign className="w-8 h-8 text-green-600" />
                <div>
                  <p className="text-sm text-green-600 font-medium">القيمة الإجمالية</p>
                  <p className="text-lg font-bold text-green-900">{formatCurrency(operation.totalAmount)}</p>
                </div>
              </div>
            </div>

            <div className="bg-yellow-50 p-4 rounded-lg border border-yellow-200">
              <div className="flex items-center gap-3">
                <CheckCircle className="w-8 h-8 text-yellow-600" />
                <div>
                  <p className="text-sm text-yellow-600 font-medium">المبلغ المنفذ</p>
                  <p className="text-lg font-bold text-yellow-900">{formatCurrency(executedAmount)}</p>
                </div>
              </div>
            </div>

            <div className="bg-red-50 p-4 rounded-lg border border-red-200">
              <div className="flex items-center gap-3">
                <Minus className="w-8 h-8 text-red-600" />
                <div>
                  <p className="text-sm text-red-600 font-medium">إجمالي الخصومات</p>
                  <p className="text-lg font-bold text-red-900">{formatCurrency(totalDeductions)}</p>
                </div>
              </div>
            </div>

            <div className="bg-purple-50 p-4 rounded-lg border border-purple-200">
              <div className="flex items-center gap-3">
                <CheckCircle className="w-8 h-8 text-purple-600" />
                <div>
                  <p className="text-sm text-purple-600 font-medium">الصافي المستحق</p>
                  <p className="text-lg font-bold text-purple-900">{formatCurrency(netAmount)}</p>
                </div>
              </div>
            </div>

            <div className="bg-indigo-50 p-4 rounded-lg border border-indigo-200">
              <div className="flex items-center gap-3">
                <DollarSign className="w-8 h-8 text-indigo-600" />
                <div>
                  <p className="text-sm text-indigo-600 font-medium">المبلغ المحصل</p>
                  <p className="text-lg font-bold text-indigo-900">{formatCurrency(operation.totalReceived)}</p>
                </div>
              </div>
            </div>

            <div className="bg-orange-50 p-4 rounded-lg border border-orange-200">
              <div className="flex items-center gap-3">
                <AlertTriangle className="w-8 h-8 text-orange-600" />
                <div>
                  <p className="text-sm text-orange-600 font-medium">المبلغ المتبقي</p>
                  <p className="text-lg font-bold text-orange-900">{formatCurrency(remainingAmount)}</p>
                </div>
              </div>
            </div>

            <div className="bg-gray-50 p-4 rounded-lg border border-gray-200">
              <div className="flex items-center gap-3">
                <CheckCircle className="w-8 h-8 text-gray-600" />
                <div>
                  <p className="text-sm text-gray-600 font-medium">نسبة التحصيل</p>
                  <p className="text-lg font-bold text-gray-900">
                    {netAmount > 0 ? ((operation.totalReceived / netAmount) * 100).toFixed(1) : 0}%
                  </p>
                </div>
              </div>
            </div>
          </div>

          {/* Financial Summary */}
          <div className="bg-white border border-gray-200 rounded-lg p-6">
            <h3 className="text-xl font-semibold text-gray-900 mb-4">الملخص المالي</h3>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="space-y-3">
                <div className="flex justify-between items-center">
                  <span className="text-gray-600">القيمة الإجمالية:</span>
                  <span className="font-bold text-gray-900">{formatCurrency(operation.totalAmount)}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-gray-600">المبلغ المنفذ:</span>
                  <span className="font-bold text-yellow-600">{formatCurrency(executedAmount)}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-gray-600">إجمالي الخصومات:</span>
                  <span className="font-bold text-red-600">- {formatCurrency(totalDeductions)}</span>
                </div>
                <div className="flex justify-between items-center pt-3 border-t border-gray-200">
                  <span className="text-gray-600 font-medium">الصافي المستحق:</span>
                  <span className="font-bold text-purple-600">{formatCurrency(netAmount)}</span>
                </div>
              </div>
              <div className="space-y-3">
                <div className="flex justify-between items-center">
                  <span className="text-gray-600">المبلغ المحصل:</span>
                  <span className="font-bold text-green-600">{formatCurrency(operation.totalReceived)}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-gray-600">المبلغ المتبقي:</span>
                  <span className="font-bold text-orange-600">{formatCurrency(remainingAmount)}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-gray-600">نسبة التحصيل:</span>
                  <span className="font-bold text-blue-600">
                    {netAmount > 0 ? ((operation.totalReceived / netAmount) * 100).toFixed(1) : 0}%
                  </span>
                </div>
                <div className="flex justify-between items-center pt-3 border-t border-gray-200">
                  <span className="text-gray-600 font-medium">نسبة التنفيذ:</span>
                  <span className="font-bold text-indigo-600">{operation.overallExecutionPercentage.toFixed(1)}%</span>
                </div>
              </div>
            </div>
          </div>

          {/* Operation Items */}
          <div className="bg-white border border-gray-200 rounded-lg">
            <div className="p-6 border-b border-gray-200">
              <h3 className="text-xl font-semibold text-gray-900 flex items-center gap-2">
                <FileText className="w-5 h-5" />
                بنود العملية
              </h3>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead className="bg-gray-50">
                  <tr>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">الكود</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">الوصف</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">القيمة</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">نسبة التنفيذ</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">القيمة المنفذة</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-200">
                  {operation.items.map((item) => (
                    <tr key={item.id}>
                      <td className="px-6 py-4 text-sm font-medium text-gray-900">{item.code}</td>
                      <td className="px-6 py-4 text-sm text-gray-900">{item.description}</td>
                      <td className="px-6 py-4 text-sm text-gray-900">{formatCurrency(item.amount)}</td>
                      <td className="px-6 py-4 text-sm text-gray-900">{item.executionPercentage}%</td>
                      <td className="px-6 py-4 text-sm text-gray-900">
                        {formatCurrency(item.amount * (item.executionPercentage / 100))}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Deductions */}
          {operation.deductions.length > 0 && (
            <div className="bg-white border border-gray-200 rounded-lg">
              <div className="p-6 border-b border-gray-200">
                <h3 className="text-xl font-semibold text-gray-900 flex items-center gap-2">
                  <Minus className="w-5 h-5" />
                  الخصومات
                </h3>
              </div>
              <div className="p-6">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {operation.deductions.filter(d => d.isActive).map((deduction) => {
                    const executedTotal = calculateExecutedTotal(operation.items);
                    const deductionAmount = deduction.type === 'percentage' 
                      ? (executedTotal * deduction.value / 100)
                      : deduction.value;
                    
                    return (
                      <div key={deduction.id} className="flex justify-between items-center p-3 bg-red-50 rounded-lg border border-red-200">
                        <div>
                          <span className="text-sm font-medium text-gray-900">{deduction.name}</span>
                          <div className="text-xs text-gray-500">
                            {deduction.type === 'percentage' ? `${deduction.value}%` : 'مبلغ ثابت'}
                          </div>
                        </div>
                        <span className="text-sm font-bold text-red-600">
                          - {formatCurrency(deductionAmount)}
                        </span>
                      </div>
                    );
                  })}
                </div>
                <div className="mt-4 pt-4 border-t border-gray-200">
                  <div className="flex justify-between items-center">
                    <span className="font-medium text-gray-900">إجمالي الخصومات:</span>
                    <span className="font-bold text-red-600">- {formatCurrency(totalDeductions)}</span>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Guarantee Checks */}
          {operation.guaranteeChecks.length > 0 && (
            <div className="bg-white border border-gray-200 rounded-lg">
              <div className="p-6 border-b border-gray-200">
                <h3 className="text-xl font-semibold text-gray-900 flex items-center gap-2">
                  <CreditCard className="w-5 h-5" />
                  شيكات الضمان
                </h3>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full">
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">رقم الشيك</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">المبلغ</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">البنك</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">تاريخ الانتهاء</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">الحالة</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-200">
                    {operation.guaranteeChecks.map((check) => (
                      <tr key={check.id}>
                        <td className="px-6 py-4 text-sm font-medium text-gray-900">{check.checkNumber}</td>
                        <td className="px-6 py-4 text-sm text-gray-900">{formatCurrency(check.amount)}</td>
                        <td className="px-6 py-4 text-sm text-gray-900">{check.bank}</td>
                        <td className="px-6 py-4 text-sm text-gray-900">{formatDate(check.expiryDate)}</td>
                        <td className="px-6 py-4">
                          <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${
                            check.isReturned 
                              ? 'bg-green-100 text-green-800' 
                              : 'bg-yellow-100 text-yellow-800'
                          }`}>
                            {check.isReturned ? 'مُسترد' : 'قائم'}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Guarantee Letters */}
          {operation.guaranteeLetters.length > 0 && (
            <div className="bg-white border border-gray-200 rounded-lg">
              <div className="p-6 border-b border-gray-200">
                <h3 className="text-xl font-semibold text-gray-900 flex items-center gap-2">
                  <FileText className="w-5 h-5" />
                  خطابات الضمان
                </h3>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full">
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">رقم الخطاب</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">البنك</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">المبلغ</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">تاريخ الاستحقاق</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">الحالة</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-200">
                    {operation.guaranteeLetters.map((letter) => (
                      <tr key={letter.id}>
                        <td className="px-6 py-4 text-sm font-medium text-gray-900">{letter.letterNumber}</td>
                        <td className="px-6 py-4 text-sm text-gray-900">{letter.bank}</td>
                        <td className="px-6 py-4 text-sm text-gray-900">{formatCurrency(letter.amount)}</td>
                        <td className="px-6 py-4 text-sm text-gray-900">{formatDate(letter.dueDate)}</td>
                        <td className="px-6 py-4">
                          <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${
                            letter.isReturned 
                              ? 'bg-green-100 text-green-800' 
                              : 'bg-yellow-100 text-yellow-800'
                          }`}>
                            {letter.isReturned ? 'مُسترد' : 'قائم'}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Warranty Certificates */}
          {(operation.warrantyCertificates || []).length > 0 && (
            <div className="bg-white border border-gray-200 rounded-lg">
              <div className="p-6 border-b border-gray-200">
                <h3 className="text-xl font-semibold text-gray-900 flex items-center gap-2">
                  <CheckCircle className="w-5 h-5" />
                  شهادات الضمان
                </h3>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full">
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">رقم الشهادة</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">الوصف</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">تاريخ البداية</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">تاريخ النهاية</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">الحالة</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-200">
                    {operation.warrantyCertificates.map((warranty) => (
                      <tr key={warranty.id}>
                        <td className="px-6 py-4 text-sm font-medium text-gray-900">{warranty.certificateNumber}</td>
                        <td className="px-6 py-4 text-sm text-gray-900">{warranty.description}</td>
                        <td className="px-6 py-4 text-sm text-gray-900">{formatDate(warranty.startDate)}</td>
                        <td className="px-6 py-4 text-sm text-gray-900">{formatDate(warranty.endDate)}</td>
                        <td className="px-6 py-4">
                          <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${
                            warranty.isActive 
                              ? 'bg-green-100 text-green-800' 
                              : 'bg-red-100 text-red-800'
                          }`}>
                            {warranty.isActive ? 'نشط' : 'منتهي'}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Received Payments */}
          {operation.receivedPayments.length > 0 && (
            <div className="bg-white border border-gray-200 rounded-lg">
              <div className="p-6 border-b border-gray-200">
                <h3 className="text-xl font-semibold text-gray-900 flex items-center gap-2">
                  <DollarSign className="w-5 h-5" />
                  المدفوعات المستلمة
                </h3>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full">
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">النوع</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">المبلغ</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">التاريخ</th>
                      <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">التفاصيل</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-200">
                    {operation.receivedPayments.map((payment) => (
                      <tr key={payment.id}>
                        <td className="px-6 py-4 text-sm text-gray-900">
                          {payment.type === 'cash' ? 'نقدي' : 'شيك'}
                        </td>
                        <td className="px-6 py-4 text-sm font-medium text-gray-900">
                          {formatCurrency(payment.amount)}
                        </td>
                        <td className="px-6 py-4 text-sm text-gray-900">{formatDate(payment.date)}</td>
                        <td className="px-6 py-4 text-sm text-gray-900">
                          {payment.type === 'check' && payment.checkNumber 
                            ? `شيك رقم: ${payment.checkNumber} - ${payment.bank}` 
                            : payment.notes || '-'
                          }
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default OperationDetails;