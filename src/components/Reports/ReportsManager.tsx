import React, { useState } from 'react';
import { BarChart3, FileText, Download, Calendar, DollarSign, TrendingUp, PieChart, FileSpreadsheet, Award, FileEdit } from 'lucide-react';
import { Operation, Client } from '../../types';
import { formatCurrency, formatDate } from '../../utils/calculations';
import { 
  exportOperationsToPDF, 
  exportOperationsToExcel,
  exportOperationsToWord,
  exportClientsToPDF, 
  exportClientsToExcel,
  exportClientsToWord,
  exportFinancialReportToPDF,
  exportFinancialReportToExcel,
  exportFinancialReportToWord,
  exportDetailedGuaranteesReportToPDF,
  exportDetailedGuaranteesReportToWord,
  exportWarrantyCertificatesReportToPDF,
  exportWarrantyCertificatesReportToWord
} from '../../utils/exportUtils';

interface ReportsManagerProps {
  operations: Operation[];
  clients: Client[];
}

const ReportsManager: React.FC<ReportsManagerProps> = ({ operations, clients }) => {
  const [activeReport, setActiveReport] = useState<'summary' | 'client' | 'financial' | 'guarantees' | 'warranties' | 'detailed-guarantees'>('summary');
  const [selectedClient, setSelectedClient] = useState<string>('all');
  const [dateFrom, setDateFrom] = useState<string>('');
  const [dateTo, setDateTo] = useState<string>('');

  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  // Filter operations based on selected criteria
  const filteredOperations = operations.filter(operation => {
    const clientMatch = selectedClient === 'all' || operation.clientId === selectedClient;
    
    let dateMatch = true;
    if (dateFrom) {
      dateMatch = dateMatch && new Date(operation.createdAt) >= new Date(dateFrom);
    }
    if (dateTo) {
      dateMatch = dateMatch && new Date(operation.createdAt) <= new Date(dateTo);
    }
    
    return clientMatch && dateMatch;
  });

  // Calculate summary statistics
  const totalOperations = filteredOperations.length;
  const completedOperations = filteredOperations.filter(op => op.status === 'completed').length;
  const inProgressOperations = filteredOperations.filter(op => op.status === 'in_progress').length;
  const totalAmount = filteredOperations.reduce((sum, op) => sum + op.totalAmount, 0);
  const totalReceived = filteredOperations.reduce((sum, op) => sum + op.totalReceived, 0);
  const totalOutstanding = totalAmount - totalReceived;

  // Client-wise statistics
  const clientStats = clients.map(client => {
    const clientOperations = filteredOperations.filter(op => op.clientId === client.id);
    const clientTotal = clientOperations.reduce((sum, op) => sum + op.totalAmount, 0);
    const clientReceived = clientOperations.reduce((sum, op) => sum + op.totalReceived, 0);
    
    return {
      client,
      operationsCount: clientOperations.length,
      totalAmount: clientTotal,
      totalReceived: clientReceived,
      outstanding: clientTotal - clientReceived
    };
  }).filter(stat => stat.operationsCount > 0);

  // Guarantee statistics
  const totalGuaranteeChecks = filteredOperations.reduce((sum, op) => sum + op.guaranteeChecks.length, 0);
  const totalGuaranteeLetters = filteredOperations.reduce((sum, op) => sum + op.guaranteeLetters.length, 0);
  const activeGuaranteeChecks = filteredOperations.reduce((sum, op) => 
    sum + op.guaranteeChecks.filter(check => !check.isReturned).length, 0
  );
  const activeGuaranteeLetters = filteredOperations.reduce((sum, op) => 
    sum + op.guaranteeLetters.filter(letter => !letter.isReturned).length, 0
  );

  // Warranty statistics
  const totalWarranties = filteredOperations.reduce((sum, op) => 
    sum + (op.warrantyCertificates || []).length, 0
  );
  const activeWarranties = filteredOperations.reduce((sum, op) => 
    sum + (op.warrantyCertificates || []).filter(warranty => warranty.isActive).length, 0
  );

  // Export handlers
  const handleExportPDF = () => {
    try {
      switch (activeReport) {
        case 'summary':
          exportOperationsToPDF(filteredOperations, clients, 'تقرير العمليات الإجمالي');
          break;
        case 'client':
          exportClientsToPDF(clients, 'تقرير العملاء');
          break;
        case 'financial':
          exportFinancialReportToPDF(filteredOperations, clients);
          break;
        case 'guarantees':
          exportOperationsToPDF(filteredOperations, clients, 'تقرير الضمانات');
          break;
        case 'detailed-guarantees':
          exportDetailedGuaranteesReportToPDF(filteredOperations, clients);
          break;
        case 'warranties':
          exportWarrantyCertificatesReportToPDF(filteredOperations, clients);
          break;
        default:
          exportOperationsToPDF(filteredOperations, clients);
      }
    } catch (error) {
      console.error('خطأ في تصدير PDF:', error);
      alert('حدث خطأ أثناء تصدير ملف PDF');
    }
  };

  const handleExportExcel = () => {
    try {
      switch (activeReport) {
        case 'summary':
          exportOperationsToExcel(filteredOperations, clients, 'تقرير العمليات الإجمالي');
          break;
        case 'client':
          exportClientsToExcel(clients, 'تقرير العملاء');
          break;
        case 'financial':
          exportFinancialReportToExcel(filteredOperations, clients);
          break;
        case 'guarantees':
          exportOperationsToExcel(filteredOperations, clients, 'تقرير الضمانات');
          break;
        case 'detailed-guarantees':
          exportOperationsToExcel(filteredOperations, clients, 'تقرير الضمانات المفصل');
          break;
        case 'warranties':
          exportOperationsToExcel(filteredOperations, clients, 'تقرير شهادات الضمان');
          break;
        default:
          exportOperationsToExcel(filteredOperations, clients);
      }
    } catch (error) {
      console.error('خطأ في تصدير Excel:', error);
      alert('حدث خطأ أثناء تصدير ملف Excel');
    }
  };

  const handleExportWord = () => {
    try {
      switch (activeReport) {
        case 'summary':
          exportOperationsToWord(filteredOperations, clients, 'تقرير العمليات الإجمالي');
          break;
        case 'client':
          exportClientsToWord(clients, 'تقرير العملاء');
          break;
        case 'financial':
          exportFinancialReportToWord(filteredOperations, clients);
          break;
        case 'guarantees':
          exportOperationsToWord(filteredOperations, clients, 'تقرير الضمانات');
          break;
        case 'detailed-guarantees':
          exportDetailedGuaranteesReportToWord(filteredOperations, clients);
          break;
        case 'warranties':
          exportWarrantyCertificatesReportToWord(filteredOperations, clients);
          break;
        default:
          exportOperationsToWord(filteredOperations, clients);
      }
    } catch (error) {
      console.error('خطأ في تصدير Word:', error);
      alert('حدث خطأ أثناء تصدير ملف Word');
    }
  };

  const exportReport = () => {
    const reportData = {
      reportType: activeReport,
      generatedAt: new Date().toISOString(),
      filters: {
        client: selectedClient,
        dateFrom,
        dateTo
      },
      summary: {
        totalOperations,
        completedOperations,
        inProgressOperations,
        totalAmount,
        totalReceived,
        totalOutstanding,
        totalWarranties,
        activeWarranties
      },
      operations: filteredOperations,
      clientStats: activeReport === 'client' ? clientStats : undefined
    };

    const blob = new Blob([JSON.stringify(reportData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `construction-report-${activeReport}-${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const reports = [
    { id: 'summary', label: 'تقرير إجمالي', icon: BarChart3 },
    { id: 'client', label: 'تقرير العملاء', icon: FileText },
    { id: 'financial', label: 'التقرير المالي', icon: DollarSign },
    { id: 'guarantees', label: 'ملخص الضمانات', icon: PieChart },
    { id: 'detailed-guarantees', label: 'الضمانات المفصلة', icon: FileText },
    { id: 'warranties', label: 'شهادات الضمان', icon: Award },
  ];

  return (
    <div className="p-6">
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-gray-900 mb-2">التقارير والإحصائيات</h2>
        <p className="text-gray-600">تقارير شاملة ومفصلة عن أداء العمليات والمشاريع</p>
      </div>

      {/* Report Type Tabs */}
      <div className="mb-6 bg-white rounded-lg border border-gray-200 p-1">
        <nav className="flex flex-wrap gap-1">
          {reports.map((report) => {
            const Icon = report.icon;
            return (
              <button
                key={report.id}
                onClick={() => setActiveReport(report.id as any)}
                className={`flex items-center gap-2 px-4 py-3 rounded-md font-medium text-sm transition-all ${
                  activeReport === report.id
                    ? 'bg-blue-600 text-white shadow-md'
                    : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
                }`}
              >
                <Icon className="w-4 h-4" />
                {report.label}
              </button>
            );
          })}
        </nav>
      </div>

      {/* Filters and Export */}
      <div className="mb-6 bg-white p-4 rounded-lg border border-gray-200">
        <div className="grid grid-cols-1 md:grid-cols-7 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">العميل</label>
            <select
              value={selectedClient}
              onChange={(e) => setSelectedClient(e.target.value)}
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

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">من تاريخ</label>
            <input
              type="date"
              value={dateFrom}
              onChange={(e) => setDateFrom(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">إلى تاريخ</label>
            <input
              type="date"
              value={dateTo}
              onChange={(e) => setDateTo(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>

          <div className="flex items-end">
            <button
              onClick={handleExportPDF}
              className="flex items-center gap-2 px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700 transition-colors w-full justify-center"
            >
              <FileText className="w-4 h-4" />
              PDF
            </button>
          </div>

          <div className="flex items-end">
            <button
              onClick={handleExportExcel}
              className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 transition-colors w-full justify-center"
            >
              <FileSpreadsheet className="w-4 h-4" />
              Excel
            </button>
          </div>

          <div className="flex items-end">
            <button
              onClick={handleExportWord}
              className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors w-full justify-center"
            >
              <FileEdit className="w-4 h-4" />
              Word
            </button>
          </div>

          <div className="flex items-end">
            <button
              onClick={exportReport}
              className="flex items-center gap-2 px-4 py-2 bg-gray-600 text-white rounded-md hover:bg-gray-700 transition-colors w-full justify-center"
            >
              <Download className="w-4 h-4" />
              JSON
            </button>
          </div>
        </div>
      </div>

      {/* Summary Report */}
      {activeReport === 'summary' && (
        <div className="space-y-6">
          {/* Summary Cards */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            <div className="bg-blue-50 p-6 rounded-lg border border-blue-200">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-blue-600">إجمالي العمليات</p>
                  <p className="text-2xl font-bold text-blue-900">{totalOperations}</p>
                </div>
                <BarChart3 className="w-8 h-8 text-blue-600" />
              </div>
            </div>

            <div className="bg-green-50 p-6 rounded-lg border border-green-200">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-green-600">العمليات المكتملة</p>
                  <p className="text-2xl font-bold text-green-900">{completedOperations}</p>
                </div>
                <TrendingUp className="w-8 h-8 text-green-600" />
              </div>
            </div>

            <div className="bg-yellow-50 p-6 rounded-lg border border-yellow-200">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-yellow-600">العمليات الجارية</p>
                  <p className="text-2xl font-bold text-yellow-900">{inProgressOperations}</p>
                </div>
                <Calendar className="w-8 h-8 text-yellow-600" />
              </div>
            </div>

            <div className="bg-purple-50 p-6 rounded-lg border border-purple-200">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-purple-600">إجمالي القيمة</p>
                  <p className="text-2xl font-bold text-purple-900">{formatCurrency(totalAmount)}</p>
                </div>
                <DollarSign className="w-8 h-8 text-purple-600" />
              </div>
            </div>

            <div className="bg-green-50 p-6 rounded-lg border border-green-200">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-green-600">المبلغ المحصل</p>
                  <p className="text-2xl font-bold text-green-900">{formatCurrency(totalReceived)}</p>
                </div>
                <TrendingUp className="w-8 h-8 text-green-600" />
              </div>
            </div>

            <div className="bg-red-50 p-6 rounded-lg border border-red-200">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-red-600">المبلغ المتبقي</p>
                  <p className="text-2xl font-bold text-red-900">{formatCurrency(totalOutstanding)}</p>
                </div>
                <DollarSign className="w-8 h-8 text-red-600" />
              </div>
            </div>
          </div>

          {/* Operations Table */}
          <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
            <div className="p-6 border-b border-gray-200">
              <h3 className="text-lg font-semibold text-gray-900">تفاصيل العمليات</h3>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead className="bg-gray-50">
                  <tr>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">العملية</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">العميل</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">القيمة</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">المحصل</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">المتبقي</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">الحالة</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-200">
                  {filteredOperations.map((operation) => (
                    <tr key={operation.id}>
                      <td className="px-6 py-4 text-sm font-medium text-gray-900">{operation.name}</td>
                      <td className="px-6 py-4 text-sm text-gray-900">{getClientName(operation.clientId)}</td>
                      <td className="px-6 py-4 text-sm text-gray-900">{formatCurrency(operation.totalAmount)}</td>
                      <td className="px-6 py-4 text-sm text-gray-900">{formatCurrency(operation.totalReceived)}</td>
                      <td className="px-6 py-4 text-sm text-gray-900">
                        {formatCurrency(operation.totalAmount - operation.totalReceived)}
                      </td>
                      <td className="px-6 py-4">
                        <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${
                          operation.status === 'completed' 
                            ? 'bg-green-100 text-green-800'
                            : operation.status === 'completed_partial_payment'
                            ? 'bg-orange-100 text-orange-800'
                            : 'bg-yellow-100 text-yellow-800'
                        }`}>
                          {operation.status === 'completed' ? 'مكتملة' :
                           operation.status === 'completed_partial_payment' ? 'مكتملة - دفع جزئي' :
                           'قيد التنفيذ'}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* Client Report */}
      {activeReport === 'client' && (
        <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
          <div className="p-6 border-b border-gray-200">
            <h3 className="text-lg font-semibold text-gray-900">تقرير العملاء</h3>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full">
              <thead className="bg-gray-50">
                <tr>
                  <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">العميل</th>
                  <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">النوع</th>
                  <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">عدد العمليات</th>
                  <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">إجمالي القيمة</th>
                  <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">المبلغ المحصل</th>
                  <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">المبلغ المتبقي</th>
                  <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">معدل التحصيل</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-200">
                {clientStats.map((stat) => (
                  <tr key={stat.client.id}>
                    <td className="px-6 py-4 text-sm font-medium text-gray-900">{stat.client.name}</td>
                    <td className="px-6 py-4 text-sm text-gray-900">
                      {stat.client.type === 'owner' ? 'مالك' : 
                       stat.client.type === 'main_contractor' ? 'مقاول رئيسي' : 'استشاري'}
                    </td>
                    <td className="px-6 py-4 text-sm text-gray-900">{stat.operationsCount}</td>
                    <td className="px-6 py-4 text-sm text-gray-900">{formatCurrency(stat.totalAmount)}</td>
                    <td className="px-6 py-4 text-sm text-gray-900">{formatCurrency(stat.totalReceived)}</td>
                    <td className="px-6 py-4 text-sm text-gray-900">{formatCurrency(stat.outstanding)}</td>
                    <td className="px-6 py-4 text-sm text-gray-900">
                      {stat.totalAmount > 0 ? ((stat.totalReceived / stat.totalAmount) * 100).toFixed(1) : 0}%
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Financial Report */}
      {activeReport === 'financial' && (
        <div className="space-y-6">
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <div className="bg-white p-6 rounded-lg border border-gray-200">
              <h3 className="text-lg font-semibold text-gray-900 mb-4">الملخص المالي</h3>
              <div className="space-y-4">
                <div className="flex justify-between items-center">
                  <span className="text-gray-600">إجمالي القيمة:</span>
                  <span className="font-bold text-gray-900">{formatCurrency(totalAmount)}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-gray-600">المبلغ المحصل:</span>
                  <span className="font-bold text-green-600">{formatCurrency(totalReceived)}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-gray-600">المبلغ المتبقي:</span>
                  <span className="font-bold text-red-600">{formatCurrency(totalOutstanding)}</span>
                </div>
                <div className="flex justify-between items-center pt-4 border-t">
                  <span className="text-gray-600">معدل التحصيل:</span>
                  <span className="font-bold text-blue-600">
                    {totalAmount > 0 ? ((totalReceived / totalAmount) * 100).toFixed(1) : 0}%
                  </span>
                </div>
              </div>
            </div>

            <div className="bg-white p-6 rounded-lg border border-gray-200">
              <h3 className="text-lg font-semibold text-gray-900 mb-4">توزيع الحالات</h3>
              <div className="space-y-4">
                <div className="flex justify-between items-center">
                  <span className="text-gray-600">مكتملة:</span>
                  <span className="font-bold text-green-600">{completedOperations}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-gray-600">قيد التنفيذ:</span>
                  <span className="font-bold text-yellow-600">{inProgressOperations}</span>
                </div>
                <div className="flex justify-between items-center pt-4 border-t">
                  <span className="text-gray-600">معدل الإنجاز:</span>
                  <span className="font-bold text-blue-600">
                    {totalOperations > 0 ? ((completedOperations / totalOperations) * 100).toFixed(1) : 0}%
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* Detailed Financial Table */}
          <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
            <div className="p-6 border-b border-gray-200">
              <h3 className="text-lg font-semibold text-gray-900">التفاصيل المالية حسب العميل</h3>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead className="bg-gray-50">
                  <tr>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">العميل</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">عدد العمليات</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">إجمالي القيمة</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">المبلغ المحصل</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">المبلغ المتبقي</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">معدل التحصيل</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase">نسبة من الإجمالي</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-200">
                  {clientStats.map((stat) => (
                    <tr key={stat.client.id}>
                      <td className="px-6 py-4 text-sm font-medium text-gray-900">{stat.client.name}</td>
                      <td className="px-6 py-4 text-sm text-gray-900">{stat.operationsCount}</td>
                      <td className="px-6 py-4 text-sm text-gray-900">{formatCurrency(stat.totalAmount)}</td>
                      <td className="px-6 py-4 text-sm text-green-600 font-medium">{formatCurrency(stat.totalReceived)}</td>
                      <td className="px-6 py-4 text-sm text-red-600 font-medium">{formatCurrency(stat.outstanding)}</td>
                      <td className="px-6 py-4 text-sm text-gray-900">
                        <div className="flex items-center">
                          <div className="w-16 bg-gray-200 rounded-full h-2 mr-2">
                            <div
                              className="bg-blue-600 h-2 rounded-full"
                              style={{ width: `${stat.totalAmount > 0 ? (stat.totalReceived / stat.totalAmount) * 100 : 0}%` }}
                            ></div>
                          </div>
                          <span className="text-xs font-medium">
                            {stat.totalAmount > 0 ? ((stat.totalReceived / stat.totalAmount) * 100).toFixed(1) : 0}%
                          </span>
                        </div>
                      </td>
                      <td className="px-6 py-4 text-sm text-gray-900">
                        {totalAmount > 0 ? ((stat.totalAmount / totalAmount) * 100).toFixed(1) : 0}%
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* Guarantees Summary Report */}
      {activeReport === 'guarantees' && (
        <div className="space-y-6">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
            <div className="bg-blue-50 p-6 rounded-lg border border-blue-200">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-blue-600">شيكات الضمان</p>
                  <p className="text-2xl font-bold text-blue-900">{totalGuaranteeChecks}</p>
                </div>
                <FileText className="w-8 h-8 text-blue-600" />
              </div>
            </div>

            <div className="bg-purple-50 p-6 rounded-lg border border-purple-200">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-purple-600">خطابات الضمان</p>
                  <p className="text-2xl font-bold text-purple-900">{totalGuaranteeLetters}</p>
                </div>
                <FileText className="w-8 h-8 text-purple-600" />
              </div>
            </div>

            <div className="bg-yellow-50 p-6 rounded-lg border border-yellow-200">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-yellow-600">ضمانات قائمة</p>
                  <p className="text-2xl font-bold text-yellow-900">{activeGuaranteeChecks + activeGuaranteeLetters}</p>
                </div>
                <Calendar className="w-8 h-8 text-yellow-600" />
              </div>
            </div>

            <div className="bg-green-50 p-6 rounded-lg border border-green-200">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-green-600">شهادات الضمان</p>
                  <p className="text-2xl font-bold text-green-900">{totalWarranties}</p>
                </div>
                <Award className="w-8 h-8 text-green-600" />
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Detailed Guarantees Report */}
      {activeReport === 'detailed-guarantees' && (
        <div className="space-y-6">
          <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
            <div className="p-6 border-b border-gray-200">
              <h3 className="text-lg font-semibold text-gray-900">تقرير الضمانات المفصل</h3>
              <p className="text-sm text-gray-600 mt-1">عرض تفصيلي لجميع شيكات وخطابات الضمان</p>
            </div>
            <div className="p-6">
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 mb-6">
                <div className="bg-blue-50 p-4 rounded-lg">
                  <div className="text-center">
                    <p className="text-2xl font-bold text-blue-900">{totalGuaranteeChecks}</p>
                    <p className="text-sm text-blue-600">شيكات الضمان</p>
                  </div>
                </div>
                <div className="bg-purple-50 p-4 rounded-lg">
                  <div className="text-center">
                    <p className="text-2xl font-bold text-purple-900">{totalGuaranteeLetters}</p>
                    <p className="text-sm text-purple-600">خطابات الضمان</p>
                  </div>
                </div>
                <div className="bg-green-50 p-4 rounded-lg">
                  <div className="text-center">
                    <p className="text-2xl font-bold text-green-900">{activeGuaranteeChecks + activeGuaranteeLetters}</p>
                    <p className="text-sm text-green-600">ضمانات نشطة</p>
                  </div>
                </div>
                <div className="bg-red-50 p-4 rounded-lg">
                  <div className="text-center">
                    <p className="text-2xl font-bold text-red-900">{(totalGuaranteeChecks + totalGuaranteeLetters) - (activeGuaranteeChecks + activeGuaranteeLetters)}</p>
                    <p className="text-sm text-red-600">ضمانات مستردة</p>
                  </div>
                </div>
              </div>
              <p className="text-gray-600">
                يحتوي هذا التقرير على تفاصيل شاملة لجميع الضمانات مع معلومات العميل والعملية والبند المرتبط.
                استخدم أزرار التصدير أعلاه للحصول على التقرير المفصل.
              </p>
            </div>
          </div>
        </div>
      )}

      {/* Warranties Report */}
      {activeReport === 'warranties' && (
        <div className="space-y-6">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            <div className="bg-green-50 p-6 rounded-lg border border-green-200">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-green-600">إجمالي الشهادات</p>
                  <p className="text-2xl font-bold text-green-900">{totalWarranties}</p>
                </div>
                <Award className="w-8 h-8 text-green-600" />
              </div>
            </div>

            <div className="bg-blue-50 p-6 rounded-lg border border-blue-200">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-blue-600">الشهادات النشطة</p>
                  <p className="text-2xl font-bold text-blue-900">{activeWarranties}</p>
                </div>
                <TrendingUp className="w-8 h-8 text-blue-600" />
              </div>
            </div>

            <div className="bg-red-50 p-6 rounded-lg border border-red-200">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium text-red-600">الشهادات المنتهية</p>
                  <p className="text-2xl font-bold text-red-900">{totalWarranties - activeWarranties}</p>
                </div>
                <Calendar className="w-8 h-8 text-red-600" />
              </div>
            </div>
          </div>

          <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
            <div className="p-6 border-b border-gray-200">
              <h3 className="text-lg font-semibold text-gray-900">تفاصيل شهادات الضمان</h3>
              <p className="text-sm text-gray-600 mt-1">عرض تفصيلي لجميع شهادات الضمان الصادرة</p>
            </div>
            <div className="p-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
                <div className="bg-gray-50 p-4 rounded-lg">
                  <h4 className="font-medium text-gray-900 mb-2">توزيع الشهادات حسب الحالة</h4>
                  <div className="space-y-2">
                    <div className="flex justify-between">
                      <span className="text-sm text-gray-600">نشطة:</span>
                      <span className="text-sm font-medium text-green-600">{activeWarranties}</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-sm text-gray-600">منتهية:</span>
                      <span className="text-sm font-medium text-red-600">{totalWarranties - activeWarranties}</span>
                    </div>
                  </div>
                </div>
                <div className="bg-gray-50 p-4 rounded-lg">
                  <h4 className="font-medium text-gray-900 mb-2">معدل الشهادات النشطة</h4>
                  <div className="flex items-center">
                    <div className="w-full bg-gray-200 rounded-full h-3 mr-3">
                      <div
                        className="bg-green-600 h-3 rounded-full"
                        style={{ width: `${totalWarranties > 0 ? (activeWarranties / totalWarranties) * 100 : 0}%` }}
                      ></div>
                    </div>
                    <span className="text-sm font-medium text-gray-900">
                      {totalWarranties > 0 ? ((activeWarranties / totalWarranties) * 100).toFixed(1) : 0}%
                    </span>
                  </div>
                </div>
              </div>
              <p className="text-gray-600">
                يحتوي هذا التقرير على تفاصيل شاملة لجميع شهادات الضمان مع تواريخ البداية والنهاية وحالة كل شهادة.
                استخدم أزرار التصدير أعلاه للحصول على التقرير المفصل.
              </p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default ReportsManager;