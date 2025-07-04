import jsPDF from 'jspdf';
import 'jspdf-autotable';
import * as XLSX from 'xlsx';
import { Operation, Client } from '../types';
import { formatCurrency, formatDate, calculateNetAmount, calculateTotalDeductions, calculateExecutedTotal } from './calculations';

// حل بديل لمشكلة PDF - استخدام HTML2Canvas أو تحويل إلى صور
const createSimplePDF = (title: string, data: any[], headers: string[]) => {
  const doc = new jsPDF();
  
  // استخدام خط افتراضي يدعم Unicode
  doc.setFont('helvetica');
  
  // إضافة العنوان بالإنجليزية
  doc.setFontSize(16);
  doc.text(title, 20, 20);
  
  // إضافة التاريخ
  doc.setFontSize(10);
  doc.text(`Date: ${new Date().toLocaleDateString()}`, 20, 30);
  
  // إنشاء الجدول
  const tableData = data.map(row => 
    headers.map(header => {
      const value = row[header];
      if (typeof value === 'string' && /[\u0600-\u06FF]/.test(value)) {
        // إذا كان النص عربي، استبدله بـ placeholder
        return '[Arabic Text]';
      }
      return value || '';
    })
  );

  (doc as any).autoTable({
    head: [headers.map(h => h.replace(/[\u0600-\u06FF]/g, '[AR]'))],
    body: tableData,
    startY: 40,
    styles: {
      fontSize: 8,
      cellPadding: 2,
    },
    headStyles: {
      fillColor: [66, 139, 202],
      textColor: 255,
    },
  });

  return doc;
};

// دالة بديلة لتصدير PDF بتنسيق HTML
const exportAsHTMLToPDF = (title: string, content: string) => {
  const htmlContent = `
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
      <meta charset="UTF-8">
      <title>${title}</title>
      <style>
        body {
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          direction: rtl;
          text-align: right;
          margin: 20px;
          line-height: 1.6;
        }
        .header {
          text-align: center;
          border-bottom: 2px solid #333;
          padding-bottom: 10px;
          margin-bottom: 20px;
        }
        .title {
          font-size: 24px;
          font-weight: bold;
          color: #333;
          margin-bottom: 10px;
        }
        .date {
          font-size: 14px;
          color: #666;
        }
        table {
          width: 100%;
          border-collapse: collapse;
          margin: 20px 0;
          font-size: 12px;
        }
        th, td {
          border: 1px solid #ddd;
          padding: 8px;
          text-align: right;
        }
        th {
          background-color: #f2f2f2;
          font-weight: bold;
        }
        .summary {
          background-color: #f9f9f9;
          padding: 15px;
          border-radius: 5px;
          margin: 20px 0;
        }
        .summary h3 {
          margin-top: 0;
          color: #333;
        }
        .print-button {
          background-color: #007bff;
          color: white;
          padding: 10px 20px;
          border: none;
          border-radius: 5px;
          cursor: pointer;
          margin: 10px;
        }
        @media print {
          .print-button { display: none; }
        }
      </style>
    </head>
    <body>
      <div class="header">
        <div class="title">${title}</div>
        <div class="date">تاريخ التقرير: ${new Date().toLocaleDateString('ar-EG')}</div>
      </div>
      <button class="print-button" onclick="window.print()">طباعة التقرير</button>
      ${content}
    </body>
    </html>
  `;

  const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${title.replace(/\s+/g, '_')}_${new Date().toISOString().split('T')[0]}.html`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);

  // فتح في نافذة جديدة للطباعة
  const printWindow = window.open('', '_blank');
  if (printWindow) {
    printWindow.document.write(htmlContent);
    printWindow.document.close();
  }
};

export const exportOperationsToPDF = (operations: Operation[], clients: Client[], title = 'تقرير العمليات') => {
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
    return statusLabels[status] || status;
  };

  let tableContent = `
    <table>
      <thead>
        <tr>
          <th>كود العملية</th>
          <th>اسم العملية</th>
          <th>العميل</th>
          <th>القيمة الإجمالية</th>
          <th>الخصومات</th>
          <th>الصافي المستحق</th>
          <th>المبلغ المحصل</th>
          <th>المتبقي</th>
          <th>نسبة الإنجاز</th>
          <th>الحالة</th>
          <th>تاريخ الإنشاء</th>
        </tr>
      </thead>
      <tbody>
  `;

  let totalAmount = 0;
  let totalReceived = 0;
  let totalDeductions = 0;

  operations.forEach(operation => {
    const deductions = calculateTotalDeductions(operation);
    const netAmount = calculateNetAmount(operation);
    const remainingAmount = netAmount - operation.totalReceived;
    
    totalAmount += operation.totalAmount;
    totalReceived += operation.totalReceived;
    totalDeductions += deductions;

    tableContent += `
      <tr>
        <td>${operation.code}</td>
        <td>${operation.name}</td>
        <td>${getClientName(operation.clientId)}</td>
        <td>${formatCurrency(operation.totalAmount)}</td>
        <td>${formatCurrency(deductions)}</td>
        <td>${formatCurrency(netAmount)}</td>
        <td>${formatCurrency(operation.totalReceived)}</td>
        <td>${formatCurrency(remainingAmount)}</td>
        <td>${operation.overallExecutionPercentage.toFixed(1)}%</td>
        <td>${getStatusLabel(operation.status)}</td>
        <td>${formatDate(operation.createdAt)}</td>
      </tr>
    `;
  });

  tableContent += `
      </tbody>
    </table>
  `;

  const summaryContent = `
    <div class="summary">
      <h3>ملخص التقرير</h3>
      <p><strong>عدد العمليات:</strong> ${operations.length}</p>
      <p><strong>إجمالي القيمة:</strong> ${formatCurrency(totalAmount)}</p>
      <p><strong>إجمالي الخصومات:</strong> ${formatCurrency(totalDeductions)}</p>
      <p><strong>إجمالي المحصل:</strong> ${formatCurrency(totalReceived)}</p>
      <p><strong>إجمالي المتبقي:</strong> ${formatCurrency(totalAmount - totalReceived)}</p>
    </div>
  `;

  const content = summaryContent + tableContent;
  exportAsHTMLToPDF(title, content);
};

export const exportOperationDetailsToPDF = (operation: Operation, client: Client) => {
  const totalDeductions = calculateTotalDeductions(operation);
  const netAmount = calculateNetAmount(operation);
  const executedAmount = calculateExecutedTotal(operation.items);

  const content = `
    <div class="summary">
      <h3>معلومات العملية</h3>
      <p><strong>كود العملية:</strong> ${operation.code}</p>
      <p><strong>اسم العملية:</strong> ${operation.name}</p>
      <p><strong>العميل:</strong> ${client.name}</p>
      <p><strong>تاريخ الإنشاء:</strong> ${formatDate(operation.createdAt)}</p>
      <p><strong>الحالة:</strong> ${operation.status === 'completed' ? 'مكتملة' : operation.status === 'completed_partial_payment' ? 'مكتملة - دفع جزئي' : operation.status === 'completed_full_payment' ? 'مكتملة ومدفوعة بالكامل' : 'قيد التنفيذ'}</p>
    </div>

    <div class="summary">
      <h3>الملخص المالي</h3>
      <p><strong>القيمة الإجمالية:</strong> ${formatCurrency(operation.totalAmount)}</p>
      <p><strong>المبلغ المنفذ:</strong> ${formatCurrency(executedAmount)}</p>
      <p><strong>إجمالي الخصومات:</strong> ${formatCurrency(totalDeductions)}</p>
      <p><strong>الصافي المستحق:</strong> ${formatCurrency(netAmount)}</p>
      <p><strong>المبلغ المحصل:</strong> ${formatCurrency(operation.totalReceived)}</p>
      <p><strong>المبلغ المتبقي:</strong> ${formatCurrency(netAmount - operation.totalReceived)}</p>
      <p><strong>نسبة التنفيذ:</strong> ${operation.overallExecutionPercentage.toFixed(1)}%</p>
    </div>

    <h3>بنود العملية</h3>
    <table>
      <thead>
        <tr>
          <th>الكود</th>
          <th>الوصف</th>
          <th>القيمة</th>
          <th>نسبة التنفيذ</th>
          <th>القيمة المنفذة</th>
        </tr>
      </thead>
      <tbody>
        ${operation.items.map(item => `
          <tr>
            <td>${item.code}</td>
            <td>${item.description}</td>
            <td>${formatCurrency(item.amount)}</td>
            <td>${item.executionPercentage}%</td>
            <td>${formatCurrency(item.amount * (item.executionPercentage / 100))}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>

    ${operation.deductions.length > 0 ? `
      <h3>الخصومات</h3>
      <table>
        <thead>
          <tr>
            <th>اسم الخصم</th>
            <th>النوع</th>
            <th>القيمة</th>
            <th>المبلغ المخصوم</th>
          </tr>
        </thead>
        <tbody>
          ${operation.deductions.filter(d => d.isActive).map(deduction => {
            const deductionAmount = deduction.type === 'percentage' 
              ? (executedAmount * deduction.value / 100)
              : deduction.value;
            return `
              <tr>
                <td>${deduction.name}</td>
                <td>${deduction.type === 'percentage' ? 'نسبة مئوية' : 'مبلغ ثابت'}</td>
                <td>${deduction.type === 'percentage' ? deduction.value + '%' : formatCurrency(deduction.value)}</td>
                <td>${formatCurrency(deductionAmount)}</td>
              </tr>
            `;
          }).join('')}
        </tbody>
      </table>
    ` : ''}

    ${operation.receivedPayments.length > 0 ? `
      <h3>المدفوعات المستلمة</h3>
      <table>
        <thead>
          <tr>
            <th>النوع</th>
            <th>المبلغ</th>
            <th>التاريخ</th>
            <th>التفاصيل</th>
          </tr>
        </thead>
        <tbody>
          ${operation.receivedPayments.map(payment => `
            <tr>
              <td>${payment.type === 'cash' ? 'نقدي' : 'شيك'}</td>
              <td>${formatCurrency(payment.amount)}</td>
              <td>${formatDate(payment.date)}</td>
              <td>${payment.type === 'check' && payment.checkNumber ? `شيك رقم: ${payment.checkNumber}` : payment.notes || '-'}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    ` : ''}
  `;

  exportAsHTMLToPDF(`تفاصيل العملية - ${operation.name}`, content);
};

// باقي دوال التصدير بنفس الطريقة...
export const exportOperationsToExcel = (operations: Operation[], clients: Client[], title = 'تقرير العمليات') => {
  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  const data = operations.map(operation => {
    const deductions = calculateTotalDeductions(operation);
    const netAmount = calculateNetAmount(operation);
    
    return {
      'كود العملية': operation.code,
      'اسم العملية': operation.name,
      'العميل': getClientName(operation.clientId),
      'القيمة الإجمالية': operation.totalAmount,
      'الخصومات': deductions,
      'الصافي المستحق': netAmount,
      'المبلغ المحصل': operation.totalReceived,
      'المتبقي': netAmount - operation.totalReceived,
      'نسبة الإنجاز': operation.overallExecutionPercentage,
      'الحالة': operation.status === 'completed' ? 'مكتملة' : operation.status === 'completed_partial_payment' ? 'مكتملة - دفع جزئي' : operation.status === 'completed_full_payment' ? 'مكتملة ومدفوعة بالكامل' : 'قيد التنفيذ',
      'تاريخ الإنشاء': formatDate(operation.createdAt)
    };
  });

  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'العمليات');
  XLSX.writeFile(wb, `${title}_${new Date().toISOString().split('T')[0]}.xlsx`);
};

// دوال Word
export const createWordDocument = (title: string, content: string) => {
  const htmlContent = `
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
      <meta charset="UTF-8">
      <title>${title}</title>
      <style>
        body {
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          direction: rtl;
          text-align: right;
          margin: 40px;
          line-height: 1.8;
          color: #333;
        }
        .header {
          text-align: center;
          border-bottom: 3px solid #2c5aa0;
          padding-bottom: 20px;
          margin-bottom: 30px;
        }
        .title {
          font-size: 28px;
          font-weight: bold;
          color: #2c5aa0;
          margin-bottom: 10px;
        }
        .subtitle {
          font-size: 16px;
          color: #666;
          margin-bottom: 5px;
        }
        .date {
          font-size: 14px;
          color: #888;
        }
        table {
          width: 100%;
          border-collapse: collapse;
          margin: 25px 0;
          font-size: 13px;
          box-shadow: 0 0 20px rgba(0,0,0,0.1);
        }
        th, td {
          border: 1px solid #ddd;
          padding: 12px;
          text-align: right;
        }
        th {
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
          font-weight: bold;
          text-align: center;
        }
        tr:nth-child(even) {
          background-color: #f8f9fa;
        }
        tr:hover {
          background-color: #e3f2fd;
        }
        .summary {
          background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
          padding: 20px;
          border-radius: 10px;
          margin: 25px 0;
          border-right: 5px solid #2c5aa0;
        }
        .summary h3 {
          margin-top: 0;
          color: #2c5aa0;
          font-size: 20px;
        }
        .summary p {
          margin: 8px 0;
          font-size: 14px;
        }
        .highlight {
          background-color: #fff3cd;
          padding: 2px 6px;
          border-radius: 3px;
          font-weight: bold;
        }
        .section-title {
          color: #2c5aa0;
          font-size: 22px;
          margin: 30px 0 15px 0;
          padding-bottom: 10px;
          border-bottom: 2px solid #e9ecef;
        }
      </style>
    </head>
    <body>
      <div class="header">
        <div class="title">${title}</div>
        <div class="subtitle">نظام إدارة العمليات الإنشائية</div>
        <div class="date">تاريخ التقرير: ${new Date().toLocaleDateString('ar-EG')}</div>
      </div>
      ${content}
      <div style="margin-top: 50px; text-align: center; color: #888; font-size: 12px;">
        تم إنشاء هذا التقرير بواسطة نظام إدارة العمليات الإنشائية
      </div>
    </body>
    </html>
  `;

  const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${title.replace(/\s+/g, '_')}_${new Date().toISOString().split('T')[0]}.html`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

export const exportOperationsToWord = (operations: Operation[], clients: Client[], title = 'تقرير العمليات') => {
  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  let tableContent = `
    <h2 class="section-title">تفاصيل العمليات</h2>
    <table>
      <thead>
        <tr>
          <th>كود العملية</th>
          <th>اسم العملية</th>
          <th>العميل</th>
          <th>القيمة الإجمالية</th>
          <th>الخصومات</th>
          <th>الصافي المستحق</th>
          <th>المبلغ المحصل</th>
          <th>المتبقي</th>
          <th>نسبة الإنجاز</th>
          <th>الحالة</th>
        </tr>
      </thead>
      <tbody>
  `;

  let totalAmount = 0;
  let totalReceived = 0;
  let totalDeductions = 0;

  operations.forEach(operation => {
    const deductions = calculateTotalDeductions(operation);
    const netAmount = calculateNetAmount(operation);
    const remainingAmount = netAmount - operation.totalReceived;
    
    totalAmount += operation.totalAmount;
    totalReceived += operation.totalReceived;
    totalDeductions += deductions;

    const statusLabel = operation.status === 'completed' ? 'مكتملة' : 
                       operation.status === 'completed_partial_payment' ? 'مكتملة - دفع جزئي' : 
                       operation.status === 'completed_full_payment' ? 'مكتملة ومدفوعة بالكامل' : 'قيد التنفيذ';

    tableContent += `
      <tr>
        <td>${operation.code}</td>
        <td>${operation.name}</td>
        <td>${getClientName(operation.clientId)}</td>
        <td>${formatCurrency(operation.totalAmount)}</td>
        <td>${formatCurrency(deductions)}</td>
        <td>${formatCurrency(netAmount)}</td>
        <td>${formatCurrency(operation.totalReceived)}</td>
        <td>${formatCurrency(remainingAmount)}</td>
        <td>${operation.overallExecutionPercentage.toFixed(1)}%</td>
        <td>${statusLabel}</td>
      </tr>
    `;
  });

  tableContent += `
      </tbody>
    </table>
  `;

  const summaryContent = `
    <div class="summary">
      <h3>ملخص التقرير</h3>
      <p><strong>عدد العمليات:</strong> <span class="highlight">${operations.length}</span></p>
      <p><strong>إجمالي القيمة:</strong> <span class="highlight">${formatCurrency(totalAmount)}</span></p>
      <p><strong>إجمالي الخصومات:</strong> <span class="highlight">${formatCurrency(totalDeductions)}</span></p>
      <p><strong>إجمالي المحصل:</strong> <span class="highlight">${formatCurrency(totalReceived)}</span></p>
      <p><strong>إجمالي المتبقي:</strong> <span class="highlight">${formatCurrency(totalAmount - totalReceived)}</span></p>
      <p><strong>معدل التحصيل:</strong> <span class="highlight">${totalAmount > 0 ? ((totalReceived / totalAmount) * 100).toFixed(1) : 0}%</span></p>
    </div>
  `;

  const content = summaryContent + tableContent;
  createWordDocument(title, content);
};

export const exportOperationDetailsToWord = (operation: Operation, client: Client) => {
  const totalDeductions = calculateTotalDeductions(operation);
  const netAmount = calculateNetAmount(operation);
  const executedAmount = calculateExecutedTotal(operation.items);

  const statusLabel = operation.status === 'completed' ? 'مكتملة' : 
                     operation.status === 'completed_partial_payment' ? 'مكتملة - دفع جزئي' : 
                     operation.status === 'completed_full_payment' ? 'مكتملة ومدفوعة بالكامل' : 'قيد التنفيذ';

  const content = `
    <div class="summary">
      <h3>معلومات العملية</h3>
      <p><strong>كود العملية:</strong> <span class="highlight">${operation.code}</span></p>
      <p><strong>اسم العملية:</strong> <span class="highlight">${operation.name}</span></p>
      <p><strong>العميل:</strong> <span class="highlight">${client.name}</span></p>
      <p><strong>تاريخ الإنشاء:</strong> <span class="highlight">${formatDate(operation.createdAt)}</span></p>
      <p><strong>الحالة:</strong> <span class="highlight">${statusLabel}</span></p>
    </div>

    <div class="summary">
      <h3>الملخص المالي</h3>
      <p><strong>القيمة الإجمالية:</strong> <span class="highlight">${formatCurrency(operation.totalAmount)}</span></p>
      <p><strong>المبلغ المنفذ:</strong> <span class="highlight">${formatCurrency(executedAmount)}</span></p>
      <p><strong>إجمالي الخصومات:</strong> <span class="highlight">${formatCurrency(totalDeductions)}</span></p>
      <p><strong>الصافي المستحق:</strong> <span class="highlight">${formatCurrency(netAmount)}</span></p>
      <p><strong>المبلغ المحصل:</strong> <span class="highlight">${formatCurrency(operation.totalReceived)}</span></p>
      <p><strong>المبلغ المتبقي:</strong> <span class="highlight">${formatCurrency(netAmount - operation.totalReceived)}</span></p>
      <p><strong>نسبة التنفيذ:</strong> <span class="highlight">${operation.overallExecutionPercentage.toFixed(1)}%</span></p>
    </div>

    <h2 class="section-title">بنود العملية</h2>
    <table>
      <thead>
        <tr>
          <th>الكود</th>
          <th>الوصف</th>
          <th>القيمة</th>
          <th>نسبة التنفيذ</th>
          <th>القيمة المنفذة</th>
        </tr>
      </thead>
      <tbody>
        ${operation.items.map(item => `
          <tr>
            <td>${item.code}</td>
            <td>${item.description}</td>
            <td>${formatCurrency(item.amount)}</td>
            <td>${item.executionPercentage}%</td>
            <td>${formatCurrency(item.amount * (item.executionPercentage / 100))}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>

    ${operation.deductions.length > 0 ? `
      <h2 class="section-title">الخصومات</h2>
      <table>
        <thead>
          <tr>
            <th>اسم الخصم</th>
            <th>النوع</th>
            <th>القيمة</th>
            <th>المبلغ المخصوم</th>
          </tr>
        </thead>
        <tbody>
          ${operation.deductions.filter(d => d.isActive).map(deduction => {
            const deductionAmount = deduction.type === 'percentage' 
              ? (executedAmount * deduction.value / 100)
              : deduction.value;
            return `
              <tr>
                <td>${deduction.name}</td>
                <td>${deduction.type === 'percentage' ? 'نسبة مئوية' : 'مبلغ ثابت'}</td>
                <td>${deduction.type === 'percentage' ? deduction.value + '%' : formatCurrency(deduction.value)}</td>
                <td>${formatCurrency(deductionAmount)}</td>
              </tr>
            `;
          }).join('')}
        </tbody>
      </table>
    ` : ''}

    ${operation.receivedPayments.length > 0 ? `
      <h2 class="section-title">المدفوعات المستلمة</h2>
      <table>
        <thead>
          <tr>
            <th>النوع</th>
            <th>المبلغ</th>
            <th>التاريخ</th>
            <th>التفاصيل</th>
          </tr>
        </thead>
        <tbody>
          ${operation.receivedPayments.map(payment => `
            <tr>
              <td>${payment.type === 'cash' ? 'نقدي' : 'شيك'}</td>
              <td>${formatCurrency(payment.amount)}</td>
              <td>${formatDate(payment.date)}</td>
              <td>${payment.type === 'check' && payment.checkNumber ? `شيك رقم: ${payment.checkNumber}` : payment.notes || '-'}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    ` : ''}
  `;

  createWordDocument(`تفاصيل العملية - ${operation.name}`, content);
};

// إضافة باقي دوال التصدير...
export const exportChecksAndPaymentsToPDF = (operations: Operation[], clients: Client[]) => {
  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  const allPayments = operations.flatMap(operation => 
    operation.receivedPayments.map(payment => ({
      ...payment,
      operationName: operation.name,
      operationCode: operation.code,
      clientName: getClientName(operation.clientId)
    }))
  );

  const content = `
    <div class="summary">
      <h3>ملخص المدفوعات</h3>
      <p><strong>إجمالي المدفوعات:</strong> ${allPayments.length}</p>
      <p><strong>الشيكات:</strong> ${allPayments.filter(p => p.type === 'check').length}</p>
      <p><strong>المدفوعات النقدية:</strong> ${allPayments.filter(p => p.type === 'cash').length}</p>
      <p><strong>إجمالي المبلغ:</strong> ${formatCurrency(allPayments.reduce((sum, p) => sum + p.amount, 0))}</p>
    </div>

    <table>
      <thead>
        <tr>
          <th>النوع</th>
          <th>المبلغ</th>
          <th>التاريخ</th>
          <th>العميل</th>
          <th>العملية</th>
          <th>التفاصيل</th>
        </tr>
      </thead>
      <tbody>
        ${allPayments.map(payment => `
          <tr>
            <td>${payment.type === 'check' ? 'شيك' : 'نقدي'}</td>
            <td>${formatCurrency(payment.amount)}</td>
            <td>${formatDate(payment.date)}</td>
            <td>${payment.clientName}</td>
            <td>${payment.operationName}</td>
            <td>${payment.type === 'check' && payment.checkNumber ? `شيك رقم: ${payment.checkNumber}` : payment.notes || '-'}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;

  exportAsHTMLToPDF('تقرير الشيكات والمدفوعات', content);
};

export const exportChecksAndPaymentsToExcel = (operations: Operation[], clients: Client[]) => {
  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  const allPayments = operations.flatMap(operation => 
    operation.receivedPayments.map(payment => ({
      'النوع': payment.type === 'check' ? 'شيك' : 'نقدي',
      'المبلغ': payment.amount,
      'التاريخ': formatDate(payment.date),
      'العميل': getClientName(operation.clientId),
      'العملية': operation.name,
      'كود العملية': operation.code,
      'رقم الشيك': payment.checkNumber || '',
      'البنك': payment.bank || '',
      'تاريخ الاستلام': payment.receiptDate ? formatDate(payment.receiptDate) : '',
      'ملاحظات': payment.notes || ''
    }))
  );

  const ws = XLSX.utils.json_to_sheet(allPayments);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'المدفوعات');
  XLSX.writeFile(wb, `تقرير_الشيكات_والمدفوعات_${new Date().toISOString().split('T')[0]}.xlsx`);
};

export const exportChecksAndPaymentsToWord = (operations: Operation[], clients: Client[]) => {
  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  const allPayments = operations.flatMap(operation => 
    operation.receivedPayments.map(payment => ({
      ...payment,
      operationName: operation.name,
      operationCode: operation.code,
      clientName: getClientName(operation.clientId)
    }))
  );

  const content = `
    <div class="summary">
      <h3>ملخص المدفوعات</h3>
      <p><strong>إجمالي المدفوعات:</strong> <span class="highlight">${allPayments.length}</span></p>
      <p><strong>الشيكات:</strong> <span class="highlight">${allPayments.filter(p => p.type === 'check').length}</span></p>
      <p><strong>المدفوعات النقدية:</strong> <span class="highlight">${allPayments.filter(p => p.type === 'cash').length}</span></p>
      <p><strong>إجمالي المبلغ:</strong> <span class="highlight">${formatCurrency(allPayments.reduce((sum, p) => sum + p.amount, 0))}</span></p>
    </div>

    <h2 class="section-title">تفاصيل المدفوعات</h2>
    <table>
      <thead>
        <tr>
          <th>النوع</th>
          <th>المبلغ</th>
          <th>التاريخ</th>
          <th>العميل</th>
          <th>العملية</th>
          <th>التفاصيل</th>
        </tr>
      </thead>
      <tbody>
        ${allPayments.map(payment => `
          <tr>
            <td>${payment.type === 'check' ? 'شيك' : 'نقدي'}</td>
            <td>${formatCurrency(payment.amount)}</td>
            <td>${formatDate(payment.date)}</td>
            <td>${payment.clientName}</td>
            <td>${payment.operationName}</td>
            <td>${payment.type === 'check' && payment.checkNumber ? `شيك رقم: ${payment.checkNumber} - ${payment.bank || ''}` : payment.notes || '-'}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;

  createWordDocument('تقرير الشيكات والمدفوعات', content);
};

// إضافة باقي دوال التصدير للضمانات وشهادات الضمان...
export const exportDetailedGuaranteesReportToPDF = (operations: Operation[], clients: Client[]) => {
  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  const allGuaranteeChecks = operations.flatMap(operation => 
    operation.guaranteeChecks.map(check => ({
      ...check,
      operationName: operation.name,
      clientName: getClientName(operation.clientId),
      type: 'شيك ضمان'
    }))
  );

  const allGuaranteeLetters = operations.flatMap(operation => 
    operation.guaranteeLetters.map(letter => ({
      ...letter,
      operationName: operation.name,
      clientName: getClientName(operation.clientId),
      type: 'خطاب ضمان'
    }))
  );

  const allGuarantees = [...allGuaranteeChecks, ...allGuaranteeLetters];

  const content = `
    <div class="summary">
      <h3>ملخص الضمانات</h3>
      <p><strong>إجمالي الضمانات:</strong> ${allGuarantees.length}</p>
      <p><strong>شيكات الضمان:</strong> ${allGuaranteeChecks.length}</p>
      <p><strong>خطابات الضمان:</strong> ${allGuaranteeLetters.length}</p>
      <p><strong>الضمانات النشطة:</strong> ${allGuarantees.filter(g => !g.isReturned).length}</p>
      <p><strong>الضمانات المستردة:</strong> ${allGuarantees.filter(g => g.isReturned).length}</p>
    </div>

    <table>
      <thead>
        <tr>
          <th>النوع</th>
          <th>الرقم</th>
          <th>المبلغ</th>
          <th>البنك</th>
          <th>العميل</th>
          <th>العملية</th>
          <th>تاريخ الانتهاء</th>
          <th>الحالة</th>
        </tr>
      </thead>
      <tbody>
        ${allGuarantees.map(guarantee => `
          <tr>
            <td>${guarantee.type}</td>
            <td>${guarantee.checkNumber || guarantee.letterNumber || ''}</td>
            <td>${formatCurrency(guarantee.amount)}</td>
            <td>${guarantee.bank}</td>
            <td>${guarantee.clientName}</td>
            <td>${guarantee.operationName}</td>
            <td>${formatDate(guarantee.expiryDate || guarantee.dueDate)}</td>
            <td>${guarantee.isReturned ? 'مُسترد' : 'قائم'}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;

  exportAsHTMLToPDF('تقرير الضمانات المفصل', content);
};

export const exportGuaranteesToExcel = (operations: Operation[], clients: Client[]) => {
  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  const allGuaranteeChecks = operations.flatMap(operation => 
    operation.guaranteeChecks.map(check => ({
      'النوع': 'شيك ضمان',
      'الرقم': check.checkNumber,
      'المبلغ': check.amount,
      'البنك': check.bank,
      'العميل': getClientName(operation.clientId),
      'العملية': operation.name,
      'تاريخ الشيك': formatDate(check.checkDate),
      'تاريخ التسليم': formatDate(check.deliveryDate),
      'تاريخ الانتهاء': formatDate(check.expiryDate),
      'الحالة': check.isReturned ? 'مُسترد' : 'قائم',
      'تاريخ الاسترداد': check.returnDate ? formatDate(check.returnDate) : ''
    }))
  );

  const allGuaranteeLetters = operations.flatMap(operation => 
    operation.guaranteeLetters.map(letter => ({
      'النوع': 'خطاب ضمان',
      'الرقم': letter.letterNumber,
      'المبلغ': letter.amount,
      'البنك': letter.bank,
      'العميل': getClientName(operation.clientId),
      'العملية': operation.name,
      'تاريخ الخطاب': formatDate(letter.letterDate),
      'تاريخ التسليم': '',
      'تاريخ الانتهاء': formatDate(letter.dueDate),
      'الحالة': letter.isReturned ? 'مُسترد' : 'قائم',
      'تاريخ الاسترداد': letter.returnDate ? formatDate(letter.returnDate) : ''
    }))
  );

  const allGuarantees = [...allGuaranteeChecks, ...allGuaranteeLetters];

  const ws = XLSX.utils.json_to_sheet(allGuarantees);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'الضمانات');
  XLSX.writeFile(wb, `تقرير_الضمانات_${new Date().toISOString().split('T')[0]}.xlsx`);
};

export const exportDetailedGuaranteesReportToWord = (operations: Operation[], clients: Client[]) => {
  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  const allGuaranteeChecks = operations.flatMap(operation => 
    operation.guaranteeChecks.map(check => ({
      ...check,
      operationName: operation.name,
      clientName: getClientName(operation.clientId),
      type: 'شيك ضمان'
    }))
  );

  const allGuaranteeLetters = operations.flatMap(operation => 
    operation.guaranteeLetters.map(letter => ({
      ...letter,
      operationName: operation.name,
      clientName: getClientName(operation.clientId),
      type: 'خطاب ضمان'
    }))
  );

  const allGuarantees = [...allGuaranteeChecks, ...allGuaranteeLetters];

  const content = `
    <div class="summary">
      <h3>ملخص الضمانات</h3>
      <p><strong>إجمالي الضمانات:</strong> <span class="highlight">${allGuarantees.length}</span></p>
      <p><strong>شيكات الضمان:</strong> <span class="highlight">${allGuaranteeChecks.length}</span></p>
      <p><strong>خطابات الضمان:</strong> <span class="highlight">${allGuaranteeLetters.length}</span></p>
      <p><strong>الضمانات النشطة:</strong> <span class="highlight">${allGuarantees.filter(g => !g.isReturned).length}</span></p>
      <p><strong>الضمانات المستردة:</strong> <span class="highlight">${allGuarantees.filter(g => g.isReturned).length}</span></p>
    </div>

    <h2 class="section-title">تفاصيل الضمانات</h2>
    <table>
      <thead>
        <tr>
          <th>النوع</th>
          <th>الرقم</th>
          <th>المبلغ</th>
          <th>البنك</th>
          <th>العميل</th>
          <th>العملية</th>
          <th>تاريخ الانتهاء</th>
          <th>الحالة</th>
        </tr>
      </thead>
      <tbody>
        ${allGuarantees.map(guarantee => `
          <tr>
            <td>${guarantee.type}</td>
            <td>${guarantee.checkNumber || guarantee.letterNumber || ''}</td>
            <td>${formatCurrency(guarantee.amount)}</td>
            <td>${guarantee.bank}</td>
            <td>${guarantee.clientName}</td>
            <td>${guarantee.operationName}</td>
            <td>${formatDate(guarantee.expiryDate || guarantee.dueDate)}</td>
            <td>${guarantee.isReturned ? 'مُسترد' : 'قائم'}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;

  createWordDocument('تقرير الضمانات المفصل', content);
};

export const exportWarrantyCertificatesReportToPDF = (operations: Operation[], clients: Client[]) => {
  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  const allWarranties = operations.flatMap(operation => 
    (operation.warrantyCertificates || []).map(warranty => ({
      ...warranty,
      operationName: operation.name,
      clientName: getClientName(operation.clientId)
    }))
  );

  const content = `
    <div class="summary">
      <h3>ملخص شهادات الضمان</h3>
      <p><strong>إجمالي الشهادات:</strong> ${allWarranties.length}</p>
      <p><strong>الشهادات النشطة:</strong> ${allWarranties.filter(w => w.isActive).length}</p>
      <p><strong>الشهادات المنتهية:</strong> ${allWarranties.filter(w => !w.isActive).length}</p>
    </div>

    <table>
      <thead>
        <tr>
          <th>رقم الشهادة</th>
          <th>العميل</th>
          <th>العملية</th>
          <th>الوصف</th>
          <th>تاريخ البداية</th>
          <th>تاريخ النهاية</th>
          <th>مدة الضمان</th>
          <th>الحالة</th>
        </tr>
      </thead>
      <tbody>
        ${allWarranties.map(warranty => `
          <tr>
            <td>${warranty.certificateNumber}</td>
            <td>${warranty.clientName}</td>
            <td>${warranty.operationName}</td>
            <td>${warranty.description}</td>
            <td>${formatDate(warranty.startDate)}</td>
            <td>${formatDate(warranty.endDate)}</td>
            <td>${warranty.warrantyPeriodMonths} شهر</td>
            <td>${warranty.isActive ? 'نشط' : 'منتهي'}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;

  exportAsHTMLToPDF('تقرير شهادات الضمان', content);
};

export const exportWarrantyCertificatesToExcel = (operations: Operation[], clients: Client[]) => {
  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  const allWarranties = operations.flatMap(operation => 
    (operation.warrantyCertificates || []).map(warranty => ({
      'رقم الشهادة': warranty.certificateNumber,
      'العميل': getClientName(operation.clientId),
      'العملية': operation.name,
      'الوصف': warranty.description,
      'تاريخ الإصدار': formatDate(warranty.issueDate),
      'تاريخ البداية': formatDate(warranty.startDate),
      'تاريخ النهاية': formatDate(warranty.endDate),
      'مدة الضمان (أشهر)': warranty.warrantyPeriodMonths,
      'مرتبط بـ': warranty.relatedTo === 'operation' ? 'العملية كاملة' : 'بند محدد',
      'الحالة': warranty.isActive ? 'نشط' : 'منتهي',
      'ملاحظات': warranty.notes || ''
    }))
  );

  const ws = XLSX.utils.json_to_sheet(allWarranties);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'شهادات الضمان');
  XLSX.writeFile(wb, `تقرير_شهادات_الضمان_${new Date().toISOString().split('T')[0]}.xlsx`);
};

export const exportWarrantyCertificatesReportToWord = (operations: Operation[], clients: Client[]) => {
  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  const allWarranties = operations.flatMap(operation => 
    (operation.warrantyCertificates || []).map(warranty => ({
      ...warranty,
      operationName: operation.name,
      clientName: getClientName(operation.clientId)
    }))
  );

  const content = `
    <div class="summary">
      <h3>ملخص شهادات الضمان</h3>
      <p><strong>إجمالي الشهادات:</strong> <span class="highlight">${allWarranties.length}</span></p>
      <p><strong>الشهادات النشطة:</strong> <span class="highlight">${allWarranties.filter(w => w.isActive).length}</span></p>
      <p><strong>الشهادات المنتهية:</strong> <span class="highlight">${allWarranties.filter(w => !w.isActive).length}</span></p>
    </div>

    <h2 class="section-title">تفاصيل شهادات الضمان</h2>
    <table>
      <thead>
        <tr>
          <th>رقم الشهادة</th>
          <th>العميل</th>
          <th>العملية</th>
          <th>الوصف</th>
          <th>تاريخ البداية</th>
          <th>تاريخ النهاية</th>
          <th>مدة الضمان</th>
          <th>الحالة</th>
        </tr>
      </thead>
      <tbody>
        ${allWarranties.map(warranty => `
          <tr>
            <td>${warranty.certificateNumber}</td>
            <td>${warranty.clientName}</td>
            <td>${warranty.operationName}</td>
            <td>${warranty.description}</td>
            <td>${formatDate(warranty.startDate)}</td>
            <td>${formatDate(warranty.endDate)}</td>
            <td>${warranty.warrantyPeriodMonths} شهر</td>
            <td>${warranty.isActive ? 'نشط' : 'منتهي'}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;

  createWordDocument('تقرير شهادات الضمان', content);
};

export const exportClientsToPDF = (clients: Client[], title = 'تقرير العملاء') => {
  const content = `
    <div class="summary">
      <h3>ملخص العملاء</h3>
      <p><strong>إجمالي العملاء:</strong> ${clients.length}</p>
      <p><strong>الملاك:</strong> ${clients.filter(c => c.type === 'owner').length}</p>
      <p><strong>المقاولون الرئيسيون:</strong> ${clients.filter(c => c.type === 'main_contractor').length}</p>
      <p><strong>الاستشاريون:</strong> ${clients.filter(c => c.type === 'consultant').length}</p>
    </div>

    <table>
      <thead>
        <tr>
          <th>اسم العميل</th>
          <th>النوع</th>
          <th>الهاتف</th>
          <th>البريد الإلكتروني</th>
          <th>العنوان</th>
          <th>جهات الاتصال</th>
          <th>تاريخ الإنشاء</th>
        </tr>
      </thead>
      <tbody>
        ${clients.map(client => `
          <tr>
            <td>${client.name}</td>
            <td>${client.type === 'owner' ? 'مالك' : client.type === 'main_contractor' ? 'مقاول رئيسي' : 'استشاري'}</td>
            <td>${client.phone || '-'}</td>
            <td>${client.email || '-'}</td>
            <td>${client.address || '-'}</td>
            <td>${client.contacts?.length || 0}</td>
            <td>${formatDate(client.createdAt)}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;

  exportAsHTMLToPDF(title, content);
};

export const exportClientsToExcel = (clients: Client[], title = 'تقرير العملاء') => {
  const data = clients.map(client => ({
    'اسم العميل': client.name,
    'النوع': client.type === 'owner' ? 'مالك' : client.type === 'main_contractor' ? 'مقاول رئيسي' : 'استشاري',
    'الهاتف': client.phone || '',
    'البريد الإلكتروني': client.email || '',
    'العنوان': client.address || '',
    'عدد جهات الاتصال': client.contacts?.length || 0,
    'تاريخ الإنشاء': formatDate(client.createdAt),
    'تاريخ التحديث': formatDate(client.updatedAt)
  }));

  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'العملاء');
  XLSX.writeFile(wb, `${title}_${new Date().toISOString().split('T')[0]}.xlsx`);
};

export const exportClientsToWord = (clients: Client[], title = 'تقرير العملاء') => {
  const content = `
    <div class="summary">
      <h3>ملخص العملاء</h3>
      <p><strong>إجمالي العملاء:</strong> <span class="highlight">${clients.length}</span></p>
      <p><strong>الملاك:</strong> <span class="highlight">${clients.filter(c => c.type === 'owner').length}</span></p>
      <p><strong>المقاولون الرئيسيون:</strong> <span class="highlight">${clients.filter(c => c.type === 'main_contractor').length}</span></p>
      <p><strong>الاستشاريون:</strong> <span class="highlight">${clients.filter(c => c.type === 'consultant').length}</span></p>
    </div>

    <h2 class="section-title">تفاصيل العملاء</h2>
    <table>
      <thead>
        <tr>
          <th>اسم العميل</th>
          <th>النوع</th>
          <th>الهاتف</th>
          <th>البريد الإلكتروني</th>
          <th>العنوان</th>
          <th>جهات الاتصال</th>
          <th>تاريخ الإنشاء</th>
        </tr>
      </thead>
      <tbody>
        ${clients.map(client => `
          <tr>
            <td>${client.name}</td>
            <td>${client.type === 'owner' ? 'مالك' : client.type === 'main_contractor' ? 'مقاول رئيسي' : 'استشاري'}</td>
            <td>${client.phone || '-'}</td>
            <td>${client.email || '-'}</td>
            <td>${client.address || '-'}</td>
            <td>${client.contacts?.length || 0}</td>
            <td>${formatDate(client.createdAt)}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;

  createWordDocument(title, content);
};

export const exportFinancialReportToPDF = (operations: Operation[], clients: Client[]) => {
  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  const totalAmount = operations.reduce((sum, op) => sum + op.totalAmount, 0);
  const totalReceived = operations.reduce((sum, op) => sum + op.totalReceived, 0);
  const totalDeductions = operations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
  const totalNet = operations.reduce((sum, op) => sum + calculateNetAmount(op), 0);

  const clientStats = clients.map(client => {
    const clientOperations = operations.filter(op => op.clientId === client.id);
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

  const content = `
    <div class="summary">
      <h3>الملخص المالي العام</h3>
      <p><strong>إجمالي القيمة:</strong> ${formatCurrency(totalAmount)}</p>
      <p><strong>إجمالي الخصومات:</strong> ${formatCurrency(totalDeductions)}</p>
      <p><strong>الصافي المستحق:</strong> ${formatCurrency(totalNet)}</p>
      <p><strong>إجمالي المحصل:</strong> ${formatCurrency(totalReceived)}</p>
      <p><strong>إجمالي المتبقي:</strong> ${formatCurrency(totalNet - totalReceived)}</p>
      <p><strong>معدل التحصيل:</strong> ${totalNet > 0 ? ((totalReceived / totalNet) * 100).toFixed(1) : 0}%</p>
    </div>

    <table>
      <thead>
        <tr>
          <th>العميل</th>
          <th>عدد العمليات</th>
          <th>إجمالي القيمة</th>
          <th>المبلغ المحصل</th>
          <th>المبلغ المتبقي</th>
          <th>معدل التحصيل</th>
        </tr>
      </thead>
      <tbody>
        ${clientStats.map(stat => `
          <tr>
            <td>${stat.client.name}</td>
            <td>${stat.operationsCount}</td>
            <td>${formatCurrency(stat.totalAmount)}</td>
            <td>${formatCurrency(stat.totalReceived)}</td>
            <td>${formatCurrency(stat.outstanding)}</td>
            <td>${stat.totalAmount > 0 ? ((stat.totalReceived / stat.totalAmount) * 100).toFixed(1) : 0}%</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;

  exportAsHTMLToPDF('التقرير المالي', content);
};

export const exportFinancialReportToExcel = (operations: Operation[], clients: Client[]) => {
  const getClientName = (clientId: string) => {
    const client = clients.find(c => c.id === clientId);
    return client ? client.name : 'عميل غير معروف';
  };

  const clientStats = clients.map(client => {
    const clientOperations = operations.filter(op => op.clientId === client.id);
    const clientTotal = clientOperations.reduce((sum, op) => sum + op.totalAmount, 0);
    const clientReceived = clientOperations.reduce((sum, op) => sum + op.totalReceived, 0);
    
    return {
      'العميل': client.name,
      'عدد العمليات': clientOperations.length,
      'إجمالي القيمة': clientTotal,
      'المبلغ المحصل': clientReceived,
      'المبلغ المتبقي': clientTotal - clientReceived,
      'معدل التحصيل (%)': clientTotal > 0 ? ((clientReceived / clientTotal) * 100).toFixed(1) : 0
    };
  }).filter(stat => stat['عدد العمليات'] > 0);

  const ws = XLSX.utils.json_to_sheet(clientStats);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'التقرير المالي');
  XLSX.writeFile(wb, `التقرير_المالي_${new Date().toISOString().split('T')[0]}.xlsx`);
};

export const exportFinancialReportToWord = (operations: Operation[], clients: Client[]) => {
  const totalAmount = operations.reduce((sum, op) => sum + op.totalAmount, 0);
  const totalReceived = operations.reduce((sum, op) => sum + op.totalReceived, 0);
  const totalDeductions = operations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
  const totalNet = operations.reduce((sum, op) => sum + calculateNetAmount(op), 0);

  const clientStats = clients.map(client => {
    const clientOperations = operations.filter(op => op.clientId === client.id);
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

  const content = `
    <div class="summary">
      <h3>الملخص المالي العام</h3>
      <p><strong>إجمالي القيمة:</strong> <span class="highlight">${formatCurrency(totalAmount)}</span></p>
      <p><strong>إجمالي الخصومات:</strong> <span class="highlight">${formatCurrency(totalDeductions)}</span></p>
      <p><strong>الصافي المستحق:</strong> <span class="highlight">${formatCurrency(totalNet)}</span></p>
      <p><strong>إجمالي المحصل:</strong> <span class="highlight">${formatCurrency(totalReceived)}</span></p>
      <p><strong>إجمالي المتبقي:</strong> <span class="highlight">${formatCurrency(totalNet - totalReceived)}</span></p>
      <p><strong>معدل التحصيل:</strong> <span class="highlight">${totalNet > 0 ? ((totalReceived / totalNet) * 100).toFixed(1) : 0}%</span></p>
    </div>

    <h2 class="section-title">التفاصيل المالية حسب العميل</h2>
    <table>
      <thead>
        <tr>
          <th>العميل</th>
          <th>عدد العمليات</th>
          <th>إجمالي القيمة</th>
          <th>المبلغ المحصل</th>
          <th>المبلغ المتبقي</th>
          <th>معدل التحصيل</th>
        </tr>
      </thead>
      <tbody>
        ${clientStats.map(stat => `
          <tr>
            <td>${stat.client.name}</td>
            <td>${stat.operationsCount}</td>
            <td>${formatCurrency(stat.totalAmount)}</td>
            <td>${formatCurrency(stat.totalReceived)}</td>
            <td>${formatCurrency(stat.outstanding)}</td>
            <td>${stat.totalAmount > 0 ? ((stat.totalReceived / stat.totalAmount) * 100).toFixed(1) : 0}%</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;

  createWordDocument('التقرير المالي', content);
};