import jsPDF from 'jspdf';
import 'jspdf-autotable';
import * as XLSX from 'xlsx';
import { Operation, Client } from '../types';
import { formatCurrency, formatDate, calculateNetAmount, calculateTotalDeductions, calculateExecutedTotal } from './calculations';

// إعداد الخط العربي لـ PDF
const setupArabicPDFFont = (doc: jsPDF) => {
  // استخدام خط يدعم العربية
  doc.setFont('helvetica');
  doc.setFontSize(12);
  doc.setR2L(true); // تفعيل الكتابة من اليمين لليسار
};

// دالة لتنسيق النص العربي للعرض الصحيح
const formatArabicText = (text: string): string => {
  if (!text) return '';
  
  // إزالة الأحرف الخاصة التي قد تسبب مشاكل في العرض
  let formattedText = text.replace(/[\u200E\u200F\u202A-\u202E]/g, '');
  
  // التأكد من أن النص يحتوي على أحرف عربية
  const hasArabic = /[\u0600-\u06FF]/.test(formattedText);
  
  if (hasArabic) {
    // إضافة علامة اتجاه النص العربي
    formattedText = '\u202B' + formattedText + '\u202C';
  }
  
  return formattedText;
};

// دالة لإنشاء محتوى Word
const createWordDocument = (title: string, content: any[]): Blob => {
  let htmlContent = `
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
      <meta charset="UTF-8">
      <style>
        body { 
          font-family: 'Arial', 'Tahoma', sans-serif; 
          direction: rtl; 
          text-align: right;
          margin: 20px;
          line-height: 1.6;
        }
        h1 { 
          color: #2563eb; 
          text-align: center; 
          border-bottom: 2px solid #2563eb;
          padding-bottom: 10px;
          margin-bottom: 30px;
        }
        h2 { 
          color: #1e40af; 
          margin-top: 30px;
          margin-bottom: 15px;
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
          background-color: #f3f4f6; 
          font-weight: bold;
          color: #374151;
        }
        tr:nth-child(even) { 
          background-color: #f9fafb; 
        }
        .summary { 
          background-color: #eff6ff; 
          padding: 15px; 
          border-radius: 5px; 
          margin: 20px 0;
        }
        .date { 
          text-align: left; 
          color: #6b7280; 
          font-size: 10px;
        }
        .currency { 
          font-weight: bold; 
          color: #059669;
        }
        .status-active { 
          color: #059669; 
          font-weight: bold;
        }
        .status-inactive { 
          color: #dc2626; 
          font-weight: bold;
        }
        .status-pending { 
          color: #d97706; 
          font-weight: bold;
        }
      </style>
    </head>
    <body>
      <h1>${title}</h1>
      <div class="date">تاريخ التقرير: ${formatDate(new Date())}</div>
  `;

  // إضافة المحتوى
  content.forEach(section => {
    if (section.type === 'summary') {
      htmlContent += `
        <div class="summary">
          <h2>${section.title}</h2>
          ${section.items.map(item => `
            <p><strong>${item.label}:</strong> ${item.value}</p>
          `).join('')}
        </div>
      `;
    } else if (section.type === 'table') {
      htmlContent += `
        <h2>${section.title}</h2>
        <table>
          <thead>
            <tr>
              ${section.headers.map(header => `<th>${header}</th>`).join('')}
            </tr>
          </thead>
          <tbody>
            ${section.rows.map(row => `
              <tr>
                ${row.map(cell => `<td>${cell}</td>`).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
      `;
    }
  });

  htmlContent += `
    </body>
    </html>
  `;

  return new Blob([htmlContent], { 
    type: 'application/msword;charset=utf-8' 
  });
};

// دالة لحفظ ملف Word
const saveWordDocument = (blob: Blob, filename: string) => {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

// تصدير تفاصيل العملية إلى PDF
export const exportOperationDetailsToPDF = (operation: Operation, client: Client) => {
  const doc = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4'
  });
  
  setupArabicPDFFont(doc);

  // العنوان
  doc.setFontSize(18);
  const titleText = `تفاصيل العملية: ${formatArabicText(operation.name)}`;
  doc.text(titleText, doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  
  // معلومات العملية
  doc.setFontSize(12);
  let yPosition = 40;
  
  doc.text(`اسم العملية: ${formatArabicText(operation.name)}`, 20, yPosition);
  yPosition += 10;
  doc.text(`كود العملية: ${operation.code}`, 20, yPosition);
  yPosition += 10;
  doc.text(`العميل: ${formatArabicText(client.name)}`, 20, yPosition);
  yPosition += 10;
  
  const clientTypeLabels = {
    'owner': 'مالك',
    'main_contractor': 'مقاول رئيسي',
    'consultant': 'استشاري'
  };
  doc.text(`نوع العميل: ${clientTypeLabels[client.type] || client.type}`, 20, yPosition);
  yPosition += 10;
  doc.text(`تاريخ الإنشاء: ${formatDate(operation.createdAt)}`, 20, yPosition);
  yPosition += 15;

  // الملخص المالي
  const executedAmount = calculateExecutedTotal(operation.items);
  const totalDeductions = calculateTotalDeductions(operation);
  const netAmount = calculateNetAmount(operation);
  const remainingAmount = netAmount - operation.totalReceived;

  doc.setFontSize(14);
  doc.text('الملخص المالي:', 20, yPosition);
  yPosition += 10;
  
  doc.setFontSize(11);
  doc.text(`إجمالي القيمة: ${formatCurrency(operation.totalAmount)}`, 30, yPosition);
  yPosition += 8;
  doc.text(`المبلغ المنفذ: ${formatCurrency(executedAmount)}`, 30, yPosition);
  yPosition += 8;
  doc.text(`إجمالي الخصومات: ${formatCurrency(totalDeductions)}`, 30, yPosition);
  yPosition += 8;
  doc.text(`الصافي المستحق: ${formatCurrency(netAmount)}`, 30, yPosition);
  yPosition += 8;
  doc.text(`المبلغ المحصل: ${formatCurrency(operation.totalReceived)}`, 30, yPosition);
  yPosition += 8;
  doc.text(`المبلغ المتبقي: ${formatCurrency(remainingAmount)}`, 30, yPosition);
  yPosition += 8;
  doc.text(`نسبة الإنجاز: ${operation.overallExecutionPercentage.toFixed(1)}%`, 30, yPosition);
  yPosition += 15;

  // بنود العملية
  if (operation.items.length > 0) {
    doc.setFontSize(14);
    doc.text('بنود العملية:', 20, yPosition);
    yPosition += 10;

    const itemsData = operation.items.map(item => [
      item.code,
      formatArabicText(item.description),
      formatCurrency(item.amount),
      `${item.executionPercentage}%`,
      formatCurrency(item.amount * (item.executionPercentage / 100))
    ]);

    (doc as any).autoTable({
      head: [['الكود', 'الوصف', 'القيمة', 'نسبة التنفيذ', 'القيمة المنفذة']],
      body: itemsData,
      startY: yPosition,
      styles: {
        fontSize: 9,
        cellPadding: 3,
        halign: 'right',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [66, 139, 202],
        textColor: 255,
        fontSize: 10,
        fontStyle: 'bold',
        halign: 'right'
      },
      alternateRowStyles: {
        fillColor: [245, 245, 245]
      },
      tableLineColor: [200, 200, 200],
      tableLineWidth: 0.1
    });

    yPosition = (doc as any).lastAutoTable.finalY + 15;
  }

  const fileName = `تفاصيل-العملية-${operation.code}-${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
};

// تصدير تفاصيل العملية إلى Word
export const exportOperationDetailsToWord = (operation: Operation, client: Client) => {
  const executedAmount = calculateExecutedTotal(operation.items);
  const totalDeductions = calculateTotalDeductions(operation);
  const netAmount = calculateNetAmount(operation);
  const remainingAmount = netAmount - operation.totalReceived;

  const clientTypeLabels = {
    'owner': 'مالك',
    'main_contractor': 'مقاول رئيسي',
    'consultant': 'استشاري'
  };

  const content = [
    {
      type: 'summary',
      title: 'معلومات العملية',
      items: [
        { label: 'اسم العملية', value: operation.name },
        { label: 'كود العملية', value: operation.code },
        { label: 'العميل', value: client.name },
        { label: 'نوع العميل', value: clientTypeLabels[client.type] || client.type },
        { label: 'تاريخ الإنشاء', value: formatDate(operation.createdAt) }
      ]
    },
    {
      type: 'summary',
      title: 'الملخص المالي',
      items: [
        { label: 'إجمالي القيمة', value: `<span class="currency">${formatCurrency(operation.totalAmount)}</span>` },
        { label: 'المبلغ المنفذ', value: `<span class="currency">${formatCurrency(executedAmount)}</span>` },
        { label: 'إجمالي الخصومات', value: `<span class="currency">${formatCurrency(totalDeductions)}</span>` },
        { label: 'الصافي المستحق', value: `<span class="currency">${formatCurrency(netAmount)}</span>` },
        { label: 'المبلغ المحصل', value: `<span class="currency">${formatCurrency(operation.totalReceived)}</span>` },
        { label: 'المبلغ المتبقي', value: `<span class="currency">${formatCurrency(remainingAmount)}</span>` },
        { label: 'نسبة الإنجاز', value: `${operation.overallExecutionPercentage.toFixed(1)}%` }
      ]
    }
  ];

  // إضافة بنود العملية
  if (operation.items.length > 0) {
    content.push({
      type: 'table',
      title: 'بنود العملية',
      headers: ['الكود', 'الوصف', 'القيمة', 'نسبة التنفيذ', 'القيمة المنفذة'],
      rows: operation.items.map(item => [
        item.code,
        item.description,
        `<span class="currency">${formatCurrency(item.amount)}</span>`,
        `${item.executionPercentage}%`,
        `<span class="currency">${formatCurrency(item.amount * (item.executionPercentage / 100))}</span>`
      ])
    });
  }

  // إضافة الخصومات
  if (operation.deductions.length > 0 && totalDeductions > 0) {
    content.push({
      type: 'table',
      title: 'الخصومات',
      headers: ['اسم الخصم', 'النوع', 'المبلغ'],
      rows: operation.deductions.filter(d => d.isActive).map(deduction => {
        const deductionAmount = deduction.type === 'percentage' 
          ? (executedAmount * deduction.value / 100)
          : deduction.value;
        
        return [
          deduction.name,
          deduction.type === 'percentage' ? `${deduction.value}%` : 'مبلغ ثابت',
          `<span class="currency">${formatCurrency(deductionAmount)}</span>`
        ];
      })
    });
  }

  // إضافة شيكات الضمان
  if (operation.guaranteeChecks.length > 0) {
    content.push({
      type: 'table',
      title: 'شيكات الضمان',
      headers: ['رقم الشيك', 'المبلغ', 'البنك', 'تاريخ الانتهاء', 'الحالة'],
      rows: operation.guaranteeChecks.map(check => [
        check.checkNumber,
        `<span class="currency">${formatCurrency(check.amount)}</span>`,
        check.bank,
        formatDate(check.expiryDate),
        check.isReturned ? '<span class="status-inactive">مُسترد</span>' : '<span class="status-active">قائم</span>'
      ])
    });
  }

  // إضافة خطابات الضمان
  if (operation.guaranteeLetters.length > 0) {
    content.push({
      type: 'table',
      title: 'خطابات الضمان',
      headers: ['رقم الخطاب', 'البنك', 'المبلغ', 'تاريخ الاستحقاق', 'الحالة'],
      rows: operation.guaranteeLetters.map(letter => [
        letter.letterNumber,
        letter.bank,
        `<span class="currency">${formatCurrency(letter.amount)}</span>`,
        formatDate(letter.dueDate),
        letter.isReturned ? '<span class="status-inactive">مُسترد</span>' : '<span class="status-active">قائم</span>'
      ])
    });
  }

  // إضافة شهادات الضمان
  if ((operation.warrantyCertificates || []).length > 0) {
    content.push({
      type: 'table',
      title: 'شهادات الضمان',
      headers: ['رقم الشهادة', 'الوصف', 'تاريخ البداية', 'تاريخ النهاية', 'الحالة'],
      rows: operation.warrantyCertificates!.map(warranty => [
        warranty.certificateNumber,
        warranty.description,
        formatDate(warranty.startDate),
        formatDate(warranty.endDate),
        warranty.isActive ? '<span class="status-active">نشط</span>' : '<span class="status-inactive">منتهي</span>'
      ])
    });
  }

  // إضافة المدفوعات المستلمة
  if (operation.receivedPayments.length > 0) {
    content.push({
      type: 'table',
      title: 'المدفوعات المستلمة',
      headers: ['النوع', 'المبلغ', 'التاريخ', 'التفاصيل'],
      rows: operation.receivedPayments.map(payment => [
        payment.type === 'cash' ? 'نقدي' : 'شيك',
        `<span class="currency">${formatCurrency(payment.amount)}</span>`,
        formatDate(payment.date),
        payment.type === 'check' && payment.checkNumber 
          ? `شيك رقم: ${payment.checkNumber} - ${payment.bank}` 
          : payment.notes || '-'
      ])
    });
  }

  const blob = createWordDocument(`تفاصيل العملية: ${operation.name}`, content);
  const fileName = `تفاصيل-العملية-${operation.code}-${new Date().toISOString().split('T')[0]}.doc`;
  saveWordDocument(blob, fileName);
};

// تصدير العمليات إلى PDF
export const exportOperationsToPDF = (operations: Operation[], clients: Client[], title: string = 'تقرير العمليات') => {
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4'
  });
  
  setupArabicPDFFont(doc);

  // العنوان
  doc.setFontSize(16);
  doc.text(formatArabicText(title), doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  
  // تاريخ التقرير
  doc.setFontSize(10);
  const dateText = `تاريخ التقرير: ${formatDate(new Date())}`;
  doc.text(dateText, 20, 35);

  // إعداد البيانات للجدول
  const tableData = operations.map(operation => {
    const client = clients.find(c => c.id === operation.clientId);
    const statusLabels = {
      'in_progress': 'قيد التنفيذ',
      'completed': 'مكتملة',
      'completed_partial_payment': 'مكتملة - دفع جزئي'
    };

    const totalDeductions = calculateTotalDeductions(operation);
    const netAmount = calculateNetAmount(operation);

    return [
      operation.code,
      formatArabicText(operation.name),
      formatArabicText(client?.name || 'غير معروف'),
      formatCurrency(operation.totalAmount),
      formatCurrency(totalDeductions),
      formatCurrency(netAmount),
      formatCurrency(operation.totalReceived),
      `${operation.overallExecutionPercentage.toFixed(1)}%`,
      statusLabels[operation.status],
      formatDate(operation.createdAt)
    ];
  });

  // إعداد الجدول
  (doc as any).autoTable({
    head: [['الكود', 'اسم العملية', 'العميل', 'إجمالي القيمة', 'الخصومات', 'الصافي المستحق', 'المحصل', 'نسبة الإنجاز', 'الحالة', 'تاريخ الإنشاء']],
    body: tableData,
    startY: 45,
    styles: {
      fontSize: 7,
      cellPadding: 2,
      overflow: 'linebreak',
      halign: 'right',
      font: 'helvetica'
    },
    headStyles: {
      fillColor: [66, 139, 202],
      textColor: 255,
      fontSize: 8,
      fontStyle: 'bold',
      halign: 'right'
    },
    alternateRowStyles: {
      fillColor: [245, 245, 245]
    },
    margin: { top: 45, right: 10, bottom: 20, left: 10 },
    tableWidth: 'auto',
    columnStyles: {
      0: { cellWidth: 25 },
      1: { cellWidth: 35 },
      2: { cellWidth: 30 },
      3: { cellWidth: 25 },
      4: { cellWidth: 25 },
      5: { cellWidth: 25 },
      6: { cellWidth: 25 },
      7: { cellWidth: 20 },
      8: { cellWidth: 25 },
      9: { cellWidth: 20 }
    }
  });

  // إضافة إحصائيات في النهاية
  const finalY = (doc as any).lastAutoTable.finalY + 20;
  const totalAmount = operations.reduce((sum, op) => sum + op.totalAmount, 0);
  const totalDeductions = operations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
  const totalNetAmount = operations.reduce((sum, op) => sum + calculateNetAmount(op), 0);
  const totalReceived = operations.reduce((sum, op) => sum + op.totalReceived, 0);
  const completedCount = operations.filter(op => op.status === 'completed').length;

  doc.setFontSize(12);
  doc.text('ملخص الإحصائيات:', 20, finalY);
  doc.setFontSize(10);
  doc.text(`إجمالي العمليات: ${operations.length}`, 20, finalY + 10);
  doc.text(`العمليات المكتملة: ${completedCount}`, 20, finalY + 20);
  doc.text(`إجمالي القيمة: ${formatCurrency(totalAmount)}`, 20, finalY + 30);
  doc.text(`إجمالي الخصومات: ${formatCurrency(totalDeductions)}`, 20, finalY + 40);
  doc.text(`إجمالي الصافي المستحق: ${formatCurrency(totalNetAmount)}`, 20, finalY + 50);
  doc.text(`إجمالي المحصل: ${formatCurrency(totalReceived)}`, 20, finalY + 60);

  const fileName = `تقرير-العمليات-${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
};

// تصدير العمليات إلى Word
export const exportOperationsToWord = (operations: Operation[], clients: Client[], title: string = 'تقرير العمليات') => {
  const totalAmount = operations.reduce((sum, op) => sum + op.totalAmount, 0);
  const totalDeductions = operations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
  const totalNetAmount = operations.reduce((sum, op) => sum + calculateNetAmount(op), 0);
  const totalReceived = operations.reduce((sum, op) => sum + op.totalReceived, 0);
  const completedCount = operations.filter(op => op.status === 'completed').length;

  const content = [
    {
      type: 'summary',
      title: 'ملخص الإحصائيات',
      items: [
        { label: 'إجمالي العمليات', value: operations.length.toString() },
        { label: 'العمليات المكتملة', value: completedCount.toString() },
        { label: 'إجمالي القيمة', value: `<span class="currency">${formatCurrency(totalAmount)}</span>` },
        { label: 'إجمالي الخصومات', value: `<span class="currency">${formatCurrency(totalDeductions)}</span>` },
        { label: 'إجمالي الصافي المستحق', value: `<span class="currency">${formatCurrency(totalNetAmount)}</span>` },
        { label: 'إجمالي المحصل', value: `<span class="currency">${formatCurrency(totalReceived)}</span>` }
      ]
    },
    {
      type: 'table',
      title: 'تفاصيل العمليات',
      headers: ['الكود', 'اسم العملية', 'العميل', 'إجمالي القيمة', 'الخصومات', 'الصافي المستحق', 'المحصل', 'نسبة الإنجاز', 'الحالة', 'تاريخ الإنشاء'],
      rows: operations.map(operation => {
        const client = clients.find(c => c.id === operation.clientId);
        const statusLabels = {
          'in_progress': '<span class="status-pending">قيد التنفيذ</span>',
          'completed': '<span class="status-active">مكتملة</span>',
          'completed_partial_payment': '<span class="status-pending">مكتملة - دفع جزئي</span>'
        };

        const totalDeductions = calculateTotalDeductions(operation);
        const netAmount = calculateNetAmount(operation);

        return [
          operation.code,
          operation.name,
          client?.name || 'غير معروف',
          `<span class="currency">${formatCurrency(operation.totalAmount)}</span>`,
          `<span class="currency">${formatCurrency(totalDeductions)}</span>`,
          `<span class="currency">${formatCurrency(netAmount)}</span>`,
          `<span class="currency">${formatCurrency(operation.totalReceived)}</span>`,
          `${operation.overallExecutionPercentage.toFixed(1)}%`,
          statusLabels[operation.status],
          formatDate(operation.createdAt)
        ];
      })
    }
  ];

  const blob = createWordDocument(title, content);
  const fileName = `تقرير-العمليات-${new Date().toISOString().split('T')[0]}.doc`;
  saveWordDocument(blob, fileName);
};

// تصدير الشيكات والمدفوعات إلى PDF
export const exportChecksAndPaymentsToPDF = (operations: Operation[], clients: Client[]) => {
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4'
  });
  
  setupArabicPDFFont(doc);

  // العنوان
  doc.setFontSize(16);
  const titleText = 'تقرير الشيكات والمدفوعات';
  doc.text(formatArabicText(titleText), doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  
  // تاريخ التقرير
  doc.setFontSize(10);
  const dateText = `تاريخ التقرير: ${formatDate(new Date())}`;
  doc.text(dateText, 20, 35);

  let yPosition = 50;

  // جمع جميع المدفوعات
  const allPayments = operations.flatMap(operation => 
    operation.receivedPayments.map(payment => {
      const client = clients.find(c => c.id === operation.clientId);
      return {
        ...payment,
        operationName: operation.name,
        operationCode: operation.code,
        clientName: client?.name || 'غير معروف'
      };
    })
  );

  if (allPayments.length > 0) {
    doc.setFontSize(14);
    doc.text('المدفوعات المستلمة:', 20, yPosition);
    yPosition += 10;

    const paymentsData = allPayments.map(payment => [
      payment.type === 'cash' ? 'نقدي' : 'شيك',
      formatCurrency(payment.amount),
      formatDate(payment.date),
      formatArabicText(payment.clientName),
      formatArabicText(payment.operationName),
      payment.type === 'check' && payment.checkNumber 
        ? `${payment.checkNumber} - ${formatArabicText(payment.bank || '')}`
        : formatArabicText(payment.notes || '-')
    ]);

    (doc as any).autoTable({
      head: [['النوع', 'المبلغ', 'التاريخ', 'العميل', 'العملية', 'التفاصيل']],
      body: paymentsData,
      startY: yPosition,
      styles: {
        fontSize: 8,
        cellPadding: 3,
        halign: 'right',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [0, 128, 0],
        textColor: 255,
        fontSize: 9,
        fontStyle: 'bold',
        halign: 'right'
      },
      alternateRowStyles: {
        fillColor: [245, 245, 245]
      }
    });

    yPosition = (doc as any).lastAutoTable.finalY + 20;
  }

  // إضافة إحصائيات
  const totalAmount = allPayments.reduce((sum, payment) => sum + payment.amount, 0);
  const totalChecks = allPayments.filter(p => p.type === 'check').length;
  const totalCash = allPayments.filter(p => p.type === 'cash').length;

  doc.setFontSize(12);
  doc.text('ملخص الإحصائيات:', 20, yPosition);
  doc.setFontSize(10);
  doc.text(`إجمالي المدفوعات: ${formatCurrency(totalAmount)}`, 20, yPosition + 10);
  doc.text(`عدد الشيكات: ${totalChecks}`, 20, yPosition + 20);
  doc.text(`المدفوعات النقدية: ${totalCash}`, 20, yPosition + 30);

  const fileName = `تقرير-الشيكات-والمدفوعات-${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
};

// تصدير الشيكات والمدفوعات إلى Word
export const exportChecksAndPaymentsToWord = (operations: Operation[], clients: Client[]) => {
  // جمع جميع المدفوعات
  const allPayments = operations.flatMap(operation => 
    operation.receivedPayments.map(payment => {
      const client = clients.find(c => c.id === operation.clientId);
      return {
        ...payment,
        operationName: operation.name,
        operationCode: operation.code,
        clientName: client?.name || 'غير معروف'
      };
    })
  );

  const totalAmount = allPayments.reduce((sum, payment) => sum + payment.amount, 0);
  const totalChecks = allPayments.filter(p => p.type === 'check').length;
  const totalCash = allPayments.filter(p => p.type === 'cash').length;

  const content = [
    {
      type: 'summary',
      title: 'ملخص الإحصائيات',
      items: [
        { label: 'إجمالي المدفوعات', value: `<span class="currency">${formatCurrency(totalAmount)}</span>` },
        { label: 'عدد الشيكات', value: totalChecks.toString() },
        { label: 'المدفوعات النقدية', value: totalCash.toString() }
      ]
    },
    {
      type: 'table',
      title: 'تفاصيل المدفوعات',
      headers: ['النوع', 'المبلغ', 'التاريخ', 'العميل', 'العملية', 'التفاصيل'],
      rows: allPayments.map(payment => [
        payment.type === 'cash' ? 'نقدي' : 'شيك',
        `<span class="currency">${formatCurrency(payment.amount)}</span>`,
        formatDate(payment.date),
        payment.clientName,
        payment.operationName,
        payment.type === 'check' && payment.checkNumber 
          ? `شيك رقم: ${payment.checkNumber} - ${payment.bank || ''}`
          : payment.notes || '-'
      ])
    }
  ];

  const blob = createWordDocument('تقرير الشيكات والمدفوعات', content);
  const fileName = `تقرير-الشيكات-والمدفوعات-${new Date().toISOString().split('T')[0]}.doc`;
  saveWordDocument(blob, fileName);
};

// تصدير الشيكات والمدفوعات إلى Excel
export const exportChecksAndPaymentsToExcel = (operations: Operation[], clients: Client[]) => {
  // جمع جميع المدفوعات
  const allPayments = operations.flatMap(operation => 
    operation.receivedPayments.map(payment => {
      const client = clients.find(c => c.id === operation.clientId);
      return {
        'نوع الدفع': payment.type === 'cash' ? 'نقدي' : 'شيك',
        'المبلغ': payment.amount,
        'التاريخ': formatDate(payment.date),
        'العميل': client?.name || 'غير معروف',
        'العملية': operation.name,
        'كود العملية': operation.code,
        'رقم الشيك': payment.checkNumber || '',
        'البنك': payment.bank || '',
        'تاريخ الاستلام': payment.receiptDate ? formatDate(payment.receiptDate) : '',
        'ملاحظات': payment.notes || ''
      };
    })
  );

  const worksheet = XLSX.utils.json_to_sheet(allPayments);
  worksheet['!cols'] = [
    { wch: 12 }, // نوع الدفع
    { wch: 15 }, // المبلغ
    { wch: 12 }, // التاريخ
    { wch: 25 }, // العميل
    { wch: 30 }, // العملية
    { wch: 15 }, // كود العملية
    { wch: 15 }, // رقم الشيك
    { wch: 20 }, // البنك
    { wch: 15 }, // تاريخ الاستلام
    { wch: 30 }  // ملاحظات
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'الشيكات والمدفوعات');

  // إضافة ورقة الإحصائيات
  const stats = [
    { 'البند': 'إجمالي المدفوعات', 'القيمة': allPayments.reduce((sum, p) => sum + p.المبلغ, 0) },
    { 'البند': 'عدد الشيكات', 'القيمة': allPayments.filter(p => p['نوع الدفع'] === 'شيك').length },
    { 'البند': 'المدفوعات النقدية', 'القيمة': allPayments.filter(p => p['نوع الدفع'] === 'نقدي').length }
  ];

  const statsWorksheet = XLSX.utils.json_to_sheet(stats);
  statsWorksheet['!cols'] = [{ wch: 20 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(workbook, statsWorksheet, 'الإحصائيات');

  const fileName = `تقرير-الشيكات-والمدفوعات-${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};

// تصدير تقرير مفصل للضمانات
export const exportDetailedGuaranteesReportToPDF = (operations: Operation[], clients: Client[]) => {
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4'
  });
  
  setupArabicPDFFont(doc);

  // العنوان
  doc.setFontSize(16);
  const titleText = 'تقرير الضمانات المفصل';
  doc.text(formatArabicText(titleText), doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  
  // تاريخ التقرير
  doc.setFontSize(10);
  const dateText = `تاريخ التقرير: ${formatDate(new Date())}`;
  doc.text(dateText, 20, 35);

  let yPosition = 50;

  // تقرير شيكات الضمان
  doc.setFontSize(14);
  doc.text('شيكات الضمان:', 20, yPosition);
  yPosition += 10;

  // جمع جميع شيكات الضمان
  const allGuaranteeChecks = operations.flatMap(operation => 
    operation.guaranteeChecks.map(check => {
      const client = clients.find(c => c.id === operation.clientId);
      const relatedItem = check.relatedTo === 'item' && check.relatedItemId 
        ? operation.items.find(item => item.id === check.relatedItemId)
        : null;
      
      return {
        ...check,
        operationName: operation.name,
        operationCode: operation.code,
        clientName: client?.name || 'غير معروف',
        relatedItemDescription: relatedItem?.description || 'العملية كاملة'
      };
    })
  );

  if (allGuaranteeChecks.length > 0) {
    const checksTableData = allGuaranteeChecks.map(check => [
      check.checkNumber,
      formatArabicText(check.clientName),
      formatArabicText(check.operationName),
      formatArabicText(check.relatedItemDescription),
      formatCurrency(check.amount),
      formatDate(check.checkDate),
      formatDate(check.expiryDate),
      formatArabicText(check.bank),
      check.isReturned ? 'مُسترد' : 'قائم'
    ]);

    (doc as any).autoTable({
      head: [['رقم الشيك', 'العميل', 'العملية', 'البند', 'المبلغ', 'تاريخ الإصدار', 'تاريخ الانتهاء', 'البنك', 'الحالة']],
      body: checksTableData,
      startY: yPosition,
      styles: {
        fontSize: 7,
        cellPadding: 2,
        overflow: 'linebreak',
        halign: 'right',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [66, 139, 202],
        textColor: 255,
        fontSize: 8,
        fontStyle: 'bold',
        halign: 'right'
      },
      alternateRowStyles: {
        fillColor: [245, 245, 245]
      },
      columnStyles: {
        0: { cellWidth: 25 },
        1: { cellWidth: 30 },
        2: { cellWidth: 35 },
        3: { cellWidth: 35 },
        4: { cellWidth: 25 },
        5: { cellWidth: 25 },
        6: { cellWidth: 25 },
        7: { cellWidth: 25 },
        8: { cellWidth: 20 }
      }
    });

    yPosition = (doc as any).lastAutoTable.finalY + 20;
  }

  // إضافة صفحة جديدة إذا لزم الأمر
  if (yPosition > 180) {
    doc.addPage();
    yPosition = 30;
  }

  // تقرير خطابات الضمان
  doc.setFontSize(14);
  doc.text('خطابات الضمان:', 20, yPosition);
  yPosition += 10;

  // جمع جميع خطابات الضمان
  const allGuaranteeLetters = operations.flatMap(operation => 
    operation.guaranteeLetters.map(letter => {
      const client = clients.find(c => c.id === operation.clientId);
      const relatedItem = letter.relatedTo === 'item' && letter.relatedItemId 
        ? operation.items.find(item => item.id === letter.relatedItemId)
        : null;
      
      return {
        ...letter,
        operationName: operation.name,
        operationCode: operation.code,
        clientName: client?.name || 'غير معروف',
        relatedItemDescription: relatedItem?.description || 'العملية كاملة'
      };
    })
  );

  if (allGuaranteeLetters.length > 0) {
    const lettersTableData = allGuaranteeLetters.map(letter => [
      letter.letterNumber,
      formatArabicText(letter.clientName),
      formatArabicText(letter.operationName),
      formatArabicText(letter.relatedItemDescription),
      formatCurrency(letter.amount),
      formatDate(letter.letterDate),
      formatDate(letter.dueDate),
      formatArabicText(letter.bank),
      letter.isReturned ? 'مُسترد' : 'قائم'
    ]);

    (doc as any).autoTable({
      head: [['رقم الخطاب', 'العميل', 'العملية', 'البند', 'المبلغ', 'تاريخ الإصدار', 'تاريخ الاستحقاق', 'البنك', 'الحالة']],
      body: lettersTableData,
      startY: yPosition,
      styles: {
        fontSize: 7,
        cellPadding: 2,
        overflow: 'linebreak',
        halign: 'right',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [139, 69, 19],
        textColor: 255,
        fontSize: 8,
        fontStyle: 'bold',
        halign: 'right'
      },
      alternateRowStyles: {
        fillColor: [245, 245, 245]
      },
      columnStyles: {
        0: { cellWidth: 25 },
        1: { cellWidth: 30 },
        2: { cellWidth: 35 },
        3: { cellWidth: 35 },
        4: { cellWidth: 25 },
        5: { cellWidth: 25 },
        6: { cellWidth: 25 },
        7: { cellWidth: 25 },
        8: { cellWidth: 20 }
      }
    });
  }

  const fileName = `تقرير-الضمانات-المفصل-${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
};

// تصدير تقرير مفصل للضمانات إلى Word
export const exportDetailedGuaranteesReportToWord = (operations: Operation[], clients: Client[]) => {
  // جمع جميع شيكات الضمان
  const allGuaranteeChecks = operations.flatMap(operation => 
    operation.guaranteeChecks.map(check => {
      const client = clients.find(c => c.id === operation.clientId);
      const relatedItem = check.relatedTo === 'item' && check.relatedItemId 
        ? operation.items.find(item => item.id === check.relatedItemId)
        : null;
      
      return {
        ...check,
        operationName: operation.name,
        operationCode: operation.code,
        clientName: client?.name || 'غير معروف',
        relatedItemDescription: relatedItem?.description || 'العملية كاملة'
      };
    })
  );

  // جمع جميع خطابات الضمان
  const allGuaranteeLetters = operations.flatMap(operation => 
    operation.guaranteeLetters.map(letter => {
      const client = clients.find(c => c.id === operation.clientId);
      const relatedItem = letter.relatedTo === 'item' && letter.relatedItemId 
        ? operation.items.find(item => item.id === letter.relatedItemId)
        : null;
      
      return {
        ...letter,
        operationName: operation.name,
        operationCode: operation.code,
        clientName: client?.name || 'غير معروف',
        relatedItemDescription: relatedItem?.description || 'العملية كاملة'
      };
    })
  );

  const content = [
    {
      type: 'summary',
      title: 'ملخص الضمانات',
      items: [
        { label: 'إجمالي شيكات الضمان', value: allGuaranteeChecks.length.toString() },
        { label: 'إجمالي خطابات الضمان', value: allGuaranteeLetters.length.toString() },
        { label: 'الضمانات النشطة', value: (allGuaranteeChecks.filter(c => !c.isReturned).length + allGuaranteeLetters.filter(l => !l.isReturned).length).toString() },
        { label: 'الضمانات المستردة', value: (allGuaranteeChecks.filter(c => c.isReturned).length + allGuaranteeLetters.filter(l => l.isReturned).length).toString() }
      ]
    }
  ];

  // إضافة جدول شيكات الضمان
  if (allGuaranteeChecks.length > 0) {
    content.push({
      type: 'table',
      title: 'شيكات الضمان',
      headers: ['رقم الشيك', 'العميل', 'العملية', 'البند', 'المبلغ', 'تاريخ الإصدار', 'تاريخ الانتهاء', 'البنك', 'الحالة'],
      rows: allGuaranteeChecks.map(check => [
        check.checkNumber,
        check.clientName,
        check.operationName,
        check.relatedItemDescription,
        `<span class="currency">${formatCurrency(check.amount)}</span>`,
        formatDate(check.checkDate),
        formatDate(check.expiryDate),
        check.bank,
        check.isReturned ? '<span class="status-inactive">مُسترد</span>' : '<span class="status-active">قائم</span>'
      ])
    });
  }

  // إضافة جدول خطابات الضمان
  if (allGuaranteeLetters.length > 0) {
    content.push({
      type: 'table',
      title: 'خطابات الضمان',
      headers: ['رقم الخطاب', 'العميل', 'العملية', 'البند', 'المبلغ', 'تاريخ الإصدار', 'تاريخ الاستحقاق', 'البنك', 'الحالة'],
      rows: allGuaranteeLetters.map(letter => [
        letter.letterNumber,
        letter.clientName,
        letter.operationName,
        letter.relatedItemDescription,
        `<span class="currency">${formatCurrency(letter.amount)}</span>`,
        formatDate(letter.letterDate),
        formatDate(letter.dueDate),
        letter.bank,
        letter.isReturned ? '<span class="status-inactive">مُسترد</span>' : '<span class="status-active">قائم</span>'
      ])
    });
  }

  const blob = createWordDocument('تقرير الضمانات المفصل', content);
  const fileName = `تقرير-الضمانات-المفصل-${new Date().toISOString().split('T')[0]}.doc`;
  saveWordDocument(blob, fileName);
};

// تصدير الضمانات إلى Excel
export const exportGuaranteesToExcel = (operations: Operation[], clients: Client[]) => {
  // جمع جميع شيكات الضمان
  const allGuaranteeChecks = operations.flatMap(operation => 
    operation.guaranteeChecks.map(check => {
      const client = clients.find(c => c.id === operation.clientId);
      const relatedItem = check.relatedTo === 'item' && check.relatedItemId 
        ? operation.items.find(item => item.id === check.relatedItemId)
        : null;
      
      return {
        'النوع': 'شيك ضمان',
        'الرقم': check.checkNumber,
        'المبلغ': check.amount,
        'البنك': check.bank,
        'العميل': client?.name || 'غير معروف',
        'العملية': operation.name,
        'كود العملية': operation.code,
        'البند المرتبط': relatedItem?.description || 'العملية كاملة',
        'تاريخ الإصدار': formatDate(check.checkDate),
        'تاريخ التسليم': formatDate(check.deliveryDate),
        'تاريخ الانتهاء': formatDate(check.expiryDate),
        'الحالة': check.isReturned ? 'مُسترد' : 'قائم',
        'تاريخ الاسترداد': check.returnDate ? formatDate(check.returnDate) : ''
      };
    })
  );

  // جمع جميع خطابات الضمان
  const allGuaranteeLetters = operations.flatMap(operation => 
    operation.guaranteeLetters.map(letter => {
      const client = clients.find(c => c.id === operation.clientId);
      const relatedItem = letter.relatedTo === 'item' && letter.relatedItemId 
        ? operation.items.find(item => item.id === letter.relatedItemId)
        : null;
      
      return {
        'النوع': 'خطاب ضمان',
        'الرقم': letter.letterNumber,
        'المبلغ': letter.amount,
        'البنك': letter.bank,
        'العميل': client?.name || 'غير معروف',
        'العملية': operation.name,
        'كود العملية': operation.code,
        'البند المرتبط': relatedItem?.description || 'العملية كاملة',
        'تاريخ الإصدار': formatDate(letter.letterDate),
        'تاريخ التسليم': '',
        'تاريخ الانتهاء': formatDate(letter.dueDate),
        'الحالة': letter.isReturned ? 'مُسترد' : 'قائم',
        'تاريخ الاسترداد': letter.returnDate ? formatDate(letter.returnDate) : '',
        'ملاحظات': letter.notes || ''
      };
    })
  );

  // دمج جميع الضمانات
  const allGuarantees = [...allGuaranteeChecks, ...allGuaranteeLetters];

  const worksheet = XLSX.utils.json_to_sheet(allGuarantees);
  worksheet['!cols'] = [
    { wch: 15 }, // النوع
    { wch: 15 }, // الرقم
    { wch: 15 }, // المبلغ
    { wch: 20 }, // البنك
    { wch: 25 }, // العميل
    { wch: 30 }, // العملية
    { wch: 15 }, // كود العملية
    { wch: 30 }, // البند المرتبط
    { wch: 12 }, // تاريخ الإصدار
    { wch: 12 }, // تاريخ التسليم
    { wch: 12 }, // تاريخ الانتهاء
    { wch: 10 }, // الحالة
    { wch: 12 }, // تاريخ الاسترداد
    { wch: 30 }  // ملاحظات
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'جميع الضمانات');

  // إضافة ورقة منفصلة لشيكات الضمان
  if (allGuaranteeChecks.length > 0) {
    const checksWorksheet = XLSX.utils.json_to_sheet(allGuaranteeChecks);
    checksWorksheet['!cols'] = worksheet['!cols'];
    XLSX.utils.book_append_sheet(workbook, checksWorksheet, 'شيكات الضمان');
  }

  // إضافة ورقة منفصلة لخطابات الضمان
  if (allGuaranteeLetters.length > 0) {
    const lettersWorksheet = XLSX.utils.json_to_sheet(allGuaranteeLetters);
    lettersWorksheet['!cols'] = worksheet['!cols'];
    XLSX.utils.book_append_sheet(workbook, lettersWorksheet, 'خطابات الضمان');
  }

  // إضافة ورقة الإحصائيات
  const stats = [
    { 'البند': 'إجمالي الضمانات', 'القيمة': allGuarantees.length },
    { 'البند': 'شيكات الضمان', 'القيمة': allGuaranteeChecks.length },
    { 'البند': 'خطابات الضمان', 'القيمة': allGuaranteeLetters.length },
    { 'البند': 'الضمانات النشطة', 'القيمة': allGuarantees.filter(g => g.الحالة === 'قائم').length },
    { 'البند': 'الضمانات المستردة', 'القيمة': allGuarantees.filter(g => g.الحالة === 'مُسترد').length },
    { 'البند': 'إجمالي المبلغ', 'القيمة': allGuarantees.reduce((sum, g) => sum + g.المبلغ, 0) }
  ];

  const statsWorksheet = XLSX.utils.json_to_sheet(stats);
  statsWorksheet['!cols'] = [{ wch: 20 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(workbook, statsWorksheet, 'الإحصائيات');

  const fileName = `تقرير-الضمانات-${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};

// تصدير تقرير شهادات الضمان
export const exportWarrantyCertificatesReportToPDF = (operations: Operation[], clients: Client[]) => {
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4'
  });
  
  setupArabicPDFFont(doc);

  // العنوان
  doc.setFontSize(16);
  const titleText = 'تقرير شهادات الضمان';
  doc.text(formatArabicText(titleText), doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  
  // تاريخ التقرير
  doc.setFontSize(10);
  const dateText = `تاريخ التقرير: ${formatDate(new Date())}`;
  doc.text(dateText, 20, 35);

  // جمع جميع شهادات الضمان
  const allWarranties = operations.flatMap(operation => 
    (operation.warrantyCertificates || []).map(warranty => {
      const client = clients.find(c => c.id === operation.clientId);
      const relatedItem = warranty.relatedTo === 'item' && warranty.relatedItemId 
        ? operation.items.find(item => item.id === warranty.relatedItemId)
        : null;
      
      return {
        ...warranty,
        operationName: operation.name,
        operationCode: operation.code,
        clientName: client?.name || 'غير معروف',
        relatedItemDescription: relatedItem?.description || 'العملية كاملة'
      };
    })
  );

  if (allWarranties.length > 0) {
    const warrantyTableData = allWarranties.map(warranty => [
      warranty.certificateNumber,
      formatArabicText(warranty.clientName),
      formatArabicText(warranty.operationName),
      formatArabicText(warranty.relatedItemDescription),
      formatArabicText(warranty.description),
      formatDate(warranty.startDate),
      formatDate(warranty.endDate),
      `${warranty.warrantyPeriodMonths} شهر`,
      warranty.isActive ? 'نشط' : 'منتهي'
    ]);

    (doc as any).autoTable({
      head: [['رقم الشهادة', 'العميل', 'العملية', 'البند', 'الوصف', 'تاريخ البداية', 'تاريخ النهاية', 'المدة', 'الحالة']],
      body: warrantyTableData,
      startY: 50,
      styles: {
        fontSize: 7,
        cellPadding: 2,
        overflow: 'linebreak',
        halign: 'right',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [34, 139, 34],
        textColor: 255,
        fontSize: 8,
        fontStyle: 'bold',
        halign: 'right'
      },
      alternateRowStyles: {
        fillColor: [245, 245, 245]
      },
      columnStyles: {
        0: { cellWidth: 25 },
        1: { cellWidth: 30 },
        2: { cellWidth: 35 },
        3: { cellWidth: 35 },
        4: { cellWidth: 40 },
        5: { cellWidth: 25 },
        6: { cellWidth: 25 },
        7: { cellWidth: 20 },
        8: { cellWidth: 20 }
      }
    });
  } else {
    doc.setFontSize(12);
    doc.text('لا توجد شهادات ضمان', 20, 60);
  }

  const fileName = `تقرير-شهادات-الضمان-${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
};

// تصدير شهادات الضمان إلى Word
export const exportWarrantyCertificatesReportToWord = (operations: Operation[], clients: Client[]) => {
  // جمع جميع شهادات الضمان
  const allWarranties = operations.flatMap(operation => 
    (operation.warrantyCertificates || []).map(warranty => {
      const client = clients.find(c => c.id === operation.clientId);
      const relatedItem = warranty.relatedTo === 'item' && warranty.relatedItemId 
        ? operation.items.find(item => item.id === warranty.relatedItemId)
        : null;
      
      return {
        ...warranty,
        operationName: operation.name,
        operationCode: operation.code,
        clientName: client?.name || 'غير معروف',
        relatedItemDescription: relatedItem?.description || 'العملية كاملة'
      };
    })
  );

  const content = [
    {
      type: 'summary',
      title: 'ملخص شهادات الضمان',
      items: [
        { label: 'إجمالي الشهادات', value: allWarranties.length.toString() },
        { label: 'الشهادات النشطة', value: allWarranties.filter(w => w.isActive).length.toString() },
        { label: 'الشهادات المنتهية', value: allWarranties.filter(w => !w.isActive).length.toString() }
      ]
    }
  ];

  if (allWarranties.length > 0) {
    content.push({
      type: 'table',
      title: 'تفاصيل شهادات الضمان',
      headers: ['رقم الشهادة', 'العميل', 'العملية', 'البند', 'الوصف', 'تاريخ البداية', 'تاريخ النهاية', 'المدة', 'الحالة'],
      rows: allWarranties.map(warranty => [
        warranty.certificateNumber,
        warranty.clientName,
        warranty.operationName,
        warranty.relatedItemDescription,
        warranty.description,
        formatDate(warranty.startDate),
        formatDate(warranty.endDate),
        `${warranty.warrantyPeriodMonths} شهر`,
        warranty.isActive ? '<span class="status-active">نشط</span>' : '<span class="status-inactive">منتهي</span>'
      ])
    });
  }

  const blob = createWordDocument('تقرير شهادات الضمان', content);
  const fileName = `تقرير-شهادات-الضمان-${new Date().toISOString().split('T')[0]}.doc`;
  saveWordDocument(blob, fileName);
};

// تصدير شهادات الضمان إلى Excel
export const exportWarrantyCertificatesToExcel = (operations: Operation[], clients: Client[]) => {
  // جمع جميع شهادات الضمان
  const allWarranties = operations.flatMap(operation => 
    (operation.warrantyCertificates || []).map(warranty => {
      const client = clients.find(c => c.id === operation.clientId);
      const relatedItem = warranty.relatedTo === 'item' && warranty.relatedItemId 
        ? operation.items.find(item => item.id === warranty.relatedItemId)
        : null;
      
      return {
        'رقم الشهادة': warranty.certificateNumber,
        'العميل': client?.name || 'غير معروف',
        'العملية': operation.name,
        'كود العملية': operation.code,
        'البند المرتبط': relatedItem?.description || 'العملية كاملة',
        'الوصف': warranty.description,
        'تاريخ الإصدار': formatDate(warranty.issueDate),
        'تاريخ البداية': formatDate(warranty.startDate),
        'تاريخ النهاية': formatDate(warranty.endDate),
        'مدة الضمان (بالأشهر)': warranty.warrantyPeriodMonths,
        'الحالة': warranty.isActive ? 'نشط' : 'منتهي',
        'ملاحظات': warranty.notes || ''
      };
    })
  );

  const worksheet = XLSX.utils.json_to_sheet(allWarranties);
  worksheet['!cols'] = [
    { wch: 20 }, // رقم الشهادة
    { wch: 25 }, // العميل
    { wch: 30 }, // العملية
    { wch: 15 }, // كود العملية
    { wch: 30 }, // البند المرتبط
    { wch: 40 }, // الوصف
    { wch: 12 }, // تاريخ الإصدار
    { wch: 12 }, // تاريخ البداية
    { wch: 12 }, // تاريخ النهاية
    { wch: 15 }, // مدة الضمان
    { wch: 10 }, // الحالة
    { wch: 30 }  // ملاحظات
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'شهادات الضمان');

  // إضافة ورقة الإحصائيات
  const stats = [
    { 'البند': 'إجمالي الشهادات', 'القيمة': allWarranties.length },
    { 'البند': 'الشهادات النشطة', 'القيمة': allWarranties.filter(w => w.الحالة === 'نشط').length },
    { 'البند': 'الشهادات المنتهية', 'القيمة': allWarranties.filter(w => w.الحالة === 'منتهي').length }
  ];

  const statsWorksheet = XLSX.utils.json_to_sheet(stats);
  statsWorksheet['!cols'] = [{ wch: 20 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(workbook, statsWorksheet, 'الإحصائيات');

  const fileName = `شهادات-الضمان-${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};

// باقي الدوال مع نفس التحسينات...
export const exportOperationsToExcel = (operations: Operation[], clients: Client[], title: string = 'تقرير العمليات') => {
  // إعداد البيانات
  const data = operations.map(operation => {
    const client = clients.find(c => c.id === operation.clientId);
    const statusLabels = {
      'in_progress': 'قيد التنفيذ',
      'completed': 'مكتملة',
      'completed_partial_payment': 'مكتملة - دفع جزئي'
    };

    const totalDeductions = calculateTotalDeductions(operation);
    const netAmount = calculateNetAmount(operation);

    return {
      'كود العملية': operation.code,
      'اسم العملية': operation.name,
      'العميل': client?.name || 'غير معروف',
      'نوع العميل': client?.type === 'owner' ? 'مالك' : client?.type === 'main_contractor' ? 'مقاول رئيسي' : 'استشاري',
      'إجمالي القيمة': operation.totalAmount,
      'إجمالي الخصومات': totalDeductions,
      'الصافي المستحق': netAmount,
      'المبلغ المحصل': operation.totalReceived,
      'المبلغ المتبقي': netAmount - operation.totalReceived,
      'نسبة الإنجاز': `${operation.overallExecutionPercentage.toFixed(1)}%`,
      'الحالة': statusLabels[operation.status],
      'تاريخ الإنشاء': formatDate(operation.createdAt),
      'تاريخ التحديث': formatDate(operation.updatedAt),
      'عدد البنود': operation.items.length,
      'شيكات الضمان': operation.guaranteeChecks.length,
      'خطابات الضمان': operation.guaranteeLetters.length,
      'شهادات الضمان': (operation.warrantyCertificates || []).length,
      'عدد المدفوعات': operation.receivedPayments.length
    };
  });

  // إنشاء ورقة العمل
  const worksheet = XLSX.utils.json_to_sheet(data);

  // تنسيق العرض
  const columnWidths = [
    { wch: 15 }, // كود العملية
    { wch: 25 }, // اسم العملية
    { wch: 20 }, // العميل
    { wch: 15 }, // نوع العميل
    { wch: 15 }, // إجمالي القيمة
    { wch: 15 }, // إجمالي الخصومات
    { wch: 15 }, // الصافي المستحق
    { wch: 15 }, // المبلغ المحصل
    { wch: 15 }, // المبلغ المتبقي
    { wch: 12 }, // نسبة الإنجاز
    { wch: 18 }, // الحالة
    { wch: 12 }, // تاريخ الإنشاء
    { wch: 12 }, // تاريخ التحديث
    { wch: 10 }, // عدد البنود
    { wch: 12 }, // شيكات الضمان
    { wch: 12 }, // خطابات الضمان
    { wch: 12 }, // شهادات الضمان
    { wch: 12 }  // عدد المدفوعات
  ];
  worksheet['!cols'] = columnWidths;

  // إنشاء المصنف
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'العمليات');

  // إضافة ورقة الإحصائيات
  const totalAmount = operations.reduce((sum, op) => sum + op.totalAmount, 0);
  const totalDeductions = operations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
  const totalNetAmount = operations.reduce((sum, op) => sum + calculateNetAmount(op), 0);
  const totalReceived = operations.reduce((sum, op) => sum + op.totalReceived, 0);

  const stats = [
    { 'البند': 'إجمالي العمليات', 'القيمة': operations.length },
    { 'البند': 'العمليات المكتملة', 'القيمة': operations.filter(op => op.status === 'completed').length },
    { 'البند': 'العمليات قيد التنفيذ', 'القيمة': operations.filter(op => op.status === 'in_progress').length },
    { 'البند': 'إجمالي القيمة', 'القيمة': totalAmount },
    { 'البند': 'إجمالي الخصومات', 'القيمة': totalDeductions },
    { 'البند': 'إجمالي الصافي المستحق', 'القيمة': totalNetAmount },
    { 'البند': 'إجمالي المحصل', 'القيمة': totalReceived },
    { 'البند': 'إجمالي المتبقي', 'القيمة': totalNetAmount - totalReceived }
  ];

  const statsWorksheet = XLSX.utils.json_to_sheet(stats);
  statsWorksheet['!cols'] = [{ wch: 20 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(workbook, statsWorksheet, 'الإحصائيات');

  // حفظ الملف
  const fileName = `تقرير-العمليات-${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};

// تصدير العملاء إلى PDF
export const exportClientsToPDF = (clients: Client[], title: string = 'تقرير العملاء') => {
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4'
  });
  
  setupArabicPDFFont(doc);

  // العنوان
  doc.setFontSize(16);
  doc.text(formatArabicText(title), doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  
  // تاريخ التقرير
  doc.setFontSize(10);
  const dateText = `تاريخ التقرير: ${formatDate(new Date())}`;
  doc.text(dateText, 20, 35);

  // إعداد البيانات للجدول
  const tableData = clients.map(client => {
    const clientTypeLabels = {
      'owner': 'مالك',
      'main_contractor': 'مقاول رئيسي',
      'consultant': 'استشاري'
    };
    
    const mainContact = client.contacts?.find(contact => contact.isMainContact);
    
    return [
      formatArabicText(client.name),
      clientTypeLabels[client.type] || 'غير معروف',
      client.phone || '-',
      client.email || '-',
      formatArabicText(client.address || '-'),
      formatArabicText(mainContact?.name || '-'),
      mainContact?.phone || '-',
      formatDate(client.createdAt)
    ];
  });

  // إعداد الجدول
  (doc as any).autoTable({
    head: [['اسم العميل', 'النوع', 'الهاتف', 'البريد الإلكتروني', 'العنوان', 'جهة الاتصال الرئيسية', 'هاتف جهة الاتصال', 'تاريخ الإضافة']],
    body: tableData,
    startY: 45,
    styles: {
      fontSize: 8,
      cellPadding: 3,
      overflow: 'linebreak',
      halign: 'right',
      font: 'helvetica'
    },
    headStyles: {
      fillColor: [66, 139, 202],
      textColor: 255,
      fontSize: 9,
      fontStyle: 'bold',
      halign: 'right'
    },
    alternateRowStyles: {
      fillColor: [245, 245, 245]
    },
    margin: { top: 45, right: 10, bottom: 20, left: 10 },
    columnStyles: {
      0: { cellWidth: 35 },
      1: { cellWidth: 25 },
      2: { cellWidth: 25 },
      3: { cellWidth: 35 },
      4: { cellWidth: 35 },
      5: { cellWidth: 30 },
      6: { cellWidth: 25 },
      7: { cellWidth: 25 }
    }
  });

  // إضافة إحصائيات
  const finalY = (doc as any).lastAutoTable.finalY + 20;
  doc.setFontSize(12);
  doc.text('ملخص الإحصائيات:', 20, finalY);
  doc.setFontSize(10);
  doc.text(`إجمالي العملاء: ${clients.length}`, 20, finalY + 10);
  
  const ownerCount = clients.filter(c => c.type === 'owner').length;
  const contractorCount = clients.filter(c => c.type === 'main_contractor').length;
  const consultantCount = clients.filter(c => c.type === 'consultant').length;
  
  doc.text(`الملاك: ${ownerCount}`, 20, finalY + 20);
  doc.text(`المقاولون الرئيسيون: ${contractorCount}`, 20, finalY + 30);
  doc.text(`الاستشاريون: ${consultantCount}`, 20, finalY + 40);

  // حفظ الملف
  const fileName = `تقرير-العملاء-${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
};

// تصدير العملاء إلى Word
export const exportClientsToWord = (clients: Client[], title: string = 'تقرير العملاء') => {
  const ownerCount = clients.filter(c => c.type === 'owner').length;
  const contractorCount = clients.filter(c => c.type === 'main_contractor').length;
  const consultantCount = clients.filter(c => c.type === 'consultant').length;

  const content = [
    {
      type: 'summary',
      title: 'ملخص الإحصائيات',
      items: [
        { label: 'إجمالي العملاء', value: clients.length.toString() },
        { label: 'الملاك', value: ownerCount.toString() },
        { label: 'المقاولون الرئيسيون', value: contractorCount.toString() },
        { label: 'الاستشاريون', value: consultantCount.toString() }
      ]
    },
    {
      type: 'table',
      title: 'تفاصيل العملاء',
      headers: ['اسم العميل', 'النوع', 'الهاتف', 'البريد الإلكتروني', 'العنوان', 'جهة الاتصال الرئيسية', 'هاتف جهة الاتصال', 'تاريخ الإضافة'],
      rows: clients.map(client => {
        const clientTypeLabels = {
          'owner': 'مالك',
          'main_contractor': 'مقاول رئيسي',
          'consultant': 'استشاري'
        };
        
        const mainContact = client.contacts?.find(contact => contact.isMainContact);
        
        return [
          client.name,
          clientTypeLabels[client.type] || 'غير معروف',
          client.phone || '-',
          client.email || '-',
          client.address || '-',
          mainContact?.name || '-',
          mainContact?.phone || '-',
          formatDate(client.createdAt)
        ];
      })
    }
  ];

  const blob = createWordDocument(title, content);
  const fileName = `تقرير-العملاء-${new Date().toISOString().split('T')[0]}.doc`;
  saveWordDocument(blob, fileName);
};

// تصدير العملاء إلى Excel
export const exportClientsToExcel = (clients: Client[], title: string = 'تقرير العملاء') => {
  const data = clients.map(client => {
    const clientTypeLabels = {
      'owner': 'مالك',
      'main_contractor': 'مقاول رئيسي',
      'consultant': 'استشاري'
    };
    
    const mainContact = client.contacts?.find(contact => contact.isMainContact);
    
    return {
      'اسم العميل': client.name,
      'نوع العميل': clientTypeLabels[client.type] || 'غير معروف',
      'رقم الهاتف': client.phone || '',
      'البريد الإلكتروني': client.email || '',
      'العنوان': client.address || '',
      'جهة الاتصال الرئيسية': mainContact?.name || '',
      'منصب جهة الاتصال': mainContact?.position || '',
      'قسم جهة الاتصال': mainContact?.department || '',
      'هاتف جهة الاتصال': mainContact?.phone || '',
      'بريد جهة الاتصال': mainContact?.email || '',
      'عدد جهات الاتصال': client.contacts?.length || 0,
      'تاريخ الإضافة': formatDate(client.createdAt),
      'تاريخ التحديث': formatDate(client.updatedAt)
    };
  });

  const worksheet = XLSX.utils.json_to_sheet(data);
  worksheet['!cols'] = [
    { wch: 25 }, // اسم العميل
    { wch: 15 }, // نوع العميل
    { wch: 15 }, // رقم الهاتف
    { wch: 30 }, // البريد الإلكتروني
    { wch: 30 }, // العنوان
    { wch: 25 }, // جهة الاتصال الرئيسية
    { wch: 20 }, // منصب جهة الاتصال
    { wch: 15 }, // قسم جهة الاتصال
    { wch: 15 }, // هاتف جهة الاتصال
    { wch: 25 }, // بريد جهة الاتصال
    { wch: 12 }, // عدد جهات الاتصال
    { wch: 15 }, // تاريخ الإضافة
    { wch: 15 }  // تاريخ التحديث
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'العملاء');

  const fileName = `تقرير-العملاء-${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};

// تصدير تقرير مالي شامل
export const exportFinancialReportToPDF = (operations: Operation[], clients: Client[]) => {
  const doc = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4'
  });
  
  setupArabicPDFFont(doc);

  // العنوان
  doc.setFontSize(18);
  const titleText = 'التقرير المالي الشامل';
  doc.text(formatArabicText(titleText), doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  
  // تاريخ التقرير
  doc.setFontSize(10);
  const dateText = `تاريخ التقرير: ${formatDate(new Date())}`;
  doc.text(dateText, 20, 35);

  // حساب الإحصائيات
  const totalOperations = operations.length;
  const completedOperations = operations.filter(op => op.status === 'completed').length;
  const inProgressOperations = operations.filter(op => op.status === 'in_progress').length;
  const totalAmount = operations.reduce((sum, op) => sum + op.totalAmount, 0);
  const totalDeductions = operations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
  const totalNetAmount = operations.reduce((sum, op) => sum + calculateNetAmount(op), 0);
  const totalReceived = operations.reduce((sum, op) => sum + op.totalReceived, 0);
  const totalOutstanding = totalNetAmount - totalReceived;
  const collectionRate = totalNetAmount > 0 ? (totalReceived / totalNetAmount) * 100 : 0;

  // الإحصائيات العامة
  let yPosition = 50;
  doc.setFontSize(14);
  doc.text('الإحصائيات العامة:', 20, yPosition);
  
  doc.setFontSize(11);
  yPosition += 15;
  doc.text(`إجمالي العمليات: ${totalOperations}`, 30, yPosition);
  yPosition += 10;
  doc.text(`العمليات المكتملة: ${completedOperations}`, 30, yPosition);
  yPosition += 10;
  doc.text(`العمليات قيد التنفيذ: ${inProgressOperations}`, 30, yPosition);
  yPosition += 10;
  doc.text(`معدل الإنجاز: ${totalOperations > 0 ? ((completedOperations / totalOperations) * 100).toFixed(1) : 0}%`, 30, yPosition);

  // الإحصائيات المالية
  yPosition += 20;
  doc.setFontSize(14);
  doc.text('الإحصائيات المالية:', 20, yPosition);
  
  doc.setFontSize(11);
  yPosition += 15;
  doc.text(`إجمالي القيمة: ${formatCurrency(totalAmount)}`, 30, yPosition);
  yPosition += 10;
  doc.text(`إجمالي الخصومات: ${formatCurrency(totalDeductions)}`, 30, yPosition);
  yPosition += 10;
  doc.text(`الصافي المستحق: ${formatCurrency(totalNetAmount)}`, 30, yPosition);
  yPosition += 10;
  doc.text(`إجمالي المحصل: ${formatCurrency(totalReceived)}`, 30, yPosition);
  yPosition += 10;
  doc.text(`إجمالي المتبقي: ${formatCurrency(totalOutstanding)}`, 30, yPosition);
  yPosition += 10;
  doc.text(`معدل التحصيل: ${collectionRate.toFixed(1)}%`, 30, yPosition);

  // جدول العمليات حسب العميل
  yPosition += 30;
  const clientStats = clients.map(client => {
    const clientOperations = operations.filter(op => op.clientId === client.id);
    const clientTotal = clientOperations.reduce((sum, op) => sum + op.totalAmount, 0);
    const clientDeductions = clientOperations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
    const clientNetAmount = clientOperations.reduce((sum, op) => sum + calculateNetAmount(op), 0);
    const clientReceived = clientOperations.reduce((sum, op) => sum + op.totalReceived, 0);
    
    return [
      formatArabicText(client.name),
      clientOperations.length.toString(),
      formatCurrency(clientTotal),
      formatCurrency(clientDeductions),
      formatCurrency(clientNetAmount),
      formatCurrency(clientReceived),
      formatCurrency(clientNetAmount - clientReceived),
      clientNetAmount > 0 ? `${((clientReceived / clientNetAmount) * 100).toFixed(1)}%` : '0%'
    ];
  }).filter(stat => parseInt(stat[1]) > 0);

  if (clientStats.length > 0) {
    (doc as any).autoTable({
      head: [['العميل', 'العمليات', 'إجمالي القيمة', 'الخصومات', 'الصافي المستحق', 'المحصل', 'المتبقي', 'معدل التحصيل']],
      body: clientStats,
      startY: yPosition,
      styles: {
        fontSize: 8,
        cellPadding: 3,
        halign: 'right',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [66, 139, 202],
        textColor: 255,
        fontSize: 9,
        fontStyle: 'bold',
        halign: 'right'
      },
      alternateRowStyles: {
        fillColor: [245, 245, 245]
      }
    });
  }

  const fileName = `التقرير-المالي-${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
};

// تصدير تقرير مالي إلى Word
export const exportFinancialReportToWord = (operations: Operation[], clients: Client[]) => {
  // ورقة الإحصائيات العامة
  const totalAmount = operations.reduce((sum, op) => sum + op.totalAmount, 0);
  const totalDeductions = operations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
  const totalNetAmount = operations.reduce((sum, op) => sum + calculateNetAmount(op), 0);
  const totalReceived = operations.reduce((sum, op) => sum + op.totalReceived, 0);
  const completedOperations = operations.filter(op => op.status === 'completed').length;

  const content = [
    {
      type: 'summary',
      title: 'الإحصائيات العامة',
      items: [
        { label: 'إجمالي العمليات', value: operations.length.toString() },
        { label: 'العمليات المكتملة', value: completedOperations.toString() },
        { label: 'العمليات قيد التنفيذ', value: operations.filter(op => op.status === 'in_progress').length.toString() },
        { label: 'معدل الإنجاز', value: `${operations.length > 0 ? ((completedOperations / operations.length) * 100).toFixed(1) : 0}%` }
      ]
    },
    {
      type: 'summary',
      title: 'الإحصائيات المالية',
      items: [
        { label: 'إجمالي القيمة', value: `<span class="currency">${formatCurrency(totalAmount)}</span>` },
        { label: 'إجمالي الخصومات', value: `<span class="currency">${formatCurrency(totalDeductions)}</span>` },
        { label: 'الصافي المستحق', value: `<span class="currency">${formatCurrency(totalNetAmount)}</span>` },
        { label: 'إجمالي المحصل', value: `<span class="currency">${formatCurrency(totalReceived)}</span>` },
        { label: 'إجمالي المتبقي', value: `<span class="currency">${formatCurrency(totalNetAmount - totalReceived)}</span>` },
        { label: 'معدل التحصيل', value: `${totalNetAmount > 0 ? ((totalReceived / totalNetAmount) * 100).toFixed(1) : 0}%` }
      ]
    }
  ];

  // ورقة إحصائيات العملاء
  const clientStats = clients.map(client => {
    const clientOperations = operations.filter(op => op.clientId === client.id);
    const clientTotal = clientOperations.reduce((sum, op) => sum + op.totalAmount, 0);
    const clientDeductions = clientOperations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
    const clientNetAmount = clientOperations.reduce((sum, op) => sum + calculateNetAmount(op), 0);
    const clientReceived = clientOperations.reduce((sum, op) => sum + op.totalReceived, 0);
    
    return [
      client.name,
      client.type === 'owner' ? 'مالك' : client.type === 'main_contractor' ? 'مقاول رئيسي' : 'استشاري',
      clientOperations.length.toString(),
      `<span class="currency">${formatCurrency(clientTotal)}</span>`,
      `<span class="currency">${formatCurrency(clientDeductions)}</span>`,
      `<span class="currency">${formatCurrency(clientNetAmount)}</span>`,
      `<span class="currency">${formatCurrency(clientReceived)}</span>`,
      `<span class="currency">${formatCurrency(clientNetAmount - clientReceived)}</span>`,
      clientNetAmount > 0 ? `${((clientReceived / clientNetAmount) * 100).toFixed(1)}%` : '0%'
    ];
  }).filter(stat => parseInt(stat[2]) > 0);

  if (clientStats.length > 0) {
    content.push({
      type: 'table',
      title: 'إحصائيات العملاء',
      headers: ['العميل', 'نوع العميل', 'عدد العمليات', 'إجمالي القيمة', 'إجمالي الخصومات', 'الصافي المستحق', 'المبلغ المحصل', 'المبلغ المتبقي', 'معدل التحصيل'],
      rows: clientStats
    });
  }

  const blob = createWordDocument('التقرير المالي الشامل', content);
  const fileName = `التقرير-المالي-${new Date().toISOString().split('T')[0]}.doc`;
  saveWordDocument(blob, fileName);
};

// تصدير تقرير مالي إلى Excel
export const exportFinancialReportToExcel = (operations: Operation[], clients: Client[]) => {
  // ورقة الإحصائيات العامة
  const totalAmount = operations.reduce((sum, op) => sum + op.totalAmount, 0);
  const totalDeductions = operations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
  const totalNetAmount = operations.reduce((sum, op) => sum + calculateNetAmount(op), 0);
  const totalReceived = operations.reduce((sum, op) => sum + op.totalReceived, 0);

  const generalStats = [
    { 'البند': 'إجمالي العمليات', 'القيمة': operations.length },
    { 'البند': 'العمليات المكتملة', 'القيمة': operations.filter(op => op.status === 'completed').length },
    { 'البند': 'العمليات قيد التنفيذ', 'القيمة': operations.filter(op => op.status === 'in_progress').length },
    { 'البند': 'إجمالي القيمة', 'القيمة': totalAmount },
    { 'البند': 'إجمالي الخصومات', 'القيمة': totalDeductions },
    { 'البند': 'الصافي المستحق', 'القيمة': totalNetAmount },
    { 'البند': 'إجمالي المحصل', 'القيمة': totalReceived },
    { 'البند': 'إجمالي المتبقي', 'القيمة': totalNetAmount - totalReceived }
  ];

  // ورقة إحصائيات العملاء
  const clientStats = clients.map(client => {
    const clientOperations = operations.filter(op => op.clientId === client.id);
    const clientTotal = clientOperations.reduce((sum, op) => sum + op.totalAmount, 0);
    const clientDeductions = clientOperations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
    const clientNetAmount = clientOperations.reduce((sum, op) => sum + calculateNetAmount(op), 0);
    const clientReceived = clientOperations.reduce((sum, op) => sum + op.totalReceived, 0);
    
    return {
      'العميل': client.name,
      'نوع العميل': client.type === 'owner' ? 'مالك' : client.type === 'main_contractor' ? 'مقاول رئيسي' : 'استشاري',
      'عدد العمليات': clientOperations.length,
      'إجمالي القيمة': clientTotal,
      'إجمالي الخصومات': clientDeductions,
      'الصافي المستحق': clientNetAmount,
      'المبلغ المحصل': clientReceived,
      'المبلغ المتبقي': clientNetAmount - clientReceived,
      'معدل التحصيل': clientNetAmount > 0 ? `${((clientReceived / clientNetAmount) * 100).toFixed(1)}%` : '0%'
    };
  }).filter(stat => stat['عدد العمليات'] > 0);

  // إنشاء المصنف
  const workbook = XLSX.utils.book_new();

  // إضافة ورقة الإحصائيات العامة
  const statsWorksheet = XLSX.utils.json_to_sheet(generalStats);
  statsWorksheet['!cols'] = [{ wch: 25 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(workbook, statsWorksheet, 'الإحصائيات العامة');

  // إضافة ورقة إحصائيات العملاء
  if (clientStats.length > 0) {
    const clientWorksheet = XLSX.utils.json_to_sheet(clientStats);
    clientWorksheet['!cols'] = [
      { wch: 25 }, // العميل
      { wch: 15 }, // نوع العميل
      { wch: 12 }, // عدد العمليات
      { wch: 15 }, // إجمالي القيمة
      { wch: 15 }, // إجمالي الخصومات
      { wch: 15 }, // الصافي المستحق
      { wch: 15 }, // المبلغ المحصل
      { wch: 15 }, // المبلغ المتبقي
      { wch: 15 }  // معدل التحصيل
    ];
    XLSX.utils.book_append_sheet(workbook, clientWorksheet, 'إحصائيات العملاء');
  }

  const fileName = `التقرير-المالي-${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};