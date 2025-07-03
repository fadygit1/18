import jsPDF from 'jspdf';
import 'jspdf-autotable';
import * as XLSX from 'xlsx';
import { Operation, Client } from '../types';
import { formatCurrency, formatDate, calculateNetAmount, calculateTotalDeductions, calculateExecutedTotal } from './calculations';

// إعداد الخط الإنجليزي لـ PDF
const setupPDFFont = (doc: jsPDF) => {
  doc.setFont('helvetica');
  doc.setFontSize(12);
};

// دالة لتحويل النص العربي للإنجليزية
const convertArabicToEnglish = (text: string): string => {
  if (!text) return '';
  
  // قاموس الترجمة للكلمات الشائعة
  const translations: { [key: string]: string } = {
    'شركة الإنشاءات الحديثة': 'Modern Construction Company',
    'مؤسسة البناء المتطور': 'Advanced Building Foundation',
    'شركة المقاولات العربية': 'Arab Contracting Company',
    'مالك': 'Owner',
    'مقاول رئيسي': 'Main Contractor',
    'استشاري': 'Consultant',
    'قيد التنفيذ': 'In Progress',
    'مكتملة': 'Completed',
    'مكتملة - دفع جزئي': 'Completed - Partial Payment',
    'نشط': 'Active',
    'منتهي': 'Expired',
    'قائم': 'Active',
    'مُسترد': 'Returned',
    'نقدي': 'Cash',
    'شيك': 'Check'
  };

  // البحث عن ترجمة مباشرة
  if (translations[text]) {
    return translations[text];
  }

  // تحويل الأرقام العربية للإنجليزية
  const arabicNumbers = ['٠', '١', '٢', '٣', '٤', '٥', '٦', '٧', '٨', '٩'];
  const englishNumbers = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'];
  
  let convertedText = text;
  arabicNumbers.forEach((arabicNum, index) => {
    convertedText = convertedText.replace(new RegExp(arabicNum, 'g'), englishNumbers[index]);
  });

  // إزالة الأحرف العربية وتحويلها لنص إنجليزي عام
  convertedText = convertedText.replace(/[\u0600-\u06FF]/g, '');
  
  // إذا كان النص فارغاً بعد التحويل، استخدم نص بديل
  if (!convertedText.trim()) {
    return 'Arabic Text';
  }

  return convertedText.trim();
};

// تصدير تفاصيل العملية إلى PDF
export const exportOperationDetailsToPDF = (operation: Operation, client: Client) => {
  const doc = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4'
  });
  
  setupPDFFont(doc);

  // العنوان
  doc.setFontSize(18);
  const titleText = 'Operation Details - ' + convertArabicToEnglish(operation.name);
  doc.text(titleText, doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  
  // معلومات العملية
  doc.setFontSize(12);
  let yPosition = 40;
  
  doc.text('Operation Name: ' + convertArabicToEnglish(operation.name), 20, yPosition);
  yPosition += 10;
  doc.text('Operation Code: ' + operation.code, 20, yPosition);
  yPosition += 10;
  doc.text('Client: ' + convertArabicToEnglish(client.name), 20, yPosition);
  yPosition += 10;
  doc.text('Client Type: ' + convertArabicToEnglish(client.type === 'owner' ? 'Owner' : client.type === 'main_contractor' ? 'Main Contractor' : 'Consultant'), 20, yPosition);
  yPosition += 10;
  doc.text('Created Date: ' + formatDate(operation.createdAt), 20, yPosition);
  yPosition += 15;

  // الملخص المالي
  const executedAmount = calculateExecutedTotal(operation.items);
  const totalDeductions = calculateTotalDeductions(operation);
  const netAmount = calculateNetAmount(operation);
  const remainingAmount = netAmount - operation.totalReceived;

  doc.setFontSize(14);
  doc.text('Financial Summary:', 20, yPosition);
  yPosition += 10;
  
  doc.setFontSize(11);
  doc.text('Total Amount: ' + formatCurrency(operation.totalAmount), 30, yPosition);
  yPosition += 8;
  doc.text('Executed Amount: ' + formatCurrency(executedAmount), 30, yPosition);
  yPosition += 8;
  doc.text('Total Deductions: ' + formatCurrency(totalDeductions), 30, yPosition);
  yPosition += 8;
  doc.text('Net Due: ' + formatCurrency(netAmount), 30, yPosition);
  yPosition += 8;
  doc.text('Received Amount: ' + formatCurrency(operation.totalReceived), 30, yPosition);
  yPosition += 8;
  doc.text('Remaining Amount: ' + formatCurrency(remainingAmount), 30, yPosition);
  yPosition += 8;
  doc.text('Completion Rate: ' + operation.overallExecutionPercentage.toFixed(1) + '%', 30, yPosition);
  yPosition += 15;

  // بنود العملية
  if (operation.items.length > 0) {
    doc.setFontSize(14);
    doc.text('Operation Items:', 20, yPosition);
    yPosition += 10;

    const itemsData = operation.items.map(item => [
      item.code,
      convertArabicToEnglish(item.description),
      formatCurrency(item.amount),
      item.executionPercentage + '%',
      formatCurrency(item.amount * (item.executionPercentage / 100))
    ]);

    (doc as any).autoTable({
      head: [['Code', 'Description', 'Amount', 'Completion %', 'Executed Amount']],
      body: itemsData,
      startY: yPosition,
      styles: {
        fontSize: 9,
        cellPadding: 3,
        halign: 'center',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [66, 139, 202],
        textColor: 255,
        fontSize: 10,
        fontStyle: 'bold'
      },
      alternateRowStyles: {
        fillColor: [245, 245, 245]
      }
    });

    yPosition = (doc as any).lastAutoTable.finalY + 15;
  }

  // الخصومات
  if (operation.deductions.length > 0 && totalDeductions > 0) {
    if (yPosition > 250) {
      doc.addPage();
      yPosition = 30;
    }

    doc.setFontSize(14);
    doc.text('Deductions:', 20, yPosition);
    yPosition += 10;

    const deductionsData = operation.deductions.filter(d => d.isActive).map(deduction => {
      const deductionAmount = deduction.type === 'percentage' 
        ? (executedAmount * deduction.value / 100)
        : deduction.value;
      
      return [
        convertArabicToEnglish(deduction.name),
        deduction.type === 'percentage' ? `${deduction.value}%` : 'Fixed Amount',
        formatCurrency(deductionAmount)
      ];
    });

    (doc as any).autoTable({
      head: [['Deduction Name', 'Type', 'Amount']],
      body: deductionsData,
      startY: yPosition,
      styles: {
        fontSize: 9,
        cellPadding: 3,
        halign: 'center',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [220, 53, 69],
        textColor: 255,
        fontSize: 10,
        fontStyle: 'bold'
      }
    });

    yPosition = (doc as any).lastAutoTable.finalY + 15;
  }

  // إضافة صفحة جديدة إذا لزم الأمر
  if (yPosition > 250) {
    doc.addPage();
    yPosition = 30;
  }

  // شيكات الضمان
  if (operation.guaranteeChecks.length > 0) {
    doc.setFontSize(14);
    doc.text('Guarantee Checks:', 20, yPosition);
    yPosition += 10;

    const checksData = operation.guaranteeChecks.map(check => [
      check.checkNumber,
      formatCurrency(check.amount),
      convertArabicToEnglish(check.bank),
      formatDate(check.expiryDate),
      check.isReturned ? 'Returned' : 'Active'
    ]);

    (doc as any).autoTable({
      head: [['Check Number', 'Amount', 'Bank', 'Expiry Date', 'Status']],
      body: checksData,
      startY: yPosition,
      styles: {
        fontSize: 9,
        cellPadding: 3,
        halign: 'center',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [139, 69, 19],
        textColor: 255,
        fontSize: 10,
        fontStyle: 'bold'
      }
    });

    yPosition = (doc as any).lastAutoTable.finalY + 15;
  }

  // خطابات الضمان
  if (operation.guaranteeLetters.length > 0) {
    if (yPosition > 200) {
      doc.addPage();
      yPosition = 30;
    }

    doc.setFontSize(14);
    doc.text('Guarantee Letters:', 20, yPosition);
    yPosition += 10;

    const lettersData = operation.guaranteeLetters.map(letter => [
      letter.letterNumber,
      convertArabicToEnglish(letter.bank),
      formatCurrency(letter.amount),
      formatDate(letter.dueDate),
      letter.isReturned ? 'Returned' : 'Active'
    ]);

    (doc as any).autoTable({
      head: [['Letter Number', 'Bank', 'Amount', 'Due Date', 'Status']],
      body: lettersData,
      startY: yPosition,
      styles: {
        fontSize: 9,
        cellPadding: 3,
        halign: 'center',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [34, 139, 34],
        textColor: 255,
        fontSize: 10,
        fontStyle: 'bold'
      }
    });

    yPosition = (doc as any).lastAutoTable.finalY + 15;
  }

  // شهادات الضمان
  if ((operation.warrantyCertificates || []).length > 0) {
    if (yPosition > 200) {
      doc.addPage();
      yPosition = 30;
    }

    doc.setFontSize(14);
    doc.text('Warranty Certificates:', 20, yPosition);
    yPosition += 10;

    const warrantiesData = operation.warrantyCertificates!.map(warranty => [
      warranty.certificateNumber,
      convertArabicToEnglish(warranty.description),
      formatDate(warranty.startDate),
      formatDate(warranty.endDate),
      warranty.isActive ? 'Active' : 'Expired'
    ]);

    (doc as any).autoTable({
      head: [['Certificate Number', 'Description', 'Start Date', 'End Date', 'Status']],
      body: warrantiesData,
      startY: yPosition,
      styles: {
        fontSize: 9,
        cellPadding: 3,
        halign: 'center',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [128, 0, 128],
        textColor: 255,
        fontSize: 10,
        fontStyle: 'bold'
      }
    });

    yPosition = (doc as any).lastAutoTable.finalY + 15;
  }

  // المدفوعات المستلمة
  if (operation.receivedPayments.length > 0) {
    if (yPosition > 200) {
      doc.addPage();
      yPosition = 30;
    }

    doc.setFontSize(14);
    doc.text('Received Payments:', 20, yPosition);
    yPosition += 10;

    const paymentsData = operation.receivedPayments.map(payment => [
      payment.type === 'cash' ? 'Cash' : 'Check',
      formatCurrency(payment.amount),
      formatDate(payment.date),
      payment.type === 'check' && payment.checkNumber 
        ? 'Check No: ' + payment.checkNumber + ' - ' + convertArabicToEnglish(payment.bank || '')
        : convertArabicToEnglish(payment.notes || '-')
    ]);

    (doc as any).autoTable({
      head: [['Type', 'Amount', 'Date', 'Details']],
      body: paymentsData,
      startY: yPosition,
      styles: {
        fontSize: 9,
        cellPadding: 3,
        halign: 'center',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [0, 128, 0],
        textColor: 255,
        fontSize: 10,
        fontStyle: 'bold'
      }
    });
  }

  const fileName = `operation-details-${operation.code}-${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
};

// تصدير العمليات إلى PDF
export const exportOperationsToPDF = (operations: Operation[], clients: Client[], title: string = 'Operations Report') => {
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4'
  });
  
  setupPDFFont(doc);

  // العنوان
  doc.setFontSize(16);
  doc.text(title, doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  
  // تاريخ التقرير
  doc.setFontSize(10);
  const dateText = 'Report Date: ' + formatDate(new Date());
  doc.text(dateText, 20, 35);

  // إعداد البيانات للجدول
  const tableData = operations.map(operation => {
    const client = clients.find(c => c.id === operation.clientId);
    const statusLabels = {
      'in_progress': 'In Progress',
      'completed': 'Completed',
      'completed_partial_payment': 'Completed - Partial Payment'
    };

    const totalDeductions = calculateTotalDeductions(operation);
    const netAmount = calculateNetAmount(operation);

    return [
      operation.code,
      convertArabicToEnglish(operation.name),
      convertArabicToEnglish(client?.name || 'Unknown'),
      formatCurrency(operation.totalAmount),
      formatCurrency(totalDeductions),
      formatCurrency(netAmount),
      formatCurrency(operation.totalReceived),
      operation.overallExecutionPercentage.toFixed(1) + '%',
      statusLabels[operation.status],
      formatDate(operation.createdAt)
    ];
  });

  // إعداد الجدول
  (doc as any).autoTable({
    head: [['Code', 'Name', 'Client', 'Total Amount', 'Deductions', 'Net Due', 'Received', 'Completion %', 'Status', 'Created Date']],
    body: tableData,
    startY: 45,
    styles: {
      fontSize: 7,
      cellPadding: 2,
      overflow: 'linebreak',
      halign: 'center',
      font: 'helvetica'
    },
    headStyles: {
      fillColor: [66, 139, 202],
      textColor: 255,
      fontSize: 8,
      fontStyle: 'bold'
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
  doc.text('Statistics Summary:', 20, finalY);
  doc.setFontSize(10);
  doc.text('Total Operations: ' + operations.length, 20, finalY + 10);
  doc.text('Completed Operations: ' + completedCount, 20, finalY + 20);
  doc.text('Total Amount: ' + formatCurrency(totalAmount), 20, finalY + 30);
  doc.text('Total Deductions: ' + formatCurrency(totalDeductions), 20, finalY + 40);
  doc.text('Total Net Due: ' + formatCurrency(totalNetAmount), 20, finalY + 50);
  doc.text('Total Received: ' + formatCurrency(totalReceived), 20, finalY + 60);

  const fileName = `operations-report-${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
};

// تصدير الشيكات والمدفوعات إلى PDF
export const exportChecksAndPaymentsToPDF = (operations: Operation[], clients: Client[]) => {
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4'
  });
  
  setupPDFFont(doc);

  // العنوان
  doc.setFontSize(16);
  const titleText = 'Checks and Payments Report';
  doc.text(titleText, doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  
  // تاريخ التقرير
  doc.setFontSize(10);
  const dateText = 'Report Date: ' + formatDate(new Date());
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
        clientName: client?.name || 'Unknown'
      };
    })
  );

  if (allPayments.length > 0) {
    doc.setFontSize(14);
    doc.text('Received Payments:', 20, yPosition);
    yPosition += 10;

    const paymentsData = allPayments.map(payment => [
      payment.type === 'cash' ? 'Cash' : 'Check',
      formatCurrency(payment.amount),
      formatDate(payment.date),
      convertArabicToEnglish(payment.clientName),
      convertArabicToEnglish(payment.operationName),
      payment.type === 'check' && payment.checkNumber 
        ? payment.checkNumber + ' - ' + convertArabicToEnglish(payment.bank || '')
        : convertArabicToEnglish(payment.notes || '-')
    ]);

    (doc as any).autoTable({
      head: [['Type', 'Amount', 'Date', 'Client', 'Operation', 'Details']],
      body: paymentsData,
      startY: yPosition,
      styles: {
        fontSize: 8,
        cellPadding: 3,
        halign: 'center',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [0, 128, 0],
        textColor: 255,
        fontSize: 9,
        fontStyle: 'bold'
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
  doc.text('Statistics Summary:', 20, yPosition);
  doc.setFontSize(10);
  doc.text('Total Payments: ' + formatCurrency(totalAmount), 20, yPosition + 10);
  doc.text('Number of Checks: ' + totalChecks, 20, yPosition + 20);
  doc.text('Cash Payments: ' + totalCash, 20, yPosition + 30);

  const fileName = `checks-payments-report-${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
};

// تصدير الشيكات والمدفوعات إلى Excel
export const exportChecksAndPaymentsToExcel = (operations: Operation[], clients: Client[]) => {
  // جمع جميع المدفوعات
  const allPayments = operations.flatMap(operation => 
    operation.receivedPayments.map(payment => {
      const client = clients.find(c => c.id === operation.clientId);
      return {
        'Payment Type': payment.type === 'cash' ? 'Cash' : 'Check',
        'Amount': payment.amount,
        'Date': formatDate(payment.date),
        'Client': client?.name || 'Unknown',
        'Operation': operation.name,
        'Operation Code': operation.code,
        'Check Number': payment.checkNumber || '',
        'Bank': payment.bank || '',
        'Receipt Date': payment.receiptDate ? formatDate(payment.receiptDate) : '',
        'Notes': payment.notes || ''
      };
    })
  );

  const worksheet = XLSX.utils.json_to_sheet(allPayments);
  worksheet['!cols'] = [
    { wch: 12 }, // Payment Type
    { wch: 15 }, // Amount
    { wch: 12 }, // Date
    { wch: 25 }, // Client
    { wch: 30 }, // Operation
    { wch: 15 }, // Operation Code
    { wch: 15 }, // Check Number
    { wch: 20 }, // Bank
    { wch: 15 }, // Receipt Date
    { wch: 30 }  // Notes
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Checks and Payments');

  // إضافة ورقة الإحصائيات
  const stats = [
    { 'Item': 'Total Payments', 'Value': allPayments.reduce((sum, p) => sum + p.Amount, 0) },
    { 'Item': 'Number of Checks', 'Value': allPayments.filter(p => p['Payment Type'] === 'Check').length },
    { 'Item': 'Cash Payments', 'Value': allPayments.filter(p => p['Payment Type'] === 'Cash').length }
  ];

  const statsWorksheet = XLSX.utils.json_to_sheet(stats);
  statsWorksheet['!cols'] = [{ wch: 20 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(workbook, statsWorksheet, 'Statistics');

  const fileName = `checks-payments-report-${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};

// تصدير تقرير مفصل للضمانات
export const exportDetailedGuaranteesReportToPDF = (operations: Operation[], clients: Client[]) => {
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4'
  });
  
  setupPDFFont(doc);

  // العنوان
  doc.setFontSize(16);
  const titleText = 'Detailed Guarantees Report';
  doc.text(titleText, doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  
  // تاريخ التقرير
  doc.setFontSize(10);
  const dateText = 'Report Date: ' + formatDate(new Date());
  doc.text(dateText, 20, 35);

  let yPosition = 50;

  // تقرير شيكات الضمان
  doc.setFontSize(14);
  doc.text('Guarantee Checks:', 20, yPosition);
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
        clientName: client?.name || 'Unknown',
        relatedItemDescription: relatedItem?.description || 'Full Operation'
      };
    })
  );

  if (allGuaranteeChecks.length > 0) {
    const checksTableData = allGuaranteeChecks.map(check => [
      check.checkNumber,
      convertArabicToEnglish(check.clientName),
      convertArabicToEnglish(check.operationName),
      convertArabicToEnglish(check.relatedItemDescription),
      formatCurrency(check.amount),
      formatDate(check.checkDate),
      formatDate(check.expiryDate),
      convertArabicToEnglish(check.bank),
      check.isReturned ? 'Returned' : 'Active'
    ]);

    (doc as any).autoTable({
      head: [['Check Number', 'Client', 'Operation', 'Item', 'Amount', 'Issue Date', 'Expiry Date', 'Bank', 'Status']],
      body: checksTableData,
      startY: yPosition,
      styles: {
        fontSize: 7,
        cellPadding: 2,
        overflow: 'linebreak',
        halign: 'center',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [66, 139, 202],
        textColor: 255,
        fontSize: 8,
        fontStyle: 'bold'
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
  doc.text('Guarantee Letters:', 20, yPosition);
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
        clientName: client?.name || 'Unknown',
        relatedItemDescription: relatedItem?.description || 'Full Operation'
      };
    })
  );

  if (allGuaranteeLetters.length > 0) {
    const lettersTableData = allGuaranteeLetters.map(letter => [
      letter.letterNumber,
      convertArabicToEnglish(letter.clientName),
      convertArabicToEnglish(letter.operationName),
      convertArabicToEnglish(letter.relatedItemDescription),
      formatCurrency(letter.amount),
      formatDate(letter.letterDate),
      formatDate(letter.dueDate),
      convertArabicToEnglish(letter.bank),
      letter.isReturned ? 'Returned' : 'Active'
    ]);

    (doc as any).autoTable({
      head: [['Letter Number', 'Client', 'Operation', 'Item', 'Amount', 'Issue Date', 'Due Date', 'Bank', 'Status']],
      body: lettersTableData,
      startY: yPosition,
      styles: {
        fontSize: 7,
        cellPadding: 2,
        overflow: 'linebreak',
        halign: 'center',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [139, 69, 19],
        textColor: 255,
        fontSize: 8,
        fontStyle: 'bold'
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

  const fileName = `detailed-guarantees-report-${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
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
        'Type': 'Guarantee Check',
        'Number': check.checkNumber,
        'Amount': check.amount,
        'Bank': check.bank,
        'Client': client?.name || 'Unknown',
        'Operation': operation.name,
        'Operation Code': operation.code,
        'Related Item': relatedItem?.description || 'Full Operation',
        'Issue Date': formatDate(check.checkDate),
        'Delivery Date': formatDate(check.deliveryDate),
        'Expiry Date': formatDate(check.expiryDate),
        'Status': check.isReturned ? 'Returned' : 'Active',
        'Return Date': check.returnDate ? formatDate(check.returnDate) : ''
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
        'Type': 'Guarantee Letter',
        'Number': letter.letterNumber,
        'Amount': letter.amount,
        'Bank': letter.bank,
        'Client': client?.name || 'Unknown',
        'Operation': operation.name,
        'Operation Code': operation.code,
        'Related Item': relatedItem?.description || 'Full Operation',
        'Issue Date': formatDate(letter.letterDate),
        'Delivery Date': '',
        'Expiry Date': formatDate(letter.dueDate),
        'Status': letter.isReturned ? 'Returned' : 'Active',
        'Return Date': letter.returnDate ? formatDate(letter.returnDate) : '',
        'Notes': letter.notes || ''
      };
    })
  );

  // دمج جميع الضمانات
  const allGuarantees = [...allGuaranteeChecks, ...allGuaranteeLetters];

  const worksheet = XLSX.utils.json_to_sheet(allGuarantees);
  worksheet['!cols'] = [
    { wch: 15 }, // Type
    { wch: 15 }, // Number
    { wch: 15 }, // Amount
    { wch: 20 }, // Bank
    { wch: 25 }, // Client
    { wch: 30 }, // Operation
    { wch: 15 }, // Operation Code
    { wch: 30 }, // Related Item
    { wch: 12 }, // Issue Date
    { wch: 12 }, // Delivery Date
    { wch: 12 }, // Expiry Date
    { wch: 10 }, // Status
    { wch: 12 }, // Return Date
    { wch: 30 }  // Notes
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'All Guarantees');

  // إضافة ورقة منفصلة لشيكات الضمان
  if (allGuaranteeChecks.length > 0) {
    const checksWorksheet = XLSX.utils.json_to_sheet(allGuaranteeChecks);
    checksWorksheet['!cols'] = worksheet['!cols'];
    XLSX.utils.book_append_sheet(workbook, checksWorksheet, 'Guarantee Checks');
  }

  // إضافة ورقة منفصلة لخطابات الضمان
  if (allGuaranteeLetters.length > 0) {
    const lettersWorksheet = XLSX.utils.json_to_sheet(allGuaranteeLetters);
    lettersWorksheet['!cols'] = worksheet['!cols'];
    XLSX.utils.book_append_sheet(workbook, lettersWorksheet, 'Guarantee Letters');
  }

  // إضافة ورقة الإحصائيات
  const stats = [
    { 'Item': 'Total Guarantees', 'Value': allGuarantees.length },
    { 'Item': 'Guarantee Checks', 'Value': allGuaranteeChecks.length },
    { 'Item': 'Guarantee Letters', 'Value': allGuaranteeLetters.length },
    { 'Item': 'Active Guarantees', 'Value': allGuarantees.filter(g => g.Status === 'Active').length },
    { 'Item': 'Returned Guarantees', 'Value': allGuarantees.filter(g => g.Status === 'Returned').length },
    { 'Item': 'Total Amount', 'Value': allGuarantees.reduce((sum, g) => sum + g.Amount, 0) }
  ];

  const statsWorksheet = XLSX.utils.json_to_sheet(stats);
  statsWorksheet['!cols'] = [{ wch: 20 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(workbook, statsWorksheet, 'Statistics');

  const fileName = `guarantees-report-${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};

// تصدير تقرير شهادات الضمان
export const exportWarrantyCertificatesReportToPDF = (operations: Operation[], clients: Client[]) => {
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4'
  });
  
  setupPDFFont(doc);

  // العنوان
  doc.setFontSize(16);
  const titleText = 'Warranty Certificates Report';
  doc.text(titleText, doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  
  // تاريخ التقرير
  doc.setFontSize(10);
  const dateText = 'Report Date: ' + formatDate(new Date());
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
        clientName: client?.name || 'Unknown',
        relatedItemDescription: relatedItem?.description || 'Full Operation'
      };
    })
  );

  if (allWarranties.length > 0) {
    const warrantyTableData = allWarranties.map(warranty => [
      warranty.certificateNumber,
      convertArabicToEnglish(warranty.clientName),
      convertArabicToEnglish(warranty.operationName),
      convertArabicToEnglish(warranty.relatedItemDescription),
      convertArabicToEnglish(warranty.description),
      formatDate(warranty.startDate),
      formatDate(warranty.endDate),
      warranty.warrantyPeriodMonths + ' months',
      warranty.isActive ? 'Active' : 'Expired'
    ]);

    (doc as any).autoTable({
      head: [['Certificate Number', 'Client', 'Operation', 'Item', 'Description', 'Start Date', 'End Date', 'Period', 'Status']],
      body: warrantyTableData,
      startY: 50,
      styles: {
        fontSize: 7,
        cellPadding: 2,
        overflow: 'linebreak',
        halign: 'center',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [34, 139, 34],
        textColor: 255,
        fontSize: 8,
        fontStyle: 'bold'
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
    doc.text('No warranty certificates found', 20, 60);
  }

  const fileName = `warranty-certificates-report-${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
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
        'Certificate Number': warranty.certificateNumber,
        'Client': client?.name || 'Unknown',
        'Operation': operation.name,
        'Operation Code': operation.code,
        'Related Item': relatedItem?.description || 'Full Operation',
        'Description': warranty.description,
        'Issue Date': formatDate(warranty.issueDate),
        'Start Date': formatDate(warranty.startDate),
        'End Date': formatDate(warranty.endDate),
        'Warranty Period (Months)': warranty.warrantyPeriodMonths,
        'Status': warranty.isActive ? 'Active' : 'Expired',
        'Notes': warranty.notes || ''
      };
    })
  );

  const worksheet = XLSX.utils.json_to_sheet(allWarranties);
  worksheet['!cols'] = [
    { wch: 20 }, // Certificate Number
    { wch: 25 }, // Client
    { wch: 30 }, // Operation
    { wch: 15 }, // Operation Code
    { wch: 30 }, // Related Item
    { wch: 40 }, // Description
    { wch: 12 }, // Issue Date
    { wch: 12 }, // Start Date
    { wch: 12 }, // End Date
    { wch: 15 }, // Warranty Period
    { wch: 10 }, // Status
    { wch: 30 }  // Notes
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Warranty Certificates');

  // إضافة ورقة الإحصائيات
  const stats = [
    { 'Item': 'Total Warranties', 'Value': allWarranties.length },
    { 'Item': 'Active Warranties', 'Value': allWarranties.filter(w => w.Status === 'Active').length },
    { 'Item': 'Expired Warranties', 'Value': allWarranties.filter(w => w.Status === 'Expired').length }
  ];

  const statsWorksheet = XLSX.utils.json_to_sheet(stats);
  statsWorksheet['!cols'] = [{ wch: 20 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(workbook, statsWorksheet, 'Statistics');

  const fileName = `warranty-certificates-${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};

// باقي الدوال مع نفس التحسينات...
export const exportOperationsToExcel = (operations: Operation[], clients: Client[], title: string = 'Operations Report') => {
  // إعداد البيانات
  const data = operations.map(operation => {
    const client = clients.find(c => c.id === operation.clientId);
    const statusLabels = {
      'in_progress': 'In Progress',
      'completed': 'Completed',
      'completed_partial_payment': 'Completed - Partial Payment'
    };

    const totalDeductions = calculateTotalDeductions(operation);
    const netAmount = calculateNetAmount(operation);

    return {
      'Operation Code': operation.code,
      'Operation Name': operation.name,
      'Client': client?.name || 'Unknown',
      'Client Type': client?.type === 'owner' ? 'Owner' : client?.type === 'main_contractor' ? 'Main Contractor' : 'Consultant',
      'Total Amount': operation.totalAmount,
      'Total Deductions': totalDeductions,
      'Net Due': netAmount,
      'Received Amount': operation.totalReceived,
      'Remaining Amount': netAmount - operation.totalReceived,
      'Completion Rate': `${operation.overallExecutionPercentage.toFixed(1)}%`,
      'Status': statusLabels[operation.status],
      'Created Date': formatDate(operation.createdAt),
      'Updated Date': formatDate(operation.updatedAt),
      'Items Count': operation.items.length,
      'Guarantee Checks': operation.guaranteeChecks.length,
      'Guarantee Letters': operation.guaranteeLetters.length,
      'Warranty Certificates': (operation.warrantyCertificates || []).length,
      'Payments Count': operation.receivedPayments.length
    };
  });

  // إنشاء ورقة العمل
  const worksheet = XLSX.utils.json_to_sheet(data);

  // تنسيق العرض
  const columnWidths = [
    { wch: 15 }, // Operation Code
    { wch: 25 }, // Operation Name
    { wch: 20 }, // Client
    { wch: 15 }, // Client Type
    { wch: 15 }, // Total Amount
    { wch: 15 }, // Total Deductions
    { wch: 15 }, // Net Due
    { wch: 15 }, // Received Amount
    { wch: 15 }, // Remaining Amount
    { wch: 12 }, // Completion Rate
    { wch: 18 }, // Status
    { wch: 12 }, // Created Date
    { wch: 12 }, // Updated Date
    { wch: 10 }, // Items Count
    { wch: 12 }, // Guarantee Checks
    { wch: 12 }, // Guarantee Letters
    { wch: 12 }, // Warranty Certificates
    { wch: 12 }  // Payments Count
  ];
  worksheet['!cols'] = columnWidths;

  // إنشاء المصنف
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Operations');

  // إضافة ورقة الإحصائيات
  const totalAmount = operations.reduce((sum, op) => sum + op.totalAmount, 0);
  const totalDeductions = operations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
  const totalNetAmount = operations.reduce((sum, op) => sum + calculateNetAmount(op), 0);
  const totalReceived = operations.reduce((sum, op) => sum + op.totalReceived, 0);

  const stats = [
    { 'Item': 'Total Operations', 'Value': operations.length },
    { 'Item': 'Completed Operations', 'Value': operations.filter(op => op.status === 'completed').length },
    { 'Item': 'In Progress Operations', 'Value': operations.filter(op => op.status === 'in_progress').length },
    { 'Item': 'Total Amount', 'Value': totalAmount },
    { 'Item': 'Total Deductions', 'Value': totalDeductions },
    { 'Item': 'Total Net Due', 'Value': totalNetAmount },
    { 'Item': 'Total Received', 'Value': totalReceived },
    { 'Item': 'Total Remaining', 'Value': totalNetAmount - totalReceived }
  ];

  const statsWorksheet = XLSX.utils.json_to_sheet(stats);
  statsWorksheet['!cols'] = [{ wch: 20 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(workbook, statsWorksheet, 'Statistics');

  // حفظ الملف
  const fileName = `operations-report-${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};

// تصدير العملاء إلى PDF
export const exportClientsToPDF = (clients: Client[], title: string = 'Clients Report') => {
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4'
  });
  
  setupPDFFont(doc);

  // العنوان
  doc.setFontSize(16);
  doc.text(title, doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  
  // تاريخ التقرير
  doc.setFontSize(10);
  const dateText = 'Report Date: ' + formatDate(new Date());
  doc.text(dateText, 20, 35);

  // إعداد البيانات للجدول
  const tableData = clients.map(client => {
    const clientTypeLabels = {
      'owner': 'Owner',
      'main_contractor': 'Main Contractor',
      'consultant': 'Consultant'
    };
    
    const mainContact = client.contacts?.find(contact => contact.isMainContact);
    
    return [
      convertArabicToEnglish(client.name),
      clientTypeLabels[client.type] || 'Unknown',
      client.phone || '-',
      client.email || '-',
      convertArabicToEnglish(client.address || '-'),
      convertArabicToEnglish(mainContact?.name || '-'),
      mainContact?.phone || '-',
      formatDate(client.createdAt)
    ];
  });

  // إعداد الجدول
  (doc as any).autoTable({
    head: [['Client Name', 'Type', 'Phone', 'Email', 'Address', 'Main Contact', 'Contact Phone', 'Added Date']],
    body: tableData,
    startY: 45,
    styles: {
      fontSize: 8,
      cellPadding: 3,
      overflow: 'linebreak',
      halign: 'center',
      font: 'helvetica'
    },
    headStyles: {
      fillColor: [66, 139, 202],
      textColor: 255,
      fontSize: 9,
      fontStyle: 'bold'
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
  doc.text('Statistics Summary:', 20, finalY);
  doc.setFontSize(10);
  doc.text('Total Clients: ' + clients.length, 20, finalY + 10);
  
  const ownerCount = clients.filter(c => c.type === 'owner').length;
  const contractorCount = clients.filter(c => c.type === 'main_contractor').length;
  const consultantCount = clients.filter(c => c.type === 'consultant').length;
  
  doc.text('Owners: ' + ownerCount, 20, finalY + 20);
  doc.text('Main Contractors: ' + contractorCount, 20, finalY + 30);
  doc.text('Consultants: ' + consultantCount, 20, finalY + 40);

  // حفظ الملف
  const fileName = `clients-report-${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
};

// تصدير العملاء إلى Excel
export const exportClientsToExcel = (clients: Client[], title: string = 'Clients Report') => {
  const data = clients.map(client => {
    const clientTypeLabels = {
      'owner': 'Owner',
      'main_contractor': 'Main Contractor',
      'consultant': 'Consultant'
    };
    
    const mainContact = client.contacts?.find(contact => contact.isMainContact);
    
    return {
      'Client Name': client.name,
      'Client Type': clientTypeLabels[client.type] || 'Unknown',
      'Phone Number': client.phone || '',
      'Email': client.email || '',
      'Address': client.address || '',
      'Main Contact': mainContact?.name || '',
      'Contact Position': mainContact?.position || '',
      'Contact Department': mainContact?.department || '',
      'Contact Phone': mainContact?.phone || '',
      'Contact Email': mainContact?.email || '',
      'Contacts Count': client.contacts?.length || 0,
      'Added Date': formatDate(client.createdAt),
      'Updated Date': formatDate(client.updatedAt)
    };
  });

  const worksheet = XLSX.utils.json_to_sheet(data);
  worksheet['!cols'] = [
    { wch: 25 }, // Client Name
    { wch: 15 }, // Client Type
    { wch: 15 }, // Phone Number
    { wch: 30 }, // Email
    { wch: 30 }, // Address
    { wch: 25 }, // Main Contact
    { wch: 20 }, // Contact Position
    { wch: 15 }, // Contact Department
    { wch: 15 }, // Contact Phone
    { wch: 25 }, // Contact Email
    { wch: 12 }, // Contacts Count
    { wch: 15 }, // Added Date
    { wch: 15 }  // Updated Date
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Clients');

  const fileName = `clients-report-${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};

// تصدير تقرير مالي شامل
export const exportFinancialReportToPDF = (operations: Operation[], clients: Client[]) => {
  const doc = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4'
  });
  
  setupPDFFont(doc);

  // العنوان
  doc.setFontSize(18);
  const titleText = 'Comprehensive Financial Report';
  doc.text(titleText, doc.internal.pageSize.getWidth() / 2, 20, { align: 'center' });
  
  // تاريخ التقرير
  doc.setFontSize(10);
  const dateText = 'Report Date: ' + formatDate(new Date());
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
  doc.text('General Statistics:', 20, yPosition);
  
  doc.setFontSize(11);
  yPosition += 15;
  doc.text('Total Operations: ' + totalOperations, 30, yPosition);
  yPosition += 10;
  doc.text('Completed Operations: ' + completedOperations, 30, yPosition);
  yPosition += 10;
  doc.text('In Progress Operations: ' + inProgressOperations, 30, yPosition);
  yPosition += 10;
  doc.text('Completion Rate: ' + (totalOperations > 0 ? ((completedOperations / totalOperations) * 100).toFixed(1) : 0) + '%', 30, yPosition);

  // الإحصائيات المالية
  yPosition += 20;
  doc.setFontSize(14);
  doc.text('Financial Statistics:', 20, yPosition);
  
  doc.setFontSize(11);
  yPosition += 15;
  doc.text('Total Amount: ' + formatCurrency(totalAmount), 30, yPosition);
  yPosition += 10;
  doc.text('Total Deductions: ' + formatCurrency(totalDeductions), 30, yPosition);
  yPosition += 10;
  doc.text('Net Due Amount: ' + formatCurrency(totalNetAmount), 30, yPosition);
  yPosition += 10;
  doc.text('Total Received: ' + formatCurrency(totalReceived), 30, yPosition);
  yPosition += 10;
  doc.text('Total Outstanding: ' + formatCurrency(totalOutstanding), 30, yPosition);
  yPosition += 10;
  doc.text('Collection Rate: ' + collectionRate.toFixed(1) + '%', 30, yPosition);

  // جدول العمليات حسب العميل
  yPosition += 30;
  const clientStats = clients.map(client => {
    const clientOperations = operations.filter(op => op.clientId === client.id);
    const clientTotal = clientOperations.reduce((sum, op) => sum + op.totalAmount, 0);
    const clientDeductions = clientOperations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
    const clientNetAmount = clientOperations.reduce((sum, op) => sum + calculateNetAmount(op), 0);
    const clientReceived = clientOperations.reduce((sum, op) => sum + op.totalReceived, 0);
    
    return [
      convertArabicToEnglish(client.name),
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
      head: [['Client', 'Operations', 'Total Amount', 'Deductions', 'Net Due', 'Received', 'Outstanding', 'Collection Rate']],
      body: clientStats,
      startY: yPosition,
      styles: {
        fontSize: 8,
        cellPadding: 3,
        halign: 'center',
        font: 'helvetica'
      },
      headStyles: {
        fillColor: [66, 139, 202],
        textColor: 255,
        fontSize: 9,
        fontStyle: 'bold'
      },
      alternateRowStyles: {
        fillColor: [245, 245, 245]
      }
    });
  }

  const fileName = `financial-report-${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
};

// تصدير تقرير مالي إلى Excel
export const exportFinancialReportToExcel = (operations: Operation[], clients: Client[]) => {
  // ورقة الإحصائيات العامة
  const totalAmount = operations.reduce((sum, op) => sum + op.totalAmount, 0);
  const totalDeductions = operations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
  const totalNetAmount = operations.reduce((sum, op) => sum + calculateNetAmount(op), 0);
  const totalReceived = operations.reduce((sum, op) => sum + op.totalReceived, 0);

  const generalStats = [
    { 'Item': 'Total Operations', 'Value': operations.length },
    { 'Item': 'Completed Operations', 'Value': operations.filter(op => op.status === 'completed').length },
    { 'Item': 'In Progress Operations', 'Value': operations.filter(op => op.status === 'in_progress').length },
    { 'Item': 'Total Amount', 'Value': totalAmount },
    { 'Item': 'Total Deductions', 'Value': totalDeductions },
    { 'Item': 'Net Due Amount', 'Value': totalNetAmount },
    { 'Item': 'Total Received', 'Value': totalReceived },
    { 'Item': 'Total Outstanding', 'Value': totalNetAmount - totalReceived }
  ];

  // ورقة إحصائيات العملاء
  const clientStats = clients.map(client => {
    const clientOperations = operations.filter(op => op.clientId === client.id);
    const clientTotal = clientOperations.reduce((sum, op) => sum + op.totalAmount, 0);
    const clientDeductions = clientOperations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
    const clientNetAmount = clientOperations.reduce((sum, op) => sum + calculateNetAmount(op), 0);
    const clientReceived = clientOperations.reduce((sum, op) => sum + op.totalReceived, 0);
    
    return {
      'Client': client.name,
      'Client Type': client.type === 'owner' ? 'Owner' : client.type === 'main_contractor' ? 'Main Contractor' : 'Consultant',
      'Operations Count': clientOperations.length,
      'Total Amount': clientTotal,
      'Total Deductions': clientDeductions,
      'Net Due Amount': clientNetAmount,
      'Received Amount': clientReceived,
      'Outstanding Amount': clientNetAmount - clientReceived,
      'Collection Rate': clientNetAmount > 0 ? `${((clientReceived / clientNetAmount) * 100).toFixed(1)}%` : '0%'
    };
  }).filter(stat => stat['Operations Count'] > 0);

  // إنشاء المصنف
  const workbook = XLSX.utils.book_new();

  // إضافة ورقة الإحصائيات العامة
  const statsWorksheet = XLSX.utils.json_to_sheet(generalStats);
  statsWorksheet['!cols'] = [{ wch: 25 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(workbook, statsWorksheet, 'General Statistics');

  // إضافة ورقة إحصائيات العملاء
  if (clientStats.length > 0) {
    const clientWorksheet = XLSX.utils.json_to_sheet(clientStats);
    clientWorksheet['!cols'] = [
      { wch: 25 }, // Client
      { wch: 15 }, // Client Type
      { wch: 12 }, // Operations Count
      { wch: 15 }, // Total Amount
      { wch: 15 }, // Total Deductions
      { wch: 15 }, // Net Due Amount
      { wch: 15 }, // Received Amount
      { wch: 15 }, // Outstanding Amount
      { wch: 15 }  // Collection Rate
    ];
    XLSX.utils.book_append_sheet(workbook, clientWorksheet, 'Client Statistics');
  }

  const fileName = `financial-report-${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};