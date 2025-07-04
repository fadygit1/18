import React, { useState } from 'react';
import { Plus, Trash2, Save, X, Calendar, CreditCard, FileText } from 'lucide-react';
import { Client, Operation, OperationItem, Deduction, GuaranteeCheck, GuaranteeLetter, ReceivedPayment } from '../../types';
import { generateOperationCode, generateItemCode, calculateOperationTotal, calculateOverallExecutionPercentage, calculateOperationStatus, formatCurrency } from '../../utils/calculations';

interface AddOperationFormProps {
  clients: Client[];
  onSave: (operation: Operation) => void;
  onCancel: () => void;
}

const AddOperationForm: React.FC<AddOperationFormProps> = ({ clients, onSave, onCancel }) => {
  const [formData, setFormData] = useState({
    clientId: '',
    name: '',
  });

  const [items, setItems] = useState<OperationItem[]>([
    {
      id: crypto.randomUUID(),
      code: '',
      description: '',
      amount: 0,
      executionPercentage: 0,
      addedAt: new Date(),
    }
  ]);

  const [deductions, setDeductions] = useState<Deduction[]>([]);
  const [guaranteeChecks, setGuaranteeChecks] = useState<GuaranteeCheck[]>([]);
  const [guaranteeLetters, setGuaranteeLetters] = useState<GuaranteeLetter[]>([]);
  const [receivedPayments, setReceivedPayments] = useState<ReceivedPayment[]>([]);
  const [activeSection, setActiveSection] = useState<string>('basic');

  const addItem = () => {
    const operationCode = formData.clientId && formData.name 
      ? generateOperationCode(
          clients.find(c => c.id === formData.clientId)?.name || '',
          formData.name
        )
      : 'TEMP';

    const newItem: OperationItem = {
      id: crypto.randomUUID(),
      code: generateItemCode(operationCode, items.length),
      description: '',
      amount: 0,
      executionPercentage: 0,
      addedAt: new Date(),
    };
    setItems([...items, newItem]);
  };

  const updateItem = (index: number, field: keyof OperationItem, value: any) => {
    const updatedItems = [...items];
    updatedItems[index] = { ...updatedItems[index], [field]: value };
    setItems(updatedItems);
  };

  const removeItem = (index: number) => {
    if (items.length > 1) {
      const newItems = items.filter((_, i) => i !== index);
      const operationCode = formData.clientId && formData.name 
        ? generateOperationCode(
            clients.find(c => c.id === formData.clientId)?.name || '',
            formData.name
          )
        : 'TEMP';
      
      const updatedItems = newItems.map((item, i) => ({
        ...item,
        code: generateItemCode(operationCode, i)
      }));
      setItems(updatedItems);
    }
  };

  const updateDeduction = (index: number, field: keyof Deduction, value: any) => {
    const updatedDeductions = [...deductions];
    updatedDeductions[index] = { ...updatedDeductions[index], [field]: value };
    setDeductions(updatedDeductions);
  };

  const addDeduction = () => {
    const newDeduction: Deduction = {
      id: crypto.randomUUID(),
      name: '',
      type: 'percentage',
      value: 0,
      isActive: true
    };
    setDeductions([...deductions, newDeduction]);
  };

  const removeDeduction = (index: number) => {
    setDeductions(deductions.filter((_, i) => i !== index));
  };

  const addGuaranteeCheck = () => {
    const today = new Date();
    const expiryDate = new Date(today);
    expiryDate.setFullYear(expiryDate.getFullYear() + 1);

    const newCheck: GuaranteeCheck = {
      id: crypto.randomUUID(),
      checkNumber: '',
      amount: 0,
      checkDate: today,
      deliveryDate: today,
      expiryDate: expiryDate,
      bank: '',
      isReturned: false,
      relatedTo: 'operation'
    };
    setGuaranteeChecks([...guaranteeChecks, newCheck]);
  };

  const updateGuaranteeCheck = (index: number, field: keyof GuaranteeCheck, value: any) => {
    const updatedChecks = [...guaranteeChecks];
    updatedChecks[index] = { ...updatedChecks[index], [field]: value };
    setGuaranteeChecks(updatedChecks);
  };

  const removeGuaranteeCheck = (index: number) => {
    setGuaranteeChecks(guaranteeChecks.filter((_, i) => i !== index));
  };

  const addGuaranteeLetter = () => {
    const today = new Date();
    const dueDate = new Date(today);
    dueDate.setFullYear(dueDate.getFullYear() + 1);

    const newLetter: GuaranteeLetter = {
      id: crypto.randomUUID(),
      bank: '',
      letterDate: today,
      letterNumber: '',
      amount: 0,
      dueDate: dueDate,
      renewals: [],
      relatedTo: 'operation',
      isReturned: false,
    };
    setGuaranteeLetters([...guaranteeLetters, newLetter]);
  };

  const updateGuaranteeLetter = (index: number, field: keyof GuaranteeLetter, value: any) => {
    const updatedLetters = [...guaranteeLetters];
    updatedLetters[index] = { ...updatedLetters[index], [field]: value };
    setGuaranteeLetters(updatedLetters);
  };

  const removeGuaranteeLetter = (index: number) => {
    setGuaranteeLetters(guaranteeLetters.filter((_, i) => i !== index));
  };

  const addReceivedPayment = () => {
    const newPayment: ReceivedPayment = {
      id: crypto.randomUUID(),
      type: 'cash',
      amount: 0,
      date: new Date(),
    };
    setReceivedPayments([...receivedPayments, newPayment]);
  };

  const updateReceivedPayment = (index: number, field: keyof ReceivedPayment, value: any) => {
    const updatedPayments = [...receivedPayments];
    updatedPayments[index] = { ...updatedPayments[index], [field]: value };
    setReceivedPayments(updatedPayments);
  };

  const removeReceivedPayment = (index: number) => {
    setReceivedPayments(receivedPayments.filter((_, i) => i !== index));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    
    if (!formData.clientId || !formData.name) {
      alert('يرجى إدخال اسم العميل واسم العملية');
      return;
    }

    if (items.some(item => !item.description || item.amount <= 0)) {
      alert('يرجى إدخال وصف وقيمة صحيحة لجميع البنود');
      return;
    }

    const client = clients.find(c => c.id === formData.clientId);
    if (!client) return;

    const operationCode = generateOperationCode(client.name, formData.name);
    const totalAmount = calculateOperationTotal(items);
    const overallExecutionPercentage = calculateOverallExecutionPercentage(items);
    const totalReceived = receivedPayments.reduce((sum, payment) => sum + payment.amount, 0);

    // تحديث أكواد البنود
    const updatedItems = items.map((item, index) => ({
      ...item,
      code: generateItemCode(operationCode, index)
    }));

    const newOperation: Operation = {
      id: crypto.randomUUID(),
      code: operationCode,
      name: formData.name,
      clientId: formData.clientId,
      items: updatedItems,
      deductions,
      guaranteeChecks,
      guaranteeLetters,
      warrantyCertificates: [],
      receivedPayments,
      totalAmount,
      totalReceived,
      overallExecutionPercentage,
      status: calculateOperationStatus({
        overallExecutionPercentage,
        totalReceived,
        items: updatedItems,
        deductions
      } as Operation),
      createdAt: new Date(),
      updatedAt: new Date(),
    };

    onSave(newOperation);
  };

  const sections = [
    { id: 'basic', label: 'البيانات الأساسية', icon: FileText },
    { id: 'items', label: 'بنود العملية', icon: Plus },
    { id: 'deductions', label: 'الخصومات', icon: Trash2 },
    { id: 'guarantees', label: 'الضمانات', icon: CreditCard },
    { id: 'payments', label: 'المدفوعات', icon: Calendar },
  ];

  return (
    <div className="p-6 max-w-7xl mx-auto">
      <div className="mb-6">
        <h2 className="text-3xl font-bold text-gray-900 mb-2">إضافة عملية جديدة</h2>
        <p className="text-gray-600">إنشاء عملية إنشائية جديدة مع جميع التفاصيل والبيانات المطلوبة</p>
      </div>

      {/* Section Navigation */}
      <div className="mb-8 bg-white rounded-lg border border-gray-200 p-1">
        <nav className="flex flex-wrap gap-1">
          {sections.map((section) => {
            const Icon = section.icon;
            return (
              <button
                key={section.id}
                onClick={() => setActiveSection(section.id)}
                className={`flex items-center gap-2 px-4 py-3 rounded-md font-medium text-sm transition-all ${
                  activeSection === section.id
                    ? 'bg-blue-600 text-white shadow-md'
                    : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
                }`}
              >
                <Icon className="w-4 h-4" />
                {section.label}
              </button>
            );
          })}
        </nav>
      </div>

      <form onSubmit={handleSubmit} className="space-y-8">
        {/* Basic Information */}
        {activeSection === 'basic' && (
          <div className="bg-white p-8 rounded-lg border border-gray-200 shadow-sm">
            <h3 className="text-xl font-semibold text-gray-900 mb-6">البيانات الأساسية</h3>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  العميل *
                </label>
                <select
                  value={formData.clientId}
                  onChange={(e) => setFormData({ ...formData, clientId: e.target.value })}
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  required
                >
                  <option value="">اختر العميل</option>
                  {clients.map((client) => (
                    <option key={client.id} value={client.id}>
                      {client.name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  اسم العملية *
                </label>
                <input
                  type="text"
                  value={formData.name}
                  onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  placeholder="أدخل اسم العملية"
                  required
                />
              </div>

              {formData.clientId && formData.name && (
                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    كود العملية (سيتم إنشاؤه تلقائياً)
                  </label>
                  <input
                    type="text"
                    value={generateOperationCode(
                      clients.find(c => c.id === formData.clientId)?.name || '',
                      formData.name
                    )}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg bg-gray-50 text-gray-600"
                    readOnly
                  />
                </div>
              )}
            </div>
          </div>
        )}

        {/* Operation Items */}
        {activeSection === 'items' && (
          <div className="bg-white p-8 rounded-lg border border-gray-200 shadow-sm">
            <div className="flex justify-between items-center mb-6">
              <h3 className="text-xl font-semibold text-gray-900">بنود العملية</h3>
              <button
                type="button"
                onClick={addItem}
                className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
              >
                <Plus className="w-4 h-4" />
                إضافة بند
              </button>
            </div>

            <div className="space-y-6">
              {items.map((item, index) => (
                <div key={item.id} className="p-6 border border-gray-200 rounded-lg bg-gray-50">
                  <div className="flex justify-between items-center mb-4">
                    <span className="text-lg font-medium text-gray-700">البند {index + 1}</span>
                    {items.length > 1 && (
                      <button
                        type="button"
                        onClick={() => removeItem(index)}
                        className="text-red-600 hover:text-red-800 p-2 rounded-lg hover:bg-red-50"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    )}
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-2">
                        الكود
                      </label>
                      <input
                        type="text"
                        value={formData.clientId && formData.name 
                          ? generateItemCode(
                              generateOperationCode(
                                clients.find(c => c.id === formData.clientId)?.name || '',
                                formData.name
                              ),
                              index
                            )
                          : 'TEMP-001'
                        }
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg bg-white text-gray-600"
                        readOnly
                      />
                    </div>

                    <div className="md:col-span-2">
                      <label className="block text-sm font-medium text-gray-700 mb-2">
                        الوصف *
                      </label>
                      <input
                        type="text"
                        value={item.description}
                        onChange={(e) => updateItem(index, 'description', e.target.value)}
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                        placeholder="وصف البند"
                        required
                      />
                    </div>

                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-2">
                        القيمة (ج.م) *
                      </label>
                      <input
                        type="number"
                        value={item.amount}
                        onChange={(e) => updateItem(index, 'amount', parseFloat(e.target.value) || 0)}
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                        min="0"
                        step="0.01"
                        placeholder="0.00"
                        required
                      />
                    </div>

                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-2">
                        نسبة التنفيذ (%)
                      </label>
                      <input
                        type="number"
                        value={item.executionPercentage}
                        onChange={(e) => updateItem(index, 'executionPercentage', parseFloat(e.target.value) || 0)}
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                        min="0"
                        max="100"
                        step="0.1"
                        placeholder="0.0"
                      />
                    </div>

                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-2">
                        رقم العقد
                      </label>
                      <input
                        type="text"
                        value={item.contractNumber || ''}
                        onChange={(e) => updateItem(index, 'contractNumber', e.target.value)}
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                        placeholder="رقم العقد"
                      />
                    </div>

                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-2">
                        تاريخ العقد
                      </label>
                      <input
                        type="date"
                        value={item.contractDate ? item.contractDate.toISOString().split('T')[0] : ''}
                        onChange={(e) => updateItem(index, 'contractDate', e.target.value ? new Date(e.target.value) : undefined)}
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                      />
                    </div>
                  </div>
                </div>
              ))}
            </div>

            <div className="mt-6 p-6 bg-blue-50 rounded-lg border border-blue-200">
              <div className="flex justify-between items-center text-lg">
                <span className="font-medium text-blue-900">إجمالي العملية:</span>
                <span className="font-bold text-blue-900">
                  {formatCurrency(calculateOperationTotal(items))}
                </span>
              </div>
            </div>
          </div>
        )}

        {/* Deductions Section */}
        {activeSection === 'deductions' && (
          <div className="bg-white p-8 rounded-lg border border-gray-200 shadow-sm">
            <div className="flex justify-between items-center mb-6">
              <h3 className="text-xl font-semibold text-gray-900">الخصومات</h3>
              <button
                type="button"
                onClick={addDeduction}
                className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors"
              >
                <Plus className="w-4 h-4" />
                إضافة خصم
              </button>
            </div>

            <div className="space-y-4">
              {deductions.map((deduction, index) => (
                <div key={deduction.id} className="p-4 border border-gray-200 rounded-lg bg-gray-50">
                  <div className="grid grid-cols-1 md:grid-cols-5 gap-4 items-center">
                    <div className="md:col-span-2">
                      <label className="block text-sm font-medium text-gray-700 mb-1">
                        اسم الخصم
                      </label>
                      <input
                        type="text"
                        value={deduction.name}
                        onChange={(e) => updateDeduction(index, 'name', e.target.value)}
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                        placeholder="اسم الخصم"
                      />
                    </div>

                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-1">
                        النوع
                      </label>
                      <select
                        value={deduction.type}
                        onChange={(e) => updateDeduction(index, 'type', e.target.value as 'percentage' | 'fixed')}
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                      >
                        <option value="percentage">نسبة مئوية</option>
                        <option value="fixed">مبلغ ثابت</option>
                      </select>
                    </div>

                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-1">
                        القيمة {deduction.type === 'percentage' ? '(%)' : '(ج.م)'}
                      </label>
                      <input
                        type="number"
                        value={deduction.value}
                        onChange={(e) => updateDeduction(index, 'value', parseFloat(e.target.value) || 0)}
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                        min="0"
                        step="0.01"
                        placeholder="0.00"
                      />
                    </div>

                    <div className="flex items-center justify-between">
                      <label className="flex items-center">
                        <input
                          type="checkbox"
                          checked={deduction.isActive}
                          onChange={(e) => updateDeduction(index, 'isActive', e.target.checked)}
                          className="mr-2 h-4 w-4 text-blue-600 focus:ring-blue-500 border-gray-300 rounded"
                        />
                        <span className="text-sm text-gray-700">فعال</span>
                      </label>
                      
                      <button
                        type="button"
                        onClick={() => removeDeduction(index)}
                        className="text-red-600 hover:text-red-800 p-1 rounded hover:bg-red-50"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                </div>
              ))}
            </div>

            {deductions.length === 0 && (
              <div className="text-center py-8 text-gray-500">
                لا توجد خصومات مضافة. يمكنك إضافة خصومات حسب الحاجة.
              </div>
            )}
          </div>
        )}

        {/* Guarantees Section */}
        {activeSection === 'guarantees' && (
          <div className="space-y-8">
            {/* Guarantee Checks */}
            <div className="bg-white p-8 rounded-lg border border-gray-200 shadow-sm">
              <div className="flex justify-between items-center mb-6">
                <h3 className="text-xl font-semibold text-gray-900">شيكات الضمان</h3>
                <button
                  type="button"
                  onClick={addGuaranteeCheck}
                  className="flex items-center gap-2 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition-colors"
                >
                  <Plus className="w-4 h-4" />
                  إضافة شيك ضمان
                </button>
              </div>

              <div className="space-y-6">
                {guaranteeChecks.map((check, index) => (
                  <div key={check.id} className="p-6 border border-gray-200 rounded-lg bg-purple-50">
                    <div className="flex justify-between items-center mb-4">
                      <span className="text-lg font-medium text-purple-900">شيك الضمان {index + 1}</span>
                      <button
                        type="button"
                        onClick={() => removeGuaranteeCheck(index)}
                        className="text-red-600 hover:text-red-800 p-2 rounded-lg hover:bg-red-50"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                    
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                          رقم الشيك
                        </label>
                        <input
                          type="text"
                          value={check.checkNumber}
                          onChange={(e) => updateGuaranteeCheck(index, 'checkNumber', e.target.value)}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500"
                          placeholder="رقم الشيك"
                        />
                      </div>

                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                          المبلغ (ج.م)
                        </label>
                        <input
                          type="number"
                          value={check.amount}
                          onChange={(e) => updateGuaranteeCheck(index, 'amount', parseFloat(e.target.value) || 0)}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500"
                          min="0"
                          step="0.01"
                          placeholder="0.00"
                        />
                      </div>

                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                          البنك
                        </label>
                        <input
                          type="text"
                          value={check.bank}
                          onChange={(e) => updateGuaranteeCheck(index, 'bank', e.target.value)}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500"
                          placeholder="اسم البنك"
                        />
                      </div>

                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                          تاريخ الشيك
                        </label>
                        <input
                          type="date"
                          value={check.checkDate.toISOString().split('T')[0]}
                          onChange={(e) => updateGuaranteeCheck(index, 'checkDate', new Date(e.target.value))}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500"
                        />
                      </div>

                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                          تاريخ التسليم
                        </label>
                        <input
                          type="date"
                          value={check.deliveryDate.toISOString().split('T')[0]}
                          onChange={(e) => updateGuaranteeCheck(index, 'deliveryDate', new Date(e.target.value))}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500"
                        />
                      </div>

                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                          تاريخ انتهاء الصلاحية
                        </label>
                        <input
                          type="date"
                          value={check.expiryDate.toISOString().split('T')[0]}
                          onChange={(e) => updateGuaranteeCheck(index, 'expiryDate', new Date(e.target.value))}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500"
                        />
                      </div>
                    </div>
                  </div>
                ))}
              </div>

              {guaranteeChecks.length === 0 && (
                <div className="text-center py-8 text-gray-500">
                  لا توجد شيكات ضمان مضافة. يمكنك إضافة شيكات ضمان حسب الحاجة.
                </div>
              )}
            </div>

            {/* Guarantee Letters */}
            <div className="bg-white p-8 rounded-lg border border-gray-200 shadow-sm">
              <div className="flex justify-between items-center mb-6">
                <h3 className="text-xl font-semibold text-gray-900">خطابات الضمان</h3>
                <button
                  type="button"
                  onClick={addGuaranteeLetter}
                  className="flex items-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition-colors"
                >
                  <Plus className="w-4 h-4" />
                  إضافة خطاب ضمان
                </button>
              </div>

              <div className="space-y-6">
                {guaranteeLetters.map((letter, index) => (
                  <div key={letter.id} className="p-6 border border-gray-200 rounded-lg bg-indigo-50">
                    <div className="flex justify-between items-center mb-4">
                      <span className="text-lg font-medium text-indigo-900">خطاب الضمان {index + 1}</span>
                      <button
                        type="button"
                        onClick={() => removeGuaranteeLetter(index)}
                        className="text-red-600 hover:text-red-800 p-2 rounded-lg hover:bg-red-50"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                    
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                          البنك
                        </label>
                        <input
                          type="text"
                          value={letter.bank}
                          onChange={(e) => updateGuaranteeLetter(index, 'bank', e.target.value)}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                          placeholder="اسم البنك"
                        />
                      </div>

                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                          رقم الخطاب
                        </label>
                        <input
                          type="text"
                          value={letter.letterNumber}
                          onChange={(e) => updateGuaranteeLetter(index, 'letterNumber', e.target.value)}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                          placeholder="رقم الخطاب"
                        />
                      </div>

                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                          المبلغ (ج.م)
                        </label>
                        <input
                          type="number"
                          value={letter.amount}
                          onChange={(e) => updateGuaranteeLetter(index, 'amount', parseFloat(e.target.value) || 0)}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                          min="0"
                          step="0.01"
                          placeholder="0.00"
                        />
                      </div>

                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                          تاريخ الخطاب
                        </label>
                        <input
                          type="date"
                          value={letter.letterDate.toISOString().split('T')[0]}
                          onChange={(e) => updateGuaranteeLetter(index, 'letterDate', new Date(e.target.value))}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                        />
                      </div>

                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                          تاريخ الاستحقاق
                        </label>
                        <input
                          type="date"
                          value={letter.dueDate.toISOString().split('T')[0]}
                          onChange={(e) => updateGuaranteeLetter(index, 'dueDate', new Date(e.target.value))}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                        />
                      </div>

                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                          مرتبط بـ
                        </label>
                        <select
                          value={letter.relatedTo}
                          onChange={(e) => updateGuaranteeLetter(index, 'relatedTo', e.target.value as 'operation' | 'item')}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                        >
                          <option value="operation">العملية كاملة</option>
                          <option value="item">بند محدد</option>
                        </select>
                      </div>

                      <div className="md:col-span-2">
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                          ملاحظات
                        </label>
                        <input
                          type="text"
                          value={letter.notes || ''}
                          onChange={(e) => updateGuaranteeLetter(index, 'notes', e.target.value)}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
                          placeholder="ملاحظات إضافية"
                        />
                      </div>
                    </div>
                  </div>
                ))}
              </div>

              {guaranteeLetters.length === 0 && (
                <div className="text-center py-8 text-gray-500">
                  لا توجد خطابات ضمان مضافة. يمكنك إضافة خطابات ضمان حسب الحاجة.
                </div>
              )}
            </div>
          </div>
        )}

        {/* Received Payments Section */}
        {activeSection === 'payments' && (
          <div className="bg-white p-8 rounded-lg border border-gray-200 shadow-sm">
            <div className="flex justify-between items-center mb-6">
              <h3 className="text-xl font-semibold text-gray-900">المدفوعات المستلمة</h3>
              <button
                type="button"
                onClick={addReceivedPayment}
                className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors"
              >
                <Plus className="w-4 h-4" />
                إضافة دفعة
              </button>
            </div>

            <div className="space-y-6">
              {receivedPayments.map((payment, index) => (
                <div key={payment.id} className="p-6 border border-gray-200 rounded-lg bg-green-50">
                  <div className="flex justify-between items-center mb-4">
                    <span className="text-lg font-medium text-green-900">الدفعة {index + 1}</span>
                    <button
                      type="button"
                      onClick={() => removeReceivedPayment(index)}
                      className="text-red-600 hover:text-red-800 p-2 rounded-lg hover:bg-red-50"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-2">
                        نوع الدفع
                      </label>
                      <select
                        value={payment.type}
                        onChange={(e) => updateReceivedPayment(index, 'type', e.target.value as 'check' | 'cash')}
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-green-500"
                      >
                        <option value="cash">نقدي</option>
                        <option value="check">شيك</option>
                      </select>
                    </div>

                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-2">
                        المبلغ (ج.م)
                      </label>
                      <input
                        type="number"
                        value={payment.amount}
                        onChange={(e) => updateReceivedPayment(index, 'amount', parseFloat(e.target.value) || 0)}
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-green-500"
                        min="0"
                        step="0.01"
                        placeholder="0.00"
                      />
                    </div>

                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-2">
                        التاريخ
                      </label>
                      <input
                        type="date"
                        value={payment.date.toISOString().split('T')[0]}
                        onChange={(e) => updateReceivedPayment(index, 'date', new Date(e.target.value))}
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-green-500"
                      />
                    </div>

                    {payment.type === 'check' && (
                      <>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">
                            رقم الشيك
                          </label>
                          <input
                            type="text"
                            value={payment.checkNumber || ''}
                            onChange={(e) => updateReceivedPayment(index, 'checkNumber', e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-green-500"
                            placeholder="رقم الشيك"
                          />
                        </div>

                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">
                            البنك
                          </label>
                          <input
                            type="text"
                            value={payment.bank || ''}
                            onChange={(e) => updateReceivedPayment(index, 'bank', e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-green-500"
                            placeholder="اسم البنك"
                          />
                        </div>

                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">
                            تاريخ الاستلام
                          </label>
                          <input
                            type="date"
                            value={payment.receiptDate ? payment.receiptDate.toISOString().split('T')[0] : ''}
                            onChange={(e) => updateReceivedPayment(index, 'receiptDate', e.target.value ? new Date(e.target.value) : undefined)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-green-500"
                          />
                        </div>
                      </>
                    )}

                    <div className="md:col-span-2">
                      <label className="block text-sm font-medium text-gray-700 mb-2">
                        ملاحظات
                      </label>
                      <input
                        type="text"
                        value={payment.notes || ''}
                        onChange={(e) => updateReceivedPayment(index, 'notes', e.target.value)}
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-green-500"
                        placeholder="ملاحظات إضافية"
                      />
                    </div>
                  </div>
                </div>
              ))}
            </div>

            {receivedPayments.length === 0 && (
              <div className="text-center py-8 text-gray-500">
                لا توجد مدفوعات مضافة. يمكنك إضافة المدفوعات المستلمة حسب الحاجة.
              </div>
            )}

            {receivedPayments.length > 0 && (
              <div className="mt-6 p-6 bg-green-100 rounded-lg border border-green-200">
                <div className="flex justify-between items-center text-lg">
                  <span className="font-medium text-green-900">إجمالي المدفوعات المستلمة:</span>
                  <span className="font-bold text-green-900">
                    {formatCurrency(receivedPayments.reduce((sum, payment) => sum + payment.amount, 0))}
                  </span>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Form Actions */}
        <div className="flex gap-4 pt-8 border-t border-gray-200 bg-white p-6 rounded-lg">
          <button
            type="submit"
            className="flex items-center gap-2 px-8 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors font-medium"
          >
            <Save className="w-5 h-5" />
            حفظ العملية
          </button>
          
          <button
            type="button"
            onClick={onCancel}
            className="flex items-center gap-2 px-8 py-3 bg-gray-300 text-gray-700 rounded-lg hover:bg-gray-400 transition-colors font-medium"
          >
            <X className="w-5 h-5" />
            إلغاء
          </button>
        </div>
      </form>
    </div>
  );
};

export default AddOperationForm;