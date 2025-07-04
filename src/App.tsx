import React, { useState, useEffect } from 'react';
import Header from './components/Layout/Header';
import Sidebar from './components/Layout/Sidebar';
import Dashboard from './components/Dashboard/Dashboard';
import AddOperationForm from './components/Operations/AddOperationForm';
import OperationsList from './components/Operations/OperationsList';
import OperationDetails from './components/Operations/OperationDetails';
import EditOperationForm from './components/Operations/EditOperationForm';
import ClientsManager from './components/Clients/ClientsManager';
import GuaranteesManager from './components/Guarantees/GuaranteesManager';
import WarrantiesManager from './components/Warranties/WarrantiesManager';
import ChecksManager from './components/Checks/ChecksManager';
import ReportsManager from './components/Reports/ReportsManager';
import { Operation, Client, DashboardStats } from './types';
import { storage } from './utils/storage';
import { initializeDatabase, dbOperations } from './utils/database';
import { calculateNetAmount, calculateTotalDeductions, calculateOperationStatus } from './utils/calculations';

function App() {
  const [activeView, setActiveView] = useState('dashboard');
  const [operations, setOperations] = useState<Operation[]>([]);
  const [clients, setClients] = useState<Client[]>([]);
  const [selectedOperation, setSelectedOperation] = useState<Operation | null>(null);
  const [isLoading, setIsLoading] = useState(true);

  // تحميل البيانات من قاعدة البيانات عند بدء التطبيق
  useEffect(() => {
    const loadData = async () => {
      try {
        setIsLoading(true);
        
        // تهيئة قاعدة البيانات
        await initializeDatabase();
        
        // تحميل البيانات
        const [loadedOperations, loadedClients] = await Promise.all([
          storage.getOperations(),
          storage.getClients()
        ]);
        
        setOperations(loadedOperations);
        setClients(loadedClients);
        
        console.log('تم تحميل البيانات بنجاح:', {
          operations: loadedOperations.length,
          clients: loadedClients.length
        });
      } catch (error) {
        console.error('خطأ في تحميل البيانات:', error);
      } finally {
        setIsLoading(false);
      }
    };

    loadData();
  }, []);

  // Navigation handler
  const handleNavigate = (view: string) => {
    setActiveView(view);
    setSelectedOperation(null);
  };

  // Calculate dashboard statistics with deductions
  const getDashboardStats = (): DashboardStats => {
    const totalOperations = operations.length;
    const completedOperations = operations.filter(op => 
      op.status === 'completed' || 
      op.status === 'completed_partial_payment' || 
      op.status === 'completed_full_payment'
    ).length;
    const inProgressOperations = operations.filter(op => op.status === 'in_progress').length;
    const totalAmount = operations.reduce((sum, op) => sum + op.totalAmount, 0);
    const totalDeductions = operations.reduce((sum, op) => sum + calculateTotalDeductions(op), 0);
    const totalNetAmount = operations.reduce((sum, op) => sum + calculateNetAmount(op), 0);
    const totalReceived = operations.reduce((sum, op) => sum + op.totalReceived, 0);
    const outstandingGuarantees = operations.reduce((sum, op) => 
      sum + op.guaranteeChecks.filter(check => !check.isReturned).length +
      op.guaranteeLetters.filter(letter => !letter.isReturned).length, 0
    );
    const activeWarranties = operations.reduce((sum, op) => 
      sum + (op.warrantyCertificates || []).filter(warranty => warranty.isActive).length, 0
    );

    return {
      totalOperations,
      completedOperations,
      inProgressOperations,
      totalAmount,
      totalReceived,
      outstandingGuarantees,
      activeWarranties,
      totalDeductions,
      totalNetAmount
    };
  };

  // Operation handlers
  const handleSaveOperation = async (operation: Operation) => {
    try {
      await storage.addOperation(operation);
      const updatedOperations = await storage.getOperations();
      setOperations(updatedOperations);
      setActiveView('view-operations');
    } catch (error) {
      console.error('خطأ في حفظ العملية:', error);
      alert('حدث خطأ في حفظ العملية');
    }
  };

  const handleUpdateOperation = async (operation: Operation) => {
    try {
      // إعادة حساب الحالة بناءً على البيانات الجديدة
      const updatedStatus = calculateOperationStatus(operation);
      const operationWithUpdatedStatus = {
        ...operation,
        status: updatedStatus,
        updatedAt: new Date()
      };

      await storage.updateOperation(operation.id, operationWithUpdatedStatus);
      const updatedOperations = await storage.getOperations();
      setOperations(updatedOperations);
      setActiveView('view-operations');
      setSelectedOperation(null);
    } catch (error) {
      console.error('خطأ في تحديث العملية:', error);
      alert('حدث خطأ في تحديث العملية');
    }
  };

  const handleDeleteOperation = async (operationId: string) => {
    if (window.confirm('هل أنت متأكد من حذف هذه العملية؟')) {
      try {
        await storage.deleteOperation(operationId);
        const updatedOperations = await storage.getOperations();
        setOperations(updatedOperations);
      } catch (error) {
        console.error('خطأ في حذف العملية:', error);
        alert('حدث خطأ في حذف العملية');
      }
    }
  };

  const handleViewOperation = (operation: Operation) => {
    setSelectedOperation(operation);
    setActiveView('view-operation-details');
  };

  const handleEditOperation = (operation: Operation) => {
    setSelectedOperation(operation);
    setActiveView('edit-operation');
  };

  // Client handlers
  const handleAddClient = async (client: Client) => {
    try {
      await storage.addClient(client);
      const updatedClients = await storage.getClients();
      setClients(updatedClients);
    } catch (error) {
      console.error('خطأ في إضافة العميل:', error);
      alert('حدث خطأ في إضافة العميل');
    }
  };

  const handleUpdateClient = async (id: string, client: Client) => {
    try {
      await storage.updateClient(id, client);
      const updatedClients = await storage.getClients();
      setClients(updatedClients);
    } catch (error) {
      console.error('خطأ في تحديث العميل:', error);
      alert('حدث خطأ في تحديث العميل');
    }
  };

  const handleDeleteClient = async (id: string) => {
    try {
      // التحقق من وجود عمليات مرتبطة بالعميل
      const hasOperations = operations.some(op => op.clientId === id);
      if (hasOperations) {
        alert('لا يمكن حذف العميل لأنه مرتبط بعمليات موجودة');
        return;
      }

      if (window.confirm('هل أنت متأكد من حذف هذا العميل؟')) {
        await storage.deleteClient(id);
        const updatedClients = await storage.getClients();
        setClients(updatedClients);
      }
    } catch (error) {
      console.error('خطأ في حذف العميل:', error);
      alert('حدث خطأ في حذف العميل');
    }
  };

  // Export/Import handlers
  const handleExport = async () => {
    try {
      const data = await storage.exportData();
      const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `construction-data-${new Date().toISOString().split('T')[0]}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error('خطأ في تصدير البيانات:', error);
      alert('حدث خطأ في تصدير البيانات');
    }
  };

  const handleImport = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const data = JSON.parse(e.target?.result as string);
          await storage.importData(data);
          
          // إعادة تحميل البيانات
          const [updatedOperations, updatedClients] = await Promise.all([
            storage.getOperations(),
            storage.getClients()
          ]);
          
          setOperations(updatedOperations);
          setClients(updatedClients);
          
          alert('تم استيراد البيانات بنجاح');
        } catch (error) {
          console.error('خطأ في استيراد البيانات:', error);
          alert('خطأ في قراءة الملف');
        }
      };
      reader.readAsText(file);
    }
    event.target.value = '';
  };

  const renderContent = () => {
    if (isLoading) {
      return (
        <div className="flex items-center justify-center min-h-screen">
          <div className="text-center">
            <div className="animate-spin rounded-full h-32 w-32 border-b-2 border-blue-600 mx-auto"></div>
            <p className="mt-4 text-lg text-gray-600">جاري تحميل البيانات...</p>
          </div>
        </div>
      );
    }

    switch (activeView) {
      case 'dashboard':
        return <Dashboard stats={getDashboardStats()} onNavigate={handleNavigate} />;
      
      case 'add-operation':
        return (
          <AddOperationForm
            clients={clients}
            onSave={handleSaveOperation}
            onCancel={() => setActiveView('dashboard')}
          />
        );
      
      case 'view-operations':
        return (
          <OperationsList
            operations={operations}
            clients={clients}
            onEdit={handleEditOperation}
            onDelete={handleDeleteOperation}
            onView={handleViewOperation}
          />
        );

      case 'view-operation-details':
        if (!selectedOperation) return <Dashboard stats={getDashboardStats()} onNavigate={handleNavigate} />;
        const client = clients.find(c => c.id === selectedOperation.clientId);
        if (!client) return <Dashboard stats={getDashboardStats()} onNavigate={handleNavigate} />;
        return (
          <OperationDetails
            operation={selectedOperation}
            client={client}
            onClose={() => {
              setSelectedOperation(null);
              setActiveView('view-operations');
            }}
          />
        );

      case 'edit-operation':
        if (!selectedOperation) return <Dashboard stats={getDashboardStats()} onNavigate={handleNavigate} />;
        return (
          <EditOperationForm
            operation={selectedOperation}
            clients={clients}
            onSave={handleUpdateOperation}
            onCancel={() => {
              setSelectedOperation(null);
              setActiveView('view-operations');
            }}
          />
        );
      
      case 'clients':
        return (
          <ClientsManager
            clients={clients}
            onAddClient={handleAddClient}
            onUpdateClient={handleUpdateClient}
            onDeleteClient={handleDeleteClient}
          />
        );

      case 'guarantees':
        return (
          <GuaranteesManager
            operations={operations}
            clients={clients}
          />
        );

      case 'warranties':
        return (
          <WarrantiesManager
            operations={operations}
            clients={clients}
            onUpdateOperation={handleUpdateOperation}
          />
        );

      case 'checks':
        return (
          <ChecksManager
            operations={operations}
            clients={clients}
          />
        );

      case 'reports':
        return (
          <ReportsManager
            operations={operations}
            clients={clients}
          />
        );
      
      default:
        return <Dashboard stats={getDashboardStats()} onNavigate={handleNavigate} />;
    }
  };

  return (
    <div className="min-h-screen bg-gray-100" dir="rtl">
      <Header onExport={handleExport} onImport={handleImport} />
      
      <div className="flex">
        <Sidebar activeView={activeView} onViewChange={setActiveView} />
        
        <main className="flex-1">
          {renderContent()}
        </main>
      </div>
    </div>
  );
}

export default App;