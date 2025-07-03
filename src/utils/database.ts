import Dexie, { Table } from 'dexie';
import { Operation, Client } from '../types';

// قاعدة البيانات المدمجة
export class ConstructionDatabase extends Dexie {
  operations!: Table<Operation>;
  clients!: Table<Client>;
  settings!: Table<{ key: string; value: any }>;

  constructor() {
    super('ConstructionManagementDB');
    
    // تحديث إصدار قاعدة البيانات لدعم الحقول الجديدة
    this.version(2).stores({
      operations: 'id, code, name, clientId, status, createdAt, updatedAt, totalAmount, totalReceived',
      clients: 'id, name, type, phone, email, createdAt, updatedAt',
      settings: 'key'
    }).upgrade(trans => {
      // ترقية البيانات الموجودة
      return trans.table('clients').toCollection().modify(client => {
        if (!client.type) {
          client.type = 'owner';
        }
        if (!client.contacts) {
          client.contacts = [];
        }
      });
    });

    // الاحتفاظ بالإصدار الأول للتوافق
    this.version(1).stores({
      operations: 'id, code, name, clientId, status, createdAt, updatedAt, totalAmount, totalReceived',
      clients: 'id, name, phone, email, createdAt, updatedAt',
      settings: 'key'
    });

    // إضافة hooks للتحديث التلقائي للتواريخ
    this.operations.hook('creating', function (primKey, obj, trans) {
      obj.createdAt = new Date();
      obj.updatedAt = new Date();
      
      // إضافة الحقول الجديدة إذا لم تكن موجودة
      if (!obj.warrantyCertificates) {
        obj.warrantyCertificates = [];
      }
      
      // تحديث relatedTo للضمانات الموجودة
      obj.guaranteeChecks = obj.guaranteeChecks.map(check => ({
        ...check,
        relatedTo: check.relatedTo || 'operation'
      }));
      
      obj.guaranteeLetters = obj.guaranteeLetters.map(letter => ({
        ...letter,
        relatedTo: letter.relatedTo || 'operation'
      }));
    });

    this.operations.hook('updating', function (modifications, primKey, obj, trans) {
      modifications.updatedAt = new Date();
    });

    this.clients.hook('creating', function (primKey, obj, trans) {
      obj.createdAt = new Date();
      obj.updatedAt = new Date();
      
      // إضافة الحقول الجديدة إذا لم تكن موجودة
      if (!obj.type) {
        obj.type = 'owner';
      }
      if (!obj.contacts) {
        obj.contacts = [];
      }
    });

    this.clients.hook('updating', function (modifications, primKey, obj, trans) {
      modifications.updatedAt = new Date();
    });
  }
}

// إنشاء مثيل من قاعدة البيانات
export const db = new ConstructionDatabase();

// وظائف مساعدة لإدارة قاعدة البيانات
export const dbOperations = {
  // العمليات
  async getAllOperations(): Promise<Operation[]> {
    return await db.operations.orderBy('createdAt').reverse().toArray();
  },

  async getOperationById(id: string): Promise<Operation | undefined> {
    return await db.operations.get(id);
  },

  async addOperation(operation: Operation): Promise<string> {
    return await db.operations.add(operation);
  },

  async updateOperation(id: string, changes: Partial<Operation>): Promise<number> {
    return await db.operations.update(id, changes);
  },

  async deleteOperation(id: string): Promise<void> {
    await db.operations.delete(id);
  },

  async getOperationsByClient(clientId: string): Promise<Operation[]> {
    return await db.operations.where('clientId').equals(clientId).toArray();
  },

  async getOperationsByStatus(status: Operation['status']): Promise<Operation[]> {
    return await db.operations.where('status').equals(status).toArray();
  },

  async getOperationsByDateRange(startDate: Date, endDate: Date): Promise<Operation[]> {
    return await db.operations
      .where('createdAt')
      .between(startDate, endDate, true, true)
      .toArray();
  },

  // العملاء
  async getAllClients(): Promise<Client[]> {
    return await db.clients.orderBy('name').toArray();
  },

  async getClientById(id: string): Promise<Client | undefined> {
    return await db.clients.get(id);
  },

  async addClient(client: Client): Promise<string> {
    return await db.clients.add(client);
  },

  async updateClient(id: string, changes: Partial<Client>): Promise<number> {
    return await db.clients.update(id, changes);
  },

  async deleteClient(id: string): Promise<void> {
    // التحقق من وجود عمليات مرتبطة بالعميل
    const operations = await this.getOperationsByClient(id);
    if (operations.length > 0) {
      throw new Error('لا يمكن حذف العميل لأنه مرتبط بعمليات موجودة');
    }
    await db.clients.delete(id);
  },

  async searchClients(searchTerm: string): Promise<Client[]> {
    return await db.clients
      .filter(client => 
        client.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
        (client.phone && client.phone.includes(searchTerm)) ||
        (client.email && client.email.toLowerCase().includes(searchTerm.toLowerCase()))
      )
      .toArray();
  },

  // الإعدادات
  async getSetting(key: string): Promise<any> {
    const setting = await db.settings.get(key);
    return setting?.value;
  },

  async setSetting(key: string, value: any): Promise<string> {
    return await db.settings.put({ key, value });
  },

  // النسخ الاحتياطي والاستيراد
  async exportData(): Promise<{ operations: Operation[]; clients: Client[]; exportDate: string; version: string }> {
    const operations = await this.getAllOperations();
    const clients = await this.getAllClients();
    
    return {
      operations,
      clients,
      exportDate: new Date().toISOString(),
      version: '2.0' // إصدار قاعدة البيانات
    };
  },

  async importData(data: { operations?: Operation[]; clients?: Client[]; version?: string }): Promise<void> {
    await db.transaction('rw', db.operations, db.clients, async () => {
      if (data.clients) {
        // تحديث بيانات العملاء للتوافق مع الإصدار الجديد
        const updatedClients = data.clients.map(client => ({
          ...client,
          type: client.type || 'owner',
          contacts: client.contacts || []
        }));
        
        await db.clients.clear();
        await db.clients.bulkAdd(updatedClients);
      }
      
      if (data.operations) {
        // تحديث بيانات العمليات للتوافق مع الإصدار الجديد
        const updatedOperations = data.operations.map(operation => ({
          ...operation,
          warrantyCertificates: operation.warrantyCertificates || [],
          guaranteeChecks: operation.guaranteeChecks.map(check => ({
            ...check,
            relatedTo: check.relatedTo || 'operation'
          })),
          guaranteeLetters: operation.guaranteeLetters.map(letter => ({
            ...letter,
            relatedTo: letter.relatedTo || 'operation'
          }))
        }));
        
        await db.operations.clear();
        await db.operations.bulkAdd(updatedOperations);
      }
    });
  },

  // إحصائيات سريعة
  async getStatistics() {
    const [operations, clients] = await Promise.all([
      this.getAllOperations(),
      this.getAllClients()
    ]);

    const totalOperations = operations.length;
    const completedOperations = operations.filter(op => op.status === 'completed').length;
    const inProgressOperations = operations.filter(op => op.status === 'in_progress').length;
    const totalAmount = operations.reduce((sum, op) => sum + op.totalAmount, 0);
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
      totalClients: clients.length
    };
  },

  // البحث المتقدم
  async advancedSearch(filters: {
    clientId?: string;
    status?: Operation['status'];
    startDate?: Date;
    endDate?: Date;
    minAmount?: number;
    maxAmount?: number;
    searchTerm?: string;
  }): Promise<Operation[]> {
    let query = db.operations.toCollection();

    if (filters.clientId) {
      query = query.filter(op => op.clientId === filters.clientId);
    }

    if (filters.status) {
      query = query.filter(op => op.status === filters.status);
    }

    if (filters.startDate) {
      query = query.filter(op => op.createdAt >= filters.startDate!);
    }

    if (filters.endDate) {
      query = query.filter(op => op.createdAt <= filters.endDate!);
    }

    if (filters.minAmount !== undefined) {
      query = query.filter(op => op.totalAmount >= filters.minAmount!);
    }

    if (filters.maxAmount !== undefined) {
      query = query.filter(op => op.totalAmount <= filters.maxAmount!);
    }

    if (filters.searchTerm) {
      const term = filters.searchTerm.toLowerCase();
      query = query.filter(op => 
        op.name.toLowerCase().includes(term) ||
        op.code.toLowerCase().includes(term)
      );
    }

    return await query.toArray();
  }
};

// تهيئة قاعدة البيانات مع بيانات تجريبية
export const initializeDatabase = async () => {
  try {
    await db.open();
    
    // التحقق من وجود بيانات
    const clientsCount = await db.clients.count();
    
    if (clientsCount === 0) {
      // إضافة عملاء تجريبيين
      const sampleClients: Client[] = [
        {
          id: crypto.randomUUID(),
          name: 'شركة الإنشاءات الحديثة',
          type: 'main_contractor',
          phone: '01234567890',
          email: 'info@modern-construction.com',
          address: 'القاهرة، مصر',
          contacts: [
            {
              id: crypto.randomUUID(),
              name: 'أحمد محمد',
              position: 'مدير الحسابات',
              department: 'accounts',
              phone: '01234567891',
              email: 'ahmed@modern-construction.com',
              isMainContact: true
            },
            {
              id: crypto.randomUUID(),
              name: 'سارة أحمد',
              position: 'مهندسة مشروعات',
              department: 'engineering',
              phone: '01234567892',
              email: 'sara@modern-construction.com',
              isMainContact: false
            }
          ],
          createdAt: new Date(),
          updatedAt: new Date()
        },
        {
          id: crypto.randomUUID(),
          name: 'مؤسسة البناء المتطور',
          type: 'owner',
          phone: '01987654321',
          email: 'contact@advanced-building.com',
          address: 'الإسكندرية، مصر',
          contacts: [
            {
              id: crypto.randomUUID(),
              name: 'محمد علي',
              position: 'المدير التنفيذي',
              department: 'management',
              phone: '01987654322',
              email: 'mohamed@advanced-building.com',
              isMainContact: true
            }
          ],
          createdAt: new Date(),
          updatedAt: new Date()
        },
        {
          id: crypto.randomUUID(),
          name: 'شركة المقاولات العربية',
          type: 'consultant',
          phone: '01555666777',
          email: 'info@arab-contracting.com',
          address: 'الجيزة، مصر',
          contacts: [
            {
              id: crypto.randomUUID(),
              name: 'فاطمة حسن',
              position: 'مديرة المشاريع',
              department: 'engineering',
              phone: '01555666778',
              email: 'fatma@arab-contracting.com',
              isMainContact: true
            }
          ],
          createdAt: new Date(),
          updatedAt: new Date()
        }
      ];

      await db.clients.bulkAdd(sampleClients);
      console.log('تم إضافة العملاء التجريبيين بنجاح');
    }

    console.log('تم تهيئة قاعدة البيانات بنجاح');
  } catch (error) {
    console.error('خطأ في تهيئة قاعدة البيانات:', error);
  }
};