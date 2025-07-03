import { Operation, Client } from '../types';
import { dbOperations } from './database';

// محول للتوافق مع النظام القديم
export const storage = {
  // العمليات
  getOperations: async (): Promise<Operation[]> => {
    try {
      return await dbOperations.getAllOperations();
    } catch (error) {
      console.error('Error loading operations:', error);
      return [];
    }
  },

  saveOperations: async (operations: Operation[]): Promise<void> => {
    // هذه الدالة لم تعد مستخدمة مع قاعدة البيانات الجديدة
    console.warn('saveOperations is deprecated. Use individual operation methods instead.');
  },

  addOperation: async (operation: Operation): Promise<void> => {
    try {
      await dbOperations.addOperation(operation);
    } catch (error) {
      console.error('Error adding operation:', error);
      throw error;
    }
  },

  updateOperation: async (id: string, updatedOperation: Operation): Promise<void> => {
    try {
      await dbOperations.updateOperation(id, updatedOperation);
    } catch (error) {
      console.error('Error updating operation:', error);
      throw error;
    }
  },

  deleteOperation: async (id: string): Promise<void> => {
    try {
      await dbOperations.deleteOperation(id);
    } catch (error) {
      console.error('Error deleting operation:', error);
      throw error;
    }
  },

  // العملاء
  getClients: async (): Promise<Client[]> => {
    try {
      return await dbOperations.getAllClients();
    } catch (error) {
      console.error('Error loading clients:', error);
      return [];
    }
  },

  saveClients: async (clients: Client[]): Promise<void> => {
    // هذه الدالة لم تعد مستخدمة مع قاعدة البيانات الجديدة
    console.warn('saveClients is deprecated. Use individual client methods instead.');
  },

  addClient: async (client: Client): Promise<void> => {
    try {
      await dbOperations.addClient(client);
    } catch (error) {
      console.error('Error adding client:', error);
      throw error;
    }
  },

  updateClient: async (id: string, updatedClient: Client): Promise<void> => {
    try {
      await dbOperations.updateClient(id, updatedClient);
    } catch (error) {
      console.error('Error updating client:', error);
      throw error;
    }
  },

  deleteClient: async (id: string): Promise<void> => {
    try {
      await dbOperations.deleteClient(id);
    } catch (error) {
      console.error('Error deleting client:', error);
      throw error;
    }
  },

  // تصدير واستيراد البيانات
  exportData: async () => {
    try {
      return await dbOperations.exportData();
    } catch (error) {
      console.error('Error exporting data:', error);
      throw error;
    }
  },

  importData: async (data: any) => {
    try {
      await dbOperations.importData(data);
    } catch (error) {
      console.error('Error importing data:', error);
      throw error;
    }
  }
};