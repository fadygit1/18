import React, { useState } from 'react';
import { Plus, Edit, Trash2, Search, Phone, Mail, MapPin, User, Building } from 'lucide-react';
import { Client, ClientContact } from '../../types';

interface ClientsManagerProps {
  clients: Client[];
  onAddClient: (client: Client) => void;
  onUpdateClient: (id: string, client: Client) => void;
  onDeleteClient: (id: string) => void;
}

const ClientsManager: React.FC<ClientsManagerProps> = ({
  clients,
  onAddClient,
  onUpdateClient,
  onDeleteClient
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [showAddForm, setShowAddForm] = useState(false);
  const [editingClient, setEditingClient] = useState<Client | null>(null);
  const [activeTab, setActiveTab] = useState<'basic' | 'contacts'>('basic');
  
  const [formData, setFormData] = useState({
    name: '',
    type: 'owner' as 'owner' | 'main_contractor' | 'consultant',
    phone: '',
    email: '',
    address: ''
  });

  const [contacts, setContacts] = useState<ClientContact[]>([]);

  const filteredClients = clients.filter(client =>
    client.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
    (client.phone && client.phone.includes(searchTerm)) ||
    (client.email && client.email.toLowerCase().includes(searchTerm.toLowerCase()))
  );

  const resetForm = () => {
    setFormData({ name: '', type: 'owner', phone: '', email: '', address: '' });
    setContacts([]);
    setShowAddForm(false);
    setEditingClient(null);
    setActiveTab('basic');
  };

  const addContact = () => {
    const newContact: ClientContact = {
      id: crypto.randomUUID(),
      name: '',
      position: '',
      department: 'management',
      phone: '',
      email: '',
      isMainContact: contacts.length === 0
    };
    setContacts([...contacts, newContact]);
  };

  const updateContact = (index: number, field: keyof ClientContact, value: any) => {
    const updatedContacts = [...contacts];
    updatedContacts[index] = { ...updatedContacts[index], [field]: value };
    
    // إذا تم تحديد هذا الشخص كجهة اتصال رئيسية، قم بإلغاء الآخرين
    if (field === 'isMainContact' && value === true) {
      updatedContacts.forEach((contact, i) => {
        if (i !== index) {
          contact.isMainContact = false;
        }
      });
    }
    
    setContacts(updatedContacts);
  };

  const removeContact = (index: number) => {
    const newContacts = contacts.filter((_, i) => i !== index);
    // إذا تم حذف جهة الاتصال الرئيسية، اجعل الأولى رئيسية
    if (contacts[index].isMainContact && newContacts.length > 0) {
      newContacts[0].isMainContact = true;
    }
    setContacts(newContacts);
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    
    if (editingClient) {
      const updatedClient: Client = {
        ...editingClient,
        ...formData,
        contacts: contacts,
        updatedAt: new Date()
      };
      onUpdateClient(editingClient.id, updatedClient);
    } else {
      const newClient: Client = {
        id: crypto.randomUUID(),
        ...formData,
        contacts: contacts,
        createdAt: new Date(),
        updatedAt: new Date()
      };
      onAddClient(newClient);
    }
    
    resetForm();
  };

  const handleEdit = (client: Client) => {
    setEditingClient(client);
    setFormData({
      name: client.name,
      type: client.type,
      phone: client.phone || '',
      email: client.email || '',
      address: client.address || ''
    });
    setContacts(client.contacts || []);
    setShowAddForm(true);
  };

  const getClientTypeLabel = (type: string) => {
    const types = {
      'owner': 'مالك',
      'main_contractor': 'مقاول رئيسي',
      'consultant': 'استشاري'
    };
    return types[type as keyof typeof types] || type;
  };

  const getDepartmentLabel = (department: string) => {
    const departments = {
      'accounts': 'الحسابات',
      'engineering': 'الهندسة',
      'management': 'الإدارة',
      'other': 'أخرى'
    };
    return departments[department as keyof typeof departments] || department;
  };

  return (
    <div className="p-6">
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-gray-900 mb-2">إدارة العملاء</h2>
        <p className="text-gray-600">إدارة قاعدة بيانات العملاء وجهات الاتصال</p>
      </div>

      {/* Header Actions */}
      <div className="mb-6 flex flex-col sm:flex-row gap-4 justify-between">
        <div className="relative flex-1 max-w-md">
          <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 w-4 h-4" />
          <input
            type="text"
            placeholder="البحث عن عميل..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full pl-10 pr-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
        </div>
        
        <button
          onClick={() => setShowAddForm(true)}
          className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors"
        >
          <Plus className="w-4 h-4" />
          إضافة عميل جديد
        </button>
      </div>

      {/* Add/Edit Form */}
      {showAddForm && (
        <div className="mb-6 bg-white p-6 rounded-lg border border-gray-200">
          <h3 className="text-lg font-semibold text-gray-900 mb-4">
            {editingClient ? 'تعديل العميل' : 'إضافة عميل جديد'}
          </h3>

          {/* Tabs */}
          <div className="mb-6 border-b border-gray-200">
            <nav className="flex gap-8">
              <button
                onClick={() => setActiveTab('basic')}
                className={`py-2 px-1 border-b-2 font-medium text-sm ${
                  activeTab === 'basic'
                    ? 'border-blue-500 text-blue-600'
                    : 'border-transparent text-gray-500 hover:text-gray-700'
                }`}
              >
                البيانات الأساسية
              </button>
              <button
                onClick={() => setActiveTab('contacts')}
                className={`py-2 px-1 border-b-2 font-medium text-sm ${
                  activeTab === 'contacts'
                    ? 'border-blue-500 text-blue-600'
                    : 'border-transparent text-gray-500 hover:text-gray-700'
                }`}
              >
                جهات الاتصال ({contacts.length})
              </button>
            </nav>
          </div>
          
          <form onSubmit={handleSubmit} className="space-y-6">
            {/* Basic Information Tab */}
            {activeTab === 'basic' && (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    اسم العميل *
                  </label>
                  <input
                    type="text"
                    value={formData.name}
                    onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                    className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                    required
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    نوع العميل *
                  </label>
                  <select
                    value={formData.type}
                    onChange={(e) => setFormData({ ...formData, type: e.target.value as any })}
                    className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                    required
                  >
                    <option value="owner">مالك</option>
                    <option value="main_contractor">مقاول رئيسي</option>
                    <option value="consultant">استشاري</option>
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    رقم الهاتف
                  </label>
                  <input
                    type="tel"
                    value={formData.phone}
                    onChange={(e) => setFormData({ ...formData, phone: e.target.value })}
                    className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    البريد الإلكتروني
                  </label>
                  <input
                    type="email"
                    value={formData.email}
                    onChange={(e) => setFormData({ ...formData, email: e.target.value })}
                    className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  />
                </div>

                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    العنوان
                  </label>
                  <input
                    type="text"
                    value={formData.address}
                    onChange={(e) => setFormData({ ...formData, address: e.target.value })}
                    className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                  />
                </div>
              </div>
            )}

            {/* Contacts Tab */}
            {activeTab === 'contacts' && (
              <div className="space-y-6">
                <div className="flex justify-between items-center">
                  <h4 className="text-md font-medium text-gray-900">جهات الاتصال</h4>
                  <button
                    type="button"
                    onClick={addContact}
                    className="flex items-center gap-2 px-3 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 transition-colors"
                  >
                    <Plus className="w-4 h-4" />
                    إضافة جهة اتصال
                  </button>
                </div>

                {contacts.map((contact, index) => (
                  <div key={contact.id} className="p-4 border border-gray-200 rounded-lg bg-gray-50">
                    <div className="flex justify-between items-center mb-4">
                      <h5 className="font-medium text-gray-900">جهة الاتصال {index + 1}</h5>
                      <button
                        type="button"
                        onClick={() => removeContact(index)}
                        className="text-red-600 hover:text-red-800"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">
                          الاسم *
                        </label>
                        <input
                          type="text"
                          value={contact.name}
                          onChange={(e) => updateContact(index, 'name', e.target.value)}
                          className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                          required
                        />
                      </div>

                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">
                          المنصب
                        </label>
                        <input
                          type="text"
                          value={contact.position}
                          onChange={(e) => updateContact(index, 'position', e.target.value)}
                          className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                        />
                      </div>

                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">
                          القسم
                        </label>
                        <select
                          value={contact.department}
                          onChange={(e) => updateContact(index, 'department', e.target.value)}
                          className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                        >
                          <option value="management">الإدارة</option>
                          <option value="accounts">الحسابات</option>
                          <option value="engineering">الهندسة</option>
                          <option value="other">أخرى</option>
                        </select>
                      </div>

                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">
                          رقم الهاتف
                        </label>
                        <input
                          type="tel"
                          value={contact.phone}
                          onChange={(e) => updateContact(index, 'phone', e.target.value)}
                          className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                        />
                      </div>

                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">
                          البريد الإلكتروني
                        </label>
                        <input
                          type="email"
                          value={contact.email}
                          onChange={(e) => updateContact(index, 'email', e.target.value)}
                          className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                        />
                      </div>

                      <div className="flex items-center">
                        <label className="flex items-center">
                          <input
                            type="checkbox"
                            checked={contact.isMainContact}
                            onChange={(e) => updateContact(index, 'isMainContact', e.target.checked)}
                            className="mr-2 h-4 w-4 text-blue-600 focus:ring-blue-500 border-gray-300 rounded"
                          />
                          <span className="text-sm text-gray-700">جهة الاتصال الرئيسية</span>
                        </label>
                      </div>
                    </div>
                  </div>
                ))}

                {contacts.length === 0 && (
                  <div className="text-center py-8 text-gray-500">
                    لا توجد جهات اتصال مضافة. اضغط "إضافة جهة اتصال" لإضافة جهة اتصال جديدة.
                  </div>
                )}
              </div>
            )}

            <div className="flex gap-3 pt-4 border-t">
              <button
                type="submit"
                className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors"
              >
                {editingClient ? 'تحديث' : 'إضافة'}
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

      {/* Clients Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {filteredClients.map((client) => {
          const mainContact = client.contacts?.find(contact => contact.isMainContact);
          
          return (
            <div key={client.id} className="bg-white p-6 rounded-lg border border-gray-200 hover:shadow-md transition-shadow">
              <div className="flex justify-between items-start mb-4">
                <div>
                  <h3 className="text-lg font-semibold text-gray-900">{client.name}</h3>
                  <span className="inline-flex items-center px-2 py-1 rounded-full text-xs font-medium bg-blue-100 text-blue-800">
                    {getClientTypeLabel(client.type)}
                  </span>
                </div>
                <div className="flex gap-2">
                  <button
                    onClick={() => handleEdit(client)}
                    className="text-blue-600 hover:text-blue-800"
                    title="تعديل"
                  >
                    <Edit className="w-4 h-4" />
                  </button>
                  <button
                    onClick={() => onDeleteClient(client.id)}
                    className="text-red-600 hover:text-red-800"
                    title="حذف"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              </div>

              <div className="space-y-2">
                {client.phone && (
                  <div className="flex items-center gap-2 text-gray-600">
                    <Phone className="w-4 h-4" />
                    <span className="text-sm">{client.phone}</span>
                  </div>
                )}
                
                {client.email && (
                  <div className="flex items-center gap-2 text-gray-600">
                    <Mail className="w-4 h-4" />
                    <span className="text-sm">{client.email}</span>
                  </div>
                )}
                
                {client.address && (
                  <div className="flex items-center gap-2 text-gray-600">
                    <MapPin className="w-4 h-4" />
                    <span className="text-sm">{client.address}</span>
                  </div>
                )}
              </div>

              {/* Main Contact */}
              {mainContact && (
                <div className="mt-4 pt-4 border-t border-gray-200">
                  <div className="flex items-center gap-2 text-gray-700 mb-2">
                    <User className="w-4 h-4" />
                    <span className="text-sm font-medium">جهة الاتصال الرئيسية:</span>
                  </div>
                  <div className="text-sm text-gray-600">
                    <div className="font-medium">{mainContact.name}</div>
                    {mainContact.position && (
                      <div className="text-xs">{mainContact.position} - {getDepartmentLabel(mainContact.department)}</div>
                    )}
                    {mainContact.phone && (
                      <div className="text-xs">{mainContact.phone}</div>
                    )}
                  </div>
                </div>
              )}

              {/* Contacts Count */}
              <div className="mt-4 pt-4 border-t border-gray-200">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2 text-gray-600">
                    <Building className="w-4 h-4" />
                    <span className="text-sm">جهات الاتصال: {client.contacts?.length || 0}</span>
                  </div>
                  <p className="text-xs text-gray-500">
                    تم الإنشاء: {new Date(client.createdAt).toLocaleDateString('ar-EG')}
                  </p>
                </div>
              </div>
            </div>
          );
        })}
      </div>

      {filteredClients.length === 0 && (
        <div className="text-center py-12">
          <p className="text-gray-500">
            {searchTerm ? 'لا توجد عملاء تطابق البحث' : 'لا توجد عملاء مسجلين'}
          </p>
        </div>
      )}
    </div>
  );
};

export default ClientsManager;