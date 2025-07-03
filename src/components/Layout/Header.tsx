import React, { useState } from 'react';
import { Building2, Settings, Download, Upload, Database } from 'lucide-react';

interface HeaderProps {
  onExport: () => void;
  onImport: (event: React.ChangeEvent<HTMLInputElement>) => void;
}

const Header: React.FC<HeaderProps> = ({ onExport, onImport }) => {
  const [showSettings, setShowSettings] = useState(false);

  const handleSettingsClick = () => {
    setShowSettings(!showSettings);
  };

  return (
    <header className="bg-white border-b border-gray-200 px-6 py-4 relative">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="flex items-center justify-center w-10 h-10 bg-blue-600 rounded-lg">
            <Building2 className="w-6 h-6 text-white" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-gray-900">نظام إدارة العمليات الإنشائية</h1>
            <div className="flex items-center gap-2">
              <p className="text-sm text-gray-600">إدارة شاملة لمشاريع البناء والمقاولات</p>
              <div className="flex items-center gap-1 px-2 py-1 bg-green-100 text-green-800 rounded-full text-xs">
                <Database className="w-3 h-3" />
                قاعدة بيانات مدمجة
              </div>
            </div>
          </div>
        </div>
        
        <div className="flex items-center gap-3">
          <button
            onClick={onExport}
            className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors"
          >
            <Download className="w-4 h-4" />
            تصدير البيانات
          </button>
          
          <label className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors cursor-pointer">
            <Upload className="w-4 h-4" />
            استيراد البيانات
            <input
              type="file"
              accept=".json"
              onChange={onImport}
              className="hidden"
            />
          </label>
          
          <div className="relative">
            <button 
              onClick={handleSettingsClick}
              className="flex items-center justify-center w-10 h-10 bg-gray-100 rounded-lg hover:bg-gray-200 transition-colors"
            >
              <Settings className="w-5 h-5 text-gray-600" />
            </button>
            
            {showSettings && (
              <div className="absolute left-0 top-12 w-64 bg-white border border-gray-200 rounded-lg shadow-lg z-50">
                <div className="p-4">
                  <h3 className="text-lg font-semibold text-gray-900 mb-4">الإعدادات</h3>
                  <div className="space-y-3">
                    <div className="flex items-center justify-between">
                      <span className="text-sm text-gray-700">الإشعارات</span>
                      <input type="checkbox" className="rounded" defaultChecked />
                    </div>
                    <div className="flex items-center justify-between">
                      <span className="text-sm text-gray-700">النسخ الاحتياطي التلقائي</span>
                      <input type="checkbox" className="rounded" defaultChecked />
                    </div>
                    <div className="flex items-center justify-between">
                      <span className="text-sm text-gray-700">الوضع المظلم</span>
                      <input type="checkbox" className="rounded" />
                    </div>
                    <hr className="my-3" />
                    <button className="w-full text-left text-sm text-gray-700 hover:text-gray-900 py-2">
                      حول النظام
                    </button>
                    <button className="w-full text-left text-sm text-gray-700 hover:text-gray-900 py-2">
                      المساعدة
                    </button>
                    <button className="w-full text-left text-sm text-red-600 hover:text-red-800 py-2">
                      مسح جميع البيانات
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
      
      {/* Overlay to close settings */}
      {showSettings && (
        <div 
          className="fixed inset-0 z-40" 
          onClick={() => setShowSettings(false)}
        />
      )}
    </header>
  );
};

export default Header;