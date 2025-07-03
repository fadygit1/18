import React from 'react';
import { 
  Home, 
  Plus, 
  Eye, 
  Users, 
  Shield, 
  CheckSquare,
  BarChart3,
  Award
} from 'lucide-react';

interface SidebarProps {
  activeView: string;
  onViewChange: (view: string) => void;
}

const Sidebar: React.FC<SidebarProps> = ({ activeView, onViewChange }) => {
  const menuItems = [
    { id: 'dashboard', label: 'الرئيسية', icon: Home },
    { id: 'add-operation', label: 'إضافة عملية', icon: Plus },
    { id: 'view-operations', label: 'استعراض العمليات', icon: Eye },
    { id: 'clients', label: 'إدارة العملاء', icon: Users },
    { id: 'guarantees', label: 'الضمانات', icon: Shield },
    { id: 'warranties', label: 'شهادات الضمان', icon: Award },
    { id: 'checks', label: 'الشيكات والمدفوعات', icon: CheckSquare },
    { id: 'reports', label: 'التقارير', icon: BarChart3 },
  ];

  return (
    <aside className="w-64 bg-gray-50 border-r border-gray-200 min-h-screen">
      <nav className="p-4">
        <ul className="space-y-2">
          {menuItems.map((item) => {
            const Icon = item.icon;
            const isActive = activeView === item.id;
            
            return (
              <li key={item.id}>
                <button
                  onClick={() => onViewChange(item.id)}
                  className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg text-right transition-colors ${
                    isActive
                      ? 'bg-blue-100 text-blue-700 border border-blue-200'
                      : 'text-gray-700 hover:bg-gray-100'
                  }`}
                >
                  <Icon className="w-5 h-5" />
                  <span className="font-medium">{item.label}</span>
                </button>
              </li>
            );
          })}
        </ul>
      </nav>
    </aside>
  );
};

export default Sidebar;