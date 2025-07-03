import React from 'react';
import { 
  Building2, 
  CheckCircle, 
  Clock, 
  DollarSign, 
  AlertTriangle,
  TrendingUp,
  Plus,
  Eye,
  Users,
  Minus
} from 'lucide-react';
import StatsCard from './StatsCard';
import { DashboardStats } from '../../types';
import { formatCurrency } from '../../utils/calculations';

interface DashboardProps {
  stats: DashboardStats;
  onNavigate?: (view: string) => void;
}

const Dashboard: React.FC<DashboardProps> = ({ stats, onNavigate }) => {
  const completionRate = stats.totalOperations > 0 
    ? (stats.completedOperations / stats.totalOperations) * 100 
    : 0;

  const collectionRate = (stats.totalNetAmount || stats.totalAmount) > 0 
    ? (stats.totalReceived / (stats.totalNetAmount || stats.totalAmount)) * 100 
    : 0;

  const handleQuickAction = (view: string) => {
    if (onNavigate) {
      onNavigate(view);
    }
  };

  return (
    <div className="p-6">
      <div className="mb-8">
        <h2 className="text-2xl font-bold text-gray-900 mb-2">لوحة التحكم</h2>
        <p className="text-gray-600">نظرة عامة على أداء العمليات والمشاريع</p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6 mb-8">
        <StatsCard
          title="إجمالي العمليات"
          value={stats.totalOperations}
          icon={Building2}
          color="blue"
        />
        
        <StatsCard
          title="العمليات المكتملة"
          value={stats.completedOperations}
          icon={CheckCircle}
          color="green"
          trend={{ value: completionRate, isPositive: true }}
        />
        
        <StatsCard
          title="العمليات الجارية"
          value={stats.inProgressOperations}
          icon={Clock}
          color="yellow"
        />
        
        <StatsCard
          title="إجمالي القيمة"
          value={formatCurrency(stats.totalAmount)}
          icon={DollarSign}
          color="purple"
        />

        {stats.totalDeductions && stats.totalDeductions > 0 && (
          <StatsCard
            title="إجمالي الخصومات"
            value={formatCurrency(stats.totalDeductions)}
            icon={Minus}
            color="red"
          />
        )}

        {stats.totalNetAmount && (
          <StatsCard
            title="الصافي المستحق"
            value={formatCurrency(stats.totalNetAmount)}
            icon={TrendingUp}
            color="purple"
          />
        )}
        
        <StatsCard
          title="المبلغ المحصل"
          value={formatCurrency(stats.totalReceived)}
          icon={TrendingUp}
          color="green"
          trend={{ value: collectionRate, isPositive: true }}
        />
        
        <StatsCard
          title="الضمانات المعلقة"
          value={stats.outstandingGuarantees}
          icon={AlertTriangle}
          color="red"
        />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white p-6 rounded-lg border border-gray-200">
          <h3 className="text-lg font-semibold text-gray-900 mb-4">معدل الإنجاز</h3>
          <div className="space-y-4">
            <div>
              <div className="flex justify-between items-center mb-2">
                <span className="text-sm font-medium text-gray-600">معدل إكمال العمليات</span>
                <span className="text-sm font-bold text-gray-900">{completionRate.toFixed(1)}%</span>
              </div>
              <div className="w-full bg-gray-200 rounded-full h-2">
                <div 
                  className="bg-blue-600 h-2 rounded-full transition-all duration-300"
                  style={{ width: `${completionRate}%` }}
                ></div>
              </div>
            </div>
            
            <div>
              <div className="flex justify-between items-center mb-2">
                <span className="text-sm font-medium text-gray-600">معدل التحصيل</span>
                <span className="text-sm font-bold text-gray-900">{collectionRate.toFixed(1)}%</span>
              </div>
              <div className="w-full bg-gray-200 rounded-full h-2">
                <div 
                  className="bg-green-600 h-2 rounded-full transition-all duration-300"
                  style={{ width: `${collectionRate}%` }}
                ></div>
              </div>
            </div>

            {stats.totalDeductions && stats.totalDeductions > 0 && (
              <div>
                <div className="flex justify-between items-center mb-2">
                  <span className="text-sm font-medium text-gray-600">نسبة الخصومات</span>
                  <span className="text-sm font-bold text-gray-900">
                    {stats.totalAmount > 0 ? ((stats.totalDeductions / stats.totalAmount) * 100).toFixed(1) : 0}%
                  </span>
                </div>
                <div className="w-full bg-gray-200 rounded-full h-2">
                  <div 
                    className="bg-red-600 h-2 rounded-full transition-all duration-300"
                    style={{ width: `${stats.totalAmount > 0 ? (stats.totalDeductions / stats.totalAmount) * 100 : 0}%` }}
                  ></div>
                </div>
              </div>
            )}
          </div>
        </div>

        <div className="bg-white p-6 rounded-lg border border-gray-200">
          <h3 className="text-lg font-semibold text-gray-900 mb-4">الإجراءات السريعة</h3>
          <div className="space-y-3">
            <button 
              onClick={() => handleQuickAction('add-operation')}
              className="w-full flex items-center justify-center gap-2 px-4 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
            >
              <Plus className="w-4 h-4" />
              إضافة عملية جديدة
            </button>
            
            <button 
              onClick={() => handleQuickAction('view-operations')}
              className="w-full flex items-center justify-center gap-2 px-4 py-3 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors"
            >
              <Eye className="w-4 h-4" />
              استعراض العمليات
            </button>
            
            <button 
              onClick={() => handleQuickAction('clients')}
              className="w-full flex items-center justify-center gap-2 px-4 py-3 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition-colors"
            >
              <Users className="w-4 h-4" />
              إدارة العملاء
            </button>

            <button 
              onClick={() => handleQuickAction('guarantees')}
              className="w-full flex items-center justify-center gap-2 px-4 py-3 bg-orange-600 text-white rounded-lg hover:bg-orange-700 transition-colors"
            >
              <AlertTriangle className="w-4 h-4" />
              إدارة الضمانات
            </button>

            <button 
              onClick={() => handleQuickAction('checks')}
              className="w-full flex items-center justify-center gap-2 px-4 py-3 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition-colors"
            >
              <CheckCircle className="w-4 h-4" />
              الشيكات والمدفوعات
            </button>

            <button 
              onClick={() => handleQuickAction('reports')}
              className="w-full flex items-center justify-center gap-2 px-4 py-3 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition-colors"
            >
              <TrendingUp className="w-4 h-4" />
              التقارير والإحصائيات
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default Dashboard;