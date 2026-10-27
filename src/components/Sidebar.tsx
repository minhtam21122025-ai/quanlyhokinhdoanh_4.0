import React from 'react';
import {
  Home,
  Users,
  BookOpen,
  DollarSign,
  Building2,
  GraduationCap,
  Calculator,
  LogOut,
  ShieldCheck,
  X,
} from 'lucide-react';
import { cn } from '../lib/utils';
import { motion, AnimatePresence } from 'motion/react';
import Logo from './Logo';
import { PWAInstallButton } from './MobilePWASupport';

interface SidebarProps {
  activeTab: string;
  setActiveTab: (tab: string) => void;
  currentUser: any;
  onLogout: () => void;
  isOpen?: boolean;
  onClose?: () => void;
  menuOrder?: string[];
}

const Sidebar: React.FC<SidebarProps> = ({
  activeTab,
  setActiveTab,
  currentUser,
  onLogout,
  isOpen,
  onClose,
  menuOrder,
}) => {
  const defaultItems = [
    { id: 'dashboard', label: 'TRANG CHỦ', icon: Home },
    { id: 'business', label: 'HỘ KINH DOANH', icon: Building2 },
    { id: 'students_group', label: 'QUẢN LÝ HỌC SINH', icon: Users },
    { id: 'student_learning_group', label: 'QUẢN LÝ HỌC TẬP HỌC SINH', icon: GraduationCap },
    { id: 'program', label: 'QUẢN LÝ CHƯƠNG TRÌNH DẠY', icon: BookOpen },
    { id: 'finance_group', label: 'QUẢN LÝ TÀI CHÍNH', icon: DollarSign },
    { id: 'tax_report', label: 'BÁO CÁO THUẾ CHO HKD', icon: Calculator },
    { id: 'users', label: 'QUẢN LÝ TÀI KHOẢN', icon: ShieldCheck, adminOnly: true },
  ];

  const menuItems = React.useMemo(() => {
    const items = [...defaultItems];

    if (menuOrder && menuOrder.length > 0) {
      items.sort((a, b) => {
        const indexA = menuOrder.indexOf(a.id);
        const indexB = menuOrder.indexOf(b.id);
        if (indexA === -1 && indexB === -1) return 0;
        if (indexA === -1) return 1;
        if (indexB === -1) return -1;
        return indexA - indexB;
      });
    }

    return items.filter((item) => {
      if (currentUser?.role === 'admin') return true;
      return !item.adminOnly;
    });
  }, [currentUser, menuOrder]);

  return (
    <>
      <AnimatePresence>
        {isOpen && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={onClose}
            className="fixed inset-0 bg-slate-950/60 backdrop-blur-xs z-[90] lg:hidden"
          />
        )}
      </AnimatePresence>

      <aside
        className={cn(
          'fixed lg:relative inset-y-0 left-0 w-72 bg-gradient-to-b from-emerald-900 via-teal-800 to-teal-900 text-white border-r border-teal-700/60 shadow-xl z-[100] flex flex-col transition-transform duration-300 lg:translate-x-0',
          isOpen ? 'translate-x-0' : '-translate-x-full'
        )}
      >
        {/* Brand Header */}
        <div className="px-6 py-6 border-b border-white/15 flex items-center justify-between">
          <div className="flex items-center gap-3.5 min-w-0">
            <div className="h-11 w-11 flex items-center justify-center bg-white/10 rounded-xl p-1 border border-white/20 shrink-0">
              <Logo className="h-full w-full" />
            </div>
            <div className="min-w-0">
              <h1 className="text-lg font-extrabold tracking-tight text-white leading-none">
                HOÀNG GIA
              </h1>
              <p className="text-[10px] font-semibold text-teal-100/90 tracking-wider uppercase mt-1 truncate">
                Trao cơ hội · Nhận niềm tin
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="lg:hidden min-w-[40px] min-h-[40px] flex items-center justify-center rounded-xl text-teal-100 hover:bg-white/15 transition-colors"
            aria-label="Đóng menu"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Navigation Menu */}
        <div className="flex-1 overflow-y-auto px-3.5 py-5 space-y-1.5 custom-scrollbar">
          {menuItems.map((item) => {
            const isActive =
              activeTab === item.id ||
              (item.id === 'program' &&
                ['classes', 'ppct', 'lesson-plan', 'journal'].includes(activeTab)) ||
              (item.id === 'students_group' &&
                ['students-list', 'students-export'].includes(activeTab)) ||
              (item.id === 'finance_group' &&
                ['finance-config', 'finance-ledger', 'finance-vouchers'].includes(activeTab));

            return (
              <button
                key={item.id}
                type="button"
                onClick={() => {
                  setActiveTab(item.id);
                  if (onClose) onClose();
                }}
                className={cn(
                  'w-full min-h-[46px] flex items-center gap-3 px-3.5 py-3 rounded-xl transition-colors duration-150 text-left group',
                  isActive
                    ? 'bg-white text-teal-900 shadow-sm font-extrabold'
                    : 'text-teal-50/95 hover:bg-white/12 hover:text-white font-semibold'
                )}
              >
                <item.icon
                  className={cn(
                    'w-5 h-5 shrink-0 transition-colors',
                    isActive ? 'text-teal-700' : 'text-teal-200 group-hover:text-white'
                  )}
                />
                <span className="text-[12.5px] leading-snug tracking-wide uppercase">
                  {item.label}
                </span>
              </button>
            );
          })}
        </div>

        {/* Mobile App Install + Account Footer */}
        <div className="p-4 border-t border-white/15 space-y-3 bg-teal-950/35">
          <PWAInstallButton
            variant="sidebar"
            className="bg-white/12 text-white border-white/20 hover:bg-white hover:text-teal-900"
          />

          <div className="flex items-center gap-3 p-3 bg-white/10 rounded-xl border border-white/15">
            <div className="w-9 h-9 rounded-lg bg-amber-400 text-slate-950 flex items-center justify-center text-sm font-extrabold shrink-0">
              {currentUser?.email?.charAt(0).toUpperCase() || 'A'}
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[11px] font-bold text-white truncate uppercase">
                {currentUser?.role === 'admin' ? 'Quản trị viên' : 'Người dùng'}
              </p>
              <p className="text-[11px] text-teal-100/85 truncate">{currentUser?.email}</p>
            </div>
          </div>

          <button
            type="button"
            onClick={onLogout}
            className="w-full min-h-[44px] flex items-center justify-center gap-2.5 px-4 py-2.5 text-rose-100 bg-rose-500/20 hover:bg-rose-500 hover:text-white border border-rose-300/25 rounded-xl transition-colors text-xs font-bold uppercase tracking-wider"
          >
            <LogOut className="w-4 h-4" />
            <span>Đăng xuất</span>
          </button>
        </div>
      </aside>
    </>
  );
};

export default Sidebar;
