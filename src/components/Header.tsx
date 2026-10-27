import React, { useState } from 'react';
import {
  Calendar,
  ChevronLeft,
  ChevronRight,
  X,
  Menu,
  Moon,
  Sun,
} from 'lucide-react';
import { cn } from '../lib/utils';
import {
  format,
  startOfMonth,
  endOfMonth,
  startOfWeek,
  endOfWeek,
  eachDayOfInterval,
  isSameMonth,
  isSameDay,
  addMonths,
  subMonths,
} from 'date-fns';
import { vi } from 'date-fns/locale';
import { motion, AnimatePresence } from 'motion/react';
import Logo from './Logo';
import { PWAInstallButton } from './MobilePWASupport';

interface HeaderProps {
  activeTab: string;
  setActiveTab: (tab: string) => void;
  currentUser: any;
  onLogout: () => void;
  logoUrl?: string;
  onMenuToggle?: () => void;
  darkMode?: boolean;
  setDarkMode?: (val: boolean) => void;
}

const TAB_LABELS: Record<string, string> = {
  dashboard: 'Trang chủ giới thiệu',
  business: 'Cấu hình Hộ kinh doanh',
  students_group: 'Quản lý Học sinh',
  'students-list': 'Quản lý Học sinh · Danh sách',
  'students-export': 'Quản lý Học sinh · Xuất đơn đăng ký',
  student_learning_group: 'Quản lý Học tập & Chấm công Học sinh',
  program: 'Quản lý Chương trình dạy',
  classes: 'Chương trình dạy · Cấu hình lớp',
  ppct: 'Chương trình dạy · Phân phối chương trình',
  'lesson-plan': 'Chương trình dạy · Lịch báo giảng',
  journal: 'Chương trình dạy · Sổ đầu bài',
  finance_group: 'Quản lý Tài chính',
  'finance-config': 'Quản lý Tài chính · Cấu hình & Thu chi',
  'finance-ledger': 'Quản lý Tài chính · Sổ doanh thu S1a-HKD',
  'finance-vouchers': 'Quản lý Tài chính · Phiếu thu & Phiếu chi',
  tax_report: 'Báo cáo Thuế Hộ kinh doanh',
  users: 'Quản lý Tài khoản hệ thống',
  menu_settings: 'Sắp xếp thứ tự Menu',
};

const Header: React.FC<HeaderProps> = ({
  activeTab,
  currentUser,
  onLogout,
  logoUrl,
  onMenuToggle,
  darkMode,
  setDarkMode,
}) => {
  const [isCalendarOpen, setIsCalendarOpen] = useState(false);
  const [currentMonth, setCurrentMonth] = useState(new Date());

  const pageLabel = TAB_LABELS[activeTab] || 'Hệ thống Quản lý Hoàng Gia';

  return (
    <header className="sticky top-0 z-[40] bg-white/95 dark:bg-slate-900/95 backdrop-blur-md border-b border-slate-200/90 dark:border-slate-800 h-16 px-4 lg:px-8 flex items-center justify-between gap-3">
      <div className="flex items-center gap-3 min-w-0">
        <button
          type="button"
          onClick={onMenuToggle}
          className="lg:hidden min-w-[42px] min-h-[42px] flex items-center justify-center text-slate-700 dark:text-slate-200 hover:text-teal-600 hover:bg-teal-50 dark:hover:bg-slate-800 rounded-xl transition-colors"
          aria-label="Mở menu"
        >
          <Menu className="w-5 h-5" />
        </button>

        <div className="lg:hidden w-8 h-8 shrink-0">
          <Logo src={logoUrl} className="w-full h-full" />
        </div>

        <div className="min-w-0">
          <p className="text-[11px] font-bold text-teal-700 dark:text-teal-400 uppercase tracking-wider hidden sm:block">
            Hệ thống Giáo dục Hoàng Gia
          </p>
          <h2 className="text-sm sm:text-base lg:text-lg font-extrabold text-slate-900 dark:text-white truncate">
            {pageLabel}
          </h2>
        </div>
      </div>

      <div className="flex items-center gap-2 sm:gap-3 shrink-0">
        <PWAInstallButton variant="header" />

        {setDarkMode && (
          <button
            type="button"
            onClick={() => setDarkMode(!darkMode)}
            className="min-w-[42px] min-h-[42px] flex items-center justify-center text-slate-600 dark:text-slate-300 hover:text-teal-600 hover:bg-slate-100 dark:hover:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-800 transition-colors"
            title={darkMode ? 'Chế độ sáng' : 'Chế độ tối'}
          >
            {darkMode ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
          </button>
        )}

        <button
          type="button"
          onClick={() => setIsCalendarOpen(true)}
          className="min-w-[42px] min-h-[42px] flex items-center justify-center text-slate-600 dark:text-slate-300 hover:text-teal-600 hover:bg-slate-100 dark:hover:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-800 transition-colors"
          title="Lịch tháng"
        >
          <Calendar className="w-4 h-4" />
        </button>

        <div className="hidden sm:block h-6 w-px bg-slate-200 dark:bg-slate-800" />

        <div className="flex items-center gap-2.5">
          <div className="text-right hidden md:block">
            <p className="text-[11px] font-bold text-slate-900 dark:text-white leading-none uppercase">
              {currentUser?.role === 'admin' ? 'Quản trị viên' : 'Người dùng'}
            </p>
            <button
              type="button"
              onClick={onLogout}
              className="text-[10px] font-bold text-rose-600 hover:underline uppercase mt-1"
            >
              Đăng xuất
            </button>
          </div>
          <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-emerald-600 to-teal-700 flex items-center justify-center text-white text-sm font-extrabold shadow-xs">
            {currentUser?.email?.charAt(0).toUpperCase() || 'A'}
          </div>
        </div>
      </div>

      {/* Calendar Modal */}
      <AnimatePresence>
        {isCalendarOpen && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-slate-950/60 backdrop-blur-xs">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white dark:bg-slate-900 rounded-2xl shadow-2xl w-full max-w-md overflow-hidden border border-slate-200 dark:border-slate-800"
            >
              <div className="p-5 border-b border-slate-100 dark:border-slate-800 flex items-center justify-between bg-slate-50 dark:bg-slate-800/40">
                <div className="flex items-center gap-3">
                  <button
                    type="button"
                    onClick={() => setCurrentMonth(subMonths(currentMonth, 1))}
                    className="min-w-[38px] min-h-[38px] flex items-center justify-center hover:bg-slate-200 dark:hover:bg-slate-700 rounded-xl transition-colors text-slate-700 dark:text-slate-200"
                  >
                    <ChevronLeft className="w-5 h-5" />
                  </button>
                  <h3 className="text-base font-extrabold text-slate-900 dark:text-white capitalize">
                    {format(currentMonth, 'MMMM yyyy', { locale: vi })}
                  </h3>
                  <button
                    type="button"
                    onClick={() => setCurrentMonth(addMonths(currentMonth, 1))}
                    className="min-w-[38px] min-h-[38px] flex items-center justify-center hover:bg-slate-200 dark:hover:bg-slate-700 rounded-xl transition-colors text-slate-700 dark:text-slate-200"
                  >
                    <ChevronRight className="w-5 h-5" />
                  </button>
                </div>
                <button
                  type="button"
                  onClick={() => setIsCalendarOpen(false)}
                  className="min-w-[38px] min-h-[38px] flex items-center justify-center hover:bg-rose-50 dark:hover:bg-rose-900/20 text-slate-400 hover:text-rose-600 rounded-xl transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="p-5">
                <div className="grid grid-cols-7 gap-1 mb-2">
                  {['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN'].map((day) => (
                    <div
                      key={day}
                      className="text-center text-[11px] font-bold text-slate-500 dark:text-slate-400 uppercase py-1.5"
                    >
                      {day}
                    </div>
                  ))}
                </div>
                <div className="grid grid-cols-7 gap-1">
                  {(() => {
                    const monthStart = startOfMonth(currentMonth);
                    const monthEnd = endOfMonth(monthStart);
                    const startDate = startOfWeek(monthStart, { weekStartsOn: 1 });
                    const endDate = endOfWeek(monthEnd, { weekStartsOn: 1 });
                    const calendarDays = eachDayOfInterval({ start: startDate, end: endDate });

                    return calendarDays.map((day) => {
                      const isCurrentMonth = isSameMonth(day, monthStart);
                      const isToday = isSameDay(day, new Date());

                      return (
                        <div
                          key={day.toString()}
                          className={cn(
                            'aspect-square flex items-center justify-center rounded-xl text-sm font-bold tabular-nums transition-colors',
                            !isCurrentMonth && 'text-slate-300 dark:text-slate-700',
                            isCurrentMonth &&
                              !isToday &&
                              'text-slate-700 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800',
                            isToday && 'bg-teal-600 text-white shadow-sm'
                          )}
                        >
                          {format(day, 'd')}
                        </div>
                      );
                    });
                  })()}
                </div>
              </div>

              <div className="p-4 bg-slate-50 dark:bg-slate-800/40 border-t border-slate-100 dark:border-slate-800 flex justify-center">
                <button
                  type="button"
                  onClick={() => setCurrentMonth(new Date())}
                  className="text-xs font-extrabold text-teal-700 dark:text-teal-400 hover:underline uppercase tracking-wider tabular-nums"
                >
                  Hôm nay: {format(new Date(), 'dd/MM/yyyy')}
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </header>
  );
};

export default Header;
