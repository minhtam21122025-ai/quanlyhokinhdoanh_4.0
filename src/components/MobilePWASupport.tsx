import React, { useState, useEffect } from 'react';
import {
  Smartphone,
  Download,
  Share,
  PlusSquare,
  CheckCircle2,
  X,
  WifiOff,
  Home,
  Users,
  GraduationCap,
  DollarSign,
  FileSpreadsheet,
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { usePWAInstall } from '../hooks/usePWAInstall';
import { cn } from '../lib/utils';

interface PWAInstallButtonProps {
  variant?: 'header' | 'sidebar' | 'banner';
  className?: string;
}

export const PWAInstallButton: React.FC<PWAInstallButtonProps> = ({ variant = 'header', className }) => {
  const { isInstallable, isInstalled, isIOS, isAndroid, promptInstall } = usePWAInstall();
  const [showGuideModal, setShowGuideModal] = useState(false);
  const [activeDeviceTab, setActiveDeviceTab] = useState<'android' | 'ios'>('android');

  useEffect(() => {
    if (isIOS) {
      setActiveDeviceTab('ios');
    } else if (isAndroid) {
      setActiveDeviceTab('android');
    }
  }, [isIOS, isAndroid]);

  if (isInstalled) {
    if (variant === 'sidebar') {
      return (
        <div className="flex items-center gap-2.5 px-3.5 py-2.5 rounded-xl bg-teal-50 dark:bg-teal-950/40 border border-teal-200 dark:border-teal-800/60 text-teal-800 dark:text-teal-300 text-xs font-semibold">
          <CheckCircle2 className="w-4 h-4 shrink-0 text-teal-600 dark:text-teal-400" />
          <span>Đã cài ứng dụng điện thoại</span>
        </div>
      );
    }
    return null;
  }

  const handleInstallClick = async () => {
    if (isInstallable) {
      const outcome = await promptInstall();
      if (outcome === 'manual') {
        setShowGuideModal(true);
      }
    } else {
      setShowGuideModal(true);
    }
  };

  return (
    <>
      {variant === 'header' && (
        <button
          type="button"
          onClick={handleInstallClick}
          className={cn(
            'min-h-[42px] px-3.5 py-2 rounded-xl bg-teal-50 dark:bg-teal-950/50 text-teal-700 dark:text-teal-300 border border-teal-200 dark:border-teal-800 hover:bg-teal-600 hover:text-white hover:border-teal-600 transition-colors flex items-center gap-2 text-xs font-bold whitespace-nowrap',
            className
          )}
          title="Cài đặt ứng dụng trên điện thoại Android & iOS"
        >
          <Smartphone className="w-4 h-4 shrink-0" />
          <span className="hidden md:inline">Cài App ĐT</span>
        </button>
      )}

      {variant === 'sidebar' && (
        <button
          type="button"
          onClick={handleInstallClick}
          className={cn(
            'w-full min-h-[44px] flex items-center justify-center gap-2.5 px-4 py-2.5 rounded-xl bg-teal-50 dark:bg-teal-950/40 text-teal-800 dark:text-teal-200 border border-teal-200 dark:border-teal-800/70 hover:bg-teal-600 hover:text-white transition-colors text-xs font-bold',
            className
          )}
        >
          <Download className="w-4 h-4 shrink-0" />
          <span>Cài ứng dụng Android / iOS</span>
        </button>
      )}

      {variant === 'banner' && (
        <button
          type="button"
          onClick={handleInstallClick}
          className={cn(
            'min-h-[44px] px-4 py-2.5 rounded-xl bg-white/15 hover:bg-white/25 text-white border border-white/25 transition-colors flex items-center gap-2 text-xs font-bold whitespace-nowrap',
            className
          )}
        >
          <Smartphone className="w-4 h-4 shrink-0" />
          <span>Dùng trên Android & iOS</span>
        </button>
      )}

      <AnimatePresence>
        {showGuideModal && (
          <div className="fixed inset-0 z-[120] flex items-center justify-center p-4 bg-slate-950/65 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.96, y: 12 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.96, y: 12 }}
              className="bg-white dark:bg-slate-900 w-full max-w-md rounded-2xl border border-slate-200 dark:border-slate-800 shadow-2xl overflow-hidden"
            >
              <div className="px-6 py-5 bg-gradient-to-r from-emerald-700 via-teal-600 to-teal-700 text-white flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-xl bg-white/15 flex items-center justify-center">
                    <Smartphone className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-base font-extrabold leading-tight">Cài đặt trên Điện thoại</h3>
                    <p className="text-xs text-teal-100 mt-0.5">Hỗ trợ đầy đủ Android & iPhone / iPad (iOS)</p>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => setShowGuideModal(false)}
                  className="min-w-[40px] min-h-[40px] rounded-xl bg-white/10 hover:bg-white/20 flex items-center justify-center transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="p-6 space-y-5">
                <div className="grid grid-cols-2 gap-2 p-1 bg-slate-100 dark:bg-slate-800 rounded-xl">
                  <button
                    type="button"
                    onClick={() => setActiveDeviceTab('android')}
                    className={cn(
                      'min-h-[42px] rounded-lg text-xs font-bold transition-colors',
                      activeDeviceTab === 'android'
                        ? 'bg-white dark:bg-slate-700 text-teal-700 dark:text-white shadow-sm'
                        : 'text-slate-600 dark:text-slate-400 hover:text-slate-900'
                    )}
                  >
                    Điện thoại Android
                  </button>
                  <button
                    type="button"
                    onClick={() => setActiveDeviceTab('ios')}
                    className={cn(
                      'min-h-[42px] rounded-lg text-xs font-bold transition-colors',
                      activeDeviceTab === 'ios'
                        ? 'bg-white dark:bg-slate-700 text-teal-700 dark:text-white shadow-sm'
                        : 'text-slate-600 dark:text-slate-400 hover:text-slate-900'
                    )}
                  >
                    iPhone / iPad (iOS)
                  </button>
                </div>

                {activeDeviceTab === 'android' ? (
                  <div className="space-y-3 text-sm text-slate-700 dark:text-slate-200">
                    <div className="p-3.5 rounded-xl bg-slate-50 dark:bg-slate-800/70 border border-slate-200/80 dark:border-slate-700 flex items-start gap-3">
                      <span className="w-6 h-6 rounded-full bg-teal-600 text-white text-xs font-bold flex items-center justify-center shrink-0 mt-0.5">
                        1
                      </span>
                      <p>Mở trang này bằng trình duyệt <strong>Google Chrome</strong> trên máy Android.</p>
                    </div>
                    <div className="p-3.5 rounded-xl bg-slate-50 dark:bg-slate-800/70 border border-slate-200/80 dark:border-slate-700 flex items-start gap-3">
                      <span className="w-6 h-6 rounded-full bg-teal-600 text-white text-xs font-bold flex items-center justify-center shrink-0 mt-0.5">
                        2
                      </span>
                      <p>Nhấn biểu tượng <strong>Menu 3 chấm (⋮)</strong> ở góc trên bên phải trình duyệt.</p>
                    </div>
                    <div className="p-3.5 rounded-xl bg-slate-50 dark:bg-slate-800/70 border border-slate-200/80 dark:border-slate-700 flex items-start gap-3">
                      <span className="w-6 h-6 rounded-full bg-teal-600 text-white text-xs font-bold flex items-center justify-center shrink-0 mt-0.5">
                        3
                      </span>
                      <p>Chọn <strong>"Cài đặt ứng dụng"</strong> hoặc <strong>"Thêm vào màn hình chính"</strong> để mở nhanh như ứng dụng gốc.</p>
                    </div>
                  </div>
                ) : (
                  <div className="space-y-3 text-sm text-slate-700 dark:text-slate-200">
                    <div className="p-3.5 rounded-xl bg-slate-50 dark:bg-slate-800/70 border border-slate-200/80 dark:border-slate-700 flex items-start gap-3">
                      <span className="w-6 h-6 rounded-full bg-teal-600 text-white text-xs font-bold flex items-center justify-center shrink-0 mt-0.5">
                        1
                      </span>
                      <p>Mở hệ thống bằng trình duyệt <strong>Safari</strong> trên iPhone hoặc iPad.</p>
                    </div>
                    <div className="p-3.5 rounded-xl bg-slate-50 dark:bg-slate-800/70 border border-slate-200/80 dark:border-slate-700 flex items-start gap-3">
                      <span className="w-6 h-6 rounded-full bg-teal-600 text-white text-xs font-bold flex items-center justify-center shrink-0 mt-0.5">
                        2
                      </span>
                      <div className="flex-1">
                        <p className="flex items-center gap-1.5 flex-wrap">
                          Nhấn nút <strong>Chia sẻ</strong> <Share className="w-4 h-4 text-teal-600 inline" /> ở thanh công cụ phía dưới.
                        </p>
                      </div>
                    </div>
                    <div className="p-3.5 rounded-xl bg-slate-50 dark:bg-slate-800/70 border border-slate-200/80 dark:border-slate-700 flex items-start gap-3">
                      <span className="w-6 h-6 rounded-full bg-teal-600 text-white text-xs font-bold flex items-center justify-center shrink-0 mt-0.5">
                        3
                      </span>
                      <div className="flex-1">
                        <p className="flex items-center gap-1.5 flex-wrap">
                          Cuộn xuống chọn <strong>Thêm vào MH chính</strong> <PlusSquare className="w-4 h-4 text-teal-600 inline" /> rồi nhấn <strong>Thêm</strong>.
                        </p>
                      </div>
                    </div>
                  </div>
                )}

                <div className="pt-2 flex justify-end">
                  <button
                    type="button"
                    onClick={() => setShowGuideModal(false)}
                    className="btn-primary w-full"
                  >
                    Đã hiểu
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </>
  );
};

export const OfflineNoticeBar: React.FC = () => {
  const [isOffline, setIsOffline] = useState<boolean>(
    typeof navigator !== 'undefined' ? !navigator.onLine : false
  );

  useEffect(() => {
    const goOffline = () => setIsOffline(true);
    const goOnline = () => setIsOffline(false);
    window.addEventListener('offline', goOffline);
    window.addEventListener('online', goOnline);
    return () => {
      window.removeEventListener('offline', goOffline);
      window.removeEventListener('online', goOnline);
    };
  }, []);

  if (!isOffline) return null;

  return (
    <div className="bg-amber-600 text-white px-4 py-2 text-xs font-semibold flex items-center justify-center gap-2">
      <WifiOff className="w-4 h-4 shrink-0" />
      <span>Đang ở chế độ ngoại tuyến (Offline) — Dữ liệu vẫn được lưu an toàn trên thiết bị.</span>
    </div>
  );
};

interface MobileBottomNavProps {
  activeTab: string;
  setActiveTab: (tab: string) => void;
}

export const MobileBottomNav: React.FC<MobileBottomNavProps> = ({ activeTab, setActiveTab }) => {
  const navItems = [
    { id: 'dashboard', label: 'Trang chủ', icon: Home, matches: ['dashboard'] },
    { id: 'students_group', label: 'Học sinh', icon: Users, matches: ['students_group', 'students-list', 'students-export'] },
    { id: 'student_learning_group', label: 'Học tập', icon: GraduationCap, matches: ['student_learning_group'] },
    { id: 'finance_group', label: 'Tài chính', icon: DollarSign, matches: ['finance_group', 'finance-config', 'finance-ledger', 'finance-vouchers'] },
    { id: 'tax_report', label: 'Thuế HKD', icon: FileSpreadsheet, matches: ['tax_report'] },
  ];

  return (
    <nav
      aria-label="Điều hướng di động"
      className="lg:hidden fixed bottom-0 left-0 right-0 z-40 bg-white/95 dark:bg-slate-900/95 backdrop-blur-md border-t border-slate-200 dark:border-slate-800 pb-safe"
    >
      <div className="grid grid-cols-5 items-center h-14 px-1">
        {navItems.map((item) => {
          const isActive = item.matches.includes(activeTab);
          return (
            <button
              key={item.id}
              type="button"
              onClick={() => setActiveTab(item.id)}
              className={cn(
                'min-h-[48px] flex flex-col items-center justify-center gap-0.5 rounded-lg transition-colors',
                isActive
                  ? 'text-teal-700 dark:text-teal-400 font-bold'
                  : 'text-slate-500 dark:text-slate-400 font-medium'
              )}
            >
              <item.icon className={cn('w-5 h-5', isActive && 'stroke-[2.5]')} />
              <span className="text-[10px] leading-tight tracking-tight truncate max-w-full px-1">
                {item.label}
              </span>
            </button>
          );
        })}
      </div>
    </nav>
  );
};
