import React, { useState } from 'react';
import {
  X,
  Info,
  CreditCard,
  Layers,
  RotateCcw,
  GraduationCap,
  Calculator,
  BookOpen,
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import Logo from './Logo';

const Footer: React.FC = () => {
  const [activeModal, setActiveModal] = useState<{ title: string; content: React.ReactNode } | null>(
    null
  );

  const openInfoModal = (link: string) => {
    let content: React.ReactNode = null;
    const title = link;

    switch (link) {
      case 'Về chúng tôi':
        content = (
          <div className="space-y-4 text-sm sm:text-base">
            <p className="text-slate-600 dark:text-slate-300 leading-relaxed">
              Chào mừng Quý Thầy Cô đến với <span className="font-bold text-teal-700 dark:text-teal-400">HOÀNG GIA</span>.
              Hệ thống được thiết kế tối ưu dành riêng cho các thầy cô và trung tâm dạy thêm trên máy tính, điện thoại Android và iOS.
            </p>
            <p className="text-slate-600 dark:text-slate-300 leading-relaxed">
              Bao gồm các phân hệ trọng tâm:{' '}
              <span className="text-teal-700 dark:text-teal-400 font-bold">Quản lý học sinh</span>,{' '}
              <span className="text-emerald-700 dark:text-emerald-400 font-bold">Quản lý học tập học sinh</span>,{' '}
              <span className="text-cyan-700 dark:text-cyan-400 font-bold">Quản lý chương trình dạy</span>,{' '}
              <span className="text-amber-700 dark:text-amber-400 font-bold">Quản lý tài chính</span> và{' '}
              <span className="text-teal-800 dark:text-teal-300 font-bold">Báo cáo thuế cho HKD</span>.
            </p>
          </div>
        );
        break;
      case 'Tính năng hệ thống':
        content = (
          <ul className="space-y-2.5">
            {[
              { label: 'Cấu hình hộ kinh doanh', icon: RotateCcw, color: 'text-teal-600' },
              { label: 'Quản lý danh sách & đơn đăng ký học sinh', icon: Layers, color: 'text-emerald-600' },
              { label: 'Quản lý học tập học sinh (Tự động sắp lớp & chấm công)', icon: GraduationCap, color: 'text-cyan-600' },
              { label: 'Quản lý chương trình dạy, lịch báo giảng & sổ đầu bài', icon: BookOpen, color: 'text-teal-700' },
              { label: 'Quản lý tài chính, sổ S1a-HKD & phiếu thu chi', icon: CreditCard, color: 'text-amber-600' },
              { label: 'Báo cáo thuế cho Hộ kinh doanh (12 tháng & xuất cả năm)', icon: Calculator, color: 'text-emerald-700' },
            ].map((item, idx) => (
              <li
                key={idx}
                className="flex items-center gap-3 p-3 rounded-xl bg-slate-50 dark:bg-slate-800/70 border border-slate-200/80 dark:border-slate-700"
              >
                <item.icon className={`w-5 h-5 shrink-0 ${item.color}`} />
                <span className="font-semibold text-sm text-slate-800 dark:text-slate-200">
                  {item.label}
                </span>
              </li>
            ))}
          </ul>
        );
        break;
      case 'Bảng giá dịch vụ':
      case 'Điều khoản sử dụng':
      case 'Quy định vận hành':
      case 'Trung tâm hỗ trợ':
        content = (
          <div className="p-6 bg-teal-50 dark:bg-teal-950/40 border border-teal-200 dark:border-teal-800 rounded-2xl text-center">
            <p className="text-teal-800 dark:text-teal-300 font-bold text-base">
              Liên hệ Quản trị viên qua Zalo: 0366.000.555 để được hỗ trợ chi tiết.
            </p>
          </div>
        );
        break;
      case 'Chính sách bảo mật':
        content = (
          <div className="p-6 bg-emerald-50 dark:bg-emerald-950/40 border border-emerald-200 dark:border-emerald-800 rounded-2xl text-center">
            <p className="text-emerald-800 dark:text-emerald-300 font-bold text-base">
              Mỗi tài khoản cơ sở được cấu hình và lưu trữ dữ liệu độc lập, an toàn.
            </p>
          </div>
        );
        break;
      default:
        content = <p>Thông tin hệ thống Hoàng Gia.</p>;
    }

    setActiveModal({ title, content });
  };

  return (
    <footer className="bg-white dark:bg-slate-900 py-10 px-5 lg:px-8 mt-12 mb-16 lg:mb-0 border border-slate-200/90 dark:border-slate-800 rounded-2xl">
      <div className="w-full">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-10 lg:gap-16">
          {/* Column 1 */}
          <div className="space-y-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 shrink-0">
                <Logo className="w-full h-full" />
              </div>
              <div>
                <h2 className="text-xl font-extrabold text-slate-900 dark:text-white tracking-tight leading-none">
                  HOÀNG GIA
                </h2>
                <p className="text-teal-700 dark:text-teal-400 font-bold text-xs uppercase tracking-wider mt-1">
                  Trao cơ hội · Nhận niềm tin
                </p>
              </div>
            </div>
            <p className="text-slate-600 dark:text-slate-400 text-sm leading-relaxed">
              Giải pháp quản lý đồng bộ học sinh, điểm danh lớp học, chương trình giảng dạy, sổ sách tài chính và kê khai doanh thu thuế hộ kinh doanh.
            </p>
            <div className="space-y-1.5 text-sm text-slate-600 dark:text-slate-400">
              <div>
                <span className="font-bold text-slate-900 dark:text-white">Địa chỉ: </span>
                <span>SN 269 Lê Duẩn, P. Tân Phong, TP. Lai Châu</span>
              </div>
              <div>
                <span className="font-bold text-slate-900 dark:text-white">Hỗ trợ Zalo: </span>
                <span className="text-teal-700 dark:text-teal-400 font-bold tabular-nums">0366.000.555</span>
              </div>
            </div>
          </div>

          {/* Column 2 */}
          <div className="space-y-4">
            <h3 className="text-xs font-extrabold text-slate-900 dark:text-white uppercase tracking-wider">
              Khám phá hệ thống
            </h3>
            <ul className="space-y-2.5">
              {['Về chúng tôi', 'Tính năng hệ thống', 'Bảng giá dịch vụ', 'Trung tâm hỗ trợ'].map(
                (link) => (
                  <li key={link}>
                    <button
                      type="button"
                      onClick={() => openInfoModal(link)}
                      className="text-slate-600 dark:text-slate-400 hover:text-teal-700 dark:hover:text-teal-400 text-sm font-medium transition-colors"
                    >
                      {link}
                    </button>
                  </li>
                )
              )}
            </ul>
          </div>

          {/* Column 3 */}
          <div className="space-y-4">
            <h3 className="text-xs font-extrabold text-slate-900 dark:text-white uppercase tracking-wider">
              Quy định & Bảo mật
            </h3>
            <ul className="space-y-2.5">
              {['Điều khoản sử dụng', 'Chính sách bảo mật', 'Quy định vận hành'].map((link) => (
                <li key={link}>
                  <button
                    type="button"
                    onClick={() => openInfoModal(link)}
                    className="text-slate-600 dark:text-slate-400 hover:text-teal-700 dark:hover:text-teal-400 text-sm font-medium transition-colors"
                  >
                    {link}
                  </button>
                </li>
              ))}
            </ul>
          </div>
        </div>

        <div className="mt-8 pt-6 border-t border-slate-200/80 dark:border-slate-800 flex flex-col sm:flex-row justify-between items-center gap-2 text-xs text-slate-500 dark:text-slate-400 font-medium">
          <p>© {new Date().getFullYear()} Hệ thống Giáo dục Hoàng Gia. Bảo lưu mọi quyền.</p>
          <p>Tương thích Máy tính · Điện thoại Android · iPhone / iPad (iOS)</p>
        </div>
      </div>

      {/* Info Modal */}
      <AnimatePresence>
        {activeModal && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-slate-950/60 backdrop-blur-xs">
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 12 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 12 }}
              className="bg-white dark:bg-slate-900 w-full max-w-lg rounded-2xl shadow-2xl overflow-hidden border border-slate-200 dark:border-slate-800"
            >
              <div className="p-6 sm:p-7">
                <div className="flex items-center justify-between mb-5">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-xl bg-teal-50 dark:bg-teal-950/50 flex items-center justify-center text-teal-700 dark:text-teal-400">
                      <Info className="w-5 h-5" />
                    </div>
                    <h3 className="text-lg font-extrabold text-slate-900 dark:text-white">
                      {activeModal.title}
                    </h3>
                  </div>
                  <button
                    type="button"
                    onClick={() => setActiveModal(null)}
                    className="min-w-[40px] min-h-[40px] rounded-xl bg-slate-100 dark:bg-slate-800 flex items-center justify-center text-slate-500 hover:bg-rose-500 hover:text-white transition-colors"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </div>
                <div>{activeModal.content}</div>
                <div className="mt-6 pt-4 border-t border-slate-100 dark:border-slate-800 flex justify-end">
                  <button
                    type="button"
                    onClick={() => setActiveModal(null)}
                    className="btn-primary"
                  >
                    Đóng
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </footer>
  );
};

export default Footer;
