import React from 'react';
import { ArrowLeft, type LucideIcon } from 'lucide-react';
import { cn } from '../lib/utils';

export interface BannerStatItem {
  label: string;
  value: string | number;
  sublabel?: string;
}

export interface BannerTabItem {
  id: string;
  label: string;
  icon?: LucideIcon;
}

interface SectionBannerProps {
  title: string;
  subtitle: string;
  icon: LucideIcon;
  onBack?: () => void;
  actions?: React.ReactNode;
  stats?: BannerStatItem[];
  tabs?: BannerTabItem[];
  activeTab?: string;
  onTabChange?: (tabId: string) => void;
}

const SectionBanner: React.FC<SectionBannerProps> = ({
  title,
  subtitle,
  icon: Icon,
  onBack,
  actions,
  stats,
  tabs,
  activeTab,
  onTabChange,
}) => {
  return (
    <div className="w-full rounded-2xl border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 overflow-hidden shadow-xs">
      {/* Top Emerald-Teal Banner Frame */}
      <div className="bg-gradient-to-r from-emerald-800 via-teal-700 to-teal-600 px-5 py-6 lg:px-8 lg:py-7 text-white">
        {onBack && (
          <div className="mb-4">
            <button
              type="button"
              onClick={onBack}
              className="inline-flex items-center gap-2 min-h-[36px] px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-xs font-semibold text-white/95 transition-colors"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              <span>Quay lại trang chủ</span>
            </button>
          </div>
        )}

        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-5 w-full">
          <div className="flex items-start sm:items-center gap-4 min-w-0 flex-1">
            <div className="w-12 h-12 lg:w-14 lg:h-14 rounded-2xl bg-white/15 border border-white/25 flex items-center justify-center shrink-0">
              <Icon className="w-6 h-6 lg:w-7 lg:h-7 text-white" />
            </div>
            <div className="min-w-0 flex-1">
              <h2 className="text-xl sm:text-2xl lg:text-[26px] font-extrabold tracking-tight text-white leading-tight">
                {title}
              </h2>
              <p className="text-sm lg:text-[15px] text-teal-50/95 mt-1 leading-relaxed font-medium w-full">
                {subtitle}
              </p>
            </div>
          </div>

          {actions && (
            <div className="flex flex-wrap items-center gap-2.5 shrink-0">
              {actions}
            </div>
          )}
        </div>
      </div>

      {/* Full-Width Information / KPI Strip Directly Below */}
      {stats && stats.length > 0 && (
        <div
          className={cn(
            'grid grid-cols-2 divide-x divide-y lg:divide-y-0 divide-slate-200/80 dark:divide-slate-800 border-t border-slate-200/80 dark:border-slate-800 bg-slate-50/70 dark:bg-slate-900/80',
            stats.length === 3 ? 'lg:grid-cols-3' : 'lg:grid-cols-4'
          )}
        >
          {stats.map((stat, idx) => (
            <div key={idx} className="px-4 py-4 lg:px-6 lg:py-4">
              <p className="text-xs font-semibold text-slate-500 dark:text-slate-400">
                {stat.label}
              </p>
              <p className="text-lg sm:text-xl font-extrabold text-slate-900 dark:text-white mt-1 tabular-nums tracking-tight">
                {stat.value}
              </p>
              {stat.sublabel && (
                <p className="text-[11px] font-medium text-slate-500 dark:text-slate-400 mt-0.5 truncate">
                  {stat.sublabel}
                </p>
              )}
            </div>
          ))}
        </div>
      )}

      {/* Segmented Sub-Navigation Tabs */}
      {tabs && tabs.length > 0 && onTabChange && (
        <div className="px-4 lg:px-8 py-3 bg-white dark:bg-slate-900 border-t border-slate-200/80 dark:border-slate-800 flex items-center gap-2 overflow-x-auto scrollbar-hide">
          {tabs.map((tab) => {
            const isActive = activeTab === tab.id;
            const TabIcon = tab.icon;
            return (
              <button
                key={tab.id}
                type="button"
                onClick={() => onTabChange(tab.id)}
                className={cn(
                  'min-h-[42px] px-4 py-2 rounded-xl text-xs sm:text-sm font-bold flex items-center gap-2 whitespace-nowrap transition-colors shrink-0',
                  isActive
                    ? 'bg-teal-600 text-white shadow-xs'
                    : 'bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 hover:bg-teal-50 dark:hover:bg-slate-700 hover:text-teal-700 dark:hover:text-white'
                )}
              >
                {TabIcon && <TabIcon className="w-4 h-4 shrink-0" />}
                <span>{tab.label}</span>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
};

export default SectionBanner;
