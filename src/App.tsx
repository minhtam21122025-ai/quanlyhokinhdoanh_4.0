/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo } from 'react';
import { 
  Building2, 
  GraduationCap, 
  BookOpen, 
  CalendarDays, 
  ClipboardList,
  Users,
  Plus,
  Trash2,
  Download,
  Upload,
  Save,
  Sparkles,
  FileText,
  FileSpreadsheet,
  DollarSign,
  Settings,
  PieChart,
  ArrowRight,
  ArrowLeft,
  Check,
  X,
  ChevronDown,
  ChevronUp,
  BarChart3,
  Receipt,
  Info,
  Lock,
  Calendar,
  RefreshCw,
  Moon,
  Sun,
  TrendingUp,
  CreditCard,
  User,
  Home,
  ShieldCheck,
  Layers,
  RotateCcw,
  Eye,
  EyeOff,
  Calculator,
  CheckCircle2,
  Smartphone
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, AlignmentType, TextRun, VerticalAlign, BorderStyle, PageBreak } from 'docx';
import { saveAs } from 'file-saver';
import { format, startOfWeek, endOfWeek, addDays, parseISO } from 'date-fns';

import { cn } from './lib/utils';

import { 
  BusinessInfo, 
  ClassSubject, 
  PPCTItem, 
  LessonPlan, 
  LessonPlanRow, 
  Student,
  StudentAttendanceRecord,
  AnnualTaxReport,
  FinancialConfig,
  IncomeItem,
  ExpenseItem
} from './types';

import Footer from './components/Footer';
import Header from './components/Header';
import Sidebar from './components/Sidebar';
import Logo from './components/Logo';
import SectionBanner from './components/SectionBanner';
import StudentLearningSection, { deriveAutoClassName } from './components/StudentLearningSection';
import TaxReportSection from './components/TaxReportSection';
import { OfflineNoticeBar, MobileBottomNav, PWAInstallButton } from './components/MobilePWASupport';

// --- Utilities ---

function numberToVietnameseWords(number: number): string {
  if (number === 0) return "Không đồng";

  const units = ["", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"];
  const tens = ["", "mười", "hai mươi", "ba mươi", "bốn mươi", "năm mươi", "sáu mươi", "bảy mươi", "tám mươi", "chín mươi"];
  const hundreds = ["không trăm", "một trăm", "hai trăm", "ba trăm", "bốn trăm", "năm trăm", "sáu trăm", "bảy trăm", "tám trăm", "chín trăm"];
  const groups = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"];

  function readThreeDigits(n: number, isFirstGroup: boolean): string {
    let res = "";
    const h = Math.floor(n / 100);
    const t = Math.floor((n % 100) / 10);
    const u = n % 10;

    if (h > 0 || !isFirstGroup) {
      res += hundreds[h] + " ";
    }

    if (t > 0) {
      res += tens[t] + " ";
      if (u === 1 && t > 1) res = res.trim() + " mốt ";
      else if (u === 5) res = res.trim() + " lăm ";
      else if (u > 0) res += units[u] + " ";
    } else if (u > 0) {
      if (!isFirstGroup || h > 0) res += "lẻ ";
      res += units[u] + " ";
    }

    return res.trim();
  }

  let res = "";
  let groupIdx = 0;
  let tempNum = number;

  while (tempNum > 0) {
    const groupVal = tempNum % 1000;
    if (groupVal > 0) {
      const groupStr = readThreeDigits(groupVal, tempNum < 1000);
      res = groupStr + " " + groups[groupIdx] + " " + res;
    }
    tempNum = Math.floor(tempNum / 1000);
    groupIdx++;
  }

  res = res.trim();
  if (res.startsWith("không trăm ")) {
    res = res.substring(11);
  }
  if (res.startsWith("lẻ ")) {
    res = res.substring(3);
  }

  return res.charAt(0).toUpperCase() + res.slice(1) + " đồng chẵn";
}

const formatDate = (dateStr: string) => {
  if (!dateStr) return "";
  const parts = dateStr.split('-');
  if (parts.length === 3) {
    return `${parts[2]}/${parts[1]}/${parts[0]}`;
  }
  return dateStr;
};

const sanitizeDocxText = (text: any): string => {
  if (text === null || text === undefined) return "";
  const str = String(text);
  return str.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x84\x86-\x9F\uFDD0-\uFDEF\uFFFE\uFFFF]/g, "");
};

const getLastDayOfMonth = (period: string) => {
  const match = period.match(/Tháng (\d{2})\/(\d{4})/i);
  if (match) {
    const month = parseInt(match[1]);
    const year = parseInt(match[2]);
    const lastDay = new Date(year, month, 0).getDate();
    return `${String(lastDay).padStart(2, '0')}/${String(month).padStart(2, '0')}/${year}`;
  }
  return "";
};

function ConfirmButton({ 
  onConfirm, 
  children, 
  className, 
  confirmText = "Bạn có chắc chắn?",
  icon: Icon
}: { 
  onConfirm: () => void, 
  children?: React.ReactNode, 
  className?: string, 
  confirmText?: string,
  icon?: any
}) {
  const [isConfirming, setIsConfirming] = useState(false);

  useEffect(() => {
    if (isConfirming) {
      const timer = setTimeout(() => setIsConfirming(false), 3000);
      return () => clearTimeout(timer);
    }
  }, [isConfirming]);

  if (isConfirming) {
    return (
      <button
        type="button"
        onClick={(e) => {
          e.stopPropagation();
          onConfirm();
          setIsConfirming(false);
        }}
        className={cn("flex items-center gap-2 px-4 py-2 bg-rose-600 text-white rounded-xl hover:bg-rose-700 transition-colors shadow-xs font-bold text-xs", className)}
      >
        {confirmText || "Xác nhận?"}
      </button>
    );
  }

  return (
    <button
      type="button"
      onClick={(e) => {
        e.stopPropagation();
        setIsConfirming(true);
      }}
      className={className}
    >
      {Icon && <Icon className="w-4 h-4" />}
      {children}
    </button>
  );
}

const normalizeGrade = (g: any) => String(g || '').replace(/\D/g, '');

const safeFormat = (dateStr: string | undefined, formatStr: string) => {
  if (!dateStr) return '...';
  try {
    const date = parseISO(dateStr);
    if (isNaN(date.getTime())) return '...';
    return format(date, formatStr);
  } catch (e) {
    return '...';
  }
};

const STORAGE_KEY = 'tutoring_center_data';
const USERS_KEY = 'tutoring_center_users';
const THEME_KEY = 'tutoring_center_theme';

const DEFAULT_MENU_ORDER = [
  'dashboard',
  'business',
  'students_group',
  'student_learning_group',
  'program',
  'finance_group',
  'tax_report',
];

const DEFAULT_CLASSES: ClassSubject[] = [
  { id: 'c6-1', grade: '6', subject: 'Toán', subSubject: 'Số học' },
  { id: 'c6-2', grade: '6', subject: 'Toán', subSubject: 'Đại số' },
  { id: 'c6-3', grade: '6', subject: 'Toán', subSubject: 'Hình học' },
  { id: 'c6-4', grade: '6', subject: 'Ngữ văn', subSubject: '' },
  { id: 'c6-7', grade: '6', subject: 'KHTN', subSubject: 'Hóa học' },
  { id: 'c6-8', grade: '6', subject: 'KHTN', subSubject: 'Sinh học' },

  { id: 'c7-1', grade: '7', subject: 'Toán', subSubject: 'Số học' },
  { id: 'c7-2', grade: '7', subject: 'Toán', subSubject: 'Đại số' },
  { id: 'c7-3', grade: '7', subject: 'Toán', subSubject: 'Hình học' },
  { id: 'c7-4', grade: '7', subject: 'Ngữ văn', subSubject: '' },
  { id: 'c7-7', grade: '7', subject: 'KHTN', subSubject: 'Hóa học' },
  { id: 'c7-8', grade: '7', subject: 'KHTN', subSubject: 'Sinh học' },

  { id: 'c8-1', grade: '8', subject: 'Toán', subSubject: 'Số học' },
  { id: 'c8-2', grade: '8', subject: 'Toán', subSubject: 'Đại số' },
  { id: 'c8-3', grade: '8', subject: 'Toán', subSubject: 'Hình học' },
  { id: 'c8-4', grade: '8', subject: 'Ngữ văn', subSubject: '' },
  { id: 'c8-7', grade: '8', subject: 'KHTN', subSubject: 'Hóa học' },
  { id: 'c8-8', grade: '8', subject: 'KHTN', subSubject: 'Sinh học' },

  { id: 'c9-1', grade: '9', subject: 'Toán', subSubject: 'Số học' },
  { id: 'c9-2', grade: '9', subject: 'Toán', subSubject: 'Đại số' },
  { id: 'c9-3', grade: '9', subject: 'Toán', subSubject: 'Hình học' },
  { id: 'c9-4', grade: '9', subject: 'Ngữ văn', subSubject: '' },
  { id: 'c9-7', grade: '9', subject: 'KHTN', subSubject: 'Hóa học' },
  { id: 'c9-8', grade: '9', subject: 'KHTN', subSubject: 'Sinh học' },
];

const DEFAULT_PPCT: PPCTItem[] = [
  { id: 'p6-1', grade: '6', subject: 'Toán', subSubject: 'Đại số', period: 1, content: 'Tập hợp các số tự nhiên', notes: '' },
  { id: 'p7-1', grade: '7', subject: 'Toán', subSubject: 'Đại số', period: 1, content: 'Số hữu tỉ', notes: '' },
  { id: 'p8-1', grade: '8', subject: 'Toán', subSubject: 'Đại số', period: 1, content: 'Đa thức', notes: '' },
  { id: 'p9-1', grade: '9', subject: 'Toán', subSubject: 'Đại số', period: 1, content: 'Căn bậc hai', notes: '' },
];

interface UserAccount {
  id: string;
  email: string;
  password: string;
  role: 'admin' | 'user';
  expiryDate?: string;
  maxDevices?: number;
  taxCode?: string;
  businessName?: string;
  businessAddress?: string;
  businessOwner?: string;
  businessLocation?: string;
  createdAt: string;
}

const SYNC_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbynIT2ZNiXJvmOB3LUi9YisDYS2Crin7G0skh-cDseNNIGXfy4PK3MMMcD0lYVhMSke/exec';

function sanitizeMenuOrder(savedOrder: unknown): string[] {
  const validIds = new Set(DEFAULT_MENU_ORDER);
  const filtered = Array.isArray(savedOrder)
    ? savedOrder.filter((id): id is string => typeof id === 'string' && validIds.has(id))
    : [];
  DEFAULT_MENU_ORDER.forEach((id) => {
    if (!filtered.includes(id)) {
      filtered.push(id);
    }
  });
  return filtered;
}

// --- Components ---

export default function App() {
  const [darkMode, setDarkMode] = useState(() => {
    const saved = localStorage.getItem(THEME_KEY);
    return saved === 'dark';
  });

  useEffect(() => {
    if (darkMode) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
    localStorage.setItem(THEME_KEY, darkMode ? 'dark' : 'light');
  }, [darkMode]);

  const [toastMessage, setToastMessage] = useState<{ text: string; type: 'success' | 'error' } | null>(null);
  const notify = (text: string, type: 'success' | 'error' = 'success') => {
    setToastMessage({ text, type });
    setTimeout(() => setToastMessage(null), 3500);
  };

  const [currentUser, setCurrentUser] = useState<UserAccount | null>(null);
  const [users, setUsers] = useState<UserAccount[]>(() => {
    const saved = localStorage.getItem(USERS_KEY);
    const admins: UserAccount[] = [
      {
        id: 'admin-1',
        email: 'cosogiaoduchoanggia269@gmail.com',
        password: '123456@',
        role: 'admin',
        createdAt: new Date().toISOString()
      },
      {
        id: 'admin-2',
        email: '0366000555',
        password: '123456@',
        role: 'admin',
        createdAt: new Date().toISOString()
      }
    ];

    if (saved) {
      let parsed: UserAccount[] = JSON.parse(saved);
      admins.forEach(admin => {
        const index = parsed.findIndex(u => u.email.toLowerCase() === admin.email.toLowerCase());
        if (index === -1) {
          parsed.push(admin);
        } else {
          parsed[index] = { ...parsed[index], password: admin.password, role: 'admin' };
        }
      });
      return parsed;
    }
    return admins;
  });

  const [activeTab, setActiveTab] = useState('dashboard');
  const [isDataLoaded, setIsDataLoaded] = useState(false);
  const [showWelcomeModal, setShowWelcomeModal] = useState(false);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [businessInfo, setBusinessInfo] = useState<BusinessInfo>({
    name: 'CƠ SỞ GIÁO DỤC HOÀNG GIA',
    address: 'SN 269 Lê Duẩn - P. Tân Phong - T. Lai Châu',
    taxId: '034150007741',
    owner: 'Hoàng Thị Mơ',
    businessLocation: 'Lai Châu'
  });
  const [classes, setClasses] = useState<ClassSubject[]>([]);
  const [ppctData, setPpctData] = useState<PPCTItem[]>([]);
  const [lessonPlans, setLessonPlans] = useState<LessonPlan[]>([]);
  const [students, setStudents] = useState<Student[]>([]);
  const [attendanceRecords, setAttendanceRecords] = useState<StudentAttendanceRecord[]>([]);
  const [taxReportsByYear, setTaxReportsByYear] = useState<Record<number, AnnualTaxReport>>({});
  const [financialConfig, setFinancialConfig] = useState<FinancialConfig>({
    reportPeriod: '',
    receiptDate: '',
    paymentDate: '',
    preparer: '',
    treasurer: '',
    taxCode: ''
  });
  const [incomeData, setIncomeData] = useState<IncomeItem[]>([]);
  const [expenseData, setExpenseData] = useState<ExpenseItem[]>([]);
  const [isSyncing, setIsSyncing] = useState(false);

  const effectiveBusinessInfo = React.useMemo(() => {
    const fixedInfo: BusinessInfo = {
      name: businessInfo.name || 'CƠ SỞ GIÁO DỤC HOÀNG GIA',
      address: businessInfo.address || 'SN 269 Lê Duẩn - P. Tân Phong - T. Lai Châu',
      taxId: businessInfo.taxId || '034150007741',
      owner: businessInfo.owner || 'Hoàng Thị Mơ',
      businessLocation: businessInfo.businessLocation || 'Lai Châu'
    };

    if (currentUser?.role === 'admin') return fixedInfo;
    
    return {
      name: currentUser?.businessName || fixedInfo.name,
      owner: currentUser?.businessOwner || fixedInfo.owner,
      taxId: currentUser?.taxCode || fixedInfo.taxId,
      address: currentUser?.businessAddress || fixedInfo.address,
      businessLocation: currentUser?.businessLocation || fixedInfo.businessLocation
    };
  }, [currentUser, businessInfo]);

  const [menuOrder, setMenuOrder] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem('MENU_ORDER');
      return sanitizeMenuOrder(saved ? JSON.parse(saved) : DEFAULT_MENU_ORDER);
    } catch {
      return DEFAULT_MENU_ORDER;
    }
  });

  useEffect(() => {
    localStorage.setItem('MENU_ORDER', JSON.stringify(menuOrder));
  }, [menuOrder]);

  const syncToGoogleSheets = async (action: 'sync_users' | 'sync_business', payload: any) => {
    setIsSyncing(true);
    try {
      await fetch(SYNC_SCRIPT_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action, ...payload })
      });
    } catch (error) {
      console.error(`Sync ${action} failed`, error);
    } finally {
      setIsSyncing(false);
    }
  };

  const fetchFromGoogleSheets = async () => {
    setIsSyncing(true);
    try {
      const response = await fetch(`${SYNC_SCRIPT_URL}?action=fetch_data`);
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }
      const data = await response.json();
      if (data.users && data.users.length > 0) {
        const newUsers = data.users.map((u: any) => ({
          ...u,
          id: u.id || `user-${Date.now()}-${Math.random()}`,
          role: u.role || 'user',
          maxDevices: parseInt(u.maxDevices) || 1
        }));
        setUsers(newUsers);
      }
      if (data.businessInfo) {
        setBusinessInfo(prev => ({
          ...prev,
          name: data.businessInfo.name || prev.name,
          owner: data.businessInfo.owner || prev.owner,
          taxId: data.businessInfo.taxId || prev.taxId,
          address: data.businessInfo.address || prev.address,
          businessLocation: data.businessInfo.businessLocation || prev.businessLocation
        }));
      }
      notify('Tải dữ liệu từ Google Sheet thành công!');
    } catch (error) {
      console.error("Fetch from Google Sheets failed:", error);
      notify(`Tải dữ liệu thất bại: ${error instanceof Error ? error.message : 'Lỗi kết nối'}`, 'error');
    } finally {
      setIsSyncing(false);
    }
  };

  // Load/Reset data when user changes
  useEffect(() => {
    setIsDataLoaded(false);
    
    if (!currentUser) {
      setBusinessInfo({
        name: 'CƠ SỞ GIÁO DỤC HOÀNG GIA',
        address: 'SN 269 Lê Duẩn - P. Tân Phong - T. Lai Châu',
        taxId: '034150007741',
        owner: 'Hoàng Thị Mơ',
        businessLocation: 'Lai Châu'
      });
      setClasses(DEFAULT_CLASSES);
      setPpctData(DEFAULT_PPCT);
      setLessonPlans([]);
      setStudents([]);
      setAttendanceRecords([]);
      setTaxReportsByYear({});
      setFinancialConfig({ reportPeriod: '', receiptDate: '', paymentDate: '', preparer: '', treasurer: '', taxCode: '' });
      setIncomeData([]);
      setExpenseData([]);
      setIsDataLoaded(true);
      return;
    }

    const userStorageKey = `${STORAGE_KEY}_${currentUser.id}`;
    const saved = localStorage.getItem(userStorageKey);
    if (saved) {
      try {
        const data = JSON.parse(saved);
        if (data.businessInfo) {
          setBusinessInfo({
            ...data.businessInfo,
            name: data.businessInfo.name || currentUser.businessName || 'CƠ SỞ GIÁO DỤC HOÀNG GIA',
            address: data.businessInfo.address || currentUser.businessAddress || 'SN 269 Lê Duẩn - P. Tân Phong - T. Lai Châu',
            owner: data.businessInfo.owner || currentUser.businessOwner || 'Hoàng Thị Mơ',
            businessLocation: data.businessInfo.businessLocation || currentUser.businessLocation || 'Lai Châu',
            taxId: data.businessInfo.taxId || currentUser.taxCode || '034150007741'
          });
        }
        
        const migrateSubject = (item: any) => {
          if (item.subject === 'Khoa học tự nhiên') {
            item.subject = 'KHTN';
          }
          if (item.subject === 'KHTN') {
            if (item.subSubject === 'Vật lý') item.subSubject = 'Vật lí';
          }
          if (item.subject === 'Toán') {
            if (item.subSubject === 'Số học và Đại số' || item.subSubject === 'Số và Đại số') {
              item.subSubject = 'Đại số';
            }
            if (item.subSubject === 'Hình học và Đo lường' || item.subSubject === 'Hình học và đo lường') {
              item.subSubject = 'Hình học';
            }
          }
          return item;
        };

        const loadedClasses = (data.classes && data.classes.length > 0 ? data.classes : DEFAULT_CLASSES).map(migrateSubject);
        const loadedPpct = (data.ppctData && data.ppctData.length > 0 ? data.ppctData : DEFAULT_PPCT).map(migrateSubject);
        
        setClasses(loadedClasses);
        setPpctData(loadedPpct);
        
        if (data.students) {
          const normalizedStudents = (data.students as Student[]).map((st) => ({
            ...st,
            assignedClass: st.assignedClass || deriveAutoClassName(st),
          }));
          setStudents(normalizedStudents);
        }
        if (data.attendanceRecords) setAttendanceRecords(data.attendanceRecords);
        if (data.taxReportsByYear) setTaxReportsByYear(data.taxReportsByYear);
        if (data.incomeData) setIncomeData(data.incomeData);
        if (data.expenseData) setExpenseData(data.expenseData);
        
        if (data.lessonPlans) {
          const migratedPlans = data.lessonPlans.map((plan: LessonPlan) => ({
            ...plan,
            rows: plan.rows.map(migrateSubject)
          }));
          setLessonPlans(migratedPlans);
        }
        if (data.financialConfig) setFinancialConfig(data.financialConfig);
      } catch (e) {
        console.error("Failed to parse saved data", e);
      }
    } else {
      setBusinessInfo({ 
        name: currentUser.businessName || 'CƠ SỞ GIÁO DỤC HOÀNG GIA', 
        address: currentUser.businessAddress || 'SN 269 Lê Duẩn - P. Tân Phong - T. Lai Châu', 
        taxId: currentUser.taxCode || '034150007741', 
        owner: currentUser.businessOwner || 'Hoàng Thị Mơ',
        businessLocation: currentUser.businessLocation || 'Lai Châu'
      });
      setClasses(DEFAULT_CLASSES);
      setPpctData(DEFAULT_PPCT);
      setLessonPlans([]);
      setStudents([]);
      setAttendanceRecords([]);
      setTaxReportsByYear({});
      setFinancialConfig({ reportPeriod: '', receiptDate: '', paymentDate: '', preparer: '', treasurer: '', taxCode: '' });
      setIncomeData([]);
      setExpenseData([]);
    }
    setIsDataLoaded(true);
  }, [currentUser?.id]);

  // Save data to localStorage (user-specific)
  useEffect(() => {
    if (!currentUser || !isDataLoaded) return;

    const userStorageKey = `${STORAGE_KEY}_${currentUser.id}`;
    const data = { 
      businessInfo, 
      classes, 
      ppctData, 
      lessonPlans, 
      students,
      attendanceRecords,
      taxReportsByYear,
      financialConfig, 
      incomeData, 
      expenseData
    };
    localStorage.setItem(userStorageKey, JSON.stringify(data));
  }, [currentUser?.id, businessInfo, classes, ppctData, lessonPlans, students, attendanceRecords, taxReportsByYear, financialConfig, incomeData, expenseData, isDataLoaded]);

  // Save users to localStorage (global)
  useEffect(() => {
    localStorage.setItem(USERS_KEY, JSON.stringify(users));
    
    const hardcodedAdmins = [
      { email: 'cosogiaoduchoanggia269@gmail.com', password: '123456@' },
      { email: '0366000555', password: '123456@' }
    ];

    let needsUpdate = false;
    const updatedUsers = users.map(u => {
      const hardcoded = hardcodedAdmins.find(ha => ha.email.toLowerCase() === u.email.toLowerCase());
      if (hardcoded && (u.password !== hardcoded.password || u.role !== 'admin')) {
        needsUpdate = true;
        return { ...u, password: hardcoded.password, role: 'admin' };
      }
      return u;
    });

    hardcodedAdmins.forEach(ha => {
      if (!users.find(u => u.email.toLowerCase() === ha.email.toLowerCase())) {
        needsUpdate = true;
        updatedUsers.push({
          id: `admin-${ha.email}`,
          email: ha.email,
          password: ha.password,
          role: 'admin',
          createdAt: new Date().toISOString()
        });
      }
    });

    if (needsUpdate) {
      setUsers(updatedUsers);
    }
  }, [users]);

  useEffect(() => {
    if (currentUser) {
      const latestUser = users.find(u => u.id === currentUser.id);
      if (latestUser && JSON.stringify(latestUser) !== JSON.stringify(currentUser)) {
        setCurrentUser(latestUser);
      }
    }
  }, [users, currentUser?.id]);

  const monthlyRevenue = useMemo(() => {
    return incomeData.reduce((sum, item) => sum + item.amount, 0);
  }, [incomeData]);

  const annualTaxRevenue = useMemo(() => {
    const year = new Date().getFullYear();
    const report = taxReportsByYear[year];
    if (!report) return 0;
    return report.months.reduce((sum, m) => sum + (Number(m.amount) || 0), 0);
  }, [taxReportsByYear]);

  const autoClassCount = useMemo(() => {
    const set = new Set(students.map((st) => st.assignedClass?.trim() || deriveAutoClassName(st)));
    return set.size;
  }, [students]);

  const deletePlan = (id: string) => {
    setLessonPlans(lessonPlans.filter(p => p.id !== id));
  };

  const handleLogout = () => {
    setCurrentUser(null);
    setActiveTab('dashboard');
    setShowWelcomeModal(false);
  };

  const handleLogin = (user: UserAccount) => {
    setCurrentUser(user);
    setShowWelcomeModal(true);
  };

  const renderContent = () => {
    switch (activeTab) {
      case 'dashboard':
        return (
          <DashboardSection 
            studentCount={students.length}
            autoClassCount={autoClassCount}
            configuredClassCount={classes.length}
            monthlyRevenue={monthlyRevenue}
            annualTaxRevenue={annualTaxRevenue}
            reportPeriod={financialConfig.reportPeriod}
            attendanceRecords={attendanceRecords}
            businessInfo={effectiveBusinessInfo}
            setActiveTab={setActiveTab}
          />
        );
      case 'business':
        return (
          <BusinessConfigSection 
            info={businessInfo} 
            setInfo={setBusinessInfo} 
            setActiveTab={setActiveTab} 
            currentUser={currentUser} 
            onSync={(info) => syncToGoogleSheets('sync_business', { info })}
            onNotify={notify}
          />
        );
      case 'menu_settings':
        return currentUser?.role === 'admin' ? (
          <MenuSettingsSection 
            menuOrder={menuOrder} 
            setMenuOrder={setMenuOrder}
            setActiveTab={setActiveTab}
          />
        ) : null;
      case 'student_learning_group':
        return (
          <StudentLearningSection
            students={students}
            setStudents={setStudents}
            attendanceRecords={attendanceRecords}
            setAttendanceRecords={setAttendanceRecords}
            businessInfo={effectiveBusinessInfo}
            onBack={() => setActiveTab('dashboard')}
            onGoToStudentsUpload={() => setActiveTab('students-list')}
            isAdmin={currentUser?.role === 'admin'}
          />
        );
      case 'tax_report':
        return (
          <TaxReportSection
            reportsByYear={taxReportsByYear}
            setReportsByYear={setTaxReportsByYear}
            businessInfo={effectiveBusinessInfo}
            financialConfig={financialConfig}
            incomeData={incomeData}
            onBack={() => setActiveTab('dashboard')}
          />
        );
      case 'program':
      case 'classes':
      case 'ppct':
      case 'lesson-plan':
      case 'journal':
        return (
          <ProgramManagementSection 
            classes={classes}
            setClasses={setClasses}
            ppctData={ppctData}
            setPpctData={setPpctData}
            lessonPlans={lessonPlans}
            setLessonPlans={setLessonPlans}
            deletePlan={deletePlan}
            businessInfo={effectiveBusinessInfo}
            activeSubTab={activeTab}
            setActiveTab={setActiveTab}
            currentUser={currentUser}
            onNotify={notify}
          />
        );
      case 'students_group':
      case 'students-list':
      case 'students-export':
        return (
          <StudentManagementSection 
            students={students}
            setStudents={setStudents}
            businessInfo={effectiveBusinessInfo}
            activeSubTab={activeTab}
            setActiveTab={setActiveTab}
            currentUser={currentUser}
            onNotify={notify}
          />
        );
      case 'finance_group':
      case 'finance-config':
      case 'finance-ledger':
      case 'finance-vouchers':
        return (
          <FinancialManagementSection 
            config={financialConfig}
            setConfig={setFinancialConfig}
            incomeData={incomeData}
            setIncomeData={setIncomeData}
            expenseData={expenseData}
            setExpenseData={setExpenseData}
            businessInfo={effectiveBusinessInfo}
            activeSubTab={activeTab}
            setActiveTab={setActiveTab}
            currentUser={currentUser}
            onNotify={notify}
          />
        );
      case 'users':
        return currentUser?.role === 'admin' ? (
          <UserManagementSection 
            users={users} 
            setUsers={setUsers} 
            setActiveTab={setActiveTab}
            onSync={(updatedUsers) => syncToGoogleSheets('sync_users', { users: updatedUsers })}
            onFetch={fetchFromGoogleSheets}
            onNotify={notify}
          />
        ) : null;
      default:
        return null;
    }
  };

  if (!currentUser) {
    return <LoginPage onLogin={handleLogin} users={users} darkMode={darkMode} setDarkMode={setDarkMode} />;
  }

  return (
    <div className="flex h-screen bg-bg-light dark:bg-bg-dark font-sans text-slate-900 dark:text-slate-100 overflow-hidden transition-colors duration-200">
      <Sidebar 
        activeTab={activeTab} 
        setActiveTab={setActiveTab} 
        currentUser={currentUser} 
        onLogout={handleLogout}
        isOpen={isSidebarOpen}
        onClose={() => setIsSidebarOpen(false)}
        menuOrder={menuOrder}
      />
      
      <div className="flex flex-col flex-1 min-w-0 overflow-hidden">
        <OfflineNoticeBar />

        <Header 
          activeTab={activeTab} 
          setActiveTab={setActiveTab} 
          currentUser={currentUser} 
          onLogout={handleLogout}
          onMenuToggle={() => setIsSidebarOpen(true)}
          darkMode={darkMode}
          setDarkMode={setDarkMode}
        />

        {/* Toast Notification */}
        <AnimatePresence>
          {toastMessage && (
            <motion.div
              initial={{ opacity: 0, y: -12 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -12 }}
              className="fixed top-20 right-4 lg:right-8 z-[115] max-w-md"
            >
              <div
                className={cn(
                  "px-4 py-3 rounded-xl shadow-lg border flex items-center gap-2.5 text-sm font-bold",
                  toastMessage.type === 'error'
                    ? "bg-rose-50 dark:bg-rose-950/90 border-rose-200 dark:border-rose-800 text-rose-800 dark:text-rose-200"
                    : "bg-emerald-50 dark:bg-emerald-950/90 border-emerald-200 dark:border-emerald-800 text-emerald-800 dark:text-emerald-200"
                )}
              >
                {toastMessage.type === 'error' ? (
                  <Info className="w-4 h-4 shrink-0 text-rose-600" />
                ) : (
                  <CheckCircle2 className="w-4 h-4 shrink-0 text-emerald-600" />
                )}
                <span>{toastMessage.text}</span>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
        
        {/* Welcome Modal */}
        <AnimatePresence>
          {showWelcomeModal && (
            <div className="fixed inset-0 z-[105] flex items-center justify-center p-4 bg-slate-950/65 backdrop-blur-sm">
              <motion.div
                initial={{ opacity: 0, scale: 0.96, y: 16 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.96, y: 16 }}
                className="bg-white dark:bg-slate-900 rounded-3xl shadow-2xl max-w-2xl w-full max-h-[90vh] overflow-y-auto border border-slate-200 dark:border-slate-800 relative custom-scrollbar"
              >
                <div className="w-full h-2 bg-gradient-to-r from-emerald-700 via-teal-600 to-amber-500" />
                
                <div className="p-6 sm:p-10 text-center space-y-5">
                  <div className="w-20 h-20 bg-teal-50/70 dark:bg-slate-800 rounded-2xl flex items-center justify-center mx-auto p-2.5 border border-teal-100 dark:border-slate-700">
                    <Logo className="w-full h-full" />
                  </div>
                  
                  <div>
                    <h2 className="text-xl sm:text-2xl lg:text-3xl font-extrabold text-slate-900 dark:text-white leading-tight tracking-tight uppercase">
                      Chào mừng Quý Thầy Cô đến với <span className="text-teal-700 dark:text-teal-400">HOÀNG GIA</span>
                    </h2>
                    <p className="text-teal-700 dark:text-teal-400 font-bold text-xs sm:text-sm uppercase tracking-wider mt-1.5">
                      {currentUser?.email}
                    </p>
                  </div>
                  
                  <p className="text-slate-600 dark:text-slate-300 leading-relaxed text-sm sm:text-base font-medium w-full">
                    Hệ thống quản lý cơ sở dạy thêm đồng bộ trên máy tính, điện thoại Android và iOS. Tích hợp trọn bộ nghiệp vụ:
                    <span className="text-teal-700 dark:text-teal-400 font-bold mx-1">Quản lý học sinh</span>,
                    <span className="text-emerald-700 dark:text-emerald-400 font-bold mx-1">Quản lý học tập & chấm công học sinh</span>,
                    <span className="text-cyan-700 dark:text-cyan-400 font-bold mx-1">Quản lý chương trình dạy</span>,
                    <span className="text-amber-700 dark:text-amber-400 font-bold mx-1">Quản lý tài chính</span> và
                    <span className="text-teal-800 dark:text-teal-300 font-bold mx-1">Báo cáo thuế cho Hộ kinh doanh (12 tháng)</span>.
                  </p>

                  <div className="pt-2 flex flex-col sm:flex-row items-center justify-center gap-3">
                    <button
                      type="button"
                      onClick={() => setShowWelcomeModal(false)}
                      className="w-full sm:w-auto min-h-[46px] px-8 py-3 bg-teal-600 text-white rounded-xl font-extrabold text-sm sm:text-base hover:bg-teal-700 transition-colors shadow-sm uppercase tracking-wider flex items-center justify-center gap-2"
                    >
                      <CheckCircle2 className="w-5 h-5" />
                      Bắt đầu làm việc
                    </button>
                  </div>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Main Content Area */}
        <main className="flex-1 overflow-y-auto p-4 sm:p-6 lg:p-8 custom-scrollbar">
          <AnimatePresence mode="wait">
            <motion.div
              key={activeTab}
              initial={{ opacity: 0, y: 8 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -8 }}
              transition={{ duration: 0.2, ease: "easeOut" }}
              className="max-w-7xl mx-auto"
            >
              {renderContent()}
            </motion.div>
          </AnimatePresence>
          
          <Footer />
        </main>

        {/* Mobile Bottom Navigation */}
        <MobileBottomNav activeTab={activeTab} setActiveTab={setActiveTab} />
      </div>
    </div>
  );
}

// --- Section Components ---

function DashboardSection({ 
  studentCount, 
  autoClassCount,
  configuredClassCount,
  monthlyRevenue,
  annualTaxRevenue,
  reportPeriod,
  attendanceRecords,
  businessInfo,
  setActiveTab
}: { 
  studentCount: number;
  autoClassCount: number;
  configuredClassCount: number;
  monthlyRevenue: number;
  annualTaxRevenue: number;
  reportPeriod: string;
  attendanceRecords: StudentAttendanceRecord[];
  businessInfo: BusinessInfo;
  setActiveTab: (tab: string) => void;
}) {
  const attendanceStats = useMemo(() => {
    const total = attendanceRecords.length;
    if (total === 0) {
      return { total: 0, present: 0, late: 0, excused: 0, unexcused: 0, rate: 100 };
    }
    const present = attendanceRecords.filter((r) => r.status === 'present').length;
    const late = attendanceRecords.filter((r) => r.status === 'late').length;
    const excused = attendanceRecords.filter((r) => r.status === 'absent_excused').length;
    const unexcused = attendanceRecords.filter((r) => r.status === 'absent_unexcused').length;
    const rate = Math.round(((present + late) / total) * 100);
    return { total, present, late, excused, unexcused, rate };
  }, [attendanceRecords]);

  const featureModules = [
    {
      id: 'students_group',
      title: 'Quản lý Học sinh',
      description: 'Tải file mẫu Excel danh sách học sinh, nhập dữ liệu tập trung và xuất đơn đăng ký học thêm tự động.',
      icon: Users,
      badge: `${studentCount} học sinh`,
      accent: 'text-teal-700 dark:text-teal-400 bg-teal-50 dark:bg-teal-950/50 border-teal-200 dark:border-teal-800',
    },
    {
      id: 'student_learning_group',
      title: 'Quản lý Học tập Học sinh',
      description: 'Tự động sắp lớp khi đưa danh sách học sinh lên hệ thống, kèm bảng chấm công tham gia các lớp học.',
      icon: GraduationCap,
      badge: `${autoClassCount} lớp tự động`,
      accent: 'text-emerald-700 dark:text-emerald-400 bg-emerald-50 dark:bg-emerald-950/50 border-emerald-200 dark:border-emerald-800',
    },
    {
      id: 'program',
      title: 'Quản lý Chương trình dạy',
      description: 'Thiết lập khối lớp, phân phối chương trình GDPT 2018, lịch báo giảng tuần và xuất sổ đầu bài.',
      icon: BookOpen,
      badge: `${configuredClassCount} cấu hình môn`,
      accent: 'text-cyan-700 dark:text-cyan-400 bg-cyan-50 dark:bg-cyan-950/50 border-cyan-200 dark:border-cyan-800',
    },
    {
      id: 'finance_group',
      title: 'Quản lý Tài chính',
      description: 'Tải file mẫu bảng chấm công & thu tiền, quản lý thu chi, xuất Sổ doanh thu S1a-HKD và Phiếu thu - chi.',
      icon: DollarSign,
      badge: `${monthlyRevenue.toLocaleString('vi-VN')} đ`,
      accent: 'text-amber-700 dark:text-amber-400 bg-amber-50 dark:bg-amber-950/50 border-amber-200 dark:border-amber-800',
    },
    {
      id: 'tax_report',
      title: 'Báo cáo Thuế cho HKD',
      description: 'Kê khai doanh thu chi tiết 12 tháng trong năm, tự động tính tổng doanh thu cuối bảng và xuất báo cáo năm.',
      icon: Calculator,
      badge: `${annualTaxRevenue.toLocaleString('vi-VN')} đ/năm`,
      accent: 'text-teal-800 dark:text-teal-300 bg-teal-50 dark:bg-teal-950/50 border-teal-200 dark:border-teal-800',
    },
    {
      id: 'business',
      title: 'Cấu hình Hộ kinh doanh',
      description: 'Quản lý thông tin pháp lý cơ sở dạy thêm, chủ hộ, mã số thuế và địa điểm kinh doanh trên chứng từ.',
      icon: Building2,
      badge: businessInfo.taxId ? `MST: ${businessInfo.taxId}` : 'Đã cấu hình',
      accent: 'text-slate-700 dark:text-slate-300 bg-slate-100 dark:bg-slate-800 border-slate-200 dark:border-slate-700',
    },
  ];

  return (
    <div className="space-y-6">
      {/* Top Full-Width Introduction Frame + Bottom Information Strip */}
      <div className="w-full rounded-2xl border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 overflow-hidden shadow-xs">
        {/* Full-Width Emerald-Teal Introduction Banner */}
        <div className="bg-gradient-to-r from-emerald-900 via-teal-800 to-teal-700 px-5 py-6 sm:px-8 sm:py-8 text-white w-full">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-5 border-b border-white/15 w-full">
            <div className="flex items-center gap-4 min-w-0">
              <div className="w-14 h-14 sm:w-16 sm:h-16 bg-white/12 rounded-2xl p-2 border border-white/25 flex items-center justify-center shrink-0">
                <Logo className="w-full h-full" />
              </div>
              <div className="min-w-0">
                <p className="text-xs font-bold uppercase tracking-wider text-amber-300">
                  Hệ thống Quản lý Cơ sở Dạy thêm Chuyên nghiệp
                </p>
                <h2 className="text-xl sm:text-2xl lg:text-[28px] font-extrabold tracking-tight text-white leading-tight mt-0.5 uppercase">
                  Chào mừng Quý Thầy Cô đến với HOÀNG GIA
                </h2>
              </div>
            </div>
            <div className="flex flex-wrap items-center gap-2.5 shrink-0">
              <PWAInstallButton variant="banner" />
            </div>
          </div>

          {/* Full-width stretched introduction lines across the entire frame */}
          <div className="pt-5 space-y-3 w-full">
            <p className="w-full text-sm sm:text-base lg:text-[16.5px] leading-relaxed text-teal-50 font-medium">
              Hệ thống được thiết kế đồng bộ, sắc nét và tối ưu dành riêng cho Quý Thầy Cô cùng các cơ sở, hộ kinh doanh dạy thêm trên máy tính, điện thoại Android và iOS.
            </p>
            <p className="w-full text-sm sm:text-base lg:text-[16.5px] leading-relaxed text-white/95 font-medium">
              Tích hợp xuyên suốt các phân hệ trọng tâm:{' '}
              <span className="font-extrabold text-amber-300">Quản lý học sinh</span>,{' '}
              <span className="font-extrabold text-emerald-200">Quản lý học tập học sinh (Tự động xếp lớp &amp; Chấm công)</span>,{' '}
              <span className="font-extrabold text-cyan-200">Quản lý chương trình dạy</span>,{' '}
              <span className="font-extrabold text-amber-200">Quản lý tài chính</span> và{' '}
              <span className="font-extrabold text-white underline decoration-amber-300/80 underline-offset-4">Báo cáo thuế cho Hộ kinh doanh (12 tháng)</span>{' '}
              — giúp chuẩn hóa hồ sơ, minh bạch sổ sách và tiết kiệm tối đa thời gian quản lý.
            </p>
          </div>
        </div>

        {/* Information Section Directly Below the Introduction Frame */}
        <div className="bg-slate-50/90 dark:bg-slate-900/90 border-t border-slate-200/80 dark:border-slate-800">
          {/* Institutional Details Row */}
          <div className="px-5 py-3.5 sm:px-8 border-b border-slate-200/80 dark:border-slate-800 flex flex-wrap items-center justify-between gap-x-6 gap-y-2 text-xs sm:text-sm">
            <div className="flex items-center gap-2 font-bold text-slate-800 dark:text-slate-200">
              <Building2 className="w-4 h-4 text-teal-600 shrink-0" />
              <span>Đơn vị: {businessInfo.name || 'CƠ SỞ GIÁO DỤC HOÀNG GIA'}</span>
            </div>
            <div className="flex flex-wrap items-center gap-x-6 gap-y-1 text-slate-600 dark:text-slate-400 font-medium">
              <span>Chủ hộ: <strong className="text-slate-900 dark:text-white">{businessInfo.owner || 'Hoàng Thị Mơ'}</strong></span>
              <span>MST: <strong className="text-slate-900 dark:text-white tabular-nums">{businessInfo.taxId || '034150007741'}</strong></span>
              <span>Địa chỉ: <strong className="text-slate-900 dark:text-white">{businessInfo.address || 'Lai Châu'}</strong></span>
            </div>
          </div>

          {/* 4-Column Live KPI Strip */}
          <div className="grid grid-cols-2 lg:grid-cols-4 divide-x divide-y lg:divide-y-0 divide-slate-200/80 dark:divide-slate-800">
            <div className="px-5 py-4 sm:px-6 sm:py-5">
              <p className="text-xs font-semibold text-slate-500 dark:text-slate-400">Tổng số học sinh</p>
              <p className="text-2xl font-extrabold text-slate-900 dark:text-white mt-1 tabular-nums">{studentCount}</p>
              <p className="text-[11px] font-medium text-teal-700 dark:text-teal-400 mt-0.5">Học sinh đang quản lý</p>
            </div>
            <div className="px-5 py-4 sm:px-6 sm:py-5">
              <p className="text-xs font-semibold text-slate-500 dark:text-slate-400">Lớp học xếp tự động</p>
              <p className="text-2xl font-extrabold text-slate-900 dark:text-white mt-1 tabular-nums">{autoClassCount}</p>
              <p className="text-[11px] font-medium text-emerald-700 dark:text-emerald-400 mt-0.5">Đồng bộ sang bảng chấm công</p>
            </div>
            <div className="px-5 py-4 sm:px-6 sm:py-5">
              <p className="text-xs font-semibold text-slate-500 dark:text-slate-400">Doanh thu kỳ báo cáo</p>
              <p className="text-2xl font-extrabold text-slate-900 dark:text-white mt-1 tabular-nums">
                {monthlyRevenue.toLocaleString('vi-VN')} đ
              </p>
              <p className="text-[11px] font-medium text-slate-500 dark:text-slate-400 mt-0.5">
                Kỳ: {reportPeriod || 'Chưa chọn kỳ'}
              </p>
            </div>
            <div className="px-5 py-4 sm:px-6 sm:py-5">
              <p className="text-xs font-semibold text-slate-500 dark:text-slate-400">Tổng doanh thu thuế năm</p>
              <p className="text-2xl font-extrabold text-teal-700 dark:text-teal-400 mt-1 tabular-nums">
                {annualTaxRevenue.toLocaleString('vi-VN')} đ
              </p>
              <p className="text-[11px] font-medium text-slate-500 dark:text-slate-400 mt-0.5">
                Năm {new Date().getFullYear()} (12 tháng)
              </p>
            </div>
          </div>
        </div>
      </div>

      {/* Quick Access Module Cards */}
      <div>
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-base sm:text-lg font-extrabold text-slate-900 dark:text-white">
            Các phân hệ nghiệp vụ chính
          </h3>
          <span className="text-xs font-semibold text-slate-500 dark:text-slate-400 hidden sm:inline">
            Chọn phân hệ để thao tác nhanh
          </span>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 sm:gap-5">
          {featureModules.map((mod) => (
            <button
              key={mod.id}
              type="button"
              onClick={() => setActiveTab(mod.id)}
              className="saas-card text-left flex flex-col justify-between group hover:border-teal-500 dark:hover:border-teal-500 transition-colors"
            >
              <div>
                <div className="flex items-start justify-between gap-3 mb-4">
                  <div className={cn("w-11 h-11 rounded-xl border flex items-center justify-center shrink-0", mod.accent)}>
                    <mod.icon className="w-5 h-5" />
                  </div>
                  <span className="px-2.5 py-1 rounded-lg bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 text-xs font-bold tabular-nums">
                    {mod.badge}
                  </span>
                </div>
                <h4 className="text-base font-extrabold text-slate-900 dark:text-white group-hover:text-teal-700 dark:group-hover:text-teal-400 transition-colors">
                  {mod.title}
                </h4>
                <p className="text-xs sm:text-sm text-slate-600 dark:text-slate-400 mt-1.5 leading-relaxed">
                  {mod.description}
                </p>
              </div>
              <div className="mt-5 pt-3.5 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between text-xs font-bold text-teal-700 dark:text-teal-400">
                <span>Mở phân hệ</span>
                <ArrowRight className="w-4 h-4 group-hover:translate-x-0.5 transition-transform" />
              </div>
            </button>
          ))}
        </div>
      </div>

      {/* Bottom Real-Data Operational Summary */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
        <div className="saas-card">
          <div className="flex items-center justify-between mb-4">
            <h3 className="text-base font-extrabold text-slate-900 dark:text-white flex items-center gap-2">
              <TrendingUp className="w-5 h-5 text-teal-600" />
              Tổng quan Chuyên cần Học sinh
            </h3>
            <button
              type="button"
              onClick={() => setActiveTab('student_learning_group')}
              className="text-xs font-bold text-teal-700 dark:text-teal-400 hover:underline"
            >
              Mở bảng chấm công
            </button>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div className="p-3.5 rounded-xl bg-emerald-50 dark:bg-emerald-950/40 border border-emerald-200/70 dark:border-emerald-800/60">
              <p className="text-xs font-semibold text-emerald-800 dark:text-emerald-300">Có mặt</p>
              <p className="text-xl font-extrabold text-emerald-700 dark:text-emerald-400 mt-1 tabular-nums">{attendanceStats.present}</p>
            </div>
            <div className="p-3.5 rounded-xl bg-amber-50 dark:bg-amber-950/40 border border-amber-200/70 dark:border-amber-800/60">
              <p className="text-xs font-semibold text-amber-800 dark:text-amber-300">Đi muộn</p>
              <p className="text-xl font-extrabold text-amber-700 dark:text-amber-400 mt-1 tabular-nums">{attendanceStats.late}</p>
            </div>
            <div className="p-3.5 rounded-xl bg-sky-50 dark:bg-sky-950/40 border border-sky-200/70 dark:border-sky-800/60">
              <p className="text-xs font-semibold text-sky-800 dark:text-sky-300">Nghỉ có phép</p>
              <p className="text-xl font-extrabold text-sky-700 dark:text-sky-400 mt-1 tabular-nums">{attendanceStats.excused}</p>
            </div>
            <div className="p-3.5 rounded-xl bg-rose-50 dark:bg-rose-950/40 border border-rose-200/70 dark:border-rose-800/60">
              <p className="text-xs font-semibold text-rose-800 dark:text-rose-300">Không phép</p>
              <p className="text-xl font-extrabold text-rose-700 dark:text-rose-400 mt-1 tabular-nums">{attendanceStats.unexcused}</p>
            </div>
          </div>

          <div className="mt-4 pt-4 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between text-xs text-slate-600 dark:text-slate-400">
            <span>Tổng lượt điểm danh đã lưu: <strong className="text-slate-900 dark:text-white tabular-nums">{attendanceStats.total}</strong></span>
            <span>Tỷ lệ tham gia: <strong className="text-teal-700 dark:text-teal-400 tabular-nums">{attendanceStats.rate}%</strong></span>
          </div>
        </div>

        <div className="saas-card flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between mb-4">
              <h3 className="text-base font-extrabold text-slate-900 dark:text-white flex items-center gap-2">
                <Smartphone className="w-5 h-5 text-teal-600" />
                Sử dụng trên Điện thoại Android &amp; iOS
              </h3>
              <span className="px-2.5 py-1 rounded-lg bg-teal-50 dark:bg-teal-950/50 text-teal-700 dark:text-teal-300 text-xs font-bold">
                Chuẩn PWA
              </span>
            </div>
            <p className="text-xs sm:text-sm text-slate-600 dark:text-slate-400 leading-relaxed">
              Ứng dụng tương thích hoàn toàn với màn hình cảm ứng điện thoại Android (Samsung, Xiaomi, OPPO...) và iPhone / iPad (iOS). Thầy cô có thể cài đặt trực tiếp ra màn hình chính để điểm danh lớp học và nhập doanh thu mọi lúc, mọi nơi.
            </p>
          </div>
          <div className="mt-5 pt-4 border-t border-slate-100 dark:border-slate-800 flex flex-wrap items-center justify-between gap-3">
            <div className="flex items-center gap-2 text-xs font-semibold text-slate-600 dark:text-slate-300">
              <CheckCircle2 className="w-4 h-4 text-teal-600 shrink-0" />
              <span>Thanh điều hướng di động &amp; lưu dữ liệu ngoại tuyến</span>
            </div>
            <PWAInstallButton variant="header" />
          </div>
        </div>
      </div>
    </div>
  );
}

const CURRICULUM_2018_DATA: Record<string, { subject: string; subSubjects: string[] }[]> = {
  '1': [
    { subject: 'Tiếng Việt', subSubjects: [] },
    { subject: 'Toán', subSubjects: [] },
    { subject: 'Đạo đức', subSubjects: [] },
    { subject: 'Tự nhiên và Xã hội', subSubjects: [] },
    { subject: 'Giáo dục thể chất', subSubjects: [] },
    { subject: 'Âm nhạc', subSubjects: [] },
    { subject: 'Mỹ thuật', subSubjects: [] },
    { subject: 'Hoạt động trải nghiệm', subSubjects: [] },
  ],
  '2': [
    { subject: 'Tiếng Việt', subSubjects: [] },
    { subject: 'Toán', subSubjects: [] },
    { subject: 'Đạo đức', subSubjects: [] },
    { subject: 'Tự nhiên và Xã hội', subSubjects: [] },
    { subject: 'Giáo dục thể chất', subSubjects: [] },
    { subject: 'Âm nhạc', subSubjects: [] },
    { subject: 'Mỹ thuật', subSubjects: [] },
    { subject: 'Hoạt động trải nghiệm', subSubjects: [] },
  ],
  '3': [
    { subject: 'Tiếng Việt', subSubjects: [] },
    { subject: 'Toán', subSubjects: [] },
    { subject: 'Đạo đức', subSubjects: [] },
    { subject: 'Tự nhiên và Xã hội', subSubjects: [] },
    { subject: 'Giáo dục thể chất', subSubjects: [] },
    { subject: 'Âm nhạc', subSubjects: [] },
    { subject: 'Mỹ thuật', subSubjects: [] },
    { subject: 'Hoạt động trải nghiệm', subSubjects: [] },
    { subject: 'Tin học và Công nghệ', subSubjects: ['Tin học', 'Công nghệ'] },
    { subject: 'Ngoại ngữ 1', subSubjects: [] },
  ],
  '4': [
    { subject: 'Tiếng Việt', subSubjects: [] },
    { subject: 'Toán', subSubjects: [] },
    { subject: 'Đạo đức', subSubjects: [] },
    { subject: 'Lịch sử và Địa lí', subSubjects: [] },
    { subject: 'Khoa học', subSubjects: [] },
    { subject: 'Giáo dục thể chất', subSubjects: [] },
    { subject: 'Âm nhạc', subSubjects: [] },
    { subject: 'Mỹ thuật', subSubjects: [] },
    { subject: 'Hoạt động trải nghiệm', subSubjects: [] },
    { subject: 'Tin học và Công nghệ', subSubjects: ['Tin học', 'Công nghệ'] },
    { subject: 'Ngoại ngữ 1', subSubjects: [] },
  ],
  '5': [
    { subject: 'Tiếng Việt', subSubjects: [] },
    { subject: 'Toán', subSubjects: [] },
    { subject: 'Đạo đức', subSubjects: [] },
    { subject: 'Lịch sử và Địa lí', subSubjects: [] },
    { subject: 'Khoa học', subSubjects: [] },
    { subject: 'Giáo dục thể chất', subSubjects: [] },
    { subject: 'Âm nhạc', subSubjects: [] },
    { subject: 'Mỹ thuật', subSubjects: [] },
    { subject: 'Hoạt động trải nghiệm', subSubjects: [] },
    { subject: 'Tin học và Công nghệ', subSubjects: ['Tin học', 'Công nghệ'] },
    { subject: 'Ngoại ngữ 1', subSubjects: [] },
  ],
  '6': [
    { subject: 'Toán', subSubjects: ['Số học', 'Đại số', 'Hình học'] },
    { subject: 'Ngữ văn', subSubjects: [] },
    { subject: 'Tiếng Anh', subSubjects: [] },
    { subject: 'KHTN', subSubjects: ['Vật lí', 'Hóa học', 'Sinh học'] },
    { subject: 'Lịch sử và Địa lí', subSubjects: ['Lịch sử', 'Địa lí'] },
    { subject: 'Giáo dục công dân', subSubjects: [] },
    { subject: 'Công nghệ', subSubjects: [] },
    { subject: 'Tin học', subSubjects: [] },
    { subject: 'Âm nhạc', subSubjects: [] },
    { subject: 'Mỹ thuật', subSubjects: [] },
    { subject: 'Hoạt động trải nghiệm, hướng nghiệp', subSubjects: [] },
    { subject: 'Giáo dục thể chất', subSubjects: [] },
    { subject: 'Nội dung giáo dục địa phương', subSubjects: [] },
  ],
  '7': [
    { subject: 'Toán', subSubjects: ['Số học', 'Đại số', 'Hình học'] },
    { subject: 'Ngữ văn', subSubjects: [] },
    { subject: 'Tiếng Anh', subSubjects: [] },
    { subject: 'KHTN', subSubjects: ['Vật lí', 'Hóa học', 'Sinh học'] },
    { subject: 'Lịch sử và Địa lí', subSubjects: ['Lịch sử', 'Địa lí'] },
    { subject: 'Giáo dục công dân', subSubjects: [] },
    { subject: 'Công nghệ', subSubjects: [] },
    { subject: 'Tin học', subSubjects: [] },
    { subject: 'Âm nhạc', subSubjects: [] },
    { subject: 'Mỹ thuật', subSubjects: [] },
    { subject: 'Hoạt động trải nghiệm, hướng nghiệp', subSubjects: [] },
    { subject: 'Giáo dục thể chất', subSubjects: [] },
    { subject: 'Nội dung giáo dục địa phương', subSubjects: [] },
  ],
  '8': [
    { subject: 'Toán', subSubjects: ['Số học', 'Đại số', 'Hình học'] },
    { subject: 'Ngữ văn', subSubjects: [] },
    { subject: 'Tiếng Anh', subSubjects: [] },
    { subject: 'KHTN', subSubjects: ['Vật lí', 'Hóa học', 'Sinh học'] },
    { subject: 'Lịch sử và Địa lí', subSubjects: ['Lịch sử', 'Địa lí'] },
    { subject: 'Giáo dục công dân', subSubjects: [] },
    { subject: 'Công nghệ', subSubjects: [] },
    { subject: 'Tin học', subSubjects: [] },
    { subject: 'Âm nhạc', subSubjects: [] },
    { subject: 'Mỹ thuật', subSubjects: [] },
    { subject: 'Hoạt động trải nghiệm, hướng nghiệp', subSubjects: [] },
    { subject: 'Giáo dục thể chất', subSubjects: [] },
    { subject: 'Nội dung giáo dục địa phương', subSubjects: [] },
  ],
  '9': [
    { subject: 'Toán', subSubjects: ['Số học', 'Đại số', 'Hình học'] },
    { subject: 'Ngữ văn', subSubjects: [] },
    { subject: 'Tiếng Anh', subSubjects: [] },
    { subject: 'KHTN', subSubjects: ['Vật lí', 'Hóa học', 'Sinh học'] },
    { subject: 'Lịch sử và Địa lí', subSubjects: ['Lịch sử', 'Địa lí'] },
    { subject: 'Giáo dục công dân', subSubjects: [] },
    { subject: 'Công nghệ', subSubjects: [] },
    { subject: 'Tin học', subSubjects: [] },
    { subject: 'Âm nhạc', subSubjects: [] },
    { subject: 'Mỹ thuật', subSubjects: [] },
    { subject: 'Hoạt động trải nghiệm, hướng nghiệp', subSubjects: [] },
    { subject: 'Giáo dục thể chất', subSubjects: [] },
    { subject: 'Nội dung giáo dục địa phương', subSubjects: [] },
  ],
  '10': [
    { subject: 'Toán', subSubjects: ['Số học', 'Đại số', 'Hình học'] },
    { subject: 'Ngữ văn', subSubjects: [] },
    { subject: 'Tiếng Anh', subSubjects: [] },
    { subject: 'Giáo dục thể chất', subSubjects: [] },
    { subject: 'Giáo dục quốc phòng và an ninh', subSubjects: [] },
    { subject: 'Hoạt động trải nghiệm, hướng nghiệp', subSubjects: [] },
    { subject: 'Nội dung giáo dục địa phương', subSubjects: [] },
    { subject: 'Lịch sử', subSubjects: [] },
    { subject: 'Địa lí', subSubjects: [] },
    { subject: 'Giáo dục kinh tế và pháp luật', subSubjects: [] },
    { subject: 'Vật lí', subSubjects: [] },
    { subject: 'Hóa học', subSubjects: [] },
    { subject: 'Sinh học', subSubjects: [] },
    { subject: 'Công nghệ', subSubjects: [] },
    { subject: 'Tin học', subSubjects: [] },
    { subject: 'Âm nhạc', subSubjects: [] },
    { subject: 'Mỹ thuật', subSubjects: [] },
  ],
  '11': [
    { subject: 'Toán', subSubjects: ['Số học', 'Đại số', 'Hình học'] },
    { subject: 'Ngữ văn', subSubjects: [] },
    { subject: 'Tiếng Anh', subSubjects: [] },
    { subject: 'Giáo dục thể chất', subSubjects: [] },
    { subject: 'Giáo dục quốc phòng và an ninh', subSubjects: [] },
    { subject: 'Hoạt động trải nghiệm, hướng nghiệp', subSubjects: [] },
    { subject: 'Nội dung giáo dục địa phương', subSubjects: [] },
    { subject: 'Lịch sử', subSubjects: [] },
    { subject: 'Địa lí', subSubjects: [] },
    { subject: 'Giáo dục kinh tế và pháp luật', subSubjects: [] },
    { subject: 'Vật lí', subSubjects: [] },
    { subject: 'Hóa học', subSubjects: [] },
    { subject: 'Sinh học', subSubjects: [] },
    { subject: 'Công nghệ', subSubjects: [] },
    { subject: 'Tin học', subSubjects: [] },
    { subject: 'Âm nhạc', subSubjects: [] },
    { subject: 'Mỹ thuật', subSubjects: [] },
  ],
  '12': [
    { subject: 'Toán', subSubjects: ['Số học', 'Đại số', 'Hình học'] },
    { subject: 'Ngữ văn', subSubjects: [] },
    { subject: 'Tiếng Anh', subSubjects: [] },
    { subject: 'Giáo dục thể chất', subSubjects: [] },
    { subject: 'Giáo dục quốc phòng và an ninh', subSubjects: [] },
    { subject: 'Hoạt động trải nghiệm, hướng nghiệp', subSubjects: [] },
    { subject: 'Nội dung giáo dục địa phương', subSubjects: [] },
    { subject: 'Lịch sử', subSubjects: [] },
    { subject: 'Địa lí', subSubjects: [] },
    { subject: 'Giáo dục kinh tế và pháp luật', subSubjects: [] },
    { subject: 'Vật lí', subSubjects: [] },
    { subject: 'Hóa học', subSubjects: [] },
    { subject: 'Sinh học', subSubjects: [] },
    { subject: 'Công nghệ', subSubjects: [] },
    { subject: 'Tin học', subSubjects: [] },
    { subject: 'Âm nhạc', subSubjects: [] },
    { subject: 'Mỹ thuật', subSubjects: [] },
  ],
};



function BusinessConfigSection({
  info,
  setInfo,
  setActiveTab,
  currentUser,
  onSync,
  onNotify,
}: {
  info: BusinessInfo;
  setInfo: (i: BusinessInfo) => void;
  setActiveTab: (t: string) => void;
  currentUser: UserAccount | null;
  onSync?: (info: BusinessInfo) => void;
  onNotify?: (msg: string, type?: 'success' | 'error' | 'info') => void;
}) {
  const isAdmin = currentUser?.role === 'admin';

  const displayInfo = React.useMemo(() => {
    if (isAdmin) return info;
    return {
      name: currentUser?.businessName || info.name,
      owner: currentUser?.businessOwner || info.owner,
      taxId: currentUser?.taxCode || info.taxId,
      address: currentUser?.businessAddress || info.address,
      businessLocation: currentUser?.businessLocation || info.businessLocation,
    };
  }, [info, currentUser, isAdmin]);

  const [formData, setFormData] = React.useState<BusinessInfo>(displayInfo);

  React.useEffect(() => {
    setFormData(displayInfo);
  }, [displayInfo]);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin) return;
    setInfo(formData);
    if (onSync) onSync(formData);
    onNotify?.('Đã lưu và đồng bộ cấu hình hộ kinh doanh thành công!', 'success');
  };

  return (
    <div className="space-y-6">
      <SectionBanner
        title="Cấu hình Hộ kinh doanh"
        subtitle={
          isAdmin
            ? 'Thiết lập thông tin pháp lý chuẩn của hộ kinh doanh để tự động điền lên sổ sách, phiếu thu chi và báo cáo thuế.'
            : 'Thông tin pháp lý cơ sở dạy thêm (Chế độ xem dành cho giáo viên).'
        }
        badge="HỒ SƠ ĐƠN VỊ"
        icon={Building2}
        onBack={() => setActiveTab('dashboard')}
        stats={[
          { label: 'Tên đơn vị', value: displayInfo.name || 'Chưa cập nhật', sub: 'Hiển thị trên chứng từ' },
          { label: 'Đại diện pháp luật', value: displayInfo.owner || 'Chưa cập nhật', sub: 'Chủ hộ kinh doanh' },
          { label: 'Mã số thuế', value: displayInfo.taxId || 'Chưa cập nhật', sub: 'Kê khai thuế HKD', highlight: true },
          { label: 'Quyền truy cập', value: isAdmin ? 'Quản trị viên' : 'Chỉ xem', sub: 'Phân quyền hệ thống' },
        ]}
      />

      <div className="bg-white dark:bg-slate-900 p-6 sm:p-8 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm">
        <form onSubmit={handleSubmit} className="space-y-6">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
            <div className="space-y-2">
              <label className="text-sm font-bold text-slate-700 dark:text-slate-300">Tên hộ kinh doanh / Cơ sở dạy thêm</label>
              <input
                type="text"
                value={formData.name}
                onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                disabled={!isAdmin}
                className="input-field disabled:opacity-70 disabled:cursor-not-allowed"
                placeholder="Ví dụ: CƠ SỞ GIÁO DỤC HOÀNG GIA"
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-bold text-slate-700 dark:text-slate-300">Chủ hộ kinh doanh</label>
              <input
                type="text"
                value={formData.owner}
                onChange={(e) => setFormData({ ...formData, owner: e.target.value })}
                disabled={!isAdmin}
                className="input-field disabled:opacity-70 disabled:cursor-not-allowed"
                placeholder="Họ và tên chủ hộ"
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-bold text-slate-700 dark:text-slate-300">Mã số thuế</label>
              <input
                type="text"
                value={formData.taxId}
                onChange={(e) => setFormData({ ...formData, taxId: e.target.value })}
                disabled={!isAdmin}
                className="input-field font-mono tabular-nums disabled:opacity-70 disabled:cursor-not-allowed"
                placeholder="Mã số thuế hộ kinh doanh"
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-bold text-slate-700 dark:text-slate-300">Địa chỉ đăng ký</label>
              <input
                type="text"
                value={formData.address}
                onChange={(e) => setFormData({ ...formData, address: e.target.value })}
                disabled={!isAdmin}
                className="input-field disabled:opacity-70 disabled:cursor-not-allowed"
                placeholder="Địa chỉ trên giấy phép HKD"
              />
            </div>
            <div className="space-y-2 md:col-span-2">
              <label className="text-sm font-bold text-slate-700 dark:text-slate-300">Địa điểm tổ chức dạy thêm thực tế</label>
              <input
                type="text"
                value={formData.businessLocation}
                onChange={(e) => setFormData({ ...formData, businessLocation: e.target.value })}
                disabled={!isAdmin}
                className="input-field disabled:opacity-70 disabled:cursor-not-allowed"
                placeholder="Địa điểm giảng dạy thực tế"
              />
            </div>
          </div>

          <div className="pt-4 border-t border-slate-100 dark:border-slate-800 flex flex-wrap gap-3">
            {isAdmin && (
              <button
                type="submit"
                className="btn-primary flex items-center gap-2"
              >
                <Save className="w-4 h-4" />
                Lưu cấu hình hộ kinh doanh
              </button>
            )}
            <button
              type="button"
              onClick={() => setActiveTab('dashboard')}
              className="btn-secondary"
            >
              Về trang chủ
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

function MenuSettingsSection({
  menuOrder,
  setMenuOrder,
  setActiveTab,
}: {
  menuOrder: string[];
  setMenuOrder: (o: string[]) => void;
  setActiveTab?: (t: string) => void;
}) {
  const ALL_MENU_ITEMS = [
    { id: 'dashboard', label: 'TRANG CHỦ', icon: Home },
    { id: 'business', label: 'HỘ KINH DOANH', icon: Building2 },
    { id: 'students_group', label: 'QUẢN LÝ HỌC SINH', icon: Users },
    { id: 'student_learning_group', label: 'QUẢN LÝ HỌC TẬP HỌC SINH', icon: GraduationCap },
    { id: 'program', label: 'QUẢN LÝ CHƯƠNG TRÌNH DẠY', icon: BookOpen },
    { id: 'finance_group', label: 'QUẢN LÝ TÀI CHÍNH', icon: DollarSign },
    { id: 'tax_report', label: 'BÁO CÁO THUẾ CHO HKD', icon: Calculator },
  ];

  const moveUp = (index: number) => {
    if (index === 0) return;
    const newOrder = [...menuOrder];
    [newOrder[index - 1], newOrder[index]] = [newOrder[index], newOrder[index - 1]];
    setMenuOrder(newOrder);
  };

  const moveDown = (index: number) => {
    if (index === menuOrder.length - 1) return;
    const newOrder = [...menuOrder];
    [newOrder[index + 1], newOrder[index]] = [newOrder[index], newOrder[index + 1]];
    setMenuOrder(newOrder);
  };

  return (
    <div className="space-y-6">
      <SectionBanner
        title="Sắp xếp Thứ tự Menu"
        subtitle="Tùy chỉnh vị trí hiển thị của các phân hệ chức năng trên thanh điều hướng bên trái."
        badge="CẤU HÌNH HỆ THỐNG"
        icon={Settings}
        onBack={setActiveTab ? () => setActiveTab('dashboard') : undefined}
      />

      <div className="bg-white dark:bg-slate-900 p-6 sm:p-8 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm max-w-3xl">
        <div className="space-y-2.5">
          {menuOrder.map((id, index) => {
            const item = ALL_MENU_ITEMS.find((i) => i.id === id);
            if (!item) return null;
            return (
              <div
                key={id}
                className="flex items-center justify-between p-4 bg-slate-50 dark:bg-slate-800/60 rounded-xl border border-slate-200/80 dark:border-slate-700"
              >
                <div className="flex items-center gap-3.5">
                  <div className="w-10 h-10 bg-teal-50 dark:bg-teal-900/30 border border-teal-200/70 dark:border-teal-800 rounded-xl flex items-center justify-center text-teal-700 dark:text-teal-300">
                    <item.icon className="w-5 h-5" />
                  </div>
                  <div>
                    <span className="text-xs font-mono font-bold text-slate-400 mr-2">#{index + 1}</span>
                    <span className="font-bold text-sm text-slate-900 dark:text-slate-100">{item.label}</span>
                  </div>
                </div>
                <div className="flex items-center gap-1.5">
                  <button
                    onClick={() => moveUp(index)}
                    disabled={index === 0}
                    className="p-2 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 hover:border-teal-500 hover:text-teal-700 rounded-lg disabled:opacity-30 transition-colors"
                    title="Di chuyển lên"
                  >
                    <ChevronUp className="w-4 h-4" />
                  </button>
                  <button
                    onClick={() => moveDown(index)}
                    disabled={index === menuOrder.length - 1}
                    className="p-2 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 hover:border-teal-500 hover:text-teal-700 rounded-lg disabled:opacity-30 transition-colors"
                    title="Di chuyển xuống"
                  >
                    <ChevronDown className="w-4 h-4" />
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function ClassConfigSection({
  classes,
  setClasses,
  currentUser,
  onNotify,
}: {
  key?: string;
  classes: ClassSubject[];
  setClasses: (c: ClassSubject[]) => void;
  setActiveTab: (t: string) => void;
  currentUser: UserAccount | null;
  onNotify?: (msg: string, type?: 'success' | 'error' | 'info') => void;
}) {
  const isAdmin = currentUser?.role === 'admin';

  const addRow = () => {
    if (!isAdmin) return;
    setClasses([...classes, { id: crypto.randomUUID(), grade: '', subject: '', subSubject: '' }]);
  };

  const resetToDefault = () => {
    if (!isAdmin) return;
    setClasses(DEFAULT_CLASSES);
    onNotify?.('Đã khôi phục cấu hình khối lớp và môn học mặc định.', 'info');
  };

  const removeRow = (id: string) => {
    if (!isAdmin) return;
    setClasses(classes.filter((c) => c.id !== id));
  };

  const handleChange = (id: string, field: keyof ClassSubject, value: string) => {
    if (!isAdmin) return;
    setClasses(classes.map((c) => (c.id === id ? { ...c, [field]: value } : c)));
  };

  return (
    <div className="space-y-5">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-white dark:bg-slate-900 p-5 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm">
        <div>
          <h3 className="text-lg font-bold text-slate-900 dark:text-white">Danh mục Khối lớp & Môn học</h3>
          <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
            Thiết lập các khối lớp, môn học và phân môn áp dụng giảng dạy tại cơ sở.
          </p>
        </div>
        {isAdmin && (
          <div className="flex flex-wrap items-center gap-2.5">
            <ConfirmButton
              onConfirm={resetToDefault}
              className="btn-secondary flex items-center gap-2 text-xs"
              icon={RotateCcw}
              confirmText="Xác nhận khôi phục?"
            >
              Khôi phục mặc định
            </ConfirmButton>
            <button onClick={addRow} className="btn-primary flex items-center gap-2 text-xs">
              <Plus className="w-4 h-4" />
              Thêm dòng mới
            </button>
          </div>
        )}
      </div>

      <div className="saas-card overflow-hidden !p-0">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-50 dark:bg-slate-800/70 border-b border-slate-200 dark:border-slate-800">
                <th className="px-6 py-4 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Khối lớp</th>
                <th className="px-6 py-4 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Môn học</th>
                <th className="px-6 py-4 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Phân môn</th>
                <th className="px-6 py-4 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider w-20"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
              {classes.map((cls) => (
                <tr key={cls.id} className="group hover:bg-teal-50/30 dark:hover:bg-slate-800/40 transition-colors">
                  <td className="px-6 py-3.5">
                    <input
                      value={cls.grade}
                      onChange={(e) => handleChange(cls.id, 'grade', e.target.value)}
                      disabled={!isAdmin}
                      placeholder="VD: 6"
                      className="w-full bg-transparent border-none focus:ring-0 text-sm font-bold text-slate-900 dark:text-white placeholder-slate-300 dark:placeholder-slate-700 disabled:opacity-80"
                    />
                  </td>
                  <td className="px-6 py-3.5">
                    <input
                      value={cls.subject}
                      onChange={(e) => handleChange(cls.id, 'subject', e.target.value)}
                      disabled={!isAdmin}
                      placeholder="VD: Toán"
                      className="w-full bg-transparent border-none focus:ring-0 text-sm font-semibold text-teal-700 dark:text-teal-300 placeholder-slate-300 dark:placeholder-slate-700 disabled:opacity-80"
                    />
                  </td>
                  <td className="px-6 py-3.5">
                    <input
                      value={cls.subSubject}
                      onChange={(e) => handleChange(cls.id, 'subSubject', e.target.value)}
                      disabled={!isAdmin}
                      placeholder="VD: Đại số"
                      className="w-full bg-transparent border-none focus:ring-0 text-sm font-medium text-slate-600 dark:text-slate-400 placeholder-slate-300 dark:placeholder-slate-700 disabled:opacity-80"
                    />
                  </td>
                  <td className="px-6 py-3.5 text-right">
                    {isAdmin && (
                      <button
                        onClick={() => removeRow(cls.id)}
                        className="w-8 h-8 rounded-lg inline-flex items-center justify-center text-slate-400 hover:text-red-600 hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {classes.length === 0 && (
          <div className="py-16 text-center">
            <Layers className="w-10 h-10 text-slate-300 dark:text-slate-700 mx-auto mb-3" />
            <h4 className="text-base font-bold text-slate-900 dark:text-white">Chưa có lớp học nào</h4>
            <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">Nhấn "Thêm dòng mới" để bắt đầu cấu hình hệ thống.</p>
          </div>
        )}
      </div>
    </div>
  );
}

function PPCTSection({ ppctData, setPpctData, classes, setPlans, plans, setActiveTab, currentUser, onNotify }: { 
  key?: string,
  ppctData: PPCTItem[], 
  setPpctData: (d: PPCTItem[]) => void, 
  classes: ClassSubject[],
  setPlans: (p: LessonPlan[]) => void,
  plans: LessonPlan[],
  setActiveTab: (t: string) => void,
  currentUser: UserAccount | null,
  onNotify?: (msg: string, type?: 'success' | 'error' | 'info') => void
}) {
  const isAdmin = currentUser?.role === 'admin';
  const [activeGrade, setActiveGrade] = useState<string>('6');

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!isAdmin) return;
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target?.result;
        const wb = XLSX.read(data, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const jsonData = XLSX.utils.sheet_to_json(ws) as any[];
        
        const mappedData: PPCTItem[] = jsonData.map(row => {
          const getRowValue = (keys: string[]) => {
            for (const key of keys) {
              if (row[key] !== undefined) return row[key];
              const foundKey = Object.keys(row).find(k => k.trim().toLowerCase() === key.toLowerCase());
              if (foundKey) return row[foundKey];
            }
            return undefined;
          };

          let subject = String(getRowValue(['Môn', 'Môn học', 'Môn Học', 'Subject']) || '').trim();
          let subSubject = String(getRowValue(['Phân môn', 'Phân Môn', 'Sub-subject', 'SubSubject', 'Phân môn học']) || '').trim();
          
          // Normalization
          if (subject.toLowerCase() === 'khoa học tự nhiên' || subject.toLowerCase() === 'khtn') subject = 'KHTN';
          if (subject.toLowerCase() === 'toán') subject = 'Toán';
          if (subject.toLowerCase() === 'ngữ văn') subject = 'Ngữ văn';
          
          if (subject === 'KHTN') {
            if (subSubject.toLowerCase() === 'vật lý' || subSubject.toLowerCase() === 'vật lí') subSubject = 'Vật lí';
            if (subSubject.toLowerCase() === 'sinh học') subSubject = 'Sinh học';
            if (subSubject.toLowerCase() === 'hóa học') subSubject = 'Hóa học';
          }
          
          if (subject === 'Toán') {
            if (subSubject === 'Số học và Đại số' || subSubject === 'Số và Đại số') subSubject = 'Đại số';
            if (subSubject === 'Hình học và Đo lường' || subSubject === 'Hình học và đo lường') subSubject = 'Hình học';
            if (subSubject.toLowerCase() === 'số học') subSubject = 'Số học';
            if (subSubject.toLowerCase() === 'đại số') subSubject = 'Đại số';
            if (subSubject.toLowerCase() === 'hình học') subSubject = 'Hình học';
          }

          return {
            id: crypto.randomUUID(),
            grade: normalizeGrade(getRowValue(['Lớp', 'Khối', 'Grade', 'Level']) || activeGrade),
            subject,
            subSubject,
            period: Number(getRowValue(['Tiết theo PPCT', 'Tiết', 'Period']) || 0),
            content: String(getRowValue(['Nội dung', 'Tên bài dạy', 'Nội dung bài học', 'Content', 'Lesson Name']) || '').trim(),
            notes: String(getRowValue(['Ghi chú', 'Notes']) || '').trim()
          };
        }).filter(item => item.subject && item.content);

        if (mappedData.length === 0) {
          onNotify?.('Không tìm thấy dữ liệu hợp lệ trong file Excel. Vui lòng kiểm tra lại định dạng file mẫu.', 'error');
          return;
        }

        const targetGrade = normalizeGrade(activeGrade);
        const otherGradesData = ppctData.filter(item => normalizeGrade(item.grade) !== targetGrade);
        setPpctData([...otherGradesData, ...mappedData]);
        onNotify?.(`Đã nhập ${mappedData.length} dòng phân phối chương trình cho Khối ${targetGrade}.`, 'success');
      } catch (err) {
        console.error("Upload error:", err);
      } finally {
        e.target.value = '';
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const deleteRow = (id: string) => {
    if (!isAdmin) return;
    setPpctData(ppctData.filter(d => d.id !== id));
  };

  const downloadSamplePPCT = () => {
    const sampleData = [
      { 'Lớp': activeGrade, 'Môn': 'Toán', 'Phân môn': 'Đại số', 'Tiết theo PPCT': 1, 'Nội dung': 'Tập hợp các số tự nhiên', 'Ghi chú': '' },
    ];
    const ws = XLSX.utils.json_to_sheet(sampleData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "PPCT Mẫu");
    XLSX.writeFile(wb, `Mau_PPCT_Lop_${activeGrade}.xlsx`);
  };

  const clearData = () => {
    if (!isAdmin) return;
    const targetGrade = normalizeGrade(activeGrade);
    setPpctData(ppctData.filter(item => normalizeGrade(item.grade) !== targetGrade));
  };

  const handleChange = (id: string, field: keyof PPCTItem, value: any) => {
    if (!isAdmin) return;
    setPpctData(ppctData.map(item => item.id === id ? { ...item, [field]: value } : item));
  };

  const addRow = () => {
    if (!isAdmin) return;
    setPpctData([...ppctData, {
      id: crypto.randomUUID(),
      grade: activeGrade,
      subject: '',
      subSubject: '',
      period: (filteredData.length > 0 ? Math.max(...filteredData.map(d => d.period)) + 1 : 1),
      content: '',
      notes: ''
    }]);
  };

  const resetToDefault = () => {
    if (!isAdmin) return;
    setPpctData(DEFAULT_PPCT);
    onNotify?.('Đã khôi phục phân phối chương trình mặc định.', 'info');
  };

  const syncToLessonPlan = () => {
    const today = new Date();
    const start = startOfWeek(today, { weekStartsOn: 1 });
    const end = endOfWeek(today, { weekStartsOn: 1 });
    
    const rows: LessonPlanRow[] = [];
    const WEEKDAYS = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6'];
    const WEEKEND = ['Thứ 7', 'Chủ Nhật'];

    WEEKDAYS.forEach((day, idx) => {
      const date = addDays(start, idx);
      ['Ca 1 (17h-19h)', 'Ca 2 (19h-21h)'].forEach(shift => {
        rows.push({
          id: crypto.randomUUID(),
          day,
          date: format(date, 'dd/MM/yyyy'),
          shift,
          grade: '',
          subject: '',
          subSubject: '',
          period: '',
          content: '',
          notes: ''
        });
      });
    });

    WEEKEND.forEach((day, idx) => {
      const date = addDays(start, 5 + idx);
      const shifts = day === 'Thứ 7' ? [
        'Ca 1 (7h-9h)', 
        'Ca 2 (9h-11h)', 
        'Ca 3 (17h-19h)', 
        'Ca 4 (19h-21h)'
      ] : [
        'Ca 1 (7h-9h)', 
        'Ca 2 (9h-11h)', 
        'Ca 3 (14h-16h)', 
        'Ca 4 (16h-18h)', 
        'Ca 5 (18h-20h)', 
        'Ca 6 (20h-22h)'
      ];

      shifts.forEach(shift => {
        rows.push({
          id: crypto.randomUUID(),
          day,
          date: format(date, 'dd/MM/yyyy'),
          shift,
          grade: '',
          subject: '',
          subSubject: '',
          period: '',
          content: '',
          notes: ''
        });
      });
    });

    const newPlan: LessonPlan = {
      id: crypto.randomUUID(),
      teacherName: '',
      week: format(start, 'w'),
      startDate: format(start, 'yyyy-MM-dd'),
      endDate: format(end, 'yyyy-MM-dd'),
      rows
    };

    setPlans([newPlan, ...plans]);
    setActiveTab('lesson-plan');
  };

  const exportAllPPCT = () => {
    if (ppctData.length === 0) {
      onNotify?.('Không có dữ liệu phân phối chương trình để xuất.', 'error');
      return;
    }
    const exportData = [...ppctData].sort((a, b) => {
      const gradeA = normalizeGrade(a.grade);
      const gradeB = normalizeGrade(b.grade);
      if (gradeA !== gradeB) return gradeA.localeCompare(gradeB, undefined, { numeric: true });
      if (a.subject !== b.subject) return a.subject.localeCompare(b.subject);
      return a.period - b.period;
    }).map(item => ({
      'Lớp': item.grade,
      'Môn': item.subject,
      'Phân môn': item.subSubject,
      'Tiết theo PPCT': item.period,
      'Nội dung': item.content,
      'Ghi chú': item.notes
    }));
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Phân phối chương trình");
    XLSX.writeFile(wb, "Phan_Phoi_Chuong_Trinh_Tong_Hop.xlsx");
  };

  const grades = Array.from(new Set(classes.map(c => c.grade))).sort();
  const filteredData = ppctData.filter(item => normalizeGrade(item.grade) === normalizeGrade(activeGrade));

  return (
    <div className="space-y-6">
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 bg-white dark:bg-slate-900 p-5 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm">
        <div>
          <h3 className="text-lg font-bold text-slate-900 dark:text-white">Phân phối chương trình</h3>
          <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">Quản lý kế hoạch giảng dạy chi tiết theo từng khối lớp.</p>
        </div>
        <div className="flex flex-wrap items-center gap-2.5">
          {isAdmin && (
            <>
              <ConfirmButton
                onConfirm={resetToDefault}
                className="btn-secondary flex items-center gap-2 text-xs"
                icon={RotateCcw}
                confirmText="Xác nhận khôi phục?"
              >
                Khôi phục mặc định
              </ConfirmButton>
              <button 
                onClick={syncToLessonPlan}
                className="btn-secondary flex items-center gap-2 text-xs"
              >
                <Calendar className="w-4 h-4" />
                Tạo lịch báo giảng
              </button>
            </>
          )}
          <button 
            onClick={downloadSamplePPCT}
            className="btn-secondary flex items-center gap-2 text-xs"
          >
            <Download className="w-4 h-4" />
            Tải file mẫu
          </button>
          <button 
            onClick={exportAllPPCT}
            className="px-4 py-2.5 bg-teal-50 dark:bg-teal-900/30 text-teal-700 dark:text-teal-300 border border-teal-200 dark:border-teal-800 rounded-xl hover:bg-teal-100 dark:hover:bg-teal-900/50 transition-colors font-bold text-xs flex items-center gap-2"
          >
            <FileSpreadsheet className="w-4 h-4" />
            Xuất PPCT các lớp
          </button>
          {isAdmin && (
            <>
              <label className="px-6 py-3 bg-primary text-white rounded-xl hover:bg-primary-hover transition-all shadow-lg shadow-primary/20 font-bold text-sm flex items-center gap-3 cursor-pointer">
                <Upload className="w-4 h-4" />
                Nhập từ Excel
                <input type="file" accept=".xlsx, .xls" className="hidden" onChange={handleFileUpload} />
              </label>
              <button 
                onClick={addRow}
                className="px-6 py-3 bg-primary text-white rounded-xl hover:bg-primary-hover transition-all shadow-lg shadow-primary/20 font-bold text-sm flex items-center gap-3"
              >
                <Plus className="w-4 h-4" />
                Thêm dòng mới
              </button>
            </>
          )}
        </div>
      </div>

      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div className="flex items-center gap-1.5 p-1.5 bg-neutral-100 dark:bg-slate-800 rounded-2xl w-fit">
          {(grades.length > 0 ? grades : ['6', '7', '8', '9']).map((grade) => (
              <button
                key={grade}
                onClick={() => setActiveGrade(grade)}
                className={cn(
                  "px-6 py-3 rounded-xl text-base font-bold transition-all",
                  normalizeGrade(activeGrade) === normalizeGrade(grade) 
                    ? "bg-white dark:bg-slate-700 text-primary dark:text-white shadow-sm" 
                    : "text-neutral-500 dark:text-slate-400 hover:text-neutral-900 dark:hover:text-slate-200"
                )}
              >
              Khối {grade}
            </button>
          ))}
        </div>
        {filteredData.length > 0 && isAdmin && (
          <ConfirmButton
            onConfirm={clearData}
            className="text-sm text-red-600 dark:text-red-400 font-bold flex items-center gap-2 px-4 py-2 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-xl transition-all"
            icon={Trash2}
            confirmText="Xác nhận xóa?"
          >
            Xóa dữ liệu khối {activeGrade}
          </ConfirmButton>
        )}
      </div>

      <div className="saas-card overflow-hidden !p-0">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse min-w-[1000px]">
            <thead>
              <tr className="bg-neutral-100 dark:bg-slate-900 border-b-2 border-neutral-200 dark:border-slate-800">
                <th className="px-6 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-16 text-center">Tiết</th>
                <th className="px-6 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-40">Môn học</th>
                <th className="px-6 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-40">Phân môn</th>
                <th className="px-6 py-5 text-xs font-semibold text-neutral-900 dark:text-white">Nội dung bài dạy</th>
                <th className="px-6 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-20"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-neutral-200 dark:divide-slate-800">
              {filteredData.map((item, idx) => (
                <tr key={item.id} className="hover:bg-neutral-50/50 dark:hover:bg-slate-900/30 transition-colors group">
                    <td className="px-6 py-4">
                      <input
                        type="number"
                        value={item.period || ''}
                        onChange={(e) => handleChange(item.id, 'period', Number(e.target.value))}
                        disabled={!isAdmin}
                        className="w-full bg-transparent border-none focus:ring-0 text-base font-semibold text-neutral-900 dark:text-white text-center font-mono disabled:opacity-80 [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                        placeholder="..."
                      />
                    </td>
                  <td className="px-6 py-4">
                    <select
                      value={item.subject}
                      onChange={(e) => handleChange(item.id, 'subject', e.target.value)}
                      disabled={!isAdmin}
                      className="w-full bg-transparent border-none focus:ring-0 text-base font-semibold text-primary disabled:opacity-80 cursor-pointer"
                    >
                      <option value="">-- Môn học --</option>
                      {CURRICULUM_2018_DATA[normalizeGrade(activeGrade)]?.map(s => (
                        <option key={s.subject} value={s.subject}>{s.subject}</option>
                      )) || (
                        <option value={item.subject}>{item.subject}</option>
                      )}
                    </select>
                  </td>
                  <td className="px-6 py-4">
                    <div className="flex flex-col gap-1">
                      <select
                        value={
                          !item.subSubject ? "" :
                          (CURRICULUM_2018_DATA[normalizeGrade(activeGrade)]?.find(s => s.subject === item.subject)?.subSubjects.includes(item.subSubject) 
                            ? item.subSubject 
                            : "custom")
                        }
                        onChange={(e) => {
                          if (e.target.value === "custom") {
                            handleChange(item.id, "subSubject", "Mới...");
                          } else {
                            handleChange(item.id, "subSubject", e.target.value);
                          }
                        }}
                        disabled={!isAdmin}
                        className="w-full bg-transparent border-none focus:ring-0 text-base font-bold text-neutral-600 dark:text-slate-400 italic disabled:opacity-80 cursor-pointer"
                      >
                        <option value="">-- Phân môn --</option>
                        {CURRICULUM_2018_DATA[normalizeGrade(activeGrade)]?.find(s => s.subject === item.subject)?.subSubjects.map(ss => (
                          <option key={ss} value={ss}>{ss}</option>
                        ))}
                        <option value="custom">-- Nhập khác --</option>
                      </select>
                      {item.subSubject && (item.subSubject === "custom" || item.subSubject === "Mới..." || !CURRICULUM_2018_DATA[normalizeGrade(activeGrade)]?.find(s => s.subject === item.subject)?.subSubjects.includes(item.subSubject)) && (
                        <input
                          type="text"
                          value={item.subSubject === "custom" || item.subSubject === "Mới..." ? "" : item.subSubject}
                          onChange={(e) => handleChange(item.id, "subSubject", e.target.value)}
                          disabled={!isAdmin}
                          placeholder="Tên phân môn..."
                          className="w-full bg-transparent border-b border-neutral-100 dark:border-slate-800 focus:ring-0 text-[10px] text-neutral-400 dark:text-slate-500 px-1 italic"
                        />
                      )}
                    </div>
                  </td>
                  <td className="px-6 py-4">
                    <input
                      value={item.content}
                      onChange={(e) => handleChange(item.id, 'content', e.target.value)}
                      disabled={!isAdmin}
                      className="w-full bg-transparent border-none focus:ring-0 text-base text-neutral-900 dark:text-white font-semibold disabled:opacity-80"
                      placeholder="Nội dung bài dạy"
                    />
                  </td>
                  <td className="px-6 py-4 text-right">
                    {isAdmin && (
                      <button
                        onClick={() => deleteRow(item.id)}
                        className="w-9 h-9 rounded-xl flex items-center justify-center text-neutral-400 hover:text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 transition-all opacity-0 group-hover:opacity-100"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    )}
                  </td>
                </tr>
              ))}
              {filteredData.length === 0 && (
                <tr>
                  <td colSpan={5} className="px-6 py-24 text-center">
                    <div className="w-20 h-20 bg-neutral-50 dark:bg-slate-900 rounded-3xl flex items-center justify-center mx-auto mb-6 text-neutral-300 dark:text-slate-700">
                      <FileSpreadsheet className="w-10 h-10" />
                    </div>
                    <h4 className="text-xl font-bold text-neutral-900 dark:text-white">Chưa có dữ liệu PPCT</h4>
                    <p className="text-sm text-neutral-500 dark:text-slate-400 mt-2">Vui lòng nhập dữ liệu từ file Excel để bắt đầu.</p>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function LessonPlanSection({ 
  plans, 
  setPlans, 
  deletePlan, 
  ppctData, 
  classes, 
  businessInfo,
  setActiveTab
}: { 
  key?: string,
  plans: LessonPlan[], 
  setPlans: (p: LessonPlan[]) => void, 
  deletePlan: (id: string) => void,
  ppctData: PPCTItem[], 
  classes: ClassSubject[],
  businessInfo: BusinessInfo,
  setActiveTab: (t: string) => void
}) {
  const [isCreating, setIsCreating] = useState(false);
  const [editingPlan, setEditingPlan] = useState<LessonPlan | null>(null);

  const startNewPlan = () => {
    const today = new Date();
    const start = startOfWeek(today, { weekStartsOn: 1 });
    const end = endOfWeek(today, { weekStartsOn: 1 });
    
    const rows: LessonPlanRow[] = [];
    const WEEKDAYS = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6'];
    const WEEKEND = ['Thứ 7', 'Chủ Nhật'];

    WEEKDAYS.forEach((day, idx) => {
      const date = addDays(start, idx);
      ['Ca 1 (17h-19h)', 'Ca 2 (19h-21h)'].forEach(shift => {
        rows.push({
          id: crypto.randomUUID(),
          day,
          date: format(date, 'dd/MM/yyyy'),
          shift,
          grade: '',
          subject: '',
          subSubject: '',
          period: '',
          content: '',
          notes: ''
        });
      });
    });

    WEEKEND.forEach((day, idx) => {
      const date = addDays(start, 5 + idx);
      const shifts = day === 'Thứ 7' ? [
        'Ca 1 (7h-9h)', 
        'Ca 2 (9h-11h)', 
        'Ca 3 (17h-19h)', 
        'Ca 4 (19h-21h)'
      ] : [
        'Ca 1 (7h-9h)', 
        'Ca 2 (9h-11h)', 
        'Ca 3 (14h-16h)', 
        'Ca 4 (16h-18h)', 
        'Ca 5 (18h-20h)', 
        'Ca 6 (20h-22h)'
      ];

      shifts.forEach(shift => {
        rows.push({
          id: crypto.randomUUID(),
          day,
          date: format(date, 'dd/MM/yyyy'),
          shift,
          grade: '',
          subject: '',
          subSubject: '',
          period: '',
          content: '',
          notes: ''
        });
      });
    });

    const newPlan: LessonPlan = {
      id: crypto.randomUUID(),
      teacherName: '',
      week: format(start, 'w'),
      startDate: format(start, 'yyyy-MM-dd'),
      endDate: format(end, 'yyyy-MM-dd'),
      rows
    };
    setEditingPlan(newPlan);
    setIsCreating(true);
  };

  const savePlan = () => {
    if (!editingPlan) return;
    if (plans.find(p => p.id === editingPlan.id)) {
      setPlans(plans.map(p => p.id === editingPlan.id ? editingPlan : p));
    } else {
      setPlans([editingPlan, ...plans]);
    }
    setIsCreating(false);
    setEditingPlan(null);
  };

  const handleRowChange = (rowId: string, field: keyof LessonPlanRow, value: string) => {
    if (!editingPlan) return;
    const newRows = editingPlan.rows.map(row => {
      if (row.id === rowId) {
        const updatedRow = { ...row, [field]: value };
        
        if (field === 'grade') {
          updatedRow.subject = '';
          updatedRow.subSubject = '';
          updatedRow.period = '';
          updatedRow.content = '';
        } else if (field === 'subject') {
          updatedRow.subSubject = '';
          updatedRow.period = '';
          updatedRow.content = '';
        } else if (field === 'subSubject') {
          updatedRow.period = '';
          updatedRow.content = '';
        } else if (field === 'period') {
          const ppctMatch = ppctData.find(p => 
            normalizeGrade(p.grade) === normalizeGrade(updatedRow.grade) && 
            String(p.subject).trim().toLowerCase() === String(updatedRow.subject).trim().toLowerCase() && 
            (updatedRow.subSubject ? String(p.subSubject).trim().toLowerCase() === String(updatedRow.subSubject).trim().toLowerCase() : true) &&
            Number(p.period) === Number(value)
          );
          if (ppctMatch) {
            updatedRow.content = ppctMatch.content;
          }
        }
        
        return updatedRow;
      }
      return row;
    });
    setEditingPlan({ ...editingPlan, rows: newRows });
  };

  const autoFillContent = (rowId: string) => {
    if (!editingPlan) return;
    const row = editingPlan.rows.find(r => r.id === rowId);
    if (!row || !row.grade || !row.subject || !row.period) return;

    const ppctMatch = ppctData.find(p => 
      normalizeGrade(p.grade) === normalizeGrade(row.grade) && 
      String(p.subject).trim().toLowerCase() === String(row.subject).trim().toLowerCase() && 
      (row.subSubject ? String(p.subSubject).trim().toLowerCase() === String(row.subSubject).trim().toLowerCase() : true) &&
      Number(p.period) === Number(row.period)
    );

    if (ppctMatch) {
      handleRowChange(rowId, 'content', ppctMatch.content);
    }
  };

  const exportToWord = async (plan: LessonPlan) => {
    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 850, bottom: 850, left: 1134, right: 850 }
          }
        },
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: sanitizeDocxText(`Hộ kinh doanh: ${businessInfo.name}`), bold: true, size: 22 }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: sanitizeDocxText(`Địa chỉ: ${businessInfo.address}`), bold: true, size: 22 }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 400, after: 200 },
            children: [
              new TextRun({ text: "KẾ HOẠCH DẠY HỌC CỦA GIÁO VIÊN", bold: true, size: 28 }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: sanitizeDocxText(`Họ tên giáo viên dạy: ${plan.teacherName}`), bold: true, size: 24 }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: sanitizeDocxText(`Tuần: ${plan.week} - Từ ngày: ${safeFormat(plan.startDate, 'dd/MM/yyyy')} - Đến ngày: ${safeFormat(plan.endDate, 'dd/MM/yyyy')}`), bold: true, size: 22 }),
            ],
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  "Thứ, ngày tháng", "Ca học", "Lớp", "Môn học", "Phân môn", "Tiết PPCT", "Tên bài dạy", "Ghi chú"
                ].map(text => new TableCell({
                  children: [new Paragraph({ 
                    alignment: AlignmentType.CENTER, 
                    children: [new TextRun({ text: sanitizeDocxText(text), bold: true, size: 22 })] 
                  })],
                  verticalAlign: VerticalAlign.CENTER,
                  shading: { fill: "F2F2F2" }
                })),
              }),
              ...(() => {
                const rowsToExport = plan.rows.filter(r => r.grade);
                const tableRows: TableRow[] = [];
                let lastDay = "";
                
                rowsToExport.forEach((row) => {
                  const dayStr = row.day + (row.date ? ` (${row.date})` : '');
                  const dayText = sanitizeDocxText(dayStr);
                  const isNewDay = dayText !== lastDay;
                  if (isNewDay) lastDay = dayText;

                  const rowChildren: TableCell[] = [];
                  
                  if (isNewDay) {
                    rowChildren.push(new TableCell({
                      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText(dayText), size: 20 })] })],
                      verticalAlign: VerticalAlign.CENTER,
                      rowSpan: rowsToExport.filter(r => (r.day + (r.date ? ` (${r.date})` : '')) === dayStr).length,
                    }));
                  }

                  const shiftMatch = row.shift.match(/(Ca \d+)\s*(.*)/);
                  const shiftLine1 = shiftMatch ? shiftMatch[1] : row.shift;
                  const shiftLine2 = shiftMatch ? shiftMatch[2] : "";

                  rowChildren.push(new TableCell({
                    children: [
                      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText(shiftLine1), size: 20 })] }),
                      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText(shiftLine2), size: 18 })] }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                  }));

                  [row.grade, row.subject, row.subSubject, row.period].forEach(text => {
                    rowChildren.push(new TableCell({
                      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText(text), size: 20 })] })],
                      verticalAlign: VerticalAlign.CENTER,
                    }));
                  });

                  [row.content, row.notes].forEach(text => {
                    rowChildren.push(new TableCell({
                      children: [new Paragraph({ children: [new TextRun({ text: sanitizeDocxText(text), size: 20 })] })],
                      verticalAlign: VerticalAlign.CENTER,
                    }));
                  });

                  tableRows.push(new TableRow({ children: rowChildren }));
                });
                return tableRows;
              })()
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "\n\n" }),
            ],
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
              insideHorizontal: { style: BorderStyle.NONE },
              insideVertical: { style: BorderStyle.NONE },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText("Người lập"), bold: true, size: 22 })] }),
                      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText(`Ngày ${safeFormat(plan.startDate, 'dd/MM/yyyy')}`), size: 20 })] }),
                      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText("(Ký, ghi rõ họ tên)"), italics: true, size: 20 })] }),
                      new Paragraph({ 
                        alignment: AlignmentType.CENTER, 
                        children: [new TextRun({ text: sanitizeDocxText(plan.teacherName), bold: true, size: 22 })],
                        spacing: { before: 1700 }
                      }),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText("Duyệt kế hoạch"), bold: true, size: 22 })] }),
                      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText(`Ngày ${safeFormat(plan.startDate, 'dd/MM/yyyy')}`), size: 20 })] }),
                      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText("(Ký, ghi rõ họ tên)"), italics: true, size: 20 })] }),
                      new Paragraph({ 
                        alignment: AlignmentType.CENTER, 
                        children: [new TextRun({ text: sanitizeDocxText(businessInfo.owner), bold: true, size: 22 })],
                        spacing: { before: 1700 }
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `LichBaoGiang_${plan.teacherName}_Tuan${plan.week}.docx`);
  };

  if (isCreating && editingPlan) {
    return (
      <div className="space-y-8 max-w-6xl mx-auto">
        <div className="flex items-center justify-between">
          <div>
            <h2 className="text-2xl font-normal text-neutral-900 dark:text-white">Tạo Lịch báo giảng</h2>
            <p className="text-sm text-neutral-500 dark:text-slate-400">Thiết lập kế hoạch giảng dạy chi tiết cho tuần mới.</p>
          </div>
          <div className="flex items-center gap-3">
            <button
              onClick={() => setIsCreating(false)}
              className="btn-secondary"
            >
              Hủy bỏ
            </button>
            <button
              onClick={savePlan}
              className="btn-primary flex items-center gap-2"
            >
              <Save className="w-4 h-4" />
              Lưu lịch dạy
            </button>
          </div>
        </div>

        <div className="saas-card grid grid-cols-1 md:grid-cols-4 gap-6">
          <div className="space-y-3">
            <label className="text-sm font-semibold text-neutral-900 dark:text-white pl-1">Họ tên giáo viên</label>
            <input
              value={editingPlan.teacherName}
              onChange={(e) => setEditingPlan({ ...editingPlan, teacherName: e.target.value })}
              placeholder="Nhập họ tên..."
              className="w-full bg-white dark:bg-slate-900 border border-neutral-300 dark:border-slate-700 rounded-xl px-4 py-4 text-base font-bold focus:ring-2 focus:ring-primary transition-all shadow-sm"
            />
          </div>
          <div className="space-y-3">
            <label className="text-sm font-semibold text-neutral-900 dark:text-white pl-1">Tuần</label>
            <input
              value={editingPlan.week}
              onChange={(e) => setEditingPlan({ ...editingPlan, week: e.target.value })}
              className="w-full bg-white dark:bg-slate-900 border border-neutral-300 dark:border-slate-700 rounded-xl px-4 py-4 text-base font-bold focus:ring-2 focus:ring-primary transition-all shadow-sm"
            />
          </div>
          <div className="space-y-3">
            <label className="text-sm font-semibold text-neutral-900 dark:text-white pl-1">Ngày bắt đầu</label>
            <input
              type="date"
              value={editingPlan.startDate}
              onChange={(e) => {
                if (!e.target.value) {
                  setEditingPlan({ ...editingPlan, startDate: '', endDate: '' });
                  return;
                }
                const selectedDate = parseISO(e.target.value);
                if (isNaN(selectedDate.getTime())) return;
                const start = startOfWeek(selectedDate, { weekStartsOn: 1 });
                const end = endOfWeek(selectedDate, { weekStartsOn: 1 });
                
                // Update all existing rows with new dates based on their day
                const updatedRows = editingPlan.rows.map(row => {
                  const dayName = String(row.day || '').toLowerCase().trim();
                  const dayMap: Record<string, number> = {
                    'thứ 2': 0, 'thứ hai': 0, 't2': 0, '2': 0,
                    'thứ 3': 1, 'thứ ba': 1, 't3': 1, '3': 1,
                    'thứ 4': 2, 'thứ tư': 2, 't4': 2, '4': 2,
                    'thứ 5': 3, 'thứ năm': 3, 't5': 3, '5': 3,
                    'thứ 6': 4, 'thứ sáu': 4, 't6': 4, '6': 4,
                    'thứ 7': 5, 'thứ bảy': 5, 't7': 5, '7': 5,
                    'chủ nhật': 6, 'cn': 6, '8': 6
                  };
                  const offset = dayMap[dayName] ?? 0;
                  const newDate = addDays(start, offset);
                  return { ...row, date: format(newDate, 'dd/MM/yyyy') };
                });

                setEditingPlan({ 
                  ...editingPlan, 
                  startDate: format(start, 'yyyy-MM-dd'),
                  endDate: format(end, 'yyyy-MM-dd'),
                  rows: updatedRows
                });
              }}
              className="w-full bg-white dark:bg-slate-900 border border-neutral-300 dark:border-slate-700 rounded-xl px-4 py-4 text-base font-bold focus:ring-2 focus:ring-primary transition-all shadow-sm"
            />
          </div>
          <div className="space-y-3">
            <label className="text-sm font-semibold text-neutral-900 dark:text-white pl-1">Ngày kết thúc</label>
            <input
              type="date"
              readOnly
              value={editingPlan.endDate}
              className="w-full bg-neutral-100 dark:bg-slate-800 border border-neutral-300 dark:border-slate-700 rounded-xl px-4 py-4 text-base font-bold text-neutral-600 dark:text-slate-400 cursor-not-allowed shadow-none"
            />
          </div>
        </div>

        <div className="saas-card overflow-hidden !p-0">
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse min-w-[1000px]">
              <thead>
                <tr className="bg-neutral-100 dark:bg-slate-900 border-b-2 border-neutral-200 dark:border-slate-800">
                  <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-24">Thứ, ngày</th>
                  <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-24">Ca dạy</th>
                  <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-20">Lớp</th>
                  <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-32">Môn</th>
                  <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-32">Phân môn</th>
                  <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-20">Tiết</th>
                  <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white">Nội dung bài dạy</th>
                  <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-24">Ghi chú</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-neutral-200 dark:divide-slate-800 font-sans">
                {editingPlan.rows.map((row) => (
                  <tr key={row.id} className="hover:bg-neutral-50/80 dark:hover:bg-slate-900/30 transition-colors">
                    <td className="px-4 py-4">
                      <div className="text-sm font-semibold text-neutral-900 dark:text-white">{row.day}</div>
                      <div className="text-xs text-neutral-500 dark:text-slate-400 font-medium">{row.date}</div>
                    </td>
                    <td className="px-4 py-4">
                      <div className="text-xs font-semibold text-primary bg-primary/10 px-3 py-1.5 rounded-lg w-fit whitespace-nowrap">
                        {row.shift}
                      </div>
                    </td>
                    <td className="px-4 py-4">
                      <select
                        value={row.grade}
                        onChange={(e) => handleRowChange(row.id, 'grade', e.target.value)}
                        className="w-full bg-transparent border-none focus:ring-0 text-sm font-semibold text-primary cursor-pointer"
                      >
                        <option value="">-</option>
                        {Array.from(new Set(classes.map(c => c.grade))).map(g => (
                          <option key={g} value={g}>{g}</option>
                        ))}
                      </select>
                    </td>
                    <td className="px-4 py-4">
                      <select
                        value={row.subject}
                        onChange={(e) => handleRowChange(row.id, 'subject', e.target.value)}
                        className="w-full bg-transparent border-none focus:ring-0 text-sm font-semibold text-neutral-800 dark:text-slate-200 cursor-pointer"
                      >
                        <option value="">-</option>
                        {Array.from(new Set(classes.filter(c => c.grade === row.grade).map(c => c.subject))).map(s => (
                          <option key={s} value={s}>{s}</option>
                        ))}
                      </select>
                    </td>
                    <td className="px-4 py-4">
                      <select
                        value={row.subSubject}
                        onChange={(e) => handleRowChange(row.id, 'subSubject', e.target.value)}
                        className="w-full bg-transparent border-none focus:ring-0 text-sm font-medium text-neutral-600 dark:text-slate-400 cursor-pointer"
                      >
                        <option value="">-</option>
                        {classes.filter(c => c.grade === row.grade && c.subject === row.subject).map(c => (
                          <option key={c.subSubject} value={c.subSubject}>{c.subSubject}</option>
                        ))}
                      </select>
                    </td>
                    <td className="px-4 py-4">
                      <input
                        type="number"
                        value={row.period}
                        onChange={(e) => handleRowChange(row.id, 'period', e.target.value)}
                        className="w-full bg-transparent border-none focus:ring-0 text-sm font-semibold text-neutral-900 dark:text-white text-center [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
                        placeholder="Tiết..."
                      />
                    </td>
                    <td className="px-4 py-4">
                      <div className="flex items-center gap-2">
                        <input
                          value={row.content}
                          onChange={(e) => handleRowChange(row.id, 'content', e.target.value)}
                          placeholder="Nội dung bài học..."
                          className="flex-1 bg-transparent border-none focus:ring-0 text-sm font-medium text-neutral-800 dark:text-slate-200 placeholder-neutral-300 dark:placeholder-slate-700"
                        />
                        {row.grade && row.subject && row.period && (
                          <button
                            onClick={() => autoFillContent(row.id)}
                            className="p-2 text-teal-600 hover:bg-teal-50 dark:hover:bg-teal-900/20 rounded-lg transition-colors"
                            title="Điền tự động từ PPCT"
                          >
                            <BookOpen className="w-4 h-4" />
                          </button>
                        )}
                      </div>
                    </td>
                    <td className="px-4 py-4">
                      <input
                        value={row.notes}
                        onChange={(e) => handleRowChange(row.id, 'notes', e.target.value)}
                        placeholder="..."
                        className="w-full bg-transparent border-none focus:ring-0 text-sm font-medium text-neutral-600 dark:text-slate-500"
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-white dark:bg-slate-900 p-5 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm">
        <div>
          <h3 className="text-lg font-bold text-slate-900 dark:text-white">Lịch báo giảng</h3>
          <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">Quản lý và theo dõi kế hoạch giảng dạy hàng tuần.</p>
        </div>
        <button
          onClick={startNewPlan}
          className="btn-primary flex items-center gap-2 text-xs"
        >
          <Plus className="w-4 h-4" />
          Tạo lịch mới
        </button>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
        <AnimatePresence mode="popLayout">
          {plans.map((plan, idx) => (
            <motion.div 
              key={plan.id}
              initial={{ opacity: 0, y: 12 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ delay: idx * 0.04 }}
              className="saas-card group hover:border-teal-400 dark:hover:border-teal-600 transition-colors"
            >
              <div className="flex justify-between items-start mb-6">
                <div className="w-12 h-12 bg-primary/10 rounded-2xl flex items-center justify-center text-primary">
                  <CalendarDays className="w-6 h-6" />
                </div>
                <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-all">
                  <button 
                    onClick={() => { setEditingPlan(plan); setIsCreating(true); }} 
                    className="w-8 h-8 rounded-lg flex items-center justify-center text-neutral-400 hover:text-primary hover:bg-primary/10 transition-all"
                  >
                    <FileText className="w-4 h-4" />
                  </button>
                  <ConfirmButton 
                    onConfirm={() => deletePlan(plan.id)} 
                    className="w-8 h-8 rounded-lg flex items-center justify-center text-neutral-400 hover:text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 transition-all"
                    icon={Trash2}
                    confirmText=""
                  />
                </div>
              </div>
              
              <h3 className="text-lg font-bold text-neutral-900 dark:text-white mb-1">
                {plan.teacherName || 'Giáo viên chưa định danh'}
              </h3>
              <div className="flex items-center gap-2 text-xs text-neutral-500 dark:text-slate-400 mb-6">
                <span className="bg-neutral-100 dark:bg-slate-800 px-2 py-1 rounded-lg font-bold">Tuần {plan.week}</span>
                <span className="font-semibold">{safeFormat(plan.startDate, 'dd/MM/yyyy')} - {safeFormat(plan.endDate, 'dd/MM/yyyy')}</span>
              </div>

              <button
                onClick={() => exportToWord(plan)}
                className="w-full btn-secondary flex items-center justify-center gap-2 py-2.5"
              >
                <Download className="w-4 h-4" />
                Tải file Word
              </button>
            </motion.div>
          ))}
        </AnimatePresence>

        {plans.length === 0 && (
          <div className="col-span-full py-24 text-center border-2 border-dashed border-neutral-200 dark:border-slate-800 rounded-3xl">
            <div className="w-20 h-20 bg-neutral-50 dark:bg-slate-900 rounded-3xl flex items-center justify-center mx-auto mb-6 text-neutral-300 dark:text-slate-700">
              <Plus className="w-10 h-10" />
            </div>
            <h4 className="text-xl font-normal text-neutral-900 dark:text-white">Chưa có lịch dạy</h4>
            <p className="text-sm text-neutral-500 dark:text-slate-400 mt-2">Bắt đầu bằng cách tạo một lịch báo giảng mới cho tuần này.</p>
          </div>
        )}
      </div>
    </div>
  );
}

function ClassJournalSection({ 
  plans, 
  setPlans, 
  deletePlan,
  businessInfo,
  setActiveTab
}: { 
  key?: string,
  plans: LessonPlan[], 
  setPlans: (p: LessonPlan[]) => void,
  deletePlan: (id: string) => void,
  businessInfo: BusinessInfo,
  setActiveTab: (t: string) => void
}) {
  const [selectedPlanId, setSelectedPlanId] = useState<string>('');
  const selectedPlan = useMemo(() => plans.find(p => p.id === selectedPlanId), [plans, selectedPlanId]);

  const handleRowChange = (rowId: string, field: 'attendance' | 'comments' | 'signature', value: string) => {
    if (!selectedPlanId) return;
    const newPlans = plans.map(plan => {
      if (plan.id === selectedPlanId) {
        const newRows = plan.rows.map(row => {
          if (row.id === rowId) {
            return { ...row, [field]: value };
          }
          return row;
        });
        return { ...plan, rows: newRows };
      }
      return plan;
    });
    setPlans(newPlans);
  };

  const exportToWord = async () => {
    if (!selectedPlan) return;

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 1134, bottom: 567, left: 1134, right: 1418 }
          }
        },
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: sanitizeDocxText(`Hộ kinh doanh: ${businessInfo.name}`), bold: true, size: 22 }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: sanitizeDocxText(`Địa chỉ: ${businessInfo.address}`), bold: true, size: 22 }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 400, after: 200 },
            children: [
              new TextRun({ text: "SỔ ĐẦU BÀI", bold: true, size: 28 }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: sanitizeDocxText(`Tuần: ${selectedPlan.week} - Từ ngày: ${safeFormat(selectedPlan.startDate, 'dd/MM/yyyy')} - Đến ngày: ${safeFormat(selectedPlan.endDate, 'dd/MM/yyyy')}`), bold: true, size: 22 }),
            ],
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  "Thứ ngày tháng", "Buổi", "Lớp", "Môn học", "Phân môn", "Tiết theo KHDH", "Tên bài, nội dung công việc", "Nhận xét của giáo viên", "Giáo viên dạy/ký tên"
                ].map(text => new TableCell({
                  children: [new Paragraph({ 
                    alignment: AlignmentType.CENTER, 
                    children: [new TextRun({ text: sanitizeDocxText(text), bold: true, size: 22 })] 
                  })],
                  verticalAlign: VerticalAlign.CENTER,
                  shading: { fill: "F2F2F2" }
                })),
              }),
              ...(() => {
                const rowsToExport = selectedPlan.rows.filter(r => r.grade);
                const tableRows: TableRow[] = [];
                let lastDay = "";
                
                const remarksList = ["Chú ý", "Nghiêm túc", "Lớp học tốt"];
                rowsToExport.forEach((row, rowIndex) => {
                  const dayStr = row.day + (row.date ? ` (${row.date})` : '');
                  const dayText = sanitizeDocxText(dayStr);
                  const isNewDay = dayText !== lastDay;
                  if (isNewDay) lastDay = dayText;

                  const rowChildren: TableCell[] = [];
                  
                  if (isNewDay) {
                    rowChildren.push(new TableCell({
                      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText(dayText), size: 20 })] })],
                      verticalAlign: VerticalAlign.CENTER,
                      rowSpan: rowsToExport.filter(r => (r.day + (r.date ? ` (${r.date})` : '')) === dayStr).length,
                    }));
                  }

                  // Split shift
                  const shiftMatch = row.shift.match(/(Ca \d+)\s*(.*)/);
                  const shiftLine1 = shiftMatch ? shiftMatch[1] : row.shift;
                  const shiftLine2 = shiftMatch ? shiftMatch[2] : "";

                  rowChildren.push(new TableCell({
                    children: [
                      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText(shiftLine1), size: 20 })] }),
                      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText(shiftLine2), size: 18 })] }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                  }));

                  // Lớp, Môn, Phân môn, Tiết - Centered
                  [row.grade, row.subject, row.subSubject, row.period].forEach(text => {
                    rowChildren.push(new TableCell({
                      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText(text), size: 20 })] })],
                      verticalAlign: VerticalAlign.CENTER,
                    }));
                  });

                  // Content, Comments, Signature - Left aligned
                  const autoRemark = row.comments || remarksList[rowIndex % remarksList.length];
                  const teacherName = row.signature || selectedPlan.teacherName;
                  
                  [row.content, autoRemark, teacherName].forEach(text => {
                    rowChildren.push(new TableCell({
                      children: [new Paragraph({ children: [new TextRun({ text: sanitizeDocxText(text), size: 20 })] })],
                      verticalAlign: VerticalAlign.CENTER,
                    }));
                  });

                  tableRows.push(new TableRow({ children: rowChildren }));
                });
                return tableRows;
              })()
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "\n\n" }),
            ],
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
              insideHorizontal: { style: BorderStyle.NONE },
              insideVertical: { style: BorderStyle.NONE },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({ children: [], width: { size: 60, type: WidthType.PERCENTAGE } }),
                  new TableCell({
                    width: { size: 40, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({ text: sanitizeDocxText("Xác nhận của Hộ Kinh doanh"), bold: true, size: 22 }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({ text: sanitizeDocxText(`Ngày ${safeFormat(selectedPlan.endDate, 'dd/MM/yyyy')}`), size: 20 }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({ text: sanitizeDocxText("(Ký, ghi rõ họ tên)"), italics: true, size: 20 }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({ text: sanitizeDocxText(businessInfo.owner), bold: true, size: 22 }),
                        ],
                        spacing: { before: 1700 } // Approx 3cm spacing
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `SoDauBai_Tuan${selectedPlan.week}.docx`);
  };

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="space-y-6"
    >
      <header className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-neutral-900 dark:text-white tracking-tight">Sổ đầu bài</h2>
          <p className="text-neutral-500 dark:text-slate-400 mt-1 font-medium">Ghi chép tình hình lớp học dựa trên lịch báo giảng.</p>
        </div>
        <div className="flex flex-wrap items-center gap-3">
          {selectedPlan && (
            <>
              <ConfirmButton
                onConfirm={() => {
                  deletePlan(selectedPlan.id);
                  setSelectedPlanId('');
                }}
                className="flex items-center gap-2 px-4 py-2 text-red-600 border border-red-200 dark:border-red-900/30 rounded-xl hover:bg-red-50 dark:hover:bg-red-900/20 transition-all font-bold text-sm"
                icon={Trash2}
              >
                Xóa lịch này
              </ConfirmButton>
              <button
                onClick={exportToWord}
                className="flex items-center gap-2 px-5 py-2.5 bg-primary text-white rounded-xl hover:bg-primary-hover transition-all shadow-lg shadow-primary/20 font-bold text-sm"
              >
                <Download className="w-4 h-4" /> Xuất Sổ đầu bài
              </button>
            </>
          )}
        </div>
      </header>

      <div className="saas-card p-0 overflow-hidden">
        <div className="p-6 border-b border-neutral-100 dark:border-slate-800 bg-neutral-50/50 dark:bg-slate-800/30">
          <div className="max-w-md">
            <label className="text-sm font-semibold text-neutral-900 dark:text-white block mb-3">Chọn lịch báo giảng để ghi sổ</label>
            <div className="relative">
              <select
                value={selectedPlanId}
                onChange={(e) => setSelectedPlanId(e.target.value)}
                className="w-full pl-4 pr-10 py-4 bg-white dark:bg-slate-900 rounded-xl border border-neutral-300 dark:border-slate-700 focus:ring-2 focus:ring-primary outline-none transition-all appearance-none text-base font-bold text-neutral-900 dark:text-white"
              >
                <option value="">-- Chọn lịch báo giảng --</option>
                {plans.map(p => (
                  <option key={p.id} value={p.id}>
                    {p.teacherName} - Tuần {p.week} ({safeFormat(p.startDate, 'dd/MM/yyyy')} - {safeFormat(p.endDate, 'dd/MM/yyyy')})
                  </option>
                ))}
              </select>
              <div className="absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none text-neutral-400">
                <ChevronDown className="w-4 h-4" />
              </div>
            </div>
          </div>
          {selectedPlan && (
            <div className="mt-4">
              <button
                onClick={() => {
                  const remarksList = ["Chú ý", "Nghiêm túc", "Lớp học tốt"];
                  const newPlans = plans.map(p => {
                    if (p.id === selectedPlanId) {
                      const newRows = p.rows.map((row, idx) => ({
                        ...row,
                        comments: row.comments || remarksList[idx % remarksList.length],
                        signature: row.signature || p.teacherName
                      }));
                      return { ...p, rows: newRows };
                    }
                    return p;
                  });
                  setPlans(newPlans);
                }}
                className="text-xs font-bold text-primary hover:underline flex items-center gap-1"
              >
                <Sparkles className="w-3 h-3" /> Tự động điền nhận xét & chữ ký cho lịch này
              </button>
            </div>
          )}
        </div>

        <div className="p-0">
          {selectedPlan ? (
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse min-w-[1200px]">
                <thead>
                  <tr className="bg-neutral-100 dark:bg-slate-800 border-b-2 border-neutral-200 dark:border-slate-800">
                    <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-32">Thứ, ngày</th>
                    <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-28">Buổi</th>
                    <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-20">Lớp</th>
                    <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-32">Môn</th>
                    <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white">Nội dung bài dạy</th>
                    <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-64">Nhận xét</th>
                    <th className="px-4 py-5 text-xs font-semibold text-neutral-900 dark:text-white w-40">Chữ ký</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-neutral-200 dark:divide-slate-800">
                  {selectedPlan.rows.filter(r => r.grade).map((row) => (
                    <tr key={row.id} className="hover:bg-neutral-50 dark:hover:bg-slate-800/50 transition-colors group">
                      <td className="px-4 py-5">
                        <div className="font-semibold text-neutral-900 dark:text-white text-base">{row.day}</div>
                        <div className="text-xs text-neutral-500 dark:text-slate-400 font-medium mt-0.5">{row.date}</div>
                      </td>
                      <td className="px-4 py-5 text-sm font-semibold text-neutral-800 dark:text-slate-200">{row.shift}</td>
                      <td className="px-4 py-5">
                        <span className="inline-flex items-center px-3 py-1 rounded-lg bg-primary/10 text-primary text-xs font-semibold border border-primary/20 uppercase tracking-wider">
                          {row.grade}
                        </span>
                      </td>
                      <td className="px-4 py-5 text-sm font-semibold text-neutral-800 dark:text-slate-200">{row.subject}</td>
                      <td className="px-4 py-5 text-sm text-neutral-900 dark:text-white leading-relaxed font-semibold">{row.content}</td>
                      <td className="px-3 py-3">
                        <input
                          value={row.comments || ''}
                          onChange={(e) => handleRowChange(row.id, 'comments', e.target.value)}
                          placeholder="Lớp học tốt..."
                          className="w-full bg-white dark:bg-slate-900 border border-neutral-300 dark:border-slate-700 focus:border-primary focus:ring-2 focus:ring-primary/10 rounded-lg px-2 py-2 text-sm transition-all outline-none dark:text-white font-semibold"
                        />
                      </td>
                      <td className="px-3 py-3">
                        <input
                          value={row.signature || ''}
                          onChange={(e) => handleRowChange(row.id, 'signature', e.target.value)}
                          placeholder="Ký tên"
                          className="w-full bg-white dark:bg-slate-900 border border-neutral-300 dark:border-slate-700 focus:border-primary focus:ring-2 focus:ring-primary/10 rounded-lg px-2 py-2 text-sm transition-all outline-none dark:text-white font-semibold italic"
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div className="py-24 text-center">
              <div className="w-20 h-20 bg-[#F9FAFB] rounded-full flex items-center justify-center mx-auto mb-4">
                <FileText className="w-10 h-10 text-[#9CA3AF]" />
              </div>
              <h3 className="text-lg font-bold text-[#1F2937]">Chưa chọn lịch báo giảng</h3>
              <p className="text-[#6B7280] mt-2 max-w-xs mx-auto">Vui lòng chọn một lịch báo giảng từ danh sách phía trên để bắt đầu ghi sổ đầu bài.</p>
            </div>
          )}
        </div>
      </div>
    </motion.div>
  );
}

function ProgramManagementSection({ 
  classes, 
  setClasses, 
  ppctData, 
  setPpctData, 
  lessonPlans, 
  setLessonPlans, 
  deletePlan, 
  businessInfo, 
  activeSubTab, 
  setActiveTab,
  currentUser,
  onNotify
}: { 
  classes: ClassSubject[], 
  setClasses: (c: ClassSubject[]) => void, 
  ppctData: PPCTItem[], 
  setPpctData: (d: PPCTItem[]) => void, 
  lessonPlans: LessonPlan[], 
  setLessonPlans: (l: LessonPlan[]) => void, 
  deletePlan: (id: string) => void, 
  businessInfo: BusinessInfo, 
  activeSubTab: string, 
  setActiveTab: (t: string) => void,
  currentUser: UserAccount | null,
  onNotify?: (msg: string, type?: 'success' | 'error' | 'info') => void
}) {
  const effectiveBusinessInfo = businessInfo;
  const subNavItems = [
    { id: 'classes', label: 'Cấu hình lớp học', icon: GraduationCap, description: 'Thiết lập khối lớp, môn học và phân môn áp dụng giảng dạy.' },
    { id: 'ppct', label: 'Phân phối chương trình', icon: BookOpen, description: 'Quản lý khung bài dạy theo từng tiết và khối lớp.' },
    { id: 'lesson-plan', label: 'Lịch báo giảng', icon: CalendarDays, description: 'Lập kế hoạch giảng dạy hàng tuần và xuất Word.' },
    { id: 'journal', label: 'Sổ đầu bài', icon: ClipboardList, description: 'Ghi chép nhận xét buổi học và xuất sổ đầu bài chuẩn.' },
  ];

  const distinctGrades = Array.from(new Set(classes.map(c => c.grade).filter(Boolean))).length;
  const distinctSubjects = Array.from(new Set(classes.map(c => c.subject).filter(Boolean))).length;

  return (
    <div className="space-y-6">
      <SectionBanner
        title="Quản lý Chương trình dạy"
        subtitle="Đồng bộ danh mục khối lớp, phân phối chương trình, lịch báo giảng hàng tuần và sổ đầu bài của giáo viên."
        badge="CHUYÊN MÔN GIẢNG DẠY"
        icon={BookOpen}
        onBack={() => setActiveTab('dashboard')}
        stats={[
          { label: 'Khối lớp đào tạo', value: `${distinctGrades} khối`, sub: `${distinctSubjects} môn học cấu hình` },
          { label: 'Tổng tiết PPCT', value: `${ppctData.length} tiết`, sub: 'Dữ liệu bài dạy chuẩn' },
          { label: 'Lịch báo giảng', value: `${lessonPlans.length} tuần`, sub: 'Đã lập trên hệ thống', highlight: true },
          { label: 'Sổ đầu bài', value: `${lessonPlans.length} sổ`, sub: 'Sẵn sàng xuất Word (.docx)' },
        ]}
        tabs={subNavItems.map(item => ({ id: item.id, label: item.label, icon: item.icon }))}
        activeTab={activeSubTab}
        onTabChange={setActiveTab}
      />

      {activeSubTab === 'program' ? (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
          {subNavItems.map((item) => (
            <button
              key={item.id}
              onClick={() => setActiveTab(item.id)}
              className="bg-white dark:bg-slate-900 p-6 rounded-2xl border border-slate-200 dark:border-slate-800 hover:border-teal-500 dark:hover:border-teal-500 shadow-sm hover:shadow-md transition-all text-left group flex flex-col justify-between"
            >
              <div>
                <div className="w-12 h-12 bg-teal-50 dark:bg-teal-900/30 border border-teal-200/70 dark:border-teal-800 rounded-xl flex items-center justify-center text-teal-700 dark:text-teal-300 mb-4 group-hover:bg-teal-600 group-hover:text-white transition-colors">
                  <item.icon className="w-6 h-6" />
                </div>
                <h3 className="text-base font-bold text-slate-900 dark:text-white group-hover:text-teal-700 dark:group-hover:text-teal-300 transition-colors">
                  {item.label}
                </h3>
                <p className="text-xs text-slate-500 dark:text-slate-400 mt-1.5 leading-relaxed">
                  {item.description}
                </p>
              </div>
              <div className="mt-5 pt-3 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between text-xs font-bold text-teal-700 dark:text-teal-400">
                <span>Mở phân mục</span>
                <ArrowRight className="w-4 h-4" />
              </div>
            </button>
          ))}
        </div>
      ) : (
        <AnimatePresence mode="wait">
          {activeSubTab === 'classes' && (
            <ClassConfigSection key="classes" classes={classes} setClasses={setClasses} setActiveTab={setActiveTab} currentUser={currentUser} onNotify={onNotify} />
          )}
          {activeSubTab === 'ppct' && (
            <PPCTSection key="ppct" ppctData={ppctData} setPpctData={setPpctData} classes={classes} plans={lessonPlans} setPlans={setLessonPlans} setActiveTab={setActiveTab} currentUser={currentUser} onNotify={onNotify} />
          )}
          {activeSubTab === 'lesson-plan' && (
            <LessonPlanSection key="lesson-plan" plans={lessonPlans} setPlans={setLessonPlans} classes={classes} ppctData={ppctData} deletePlan={deletePlan} businessInfo={effectiveBusinessInfo} setActiveTab={setActiveTab} />
          )}
          {activeSubTab === 'journal' && (
            <ClassJournalSection key="journal" plans={lessonPlans} setPlans={setLessonPlans} deletePlan={deletePlan} businessInfo={effectiveBusinessInfo} setActiveTab={setActiveTab} />
          )}
        </AnimatePresence>
      )}
    </div>
  );
}

function StudentManagementSection({
  students,
  setStudents,
  businessInfo,
  activeSubTab,
  setActiveTab,
  currentUser,
  onNotify,
}: {
  students: Student[];
  setStudents: (s: Student[]) => void;
  businessInfo: BusinessInfo;
  activeSubTab: string;
  setActiveTab: (t: string) => void;
  currentUser: UserAccount | null;
  onNotify?: (msg: string, type?: 'success' | 'error' | 'info') => void;
}) {
  const isAdmin = currentUser?.role === 'admin';
  const subNavItems = [
    { id: 'students-list', label: 'Danh sách & Nhập Excel', icon: Upload },
    { id: 'students-export', label: 'Xuất đơn đăng ký học thêm', icon: FileText },
  ];

  const autoClassesCount = useMemo(() => {
    const set = new Set(students.map((s) => deriveAutoClassName(s)));
    return set.size;
  }, [students]);

  const schoolCount = useMemo(() => {
    const set = new Set(students.map((s) => s.school.trim()).filter(Boolean));
    return set.size;
  }, [students]);

  const subjectCount = useMemo(() => {
    const set = new Set(students.map((s) => s.subject.trim()).filter(Boolean));
    return set.size;
  }, [students]);

  const downloadStudentTemplate = () => {
    const rows = [
      [
        'STT',
        'HỌ VÀ TÊN',
        'LỚP',
        'TRƯỜNG',
        'HỌ VÀ TÊN PHỤ HUYNH',
        'SĐT',
        'MÔN ĐĂNG KÍ HỌC',
        'NGÀY ĐĂNG KÍ HỌC',
      ],
      [
        1,
        'Nguyễn Minh Anh',
        '8A1',
        'THCS Đoàn Kết',
        'Nguyễn Văn Hùng',
        '0912345678',
        'Toán',
        format(new Date(), 'dd/MM/yyyy'),
      ],
      [
        2,
        'Trần Bảo Ngọc',
        '9A2',
        'THCS Tân Phong',
        'Trần Thị Lan',
        '0987654321',
        'KHTN',
        format(new Date(), 'dd/MM/yyyy'),
      ],
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [
      { wch: 6 },
      { wch: 24 },
      { wch: 10 },
      { wch: 22 },
      { wch: 24 },
      { wch: 14 },
      { wch: 18 },
      { wch: 18 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'DS_Hoc_Sinh');
    XLSX.writeFile(wb, 'Mau_Danh_Sach_Thong_Tin_Hoc_Sinh.xlsx');
    onNotify?.('Đã tải file mẫu Excel danh sách thông tin học sinh!', 'success');
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!isAdmin) {
      onNotify?.('Chỉ quản trị viên mới có thể nhập danh sách học sinh.', 'error');
      return;
    }
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const dataBuffer = evt.target?.result;
        const wb = XLSX.read(dataBuffer, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];

        // Columns: STT, HỌ VÀ TÊN, LỚP, TRƯỜNG, HỌ VÀ TÊN PHỤ HUYNH, SĐT, MÔN ĐĂNG KÍ HỌC, NGÀY ĐĂNG KÍ HỌC
        const newStudents: Student[] = data
          .slice(1)
          .filter((row) => row && row[1] && String(row[1]).trim() !== '')
          .map((row, idx) => {
            let phone = String(row[5] || '').trim();
            if (phone && !phone.startsWith('0')) {
              phone = '0' + phone;
            }
            const draftStudent: Student = {
              id: crypto.randomUUID(),
              stt: String(row[0] || students.length + idx + 1),
              name: String(row[1] || '').trim(),
              grade: String(row[2] || '').trim(),
              school: String(row[3] || '').trim(),
              parentName: String(row[4] || '').trim(),
              phone,
              subject: String(row[6] || '').trim(),
              registrationDate: String(row[7] || '').trim(),
            };
            draftStudent.assignedClass = deriveAutoClassName(draftStudent);
            return draftStudent;
          });

        if (newStudents.length === 0) {
          onNotify?.('Không tìm thấy dữ liệu học sinh hợp lệ trong file Excel.', 'error');
          return;
        }

        const merged = [...students, ...newStudents];
        setStudents(merged);
        const classSet = new Set(merged.map((s) => deriveAutoClassName(s)));
        onNotify?.(
          `Đã tải lên ${newStudents.length} học sinh và tự động sắp vào ${classSet.size} lớp học!`,
          'success'
        );
      } catch (err) {
        console.error('Student Excel upload error:', err);
        onNotify?.('Có lỗi khi đọc file Excel. Vui lòng sử dụng đúng file mẫu.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const getRegistrationFormChildren = (student: Student) => {
    const parseDate = (dateStr: string) => {
      if (!dateStr) return { d: '......', m: '......', y: '......' };
      const parts = dateStr.split(/[\/\-\.]/);
      if (parts.length === 3) {
        return { d: parts[0].padStart(2, '0'), m: parts[1].padStart(2, '0'), y: parts[2] };
      }
      const d = new Date(dateStr);
      if (!isNaN(d.getTime())) {
        return { 
          d: String(d.getDate()).padStart(2, '0'), 
          m: String(d.getMonth() + 1).padStart(2, '0'), 
          y: String(d.getFullYear()) 
        };
      }
      return { d: '......', m: '......', y: '......' };
    };

    const { d, m, y } = parseDate(student.registrationDate);

    return [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"), bold: true, size: 28 }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("Độc lập - Tự do - Hạnh phúc"), bold: true, size: 26 }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("----------***----------"), size: 24 }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 400, line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("ĐƠN ĐĂNG KÍ HỌC THÊM"), bold: true, size: 32 }),
        ],
      }),
      new Paragraph({
        indent: { left: 1134 },
        spacing: { line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("Kính gửi: "), italics: true, size: 28 }),
          new TextRun({ text: sanitizeDocxText(businessInfo.name || "................................................................"), bold: true, size: 28 }),
        ],
      }),
      new Paragraph({ spacing: { before: 200, line: 312 } }),
      new Paragraph({
        indent: { firstLine: 567 },
        spacing: { line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("Tôi tên là: "), size: 28 }),
          new TextRun({ text: sanitizeDocxText(student.parentName || "................................................................"), size: 28 }),
        ],
      }),
      new Paragraph({
        indent: { firstLine: 567 },
        spacing: { line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("Số điện thoại: "), size: 28 }),
          new TextRun({ text: sanitizeDocxText(student.phone || "................................................................"), size: 28 }),
        ],
      }),
      new Paragraph({
        indent: { firstLine: 567 },
        spacing: { line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("Là Phụ huynh của học sinh: "), size: 28 }),
          new TextRun({ text: sanitizeDocxText(student.name || "................................................................"), size: 28 }),
        ],
      }),
      new Paragraph({
        indent: { firstLine: 567 },
        spacing: { line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("Lớp: "), size: 28 }),
          new TextRun({ text: sanitizeDocxText(student.grade || "................................................................"), size: 28 }),
        ],
      }),
      new Paragraph({
        indent: { firstLine: 567 },
        spacing: { line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("Đang học tại trường: "), size: 28 }),
          new TextRun({ text: sanitizeDocxText(student.school || "................................................................"), size: 28 }),
        ],
      }),
      new Paragraph({
        indent: { firstLine: 567 },
        spacing: { line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("Môn đăng kí học: "), size: 28 }),
          new TextRun({ text: sanitizeDocxText(student.subject || "................................................................"), size: 28 }),
        ],
      }),
      new Paragraph({
        indent: { firstLine: 567 },
        spacing: { before: 200, line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText(`Tôi viết đơn này đăng kí học thêm môn ${student.subject || '........'} trong năm 2026, do `), size: 28 }),
          new TextRun({ text: sanitizeDocxText(businessInfo.name || "................"), size: 28 }),
          new TextRun({ text: sanitizeDocxText(" tổ chức tại "), size: 28 }),
          new TextRun({ text: sanitizeDocxText(businessInfo.address || "................"), size: 28 }),
        ],
      }),
      new Paragraph({
        indent: { firstLine: 567 },
        spacing: { before: 200, line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("Tôi xin cam kết đối với con tôi sẽ:"), size: 28 }),
        ],
      }),
      new Paragraph({
        indent: { firstLine: 567 },
        spacing: { line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("+ Chấp hành nghiêm túc nội quy lớp học."), size: 28 }),
        ],
      }),
      new Paragraph({
        indent: { firstLine: 567 },
        spacing: { line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("+ Tham gia học tập đầy đủ, đúng giờ."), size: 28 }),
        ],
      }),
      new Paragraph({
        indent: { firstLine: 567 },
        spacing: { line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("+ Hoàn thành bài tập và chủ động trong học tập."), size: 28 }),
        ],
      }),
      new Paragraph({
        indent: { firstLine: 567 },
        spacing: { line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("Rất mong cơ sở xem xét và chấp thuận."), size: 28 }),
        ],
      }),
      new Paragraph({
        indent: { firstLine: 567 },
        spacing: { before: 400, line: 312 },
        children: [
          new TextRun({ text: sanitizeDocxText("Tôi xin trân trọng cảm ơn!"), italics: true, size: 28 }),
        ],
      }),
      new Paragraph({ spacing: { before: 400, line: 312 } }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          top: { style: BorderStyle.NONE },
          bottom: { style: BorderStyle.NONE },
          left: { style: BorderStyle.NONE },
          right: { style: BorderStyle.NONE },
          insideHorizontal: { style: BorderStyle.NONE },
          insideVertical: { style: BorderStyle.NONE },
        },
        rows: [
          new TableRow({
            children: [
              new TableCell({ children: [], width: { size: 50, type: WidthType.PERCENTAGE } }),
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { line: 312 },
                    children: [
                      new TextRun({ text: sanitizeDocxText(`Lai châu, ngày ${d} tháng ${m} năm ${y}`), italics: true, size: 26 }),
                    ],
                  }),
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { line: 312 },
                    children: [
                      new TextRun({ text: "NGƯỜI LÀM ĐƠN", bold: true, size: 28 }),
                    ],
                  }),
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { line: 312 },
                    children: [
                      new TextRun({ text: "(Kí và ghi rõ họ tên)", italics: true, size: 24 }),
                    ],
                  }),
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 1400, line: 312 },
                    children: [
                      new TextRun({ text: sanitizeDocxText(student.parentName || ""), bold: true, size: 28 }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ];
  };

  const exportRegistrationForm = async (student: Student) => {
    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 1134, bottom: 567, left: 1531, right: 1134 }
          }
        },
        children: getRegistrationFormChildren(student),
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `DonDangKiHocThem_${student.name}.docx`);
  };

  const exportAllRegistrationForms = async () => {
    if (students.length === 0) return;
    
    const children: any[] = [];
    students.forEach((student, index) => {
      children.push(...getRegistrationFormChildren(student));
      if (index < students.length - 1) {
        children.push(new Paragraph({ children: [new PageBreak()] }));
      }
    });

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 1134, bottom: 567, left: 1134, right: 1418 }
          }
        },
        children: children,
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `Tat_Ca_Don_Dang_Ky.docx`);
  };

  return (
    <div className="space-y-6">
      <SectionBanner
        title="Quản lý Học sinh"
        subtitle="Quản lý hồ sơ thông tin học sinh tập trung, xuất file mẫu Excel chuẩn, tự động phân lớp học thêm và in đơn đăng ký."
        badge="HỒ SƠ HỌC SINH"
        icon={Users}
        onBack={() => setActiveTab('dashboard')}
        stats={[
          { label: 'Tổng học sinh', value: `${students.length} HS`, sub: 'Đang theo học tại cơ sở', highlight: true },
          { label: 'Lớp tự động xếp', value: `${autoClassesCount} lớp`, sub: 'Tự động theo môn & khối' },
          { label: 'Trường phổ thông', value: `${schoolCount} trường`, sub: 'Đơn vị trường đang học' },
          { label: 'Môn đăng ký', value: `${subjectCount} môn`, sub: 'Môn học thêm đăng ký' },
        ]}
        tabs={subNavItems}
        activeTab={activeSubTab}
        onTabChange={setActiveTab}
        actions={
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={downloadStudentTemplate}
              className="inline-flex items-center gap-2 px-3.5 py-2.5 rounded-xl bg-white text-teal-900 hover:bg-teal-50 font-bold text-xs shadow-sm transition-colors"
            >
              <FileSpreadsheet className="w-4 h-4 text-teal-700" />
              Xuất file mẫu DS Học sinh (.xlsx)
            </button>
            <button
              onClick={() => setActiveTab('student_learning_group')}
              className="inline-flex items-center gap-2 px-3.5 py-2.5 rounded-xl bg-white/15 hover:bg-white/25 text-white border border-white/25 font-bold text-xs transition-colors"
            >
              <GraduationCap className="w-4 h-4" />
              Xếp lớp & Chấm công
            </button>
          </div>
        }
      />

      {activeSubTab === 'students_group' ? (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
          <button
            onClick={() => setActiveTab('students-list')}
            className="bg-white dark:bg-slate-900 p-6 rounded-2xl border border-slate-200 dark:border-slate-800 hover:border-teal-500 shadow-sm hover:shadow-md transition-all text-left group flex flex-col justify-between"
          >
            <div>
              <div className="w-12 h-12 bg-teal-50 dark:bg-teal-900/30 border border-teal-200/70 dark:border-teal-800 rounded-xl flex items-center justify-center text-teal-700 dark:text-teal-300 mb-4 group-hover:bg-teal-600 group-hover:text-white transition-colors">
                <Upload className="w-6 h-6" />
              </div>
              <h3 className="text-base font-bold text-slate-900 dark:text-white">Tải danh sách học sinh</h3>
              <p className="text-xs text-slate-500 dark:text-slate-400 mt-1.5 leading-relaxed">
                Đưa danh sách thông tin học sinh từ file Excel lên hệ thống và tự động xếp lớp học thêm.
              </p>
            </div>
            <div className="mt-5 pt-3 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between text-xs font-bold text-teal-700 dark:text-teal-400">
              <span>Mở danh sách</span>
              <ArrowRight className="w-4 h-4" />
            </div>
          </button>

          <button
            onClick={downloadStudentTemplate}
            className="bg-white dark:bg-slate-900 p-6 rounded-2xl border border-slate-200 dark:border-slate-800 hover:border-teal-500 shadow-sm hover:shadow-md transition-all text-left group flex flex-col justify-between"
          >
            <div>
              <div className="w-12 h-12 bg-emerald-50 dark:bg-emerald-900/30 border border-emerald-200/70 dark:border-emerald-800 rounded-xl flex items-center justify-center text-emerald-700 dark:text-emerald-300 mb-4 group-hover:bg-emerald-600 group-hover:text-white transition-colors">
                <FileSpreadsheet className="w-6 h-6" />
              </div>
              <h3 className="text-base font-bold text-slate-900 dark:text-white">Xuất file mẫu Excel DS Học sinh</h3>
              <p className="text-xs text-slate-500 dark:text-slate-400 mt-1.5 leading-relaxed">
                Tải về bảng mẫu Excel chuẩn gồm đầy đủ các cột thông tin học sinh, phụ huynh, SĐT và môn học.
              </p>
            </div>
            <div className="mt-5 pt-3 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between text-xs font-bold text-emerald-700 dark:text-emerald-400">
              <span>Tải file mẫu (.xlsx)</span>
              <Download className="w-4 h-4" />
            </div>
          </button>

          <button
            onClick={() => setActiveTab('students-export')}
            className="bg-white dark:bg-slate-900 p-6 rounded-2xl border border-slate-200 dark:border-slate-800 hover:border-teal-500 shadow-sm hover:shadow-md transition-all text-left group flex flex-col justify-between"
          >
            <div>
              <div className="w-12 h-12 bg-teal-50 dark:bg-teal-900/30 border border-teal-200/70 dark:border-teal-800 rounded-xl flex items-center justify-center text-teal-700 dark:text-teal-300 mb-4 group-hover:bg-teal-600 group-hover:text-white transition-colors">
                <FileText className="w-6 h-6" />
              </div>
              <h3 className="text-base font-bold text-slate-900 dark:text-white">Xuất đơn đăng ký học thêm</h3>
              <p className="text-xs text-slate-500 dark:text-slate-400 mt-1.5 leading-relaxed">
                Kết xuất đơn đăng ký học thêm định dạng Word (.docx) cho từng học sinh hoặc toàn bộ cơ sở.
              </p>
            </div>
            <div className="mt-5 pt-3 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between text-xs font-bold text-teal-700 dark:text-teal-400">
              <span>Xuất hồ sơ đơn</span>
              <ArrowRight className="w-4 h-4" />
            </div>
          </button>
        </div>
      ) : (
        <div className="space-y-5">
          <div className="bg-white dark:bg-slate-900 p-5 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm flex flex-col lg:flex-row lg:items-center justify-between gap-4">
            <div>
              <h3 className="text-lg font-bold text-slate-900 dark:text-white">
                {activeSubTab === 'students-list'
                  ? 'Bảng Danh sách Thông tin Học sinh'
                  : 'Xuất Đơn Đăng ký Học thêm (.docx)'}
              </h3>
              <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
                {activeSubTab === 'students-list'
                  ? 'Tải file mẫu Excel để điền danh sách học sinh, sau khi đưa lên hệ thống sẽ tự động xếp lớp học thêm.'
                  : 'Tải đơn đăng ký học thêm chuẩn cho từng cá nhân hoặc xuất hàng loạt toàn bộ học sinh.'}
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-2.5">
              <button
                onClick={downloadStudentTemplate}
                className="inline-flex items-center gap-2 px-4 py-2.5 rounded-xl bg-emerald-50 dark:bg-emerald-900/30 text-emerald-700 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800 hover:bg-emerald-100 dark:hover:bg-emerald-900/50 font-bold text-xs transition-colors"
              >
                <FileSpreadsheet className="w-4 h-4" />
                Xuất file mẫu Excel (DS Học sinh)
              </button>

              {activeSubTab === 'students-export' && (
                <button
                  onClick={exportAllRegistrationForms}
                  disabled={students.length === 0}
                  className="btn-primary flex items-center gap-2 text-xs disabled:opacity-50"
                >
                  <Download className="w-4 h-4" /> Xuất toàn bộ đơn (.docx)
                </button>
              )}

              {activeSubTab === 'students-list' && isAdmin && (
                <>
                  {students.length > 0 && (
                    <ConfirmButton
                      onConfirm={() => {
                        setStudents([]);
                        onNotify?.('Đã xóa toàn bộ danh sách học sinh.', 'info');
                      }}
                      className="flex items-center gap-2 px-3.5 py-2.5 text-red-600 border border-red-200 dark:border-red-900/40 rounded-xl hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors font-bold text-xs"
                      icon={Trash2}
                    >
                      Xóa toàn bộ
                    </ConfirmButton>
                  )}
                  <label className="btn-primary flex items-center gap-2 cursor-pointer text-xs">
                    <Upload className="w-4 h-4" /> Đưa danh sách lên (Excel)
                    <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} className="hidden" />
                  </label>
                </>
              )}
            </div>
          </div>

          <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-800 overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse min-w-[980px]">
                <thead>
                  <tr className="bg-slate-50 dark:bg-slate-800/70 border-b border-slate-200 dark:border-slate-800">
                    <th className="px-4 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider w-14 text-center">STT</th>
                    <th className="px-4 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Họ và tên học sinh</th>
                    <th className="px-4 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider w-24 text-center">Lớp PT</th>
                    <th className="px-4 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Trường học</th>
                    <th className="px-4 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Họ tên Phụ huynh</th>
                    <th className="px-4 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">SĐT liên hệ</th>
                    <th className="px-4 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Môn đăng ký</th>
                    <th className="px-4 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Lớp xếp tự động</th>
                    <th className="px-4 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider w-28 text-center">Thao tác</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                  {students.map((student, idx) => (
                    <tr key={student.id} className="hover:bg-teal-50/30 dark:hover:bg-slate-800/40 transition-colors">
                      <td className="px-4 py-3.5 text-xs font-mono font-semibold text-slate-500 dark:text-slate-400 text-center tabular-nums">
                        {student.stt || idx + 1}
                      </td>
                      <td className="px-4 py-3.5 text-sm font-bold text-slate-900 dark:text-white">{student.name}</td>
                      <td className="px-4 py-3.5 text-center">
                        <span className="inline-flex items-center px-2.5 py-0.5 rounded-md bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-200 text-xs font-bold">
                          {student.grade || '-'}
                        </span>
                      </td>
                      <td className="px-4 py-3.5 text-sm text-slate-600 dark:text-slate-300">{student.school}</td>
                      <td className="px-4 py-3.5 text-sm text-slate-700 dark:text-slate-300 font-medium">{student.parentName}</td>
                      <td className="px-4 py-3.5 text-sm font-mono tabular-nums text-slate-700 dark:text-slate-300">{student.phone}</td>
                      <td className="px-4 py-3.5 text-sm font-semibold text-teal-700 dark:text-teal-300">{student.subject}</td>
                      <td className="px-4 py-3.5">
                        <span className="inline-flex items-center px-2.5 py-1 rounded-lg bg-teal-50 dark:bg-teal-900/30 text-teal-800 dark:text-teal-200 border border-teal-200/70 dark:border-teal-800 text-xs font-bold">
                          {deriveAutoClassName(student)}
                        </span>
                      </td>
                      <td className="px-4 py-3.5 text-sm text-center">
                        <div className="flex justify-center gap-1.5">
                          {activeSubTab === 'students-export' && (
                            <button
                              onClick={() => exportRegistrationForm(student)}
                              title="Xuất đơn đăng ký (.docx)"
                              className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-bold text-teal-700 dark:text-teal-300 bg-teal-50 dark:bg-teal-900/30 hover:bg-teal-100 rounded-lg transition-colors"
                            >
                              <FileText className="w-3.5 h-3.5" />
                              Đơn
                            </button>
                          )}
                          {activeSubTab === 'students-list' && isAdmin && (
                            <ConfirmButton
                              onConfirm={() => setStudents(students.filter((s) => s.id !== student.id))}
                              className="w-8 h-8 flex items-center justify-center text-slate-400 hover:text-red-600 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-lg transition-colors"
                              icon={Trash2}
                              confirmText=""
                            />
                          )}
                        </div>
                      </td>
                    </tr>
                  ))}
                  {students.length === 0 && (
                    <tr>
                      <td colSpan={9} className="px-6 py-16 text-center">
                        <div className="w-16 h-16 bg-teal-50 dark:bg-teal-900/20 rounded-2xl flex items-center justify-center mx-auto mb-4 text-teal-600 dark:text-teal-400">
                          <Users className="w-8 h-8" />
                        </div>
                        <h3 className="text-base font-bold text-slate-900 dark:text-white">Chưa có danh sách học sinh</h3>
                        <p className="text-xs text-slate-500 dark:text-slate-400 mt-1 max-w-md mx-auto">
                          Nhấn nút <strong>"Xuất file mẫu Excel (DS Học sinh)"</strong> để tải mẫu chuẩn, điền thông tin học sinh và đưa lên hệ thống để tự động sắp lớp.
                        </p>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function LoginPage({
  onLogin,
  users,
  darkMode,
  setDarkMode,
}: {
  onLogin: (user: UserAccount) => void;
  users: UserAccount[];
  darkMode: boolean;
  setDarkMode: (v: boolean) => void;
}) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState('');

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const user = users.find(
      (u) => u.email.toLowerCase() === email.trim().toLowerCase() && u.password === password.trim()
    );

    if (user) {
      if (user.expiryDate && new Date(user.expiryDate) < new Date()) {
        setError('Tài khoản đã hết hạn sử dụng.');
        return;
      }
      onLogin(user);
    } else {
      setError('Tài khoản hoặc mật khẩu không chính xác.');
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-emerald-950 via-teal-900 to-teal-800 flex items-center justify-center p-4 sm:p-6 relative overflow-hidden">
      <div
        className="absolute inset-0 opacity-15 pointer-events-none"
        style={{
          backgroundImage: 'radial-gradient(circle at 2px 2px, rgba(255,255,255,0.45) 1px, transparent 0)',
          backgroundSize: '28px 28px',
        }}
      />

      <div className="max-w-md w-full bg-white dark:bg-slate-900 rounded-3xl shadow-2xl border border-teal-500/30 overflow-hidden relative z-10">
        {/* Emerald-Teal Top Header Band */}
        <div className="bg-gradient-to-r from-emerald-800 via-teal-700 to-teal-600 px-6 sm:px-8 pt-7 pb-6 text-white relative">
          <div className="absolute top-4 right-4">
            <button
              type="button"
              onClick={() => setDarkMode(!darkMode)}
              className="p-2 rounded-xl bg-white/15 hover:bg-white/25 text-white transition-colors"
              title="Đổi giao diện Sáng / Tối"
            >
              {darkMode ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
            </button>
          </div>

          <div className="flex flex-col items-center text-center">
            <div className="w-20 h-20 bg-white rounded-2xl flex items-center justify-center p-2.5 shadow-lg mb-4 border-2 border-teal-200/60">
              <Logo className="w-full h-full" />
            </div>
            <span className="text-[11px] font-extrabold uppercase tracking-widest text-teal-100 bg-white/10 px-3 py-1 rounded-md border border-white/15 mb-2">
              HỆ THỐNG QUẢN LÝ HOÀNG GIA
            </span>
            <h1 className="text-xl sm:text-2xl font-extrabold tracking-tight text-white">
              Quản lý Cơ sở Dạy thêm
            </h1>
            <p className="text-xs text-teal-100 mt-1 font-medium">
              Đăng nhập điều hành học vụ, điểm danh, tài chính & báo cáo thuế HKD
            </p>
          </div>
        </div>

        {/* Form Body */}
        <div className="p-6 sm:p-8 bg-teal-50/35 dark:bg-slate-900">
          <div className="mb-5 flex flex-wrap justify-center gap-1.5">
            <span className="px-2.5 py-1 bg-teal-100/80 dark:bg-teal-900/40 text-teal-800 dark:text-teal-200 text-[11px] font-bold rounded-md border border-teal-200/80 dark:border-teal-800">
              Quản lý học sinh
            </span>
            <span className="px-2.5 py-1 bg-emerald-100/80 dark:bg-emerald-900/40 text-emerald-800 dark:text-emerald-200 text-[11px] font-bold rounded-md border border-emerald-200/80 dark:border-emerald-800">
              Học tập & Chấm công
            </span>
            <span className="px-2.5 py-1 bg-teal-100/80 dark:bg-teal-900/40 text-teal-800 dark:text-teal-200 text-[11px] font-bold rounded-md border border-teal-200/80 dark:border-teal-800">
              Chương trình dạy
            </span>
            <span className="px-2.5 py-1 bg-emerald-100/80 dark:bg-emerald-900/40 text-emerald-800 dark:text-emerald-200 text-[11px] font-bold rounded-md border border-emerald-200/80 dark:border-emerald-800">
              Tài chính & Báo cáo thuế
            </span>
          </div>

          <form onSubmit={handleSubmit} className="space-y-4">
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-teal-950 dark:text-slate-200 uppercase tracking-wider">
                Tài khoản đăng nhập
              </label>
              <div className="relative">
                <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none">
                  <User className="w-4 h-4 text-teal-600 dark:text-teal-400" />
                </div>
                <input
                  type="text"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  className="w-full pl-10 pr-4 py-3 bg-white dark:bg-slate-800 border border-teal-200 dark:border-slate-700 rounded-xl focus:ring-2 focus:ring-teal-500/30 focus:border-teal-600 dark:text-white text-sm font-medium outline-none transition-colors"
                  placeholder="Nhập email hoặc số điện thoại"
                  required
                />
              </div>
            </div>

            <div className="space-y-1.5">
              <label className="text-xs font-bold text-teal-950 dark:text-slate-200 uppercase tracking-wider">
                Mật khẩu
              </label>
              <div className="relative">
                <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none">
                  <Lock className="w-4 h-4 text-teal-600 dark:text-teal-400" />
                </div>
                <input
                  type={showPassword ? 'text' : 'password'}
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  className="w-full pl-10 pr-11 py-3 bg-white dark:bg-slate-800 border border-teal-200 dark:border-slate-700 rounded-xl focus:ring-2 focus:ring-teal-500/30 focus:border-teal-600 dark:text-white text-sm font-medium outline-none transition-colors"
                  placeholder="Nhập mật khẩu"
                  required
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(!showPassword)}
                  className="absolute inset-y-0 right-0 pr-3.5 flex items-center text-slate-400 hover:text-teal-700 transition-colors"
                >
                  {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                </button>
              </div>
            </div>

            {error && (
              <div className="p-3.5 bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-900/40 rounded-xl flex items-center gap-2.5 text-red-700 dark:text-red-300 text-xs font-semibold">
                <Info className="w-4 h-4 shrink-0" />
                {error}
              </div>
            )}

            <button
              type="submit"
              className="w-full py-3.5 bg-gradient-to-r from-emerald-700 via-teal-600 to-teal-700 hover:from-emerald-800 hover:to-teal-800 text-white rounded-xl font-bold text-sm shadow-md shadow-teal-900/20 flex items-center justify-center gap-2 transition-all cursor-pointer"
            >
              Đăng nhập hệ thống
              <ArrowRight className="w-4 h-4" />
            </button>
          </form>

          <div className="mt-5 pt-4 border-t border-teal-200/60 dark:border-slate-800 space-y-3">
            <PWAInstallButton variant="sidebar" className="w-full justify-center" />
            <p className="text-[11px] text-center text-slate-500 dark:text-slate-400 font-semibold">
              Hỗ trợ máy tính, điện thoại Android & iPhone (iOS) • Đào Minh Tâm - Zalo 0366000555
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}

function UserManagementSection({
  users,
  setUsers,
  setActiveTab,
  onSync,
  onFetch,
  onNotify,
}: {
  users: UserAccount[];
  setUsers: React.Dispatch<React.SetStateAction<UserAccount[]>>;
  setActiveTab: (t: string) => void;
  onSync?: (users: UserAccount[]) => void;
  onFetch?: () => void;
  onNotify?: (msg: string, type?: 'success' | 'error' | 'info') => void;
}) {
  const [newEmail, setNewEmail] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [newRole, setNewRole] = useState<'admin' | 'user'>('user');
  const [newExpiryDate, setNewExpiryDate] = useState('');
  const [newMaxDevices, setNewMaxDevices] = useState(1);
  const [newTaxCode, setNewTaxCode] = useState('');
  const [newBusinessName, setNewBusinessName] = useState('');
  const [newBusinessAddress, setNewBusinessAddress] = useState('');
  const [newBusinessLocation, setNewBusinessLocation] = useState('');
  const [newBusinessOwner, setNewBusinessOwner] = useState('');

  const adminCount = users.filter((u) => u.role === 'admin').length;
  const teacherCount = users.length - adminCount;

  const handleAddUser = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newEmail || !newPassword) return;

    const newUser: UserAccount = {
      id: Math.random().toString(36).substr(2, 9),
      email: newEmail.trim(),
      password: newPassword.trim(),
      role: newRole,
      expiryDate: newExpiryDate || undefined,
      maxDevices: newMaxDevices,
      taxCode: newTaxCode || undefined,
      businessName: newBusinessName || undefined,
      businessAddress: newBusinessAddress || undefined,
      businessLocation: newBusinessLocation || undefined,
      businessOwner: newBusinessOwner || undefined,
      createdAt: new Date().toISOString(),
    };

    const updatedUsers = [...users, newUser];
    setUsers(updatedUsers);
    if (onSync) onSync(updatedUsers);
    onNotify?.(`Đã tạo tài khoản "${newUser.email}" thành công!`, 'success');

    setNewEmail('');
    setNewPassword('');
    setNewExpiryDate('');
    setNewMaxDevices(1);
    setNewTaxCode('');
    setNewBusinessName('');
    setNewBusinessAddress('');
    setNewBusinessLocation('');
    setNewBusinessOwner('');
  };

  const handleDeleteUser = (id: string) => {
    if (users.length <= 1) {
      onNotify?.('Không thể xóa tài khoản quản trị cuối cùng.', 'error');
      return;
    }
    const updatedUsers = users.filter((u) => u.id !== id);
    setUsers(updatedUsers);
    if (onSync) onSync(updatedUsers);
    onNotify?.('Đã xóa tài khoản khỏi hệ thống.', 'info');
  };

  return (
    <div className="space-y-6">
      <SectionBanner
        title="Quản lý Tài khoản Hệ thống"
        subtitle="Phân quyền quản trị viên, cấp tài khoản giáo viên / hộ kinh doanh và đồng bộ dữ liệu đám mây."
        badge="QUẢN TRỊ TÀI KHOẢN"
        icon={ShieldCheck}
        onBack={() => setActiveTab('dashboard')}
        stats={[
          { label: 'Tổng tài khoản', value: `${users.length} tài khoản`, sub: 'Đang hoạt động', highlight: true },
          { label: 'Quản trị viên', value: `${adminCount} Admin`, sub: 'Toàn quyền cấu hình' },
          { label: 'Người dùng / GV', value: `${teacherCount} tài khoản`, sub: 'Tài khoản nghiệp vụ' },
          { label: 'Đồng bộ đám mây', value: 'Google Sheet', sub: 'Sao lưu tức thời' },
        ]}
        actions={
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => setActiveTab('menu_settings')}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-xl bg-white/15 hover:bg-white/25 text-white border border-white/25 font-bold text-xs transition-colors"
            >
              <Settings className="w-4 h-4" />
              Sắp xếp Menu
            </button>
            <button
              onClick={() => {
                if (onSync) onSync(users);
                onNotify?.('Đã gửi yêu cầu đồng bộ danh sách tài khoản lên Google Sheet!', 'success');
              }}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-xl bg-white text-teal-900 hover:bg-teal-50 font-bold text-xs shadow-sm transition-colors"
            >
              <Upload className="w-4 h-4 text-teal-700" />
              Đẩy lên Sheet
            </button>
            <button
              onClick={() => {
                onFetch?.();
                onNotify?.('Đang tải dữ liệu tài khoản mới nhất từ Google Sheet...', 'info');
              }}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-xl bg-white/15 hover:bg-white/25 text-white border border-white/25 font-bold text-xs transition-colors"
            >
              <RefreshCw className="w-4 h-4" />
              Tải từ Sheet
            </button>
          </div>
        }
      />

      <div className="bg-white dark:bg-slate-900 p-6 sm:p-8 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm space-y-6">
        <div>
          <h3 className="text-base font-bold text-slate-900 dark:text-white">Thêm tài khoản mới</h3>
          <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
            Cấp tài khoản đăng nhập kèm thông tin hộ kinh doanh riêng (nếu có).
          </p>
        </div>

        <form onSubmit={handleAddUser} className="space-y-5 p-5 sm:p-6 bg-slate-50 dark:bg-slate-800/50 rounded-2xl border border-slate-200/80 dark:border-slate-700">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">Email / Tên đăng nhập *</label>
              <input
                type="text"
                value={newEmail}
                onChange={(e) => setNewEmail(e.target.value)}
                className="input-field"
                placeholder="VD: giaovien01@gmail.com"
                required
              />
            </div>
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">Mật khẩu *</label>
              <input
                type="password"
                value={newPassword}
                onChange={(e) => setNewPassword(e.target.value)}
                className="input-field"
                placeholder="Nhập mật khẩu"
                required
              />
            </div>
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">Vai trò</label>
              <select
                value={newRole}
                onChange={(e) => setNewRole(e.target.value as 'admin' | 'user')}
                className="input-field"
              >
                <option value="user">Người dùng / Giáo viên</option>
                <option value="admin">Quản trị viên (Admin)</option>
              </select>
            </div>
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">Mã số thuế</label>
              <input
                type="text"
                value={newTaxCode}
                onChange={(e) => setNewTaxCode(e.target.value)}
                className="input-field font-mono tabular-nums"
                placeholder="Mã số thuế HKD"
              />
            </div>
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">Số thiết bị tối đa</label>
              <input
                type="number"
                value={newMaxDevices}
                onChange={(e) => setNewMaxDevices(parseInt(e.target.value) || 1)}
                min="1"
                className="input-field font-mono tabular-nums"
              />
            </div>
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">Ngày hết hạn</label>
              <input
                type="date"
                value={newExpiryDate}
                onChange={(e) => setNewExpiryDate(e.target.value)}
                className="input-field"
              />
            </div>
            <div className="space-y-1.5 md:col-span-2">
              <label className="text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">Tên hộ kinh doanh gắn kèm</label>
              <input
                type="text"
                value={newBusinessName}
                onChange={(e) => setNewBusinessName(e.target.value)}
                className="input-field"
                placeholder="Tên cơ sở / Hộ kinh doanh"
              />
            </div>
            <div className="space-y-1.5">
              <label className="text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">Chủ hộ kinh doanh</label>
              <input
                type="text"
                value={newBusinessOwner}
                onChange={(e) => setNewBusinessOwner(e.target.value)}
                className="input-field"
                placeholder="Họ và tên chủ hộ"
              />
            </div>
            <div className="md:col-span-3 grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">Địa chỉ đăng ký</label>
                <input
                  type="text"
                  value={newBusinessAddress}
                  onChange={(e) => setNewBusinessAddress(e.target.value)}
                  className="input-field"
                  placeholder="Địa chỉ trên giấy phép"
                />
              </div>
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">Địa điểm dạy thực tế</label>
                <input
                  type="text"
                  value={newBusinessLocation}
                  onChange={(e) => setNewBusinessLocation(e.target.value)}
                  className="input-field"
                  placeholder="Địa điểm tổ chức dạy thêm"
                />
              </div>
            </div>
          </div>
          <div className="flex justify-end pt-2">
            <button type="submit" className="btn-primary flex items-center gap-2 text-xs">
              <Plus className="w-4 h-4" />
              Tạo tài khoản mới
            </button>
          </div>
        </form>

        <div className="overflow-x-auto border border-slate-200 dark:border-slate-800 rounded-2xl">
          <table className="w-full text-left border-collapse min-w-[680px]">
            <thead>
              <tr className="bg-slate-50 dark:bg-slate-800/70 border-b border-slate-200 dark:border-slate-800">
                <th className="px-5 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Tài khoản</th>
                <th className="px-5 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Vai trò</th>
                <th className="px-5 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Đơn vị / MST</th>
                <th className="px-5 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider text-right">Thao tác</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
              {users.map((user) => (
                <tr key={user.id} className="hover:bg-teal-50/30 dark:hover:bg-slate-800/40 transition-colors">
                  <td className="px-5 py-3.5">
                    <div className="flex items-center gap-3">
                      <div className="w-9 h-9 bg-teal-50 dark:bg-teal-900/30 border border-teal-200/70 dark:border-teal-800 rounded-xl flex items-center justify-center text-teal-700 dark:text-teal-300">
                        <User className="w-4 h-4" />
                      </div>
                      <div className="flex flex-col">
                        <span className="font-bold text-sm text-slate-900 dark:text-white">{user.email}</span>
                        <span className="text-[11px] text-slate-500 dark:text-slate-400">
                          {user.expiryDate ? `Hết hạn: ${formatDate(user.expiryDate)}` : 'Không giới hạn thời hạn'}
                        </span>
                      </div>
                    </div>
                  </td>
                  <td className="px-5 py-3.5">
                    <span
                      className={cn(
                        'px-2.5 py-1 rounded-md text-[11px] font-bold uppercase tracking-wider border',
                        user.role === 'admin'
                          ? 'bg-emerald-50 text-emerald-800 border-emerald-200 dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-800'
                          : 'bg-teal-50 text-teal-800 border-teal-200 dark:bg-teal-900/30 dark:text-teal-300 dark:border-teal-800'
                      )}
                    >
                      {user.role === 'admin' ? 'Quản trị viên' : 'Người dùng'}
                    </span>
                  </td>
                  <td className="px-5 py-3.5">
                    <div className="text-xs font-semibold text-slate-700 dark:text-slate-200">
                      {user.businessName || 'Mặc định hệ thống'}
                    </div>
                    <div className="text-[11px] font-mono tabular-nums text-slate-500">
                      MST: {user.taxCode || '-'}
                    </div>
                  </td>
                  <td className="px-5 py-3.5 text-right">
                    <ConfirmButton
                      onConfirm={() => handleDeleteUser(user.id)}
                      className="p-2 text-slate-400 hover:text-red-600 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-lg transition-colors inline-flex"
                      icon={Trash2}
                      confirmText="Xóa?"
                    />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function FinancialManagementSection({
  config,
  setConfig,
  incomeData,
  setIncomeData,
  expenseData,
  setExpenseData,
  businessInfo,
  activeSubTab,
  setActiveTab,
  currentUser,
  onNotify,
}: {
  config: FinancialConfig;
  setConfig: (c: FinancialConfig) => void;
  incomeData: IncomeItem[];
  setIncomeData: (d: IncomeItem[]) => void;
  expenseData: ExpenseItem[];
  setExpenseData: (d: ExpenseItem[]) => void;
  businessInfo: BusinessInfo;
  activeSubTab: string;
  setActiveTab: (t: string) => void;
  currentUser: UserAccount | null;
  onNotify?: (msg: string, type?: 'success' | 'error' | 'info') => void;
}) {
  const [isSyncing, setIsSyncing] = useState(false);

  useEffect(() => {
    if (currentUser) {
      const savedData = localStorage.getItem(`finance_data_${currentUser.id}`);
      if (savedData) {
        try {
          const parsed = JSON.parse(savedData);
          if (parsed.incomeData) setIncomeData(parsed.incomeData);
          if (parsed.expenseData) setExpenseData(parsed.expenseData);
          if (parsed.config) setConfig(parsed.config);
        } catch (e) {
          console.error('Error loading saved finance data', e);
        }
      }
    }
  }, [currentUser]);

  const isAdmin = currentUser?.role === 'admin';
  const effectiveBusinessInfo = businessInfo;

  const subNavItems = [
    { id: 'finance-config', label: 'Cấu hình & Tải dữ liệu', icon: Settings },
    { id: 'finance-ledger', label: 'Sổ chi tiết doanh thu', icon: BarChart3 },
    { id: 'finance-vouchers', label: 'Phiếu Thu - Phiếu Chi', icon: Receipt },
  ];

  const totalIncome = useMemo(
    () => incomeData.reduce((sum, item) => sum + (Number(item.amount) || 0), 0),
    [incomeData]
  );
  const totalExpense = useMemo(
    () => expenseData.reduce((sum, item) => sum + (Number(item.amount) || 0), 0),
    [expenseData]
  );
  const netBalance = totalIncome - totalExpense;

  const syncData = async () => {
    if (!currentUser) {
      onNotify?.('Vui lòng đăng nhập để đồng bộ dữ liệu.', 'error');
      return;
    }

    setIsSyncing(true);
    try {
      const dataToSync = {
        incomeData,
        expenseData,
        config,
        updatedAt: new Date().toISOString(),
      };

      localStorage.setItem(`finance_data_${currentUser.id}`, JSON.stringify(dataToSync));
      await new Promise((resolve) => setTimeout(resolve, 600));
      onNotify?.('Đã lưu và đồng bộ dữ liệu tài chính thành công!', 'success');
    } catch (error) {
      console.error('Sync error:', error);
      onNotify?.('Có lỗi xảy ra khi đồng bộ dữ liệu.', 'error');
    } finally {
      setIsSyncing(false);
    }
  };

  const handleIncomeUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!isAdmin) {
      onNotify?.('Chỉ quản trị viên mới có thể nhập dữ liệu doanh thu.', 'error');
      return;
    }
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result;
        const wb = XLSX.read(buffer, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];

        // Fixed indices: A=0 (STT), B=1 (HỌ VÀ TÊN), C=2 (Địa chỉ), AK=36 (Tổng tiền thu)
        const sttIdx = 0;
        const nameIdx = 1;
        const addressIdx = 2;
        const amountIdx = 36;

        let headerRowIdx = -1;
        for (let i = 0; i < data.length; i++) {
          if (String(data[i]?.[nameIdx] || '').toUpperCase().includes('HỌ VÀ TÊN')) {
            headerRowIdx = i;
            break;
          }
        }

        if (headerRowIdx === -1) {
          onNotify?.(
            "Không tìm thấy dòng tiêu đề có 'HỌ VÀ TÊN' ở cột B. Vui lòng dùng đúng file mẫu Bảng chấm công & Thu tiền.",
            'error'
          );
          return;
        }

        const newIncome: IncomeItem[] = data
          .slice(headerRowIdx + 1)
          .filter((row) => row && row[nameIdx] && Number(row[amountIdx]) > 0)
          .map((row, idx) => ({
            id: crypto.randomUUID(),
            stt: String(row[sttIdx] || incomeData.length + idx + 1),
            name: String(row[nameIdx] || '').trim(),
            address: String(row[addressIdx] || '').trim(),
            amount: Number(row[amountIdx] || 0),
            date: config.receiptDate,
          }));

        if (newIncome.length === 0) {
          onNotify?.('Không tìm thấy dòng thu tiền hợp lệ (cần có Họ tên ở cột B và Tổng tiền thu > 0 ở cột AK).', 'error');
          return;
        }

        setIncomeData([...incomeData, ...newIncome]);
        onNotify?.(`Đã tải lên ${newIncome.length} mục chấm công & thu tiền học phí!`, 'success');
      } catch (err) {
        console.error('Income Excel error:', err);
        onNotify?.('Không thể đọc file Excel. Vui lòng kiểm tra lại định dạng.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const handleExpenseUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!isAdmin) {
      onNotify?.('Chỉ quản trị viên mới có thể nhập dữ liệu chi phí.', 'error');
      return;
    }
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result;
        const wb = XLSX.read(buffer, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];

        let headerRowIdx = -1;
        for (let i = 0; i < data.length; i++) {
          if (data[i]?.some((cell) => String(cell).includes('Họ và tên'))) {
            headerRowIdx = i;
            break;
          }
        }

        if (headerRowIdx === -1) {
          onNotify?.("Không tìm thấy cột 'Họ và tên' trong file chi phí.", 'error');
          return;
        }

        const headers = data[headerRowIdx].map((h) => String(h || '').trim());
        const sttIdx = headers.findIndex((h) => h === 'TT' || h === 'STT');
        const nameIdx = headers.findIndex((h) => h === 'Họ và tên');
        const addressIdx = headers.findIndex((h) => h === 'Địa chỉ');
        const contentIdx = headers.findIndex((h) => h === 'Nội dung chi');
        const amountIdx = headers.findIndex((h) => h === 'Số tiền chi');

        if (nameIdx === -1 || amountIdx === -1) {
          onNotify?.("Không tìm thấy cột 'Họ và tên' hoặc 'Số tiền chi'.", 'error');
          return;
        }

        const newExpense: ExpenseItem[] = data
          .slice(headerRowIdx + 1)
          .filter((row) => row && row[nameIdx] && Number(row[amountIdx]) > 0)
          .map((row, idx) => ({
            id: crypto.randomUUID(),
            stt: String(row[sttIdx] || expenseData.length + idx + 1),
            name: String(row[nameIdx] || '').trim(),
            address: String(row[addressIdx] || '').trim(),
            content: String(row[contentIdx] || '').trim(),
            amount: Number(row[amountIdx] || 0),
            date: config.paymentDate,
          }));

        setExpenseData([...expenseData, ...newExpense]);
        onNotify?.(`Đã tải lên ${newExpense.length} khoản chi phí vận hành!`, 'success');
      } catch (err) {
        console.error('Expense Excel error:', err);
        onNotify?.('Có lỗi khi đọc file Excel chi phí.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const downloadIncomeTemplate = () => {
    const header1: any[] = [];
    header1[0] = 'STT';
    header1[1] = 'HỌ VÀ TÊN';
    header1[2] = 'Địa chỉ';
    header1[3] = 'BUỔI HỌC TRONG THÁNG';
    header1[34] = 'Số buổi học';
    header1[35] = 'Số tiền 1 buổi';
    header1[36] = 'Tổng tiền thu';
    header1[37] = 'Ghi chú';

    const header2: any[] = [];
    for (let i = 1; i <= 31; i++) {
      header2[2 + i] = i;
    }
    header2[34] = 'Số buổi học';
    header2[35] = 'Số tiền 1 buổi';
    header2[36] = 'Tổng tiền thu';
    header2[37] = 'Ghi chú';

    const sampleRow1: any[] = [];
    sampleRow1[0] = 1;
    sampleRow1[1] = 'Nguyễn Minh Anh';
    sampleRow1[2] = 'P. Tân Phong - TP. Lai Châu';
    sampleRow1[3] = 'x';
    sampleRow1[6] = 'x';
    sampleRow1[10] = 'x';
    sampleRow1[13] = 'x';
    sampleRow1[17] = 'x';
    sampleRow1[20] = 'x';
    sampleRow1[24] = 'x';
    sampleRow1[27] = 'x';
    sampleRow1[34] = 8;
    sampleRow1[35] = 70000;
    sampleRow1[36] = 560000;
    sampleRow1[37] = 'Đã thu đủ';

    const ws = XLSX.utils.aoa_to_sheet([header1, header2, sampleRow1]);

    ws['!merges'] = [
      { s: { r: 0, c: 3 }, e: { r: 0, c: 33 } },
      { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },
      { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } },
      { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } },
      { s: { r: 0, c: 34 }, e: { r: 1, c: 34 } },
      { s: { r: 0, c: 35 }, e: { r: 1, c: 35 } },
      { s: { r: 0, c: 36 }, e: { r: 1, c: 36 } },
      { s: { r: 0, c: 37 }, e: { r: 1, c: 37 } },
    ];

    const cols: { wch: number }[] = [
      { wch: 6 },
      { wch: 24 },
      { wch: 26 },
    ];
    for (let i = 1; i <= 31; i++) {
      cols.push({ wch: 4 });
    }
    cols.push({ wch: 12 }, { wch: 15 }, { wch: 16 }, { wch: 18 });
    ws['!cols'] = cols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Cham_Cong_Thu_Tien');
    XLSX.writeFile(wb, 'Mau_Cham_Cong_Thu_Tien.xlsx');
    onNotify?.('Đã tải file mẫu Excel Bảng chấm công và Thu tiền!', 'success');
  };

  const downloadExpenseTemplate = () => {
    const rows = [
      ['TT', 'Họ và tên', 'Địa chỉ', 'Nội dung chi', 'Số tiền chi'],
      [1, 'Điện lực Lai Châu', 'TP. Lai Châu', 'Tiền điện phòng học tháng hiện tại', 850000],
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 6 }, { wch: 24 }, { wch: 24 }, { wch: 32 }, { wch: 16 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Mau_Chi_Tien');
    XLSX.writeFile(wb, 'Mau_Noi_Dung_Chi.xlsx');
    onNotify?.('Đã tải file mẫu Excel Nội dung chi!', 'success');
  };

  const exportS1aHKD = async () => {
    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 1134, bottom: 567, left: 1531, right: 1134 }
          }
        },
        children: [
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
              insideHorizontal: { style: BorderStyle.NONE },
              insideVertical: { style: BorderStyle.NONE },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 60, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({ text: sanitizeDocxText("HỘ, CÁ NHÂN KINH DOANH: "), bold: true, size: 24 }),
                          new TextRun({ text: sanitizeDocxText(effectiveBusinessInfo.name?.toUpperCase() || ""), bold: true, size: 24 }),
                        ],
                      }),
                      new Paragraph({
                        children: [
                          new TextRun({ text: sanitizeDocxText("Địa chỉ: "), bold: true, size: 24 }),
                          new TextRun({ text: sanitizeDocxText(effectiveBusinessInfo.address || ""), size: 24 }),
                        ],
                      }),
                      new Paragraph({
                        children: [
                          new TextRun({ text: sanitizeDocxText("Mã số thuế: "), bold: true, size: 24 }),
                          new TextRun({ text: sanitizeDocxText(effectiveBusinessInfo.taxId) || "................", size: 24 }),
                        ],
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 40, type: WidthType.PERCENTAGE },
                    children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({ text: sanitizeDocxText("Mẫu số S1a-HKD"), bold: true, size: 22 }),
                            ],
                          }),
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({ text: sanitizeDocxText("(Kèm theo Thông tư số 152/2025/TT-BTC"), italics: true, size: 20 }),
                            ],
                          }),
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({ text: sanitizeDocxText("ngày 31 tháng 12 năm 2025 của Bộ trưởng Bộ Tài chính)"), italics: true, size: 20 }),
                            ],
                          }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 600, after: 300 },
            children: [
              new TextRun({ text: sanitizeDocxText("SỔ CHI TIẾT DOANH THU BÁN HÀNG HÓA, DỊCH VỤ"), bold: true, size: 36 }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: sanitizeDocxText(`Địa điểm kinh doanh: ${effectiveBusinessInfo.businessLocation || effectiveBusinessInfo.address}`), size: 26 }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
            children: [
              new TextRun({ text: sanitizeDocxText(`Kỳ kê khai: ${config.reportPeriod}`), size: 26 }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: { after: 200 },
            children: [
              new TextRun({ text: sanitizeDocxText("Đơn vị tính: VNĐ"), italics: true, size: 24 }),
            ],
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText("Ngày tháng"), bold: true, size: 26 })] })], verticalAlign: VerticalAlign.CENTER }),
                  new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText("Giao dịch"), bold: true, size: 26 })] })], verticalAlign: VerticalAlign.CENTER }),
                  new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText("Số tiền"), bold: true, size: 26 })] })], verticalAlign: VerticalAlign.CENTER }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText("A"), italics: true, size: 24 })] })] }),
                  new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText("B"), italics: true, size: 24 })] })] }),
                  new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText("1"), italics: true, size: 24 })] })] }),
                ],
              }),
              ...incomeData.map(item => new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText(formatDate(item.date) || getLastDayOfMonth(config.reportPeriod)), size: 26 })] })] }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: sanitizeDocxText(`Thu tiền học phí - ${item.name} ${item.address}`), size: 26 })] })] }),
                  new TableCell({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: sanitizeDocxText(item.amount.toLocaleString()), size: 26 })] })] }),
                ],
              })),
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: sanitizeDocxText("Tổng cộng"), bold: true, size: 26 })] })], columnSpan: 2 }),
                  new TableCell({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: sanitizeDocxText(incomeData.reduce((sum, item) => sum + item.amount, 0).toLocaleString()), bold: true, size: 26 })] })] }),
                ],
              }),
            ],
          }),
          new Paragraph({ spacing: { before: 800 } }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
              insideHorizontal: { style: BorderStyle.NONE },
              insideVertical: { style: BorderStyle.NONE },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({ children: [], width: { size: 50, type: WidthType.PERCENTAGE } }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({ text: sanitizeDocxText("Ngày ...... tháng ...... năm ......"), italics: true, size: 24 }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({ text: sanitizeDocxText("NGƯỜI ĐẠI DIỆN HỘ KINH DOANH/"), bold: true, size: 24 }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({ text: sanitizeDocxText("CÁ NHÂN KINH DOANH"), bold: true, size: 24 }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({ text: sanitizeDocxText("(Ký, ghi rõ họ tên, đóng dấu(nếu có))"), italics: true, size: 20 }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 1200 },
                        children: [
                          new TextRun({ text: sanitizeDocxText(effectiveBusinessInfo.owner || "................"), bold: true, size: 24 }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `SoDoanhThu_${config.reportPeriod.replace(/\//g, '-')}.docx`);
  };

  const getReceiptVoucherChildren = (item: IncomeItem, index: number = 0) => {
    const dateParts = (config.receiptDate || "").split('-');
    const d = dateParts[2] || '...';
    const m = dateParts[1] || '...';
    const y = dateParts[0] || '...';
    const voucherNo = `PT${d}${m}${y.slice(-2)}-${String(index + 1).padStart(3, '0')}`;

    const noBorder = {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    };

    return [
      // Header Table
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: noBorder,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 60, type: WidthType.PERCENTAGE },
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({ text: "HỘ, CÁ NHÂN KINH DOANH: ", bold: true, size: 20 }),
                      new TextRun({ text: effectiveBusinessInfo.name || "................", size: 20 }),
                    ],
                  }),
                  new Paragraph({
                    children: [
                      new TextRun({ text: "Địa chỉ: ", bold: true, size: 20 }),
                      new TextRun({ text: effectiveBusinessInfo.address || "................", size: 20 }),
                    ],
                  }),
                  new Paragraph({
                    children: [
                      new TextRun({ text: "Mã số thuế: ", bold: true, size: 20 }),
                      new TextRun({ text: effectiveBusinessInfo.taxId || "................", size: 20 }),
                    ],
                  }),
                ],
              }),
              new TableCell({
                width: { size: 40, type: WidthType.PERCENTAGE },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({ text: "Mẫu số 01 – TT", bold: true, size: 20 }),
                    ],
                  }),
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({ 
                        text: "(Ban hành kèm theo Thông tư số 88/2021/TT-BTC ngày 11 tháng 10 năm 2021 của Bộ trưởng Bộ Tài chính)", 
                        italics: true, 
                        size: 16 
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      }),

      // Title Section
      new Paragraph({ spacing: { before: 200 } }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: noBorder,
        rows: [
          new TableRow({
            children: [
              new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, children: [] }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({ text: sanitizeDocxText("PHIẾU THU"), bold: true, size: 32 }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({ text: sanitizeDocxText(`Ngày ${d} tháng ${m} năm ${y}`), italics: true, size: 20 }),
                        ],
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 25, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({ text: sanitizeDocxText("Quyển số: "), size: 20 }),
                          new TextRun({ text: sanitizeDocxText("................"), size: 20 }),
                        ],
                      }),
                      new Paragraph({
                        children: [
                          new TextRun({ text: sanitizeDocxText("Số: "), size: 20 }),
                          new TextRun({ text: sanitizeDocxText(voucherNo), size: 20 }),
                        ],
                      }),
                    ],
                  }),
            ],
          }),
        ],
      }),

      new Paragraph({ spacing: { before: 200 } }),
      new Paragraph({
        children: [
          new TextRun({ text: sanitizeDocxText("Họ và tên người nộp tiền: "), size: 26 }),
          new TextRun({ text: sanitizeDocxText(item.name || "................................................................"), bold: true, size: 26 }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({ text: sanitizeDocxText("Địa chỉ: "), size: 26 }),
          new TextRun({ text: sanitizeDocxText(item.address || "................................................................"), size: 26 }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({ text: sanitizeDocxText("Lý do nộp: "), size: 26 }),
          new TextRun({ text: sanitizeDocxText(`Nộp tiền học phí ${config.reportPeriod}`), size: 26 }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({ text: sanitizeDocxText("Số tiền: "), size: 26 }),
          new TextRun({ text: sanitizeDocxText(`${item.amount.toLocaleString()} VNĐ`), bold: true, size: 26 }),
          new TextRun({ text: sanitizeDocxText(" (Viết bằng chữ): "), size: 26 }),
          new TextRun({ text: sanitizeDocxText(numberToVietnameseWords(item.amount)), italics: true, size: 26 }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({ text: sanitizeDocxText("Kèm theo: "), size: 26 }),
          new TextRun({ text: sanitizeDocxText(`Bảng chấm công và thu tiền ${config.reportPeriod}`), size: 26 }),
          new TextRun({ text: sanitizeDocxText(" Chứng từ gốc: "), size: 26 }),
          new TextRun({ text: sanitizeDocxText("................"), size: 26 }),
        ],
      }),

      // Signatures
      new Paragraph({ spacing: { before: 200 } }),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [
          new TextRun({ text: `Ngày ${d} tháng ${m} năm ${y}`, italics: true, size: 18 }),
        ],
      }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: noBorder,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NGƯỜI ĐẠI DIỆN", bold: true, size: 22 })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "HỘ KINH DOANH/", bold: true, size: 22 })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "CÁ NHÂN KINH DOANH", bold: true, size: 22 })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "(Ký, họ tên, đóng dấu)", italics: true, size: 22 })] }),
                ],
              }),
              new TableCell({
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NGƯỜI LẬP BIỂU", bold: true, size: 22 })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "(Ký, họ tên)", italics: true, size: 22 })] }),
                ],
              }),
              new TableCell({
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NGƯỜI NỘP TIỀN", bold: true, size: 22 })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "(Ký, họ tên)", italics: true, size: 22 })] }),
                ],
              }),
              new TableCell({
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "THỦ QUỸ", bold: true, size: 22 })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "(Ký, họ tên)", italics: true, size: 22 })] }),
                ],
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({ spacing: { before: 400 } }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: effectiveBusinessInfo.owner || "", bold: true, size: 22 })] }),
                ],
              }),
              new TableCell({
                children: [
                  new Paragraph({ spacing: { before: 400 } }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: config.preparer || "", bold: true, size: 22 })] }),
                ],
              }),
              new TableCell({
                children: [
                  new Paragraph({ spacing: { before: 400 } }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.name, bold: true, size: 22 })] }),
                ],
              }),
              new TableCell({
                children: [
                  new Paragraph({ spacing: { before: 400 } }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: config.treasurer || "", bold: true, size: 22 })] }),
                ],
              }),
            ],
          }),
        ],
      }),
      new Paragraph({ spacing: { before: 200 } }),
      new Paragraph({
        children: [
          new TextRun({ text: sanitizeDocxText("Đã nhận đủ số tiền (viết bằng chữ): "), size: 20 }),
          new TextRun({ text: sanitizeDocxText("...................................................................................................."), size: 20 }),
        ],
      }),
    ];
  };

  const getPaymentVoucherChildren = (item: ExpenseItem, index: number = 0) => {
    const dateParts = (config.paymentDate || "").split('-');
    const d = dateParts[2] || '...';
    const m = dateParts[1] || '...';
    const y = dateParts[0] || '...';
    const voucherNo = `PC${d}${m}${y.slice(-2)}-${String(index + 1).padStart(3, '0')}`;

    const noBorder = {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    };

    return [
      // Header Table
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: noBorder,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 60, type: WidthType.PERCENTAGE },
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({ text: "HỘ, CÁ NHÂN KINH DOANH: ", bold: true, size: 20 }),
                      new TextRun({ text: effectiveBusinessInfo.name || "................", size: 20 }),
                    ],
                  }),
                  new Paragraph({
                    children: [
                      new TextRun({ text: "Địa chỉ: ", bold: true, size: 20 }),
                      new TextRun({ text: effectiveBusinessInfo.address || "................", size: 20 }),
                    ],
                  }),
                  new Paragraph({
                    children: [
                      new TextRun({ text: "Mã số thuế: ", bold: true, size: 20 }),
                      new TextRun({ text: effectiveBusinessInfo.taxId || "................", size: 20 }),
                    ],
                  }),
                ],
              }),
              new TableCell({
                width: { size: 40, type: WidthType.PERCENTAGE },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({ text: "Mẫu số 01 – TT", bold: true, size: 20 }),
                    ],
                  }),
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({ 
                        text: "(Ban hành kèm theo Thông tư số 88/2021/TT-BTC ngày 11 tháng 10 năm 2021 của Bộ trưởng Bộ Tài chính)", 
                        italics: true, 
                        size: 16 
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      }),

      // Title Section
      new Paragraph({ spacing: { before: 200 } }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: noBorder,
        rows: [
          new TableRow({
            children: [
              new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, children: [] }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({ text: sanitizeDocxText("PHIẾU CHI"), bold: true, size: 32 }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({ text: sanitizeDocxText(`Ngày ${d} tháng ${m} năm ${y}`), italics: true, size: 20 }),
                        ],
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 25, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({ text: sanitizeDocxText("Quyển số: "), size: 20 }),
                          new TextRun({ text: sanitizeDocxText("................"), size: 20 }),
                        ],
                      }),
                      new Paragraph({
                        children: [
                          new TextRun({ text: sanitizeDocxText("Số: "), size: 20 }),
                          new TextRun({ text: sanitizeDocxText(voucherNo), size: 20 }),
                        ],
                      }),
                    ],
                  }),
            ],
          }),
        ],
      }),

      new Paragraph({ spacing: { before: 200 } }),
      new Paragraph({
        children: [
          new TextRun({ text: sanitizeDocxText("Họ và tên người nhận tiền: "), size: 26 }),
          new TextRun({ text: sanitizeDocxText(item.name || "................................................................"), bold: true, size: 26 }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({ text: sanitizeDocxText("Địa chỉ: "), size: 26 }),
          new TextRun({ text: sanitizeDocxText(item.address || "................................................................"), size: 26 }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({ text: sanitizeDocxText("Lý do chi: "), size: 26 }),
          new TextRun({ text: sanitizeDocxText(`Chi tiền học phí ${config.reportPeriod}`), size: 26 }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({ text: sanitizeDocxText("Số tiền: "), size: 26 }),
          new TextRun({ text: sanitizeDocxText(`${item.amount.toLocaleString()} VNĐ`), bold: true, size: 26 }),
          new TextRun({ text: sanitizeDocxText(" (Viết bằng chữ): "), size: 26 }),
          new TextRun({ text: sanitizeDocxText(numberToVietnameseWords(item.amount)), italics: true, size: 26 }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun({ text: sanitizeDocxText("Kèm theo: "), size: 26 }),
          new TextRun({ text: sanitizeDocxText(`Bảng chấm công và thu tiền ${config.reportPeriod}`), size: 26 }),
          new TextRun({ text: sanitizeDocxText(" Chứng từ gốc: "), size: 26 }),
          new TextRun({ text: sanitizeDocxText("................"), size: 26 }),
        ],
      }),

      // Signatures
      new Paragraph({ spacing: { before: 200 } }),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [
          new TextRun({ text: `Ngày ${d} tháng ${m} năm ${y}`, italics: true, size: 18 }),
        ],
      }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: noBorder,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NGƯỜI ĐẠI DIỆN", bold: true, size: 22 })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "HỘ KINH DOANH/", bold: true, size: 22 })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "CÁ NHÂN KINH DOANH", bold: true, size: 22 })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "(Ký, họ tên, đóng dấu)", italics: true, size: 22 })] }),
                ],
              }),
              new TableCell({
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NGƯỜI LẬP BIỂU", bold: true, size: 22 })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "(Ký, họ tên)", italics: true, size: 22 })] }),
                ],
              }),
              new TableCell({
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NGƯỜI NHẬN TIỀN", bold: true, size: 22 })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "(Ký, họ tên)", italics: true, size: 22 })] }),
                ],
              }),
              new TableCell({
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "THỦ QUỸ", bold: true, size: 22 })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "(Ký, họ tên)", italics: true, size: 22 })] }),
                ],
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({ spacing: { before: 400 } }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: effectiveBusinessInfo.owner || "", bold: true, size: 22 })] }),
                ],
              }),
              new TableCell({
                children: [
                  new Paragraph({ spacing: { before: 400 } }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: config.preparer || "", bold: true, size: 22 })] }),
                ],
              }),
              new TableCell({
                children: [
                  new Paragraph({ spacing: { before: 400 } }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.name, bold: true, size: 22 })] }),
                ],
              }),
              new TableCell({
                children: [
                  new Paragraph({ spacing: { before: 400 } }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: config.treasurer || "", bold: true, size: 22 })] }),
                ],
              }),
            ],
          }),
        ],
      }),
    ];
  };

  const exportReceipt = async (item: IncomeItem) => {
    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 1134, bottom: 567, left: 1134, right: 1418 }
          }
        },
        children: getReceiptVoucherChildren(item),
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `PhieuThu_${item.name}.docx`);
  };

  const exportPaymentVoucher = async (item: ExpenseItem) => {
    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 1134, bottom: 567, left: 1134, right: 1418 }
          }
        },
        children: getPaymentVoucherChildren(item),
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `PhieuChi_${item.name}.docx`);
  };

  const exportAllVouchers = async () => {
    const children: any[] = [];
    
    // Combine income and expense items
    const allItems = [
      ...incomeData.map(item => ({ type: 'receipt' as const, data: item })),
      ...expenseData.map(item => ({ type: 'payment' as const, data: item }))
    ];

    let receiptCount = 0;
    let paymentCount = 0;

    allItems.forEach((item, index) => {
      if (item.type === 'receipt') {
        children.push(...getReceiptVoucherChildren(item.data as IncomeItem, receiptCount));
        receiptCount++;
      } else {
        children.push(...getPaymentVoucherChildren(item.data as ExpenseItem, paymentCount));
        paymentCount++;
      }

      // Add separator or page break
      if (index < allItems.length - 1) {
        if (index % 2 === 0) {
          // Even index (0, 2, 4...) means it's the first voucher on the page
          children.push(new Paragraph({ 
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 200 },
            children: [new TextRun({ text: "----------------------------------------------------------------------------------------------------", size: 20 })] 
          }));
        } else {
          // Odd index (1, 3, 5...) means it's the second voucher on the page
          children.push(new Paragraph({ 
            children: [new PageBreak()],
            spacing: { before: 0, after: 0 }
          }));
        }
      }
    });

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 1134, bottom: 567, left: 1134, right: 1418 }
          }
        },
        children: children,
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `Tat_Ca_Phieu_Thu_Chi.docx`);
  };

  return (
    <div className="space-y-6">
      <SectionBanner
        title="Quản lý Tài chính"
        subtitle="Quản lý thu học phí, chi phí vận hành, xuất file mẫu Excel bảng chấm công & thu tiền, sổ doanh thu S1a-HKD và phiếu thu chi."
        badge="TÀI CHÍNH - KẾ TOÁN"
        icon={DollarSign}
        onBack={() => setActiveTab('dashboard')}
        stats={[
          {
            label: 'Tổng thu học phí',
            value: `${totalIncome.toLocaleString('vi-VN')} đ`,
            sub: `${incomeData.length} khoản thu đã nhập`,
            highlight: true,
          },
          {
            label: 'Tổng chi vận hành',
            value: `${totalExpense.toLocaleString('vi-VN')} đ`,
            sub: `${expenseData.length} khoản chi đã nhập`,
          },
          {
            label: 'Chênh lệch Thu - Chi',
            value: `${netBalance.toLocaleString('vi-VN')} đ`,
            sub: netBalance >= 0 ? 'Dòng tiền dương' : 'Thâm hụt kỳ này',
          },
          {
            label: 'Kỳ kê khai',
            value: config.reportPeriod || 'Chưa đặt kỳ',
            sub: `Người lập: ${config.preparer || 'Chưa nhập'}`,
          },
        ]}
        tabs={subNavItems}
        activeTab={activeSubTab}
        onTabChange={setActiveTab}
        actions={
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={downloadIncomeTemplate}
              className="inline-flex items-center gap-2 px-3.5 py-2.5 rounded-xl bg-white text-teal-900 hover:bg-teal-50 font-bold text-xs shadow-sm transition-colors"
            >
              <FileSpreadsheet className="w-4 h-4 text-teal-700" />
              Xuất file mẫu Chấm công & Thu tiền (.xlsx)
            </button>
            <button
              onClick={() => setActiveTab('tax_report')}
              className="inline-flex items-center gap-2 px-3.5 py-2.5 rounded-xl bg-white/15 hover:bg-white/25 text-white border border-white/25 font-bold text-xs transition-colors"
            >
              <Calculator className="w-4 h-4" />
              Báo cáo thuế 12 tháng
            </button>
          </div>
        }
      />

      {activeSubTab === 'finance_group' && (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
          <button
            onClick={() => setActiveTab('finance-config')}
            className="bg-white dark:bg-slate-900 p-6 rounded-2xl border border-slate-200 dark:border-slate-800 hover:border-teal-500 shadow-sm hover:shadow-md transition-all text-left group flex flex-col justify-between"
          >
            <div>
              <div className="w-12 h-12 bg-teal-50 dark:bg-teal-900/30 border border-teal-200/70 dark:border-teal-800 rounded-xl flex items-center justify-center text-teal-700 dark:text-teal-300 mb-4 group-hover:bg-teal-600 group-hover:text-white transition-colors">
                <Settings className="w-6 h-6" />
              </div>
              <h3 className="text-base font-bold text-slate-900 dark:text-white">Cấu hình & Tải dữ liệu</h3>
              <p className="text-xs text-slate-500 dark:text-slate-400 mt-1.5 leading-relaxed">
                Thiết lập kỳ báo cáo, người ký và tải dữ liệu bảng chấm công thu tiền / chi phí từ Excel.
              </p>
            </div>
            <div className="mt-5 pt-3 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between text-xs font-bold text-teal-700 dark:text-teal-400">
              <span>Mở cấu hình</span>
              <ArrowRight className="w-4 h-4" />
            </div>
          </button>

          <button
            onClick={downloadIncomeTemplate}
            className="bg-white dark:bg-slate-900 p-6 rounded-2xl border border-slate-200 dark:border-slate-800 hover:border-emerald-500 shadow-sm hover:shadow-md transition-all text-left group flex flex-col justify-between"
          >
            <div>
              <div className="w-12 h-12 bg-emerald-50 dark:bg-emerald-900/30 border border-emerald-200/70 dark:border-emerald-800 rounded-xl flex items-center justify-center text-emerald-700 dark:text-emerald-300 mb-4 group-hover:bg-emerald-600 group-hover:text-white transition-colors">
                <FileSpreadsheet className="w-6 h-6" />
              </div>
              <h3 className="text-base font-bold text-slate-900 dark:text-white">Xuất file mẫu Chấm công & Thu tiền</h3>
              <p className="text-xs text-slate-500 dark:text-slate-400 mt-1.5 leading-relaxed">
                Tải file mẫu Excel 31 ngày trong tháng kèm cột số buổi học, đơn giá và tổng tiền thu chuẩn.
              </p>
            </div>
            <div className="mt-5 pt-3 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between text-xs font-bold text-emerald-700 dark:text-emerald-400">
              <span>Tải mẫu Excel (.xlsx)</span>
              <Download className="w-4 h-4" />
            </div>
          </button>

          <button
            onClick={() => setActiveTab('finance-ledger')}
            className="bg-white dark:bg-slate-900 p-6 rounded-2xl border border-slate-200 dark:border-slate-800 hover:border-teal-500 shadow-sm hover:shadow-md transition-all text-left group flex flex-col justify-between"
          >
            <div>
              <div className="w-12 h-12 bg-teal-50 dark:bg-teal-900/30 border border-teal-200/70 dark:border-teal-800 rounded-xl flex items-center justify-center text-teal-700 dark:text-teal-300 mb-4 group-hover:bg-teal-600 group-hover:text-white transition-colors">
                <BarChart3 className="w-6 h-6" />
              </div>
              <h3 className="text-base font-bold text-slate-900 dark:text-white">Sổ chi tiết doanh thu</h3>
              <p className="text-xs text-slate-500 dark:text-slate-400 mt-1.5 leading-relaxed">
                Xuất Sổ doanh thu mẫu S1a-HKD theo Thông tư 152/2025/TT-BTC phục vụ kê khai quyết toán.
              </p>
            </div>
            <div className="mt-5 pt-3 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between text-xs font-bold text-teal-700 dark:text-teal-400">
              <span>Xuất sổ S1a-HKD</span>
              <ArrowRight className="w-4 h-4" />
            </div>
          </button>

          <button
            onClick={() => setActiveTab('finance-vouchers')}
            className="bg-white dark:bg-slate-900 p-6 rounded-2xl border border-slate-200 dark:border-slate-800 hover:border-teal-500 shadow-sm hover:shadow-md transition-all text-left group flex flex-col justify-between"
          >
            <div>
              <div className="w-12 h-12 bg-teal-50 dark:bg-teal-900/30 border border-teal-200/70 dark:border-teal-800 rounded-xl flex items-center justify-center text-teal-700 dark:text-teal-300 mb-4 group-hover:bg-teal-600 group-hover:text-white transition-colors">
                <Receipt className="w-6 h-6" />
              </div>
              <h3 className="text-base font-bold text-slate-900 dark:text-white">Phiếu Thu - Phiếu Chi</h3>
              <p className="text-xs text-slate-500 dark:text-slate-400 mt-1.5 leading-relaxed">
                Tự động lập và kết xuất hàng loạt phiếu thu, phiếu chi tiền mặt (2 phiếu/trang A4).
              </p>
            </div>
            <div className="mt-5 pt-3 border-t border-slate-100 dark:border-slate-800 flex items-center justify-between text-xs font-bold text-teal-700 dark:text-teal-400">
              <span>Xuất chứng từ</span>
              <ArrowRight className="w-4 h-4" />
            </div>
          </button>
        </div>
      )}

      {activeSubTab === 'finance-config' && (
        <div className="space-y-6">
          <div className="bg-white dark:bg-slate-900 p-6 sm:p-8 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm">
            <h3 className="text-base font-bold text-slate-900 dark:text-white mb-5 flex items-center gap-2.5">
              <div className="w-9 h-9 bg-teal-50 dark:bg-teal-900/30 border border-teal-200/70 dark:border-teal-800 rounded-xl flex items-center justify-center text-teal-700 dark:text-teal-300">
                <Settings className="w-4 h-4" />
              </div>
              Cấu hình Kỳ Báo cáo & Người ký Chứng từ
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-700 dark:text-slate-300">Kỳ báo cáo (VD: Tháng 03/2026)</label>
                <input
                  type="text"
                  value={config.reportPeriod}
                  onChange={(e) => isAdmin && setConfig({ ...config, reportPeriod: e.target.value })}
                  disabled={!isAdmin}
                  className="input-field disabled:opacity-70"
                  placeholder="Tháng 03/2026"
                />
              </div>
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-700 dark:text-slate-300">Ngày xuất phiếu thu</label>
                <input
                  type="date"
                  value={config.receiptDate}
                  onChange={(e) => isAdmin && setConfig({ ...config, receiptDate: e.target.value })}
                  disabled={!isAdmin}
                  className="input-field disabled:opacity-70"
                />
              </div>
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-700 dark:text-slate-300">Ngày xuất phiếu chi</label>
                <input
                  type="date"
                  value={config.paymentDate}
                  onChange={(e) => isAdmin && setConfig({ ...config, paymentDate: e.target.value })}
                  disabled={!isAdmin}
                  className="input-field disabled:opacity-70"
                />
              </div>
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-700 dark:text-slate-300">Người lập biểu</label>
                <input
                  type="text"
                  value={config.preparer}
                  onChange={(e) => isAdmin && setConfig({ ...config, preparer: e.target.value })}
                  disabled={!isAdmin}
                  className="input-field disabled:opacity-70"
                  placeholder="Họ và tên người lập biểu"
                />
              </div>
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-700 dark:text-slate-300">Thủ quỹ</label>
                <input
                  type="text"
                  value={config.treasurer}
                  onChange={(e) => isAdmin && setConfig({ ...config, treasurer: e.target.value })}
                  disabled={!isAdmin}
                  className="input-field disabled:opacity-70"
                  placeholder="Họ và tên thủ quỹ"
                />
              </div>
              {isAdmin && (
                <div className="flex items-end">
                  <button
                    onClick={() => {
                      if (!config.reportPeriod) {
                        onNotify?.('Vui lòng nhập Kỳ báo cáo (VD: Tháng 03/2026).', 'error');
                        return;
                      }
                      onNotify?.('Đã lưu cấu hình kỳ báo cáo tài chính!', 'success');
                    }}
                    className="btn-primary w-full flex items-center justify-center gap-2 text-xs py-3"
                  >
                    <Save className="w-4 h-4" /> Lưu cấu hình kỳ báo cáo
                  </button>
                </div>
              )}
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* Nội dung thu */}
            <div className="bg-white dark:bg-slate-900 p-6 sm:p-7 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm flex flex-col justify-between gap-5">
              <div>
                <div className="flex flex-wrap items-start justify-between gap-3 mb-3">
                  <h3 className="text-base font-bold text-slate-900 dark:text-white flex items-center gap-2.5">
                    <div className="w-9 h-9 bg-emerald-50 dark:bg-emerald-900/30 border border-emerald-200/70 dark:border-emerald-800 rounded-xl flex items-center justify-center text-emerald-700 dark:text-emerald-300">
                      <TrendingUp className="w-4 h-4" />
                    </div>
                    Bảng Chấm công & Thu tiền Học phí
                  </h3>
                  <button
                    onClick={downloadIncomeTemplate}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-emerald-50 dark:bg-emerald-900/30 text-emerald-800 dark:text-emerald-200 border border-emerald-200 dark:border-emerald-800 hover:bg-emerald-100 rounded-xl font-bold text-xs transition-colors"
                  >
                    <FileSpreadsheet className="w-4 h-4" />
                    Xuất file mẫu Chấm công & Thu tiền (.xlsx)
                  </button>
                </div>
                <p className="text-xs text-slate-500 dark:text-slate-400 leading-relaxed">
                  Tải file mẫu Excel chấm công 31 ngày & thu tiền, điền dữ liệu và đưa lên hệ thống để tự động tổng hợp doanh thu.
                </p>
              </div>

              <div className="space-y-3">
                {isAdmin && (
                  <div className="flex flex-col sm:flex-row gap-3">
                    <label className="flex-1 flex items-center justify-center gap-2.5 px-5 py-3.5 border-2 border-dashed border-emerald-300 dark:border-emerald-800 bg-emerald-50/40 dark:bg-emerald-900/10 rounded-xl hover:border-emerald-600 hover:bg-emerald-50 transition-colors cursor-pointer">
                      <Upload className="w-4 h-4 text-emerald-700 dark:text-emerald-400" />
                      <span className="text-xs font-bold text-emerald-900 dark:text-emerald-200">
                        Đưa bảng Chấm công & Thu tiền lên (Excel)
                      </span>
                      <input type="file" accept=".xlsx, .xls" onChange={handleIncomeUpload} className="hidden" />
                    </label>
                    {incomeData.length > 0 && (
                      <ConfirmButton
                        onConfirm={() => {
                          setIncomeData([]);
                          onNotify?.('Đã xóa dữ liệu thu học phí.', 'info');
                        }}
                        className="px-4 py-3 text-red-600 border border-red-200 dark:border-red-900/40 rounded-xl hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors font-bold text-xs"
                        icon={Trash2}
                      >
                        Xóa thu
                      </ConfirmButton>
                    )}
                  </div>
                )}

                {incomeData.length > 0 && (
                  <div className="p-3.5 bg-emerald-50 dark:bg-emerald-900/20 border border-emerald-200/70 dark:border-emerald-800 text-emerald-800 dark:text-emerald-300 rounded-xl text-xs font-bold flex items-center justify-between">
                    <span className="flex items-center gap-2">
                      <CheckCircle2 className="w-4 h-4" />
                      Đã tải {incomeData.length} học sinh nộp học phí
                    </span>
                    <span className="font-mono tabular-nums">{totalIncome.toLocaleString('vi-VN')} đ</span>
                  </div>
                )}
              </div>
            </div>

            {/* Nội dung chi */}
            <div className="bg-white dark:bg-slate-900 p-6 sm:p-7 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm flex flex-col justify-between gap-5">
              <div>
                <div className="flex flex-wrap items-start justify-between gap-3 mb-3">
                  <h3 className="text-base font-bold text-slate-900 dark:text-white flex items-center gap-2.5">
                    <div className="w-9 h-9 bg-teal-50 dark:bg-teal-900/30 border border-teal-200/70 dark:border-teal-800 rounded-xl flex items-center justify-center text-teal-700 dark:text-teal-300">
                      <CreditCard className="w-4 h-4" />
                    </div>
                    Bảng Nội dung Chi phí Vận hành
                  </h3>
                  <button
                    onClick={downloadExpenseTemplate}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-teal-50 dark:bg-teal-900/30 text-teal-800 dark:text-teal-200 border border-teal-200 dark:border-teal-800 hover:bg-teal-100 rounded-xl font-bold text-xs transition-colors"
                  >
                    <Download className="w-4 h-4" />
                    Xuất file mẫu Chi phí (.xlsx)
                  </button>
                </div>
                <p className="text-xs text-slate-500 dark:text-slate-400 leading-relaxed">
                  Tải lên danh sách các khoản chi phí điện nước, văn phòng phẩm, thù lao... để lập phiếu chi.
                </p>
              </div>

              <div className="space-y-3">
                {isAdmin && (
                  <div className="flex flex-col sm:flex-row gap-3">
                    <label className="flex-1 flex items-center justify-center gap-2.5 px-5 py-3.5 border-2 border-dashed border-teal-300 dark:border-teal-800 bg-teal-50/40 dark:bg-teal-900/10 rounded-xl hover:border-teal-600 hover:bg-teal-50 transition-colors cursor-pointer">
                      <Upload className="w-4 h-4 text-teal-700 dark:text-teal-400" />
                      <span className="text-xs font-bold text-teal-900 dark:text-teal-200">
                        Đưa danh sách Chi phí lên (Excel)
                      </span>
                      <input type="file" accept=".xlsx, .xls" onChange={handleExpenseUpload} className="hidden" />
                    </label>
                    {expenseData.length > 0 && (
                      <ConfirmButton
                        onConfirm={() => {
                          setExpenseData([]);
                          onNotify?.('Đã xóa dữ liệu chi phí.', 'info');
                        }}
                        className="px-4 py-3 text-red-600 border border-red-200 dark:border-red-900/40 rounded-xl hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors font-bold text-xs"
                        icon={Trash2}
                      >
                        Xóa chi
                      </ConfirmButton>
                    )}
                  </div>
                )}

                {expenseData.length > 0 && (
                  <div className="p-3.5 bg-teal-50 dark:bg-teal-900/20 border border-teal-200/70 dark:border-teal-800 text-teal-800 dark:text-teal-300 rounded-xl text-xs font-bold flex items-center justify-between">
                    <span className="flex items-center gap-2">
                      <CheckCircle2 className="w-4 h-4" />
                      Đã tải {expenseData.length} khoản chi phí
                    </span>
                    <span className="font-mono tabular-nums">{totalExpense.toLocaleString('vi-VN')} đ</span>
                  </div>
                )}
              </div>
            </div>
          </div>

          <div className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 overflow-hidden shadow-sm">
            <div className="p-5 border-b border-slate-200 dark:border-slate-800 flex flex-col sm:flex-row justify-between sm:items-center gap-3 bg-slate-50/70 dark:bg-slate-800/40">
              <div>
                <h3 className="font-bold text-slate-900 dark:text-white text-base">Bảng Tổng hợp Thu - Chi Chi tiết</h3>
                <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
                  Tổng số: {incomeData.length + expenseData.length} giao dịch • Tổng thu:{' '}
                  <strong className="text-emerald-700 dark:text-emerald-400 font-mono tabular-nums">
                    {totalIncome.toLocaleString('vi-VN')} đ
                  </strong>
                </p>
              </div>
              <button
                onClick={syncData}
                disabled={isSyncing}
                className="btn-primary flex items-center gap-2 text-xs disabled:opacity-50"
              >
                <RefreshCw className={cn('w-4 h-4', isSyncing && 'animate-spin')} />
                {isSyncing ? 'Đang lưu...' : 'Lưu & Đồng bộ dữ liệu'}
              </button>
            </div>
            <div className="overflow-x-auto max-h-[480px] custom-scrollbar">
              <table className="w-full text-left border-collapse min-w-[780px]">
                <thead className="sticky top-0 bg-slate-50 dark:bg-slate-800 z-10 border-b border-slate-200 dark:border-slate-700">
                  <tr>
                    <th className="px-5 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider w-16 text-center">STT</th>
                    <th className="px-5 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Họ và tên / Đơn vị</th>
                    <th className="px-5 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Địa chỉ / Nội dung</th>
                    <th className="px-5 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Phân loại</th>
                    <th className="px-5 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider text-right">Số tiền (VNĐ)</th>
                    <th className="px-5 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase tracking-wider w-28 text-center">Phiếu lẻ</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                  {incomeData.map((item) => (
                    <tr key={`income-${item.id}`} className="hover:bg-emerald-50/30 dark:hover:bg-emerald-900/10 transition-colors">
                      <td className="px-5 py-3.5 text-xs text-slate-500 dark:text-slate-400 text-center font-mono tabular-nums">{item.stt}</td>
                      <td className="px-5 py-3.5 font-bold text-sm text-slate-900 dark:text-white">{item.name}</td>
                      <td className="px-5 py-3.5 text-sm text-slate-600 dark:text-slate-300">{item.address}</td>
                      <td className="px-5 py-3.5">
                        <span className="px-2.5 py-1 rounded-md bg-emerald-50 dark:bg-emerald-900/30 text-emerald-800 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800 text-[11px] font-bold">
                          Thu học phí
                        </span>
                      </td>
                      <td className="px-5 py-3.5 text-sm text-emerald-700 dark:text-emerald-400 text-right font-bold font-mono tabular-nums">
                        +{item.amount.toLocaleString('vi-VN')}
                      </td>
                      <td className="px-5 py-3.5 text-center">
                        <button
                          onClick={() => exportReceipt(item)}
                          className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-teal-50 dark:bg-teal-900/30 text-teal-700 dark:text-teal-300 hover:bg-teal-100 text-xs font-bold transition-colors"
                          title="Tải phiếu thu cá nhân"
                        >
                          <Download className="w-3.5 h-3.5" /> PT
                        </button>
                      </td>
                    </tr>
                  ))}
                  {expenseData.map((item) => (
                    <tr key={`expense-${item.id}`} className="hover:bg-amber-50/30 dark:hover:bg-amber-900/10 transition-colors">
                      <td className="px-5 py-3.5 text-xs text-slate-500 dark:text-slate-400 text-center font-mono tabular-nums">{item.stt}</td>
                      <td className="px-5 py-3.5 font-bold text-sm text-slate-900 dark:text-white">{item.name}</td>
                      <td className="px-5 py-3.5 text-sm text-slate-600 dark:text-slate-300">{item.content || item.address}</td>
                      <td className="px-5 py-3.5">
                        <span className="px-2.5 py-1 rounded-md bg-amber-50 dark:bg-amber-900/30 text-amber-800 dark:text-amber-300 border border-amber-200 dark:border-amber-800 text-[11px] font-bold">
                          Chi phí
                        </span>
                      </td>
                      <td className="px-5 py-3.5 text-sm text-amber-700 dark:text-amber-400 text-right font-bold font-mono tabular-nums">
                        -{item.amount.toLocaleString('vi-VN')}
                      </td>
                      <td className="px-5 py-3.5 text-center">
                        <button
                          onClick={() => exportPaymentVoucher(item)}
                          className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-amber-50 dark:bg-amber-900/30 text-amber-700 dark:text-amber-300 hover:bg-amber-100 text-xs font-bold transition-colors"
                          title="Tải phiếu chi cá nhân"
                        >
                          <Download className="w-3.5 h-3.5" /> PC
                        </button>
                      </td>
                    </tr>
                  ))}
                  {incomeData.length === 0 && expenseData.length === 0 && (
                    <tr>
                      <td colSpan={6} className="px-6 py-16 text-center">
                        <div className="flex flex-col items-center gap-2.5">
                          <div className="w-14 h-14 bg-teal-50 dark:bg-teal-900/20 rounded-2xl flex items-center justify-center text-teal-600 dark:text-teal-400">
                            <FileSpreadsheet className="w-7 h-7" />
                          </div>
                          <p className="text-sm font-bold text-slate-900 dark:text-white">Chưa có dữ liệu thu - chi trong kỳ</p>
                          <p className="text-xs text-slate-500 dark:text-slate-400 max-w-md">
                            Nhấn nút <strong>"Xuất file mẫu Chấm công & Thu tiền (.xlsx)"</strong> để lấy biểu mẫu chuẩn và đưa dữ liệu lên hệ thống.
                          </p>
                        </div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {activeSubTab === 'finance-ledger' && (
        <div className="bg-white dark:bg-slate-900 p-8 sm:p-12 rounded-2xl border border-slate-200 dark:border-slate-800 text-center space-y-6 shadow-sm">
          <div className="w-20 h-20 bg-teal-50 dark:bg-teal-900/30 border border-teal-200/70 dark:border-teal-800 rounded-2xl flex items-center justify-center text-teal-700 dark:text-teal-300 mx-auto">
            <FileSpreadsheet className="w-9 h-9" />
          </div>
          <div className="max-w-xl mx-auto space-y-2">
            <h3 className="text-xl font-bold text-slate-900 dark:text-white">
              Xuất Sổ Chi tiết Doanh thu (Mẫu S1a-HKD)
            </h3>
            <p className="text-sm text-slate-500 dark:text-slate-400 leading-relaxed">
              Tổng hợp toàn bộ {incomeData.length} khoản thu học phí (Tổng cộng:{' '}
              <strong className="text-teal-700 dark:text-teal-300 font-mono tabular-nums">
                {totalIncome.toLocaleString('vi-VN')} VNĐ
              </strong>
              ) thành Sổ chi tiết doanh thu theo Thông tư 152/2025/TT-BTC.
            </p>
          </div>
          <div>
            <button
              onClick={exportS1aHKD}
              className="btn-primary inline-flex items-center gap-2.5 px-8 py-3.5 text-sm"
            >
              <Download className="w-5 h-5" />
              Tải xuống Sổ doanh thu S1a-HKD (.docx)
            </button>
          </div>
          <p className="text-xs text-slate-400 dark:text-slate-500">
            Định dạng Microsoft Word (.docx) chuẩn A4 • Sẵn sàng in ấn & lưu trữ hồ sơ thuế
          </p>
        </div>
      )}

      {activeSubTab === 'finance-vouchers' && (
        <div className="bg-white dark:bg-slate-900 p-8 sm:p-12 rounded-2xl border border-slate-200 dark:border-slate-800 text-center space-y-6 shadow-sm">
          <div className="w-20 h-20 bg-emerald-50 dark:bg-emerald-900/30 border border-emerald-200/70 dark:border-emerald-800 rounded-2xl flex items-center justify-center text-emerald-700 dark:text-emerald-300 mx-auto">
            <Receipt className="w-9 h-9" />
          </div>
          <div className="max-w-xl mx-auto space-y-2">
            <h3 className="text-xl font-bold text-slate-900 dark:text-white">
              Xuất Hàng loạt Phiếu Thu & Phiếu Chi (.docx)
            </h3>
            <p className="text-sm text-slate-500 dark:text-slate-400 leading-relaxed">
              Tự động tạo {incomeData.length} phiếu thu và {expenseData.length} phiếu chi tiền mặt theo Mẫu 01-TT. Bố cục 2 phiếu trên 1 trang A4 giúp tiết kiệm giấy in.
            </p>
          </div>
          <div>
            <button
              onClick={exportAllVouchers}
              disabled={incomeData.length === 0 && expenseData.length === 0}
              className="btn-primary inline-flex items-center gap-2.5 px-8 py-3.5 text-sm disabled:opacity-50"
            >
              <Download className="w-5 h-5" />
              Xuất toàn bộ Phiếu Thu - Phiếu Chi (.docx)
            </button>
          </div>
          <p className="text-xs text-slate-400 dark:text-slate-500">
            Tự động đánh số chứng từ liên tục và đọc số tiền bằng chữ tiếng Việt
          </p>
        </div>
      )}
    </div>
  );
}
