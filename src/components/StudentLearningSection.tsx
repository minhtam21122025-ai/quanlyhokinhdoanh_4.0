import React, { useState, useMemo } from 'react';
import {
  GraduationCap,
  Users,
  CheckCircle2,
  XCircle,
  Clock,
  Sparkles,
  Download,
  Search,
  Calendar,
  Layers,
  Check,
  FileSpreadsheet,
  UserCheck,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { format } from 'date-fns';
import SectionBanner from './SectionBanner';
import {
  Student,
  StudentAttendanceRecord,
  AttendanceStatus,
  BusinessInfo,
} from '../types';
import { cn } from '../lib/utils';

interface StudentLearningSectionProps {
  students: Student[];
  setStudents: (students: Student[]) => void;
  attendanceRecords: StudentAttendanceRecord[];
  setAttendanceRecords: (records: StudentAttendanceRecord[]) => void;
  businessInfo: BusinessInfo;
  onBack: () => void;
  onGoToStudentsUpload: () => void;
  isAdmin: boolean;
}

export function deriveAutoClassName(student: Student): string {
  const rawGrade = String(student.grade || '').trim();
  const baseGradeMatch = rawGrade.match(/\d+/);
  const gradeLabel = baseGradeMatch ? baseGradeMatch[0] : rawGrade || 'Chưa phân khối';
  const rawSubject = String(student.subject || '').trim();
  const primarySubject = rawSubject
    ? rawSubject.split(/[,;/-]/)[0].trim()
    : 'Tổng hợp';
  return `Lớp ${primarySubject} ${gradeLabel}`.replace(/\s+/g, ' ').trim();
}

const SHIFTS = [
  'Ca 1 (07h30 - 09h30)',
  'Ca 2 (09h30 - 11h30)',
  'Ca 3 (14h00 - 16h00)',
  'Ca 4 (17h00 - 19h00)',
  'Ca 5 (19h00 - 21h00)',
];

const STATUS_META: Record<
  AttendanceStatus,
  { label: string; short: string; activeClass: string; textClass: string }
> = {
  present: {
    label: 'Có mặt',
    short: 'CM',
    activeClass: 'bg-emerald-600 text-white border-emerald-600',
    textClass: 'text-emerald-700 dark:text-emerald-400',
  },
  late: {
    label: 'Đi muộn',
    short: 'ĐM',
    activeClass: 'bg-amber-500 text-white border-amber-500',
    textClass: 'text-amber-700 dark:text-amber-400',
  },
  absent_excused: {
    label: 'Nghỉ có phép',
    short: 'CP',
    activeClass: 'bg-sky-600 text-white border-sky-600',
    textClass: 'text-sky-700 dark:text-sky-400',
  },
  absent_unexcused: {
    label: 'Nghỉ không phép',
    short: 'KP',
    activeClass: 'bg-rose-600 text-white border-rose-600',
    textClass: 'text-rose-700 dark:text-rose-400',
  },
};

const StudentLearningSection: React.FC<StudentLearningSectionProps> = ({
  students,
  setStudents,
  attendanceRecords,
  setAttendanceRecords,
  businessInfo,
  onBack,
  onGoToStudentsUpload,
  isAdmin,
}) => {
  const [subView, setSubView] = useState<'classes' | 'attendance' | 'summary'>('attendance');
  const [selectedClass, setSelectedClass] = useState<string>('ALL');
  const [selectedDate, setSelectedDate] = useState<string>(() => format(new Date(), 'yyyy-MM-dd'));
  const [selectedShift, setSelectedShift] = useState<string>(SHIFTS[3]);
  const [selectedMonth, setSelectedMonth] = useState<string>(() => format(new Date(), 'yyyy-MM'));
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [notice, setNotice] = useState<string>('');

  const showNotice = (msg: string) => {
    setNotice(msg);
    setTimeout(() => setNotice(''), 3500);
  };

  // Normalized students with auto-assigned classes
  const studentsWithClass = useMemo(() => {
    return students.map((st) => ({
      ...st,
      effectiveClass: st.assignedClass?.trim() || deriveAutoClassName(st),
    }));
  }, [students]);

  const classGroups = useMemo(() => {
    const map = new Map<string, typeof studentsWithClass>();
    studentsWithClass.forEach((st) => {
      const list = map.get(st.effectiveClass) || [];
      list.push(st);
      map.set(st.effectiveClass, list);
    });
    return Array.from(map.entries()).sort((a, b) =>
      a[0].localeCompare(b[0], 'vi', { numeric: true })
    );
  }, [studentsWithClass]);

  const classNames = useMemo(() => classGroups.map(([name]) => name), [classGroups]);

  const activeClassFilter = useMemo(() => {
    if (selectedClass === 'ALL') return 'ALL';
    return classNames.includes(selectedClass) ? selectedClass : 'ALL';
  }, [selectedClass, classNames]);

  const filteredStudents = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return studentsWithClass.filter((st) => {
      const matchClass = activeClassFilter === 'ALL' || st.effectiveClass === activeClassFilter;
      const matchQuery =
        !q ||
        st.name.toLowerCase().includes(q) ||
        st.effectiveClass.toLowerCase().includes(q) ||
        st.phone.toLowerCase().includes(q) ||
        st.grade.toLowerCase().includes(q);
      return matchClass && matchQuery;
    });
  }, [studentsWithClass, activeClassFilter, searchQuery]);

  const handleAutoArrangeClasses = () => {
    if (students.length === 0) {
      showNotice('Chưa có học sinh trong danh sách. Hãy tải danh sách học sinh lên trước.');
      return;
    }
    const updated = students.map((st) => ({
      ...st,
      assignedClass: deriveAutoClassName(st),
    }));
    setStudents(updated);
    showNotice(`Đã tự động sắp xếp ${updated.length} học sinh vào ${new Set(updated.map((s) => s.assignedClass)).size} lớp học.`);
  };

  const handleManualClassChange = (studentId: string, newClass: string) => {
    setStudents(
      students.map((st) =>
        st.id === studentId ? { ...st, assignedClass: newClass } : st
      )
    );
  };

  const getRecordForSession = (studentId: string) => {
    return attendanceRecords.find(
      (r) =>
        r.studentId === studentId &&
        r.date === selectedDate &&
        r.shift === selectedShift
    );
  };

  const setStudentAttendance = (
    student: { id: string; effectiveClass: string },
    status: AttendanceStatus
  ) => {
    const existing = getRecordForSession(student.id);
    if (existing) {
      setAttendanceRecords(
        attendanceRecords.map((r) =>
          r.id === existing.id
            ? { ...r, status, className: student.effectiveClass }
            : r
        )
      );
    } else {
      const next: StudentAttendanceRecord = {
        id: crypto.randomUUID(),
        studentId: student.id,
        className: student.effectiveClass,
        date: selectedDate,
        shift: selectedShift,
        status,
      };
      setAttendanceRecords([...attendanceRecords, next]);
    }
  };

  const updateSessionNote = (student: { id: string; effectiveClass: string }, note: string) => {
    const existing = getRecordForSession(student.id);
    if (existing) {
      setAttendanceRecords(
        attendanceRecords.map((r) => (r.id === existing.id ? { ...r, note } : r))
      );
    } else {
      setAttendanceRecords([
        ...attendanceRecords,
        {
          id: crypto.randomUUID(),
          studentId: student.id,
          className: student.effectiveClass,
          date: selectedDate,
          shift: selectedShift,
          status: 'present',
          note,
        },
      ]);
    }
  };

  const markAllPresentForVisible = () => {
    if (filteredStudents.length === 0) return;
    const updatedMap = new Map<string, StudentAttendanceRecord>();
    attendanceRecords.forEach((r) => {
      updatedMap.set(`${r.studentId}__${r.date}__${r.shift}`, r);
    });

    filteredStudents.forEach((st) => {
      const key = `${st.id}__${selectedDate}__${selectedShift}`;
      const prev = updatedMap.get(key);
      updatedMap.set(key, {
        id: prev?.id || crypto.randomUUID(),
        studentId: st.id,
        className: st.effectiveClass,
        date: selectedDate,
        shift: selectedShift,
        status: 'present',
        note: prev?.note || '',
      });
    });

    setAttendanceRecords(Array.from(updatedMap.values()));
    showNotice(`Đã chấm công "Có mặt" cho ${filteredStudents.length} học sinh trong ca học.`);
  };

  // Current session KPI counts
  const sessionStats = useMemo(() => {
    let present = 0;
    let absent = 0;
    filteredStudents.forEach((st) => {
      const rec = attendanceRecords.find(
        (r) => r.studentId === st.id && r.date === selectedDate && r.shift === selectedShift
      );
      if (rec) {
        if (rec.status === 'present' || rec.status === 'late') present++;
        else absent++;
      }
    });
    return { present, absent };
  }, [filteredStudents, attendanceRecords, selectedDate, selectedShift]);

  // Monthly attendance matrix for selectedMonth
  const monthlySummary = useMemo(() => {
    const [yrStr, moStr] = selectedMonth.split('-');
    const year = Number(yrStr) || new Date().getFullYear();
    const month = Number(moStr) || new Date().getMonth() + 1;
    const daysInMonth = new Date(year, month, 0).getDate();

    const rows = filteredStudents.map((st, index) => {
      const monthRecords = attendanceRecords.filter(
        (r) => r.studentId === st.id && r.date.startsWith(selectedMonth)
      );
      const dayMarks: Record<number, string> = {};
      let presentSessions = 0;
      let lateSessions = 0;
      let absentExcused = 0;
      let absentUnexcused = 0;

      monthRecords.forEach((r) => {
        const d = Number(r.date.split('-')[2]);
        if (r.status === 'present') {
          presentSessions++;
          dayMarks[d] = 'X';
        } else if (r.status === 'late') {
          lateSessions++;
          presentSessions++;
          dayMarks[d] = 'M';
        } else if (r.status === 'absent_excused') {
          absentExcused++;
          if (!dayMarks[d]) dayMarks[d] = 'P';
        } else if (r.status === 'absent_unexcused') {
          absentUnexcused++;
          if (!dayMarks[d]) dayMarks[d] = 'K';
        }
      });

      return {
        stt: index + 1,
        student: st,
        dayMarks,
        presentSessions,
        lateSessions,
        absentExcused,
        absentUnexcused,
      };
    });

    return { daysInMonth, rows };
  }, [filteredStudents, attendanceRecords, selectedMonth]);

  const exportAttendanceExcel = () => {
    const { daysInMonth, rows } = monthlySummary;
    const [yrStr, moStr] = selectedMonth.split('-');

    const titleRow = [
      `${businessInfo.name || 'CƠ SỞ GIÁO DỤC HOÀNG GIA'} - BẢNG CHẤM CÔNG HỌC SINH THÁNG ${moStr}/${yrStr}`,
    ];
    const subRow = [
      `Lớp xuất báo cáo: ${activeClassFilter === 'ALL' ? 'Tất cả các lớp' : activeClassFilter}`,
    ];

    const header = [
      'STT',
      'HỌ VÀ TÊN HỌC SINH',
      'LỚP XẾP TỰ ĐỘNG',
      'KHỐI',
      'SĐT PHỤ HUYNH',
      ...Array.from({ length: 31 }, (_, i) => `${i + 1}`),
      'TỔNG BUỔI HỌC',
      'NGHỈ CÓ PHÉP',
      'NGHỈ KHÔNG PHÉP',
      'GHI CHÚ',
    ];

    const dataRows = rows.map((r) => [
      r.stt,
      r.student.name,
      r.student.effectiveClass,
      r.student.grade,
      r.student.phone,
      ...Array.from({ length: 31 }, (_, i) =>
        i + 1 <= daysInMonth ? r.dayMarks[i + 1] || '' : ''
      ),
      r.presentSessions,
      r.absentExcused,
      r.absentUnexcused,
      '',
    ]);

    const ws = XLSX.utils.aoa_to_sheet([titleRow, subRow, [], header, ...dataRows]);
    ws['!cols'] = [
      { wch: 6 },
      { wch: 26 },
      { wch: 22 },
      { wch: 8 },
      { wch: 14 },
      ...Array.from({ length: 31 }, () => ({ wch: 4 })),
      { wch: 15 },
      { wch: 14 },
      { wch: 16 },
      { wch: 18 },
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Bang_Cham_Cong_HS');
    XLSX.writeFile(wb, `Bang_Cham_Cong_Hoc_Sinh_${moStr}_${yrStr}.xlsx`);
  };

  return (
    <div className="space-y-6">
      <SectionBanner
        title="Quản lý Học tập & Chấm công Học sinh"
        subtitle="Tự động phân lớp từ danh sách học sinh theo môn học - khối lớp và chấm công tham gia từng buổi học."
        icon={GraduationCap}
        onBack={onBack}
        actions={
          <>
            <button
              type="button"
              onClick={handleAutoArrangeClasses}
              className="min-h-[42px] px-4 py-2 rounded-xl bg-white text-teal-800 hover:bg-teal-50 font-bold text-xs sm:text-sm flex items-center gap-2 shadow-xs transition-colors"
            >
              <Sparkles className="w-4 h-4 text-teal-600" />
              <span>Tự động sắp lớp</span>
            </button>
            <button
              type="button"
              onClick={exportAttendanceExcel}
              className="min-h-[42px] px-4 py-2 rounded-xl bg-amber-400 text-slate-950 hover:bg-amber-300 font-bold text-xs sm:text-sm flex items-center gap-2 shadow-xs transition-colors"
            >
              <Download className="w-4 h-4" />
              <span>Xuất bảng chấm công (.xlsx)</span>
            </button>
          </>
        }
        stats={[
          {
            label: 'Tổng học sinh',
            value: students.length,
            sublabel: 'Đồng bộ từ Quản lý học sinh',
          },
          {
            label: 'Số lớp đã xếp',
            value: classGroups.length,
            sublabel: 'Tự động theo môn & khối',
          },
          {
            label: 'Có mặt ca đang chọn',
            value: sessionStats.present,
            sublabel: `${selectedDate.split('-').reverse().join('/')} · ${selectedShift.split(' ')[0]} ${selectedShift.split(' ')[1] || ''}`,
          },
          {
            label: 'Vắng mặt ca đang chọn',
            value: sessionStats.absent,
            sublabel: 'Có phép & không phép',
          },
        ]}
        tabs={[
          { id: 'attendance', label: 'Bảng chấm công buổi học', icon: UserCheck },
          { id: 'classes', label: 'Danh sách sắp lớp tự động', icon: Layers },
          { id: 'summary', label: 'Tổng hợp chuyên cần tháng', icon: FileSpreadsheet },
        ]}
        activeTab={subView}
        onTabChange={(id) => setSubView(id as 'classes' | 'attendance' | 'summary')}
      />

      {notice && (
        <div className="p-4 rounded-xl bg-teal-50 dark:bg-teal-950/50 border border-teal-200 dark:border-teal-800 text-teal-800 dark:text-teal-200 text-sm font-semibold flex items-center gap-2.5">
          <CheckCircle2 className="w-5 h-5 text-teal-600 shrink-0" />
          <span>{notice}</span>
        </div>
      )}

      {/* Empty State if no students uploaded yet */}
      {students.length === 0 ? (
        <div className="saas-card py-16 text-center space-y-4">
          <div className="w-16 h-16 rounded-2xl bg-teal-50 dark:bg-teal-950/50 text-teal-600 flex items-center justify-center mx-auto">
            <Users className="w-8 h-8" />
          </div>
          <div className="max-w-md mx-auto space-y-1.5">
            <h3 className="text-lg font-extrabold text-slate-900 dark:text-white">
              Chưa có danh sách học sinh để sắp lớp & chấm công
            </h3>
            <p className="text-sm text-slate-600 dark:text-slate-400">
              Khi Quý Thầy Cô tải danh sách học sinh lên tại mục Quản lý học sinh, hệ thống sẽ tự động phân lớp theo khối và môn đăng ký học.
            </p>
          </div>
          <div className="pt-2 flex justify-center">
            <button
              type="button"
              onClick={onGoToStudentsUpload}
              className="btn-primary"
            >
              Chuyển tới Tải danh sách học sinh
            </button>
          </div>
        </div>
      ) : (
        <>
          {/* Filter Bar */}
          <div className="saas-card !p-4 flex flex-col lg:flex-row items-stretch lg:items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={() => setSelectedClass('ALL')}
                className={cn(
                  'min-h-[40px] px-3.5 py-2 rounded-xl text-xs font-bold border transition-colors',
                  activeClassFilter === 'ALL'
                    ? 'bg-teal-600 text-white border-teal-600'
                    : 'bg-slate-50 dark:bg-slate-800 text-slate-700 dark:text-slate-300 border-slate-200 dark:border-slate-700 hover:border-teal-500'
                )}
              >
                Tất cả lớp ({students.length})
              </button>
              {classGroups.map(([clsName, members]) => (
                <button
                  key={clsName}
                  type="button"
                  onClick={() => setSelectedClass(clsName)}
                  className={cn(
                    'min-h-[40px] px-3.5 py-2 rounded-xl text-xs font-bold border transition-colors',
                    activeClassFilter === clsName
                      ? 'bg-teal-600 text-white border-teal-600'
                      : 'bg-slate-50 dark:bg-slate-800 text-slate-700 dark:text-slate-300 border-slate-200 dark:border-slate-700 hover:border-teal-500'
                  )}
                >
                  {clsName} ({members.length})
                </button>
              ))}
            </div>

            <div className="relative min-w-[240px]">
              <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Tìm tên học sinh, lớp, SĐT..."
                className="w-full min-h-[42px] pl-10 pr-4 py-2 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl text-sm font-medium text-slate-900 dark:text-white outline-none focus:border-teal-600"
              />
            </div>
          </div>

          {/* Subview 1: Daily Session Attendance */}
          {subView === 'attendance' && (
            <div className="space-y-4">
              <div className="saas-card !p-4 flex flex-col md:flex-row md:items-end justify-between gap-4">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 flex-1">
                  <div>
                    <label className="block text-xs font-bold text-slate-600 dark:text-slate-300 mb-1.5">
                      Ngày học điểm danh
                    </label>
                    <div className="relative">
                      <Calendar className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none" />
                      <input
                        type="date"
                        value={selectedDate}
                        onChange={(e) => setSelectedDate(e.target.value)}
                        className="w-full min-h-[42px] pl-10 pr-3 py-2 rounded-xl border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-sm font-semibold text-slate-900 dark:text-white tabular-nums outline-none focus:border-teal-600"
                      />
                    </div>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-600 dark:text-slate-300 mb-1.5">
                      Ca học trong ngày
                    </label>
                    <select
                      value={selectedShift}
                      onChange={(e) => setSelectedShift(e.target.value)}
                      className="w-full min-h-[42px] px-3.5 py-2 rounded-xl border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-sm font-semibold text-slate-900 dark:text-white outline-none focus:border-teal-600"
                    >
                      {SHIFTS.map((s) => (
                        <option key={s} value={s}>
                          {s}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                <button
                  type="button"
                  onClick={markAllPresentForVisible}
                  className="btn-primary shrink-0"
                >
                  <Check className="w-4 h-4" />
                  <span>Điểm danh Có mặt tất cả ({filteredStudents.length})</span>
                </button>
              </div>

              <div className="saas-card !p-0 overflow-hidden">
                <div className="overflow-x-auto custom-scrollbar">
                  <table className="w-full text-left border-collapse min-w-[860px]">
                    <thead>
                      <tr className="bg-slate-50 dark:bg-slate-800/70 border-b border-slate-200 dark:border-slate-800">
                        <th className="px-4 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase w-14 text-center">
                          STT
                        </th>
                        <th className="px-4 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">
                          Họ và tên học sinh
                        </th>
                        <th className="px-4 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">
                          Lớp học
                        </th>
                        <th className="px-4 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">
                          Trạng thái tham gia buổi học
                        </th>
                        <th className="px-4 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase w-56">
                          Ghi chú buổi học
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200/80 dark:divide-slate-800">
                      {filteredStudents.map((st, idx) => {
                        const rec = getRecordForSession(st.id);
                        return (
                          <tr
                            key={st.id}
                            className="hover:bg-slate-50/70 dark:hover:bg-slate-800/40 transition-colors"
                          >
                            <td className="px-4 py-3.5 text-sm font-mono font-semibold text-slate-500 text-center tabular-nums">
                              {idx + 1}
                            </td>
                            <td className="px-4 py-3.5">
                              <div className="font-bold text-sm text-slate-900 dark:text-white">
                                {st.name}
                              </div>
                              <div className="text-xs text-slate-500 dark:text-slate-400 mt-0.5 tabular-nums">
                                Khối/Lớp trường: {st.grade || '-'} · SĐT PH: {st.phone || '-'}
                              </div>
                            </td>
                            <td className="px-4 py-3.5">
                              <span className="text-xs font-bold text-teal-700 dark:text-teal-300">
                                {st.effectiveClass}
                              </span>
                            </td>
                            <td className="px-4 py-3.5">
                              <div className="flex flex-wrap items-center gap-1.5">
                                {(Object.keys(STATUS_META) as AttendanceStatus[]).map((statusKey) => {
                                  const meta = STATUS_META[statusKey];
                                  const isSelected = rec?.status === statusKey;
                                  return (
                                    <button
                                      key={statusKey}
                                      type="button"
                                      onClick={() => setStudentAttendance(st, statusKey)}
                                      className={cn(
                                        'min-h-[38px] px-3 py-1.5 rounded-lg text-xs font-bold border transition-colors',
                                        isSelected
                                          ? meta.activeClass
                                          : 'bg-white dark:bg-slate-900 text-slate-700 dark:text-slate-300 border-slate-200 dark:border-slate-700 hover:border-teal-500'
                                      )}
                                    >
                                      {meta.label}
                                    </button>
                                  );
                                })}
                              </div>
                            </td>
                            <td className="px-4 py-3.5">
                              <input
                                type="text"
                                value={rec?.note || ''}
                                onChange={(e) => updateSessionNote(st, e.target.value)}
                                placeholder="Nhận xét / lý do..."
                                className="w-full min-h-[38px] px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 text-xs font-medium text-slate-900 dark:text-white outline-none focus:border-teal-600"
                              />
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* Subview 2: Auto-Arranged Classes */}
          {subView === 'classes' && (
            <div className="space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                {classGroups.map(([clsName, members]) => (
                  <div key={clsName} className="saas-card space-y-3">
                    <div className="flex items-center justify-between border-b border-slate-100 dark:border-slate-800 pb-3">
                      <div>
                        <h3 className="text-base font-extrabold text-slate-900 dark:text-white">
                          {clsName}
                        </h3>
                        <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5 tabular-nums">
                          Sĩ số: {members.length} học sinh
                        </p>
                      </div>
                      <button
                        type="button"
                        onClick={() => {
                          setSelectedClass(clsName);
                          setSubView('attendance');
                        }}
                        className="min-h-[38px] px-3 py-1.5 rounded-lg bg-teal-50 dark:bg-teal-950/50 text-teal-700 dark:text-teal-300 hover:bg-teal-600 hover:text-white text-xs font-bold transition-colors"
                      >
                        Chấm công lớp
                      </button>
                    </div>
                    <div className="space-y-2 max-h-56 overflow-y-auto custom-scrollbar pr-1">
                      {members.map((m, idx) => (
                        <div
                          key={m.id}
                          className="flex items-center justify-between gap-2 py-1.5 border-b border-slate-100 dark:border-slate-800/60 last:border-none text-xs"
                        >
                          <span className="font-semibold text-slate-800 dark:text-slate-200 truncate">
                            {idx + 1}. {m.name}
                          </span>
                          {isAdmin ? (
                            <input
                              type="text"
                              value={m.assignedClass ?? m.effectiveClass}
                              onChange={(e) => handleManualClassChange(m.id, e.target.value)}
                              className="w-36 px-2 py-1 rounded border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 text-[11px] font-semibold text-teal-700 dark:text-teal-300 text-right outline-none focus:border-teal-600"
                              title="Chỉnh sửa tên lớp nếu muốn chuyển lớp"
                            />
                          ) : (
                            <span className="text-slate-500 tabular-nums">{m.grade}</span>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Subview 3: Monthly Attendance Matrix */}
          {subView === 'summary' && (
            <div className="space-y-4">
              <div className="saas-card !p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                <div className="flex items-center gap-3">
                  <label className="text-sm font-bold text-slate-700 dark:text-slate-200">
                    Tháng tổng hợp:
                  </label>
                  <input
                    type="month"
                    value={selectedMonth}
                    onChange={(e) => setSelectedMonth(e.target.value)}
                    className="min-h-[42px] px-3.5 py-2 rounded-xl border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-sm font-bold text-slate-900 dark:text-white tabular-nums outline-none focus:border-teal-600"
                  />
                </div>
                <div className="flex flex-wrap items-center gap-4 text-xs font-semibold text-slate-600 dark:text-slate-300">
                  <span className="flex items-center gap-1.5">
                    <strong className="text-emerald-600 font-mono">X</strong>: Có mặt
                  </span>
                  <span className="flex items-center gap-1.5">
                    <strong className="text-amber-600 font-mono">M</strong>: Đi muộn
                  </span>
                  <span className="flex items-center gap-1.5">
                    <strong className="text-sky-600 font-mono">P</strong>: Nghỉ có phép
                  </span>
                  <span className="flex items-center gap-1.5">
                    <strong className="text-rose-600 font-mono">K</strong>: Nghỉ không phép
                  </span>
                </div>
              </div>

              <div className="saas-card !p-0 overflow-hidden">
                <div className="overflow-x-auto custom-scrollbar">
                  <table className="w-full text-left border-collapse min-w-[1100px]">
                    <thead>
                      <tr className="bg-slate-50 dark:bg-slate-800/80 border-b border-slate-200 dark:border-slate-800">
                        <th className="px-3 py-3 text-xs font-bold text-slate-600 dark:text-slate-300 text-center w-12">
                          STT
                        </th>
                        <th className="px-4 py-3 text-xs font-bold text-slate-600 dark:text-slate-300 min-w-[180px]">
                          Họ và tên
                        </th>
                        <th className="px-3 py-3 text-xs font-bold text-slate-600 dark:text-slate-300 min-w-[130px]">
                          Lớp
                        </th>
                        {Array.from({ length: monthlySummary.daysInMonth }, (_, i) => (
                          <th
                            key={i + 1}
                            className="px-1.5 py-3 text-[11px] font-mono font-bold text-slate-600 dark:text-slate-300 text-center w-7 border-l border-slate-200/60 dark:border-slate-800"
                          >
                            {i + 1}
                          </th>
                        ))}
                        <th className="px-3 py-3 text-xs font-bold text-teal-700 dark:text-teal-300 text-right border-l border-slate-200 dark:border-slate-800">
                          Số buổi
                        </th>
                        <th className="px-3 py-3 text-xs font-bold text-rose-600 text-right">
                          Vắng
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200/80 dark:divide-slate-800">
                      {monthlySummary.rows.map((row) => (
                        <tr
                          key={row.student.id}
                          className="hover:bg-slate-50/70 dark:hover:bg-slate-800/40"
                        >
                          <td className="px-3 py-2.5 text-xs font-mono text-center text-slate-500 tabular-nums">
                            {row.stt}
                          </td>
                          <td className="px-4 py-2.5 text-xs sm:text-sm font-bold text-slate-900 dark:text-white">
                            {row.student.name}
                          </td>
                          <td className="px-3 py-2.5 text-xs font-semibold text-teal-700 dark:text-teal-300">
                            {row.student.effectiveClass}
                          </td>
                          {Array.from({ length: monthlySummary.daysInMonth }, (_, i) => {
                            const mark = row.dayMarks[i + 1] || '';
                            return (
                              <td
                                key={i + 1}
                                className={cn(
                                  'px-1 py-2.5 text-xs font-mono font-bold text-center border-l border-slate-100 dark:border-slate-800/70',
                                  mark === 'X' && 'text-emerald-600',
                                  mark === 'M' && 'text-amber-600',
                                  mark === 'P' && 'text-sky-600',
                                  mark === 'K' && 'text-rose-600'
                                )}
                              >
                                {mark}
                              </td>
                            );
                          })}
                          <td className="px-3 py-2.5 text-sm font-mono font-extrabold text-teal-700 dark:text-teal-300 text-right border-l border-slate-200 dark:border-slate-800 tabular-nums">
                            {row.presentSessions}
                          </td>
                          <td className="px-3 py-2.5 text-sm font-mono font-bold text-rose-600 text-right tabular-nums">
                            {row.absentExcused + row.absentUnexcused}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
};

export default StudentLearningSection;
