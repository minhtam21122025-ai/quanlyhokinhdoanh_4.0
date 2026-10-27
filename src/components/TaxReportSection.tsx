import React, { useState, useMemo } from 'react';
import {
  FileSpreadsheet,
  Download,
  FileText,
  RefreshCw,
  CheckCircle2,
  Building2,
  RotateCcw,
  Calculator,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import {
  Document,
  Packer,
  Paragraph,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  TextRun,
  VerticalAlign,
  BorderStyle,
} from 'docx';
import { saveAs } from 'file-saver';
import SectionBanner from './SectionBanner';
import {
  AnnualTaxReport,
  BusinessInfo,
  FinancialConfig,
  IncomeItem,
  MonthlyTaxEntry,
} from '../types';

interface TaxReportSectionProps {
  reportsByYear: Record<number, AnnualTaxReport>;
  setReportsByYear: React.Dispatch<React.SetStateAction<Record<number, AnnualTaxReport>>>;
  businessInfo: BusinessInfo;
  financialConfig: FinancialConfig;
  incomeData: IncomeItem[];
  onBack: () => void;
}

function createDefaultYearReport(
  year: number,
  businessInfo: BusinessInfo,
  financialConfig: FinancialConfig
): AnnualTaxReport {
  return {
    year,
    businessName: businessInfo.name || 'CƠ SỞ GIÁO DỤC HOÀNG GIA',
    taxId: businessInfo.taxId || financialConfig.taxCode || '',
    owner: businessInfo.owner || '',
    address: businessInfo.address || '',
    businessSector: 'Dịch vụ giáo dục, dạy thêm học thêm',
    preparer: financialConfig.preparer || businessInfo.owner || '',
    months: Array.from({ length: 12 }, (_, idx) => ({
      month: idx + 1,
      amount: 0,
      note: '',
    })),
  };
}

function readAmountInVietnamese(amount: number): string {
  if (!amount || amount <= 0) return 'Không đồng';
  const units = ['', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín'];
  const tens = ['', 'mười', 'hai mươi', 'ba mươi', 'bốn mươi', 'năm mươi', 'sáu mươi', 'bảy mươi', 'tám mươi', 'chín mươi'];
  const hundreds = ['không trăm', 'một trăm', 'hai trăm', 'ba trăm', 'bốn trăm', 'năm trăm', 'sáu trăm', 'bảy trăm', 'tám trăm', 'chín trăm'];
  const groups = ['', 'nghìn', 'triệu', 'tỷ', 'nghìn tỷ'];

  const readBlock = (n: number, first: boolean) => {
    let s = '';
    const h = Math.floor(n / 100);
    const t = Math.floor((n % 100) / 10);
    const u = n % 10;
    if (h > 0 || !first) s += hundreds[h] + ' ';
    if (t > 0) {
      s += tens[t] + ' ';
      if (u === 1 && t > 1) s = s.trim() + ' mốt ';
      else if (u === 5) s = s.trim() + ' lăm ';
      else if (u > 0) s += units[u] + ' ';
    } else if (u > 0) {
      if (!first || h > 0) s += 'lẻ ';
      s += units[u] + ' ';
    }
    return s.trim();
  };

  let res = '';
  let idx = 0;
  let temp = Math.floor(amount);
  while (temp > 0) {
    const val = temp % 1000;
    if (val > 0) {
      res = `${readBlock(val, temp < 1000)} ${groups[idx]} ${res}`.trim();
    }
    temp = Math.floor(temp / 1000);
    idx++;
  }
  res = res.replace(/^không trăm\s+/i, '').replace(/^lẻ\s+/i, '').trim();
  return res.charAt(0).toUpperCase() + res.slice(1) + ' đồng chẵn';
}

const TaxReportSection: React.FC<TaxReportSectionProps> = ({
  reportsByYear,
  setReportsByYear,
  businessInfo,
  financialConfig,
  incomeData,
  onBack,
}) => {
  const currentYear = new Date().getFullYear();
  const [selectedYear, setSelectedYear] = useState<number>(currentYear);
  const [notice, setNotice] = useState<string>('');

  const showNotice = (msg: string) => {
    setNotice(msg);
    setTimeout(() => setNotice(''), 3500);
  };

  const activeReport: AnnualTaxReport = useMemo(() => {
    const existing = reportsByYear[selectedYear];
    if (existing) return existing;
    return createDefaultYearReport(selectedYear, businessInfo, financialConfig);
  }, [reportsByYear, selectedYear, businessInfo, financialConfig]);

  const updateReport = (updater: (prev: AnnualTaxReport) => AnnualTaxReport) => {
    setReportsByYear((prevMap) => {
      const current =
        prevMap[selectedYear] ||
        createDefaultYearReport(selectedYear, businessInfo, financialConfig);
      return {
        ...prevMap,
        [selectedYear]: updater(current),
      };
    });
  };

  const handleMonthAmountChange = (monthNum: number, rawVal: string) => {
    const digitsOnly = rawVal.replace(/[^\d]/g, '');
    const numeric = digitsOnly ? Number(digitsOnly) : 0;
    updateReport((prev) => ({
      ...prev,
      months: prev.months.map((m) =>
        m.month === monthNum ? { ...m, amount: numeric } : m
      ),
    }));
  };

  const handleMonthNoteChange = (monthNum: number, note: string) => {
    updateReport((prev) => ({
      ...prev,
      months: prev.months.map((m) =>
        m.month === monthNum ? { ...m, note } : m
      ),
    }));
  };

  const syncFromFinanceIncome = () => {
    if (incomeData.length === 0) {
      showNotice('Chưa có dữ liệu thu học phí trong mục Quản lý Tài chính để đồng bộ.');
      return;
    }

    const monthTotals = new Array(12).fill(0);
    let matchedAnyDate = false;

    // Parse month from item.date (yyyy-MM-dd or dd/MM/yyyy) or financialConfig.reportPeriod
    const periodMatch = financialConfig.reportPeriod.match(/(\d{1,2})\s*\/\s*(\d{4})/);
    const fallbackMonth =
      periodMatch && Number(periodMatch[2]) === selectedYear
        ? Math.min(12, Math.max(1, Number(periodMatch[1])))
        : new Date().getMonth() + 1;

    incomeData.forEach((item) => {
      let itemMonth = fallbackMonth;
      let itemYear = selectedYear;
      if (item.date) {
        if (/^\d{4}-\d{2}-\d{2}/.test(item.date)) {
          const parts = item.date.split('-');
          itemYear = Number(parts[0]);
          itemMonth = Number(parts[1]);
          matchedAnyDate = true;
        } else if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(item.date)) {
          const parts = item.date.split('/');
          itemMonth = Number(parts[1]);
          itemYear = Number(parts[2]);
          matchedAnyDate = true;
        }
      }
      if (itemYear === selectedYear && itemMonth >= 1 && itemMonth <= 12) {
        monthTotals[itemMonth - 1] += Number(item.amount || 0);
      }
    });

    if (!matchedAnyDate && incomeData.length > 0) {
      const sumAll = incomeData.reduce((s, it) => s + Number(it.amount || 0), 0);
      monthTotals[fallbackMonth - 1] = sumAll;
    }

    updateReport((prev) => ({
      ...prev,
      months: prev.months.map((m, idx) => ({
        ...m,
        amount: monthTotals[idx] > 0 ? monthTotals[idx] : m.amount,
        note:
          monthTotals[idx] > 0 && !m.note
            ? `Đồng bộ từ Quản lý tài chính (${incomeData.length} khoản thu)`
            : m.note,
      })),
    }));

    showNotice(`Đã cập nhật số liệu doanh thu từ mục Quản lý Tài chính vào năm ${selectedYear}.`);
  };

  const clearYearAmounts = () => {
    updateReport((prev) => ({
      ...prev,
      months: prev.months.map((m) => ({ ...m, amount: 0, note: '' })),
    }));
    showNotice(`Đã đặt lại số tiền 12 tháng của năm ${selectedYear} về 0.`);
  };

  const totalAnnualRevenue = useMemo(() => {
    return activeReport.months.reduce((sum, m) => sum + (Number(m.amount) || 0), 0);
  }, [activeReport.months]);

  const monthsWithRevenue = useMemo(() => {
    return activeReport.months.filter((m) => m.amount > 0).length;
  }, [activeReport.months]);

  const averageMonthlyRevenue = useMemo(() => {
    return monthsWithRevenue > 0 ? Math.round(totalAnnualRevenue / monthsWithRevenue) : 0;
  }, [totalAnnualRevenue, monthsWithRevenue]);

  const quarterlyTotals = useMemo(() => {
    const q = [0, 0, 0, 0];
    activeReport.months.forEach((m) => {
      const qIdx = Math.floor((m.month - 1) / 3);
      q[qIdx] += Number(m.amount) || 0;
    });
    return q;
  }, [activeReport.months]);

  // Export complete annual revenue in Excel (.xlsx)
  const exportAnnualTaxExcel = () => {
    const rows: any[][] = [
      ['CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM'],
      ['Độc lập - Tự do - Hạnh phúc'],
      [],
      [`BÁO CÁO TỔNG HỢP DOANH THU KÊ KHAI THUẾ HỘ KINH DOANH NĂM ${selectedYear}`],
      [],
      ['Tên hộ kinh doanh:', activeReport.businessName || businessInfo.name],
      ['Chủ hộ kinh doanh:', activeReport.owner || businessInfo.owner],
      ['Mã số thuế:', activeReport.taxId || businessInfo.taxId],
      ['Địa chỉ kinh doanh:', activeReport.address || businessInfo.address],
      ['Ngành nghề kinh doanh:', activeReport.businessSector],
      [],
      ['STT', 'KỲ DOANH THU (THÁNG)', 'QUÝ', 'DOANH THU TRONG THÁNG (VNĐ)', 'GHI CHÚ / DIỄN GIẢI'],
    ];

    activeReport.months.forEach((m: MonthlyTaxEntry, idx: number) => {
      const quarter = `Quý ${Math.floor((m.month - 1) / 3) + 1}`;
      rows.push([
        idx + 1,
        `Tháng ${String(m.month).padStart(2, '0')}/${selectedYear}`,
        quarter,
        m.amount,
        m.note || `Doanh thu dịch vụ dạy thêm tháng ${String(m.month).padStart(2, '0')}/${selectedYear}`,
      ]);
    });

    rows.push([
      '',
      `TỔNG DOANH THU CẢ NĂM ${selectedYear}`,
      'Cả năm',
      totalAnnualRevenue,
      `Bằng chữ: ${readAmountInVietnamese(totalAnnualRevenue)}`,
    ]);
    rows.push([]);
    rows.push(['TỔNG HỢP THEO QUÝ', '', '', 'SỐ TIỀN (VNĐ)', '']);
    quarterlyTotals.forEach((qAmount, idx) => {
      rows.push(['', `Quý ${idx + 1}/${selectedYear}`, '', qAmount, '']);
    });
    rows.push([]);
    rows.push([
      '',
      'NGƯỜI LẬP BIỂU',
      '',
      'ĐẠI DIỆN HỘ KINH DOANH',
      '',
    ]);
    rows.push([
      '',
      activeReport.preparer || '',
      '',
      activeReport.owner || businessInfo.owner || '',
      '',
    ]);

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [
      { wch: 8 },
      { wch: 32 },
      { wch: 14 },
      { wch: 28 },
      { wch: 44 },
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `Doanh_Thu_${selectedYear}`);
    XLSX.writeFile(wb, `Bao_Cao_Thue_Doanh_Thu_Nam_${selectedYear}.xlsx`);
  };

  // Export complete annual revenue in Word (.docx)
  const exportAnnualTaxWord = async () => {
    const today = new Date();
    const dStr = String(today.getDate()).padStart(2, '0');
    const mStr = String(today.getMonth() + 1).padStart(2, '0');
    const yStr = String(today.getFullYear());

    const noBorder = {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    };

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              margin: { top: 1000, bottom: 1000, left: 1200, right: 1000 },
            },
          },
          children: [
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              borders: noBorder,
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      width: { size: 52, type: WidthType.PERCENTAGE },
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: (activeReport.businessName || businessInfo.name || 'HỘ KINH DOANH').toUpperCase(),
                              bold: true,
                              size: 22,
                            }),
                          ],
                        }),
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: `MST: ${activeReport.taxId || businessInfo.taxId || '....................'}`,
                              size: 21,
                            }),
                          ],
                        }),
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: `Địa chỉ: ${activeReport.address || businessInfo.address || '....................'}`,
                              size: 20,
                            }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      width: { size: 48, type: WidthType.PERCENTAGE },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM',
                              bold: true,
                              size: 22,
                            }),
                          ],
                        }),
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: 'Độc lập - Tự do - Hạnh phúc',
                              bold: true,
                              size: 22,
                            }),
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
              spacing: { before: 400, after: 160 },
              children: [
                new TextRun({
                  text: `BẢNG KÊ TỔNG HỢP DOANH THU BÁO CÁO THUẾ NĂM ${selectedYear}`,
                  bold: true,
                  size: 30,
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 300 },
              children: [
                new TextRun({
                  text: `Chủ hộ kinh doanh: ${activeReport.owner || businessInfo.owner || '....................'} — Ngành nghề: ${activeReport.businessSector}`,
                  italics: true,
                  size: 22,
                }),
              ],
            }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: ['STT', 'Tháng kê khai', 'Doanh thu trong tháng (VNĐ)', 'Ghi chú'].map(
                    (h) =>
                      new TableCell({
                        verticalAlign: VerticalAlign.CENTER,
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ text: h, bold: true, size: 22 })],
                          }),
                        ],
                      })
                  ),
                }),
                ...activeReport.months.map(
                  (m, idx) =>
                    new TableRow({
                      children: [
                        new TableCell({
                          children: [
                            new Paragraph({
                              alignment: AlignmentType.CENTER,
                              children: [new TextRun({ text: String(idx + 1), size: 22 })],
                            }),
                          ],
                        }),
                        new TableCell({
                          children: [
                            new Paragraph({
                              alignment: AlignmentType.CENTER,
                              children: [
                                new TextRun({
                                  text: `Tháng ${String(m.month).padStart(2, '0')}/${selectedYear}`,
                                  bold: true,
                                  size: 22,
                                }),
                              ],
                            }),
                          ],
                        }),
                        new TableCell({
                          children: [
                            new Paragraph({
                              alignment: AlignmentType.RIGHT,
                              children: [
                                new TextRun({
                                  text: m.amount.toLocaleString('vi-VN'),
                                  bold: m.amount > 0,
                                  size: 22,
                                }),
                              ],
                            }),
                          ],
                        }),
                        new TableCell({
                          children: [
                            new Paragraph({
                              children: [new TextRun({ text: m.note || '', size: 21 })],
                            }),
                          ],
                        }),
                      ],
                    })
                ),
                new TableRow({
                  children: [
                    new TableCell({
                      columnSpan: 2,
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: `TỔNG DOANH THU CẢ NĂM ${selectedYear}`,
                              bold: true,
                              size: 24,
                            }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.RIGHT,
                          children: [
                            new TextRun({
                              text: `${totalAnnualRevenue.toLocaleString('vi-VN')} VNĐ`,
                              bold: true,
                              size: 24,
                            }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      children: [
                        new Paragraph({
                          children: [new TextRun({ text: 'Toàn bộ 12 tháng', italics: true, size: 21 })],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
            new Paragraph({
              spacing: { before: 240, after: 360 },
              children: [
                new TextRun({ text: 'Số tiền viết bằng chữ: ', bold: true, size: 22 }),
                new TextRun({
                  text: readAmountInVietnamese(totalAnnualRevenue),
                  italics: true,
                  size: 22,
                }),
              ],
            }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              borders: noBorder,
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      width: { size: 50, type: WidthType.PERCENTAGE },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [new TextRun({ text: 'NGƯỜI LẬP BIỂU', bold: true, size: 22 })],
                        }),
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [new TextRun({ text: '(Ký, ghi rõ họ tên)', italics: true, size: 20 })],
                        }),
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          spacing: { before: 1200 },
                          children: [
                            new TextRun({ text: activeReport.preparer || '', bold: true, size: 22 }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      width: { size: 50, type: WidthType.PERCENTAGE },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: `Ngày ${dStr} tháng ${mStr} năm ${yStr}`,
                              italics: true,
                              size: 20,
                            }),
                          ],
                        }),
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: 'ĐẠI DIỆN HỘ KINH DOANH',
                              bold: true,
                              size: 22,
                            }),
                          ],
                        }),
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: '(Ký, ghi rõ họ tên, đóng dấu)',
                              italics: true,
                              size: 20,
                            }),
                          ],
                        }),
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          spacing: { before: 1000 },
                          children: [
                            new TextRun({
                              text: activeReport.owner || businessInfo.owner || '',
                              bold: true,
                              size: 22,
                            }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        },
      ],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `Bao_Cao_Thue_Doanh_Thu_Nam_${selectedYear}.docx`);
  };

  const availableYears = [currentYear - 2, currentYear - 1, currentYear, currentYear + 1, currentYear + 2];

  return (
    <div className="space-y-6">
      <SectionBanner
        title="Báo cáo Thuế Hộ Kinh Doanh"
        subtitle="Nhập số tiền doanh thu của 12 tháng trong năm, tự động tính tổng doanh thu cuối kỳ và xuất toàn bộ báo cáo doanh thu năm."
        icon={Calculator}
        onBack={onBack}
        actions={
          <>
            <button
              type="button"
              onClick={syncFromFinanceIncome}
              className="min-h-[42px] px-4 py-2 rounded-xl bg-white/15 hover:bg-white/25 text-white border border-white/25 font-bold text-xs sm:text-sm flex items-center gap-2 transition-colors"
            >
              <RefreshCw className="w-4 h-4" />
              <span>Đồng bộ từ Tài chính</span>
            </button>
            <button
              type="button"
              onClick={exportAnnualTaxExcel}
              className="min-h-[42px] px-4 py-2 rounded-xl bg-amber-400 text-slate-950 hover:bg-amber-300 font-bold text-xs sm:text-sm flex items-center gap-2 shadow-xs transition-colors"
            >
              <Download className="w-4 h-4" />
              <span>Xuất toàn bộ doanh thu năm (.xlsx)</span>
            </button>
            <button
              type="button"
              onClick={exportAnnualTaxWord}
              className="min-h-[42px] px-4 py-2 rounded-xl bg-white text-teal-800 hover:bg-teal-50 font-bold text-xs sm:text-sm flex items-center gap-2 shadow-xs transition-colors"
            >
              <FileText className="w-4 h-4 text-teal-700" />
              <span>Xuất bản in Word (.docx)</span>
            </button>
          </>
        }
        stats={[
          {
            label: 'Năm kê khai thuế',
            value: `Năm ${selectedYear}`,
            sublabel: `MST: ${activeReport.taxId || businessInfo.taxId || 'Chưa cập nhật'}`,
          },
          {
            label: 'Tổng doanh thu cả năm',
            value: `${totalAnnualRevenue.toLocaleString('vi-VN')} đ`,
            sublabel: `${monthsWithRevenue}/12 tháng phát sinh doanh thu`,
          },
          {
            label: 'Bình quân tháng có thu',
            value: `${averageMonthlyRevenue.toLocaleString('vi-VN')} đ`,
            sublabel: 'Trung bình các tháng hoạt động',
          },
          {
            label: 'Ngưỡng tham chiếu HKD',
            value: totalAnnualRevenue > 100_000_000 ? 'Thuộc diện nộp thuế' : 'Dưới 100 triệu/năm',
            sublabel: 'Theo quy định thuế hộ kinh doanh',
          },
        ]}
      />

      {notice && (
        <div className="p-4 rounded-xl bg-teal-50 dark:bg-teal-950/50 border border-teal-200 dark:border-teal-800 text-teal-800 dark:text-teal-200 text-sm font-semibold flex items-center gap-2.5">
          <CheckCircle2 className="w-5 h-5 text-teal-600 shrink-0" />
          <span>{notice}</span>
        </div>
      )}

      {/* Household Business Tax Header Info */}
      <div className="saas-card space-y-5">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-200/80 dark:border-slate-800 pb-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-teal-50 dark:bg-teal-950/50 text-teal-700 dark:text-teal-400 flex items-center justify-center">
              <Building2 className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-base font-extrabold text-slate-900 dark:text-white">
                Thông tin Hộ kinh doanh & Năm báo cáo thuế
              </h3>
              <p className="text-xs text-slate-500 dark:text-slate-400">
                Thông tin này được in trực tiếp trên Bảng kê doanh thu năm xuất ra Excel và Word.
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2.5">
            <label className="text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">
              Chọn năm:
            </label>
            <select
              value={selectedYear}
              onChange={(e) => setSelectedYear(Number(e.target.value))}
              className="min-h-[42px] px-4 py-2 rounded-xl border border-teal-600 bg-teal-50/70 dark:bg-slate-800 text-sm font-extrabold text-teal-800 dark:text-teal-300 tabular-nums outline-none"
            >
              {availableYears.map((yr) => (
                <option key={yr} value={yr}>
                  Năm {yr}
                </option>
              ))}
            </select>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div>
            <label className="block text-xs font-bold text-slate-600 dark:text-slate-300 mb-1.5">
              Tên hộ kinh doanh
            </label>
            <input
              type="text"
              value={activeReport.businessName}
              onChange={(e) => updateReport((p) => ({ ...p, businessName: e.target.value }))}
              className="w-full min-h-[42px] px-3.5 py-2 rounded-xl border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/70 text-sm font-semibold text-slate-900 dark:text-white outline-none focus:border-teal-600"
            />
          </div>
          <div>
            <label className="block text-xs font-bold text-slate-600 dark:text-slate-300 mb-1.5">
              Mã số thuế HKD
            </label>
            <input
              type="text"
              value={activeReport.taxId}
              onChange={(e) => updateReport((p) => ({ ...p, taxId: e.target.value }))}
              placeholder="Nhập mã số thuế..."
              className="w-full min-h-[42px] px-3.5 py-2 rounded-xl border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/70 text-sm font-mono font-semibold text-slate-900 dark:text-white outline-none focus:border-teal-600"
            />
          </div>
          <div>
            <label className="block text-xs font-bold text-slate-600 dark:text-slate-300 mb-1.5">
              Chủ hộ / Người đại diện
            </label>
            <input
              type="text"
              value={activeReport.owner}
              onChange={(e) => updateReport((p) => ({ ...p, owner: e.target.value }))}
              className="w-full min-h-[42px] px-3.5 py-2 rounded-xl border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/70 text-sm font-semibold text-slate-900 dark:text-white outline-none focus:border-teal-600"
            />
          </div>
          <div className="md:col-span-2">
            <label className="block text-xs font-bold text-slate-600 dark:text-slate-300 mb-1.5">
              Địa chỉ kinh doanh
            </label>
            <input
              type="text"
              value={activeReport.address}
              onChange={(e) => updateReport((p) => ({ ...p, address: e.target.value }))}
              className="w-full min-h-[42px] px-3.5 py-2 rounded-xl border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/70 text-sm font-semibold text-slate-900 dark:text-white outline-none focus:border-teal-600"
            />
          </div>
          <div>
            <label className="block text-xs font-bold text-slate-600 dark:text-slate-300 mb-1.5">
              Người lập biểu
            </label>
            <input
              type="text"
              value={activeReport.preparer}
              onChange={(e) => updateReport((p) => ({ ...p, preparer: e.target.value }))}
              placeholder="Họ và tên người lập"
              className="w-full min-h-[42px] px-3.5 py-2 rounded-xl border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/70 text-sm font-semibold text-slate-900 dark:text-white outline-none focus:border-teal-600"
            />
          </div>
        </div>
      </div>

      {/* 12-Month Revenue Input Table + Bottom Annual Total */}
      <div className="saas-card !p-0 overflow-hidden">
        <div className="px-6 py-4 bg-slate-50 dark:bg-slate-800/60 border-b border-slate-200 dark:border-slate-800 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h3 className="text-base font-extrabold text-slate-900 dark:text-white">
              Bảng đăng nhập doanh thu 12 tháng — Năm {selectedYear}
            </h3>
            <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
              Nhập số tiền doanh thu thực tế của từng tháng (VNĐ). Hệ thống tự động cộng dồn Tổng doanh thu cả năm ở dòng cuối.
            </p>
          </div>
          <button
            type="button"
            onClick={clearYearAmounts}
            className="min-h-[38px] px-3.5 py-1.5 rounded-xl border border-slate-300 dark:border-slate-700 text-xs font-bold text-slate-600 dark:text-slate-300 hover:text-rose-600 hover:border-rose-300 flex items-center gap-1.5 self-start sm:self-center transition-colors"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>Làm mới 12 tháng</span>
          </button>
        </div>

        <div className="overflow-x-auto custom-scrollbar">
          <table className="w-full text-left border-collapse min-w-[680px]">
            <thead>
              <tr className="bg-slate-100/80 dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700">
                <th className="px-5 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase w-20 text-center">
                  Tháng
                </th>
                <th className="px-5 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase w-28">
                  Quý
                </th>
                <th className="px-5 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase w-72 text-right">
                  Số tiền doanh thu tháng (VNĐ)
                </th>
                <th className="px-5 py-3.5 text-xs font-bold text-slate-600 dark:text-slate-300 uppercase">
                  Ghi chú / Diễn giải
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200/80 dark:divide-slate-800">
              {activeReport.months.map((entry) => {
                const quarterLabel = `Quý ${Math.floor((entry.month - 1) / 3) + 1}`;
                return (
                  <tr
                    key={entry.month}
                    className="hover:bg-teal-50/30 dark:hover:bg-slate-800/40 transition-colors"
                  >
                    <td className="px-5 py-3 text-center">
                      <span className="inline-flex items-center justify-center min-w-[78px] px-2.5 py-1 rounded-lg bg-teal-50 dark:bg-teal-950/60 text-teal-800 dark:text-teal-300 font-bold text-xs tabular-nums">
                        Tháng {String(entry.month).padStart(2, '0')}
                      </span>
                    </td>
                    <td className="px-5 py-3 text-xs font-semibold text-slate-500 dark:text-slate-400">
                      {quarterLabel}
                    </td>
                    <td className="px-5 py-3">
                      <div className="relative">
                        <input
                          type="text"
                          inputMode="numeric"
                          value={entry.amount > 0 ? entry.amount.toLocaleString('vi-VN') : ''}
                          onChange={(e) => handleMonthAmountChange(entry.month, e.target.value)}
                          placeholder="0"
                          className="w-full min-h-[42px] pl-4 pr-12 py-2 rounded-xl border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-right font-mono font-bold text-sm text-slate-900 dark:text-white tabular-nums outline-none focus:border-teal-600 focus:ring-2 focus:ring-teal-500/15"
                        />
                        <span className="absolute right-3.5 top-1/2 -translate-y-1/2 text-xs font-bold text-slate-400 pointer-events-none">
                          VNĐ
                        </span>
                      </div>
                    </td>
                    <td className="px-5 py-3">
                      <input
                        type="text"
                        value={entry.note || ''}
                        onChange={(e) => handleMonthNoteChange(entry.month, e.target.value)}
                        placeholder={`Ghi chú doanh thu tháng ${String(entry.month).padStart(2, '0')}...`}
                        className="w-full min-h-[42px] px-3.5 py-2 rounded-xl border border-slate-200 dark:border-slate-700 bg-slate-50/70 dark:bg-slate-800/60 text-sm font-medium text-slate-800 dark:text-slate-200 outline-none focus:border-teal-600"
                      />
                    </td>
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              {/* Prominent Total Annual Revenue Footer */}
              <tr className="bg-gradient-to-r from-emerald-800 via-teal-700 to-teal-600 text-white">
                <td colSpan={2} className="px-5 py-5">
                  <div className="text-sm sm:text-base font-extrabold uppercase tracking-wide">
                    TỔNG DOANH THU CẢ NĂM {selectedYear}
                  </div>
                  <div className="text-xs text-teal-100 mt-0.5">
                    Tổng cộng doanh thu 12 tháng kê khai thuế hộ kinh doanh
                  </div>
                </td>
                <td className="px-5 py-5 text-right">
                  <div className="text-xl sm:text-2xl font-mono font-extrabold text-amber-300 tabular-nums">
                    {totalAnnualRevenue.toLocaleString('vi-VN')} VNĐ
                  </div>
                </td>
                <td className="px-5 py-5">
                  <div className="text-xs sm:text-sm font-medium text-white/95 italic">
                    Bằng chữ: {readAmountInVietnamese(totalAnnualRevenue)}
                  </div>
                </td>
              </tr>
            </tfoot>
          </table>
        </div>

        {/* Quarterly Summary & Full-Year Export Action Bar */}
        <div className="p-6 bg-slate-50 dark:bg-slate-900 border-t border-slate-200 dark:border-slate-800 space-y-5">
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
            {quarterlyTotals.map((qAmt, idx) => (
              <div
                key={idx}
                className="p-3.5 rounded-xl bg-white dark:bg-slate-800/80 border border-slate-200/90 dark:border-slate-700"
              >
                <p className="text-xs font-bold text-slate-500 dark:text-slate-400">
                  Tổng Quý {idx + 1}/{selectedYear}
                </p>
                <p className="text-base font-mono font-extrabold text-slate-900 dark:text-white mt-1 tabular-nums">
                  {qAmt.toLocaleString('vi-VN')} đ
                </p>
              </div>
            ))}
          </div>

          <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-4 pt-2 border-t border-slate-200/80 dark:border-slate-800">
            <div className="text-xs sm:text-sm text-slate-600 dark:text-slate-300 font-medium">
              Sau khi kiểm tra số liệu 12 tháng, nhấn nút bên phải để xuất toàn bộ doanh thu trong năm phục vụ lưu trữ và kê khai thuế.
            </div>
            <div className="flex flex-wrap items-center gap-3 shrink-0">
              <button
                type="button"
                onClick={exportAnnualTaxExcel}
                className="btn-primary"
              >
                <FileSpreadsheet className="w-4 h-4" />
                <span>Xuất toàn bộ doanh thu năm (.xlsx)</span>
              </button>
              <button
                type="button"
                onClick={exportAnnualTaxWord}
                className="btn-secondary"
              >
                <FileText className="w-4 h-4 text-teal-600" />
                <span>Xuất báo cáo năm (.docx)</span>
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default TaxReportSection;
