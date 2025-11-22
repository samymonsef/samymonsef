import React, { useState, useCallback, ChangeEvent, useMemo, useRef, useEffect } from 'react';
import { ExcelSheetData, ExcelRow } from './types';

// Declare external libraries to inform TypeScript that they're available globally from the CDN script
declare var XLSX: any;
declare var html2pdf: any;
declare var docx: any;
declare var Chart: any;
declare var LZString: any;

//#region Icons
const UploadIcon: React.FC<{ className?: string }> = ({ className }) => (
  <svg className={className} stroke="currentColor" fill="none" strokeWidth="2" viewBox="0 0 24 24" strokeLinecap="round" strokeLinejoin="round" xmlns="http://www.w3.org/2000/svg">
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
    <polyline points="17 8 12 3 7 8"></polyline>
    <line x1="12" y1="3" x2="12" y2="15"></line>
  </svg>
);

const ClearIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" d="M9.75 9.75l4.5 4.5m0-4.5l-4.5 4.5M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
    </svg>
);

const TableIcon: React.FC<{ className?: string }> = ({ className }) => (
  <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
    <path strokeLinecap="round" strokeLinejoin="round" d="M3.375 19.5h17.25m-17.25 0a1.125 1.125 0 01-1.125-1.125v-1.5c0-.621.504-1.125 1.125-1.125H6.75m10.5 0h1.125c.621 0 1.125.504 1.125 1.125v1.5c0 .621-.504 1.125-1.125 1.125m-17.25 0h.008v.016m0 0h17.248m-17.25 0h17.25m-17.25 0v-2.25H6.75m10.5 0v-2.25h2.625m0 0H6.75m10.5 0H6.75m0 0v-2.25h10.5m0 0v-2.25h2.625m0 0H6.75m10.5 0H6.75m0 0V6.75h10.5m0 0V4.5h2.625m0 0H6.75" />
  </svg>
);

const DashboardIcon: React.FC<{ className?: string }> = ({ className }) => (
  <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
    <path strokeLinecap="round" strokeLinejoin="round" d="M3.75 6A2.25 2.25 0 016 3.75h2.25A2.25 2.25 0 0110.5 6v2.25a2.25 2.25 0 01-2.25 2.25H6a2.25 2.25 0 01-2.25-2.25V6zM3.75 15.75A2.25 2.25 0 016 13.5h2.25a2.25 2.25 0 012.25 2.25V18a2.25 2.25 0 01-2.25 2.25H6A2.25 2.25 0 013.75 18v-2.25zM13.5 6a2.25 2.25 0 012.25-2.25H18A2.25 2.25 0 0120.25 6v2.25A2.25 2.25 0 0118 10.5h-2.25a2.25 2.25 0 01-2.25-2.25V6zM13.5 15.75a2.25 2.25 0 012.25-2.25H18a2.25 2.25 0 012.25 2.25V18A2.25 2.25 0 0118 20.25h-2.25A2.25 2.25 0 0113.5 18v-2.25z" />
  </svg>
);

const InfoIcon: React.FC<{ className?: string }> = ({ className }) => (
  <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
    <path strokeLinecap="round" strokeLinejoin="round" d="M11.25 11.25l.041-.02a.75.75 0 011.063.852l-.708 2.836a.75.75 0 001.063.853l.041-.021M21 12a9 9 0 11-18 0 9 9 0 0118 0zm-9-3.75h.008v.008H12V8.25z" />
  </svg>
);

const ExportIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 16.5m0 0L7.5 12m4.5 4.5V3" />
    </svg>
);

const FilePdfIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
    </svg>
);

const ShareIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" d="M7.217 10.907a2.25 2.25 0 100 2.186m0-2.186c.18.324.283.696.283 1.093s-.103.77-.283 1.093m0-2.186l9.566-5.314m-9.566 7.5l9.566 5.314m0 0a2.25 2.25 0 103.935 2.186 2.25 2.25 0 00-3.935-2.186zm0-12.814a2.25 2.25 0 103.933-2.185 2.25 2.25 0 00-3.933 2.185z" />
    </svg>
);

const PestIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" d="M12 3v17.25m0 0c-1.472 0-2.882.265-4.185.75M12 20.25c1.472 0 2.882.265 4.185.75M18.75 4.97A48.416 48.416 0 0012 4.5c-2.291 0-4.545.16-6.75.47m13.5 0c1.01.143 2.01.317 3 .52m-3-.52l2.62 10.726c.122.499-.106 1.028-.589 1.202a5.988 5.988 0 01-2.031.352 5.988 5.988 0 01-2.031-.352c-.483-.174-.711-.703-.59-1.202L18.75 4.971zm-16.5.52c.99-.203 1.99-.377 3-.52m0 0l2.62 10.726c.122.499-.106 1.028-.589 1.202a5.989 5.989 0 01-2.031.352 5.989 5.989 0 01-2.031-.352c-.483-.174-.711-.703-.59-1.202L5.25 4.971z" />
    </svg>
);

const ListIcon: React.FC<{ className?: string }> = ({ className }) => (
    <svg className={className} xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" d="M8.25 6.75h12M8.25 12h12m-12 5.25h12M3.75 6.75h.007v.008H3.75V6.75zm.375 0a.375.375 0 11-.75 0 .375.375 0 01.75 0zM3.75 12h.007v.008H3.75V12zm.375 0a.375.375 0 11-.75 0 .375.375 0 01.75 0zm-.375 5.25h.007v.008H3.75v-.008zm.375 0a.375.375 0 11-.75 0 .375.375 0 01.75 0z" />
    </svg>
);

//#endregion

//#region Type Definitions

interface DashboardData {
  summary: {
    totalRecommendations: number;
    totalRaw: number;
    totalFormulations: number;
  };
  formulationsByType: { type: string; count: number }[];
  formulationsByToxicity: { toxicity: string; count: number }[];
  rawMaterialsByCompany: { company: string; country: string; count: number }[];
  rawMaterialsByCountry: { country: string; count: number }[];
  rawMaterialsByAI: { activeIngredient: string; count: number }[];
  recommendationsByCrop: { crop: string; count: number }[];
  recommendationsByCropAndPest: { crop: string; pests: { pest: string; count: number }[] }[];
  columnMapping: { [key: string]: string };
}

type ViewMode = 'upload' | 'dashboard' | 'table';
//#endregion

//#region Chart Component
interface PieChartProps {
    chartId: string;
    data: { label: string; value: number; color?: string }[];
    title: string;
    onSliceClick?: (label: string) => void;
}

const PieChart: React.FC<PieChartProps> = ({ chartId, data, title, onSliceClick }) => {
    const chartRef = useRef<HTMLCanvasElement>(null);
    const chartInstanceRef = useRef<any>(null);

    useEffect(() => {
        if (!chartRef.current || typeof Chart === 'undefined') return;
        const ctx = chartRef.current.getContext('2d');
        if (!ctx) return;

        if (chartInstanceRef.current) {
            chartInstanceRef.current.destroy();
        }

        const isDarkMode = document.documentElement.classList.contains('dark');
        const textColor = isDarkMode ? 'rgba(229, 231, 235, 0.9)' : 'rgba(55, 65, 81, 0.9)';
        
        const chartColors = data.every(d => d.color) 
            ? data.map(d => d.color!)
            : [
                '#4f46e5', '#059669', '#f59e0b', '#e11d48', '#3b82f6', '#8b5cf6',
                '#db2777', '#6366f1', '#d946ef', '#06b6d4', '#f97316', '#65a30d'
            ];

        chartInstanceRef.current = new Chart(ctx, {
            type: 'pie',
            data: {
                labels: data.map(d => d.label),
                datasets: [{
                    label: title,
                    data: data.map(d => d.value),
                    backgroundColor: chartColors,
                    borderColor: isDarkMode ? '#1f2937' : '#ffffff',
                    borderWidth: 2,
                    hoverOffset: 4,
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                onHover: (event: any, elements: any[]) => {
                    if (event.native && event.native.target) {
                        event.native.target.style.cursor = elements.length ? 'pointer' : 'default';
                    }
                },
                onClick: (event: any, elements: any[]) => {
                    if (elements.length > 0 && onSliceClick) {
                        const index = elements[0].index;
                        const label = data[index].label;
                        onSliceClick(label);
                    }
                },
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            color: textColor,
                            padding: 15,
                            font: {
                                family: 'inherit',
                                size: 12,
                            }
                        }
                    },
                    tooltip: {
                        rtl: true,
                        titleFont: { family: 'inherit' },
                        bodyFont: { family: 'inherit' },
                        backgroundColor: isDarkMode ? '#374151' : '#fff',
                        titleColor: textColor,
                        bodyColor: textColor,
                        borderColor: isDarkMode ? '#4b5563' : '#e5e7eb',
                        borderWidth: 1,
                        padding: 10,
                        callbacks: {
                            label: function(context: any) {
                                let label = context.dataset.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                if (context.parsed !== null) {
                                    label += context.label + ' - ' + context.parsed.toLocaleString();
                                }
                                return label;
                            }
                        }
                    }
                }
            }
        });

        return () => {
            if (chartInstanceRef.current) {
                chartInstanceRef.current.destroy();
            }
        };
    }, [data, title, chartId, onSliceClick]);

    return <div className="relative h-72 w-full"><canvas ref={chartRef} id={chartId}></canvas></div>;
};

//#endregion

//#region Worker Code
const workerScript = `
importScripts("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");

self.onmessage = function(e) {
    try {
        const { fileData } = e.data;
        
        if (!XLSX) {
             throw new Error("فشل تحميل مكتبة Excel. تأكد من الاتصال بالإنترنت.");
        }

        const workbook = XLSX.read(fileData, { type: 'array' });
        const sheetNames = workbook.SheetNames;
        
        if (sheetNames.length === 0) throw new Error("الملف فارغ");

        const allSheetsData = sheetNames.map(name => 
            XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: "" })
        );
        
        self.postMessage({ success: true, sheetNames, allSheetsData });
    } catch (error) {
        self.postMessage({ success: false, error: error.message || "خطأ غير معروف أثناء معالجة الملف" });
    }
};
`;
//#endregion

export const App: React.FC = () => {
  const [allSheetsData, setAllSheetsData] = useState<ExcelSheetData[]>([]);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [activeSheetIndex, setActiveSheetIndex] = useState<number>(0);
  const [error, setError] = useState<string | null>(null);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [notification, setNotification] = useState<{type: 'success' | 'info', message: string} | null>(null);
  
  const [viewMode, setViewMode] = useState<ViewMode>('upload');
  const [dashboardData, setDashboardData] = useState<DashboardData | null>(null);
  const [drillDownData, setDrillDownData] = useState<{title: string, rows: ExcelRow[], columns: string[]} | null>(null);
  const [pestModalData, setPestModalData] = useState<{crop: string, pests: {pest: string, count: number}[]} | null>(null);
  const [isSharedMode, setIsSharedMode] = useState(false);
  const [showAllCrops, setShowAllCrops] = useState(false);
  const [showToxicityTable, setShowToxicityTable] = useState(false);

  const LOCAL_STORAGE_KEY = 'pesticideAnalysisData_v1';

  // Initial load
  useEffect(() => {
      const init = async () => {
          // 1. Check URL for shared data
          const hash = window.location.hash;
          if (hash.startsWith('#share=')) {
              try {
                  if (typeof LZString === 'undefined') {
                      throw new Error("مكتبة الضغط غير متوفرة");
                  }
                  const compressed = hash.substring(7);
                  const decompressed = LZString.decompressFromEncodedURIComponent(compressed);
                  if (decompressed) {
                      const sharedState = JSON.parse(decompressed);
                      if (sharedState.dashboardData) {
                          setDashboardData(sharedState.dashboardData);
                          setFileName(sharedState.fileName || 'Shared Dashboard');
                          setViewMode('dashboard');
                          setIsSharedMode(true);
                          setAllSheetsData([]); // Shared links usually don't carry full data
                          setSheetNames([]);
                          setNotification({ type: 'success', message: 'تم تحميل لوحة المعلومات المشتركة' });
                          setIsLoading(false);
                          return;
                      }
                  }
              } catch (e) {
                  console.warn("Failed to load shared data", e);
              }
          }

          // 2. Check LocalStorage
          try {
              const savedStateJSON = localStorage.getItem(LOCAL_STORAGE_KEY);
              if (savedStateJSON) {
                  const savedState = JSON.parse(savedStateJSON);
                  if (savedState.fileName && savedState.dashboardData) {
                      setFileName(savedState.fileName);
                      setSheetNames(savedState.sheetNames || []);
                      setAllSheetsData(savedState.allSheetsData || []);
                      setDashboardData(savedState.dashboardData);
                      setViewMode('dashboard');
                  }
              }
          } catch (e) {
              console.warn("LocalStorage error", e);
          } finally {
              setIsLoading(false);
          }
      };

      // Small delay to ensure CDNs (LZString) might be ready
      setTimeout(init, 100);
  }, []);

  useEffect(() => {
    if (notification) {
      const timer = setTimeout(() => setNotification(null), 3000);
      return () => clearTimeout(timer);
    }
  }, [notification]);

  // Auto-save to LocalStorage
  useEffect(() => {
    if (viewMode === 'dashboard' && dashboardData && fileName && !isSharedMode) {
        const stateToSave = {
            fileName,
            sheetNames,
            allSheetsData: allSheetsData.length < 10000 ? allSheetsData : [], // Don't save massive datasets
            dashboardData,
        };
        try {
            localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(stateToSave));
            setSaveError(null);
        } catch (err) {
            console.warn("LocalStorage full", err);
            setSaveError("مساحة التخزين ممتلئة. لن يتم حفظ البيانات تلقائياً.");
        }
    }
  }, [dashboardData, fileName, sheetNames, allSheetsData, viewMode, isSharedMode]);


  const resetState = () => {
    setAllSheetsData([]);
    setSheetNames([]);
    setActiveSheetIndex(0);
    setError(null);
    setSaveError(null);
    setFileName(null);
    setIsLoading(false);
    setViewMode('upload');
    setDashboardData(null);
    setNotification(null);
    setDrillDownData(null);
    setPestModalData(null);
    setIsSharedMode(false);
    setShowAllCrops(false);
    setShowToxicityTable(false);
    localStorage.removeItem(LOCAL_STORAGE_KEY);
    window.history.pushState("", document.title, window.location.pathname + window.location.search);
  };
  
  const runAnalysis = useCallback((sheetsData: ExcelSheetData[], availableSheetNames: string[]) => {
    try {
        const targetSheetName = 'المبيدات';
        // Case insensitive search for the sheet
        const targetSheetIndex = availableSheetNames.findIndex(name => name.trim().toLowerCase() === targetSheetName.toLowerCase() || name.includes('المبيدات'));

        if (targetSheetIndex === -1) {
            throw new Error(`لم يتم العثور على الشيت المطلوب باسم "${targetSheetName}".`);
        }

        const sheetData = sheetsData[targetSheetIndex];
        if (!sheetData || sheetData.length === 0) {
            throw new Error(`شيت "${targetSheetName}" فارغ.`);
        }

        const headers = Object.keys(sheetData[0]);
        const findHeader = (possibleNames: string[]): string => {
            for (const name of possibleNames) {
                const found = headers.find(h => h.trim().toLowerCase() === name.toLowerCase());
                if (found) return found;
            }
            return possibleNames[0];
        };
        
        const headerMap = {
            status: findHeader(['موقف المركب من التسجيل', 'Product Status AR', 'حالة المركب', 'Status']),
            form: findHeader(['الصورة', 'Formulation', 'formulation', 'Type']),
            tradeName: findHeader(['إسم المستحضر التجاري', 'الإسم التجاري', 'الإسم التجارى', 'اسم المستحضر التجاري', 'الاسم التجاري', 'Trade Name', 'TradeName']),
            company: findHeader(['الشركة المنتجة', 'Company', 'Supplier']),
            country: findHeader(['البلد', 'Country', 'Origin']),
            activeIngredient: findHeader(['المادة الفعالة', 'الإسم الشائع', 'Active Ingredient', 'Common Name']),
            pesticideType: findHeader(['تصنيف المبيد', 'نوع المبيد', 'group', 'classification', 'المجموعة', 'Category']),
            toxicity: findHeader(['سمية المستحضر', 'السمية', 'Toxicity']),
            crop: findHeader(['المحصول', 'Crop']),
            pest: findHeader(['الآفة', 'Pest', 'Common Name of Pest']),
            rate: findHeader(['معدل الاستخدام', 'معدل الإستخدام', 'Rate', 'Dosage']),
            unit: findHeader(['وحدة الاستخدام', 'وحدة الإستخدام', 'Unit']),
        };

        const registeredPesticides = sheetData.filter(row => String(row[headerMap.status] || '').trim() !== 'موقوف');

        const formulations = new Set<string>();
        const rawMaterials = new Set<string>();
        const formulationsByType = new Map<string, number>();
        const uniqueFormulationToxicity = new Map<string, string>();
        
        const rawMaterialsByCompanyData = new Map<string, { country: string; materials: Set<string> }>();
        const rawMaterialsByAI = new Map<string, number>();
        
        const recommendationsByCrop = new Map<string, number>();
        const pestsByCrop = new Map<string, Map<string, number>>();
        
        let totalRecommendationsCount = 0;

        for (const row of registeredPesticides) {
            const tradeName = String(row[headerMap.tradeName] || '').trim();
            const isRaw = String(row[headerMap.form] || '').trim().toLowerCase() === 'tech';

            if (tradeName) {
                if (isRaw) {
                    rawMaterials.add(tradeName);
                    const company = String(row[headerMap.company] || 'غير معروف').trim();
                    const country = String(row[headerMap.country] || 'غير محدد').trim();
                    const activeIngredient = String(row[headerMap.activeIngredient] || 'غير معروف').trim();
                    
                    if (!rawMaterialsByCompanyData.has(company)) {
                        rawMaterialsByCompanyData.set(company, { country: country, materials: new Set() });
                    }
                    rawMaterialsByCompanyData.get(company)!.materials.add(tradeName);
                    
                    rawMaterialsByAI.set(activeIngredient, (rawMaterialsByAI.get(activeIngredient) || 0) + 1);

                } else {
                    formulations.add(tradeName);
                    const pesticideType = String(row[headerMap.pesticideType] || 'غير مصنف').trim();
                    if (pesticideType) {
                        formulationsByType.set(pesticideType, (formulationsByType.get(pesticideType) || 0) + 1);
                    }
                    
                    const toxicity = String(row[headerMap.toxicity] || '').trim();
                    
                    // Ensure we record unique toxicity for each tradeName
                    // Normalize Toxicity to the 6 requested categories
                    if (!uniqueFormulationToxicity.has(tradeName)) {
                        let standardizedToxicity = 'غير مصنف';
                        const t = toxicity.toLowerCase();
                        
                        if (t.includes('ia') || t.includes('1a')) {
                            standardizedToxicity = 'High Ia';
                        } else if (t.includes('ib') || t.includes('1b')) {
                            standardizedToxicity = 'High Ib';
                        } else if (t.includes('iii') || t.includes('3')) { 
                            // Check III before II because III contains II string
                            standardizedToxicity = 'Low III';
                        } else if (t.includes('mod') || (t.includes('ii') && !t.includes('low'))) {
                            standardizedToxicity = 'Mod II';
                        } else if (t.includes('low') && t.includes('ii')) {
                            standardizedToxicity = 'Low II';
                        } else if (t.includes('u') || t.includes('unlikely') || t.includes('iv')) {
                            standardizedToxicity = 'U';
                        }
                        
                        if (standardizedToxicity !== 'غير مصنف') {
                            uniqueFormulationToxicity.set(tradeName, standardizedToxicity);
                        }
                    }
                }
            }
            
            if (!isRaw) {
                const crop = String(row[headerMap.crop] || '').trim();
                const pest = String(row[headerMap.pest] || '').trim();

                if (crop && pest) {
                    totalRecommendationsCount++;
                    recommendationsByCrop.set(crop, (recommendationsByCrop.get(crop) || 0) + 1);
                    
                    if (!pestsByCrop.has(crop)) {
                        pestsByCrop.set(crop, new Map());
                    }
                    const pestMap = pestsByCrop.get(crop)!;
                    pestMap.set(pest, (pestMap.get(pest) || 0) + 1);
                }
            }
        }
        
        const processCountMap = (map: Map<string, number>, limit = 15) => {
            const sorted = Array.from(map.entries())
              .map(([key, count]) => ({ key, count }))
              .sort((a, b) => b.count - a.count);
              
            if (sorted.length <= limit) return sorted;
            
            const top = sorted.slice(0, limit);
            const otherCount = sorted.slice(limit).reduce((sum, item) => sum + item.count, 0);
            if (otherCount > 0) top.push({ key: 'أخرى', count: otherCount });
            return top;
        };

        const rawMaterialsByCountry = new Map<string, number>();
        for (const data of rawMaterialsByCompanyData.values()) {
            rawMaterialsByCountry.set(data.country, (rawMaterialsByCountry.get(data.country) || 0) + data.materials.size);
        }

        const processedRawMaterialsByCompany = Array.from(rawMaterialsByCompanyData.entries())
            .map(([company, data]) => ({
                company,
                country: data.country,
                count: data.materials.size,
            }))
            .sort((a, b) => b.count - a.count)
            .slice(0, 20);
        
        // Sort crops by count
        const sortedCrops = Array.from(recommendationsByCrop.entries())
             .sort((a, b) => b[1] - a[1]);

        const formattedPestsByCrop = Array.from(pestsByCrop.entries()).map(([crop, pestMap]) => ({
            crop,
            pests: Array.from(pestMap.entries()).map(([pest, count]) => ({ pest, count })).sort((a, b) => b.count - a.count)
        }));

        // Custom sort for Toxicity
        // Expected order: High Ia -> High Ib -> Mod II -> Low II -> Low III -> U
        const toxicityOrder = ['High Ia', 'High Ib', 'Mod II', 'Low II', 'Low III', 'U'];
        
        // Calculate formulation counts from unique map
        const formulationsByToxicityMap = new Map<string, number>();
        // Initialize with 0 for all keys to ensure they appear in order
        toxicityOrder.forEach(key => formulationsByToxicityMap.set(key, 0));

        for (const t of uniqueFormulationToxicity.values()) {
             if (formulationsByToxicityMap.has(t)) {
                 formulationsByToxicityMap.set(t, (formulationsByToxicityMap.get(t) || 0) + 1);
             }
        }

        // Create sorted array based on fixed order, filtering out zero counts if desired (or keeping them if strict structure needed)
        const sortedToxicity = toxicityOrder
            .map(toxicity => ({ toxicity, count: formulationsByToxicityMap.get(toxicity) || 0 }))
            .filter(item => item.count > 0); // Only show categories that have data

        const finalData: DashboardData = {
            summary: {
                totalRecommendations: totalRecommendationsCount,
                totalRaw: rawMaterials.size,
                totalFormulations: formulations.size,
            },
            formulationsByType: processCountMap(formulationsByType).map(item => ({type: item.key, count: item.count})),
            formulationsByToxicity: sortedToxicity,
            rawMaterialsByCompany: processedRawMaterialsByCompany,
            rawMaterialsByCountry: processCountMap(rawMaterialsByCountry).map(item => ({ country: item.key, count: item.count })),
            rawMaterialsByAI: processCountMap(rawMaterialsByAI).map(item => ({activeIngredient: item.key, count: item.count})),
            recommendationsByCrop: sortedCrops.map(([crop, count]) => ({crop, count})),
            recommendationsByCropAndPest: formattedPestsByCrop,
            columnMapping: headerMap,
        };
        
        setDashboardData(finalData);
        setViewMode('dashboard');

    } catch(err) {
        throw err;
    }
  }, []);

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    setError(null);
    setSaveError(null);
    setIsSharedMode(false);
    setDashboardData(null);
    setAllSheetsData([]); 
    setIsLoading(true);
    
    // Reset input
    event.target.value = '';

    const reader = new FileReader();

    reader.onload = (e) => {
        const data = e.target?.result;
        if (!data) {
            setError("فشل قراءة الملف");
            setIsLoading(false);
            return;
        }

        // Create worker from string
        const blob = new Blob([workerScript], { type: 'application/javascript' });
        const worker = new Worker(URL.createObjectURL(blob));

        worker.onmessage = (message) => {
            const { success, sheetNames: parsedSheetNames, allSheetsData: parsedData, error: workerError } = message.data;

            if (success) {
                setSheetNames(parsedSheetNames);
                setAllSheetsData(parsedData);
                try {
                    runAnalysis(parsedData, parsedSheetNames);
                } catch (analysisErr: any) {
                    setError(analysisErr.message || "خطأ في تحليل البيانات");
                    setDashboardData(null);
                }
            } else {
                setError(workerError || "فشل معالجة الملف في الخلفية");
            }
            setIsLoading(false);
            worker.terminate();
        };

        worker.onerror = (err) => {
            console.error("Worker Error:", err);
            setError("حدث خطأ غير متوقع أثناء المعالجة.");
            setIsLoading(false);
            worker.terminate();
        };

        // Send data to worker
        worker.postMessage({ fileData: data });
    };

    reader.onerror = () => {
        setError("فشل قراءة الملف من القرص الصلب.");
        setIsLoading(false);
        setFileName(null);
    };

    reader.readAsArrayBuffer(file);
  };

  const exportToCSV = () => {
      if (!allSheetsData.length || activeSheetIndex < 0) return;
      const currentData = allSheetsData[activeSheetIndex];
      if (!currentData || currentData.length === 0) return;

      const worksheet = XLSX.utils.json_to_sheet(currentData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
      const baseName = fileName ? fileName.replace(/\.[^/.]+$/, "") : "export";
      const safeSheetName = (sheetNames[activeSheetIndex] || "sheet").replace(/[\\/?*:|"<>]/g, "_");
      XLSX.writeFile(workbook, `${baseName}_${safeSheetName}.csv`);
  };

  const exportToPDF = () => {
      if (!allSheetsData.length || activeSheetIndex < 0) return;
      const currentData = allSheetsData[activeSheetIndex];
      if (!currentData || currentData.length === 0) return;

      setIsLoading(true);

      setTimeout(() => {
          const content = document.createElement('div');
          content.setAttribute('dir', 'rtl');
          content.style.fontFamily = 'sans-serif';
          content.innerHTML = `
            <h2 style="text-align: center; margin-bottom: 20px;">${fileName || 'Report'} - ${sheetNames[activeSheetIndex] || ''}</h2>
            <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
                <thead>
                    <tr style="background-color: #f3f4f6;">
                        ${Object.keys(currentData[0]).map(header => 
                            `<th style="border: 1px solid #9ca3af; padding: 8px; text-align: right;">${header}</th>`
                        ).join('')}
                    </tr>
                </thead>
                <tbody>
                    ${currentData.slice(0, 500).map(row => `
                        <tr>
                            ${Object.values(row).map(val => 
                                `<td style="border: 1px solid #d1d5db; padding: 8px; text-align: right;">${val !== null && val !== undefined ? val : ''}</td>`
                            ).join('')}
                        </tr>
                    `).join('')}
                </tbody>
            </table>
            ${currentData.length > 500 ? '<p style="text-align:center; margin-top: 10px;">تم عرض أول 500 سجل فقط</p>' : ''}
          `;

          const opt = {
              margin: [10, 10, 10, 10],
              filename: `${fileName?.replace(/\.[^/.]+$/, "") || 'export'}_${sheetNames[activeSheetIndex] || 'sheet'}.pdf`,
              image: { type: 'jpeg', quality: 0.98 },
              html2canvas: { scale: 2, useCORS: true, logging: false },
              jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' }
          };

          html2pdf().from(content).set(opt).save()
              .then(() => setIsLoading(false))
              .catch((err: any) => {
                  console.error("PDF Export Error:", err);
                  setError("فشل تصدير PDF. قد يكون حجم البيانات كبيرًا جدًا.");
                  setIsLoading(false);
              });
      }, 100);
  };

  const handleDrillDown = (chartType: 'type' | 'toxicity', label: string) => {
      if (isSharedMode && (!allSheetsData || allSheetsData.length === 0)) {
          setNotification({ type: 'info', message: 'البيانات التفصيلية غير متوفرة في وضع المشاركة.' });
          return;
      }

      if (!dashboardData || !allSheetsData.length) return;
      
      // Try to find sheet case-insensitively or partially
      const targetSheetIndex = sheetNames.findIndex(name => name.trim().toLowerCase() === 'المبيدات' || name.includes('المبيدات'));
      const sheetData = allSheetsData[targetSheetIndex];
      
      if (!sheetData || sheetData.length === 0) return;
      
      const mapping = dashboardData.columnMapping;
      const columns = Object.keys(sheetData[0]);

      const filteredRows = sheetData.filter(row => {
           const status = String(row[mapping.status] || '').trim();
           if (status === 'موقوف') return false;

           const tradeName = String(row[mapping.tradeName] || '').trim();
           const isRaw = String(row[mapping.form] || '').trim().toLowerCase() === 'tech';
           
           if (!tradeName) return false;
           if (isRaw) return false;

           if (label === 'أخرى') return false;

           let rowValue = String(row[mapping[chartType === 'type' ? 'pesticideType' : 'toxicity']] || 'غير مصنف').trim();
           
           // Normalize logic for drilldown matching
           if (chartType === 'toxicity') {
               const t = rowValue.toLowerCase();
               let standardized = 'غير مصنف';
                if (t.includes('ia') || t.includes('1a')) standardized = 'High Ia';
                else if (t.includes('ib') || t.includes('1b')) standardized = 'High Ib';
                else if (t.includes('iii') || t.includes('3')) standardized = 'Low III';
                else if (t.includes('mod') || (t.includes('ii') && !t.includes('low'))) standardized = 'Mod II';
                else if (t.includes('low') && t.includes('ii')) standardized = 'Low II';
                else if (t.includes('u') || t.includes('unlikely') || t.includes('iv')) standardized = 'U';
                
               if (standardized !== label) return false;
           } else {
               if (rowValue !== label) return false;
           }
           return true;
      });
      
      const limitedRows = filteredRows.slice(0, 500);

      setDrillDownData({
          title: `${chartType === 'type' ? 'النوع' : 'السمية'}: ${label}`,
          rows: limitedRows,
          columns: columns
      });
      
      if (filteredRows.length > 500) {
          setNotification({ type: 'info', message: `تم عرض أول 500 نتيجة فقط من أصل ${filteredRows.length}` });
      }
  };

  const exportCropsReportToExcel = () => {
      if (!dashboardData) return;
      
      const total = dashboardData.summary.totalRecommendations;
      const data = dashboardData.recommendationsByCrop.map(item => ({
          'المحصول': item.crop,
          'عدد التوصيات': item.count,
          'النسبة': ((item.count / total) * 100).toFixed(1) + '%'
      }));

      const worksheet = XLSX.utils.json_to_sheet(data);
      const wscols = [
          { wch: 30 }, // Crop
          { wch: 15 }, // Count
          { wch: 10 }  // Percentage
      ];
      worksheet['!cols'] = wscols;

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "تقرير المحاصيل");
      XLSX.writeFile(workbook, `Crop_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const exportCropsReportToPDF = () => {
      if (!dashboardData) return;

      setIsLoading(true);
      const total = dashboardData.summary.totalRecommendations;

      setTimeout(() => {
          const content = document.createElement('div');
          content.setAttribute('dir', 'rtl');
          content.style.fontFamily = 'sans-serif';
          content.innerHTML = `
            <h2 style="text-align: center; margin-bottom: 20px;">تقرير توصيات المحاصيل</h2>
            <p style="text-align: center; margin-bottom: 10px; color: #666;">إجمالي التوصيات: ${total}</p>
            <table style="width: 100%; border-collapse: collapse; font-size: 12px; text-align: right;">
                <thead>
                    <tr style="background-color: #f3f4f6;">
                        <th style="border: 1px solid #9ca3af; padding: 8px; width: 50%;">المحصول</th>
                        <th style="border: 1px solid #9ca3af; padding: 8px; width: 25%;">عدد التوصيات</th>
                        <th style="border: 1px solid #9ca3af; padding: 8px; width: 25%;">النسبة</th>
                    </tr>
                </thead>
                <tbody>
                    ${dashboardData.recommendationsByCrop.map(item => `
                        <tr>
                            <td style="border: 1px solid #d1d5db; padding: 8px;">${item.crop}</td>
                            <td style="border: 1px solid #d1d5db; padding: 8px;">${item.count}</td>
                            <td style="border: 1px solid #d1d5db; padding: 8px;">${((item.count / total) * 100).toFixed(1)}%</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
          `;

          const opt = {
              margin: [10, 10, 10, 10],
              filename: `Crop_Report_${new Date().toISOString().split('T')[0]}.pdf`,
              image: { type: 'jpeg', quality: 0.98 },
              html2canvas: { scale: 2 },
              jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
          };

          html2pdf().from(content).set(opt).save()
              .then(() => setIsLoading(false))
              .catch((err: any) => {
                  console.error("PDF Export Error:", err);
                  setError("فشل تصدير PDF.");
                  setIsLoading(false);
              });
      }, 100);
  };

  const handleShowPests = (crop: string) => {
      if (!dashboardData) return;
      const cropData = dashboardData.recommendationsByCropAndPest.find(c => c.crop === crop);
      if (cropData) {
          setPestModalData({ crop: cropData.crop, pests: cropData.pests });
      }
  };

  const handleShare = async () => {
      if (!dashboardData) return;

      const shareObj = {
          fileName,
          dashboardData
      };

      try {
          if (typeof LZString === 'undefined') throw new Error('Compressor not ready');
          const jsonString = JSON.stringify(shareObj);
          const compressed = LZString.compressToEncodedURIComponent(jsonString);
          const url = `${window.location.origin}${window.location.pathname}#share=${compressed}`;

          await navigator.clipboard.writeText(url);
          setNotification({ type: 'success', message: 'تم نسخ رابط المشاركة!' });
      } catch (err) {
          console.error("Share error:", err);
          setNotification({ type: 'info', message: 'فشل إنشاء رابط المشاركة. البيانات كبيرة جدًا.' });
      }
  };

  const clearData = () => {
      resetState();
  };

  const getToxicityColor = (label: string) => {
      // Precise color matching based on user requirement
      switch (label) {
          case 'High Ia': return '#7f1d1d'; // Dark Red
          case 'High Ib': return '#f87171'; // Light Red
          case 'Mod II': return '#f59e0b';  // Yellow
          case 'Low II': return '#1e3a8a';  // Dark Blue (User Request)
          case 'Low III': return '#3b82f6'; // Light Blue (User Request: Distinct from Dark Blue)
          case 'U': return '#22c55e';       // Green
          default: return '#9ca3af';
      }
  };

  const exportToxicityReportToExcel = () => {
      if (!dashboardData) return;
      
      const total = dashboardData.summary.totalFormulations;
      const data = dashboardData.formulationsByToxicity.map(item => ({
          'تصنيف السمية': item.toxicity,
          'عدد المستحضرات': item.count,
          'النسبة': ((item.count / total) * 100).toFixed(2) + '%'
      }));

      // Add Total Row
      data.push({
          'تصنيف السمية': 'الإجمالي',
          'عدد المستحضرات': total,
          'النسبة': '100.00%'
      });

      const worksheet = XLSX.utils.json_to_sheet(data);
      const wscols = [
          { wch: 20 }, // Toxicity
          { wch: 15 }, // Count
          { wch: 10 }  // Percentage
      ];
      worksheet['!cols'] = wscols;

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "تقرير السمية");
      XLSX.writeFile(workbook, `Toxicity_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-gray-50 text-right font-sans" dir="rtl">
      <nav className="bg-white shadow-sm border-b border-gray-200 px-4 py-3 flex justify-between items-center sticky top-0 z-10">
        <div className="flex items-center">
            <h1 className="text-xl font-bold text-gray-800 ml-2">لوحة تحكم المبيدات</h1>
        </div>
        <div className="flex gap-2">
            {fileName && dashboardData && (
                <button 
                    onClick={handleShare} 
                    className="p-2 text-indigo-600 hover:bg-indigo-50 rounded-full transition-colors" 
                    title="مشاركة"
                >
                    <ShareIcon className="w-6 h-6" />
                </button>
            )}
            {fileName && (
                <button onClick={clearData} className="p-2 text-red-600 hover:bg-red-50 rounded-full transition-colors" title="مسح البيانات">
                    <ClearIcon className="w-6 h-6" />
                </button>
            )}
        </div>
      </nav>

      <main className="p-4 md:p-6 max-w-7xl mx-auto pb-20 md:pb-6">
        {notification && (
            <div className={`fixed bottom-4 left-1/2 transform -translate-x-1/2 z-50 px-6 py-3 rounded-lg shadow-lg text-white font-medium transition-opacity duration-300 ${notification.type === 'success' ? 'bg-green-600' : 'bg-blue-600'}`}>
                {notification.message}
            </div>
        )}

        {error && (
            <div className="bg-red-50 border-l-4 border-red-500 p-4 mb-6 rounded-r-lg animate-in fade-in slide-in-from-top-4 duration-300">
                <div className="flex">
                    <div className="flex-shrink-0 ml-3">
                        <svg className="h-5 w-5 text-red-400" viewBox="0 0 20 20" fill="currentColor">
                            <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
                        </svg>
                    </div>
                    <div className="flex-1">
                        <p className="text-sm text-red-700">{error}</p>
                    </div>
                </div>
            </div>
        )}

        {saveError && (
             <div className="bg-yellow-50 border-l-4 border-yellow-400 p-4 mb-6 rounded-r-lg">
                <div className="flex">
                    <div className="flex-shrink-0 ml-3">
                        <InfoIcon className="h-5 w-5 text-yellow-400" />
                    </div>
                    <div className="flex-1">
                        <p className="text-sm text-yellow-700">{saveError}</p>
                    </div>
                </div>
            </div>
        )}

        {viewMode === 'upload' && (
             <div className="flex flex-col items-center justify-center h-[60vh] border-2 border-dashed border-gray-300 rounded-xl bg-white hover:bg-gray-50 transition-all cursor-pointer relative group">
                <input 
                    type="file" 
                    accept=".xlsx, .xls" 
                    onChange={handleFileChange} 
                    className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                />
                <div className="text-center p-6">
                    <div className="bg-indigo-50 rounded-full p-4 inline-block mb-4 group-hover:bg-indigo-100 transition-colors">
                        <UploadIcon className="w-12 h-12 text-indigo-600" />
                    </div>
                    <h3 className="text-xl font-semibold text-gray-800 mb-2">رفع ملف Excel</h3>
                    <p className="text-gray-500 max-w-sm mx-auto">اضغط لاختيار ملف أو اسحبه هنا. تأكد من وجود شيت باسم "المبيدات".</p>
                </div>
            </div>
        )}

        {viewMode === 'dashboard' && dashboardData && (
            <div className="space-y-6 animate-fade-in">
                {isSharedMode && (
                    <div className="bg-blue-50 text-blue-700 px-4 py-2 rounded-lg border border-blue-100 flex items-center">
                        <InfoIcon className="w-5 h-5 ml-2" />
                        <span>أنت تشاهد نسخة للعرض فقط.</span>
                    </div>
                )}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
                        <h3 className="text-sm font-medium text-gray-500">عدد التوصيات</h3>
                        <p className="text-3xl font-bold text-indigo-600 mt-2">{dashboardData.summary.totalRecommendations.toLocaleString()}</p>
                    </div>
                    <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
                        <h3 className="text-sm font-medium text-gray-500">المستحضرات</h3>
                        <p className="text-3xl font-bold text-emerald-600 mt-2">{dashboardData.summary.totalFormulations.toLocaleString()}</p>
                    </div>
                    <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
                        <h3 className="text-sm font-medium text-gray-500">الخامات</h3>
                        <p className="text-3xl font-bold text-amber-600 mt-2">{dashboardData.summary.totalRaw.toLocaleString()}</p>
                    </div>
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                     <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100">
                        <h3 className="text-lg font-bold text-gray-800 mb-4">المستحضرات حسب النوع</h3>
                        <PieChart 
                            chartId="chart-type" 
                            title="النوع" 
                            data={dashboardData.formulationsByType.map(x => ({ label: x.type, value: x.count }))} 
                            onSliceClick={(label) => handleDrillDown('type', label)}
                        />
                     </div>
                     <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 relative">
                        <div className="flex justify-between items-start mb-4">
                            <h3 className="text-lg font-bold text-gray-800">المستحضرات حسب السمية</h3>
                             <button 
                                onClick={() => setShowToxicityTable(true)}
                                className="p-2 text-indigo-600 hover:bg-indigo-50 rounded-full transition-colors"
                                title="عرض الجدول"
                             >
                                 <ListIcon className="w-5 h-5" />
                             </button>
                        </div>
                        <PieChart 
                            chartId="chart-toxicity" 
                            title="السمية" 
                            data={dashboardData.formulationsByToxicity.map(x => ({ 
                                label: x.toxicity, 
                                value: x.count,
                                color: getToxicityColor(x.toxicity)
                            }))} 
                            onSliceClick={(label) => handleDrillDown('toxicity', label)}
                        />
                     </div>
                </div>
                
                <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
                    <div className="p-5 border-b border-gray-100 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                        <div>
                            <h3 className="text-lg font-bold text-gray-800">أكثر المحاصيل توصية</h3>
                            <p className="text-sm text-gray-500 mt-1">اضغط على أي محصول لعرض الآفات المسجلة عليه</p>
                        </div>
                        <div className="flex gap-2">
                             <button onClick={exportCropsReportToExcel} className="flex items-center gap-2 px-3 py-1.5 bg-white border border-gray-300 text-gray-700 rounded-md hover:bg-indigo-50 hover:text-indigo-700 transition-colors text-sm font-medium shadow-sm">
                                <ExportIcon className="w-4 h-4" />
                                <span>Excel</span>
                            </button>
                             <button onClick={exportCropsReportToPDF} className="flex items-center gap-2 px-3 py-1.5 bg-white border border-gray-300 text-gray-700 rounded-md hover:bg-red-50 hover:text-red-700 transition-colors text-sm font-medium shadow-sm">
                                <FilePdfIcon className="w-4 h-4" />
                                <span>PDF</span>
                            </button>
                        </div>
                    </div>
                    <div className="overflow-x-auto">
                        <table className="min-w-full divide-y divide-gray-200">
                            <thead className="bg-gray-50">
                                <tr>
                                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">المحصول</th>
                                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">عدد التوصيات</th>
                                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">النسبة</th>
                                </tr>
                            </thead>
                            <tbody className="bg-white divide-y divide-gray-200">
                                {dashboardData.recommendationsByCrop.slice(0, 20).map((item, idx) => (
                                    <tr 
                                        key={idx} 
                                        onClick={() => handleShowPests(item.crop)}
                                        className="hover:bg-indigo-50 cursor-pointer transition-colors"
                                        title="اضغط لعرض الآفات"
                                    >
                                        <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-indigo-600 underline decoration-dotted underline-offset-4">{item.crop}</td>
                                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{item.count}</td>
                                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                                            {((item.count / dashboardData.summary.totalRecommendations) * 100).toFixed(1)}%
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                    {dashboardData.recommendationsByCrop.length > 20 && (
                         <div 
                            onClick={() => setShowAllCrops(true)}
                            className="p-3 text-center border-t border-gray-100 bg-gray-50 hover:bg-indigo-50 cursor-pointer transition-colors"
                         >
                             <span className="text-sm text-gray-500 font-medium hover:text-indigo-600">...و {dashboardData.recommendationsByCrop.length - 20} محصول آخر (اضغط للعرض)</span>
                         </div>
                    )}
                </div>
            </div>
        )}

        {viewMode === 'table' && allSheetsData.length > 0 && (
            <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden flex flex-col h-[80vh]">
                <div className="p-4 border-b border-gray-200 flex flex-wrap justify-between items-center gap-4 bg-gray-50">
                    <h3 className="text-lg font-bold text-gray-800">استعراض البيانات</h3>
                    
                    <div className="flex items-center gap-3">
                        <button
                            onClick={exportToCSV}
                            className="flex items-center gap-2 px-3 py-1.5 bg-white border border-gray-300 text-gray-700 rounded-md hover:bg-gray-50 transition-colors text-sm font-medium shadow-sm"
                        >
                             <ExportIcon className="w-4 h-4" />
                             <span>CSV</span>
                        </button>

                        <button
                            onClick={exportToPDF}
                            className="flex items-center gap-2 px-3 py-1.5 bg-white border border-gray-300 text-gray-700 rounded-md hover:bg-gray-50 transition-colors text-sm font-medium shadow-sm"
                        >
                             <FilePdfIcon className="w-4 h-4 text-red-600" />
                             <span>PDF</span>
                        </button>
                        
                        <div className="h-6 w-px bg-gray-300 mx-1 hidden sm:block"></div>

                        <div className="flex items-center gap-2">
                            <span className="text-sm text-gray-500">الشيت:</span>
                            <select 
                                value={activeSheetIndex} 
                                onChange={(e) => setActiveSheetIndex(Number(e.target.value))}
                                className="border border-gray-300 rounded-md shadow-sm py-1.5 px-3 bg-white focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm"
                            >
                                {sheetNames.map((name, idx) => (
                                    <option key={idx} value={idx}>{name}</option>
                                ))}
                            </select>
                        </div>
                    </div>
                </div>
                <div className="overflow-auto flex-1">
                    <table className="min-w-full divide-y divide-gray-200">
                        <thead className="bg-gray-50 sticky top-0 z-10 shadow-sm">
                            <tr>
                                {Object.keys(allSheetsData[activeSheetIndex][0] || {}).map((header, idx) => (
                                    <th key={idx} className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider whitespace-nowrap bg-gray-50">
                                        {header}
                                    </th>
                                ))}
                            </tr>
                        </thead>
                        <tbody className="bg-white divide-y divide-gray-200">
                            {allSheetsData[activeSheetIndex].slice(0, 100).map((row, rIdx) => (
                                <tr key={rIdx} className="hover:bg-gray-50">
                                    {Object.values(row).map((val, cIdx) => (
                                        <td key={cIdx} className="px-6 py-4 whitespace-nowrap text-sm text-gray-500 border-b border-gray-100">
                                            {String(val)}
                                        </td>
                                    ))}
                                </tr>
                            ))}
                        </tbody>
                    </table>
                    {allSheetsData[activeSheetIndex].length > 100 && (
                        <div className="p-4 text-center text-gray-500 bg-gray-50 border-t">
                            يتم عرض أول 100 صف فقط للأداء.
                        </div>
                    )}
                </div>
            </div>
        )}
      </main>

      {/* Toxicity Table Modal */}
      {showToxicityTable && dashboardData && (
          <div className="fixed inset-0 z-50 overflow-y-auto bg-gray-900/75 backdrop-blur-sm flex items-center justify-center p-4" onClick={() => setShowToxicityTable(false)}>
              <div className="bg-white rounded-xl shadow-2xl w-full max-w-2xl max-h-[80vh] flex flex-col animate-in fade-in zoom-in duration-200" onClick={e => e.stopPropagation()}>
                  <div className="p-5 border-b border-gray-100 flex justify-between items-center bg-gray-50">
                      <h3 className="text-lg font-bold text-gray-800">تقرير السمية التفصيلي</h3>
                      <div className="flex items-center gap-2">
                          <button 
                              onClick={exportToxicityReportToExcel}
                              className="flex items-center gap-2 px-3 py-1.5 bg-white border border-gray-300 text-gray-700 rounded-md hover:bg-green-50 hover:text-green-700 transition-colors text-sm font-medium shadow-sm"
                              title="تصدير إلى Excel"
                          >
                              <ExportIcon className="w-4 h-4" />
                              <span>Excel</span>
                          </button>
                          <button onClick={() => setShowToxicityTable(false)} className="text-gray-400 hover:text-gray-600 p-2 hover:bg-gray-100 rounded-full transition-colors">
                              <ClearIcon className="w-5 h-5" />
                          </button>
                      </div>
                  </div>
                  <div className="overflow-y-auto flex-1 p-1">
                      <table className="min-w-full divide-y divide-gray-200">
                          <thead className="bg-gray-50 sticky top-0">
                              <tr>
                                  <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">الصنف (Toxicity)</th>
                                  <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">عدد المستحضرات</th>
                                  <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">النسبة</th>
                              </tr>
                          </thead>
                          <tbody className="bg-white divide-y divide-gray-100">
                              {dashboardData.formulationsByToxicity.map((item, idx) => (
                                  <tr key={idx} className="hover:bg-gray-50">
                                      <td className="px-6 py-3 whitespace-nowrap text-sm font-medium text-gray-800 flex items-center gap-2">
                                          <span className="w-3 h-3 rounded-full" style={{backgroundColor: getToxicityColor(item.toxicity)}}></span>
                                          {item.toxicity}
                                      </td>
                                      <td className="px-6 py-3 whitespace-nowrap text-sm text-gray-600">{item.count}</td>
                                      <td className="px-6 py-3 whitespace-nowrap text-sm text-gray-600">
                                          {((item.count / dashboardData.summary.totalFormulations) * 100).toFixed(2)}%
                                      </td>
                                  </tr>
                              ))}
                          </tbody>
                          <tfoot className="bg-gray-100 font-bold">
                              <tr>
                                  <td className="px-6 py-3 text-sm text-gray-800">الإجمالي (المستحضرات المسجلة)</td>
                                  <td className="px-6 py-3 text-sm text-gray-800">
                                      {dashboardData.summary.totalFormulations}
                                  </td>
                                  <td className="px-6 py-3 text-sm text-gray-800">100.00%</td>
                              </tr>
                          </tfoot>
                      </table>
                  </div>
              </div>
          </div>
      )}

      {/* All Crops Modal */}
      {showAllCrops && dashboardData && (
          <div className="fixed inset-0 z-50 overflow-y-auto bg-gray-900/75 backdrop-blur-sm flex items-center justify-center p-4" onClick={() => setShowAllCrops(false)}>
              <div className="bg-white rounded-xl shadow-2xl w-full max-w-2xl max-h-[80vh] flex flex-col animate-in fade-in zoom-in duration-200" onClick={e => e.stopPropagation()}>
                  <div className="p-5 border-b border-gray-100 flex justify-between items-center bg-gray-50">
                      <h3 className="text-lg font-bold text-gray-800">باقي المحاصيل ({dashboardData.recommendationsByCrop.length - 20})</h3>
                      <button onClick={() => setShowAllCrops(false)} className="text-gray-400 hover:text-gray-600 p-2 hover:bg-gray-100 rounded-full transition-colors">
                          <ClearIcon className="w-5 h-5" />
                      </button>
                  </div>
                  <div className="overflow-y-auto flex-1">
                      <table className="min-w-full divide-y divide-gray-200">
                          <thead className="bg-gray-50 sticky top-0">
                              <tr>
                                  <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">المحصول</th>
                                  <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">عدد التوصيات</th>
                                  <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">النسبة</th>
                              </tr>
                          </thead>
                          <tbody className="bg-white divide-y divide-gray-100">
                              {dashboardData.recommendationsByCrop.slice(20).map((item, idx) => (
                                  <tr 
                                      key={idx} 
                                      onClick={() => {
                                          setShowAllCrops(false);
                                          handleShowPests(item.crop);
                                      }}
                                      className="hover:bg-indigo-50 cursor-pointer transition-colors"
                                  >
                                      <td className="px-6 py-3 whitespace-nowrap text-sm font-medium text-indigo-600 underline decoration-dotted underline-offset-4">{item.crop}</td>
                                      <td className="px-6 py-3 whitespace-nowrap text-sm text-gray-500">{item.count}</td>
                                      <td className="px-6 py-3 whitespace-nowrap text-sm text-gray-500">
                                          {((item.count / dashboardData.summary.totalRecommendations) * 100).toFixed(1)}%
                                      </td>
                                  </tr>
                              ))}
                          </tbody>
                      </table>
                  </div>
              </div>
          </div>
      )}

      {/* Pest Breakdown Modal */}
      {pestModalData && (
          <div className="fixed inset-0 z-50 overflow-y-auto bg-gray-900/75 backdrop-blur-sm flex items-center justify-center p-4" onClick={() => setPestModalData(null)}>
              <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg animate-in fade-in zoom-in duration-200 overflow-hidden" onClick={e => e.stopPropagation()}>
                  <div className="p-5 border-b border-gray-100 flex justify-between items-center bg-indigo-50/50">
                      <div className="flex items-center gap-3">
                          <div className="bg-indigo-100 p-2 rounded-lg">
                              <PestIcon className="w-5 h-5 text-indigo-600" />
                          </div>
                          <div>
                              <h3 className="text-lg font-bold text-gray-800">{pestModalData.crop}</h3>
                              <p className="text-sm text-gray-500">توزيع الآفات المسجلة</p>
                          </div>
                      </div>
                      <button onClick={() => setPestModalData(null)} className="text-gray-400 hover:text-gray-600 p-2 hover:bg-gray-100 rounded-full transition-colors">
                          <ClearIcon className="w-5 h-5" />
                      </button>
                  </div>
                  <div className="p-0 max-h-[60vh] overflow-y-auto">
                      <table className="min-w-full divide-y divide-gray-50">
                          <thead className="bg-gray-50 sticky top-0">
                              <tr>
                                  <th className="px-6 py-3 text-right text-xs font-semibold text-gray-500 uppercase">الآفة</th>
                                  <th className="px-6 py-3 text-right text-xs font-semibold text-gray-500 uppercase">عدد التوصيات</th>
                              </tr>
                          </thead>
                          <tbody className="divide-y divide-gray-50 bg-white">
                              {pestModalData.pests.map((pest, idx) => (
                                  <tr key={idx} className="hover:bg-gray-50">
                                      <td className="px-6 py-3 text-sm text-gray-700">{pest.pest}</td>
                                      <td className="px-6 py-3 text-sm font-medium text-indigo-600">{pest.count}</td>
                                  </tr>
                              ))}
                          </tbody>
                      </table>
                  </div>
                  <div className="p-4 border-t border-gray-100 bg-gray-50/50 text-center">
                      <button onClick={() => setPestModalData(null)} className="text-sm text-gray-500 hover:text-gray-700 font-medium">
                          إغلاق
                      </button>
                  </div>
              </div>
          </div>
      )}

      {/* Drill-down Modal */}
        {drillDownData && (
            <div className="fixed inset-0 z-50 overflow-y-auto bg-gray-900/75 backdrop-blur-sm flex items-center justify-center p-4" onClick={() => setDrillDownData(null)}>
                <div className="bg-white rounded-xl shadow-2xl w-full max-w-6xl max-h-[90vh] flex flex-col animate-in fade-in zoom-in duration-200" onClick={e => e.stopPropagation()}>
                    <div className="p-5 border-b border-gray-100 flex justify-between items-center bg-gray-50/50 rounded-t-xl">
                        <div className="flex items-center gap-3">
                           <div className="bg-indigo-100 p-2 rounded-lg">
                               <TableIcon className="w-5 h-5 text-indigo-600" />
                           </div>
                           <h3 className="text-xl font-bold text-gray-800">{drillDownData.title} <span className="text-sm font-normal text-gray-500 bg-white px-2 py-0.5 rounded-full border border-gray-200 shadow-sm mx-2">{drillDownData.rows.length} نتيجة</span></h3>
                        </div>
                        <button onClick={() => setDrillDownData(null)} className="text-gray-400 hover:text-gray-600 p-2 hover:bg-gray-100 rounded-full transition-colors">
                            <ClearIcon className="w-6 h-6" />
                        </button>
                    </div>
                    
                    <div className="overflow-auto flex-1">
                        <table className="min-w-full divide-y divide-gray-200">
                            <thead className="bg-gray-50 sticky top-0 shadow-sm z-10">
                                <tr>
                                    {drillDownData.columns.map((header, idx) => (
                                        <th key={idx} className="px-6 py-3 text-right text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap bg-gray-50">
                                            {header}
                                        </th>
                                    ))}
                                </tr>
                            </thead>
                            <tbody className="bg-white divide-y divide-gray-100">
                                {drillDownData.rows.map((row, rIdx) => (
                                    <tr key={rIdx} className="hover:bg-indigo-50/50 transition-colors">
                                        {drillDownData.columns.map((col, cIdx) => (
                                            <td key={cIdx} className="px-6 py-3 whitespace-nowrap text-sm text-gray-600 border-b border-gray-50">
                                                {String(row[col] !== undefined && row[col] !== null ? row[col] : '')}
                                            </td>
                                        ))}
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                        {drillDownData.rows.length === 0 && (
                            <div className="flex flex-col items-center justify-center p-12 text-gray-400">
                                <TableIcon className="w-12 h-12 mb-3 opacity-50" />
                                <p>لا توجد بيانات مطابقة للعرض</p>
                            </div>
                        )}
                    </div>

                    <div className="p-4 border-t border-gray-100 bg-gray-50/50 rounded-b-xl flex justify-end">
                        <button 
                            onClick={() => setDrillDownData(null)}
                            className="px-5 py-2 bg-white border border-gray-200 text-gray-700 rounded-lg hover:bg-gray-50 hover:border-gray-300 font-medium shadow-sm transition-all"
                        >
                            إغلاق
                        </button>
                    </div>
                </div>
            </div>
        )}
      
      {fileName && (
          <div className="fixed bottom-0 left-0 right-0 bg-white border-t border-gray-200 shadow-[0_-2px_10px_rgba(0,0,0,0.05)] z-20 md:hidden">
            <div className="flex justify-around items-center">
                <button 
                    onClick={() => setViewMode('upload')} 
                    className={`flex flex-col items-center py-3 px-4 w-full ${viewMode === 'upload' ? 'text-indigo-600' : 'text-gray-400 hover:text-gray-600'}`}
                >
                    <UploadIcon className="w-6 h-6" />
                    <span className="text-[10px] mt-1 font-medium">رفع</span>
                </button>
                <button 
                    onClick={() => setViewMode('dashboard')} 
                    className={`flex flex-col items-center py-3 px-4 w-full ${viewMode === 'dashboard' ? 'text-indigo-600' : 'text-gray-400 hover:text-gray-600'}`} 
                    disabled={!dashboardData}
                >
                    <DashboardIcon className="w-6 h-6" />
                    <span className="text-[10px] mt-1 font-medium">الرئيسية</span>
                </button>
                <button 
                    onClick={() => setViewMode('table')} 
                    className={`flex flex-col items-center py-3 px-4 w-full ${viewMode === 'table' ? 'text-indigo-600' : 'text-gray-400 hover:text-gray-600'}`} 
                    disabled={!allSheetsData.length}
                >
                    <TableIcon className="w-6 h-6" />
                    <span className="text-[10px] mt-1 font-medium">البيانات</span>
                </button>
            </div>
          </div>
      )}
      
      {fileName && (
        <div className="hidden md:block fixed left-0 top-1/2 -translate-y-1/2 z-20 bg-white shadow-lg rounded-r-xl border border-gray-200 py-2">
            <div className="flex flex-col gap-2">
                <button 
                    onClick={() => setViewMode('upload')} 
                    className={`p-3 hover:bg-gray-50 transition-colors relative group ${viewMode === 'upload' ? 'text-indigo-600 border-r-2 border-indigo-600' : 'text-gray-400'}`}
                    title="رفع ملف"
                >
                    <UploadIcon className="w-6 h-6" />
                    <span className="absolute left-full ml-2 bg-gray-800 text-white text-xs px-2 py-1 rounded opacity-0 group-hover:opacity-100 whitespace-nowrap transition-opacity pointer-events-none">رفع ملف</span>
                </button>
                <button 
                    onClick={() => setViewMode('dashboard')} 
                    className={`p-3 hover:bg-gray-50 transition-colors relative group ${viewMode === 'dashboard' ? 'text-indigo-600 border-r-2 border-indigo-600' : 'text-gray-400'}`} 
                    disabled={!dashboardData}
                    title="لوحة المعلومات"
                >
                    <DashboardIcon className="w-6 h-6" />
                    <span className="absolute left-full ml-2 bg-gray-800 text-white text-xs px-2 py-1 rounded opacity-0 group-hover:opacity-100 whitespace-nowrap transition-opacity pointer-events-none">لوحة المعلومات</span>
                </button>
                <button 
                    onClick={() => setViewMode('table')} 
                    className={`p-3 hover:bg-gray-50 transition-colors relative group ${viewMode === 'table' ? 'text-indigo-600 border-r-2 border-indigo-600' : 'text-gray-400'}`} 
                    disabled={!allSheetsData.length}
                    title="البيانات"
                >
                    <TableIcon className="w-6 h-6" />
                    <span className="absolute left-full ml-2 bg-gray-800 text-white text-xs px-2 py-1 rounded opacity-0 group-hover:opacity-100 whitespace-nowrap transition-opacity pointer-events-none">البيانات</span>
                </button>
            </div>
        </div>
      )}

      {isLoading && (
        <div className="fixed inset-0 bg-gray-900 bg-opacity-50 flex items-center justify-center z-50">
            <div className="animate-spin rounded-full h-12 w-12 border-4 border-indigo-500 border-t-transparent"></div>
        </div>
      )}
    </div>
  );
};