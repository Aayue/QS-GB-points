import React, { useState, useRef, useEffect, useCallback } from 'react';
import { createRoot } from 'react-dom/client';
import { GoogleGenAI } from "@google/genai";
import * as XLSX from 'xlsx';
import {
  Upload,
  FileSpreadsheet,
  Loader2,
  CheckCircle,
  AlertCircle,
  X,
  Download,
  Copy,
  Check,
  Settings,
  Users,
  HelpCircle,
  Eye,
  EyeOff,
  FileText,
  RotateCcw,
  Moon,
  Sun,
  Edit2,
  Plus,
  RefreshCw
} from 'lucide-react';

// Interfaces for type safety
interface RawRow {
  teacher: string;
  code?: string;
  count: number;
  matchedStandardName?: string | null;
  sourceFile?: string;
}

interface ProcessedRow {
  id: number;
  "序号": number;
  "姓名": string;
  "课时": number;
  "得分": number;
  "明细笔数": number;
}

interface UploadedFile {
  id: string;
  file: File;
  previewUrl: string | null;
}

const DEFAULT_REFERENCE_NAMES = `任艳 Rebecca
王芸 Yun
樊澄 Claire
孟远 Krystal
刘康 Ivy
杨晶晶 Jenny
徐育梅 Catherine
储晓栋 Amy
孟雅 Alice
陈蓉 Lotus
谢淑平 Sherry
郭晶晶 Joy
李臣亚 Linna
杨阳Sunny
詹和太James
曹阳Bill
蒋雨桢 Tina
谭颖Tenny
刘鹏宇Dean
张月明 Jasmine 
徐晓娜 Elina
赵茹怡Daicen
金辰昊Colin
张金慧 Emily
冯灿Sunny
李雪维 Nicole 
张国靖 Alicia 
秦笑Sophie 
刘家玮Jane
Jane Z
施璎真Serena
丁佐伊Zoe
宋吉 Merissa 
王楚怡phyllis
周筱妍Ollie
张颖 Yvonne 
李苏粤Sibyl
许明月 Luna
季远兰 Mandy
吴飞Fiona（英语）
孙胜蓝 Sierra（英语）
汪苇 Winnie
朱嬿蓉`;

const LOCAL_STORAGE_KEY = 'teacher_scores_reference_names_v4';

const App = () => {
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFile[]>([]);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [extractedData, setExtractedData] = useState<ProcessedRow[] | null>(null);
  const [unmatchedRows, setUnmatchedRows] = useState<RawRow[]>([]);
  const [showUnmatched, setShowUnmatched] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [referenceNames, setReferenceNames] = useState<string>(() => {
    const saved = localStorage.getItem(LOCAL_STORAGE_KEY);
    if (saved !== null) {
      if (!saved.includes('Jane Z')) {
        const updated = saved.includes('刘家玮Jane')
          ? saved.replace('刘家玮Jane', '刘家玮Jane\nJane Z')
          : saved + '\nJane Z';
        localStorage.setItem(LOCAL_STORAGE_KEY, updated);
        return updated;
      }
      return saved;
    }
    // Also check older v3 storage
    const v3Saved = localStorage.getItem('teacher_scores_reference_names_v3');
    if (v3Saved !== null) {
      const updated = v3Saved.includes('Jane Z')
        ? v3Saved
        : v3Saved.includes('刘家玮Jane')
        ? v3Saved.replace('刘家玮Jane', '刘家玮Jane\nJane Z')
        : v3Saved + '\nJane Z';
      localStorage.setItem(LOCAL_STORAGE_KEY, updated);
      return updated;
    }
    return DEFAULT_REFERENCE_NAMES;
  });
  const [showSettings, setShowSettings] = useState(true);
  const [isDarkMode, setIsDarkMode] = useState(false);
  const [copied, setCopied] = useState(false);
  const [editingRowIndex, setEditingRowIndex] = useState<number | null>(null);
  const [editingCountValue, setEditingCountValue] = useState<string>('');
  const [nameListNotification, setNameListNotification] = useState<string | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const nameListInputRef = useRef<HTMLInputElement>(null);

  // Sync referenceNames to localStorage
  useEffect(() => {
    localStorage.setItem(LOCAL_STORAGE_KEY, referenceNames);
  }, [referenceNames]);

  // Initialize Theme based on local time
  useEffect(() => {
    const hour = new Date().getHours();
    const isNight = hour >= 18 || hour < 6;
    setIsDarkMode(isNight);
  }, []);

  const toggleTheme = () => setIsDarkMode(!isDarkMode);

  const generateId = () => Math.random().toString(36).substring(2, 11);

  const restoreDefaultNames = () => {
    setReferenceNames(DEFAULT_REFERENCE_NAMES);
    localStorage.setItem(LOCAL_STORAGE_KEY, DEFAULT_REFERENCE_NAMES);
    const count = DEFAULT_REFERENCE_NAMES.split('\n').map(n => n.trim()).filter(Boolean).length;
    setNameListNotification(`已恢复为默认 ${count} 位教师名单（含 Jane Liu 与 Jane Z）`);
    setTimeout(() => setNameListNotification(null), 3000);
  };

  const handleClearClick = () => {
    setReferenceNames('');
    localStorage.setItem(LOCAL_STORAGE_KEY, '');
    setNameListNotification('已清空名单。您可直接点击“上传姓名列表”导入文件，或在下方文本框内粘贴新名单。');
    setTimeout(() => setNameListNotification(null), 4000);
  };

  const parseNameListFile = async (file: File) => {
    try {
      let names: string[] = [];
      const fileName = file.name.toLowerCase();

      if (fileName.endsWith('.txt')) {
        const text = await file.text();
        names = text
          .split(/\r?\n/)
          .map(n => n.trim())
          .filter(n => n.length > 0);
      } else if (
        fileName.endsWith('.xlsx') ||
        fileName.endsWith('.xls') ||
        fileName.endsWith('.csv') ||
        file.type.includes('spreadsheet') ||
        file.type.includes('excel') ||
        file.type.includes('csv')
      ) {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          throw new Error('表格中没有可读取的工作表');
        }
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        const data: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        if (data.length > 0) {
          let nameColIndex = -1;
          let enColIndex = -1;
          let startRowIndex = 0;
          let foundHeader = false;

          // Check header row for Name / Chinese Name / English Name columns
          const headerRow = data[0] || [];
          for (let col = 0; col < headerRow.length; col++) {
            const val = String(headerRow[col] || '').trim().toLowerCase();
            if (/中文名|^姓名$|教师姓名|教师|老师/i.test(val)) {
              nameColIndex = col;
              startRowIndex = 1;
              foundHeader = true;
            } else if (/英文名|^en$|english|eng\s*name/i.test(val)) {
              enColIndex = col;
              startRowIndex = 1;
              foundHeader = true;
            } else if (nameColIndex === -1 && /name|teacher|instructor/i.test(val)) {
              nameColIndex = col;
              startRowIndex = 1;
              foundHeader = true;
            }
          }

          if (!foundHeader && data.length > 1) {
            const subHeader = data[1] || [];
            for (let col = 0; col < subHeader.length; col++) {
              const val = String(subHeader[col] || '').trim().toLowerCase();
              if (/中文名|^姓名$|教师姓名|教师|老师/i.test(val)) {
                nameColIndex = col;
                startRowIndex = 2;
                foundHeader = true;
              } else if (/英文名|^en$|english|eng\s*name/i.test(val)) {
                enColIndex = col;
                startRowIndex = 2;
                foundHeader = true;
              } else if (nameColIndex === -1 && /name|teacher|instructor/i.test(val)) {
                nameColIndex = col;
                startRowIndex = 2;
                foundHeader = true;
              }
            }
          }

          if (nameColIndex === -1) {
            // Find first column containing non-numeric strings
            for (let col = 0; col < (data[0]?.length || 1); col++) {
              const sample = String(data[0]?.[col] || '').trim();
              if (sample && !/^\d+$/.test(sample) && !/^(no|序号|id)$/i.test(sample)) {
                nameColIndex = col;
                break;
              }
            }
            if (nameColIndex === -1) nameColIndex = 0;
          }

          for (let r = startRowIndex; r < data.length; r++) {
            const rawVal = String(data[r]?.[nameColIndex] || '').trim();
            const rawEnVal = enColIndex !== -1 ? String(data[r]?.[enColIndex] || '').trim() : '';

            let fullName = rawVal;
            if (rawEnVal && rawEnVal !== rawVal && !rawVal.toLowerCase().includes(rawEnVal.toLowerCase())) {
              fullName = `${rawVal} ${rawEnVal}`.trim();
            }

            if (
              fullName &&
              !/^(合计|总计|total|平均|average|序号|no\.?)$/i.test(fullName) &&
              !/^\d+$/.test(fullName)
            ) {
              names.push(fullName);
            }
          }
        }
      } else {
        throw new Error('不支持的文件格式，请上传 .xlsx, .xls, .csv 或 .txt 文件');
      }

      if (names.length === 0) {
        setNameListNotification('未能从文件中提取出有效姓名，请检查文件格式。');
        setTimeout(() => setNameListNotification(null), 4000);
        return;
      }

      const newNamesText = names.join('\n');
      setReferenceNames(newNamesText);
      localStorage.setItem(LOCAL_STORAGE_KEY, newNamesText);
      setNameListNotification(`✅ 成功导入 ${names.length} 位教师名单！已按文件顺序排布。`);
      setTimeout(() => setNameListNotification(null), 4500);
    } catch (err) {
      setNameListNotification('文件解析失败: ' + (err as Error).message);
      setTimeout(() => setNameListNotification(null), 4000);
    }
  };

  const handleNameListUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      await parseNameListFile(file);
    }
    if (nameListInputRef.current) nameListInputRef.current.value = '';
  };

  const handleTextareaDrop = async (e: React.DragEvent<HTMLTextAreaElement>) => {
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      const fileName = file.name.toLowerCase();
      if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls') || fileName.endsWith('.csv') || fileName.endsWith('.txt')) {
        e.preventDefault();
        await parseNameListFile(file);
      }
    }
  };

  // Scoring function: Count <= 10: 0 pts; 11-20: 1 pt; 21-35: 3 pts; >= 36: 5 pts
  const calculateScore = (count: number): number => {
    if (count <= 10) return 0;
    if (count <= 20) return 1;
    if (count <= 35) return 3;
    return 5;
  };

  const processFiles = useCallback((files: File[]) => {
    const newUploadedFiles: UploadedFile[] = [];
    const validTypes = [
      'image/png',
      'image/jpeg',
      'image/jpg',
      'image/webp',
      'application/pdf',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel'
    ];

    files.forEach(file => {
      const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
      const isValid = validTypes.some(t => file.type.includes(t)) || isExcel;

      if (isValid) {
        let previewUrl: string | null = null;
        if (file.type.startsWith('image/') || file.type === 'application/pdf') {
          const reader = new FileReader();
          reader.onload = (e) => {
            setUploadedFiles(prev => prev.map(item => {
              if (item.file === file) {
                return { ...item, previewUrl: e.target?.result as string };
              }
              return item;
            }));
          };
          reader.readAsDataURL(file);
        }

        newUploadedFiles.push({
          id: generateId(),
          file,
          previewUrl
        });
      }
    });

    if (newUploadedFiles.length > 0) {
      setUploadedFiles(prev => [...prev, ...newUploadedFiles]);
      setError(null);
      setExtractedData(null);
      setUnmatchedRows([]);
    } else if (files.length > 0) {
      setError('未检测到支持的文件格式。请上传图片、PDF 或 Excel 文件。');
    }
  }, []);

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      processFiles(Array.from(e.target.files));
    }
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      processFiles(Array.from(e.dataTransfer.files));
    }
  };

  // Paste Event Listener for quick screenshot upload
  useEffect(() => {
    const handlePaste = (e: ClipboardEvent) => {
      if (e.clipboardData && e.clipboardData.items) {
        const files: File[] = [];
        for (let i = 0; i < e.clipboardData.items.length; i++) {
          const item = e.clipboardData.items[i];
          if (item.type.indexOf('image') !== -1) {
            const blob = item.getAsFile();
            if (blob) files.push(blob);
          }
        }
        if (files.length > 0) {
          processFiles(files);
          e.preventDefault();
        }
      }
    };

    document.addEventListener('paste', handlePaste);
    return () => {
      document.removeEventListener('paste', handlePaste);
    };
  }, [processFiles]);

  const removeFile = (id: string) => {
    setUploadedFiles(prev => prev.filter(f => f.id !== id));
    setExtractedData(null);
    setUnmatchedRows([]);
  };

  const clearAllFiles = () => {
    setUploadedFiles([]);
    setExtractedData(null);
    setUnmatchedRows([]);
    setError(null);
  };

  // Comprehensive Chinese surname pinyin map for precise English name and surname identification
  const SURNAME_PINYIN_MAP: Record<string, string> = {
    '刘': 'liu', '杨': 'yang', '冯': 'feng', '张': 'zhang', '周': 'zhou',
    '赵': 'zhao', '朱': 'zhu', '曾': 'zeng', '郑': 'zheng', '钟': 'zhong',
    '庄': 'zhuang', '王': 'wang', '李': 'li', '陈': 'chen', '吴': 'wu',
    '孙': 'sun', '徐': 'xu', '曹': 'cao', '金': 'jin', '丁': 'ding',
    '宋': 'song', '谢': 'xie', '郭': 'guo', '孟': 'meng', '樊': 'fan',
    '任': 'ren', '詹': 'zhan', '谭': 'tan', '蒋': 'jiang', '秦': 'qin',
    '施': 'shi', '许': 'xu', '季': 'ji', '汪': 'wang', '储': 'chu',
    '钱': 'qian', '沈': 'shen', '韩': 'han', '何': 'he', '高': 'gao',
    '马': 'ma', '罗': 'luo', '梁': 'liang', '黄': 'huang', '陆': 'lu',
    '唐': 'tang', '董': 'dong', '萧': 'xiao', '肖': 'xiao', '程': 'cheng',
    '袁': 'yuan', '邓': 'deng', '潘': 'pan', '杜': 'du', '戴': 'dai',
    '夏': 'xia', '田': 'tian', '胡': 'hu', '叶': 'ye', '方': 'fang'
  };

  interface TeacherProfile {
    fullName: string;
    zhChars: string;
    zhSurname: string | null;
    zhSurnamePinyin: string | null;
    enWords: string[];
    enFirstName: string | null;
    enSurname: string | null;
    keywords: string[];
  }

  const parseTeacherProfile = (standardNameLine: string): TeacherProfile => {
    const line = standardNameLine.trim();
    const zhMatches = line.match(/[\u4e00-\u9fa5]+/g);
    const zhChars = zhMatches ? zhMatches.join('') : '';

    // Extract English words / initials (case-insensitive, preserving single letters like "Z")
    const enMatches = line.match(/[a-zA-Z]+/g);
    const rawEnWords = enMatches ? enMatches.map(w => w.toLowerCase()) : [];
    
    // Deduplicate English words while preserving order
    const uniqueEnWords: string[] = [];
    for (const w of rawEnWords) {
      if (!uniqueEnWords.includes(w)) uniqueEnWords.push(w);
    }

    const zhSurname = zhChars.length > 0 ? zhChars[0] : null;
    const zhSurnamePinyin = zhSurname && SURNAME_PINYIN_MAP[zhSurname] ? SURNAME_PINYIN_MAP[zhSurname] : null;

    let enFirstName: string | null = null;
    let enSurname: string | null = null;

    if (zhSurnamePinyin) {
      enSurname = zhSurnamePinyin;
      // The English first name is the non-surname English word (e.g. "sunny" from "杨阳Sunny (Sunny Yang)")
      const nonSurnameWords = uniqueEnWords.filter(w => w !== zhSurnamePinyin);
      if (nonSurnameWords.length > 0) {
        enFirstName = nonSurnameWords[0];
      }
    } else if (uniqueEnWords.length >= 2) {
      enFirstName = uniqueEnWords[0];
      enSurname = uniqueEnWords[1];
    } else if (uniqueEnWords.length === 1) {
      enFirstName = uniqueEnWords[0];
    }

    // Build searchable keywords
    const keywords: string[] = [];
    if (zhChars) keywords.push(zhChars);
    if (line.includes('朱嬿蓉') || line.includes('朱燕蓉')) {
      keywords.push('朱嬿蓉', '朱燕蓉');
    }
    uniqueEnWords.forEach(w => {
      if (!keywords.includes(w)) keywords.push(w);
    });
    if (zhSurnamePinyin && !keywords.includes(zhSurnamePinyin)) {
      keywords.push(zhSurnamePinyin);
    }
    if (enFirstName && enSurname) {
      const combined = `${enFirstName} ${enSurname}`;
      if (!keywords.includes(combined)) keywords.push(combined);
    }

    return {
      fullName: line,
      zhChars,
      zhSurname,
      zhSurnamePinyin,
      enWords: uniqueEnWords,
      enFirstName,
      enSurname,
      keywords
    };
  };

  const calculateMatchScore = (rawTeacherStr: string, profile: TeacherProfile, rawCode?: string): number => {
    const rawClean = (rawTeacherStr || "").trim();
    const rawLower = rawClean.toLowerCase();
    const codeLower = (rawCode || "").toLowerCase();

    const zhMatches = rawClean.match(/[\u4e00-\u9fa5]+/g);
    const rawZh = zhMatches ? zhMatches.join('') : '';

    const enMatches = rawClean.match(/[a-zA-Z]+/g);
    const rawEnWords = enMatches ? enMatches.map(w => w.toLowerCase()) : [];

    let score = 0;

    // 1. Direct Chinese Match
    if (profile.zhChars && profile.zhChars.length > 0) {
      if (rawClean.includes(profile.zhChars) || (rawZh && rawZh === profile.zhChars)) {
        score += 180;
      }
    }

    // Handle 朱嬿蓉 / 朱燕蓉 variant
    if (
      (profile.fullName.includes('朱嬿蓉') || profile.fullName.includes('朱燕蓉')) &&
      (rawClean.includes('朱嬿蓉') || rawClean.includes('朱燕蓉'))
    ) {
      score += 180;
    }

    // 2. English & Surname Match (Crucial for Jane Liu vs Jane Z, Sunny Yang vs Sunny Feng)
    const rawFirstName = rawEnWords.length > 0 ? rawEnWords[0] : null;
    let rawSurname: string | null = null;
    if (rawEnWords.length >= 2) {
      rawSurname = rawEnWords[1];
    } else if (rawZh.length > 0 && SURNAME_PINYIN_MAP[rawZh[0]]) {
      rawSurname = SURNAME_PINYIN_MAP[rawZh[0]];
    }

    if (rawFirstName && profile.enFirstName && rawFirstName === profile.enFirstName) {
      score += 40;

      const stdSurname = profile.enSurname;
      if (rawSurname && stdSurname) {
        let surnameMatches = false;

        if (rawSurname === stdSurname) {
          surnameMatches = true;
        } else if (stdSurname === 'z' && rawSurname.startsWith('z')) {
          // e.g. raw is "Jane Zhou" or "Jane Zhang" or "Jane Zhao" and std is "Jane Z"
          surnameMatches = true;
        } else if (rawSurname === 'z' && stdSurname.startsWith('z')) {
          // e.g. raw is "Jane Z" and std is "Jane Zhou"
          surnameMatches = true;
        }

        if (surnameMatches) {
          score += 80; // Big bonus for matching surname
        } else {
          // Strict conflict penalty (e.g. raw has Liu but standard has Z, or raw has Yang but standard has Feng)
          score -= 5000;
        }
      } else if (!rawSurname) {
        // Disambiguation via course/subject code when raw name has only first name
        if (profile.enFirstName === 'sunny') {
          if (/chem|化/.test(codeLower) || /chem|化/.test(rawLower)) {
            if (profile.enSurname === 'feng') score += 50;
          } else if (/math|数/.test(codeLower) || /math|数/.test(rawLower)) {
            if (profile.enSurname === 'yang') score += 50;
          }
        } else if (profile.enFirstName === 'jane') {
          if (/geog|glob|地|全球/.test(codeLower) || /geog|glob|地|全球/.test(rawLower)) {
            if (profile.enSurname === 'liu') score += 50;
          }
        }
      }
    }

    // 3. Keyword positive score
    profile.keywords.forEach(kw => {
      if (!kw || kw.length === 0) return;
      if (/[\u4e00-\u9fa5]/.test(kw)) {
        if (rawClean.includes(kw)) score += 30;
      } else if (kw.length > 1) {
        if (rawLower.includes(kw)) score += kw.length;
      }
    });

    return score;
  };

  // Main Analysis function
  const analyzeData = async () => {
    if (uploadedFiles.length === 0) return;

    setIsAnalyzing(true);
    setError(null);
    setUnmatchedRows([]);

    try {
      const allRawRowsResults = await Promise.all(uploadedFiles.map(async (uFile) => {
        const { file, previewUrl } = uFile;
        let fileRows: RawRow[] = [];

        try {
          // BRANCH 1: AI VISION EXTRACTION (IMAGE OR PDF)
          if ((file.type.startsWith('image/') || file.type === 'application/pdf') && previewUrl) {
            const apiKey = process.env.API_KEY || process.env.GEMINI_API_KEY;
            const ai = new GoogleGenAI({ apiKey: apiKey || '' });
            const base64Data = previewUrl.split(',')[1];

            const prompt = `
            Analyze this ${file.type === 'application/pdf' ? 'PDF document' : 'image'} containing teacher schedules, lesson counts, tables, or summary charts.
            Extract EVERY teacher's name and their corresponding lesson count/hours.
            
            Columns/Elements often represent:
            - Teacher name (姓名 / 教师 / Teacher / Instructor)
            - Lesson Count / Hours / Quantity (节数 / 课时 / 数量 / Count / 课时数 / 总节数 / 圈内数字)
            - Optional Subject/Course (科目 / 课程 / Code / Description)
            
            Rules:
            1. Extract EVERY SINGLE teacher. Do NOT truncate or omit any teachers.
            2. PRESERVE FULL NAMES AND SURNAMES:
               - Distinguish Jane Liu vs Jane Z. Keep the full surname or initial (e.g. "Jane Liu", "Jane Z", "Jane Zhou").
               - Distinguish Sunny Yang vs Sunny Feng. Keep the full surname (e.g. "Sunny Yang", "Sunny Feng").
               - If only a first name like "Sunny" or "Jane" is shown, include the subject or course name in "code" to help identify them.
            3. Do NOT double-count: if a document has both itemized course rows AND total rows for a teacher, extract the itemized breakdown; if it only has total rows or summary bubbles/charts, extract the total count so no teacher is missed.
            4. If count is a decimal or fraction, extract the numerical value.
            
            Return ONLY a valid JSON array of objects with keys:
            [
              { "teacher": "Teacher Name as shown", "code": "optional subject or code", "count": 14 }
            ]
            `;

            let responseText = '';
            for (let attempt = 1; attempt <= 3; attempt++) {
              try {
                const response = await ai.models.generateContent({
                  model: 'gemini-3.8-flash',
                  contents: {
                    parts: [
                      {
                        inlineData: {
                          mimeType: file.type || 'image/png',
                          data: base64Data
                        }
                      },
                      { text: prompt }
                    ]
                  },
                  config: {
                    responseMimeType: "application/json"
                  }
                });

                responseText = response.text || '';
                if (responseText) break;
              } catch (apiErr: any) {
                console.warn(`Gemini attempt ${attempt} failed:`, apiErr);
                if (attempt === 3) throw apiErr;
                await new Promise(res => setTimeout(res, 1500 * attempt));
              }
            }

            if (!responseText) throw new Error(`未从文件 ${file.name} 中获取到有效数据`);

            try {
              const parsed = JSON.parse(responseText);
              let rows: any[] = [];
              if (Array.isArray(parsed)) {
                rows = parsed;
              } else if (typeof parsed === 'object' && parsed !== null) {
                const key = Object.keys(parsed).find(k => Array.isArray(parsed[k]));
                if (key) rows = parsed[key];
              }

              fileRows = rows.map((r: any) => {
                const teacherKey = Object.keys(r).find(k => /teacher|instructor|姓名|^name$|教师|老师/i.test(k));
                const countKey = Object.keys(r).find(k => /count|num|节数|quantity|课时|总计|课时数/i.test(k));
                const codeKey = Object.keys(r).find(k => /code|代码|科目|课程/i.test(k));

                const rawCount = countKey ? (Number(r[countKey]) || 0) : 0;
                return {
                  teacher: teacherKey ? String(r[teacherKey]).trim() : "",
                  code: codeKey ? String(r[codeKey]).trim() : "",
                  count: rawCount,
                  sourceFile: file.name
                };
              }).filter(r => r.teacher && r.count > 0);

            } catch (jsonErr) {
              console.warn(`JSON parse error for ${file.name}:`, jsonErr);
            }
          }
          // BRANCH 2: LOCAL EXCEL SPREADSHEET EXTRACTION
          else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || file.type.includes('spreadsheet') || file.type.includes('excel')) {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer);
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
            let lastTeacher = "";

            fileRows = jsonData.map((row: any) => {
              const teacherKey = Object.keys(row).find(k => /teacher|instructor|姓名|^name$|教师|老师/i.test(k));
              const codeKey = Object.keys(row).find(k => /code|代码|科目|课程/i.test(k));
              const countKey = Object.keys(row).find(k => /column|count|节数|数量|num|课时|课时数|总课时|总计/i.test(k));

              let rawTeacherVal = teacherKey ? String(row[teacherKey]).trim() : "";
              if (rawTeacherVal.toLowerCase().includes('total') || rawTeacherVal.includes('合计') || rawTeacherVal.includes('总计')) {
                return null;
              }

              let teacher = rawTeacherVal;
              let code = codeKey ? String(row[codeKey]).trim() : "";
              let countStr = countKey ? String(row[countKey]) : "0";
              let count = parseFloat(countStr) || 0;

              if (!teacher && lastTeacher) {
                teacher = lastTeacher;
              } else if (teacher) {
                lastTeacher = teacher;
              }

              return {
                teacher: teacher,
                code: code,
                count: count,
                sourceFile: file.name
              } as RawRow;
            }).filter((r): r is RawRow => {
              if (!r) return false;
              if (!r.teacher || r.teacher.toLowerCase().includes('total') || r.teacher.includes('合计') || r.teacher.includes('总计')) {
                return false;
              }
              return r.count > 0;
            });
          }
        } catch (fileErr) {
          console.error(`Error processing file ${file.name}:`, fileErr);
        }
        return fileRows;
      }));

      const rawRows = allRawRowsResults.flat();

      if (rawRows.length === 0) {
        throw new Error("未能从上传的文件中提取到有效数据，请检查文件是否包含教师姓名与课时数据。");
      }

      // Match rows against Standard Names list
      const stdNamesList = referenceNames
        .split('\n')
        .map(n => n.trim())
        .filter(n => n.length > 0);

      const stdProfiles = stdNamesList.map(name => parseTeacherProfile(name));

      const matchedRows: RawRow[] = rawRows.map(row => {
        let bestMatch: { name: string; score: number } | null = null;

        stdProfiles.forEach(profile => {
          const currentScore = calculateMatchScore(row.teacher, profile, row.code);
          if (currentScore > 0) {
            if (!bestMatch || currentScore > bestMatch.score) {
              bestMatch = { name: profile.fullName, score: currentScore };
            }
          }
        });

        return {
          ...row,
          matchedStandardName: bestMatch ? bestMatch.name : null
        };
      });

      setUnmatchedRows(matchedRows.filter(r => !r.matchedStandardName));

      // Direct Aggregation: Strictly preserve the given teacher names list order
      // Direct calculation without co-teaching
      const results: ProcessedRow[] = stdNamesList.map((stdName, index) => {
        const teacherRows = matchedRows.filter(r => r.matchedStandardName === stdName);
        let totalCount = 0;
        teacherRows.forEach(r => {
          totalCount += Number(r.count) || 0;
        });

        // Round to 1 decimal place if has decimals, otherwise integer
        const cleanCount = Math.round(totalCount * 10) / 10;
        const score = calculateScore(cleanCount);

        return {
          id: index + 1,
          "序号": index + 1,
          "姓名": stdName,
          "课时": cleanCount,
          "得分": score,
          "明细笔数": teacherRows.length
        };
      });

      setExtractedData(results);

    } catch (err) {
      console.error(err);
      setError("处理失败: " + (err as Error).message);
    } finally {
      setIsAnalyzing(false);
    }
  };

  // Download Excel table strictly preserving the teacher names order
  const downloadExcel = () => {
    if (!extractedData || extractedData.length === 0) return;

    // Excel formatting: 序号, 教师姓名, 总课时, 得分
    const excelRows = extractedData.map(row => ({
      "序号": row["序号"],
      "教师姓名": row["姓名"],
      "总课时": row["课时"],
      "得分": row["得分"]
    }));

    const ws = XLSX.utils.json_to_sheet(excelRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "教师得分汇总");

    // Set column widths
    ws['!cols'] = [
      { wch: 8 },  // 序号
      { wch: 26 }, // 教师姓名
      { wch: 14 }, // 总课时
      { wch: 12 }  // 得分
    ];

    XLSX.writeFile(wb, "教师得分汇总表.xlsx");
  };

  // Copy table to clipboard
  const copyTableToClipboard = () => {
    if (!extractedData || extractedData.length === 0) return;

    const headers = ["序号", "教师姓名", "总课时", "得分"].join("\t");
    const rows = extractedData.map(r => [r["序号"], r["姓名"], r["课时"], r["得分"]].join("\t")).join("\n");
    const fullText = headers + "\n" + rows;

    navigator.clipboard.writeText(fullText).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    });
  };

  // Handle manual inline edit of count
  const startEditing = (idx: number, currentCount: number) => {
    setEditingRowIndex(idx);
    setEditingCountValue(String(currentCount));
  };

  const saveEditing = (idx: number) => {
    if (editingRowIndex === null || !extractedData) return;
    const newCount = parseFloat(editingCountValue) || 0;
    const updated = [...extractedData];
    updated[idx] = {
      ...updated[idx],
      "课时": newCount,
      "得分": calculateScore(newCount)
    };
    setExtractedData(updated);
    setEditingRowIndex(null);
  };

  const currentNamesList = referenceNames.split('\n').map(n => n.trim()).filter(Boolean);

  // Statistics
  const totalTeachers = extractedData?.length || 0;
  const activeTeachers = extractedData?.filter(r => r["课时"] > 0).length || 0;
  const score5Teachers = extractedData?.filter(r => r["得分"] === 5).length || 0;
  const score3Teachers = extractedData?.filter(r => r["得分"] === 3).length || 0;
  const score1Teachers = extractedData?.filter(r => r["得分"] === 1).length || 0;
  const score0Teachers = extractedData?.filter(r => r["得分"] === 0).length || 0;

  const renderPreview = (uFile: UploadedFile) => {
    const { file, previewUrl } = uFile;

    if (file.type.startsWith('image/') && previewUrl) {
      return <img src={previewUrl} alt="Preview" className="h-20 w-20 object-cover rounded-lg border dark:border-gray-700 shadow-sm" />;
    }

    return (
      <div className="h-20 w-20 bg-gray-100 dark:bg-gray-750 rounded-lg border dark:border-gray-700 flex flex-col items-center justify-center p-2 text-center shadow-sm">
        {file.type.includes('pdf') ? (
          <FileText className="w-7 h-7 text-red-600 dark:text-red-400" />
        ) : (
          <FileSpreadsheet className="w-7 h-7 text-emerald-600 dark:text-emerald-400" />
        )}
        <span className="text-[10px] text-gray-600 dark:text-gray-300 mt-1 uppercase font-semibold truncate w-full">
          {file.name.split('.').pop()}
        </span>
      </div>
    );
  };

  return (
    <div className={isDarkMode ? 'dark' : ''}>
      <div className="min-h-screen bg-slate-50 dark:bg-gray-950 flex flex-col items-center py-10 px-4 sm:px-6 lg:px-8 text-gray-800 dark:text-gray-100 transition-colors duration-200">
        <div className="max-w-5xl w-full space-y-6">

          {/* Header */}
          <div className="relative flex flex-col items-center text-center">
            {/* Theme Toggle Button */}
            <button
              onClick={toggleTheme}
              className="absolute right-0 top-0 p-2.5 rounded-xl bg-white dark:bg-gray-850 shadow-sm border border-gray-200 dark:border-gray-750 text-gray-600 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-750 transition-colors"
              title={isDarkMode ? "切换到明亮模式" : "切换到暗黑模式"}
            >
              {isDarkMode ? <Sun className="w-5 h-5 text-amber-400" /> : <Moon className="w-5 h-5" />}
            </button>

            <div className="mx-auto h-16 w-16 bg-blue-100 dark:bg-blue-900/40 rounded-2xl flex items-center justify-center mb-4 shadow-sm border border-blue-200 dark:border-blue-800">
              <FileSpreadsheet className="h-9 w-9 text-blue-600 dark:text-blue-400" />
            </div>
            <h1 className="text-3xl font-extrabold text-gray-900 dark:text-white tracking-tight">
              教师得分计算器
            </h1>
            <p className="mt-2 text-sm sm:text-base text-gray-600 dark:text-gray-400 max-w-xl">
              直接提取课时数据，严格按标准名单顺序汇总计算教师得分，一键生成规范 Excel 表格
            </p>
          </div>

          {/* Scoring Rule Banner */}
          <div className="bg-gradient-to-r from-blue-50 via-indigo-50 to-blue-50 dark:from-blue-950/40 dark:via-indigo-950/30 dark:to-blue-950/40 border border-blue-200/80 dark:border-blue-800/80 rounded-2xl p-4 shadow-sm">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div className="flex items-center gap-2.5">
                <div className="p-1.5 bg-blue-600 text-white rounded-lg">
                  <HelpCircle className="w-4 h-4" />
                </div>
                <div>
                  <h3 className="text-sm font-bold text-gray-900 dark:text-white">计分规则标准 (Scoring Rules)</h3>
                  <p className="text-xs text-gray-500 dark:text-gray-400">已去除任何 co-teaching 折半，直接按教师课时总数核算</p>
                </div>
              </div>

              <div className="flex items-center gap-2 sm:gap-3 flex-wrap">
                <div className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 shadow-xs">
                  <span className="w-2.5 h-2.5 rounded-full bg-slate-400 dark:bg-slate-500"></span>
                  <span className="text-xs font-medium text-gray-700 dark:text-gray-300">课时 ≤ 10:</span>
                  <span className="text-xs font-bold text-gray-900 dark:text-white">0 分</span>
                </div>
                <div className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white dark:bg-gray-800 border border-amber-200 dark:border-amber-800/60 shadow-xs">
                  <span className="w-2.5 h-2.5 rounded-full bg-amber-500"></span>
                  <span className="text-xs font-medium text-amber-700 dark:text-amber-400">11 - 20:</span>
                  <span className="text-xs font-bold text-amber-800 dark:text-amber-300">1 分</span>
                </div>
                <div className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white dark:bg-gray-800 border border-blue-200 dark:border-blue-800/60 shadow-xs">
                  <span className="w-2.5 h-2.5 rounded-full bg-blue-600"></span>
                  <span className="text-xs font-medium text-blue-700 dark:text-blue-400">{'>'} 20:</span>
                  <span className="text-xs font-bold text-blue-800 dark:text-blue-300">3 分</span>
                </div>
              </div>
            </div>
          </div>

          {/* Standard Names List (可编辑修改名单) */}
          <div className="bg-white dark:bg-gray-850 shadow-sm rounded-2xl overflow-hidden border border-gray-200 dark:border-gray-750">
            <div
              className="px-6 py-4 bg-gray-50/80 dark:bg-gray-850 border-b border-gray-200 dark:border-gray-750 flex items-center justify-between cursor-pointer select-none"
              onClick={() => setShowSettings(!showSettings)}
            >
              <div className="flex items-center gap-2.5 font-semibold text-gray-800 dark:text-gray-200">
                <Users className="w-5 h-5 text-blue-600 dark:text-blue-400" />
                <span>标准教师名单 (Standard Name List)</span>
                <span className="ml-1 text-xs px-2.5 py-0.5 rounded-full bg-blue-100 text-blue-800 dark:bg-blue-900/60 dark:text-blue-300 font-bold">
                  共 {currentNamesList.length} 位
                </span>
              </div>
              <div className="flex items-center gap-3">
                <span className="text-xs text-gray-500 dark:text-gray-400 hidden sm:inline">
                  {showSettings ? '点击收起' : '点击展开编辑'}
                </span>
                <span className="text-sm font-medium text-blue-600 dark:text-blue-400">
                  {showSettings ? '收起' : '展开修改'}
                </span>
              </div>
            </div>

            {showSettings && (
              <div className="p-6 space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <p className="text-xs text-gray-500 dark:text-gray-400">
                    每行一位教师姓名（含中英文），Excel 导出与汇总结果将<strong>严格按照此处的上下顺序</strong>排列输出：
                  </p>
                  <div className="flex items-center gap-2 flex-wrap">
                    {/* Hidden File Input for Name List */}
                    <input
                      type="file"
                      ref={nameListInputRef}
                      accept=".xlsx,.xls,.csv,.txt"
                      className="hidden"
                      onChange={handleNameListUpload}
                    />

                    <button
                      onClick={() => nameListInputRef.current?.click()}
                      className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 active:bg-blue-800 rounded-lg shadow-sm transition-all cursor-pointer ring-1 ring-blue-500"
                      title="支持上传 Excel (.xlsx, .xls)、CSV 或 TXT 文件，自动提取并替换教师名单"
                    >
                      <Upload className="w-3.5 h-3.5" /> 上传姓名列表 (Excel/CSV/TXT)
                    </button>

                    <button
                      onClick={restoreDefaultNames}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-blue-600 dark:text-blue-400 bg-blue-50 dark:bg-blue-950/40 hover:bg-blue-100 dark:hover:bg-blue-900/50 rounded-lg border border-blue-200 dark:border-blue-800 transition-colors cursor-pointer"
                      title="恢复为默认的标准教师名单"
                    >
                      <RotateCcw className="w-3.5 h-3.5" /> 恢复默认 ({DEFAULT_REFERENCE_NAMES.split('\n').filter(Boolean).length}人)
                    </button>

                    <button
                      onClick={handleClearClick}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-red-600 dark:text-red-400 bg-red-50 dark:bg-red-950/40 hover:bg-red-100 dark:hover:bg-red-900/50 rounded-lg border border-red-200 dark:border-red-800/60 transition-colors cursor-pointer"
                      title="一键清空名单，方便上传新文件或粘贴"
                    >
                      <X className="w-3.5 h-3.5" /> 清空列表
                    </button>
                  </div>
                </div>

                {/* Intelligent Surname Differentiation Notice */}
                <div className="flex items-center gap-2 p-2.5 rounded-xl bg-blue-50/70 dark:bg-blue-950/40 border border-blue-200/70 dark:border-blue-800/50 text-[11px] text-blue-800 dark:text-blue-300">
                  <span className="font-bold flex items-center gap-1 flex-shrink-0">
                    <CheckCircle className="w-3.5 h-3.5 text-blue-600 dark:text-blue-400" />
                    智能姓氏重名区分：
                  </span>
                  <span>
                    已配置英文同名区分规则：<strong>Jane Liu (刘家玮)</strong> 与 <strong>Jane Z</strong> 自动按姓氏精确区分，各自独立核算课时；<strong>Sunny Yang (杨阳)</strong> 与 <strong>Sunny Feng (冯灿)</strong> 同样精确区分。
                  </span>
                </div>

                {nameListNotification && (
                  <div className="flex items-center justify-between gap-2 p-3 rounded-xl text-xs font-medium bg-emerald-50 dark:bg-emerald-950/60 text-emerald-800 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800 animate-fadeIn">
                    <div className="flex items-center gap-2">
                      <CheckCircle className="w-4 h-4 text-emerald-600 dark:text-emerald-400 flex-shrink-0" />
                      <span>{nameListNotification}</span>
                    </div>
                    <button
                      onClick={() => setNameListNotification(null)}
                      className="text-emerald-500 hover:text-emerald-700 dark:hover:text-emerald-200"
                    >
                      <X className="w-3.5 h-3.5" />
                    </button>
                  </div>
                )}

                {/* Empty State when list is cleared */}
                {currentNamesList.length === 0 ? (
                  <div className="border-2 border-dashed border-blue-200 dark:border-blue-800/60 rounded-xl p-8 text-center bg-blue-50/30 dark:bg-blue-950/20 flex flex-col items-center justify-center space-y-3">
                    <div className="h-12 w-12 rounded-xl bg-blue-100 dark:bg-blue-900/50 text-blue-600 dark:text-blue-400 flex items-center justify-center shadow-xs">
                      <Upload className="w-6 h-6" />
                    </div>
                    <div>
                      <p className="text-sm font-bold text-gray-800 dark:text-gray-200">当前教师名单为空</p>
                      <p className="text-xs text-gray-500 dark:text-gray-400 mt-1 max-w-md">
                        您可直接点击下方按钮上传 Excel / CSV / TXT 文件，或在下方输入框中直接粘贴教师姓名。
                      </p>
                    </div>
                    <div className="flex items-center gap-3 pt-1">
                      <button
                        onClick={() => nameListInputRef.current?.click()}
                        className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 active:bg-blue-800 rounded-lg shadow-sm transition-all cursor-pointer"
                      >
                        <Upload className="w-4 h-4" /> 选择名单文件上传
                      </button>
                      <button
                        onClick={restoreDefaultNames}
                        className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-gray-700 dark:text-gray-300 bg-white dark:bg-gray-800 hover:bg-gray-100 dark:hover:bg-gray-750 rounded-lg border border-gray-200 dark:border-gray-700 transition-colors cursor-pointer"
                      >
                        <RotateCcw className="w-3.5 h-3.5" /> 恢复默认{DEFAULT_REFERENCE_NAMES.split('\n').filter(Boolean).length}人
                      </button>
                    </div>
                  </div>
                ) : null}

                <div className="relative">
                  <textarea
                    value={referenceNames}
                    onChange={(e) => setReferenceNames(e.target.value)}
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={handleTextareaDrop}
                    rows={8}
                    placeholder="输入、粘贴名单，或将 Excel / CSV / TXT 文件直接拖入此处...&#10;系统已内置 Jane Liu 与 Jane Z、Sunny Yang 与 Sunny Feng 姓氏智能区分"
                    className="w-full p-4 border border-gray-300 dark:border-gray-700 bg-gray-50/50 dark:bg-gray-900/60 rounded-xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm font-mono leading-relaxed text-gray-900 dark:text-gray-100 resize-y"
                  />
                  <div className="absolute right-3 bottom-3 text-[11px] text-gray-400 dark:text-gray-500 bg-white/80 dark:bg-gray-900/80 px-2 py-0.5 rounded backdrop-blur-sm pointer-events-none">
                    支持直接拖入文件 / 已自动保存本地
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Main Workspace Area */}
          <div className="bg-white dark:bg-gray-850 shadow-sm rounded-2xl overflow-hidden border border-gray-200 dark:border-gray-750 flex flex-col md:flex-row min-h-[520px]">

            {/* Left Panel: Upload Zone */}
            <div className="w-full md:w-5/12 bg-gray-50/70 dark:bg-gray-900/40 border-b md:border-b-0 md:border-r border-gray-200 dark:border-gray-750 p-6 flex flex-col">
              <div className="flex items-center justify-between mb-3">
                <span className="text-xs font-bold tracking-wider uppercase text-gray-500 dark:text-gray-400">
                  文件上传 / Upload
                </span>
                {uploadedFiles.length > 0 && (
                  <button
                    onClick={clearAllFiles}
                    className="text-xs font-medium text-red-500 hover:text-red-700 dark:text-red-400 dark:hover:text-red-300 hover:underline"
                  >
                    全部清空
                  </button>
                )}
              </div>

              {/* Drop Zone */}
              <div
                className="border-2 border-dashed border-gray-300 dark:border-gray-700 rounded-2xl p-6 flex flex-col items-center justify-center hover:border-blue-500 dark:hover:border-blue-500 hover:bg-blue-50/50 dark:hover:bg-blue-950/20 cursor-pointer transition-colors bg-white dark:bg-gray-850 mb-4 group text-center"
                onDragOver={handleDragOver}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
              >
                <div className="h-12 w-12 rounded-xl bg-blue-50 dark:bg-blue-950/50 text-blue-600 dark:text-blue-400 flex items-center justify-center mb-3 group-hover:scale-105 transition-transform">
                  <Upload className="h-6 w-6" />
                </div>
                <p className="text-sm font-semibold text-gray-800 dark:text-gray-200">
                  点击上传 或 拖拽文件到此处
                </p>
                <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">
                  支持截图、图片 (PNG/JPG)、PDF、Excel 表格
                </p>
                <p className="text-[11px] text-blue-600/80 dark:text-blue-400/80 mt-2 font-medium bg-blue-50 dark:bg-blue-900/30 px-2 py-0.5 rounded-md">
                  💡 支持直接 Ctrl+V / Cmd+V 粘贴截图
                </p>
                <input
                  type="file"
                  multiple
                  ref={fileInputRef}
                  className="hidden"
                  accept="image/*,.pdf,.xlsx,.xls,application/pdf,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
                  onChange={handleFileSelect}
                />
              </div>

              {/* Uploaded Files Gallery */}
              {uploadedFiles.length > 0 ? (
                <div className="flex-1 overflow-y-auto space-y-2">
                  <div className="text-xs font-semibold text-gray-500 dark:text-gray-400 mb-2">
                    已添加 {uploadedFiles.length} 个文件：
                  </div>
                  <div className="grid grid-cols-3 gap-2.5">
                    {uploadedFiles.map((uFile) => (
                      <div key={uFile.id} className="relative group">
                        {renderPreview(uFile)}
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            removeFile(uFile.id);
                          }}
                          className="absolute -top-1.5 -right-1.5 bg-red-500 hover:bg-red-600 text-white rounded-full p-1 shadow-md opacity-90 sm:opacity-0 sm:group-hover:opacity-100 transition-opacity"
                          title="删除此文件"
                        >
                          <X className="w-3 h-3" />
                        </button>
                      </div>
                    ))}
                    <div
                      onClick={() => fileInputRef.current?.click()}
                      className="h-20 w-20 border-2 border-dashed border-gray-200 dark:border-gray-700 rounded-lg flex flex-col items-center justify-center text-gray-400 hover:text-blue-500 hover:border-blue-400 cursor-pointer bg-white dark:bg-gray-850 transition-colors"
                      title="添加更多文件"
                    >
                      <Plus className="w-5 h-5 mb-0.5" />
                      <span className="text-[10px]">添加</span>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="flex-1 flex flex-col items-center justify-center text-gray-400 dark:text-gray-600 text-xs py-8">
                  <span>等待上传或粘贴课时表格/图片</span>
                </div>
              )}
            </div>

            {/* Right Panel: Actions & Output Table */}
            <div className="w-full md:w-7/12 p-6 flex flex-col relative bg-white dark:bg-gray-850">

              {/* Empty / Ready to Analyze Overlay */}
              {uploadedFiles.length > 0 && !extractedData && !isAnalyzing && (
                <div className="absolute inset-0 flex flex-col items-center justify-center z-10 bg-white/95 dark:bg-gray-850/95 backdrop-blur-xs p-6 text-center">
                  <div className="max-w-sm space-y-4">
                    <div className="h-14 w-14 rounded-2xl bg-blue-100 dark:bg-blue-900/40 text-blue-600 dark:text-blue-400 flex items-center justify-center mx-auto">
                      <FileSpreadsheet className="w-7 h-7" />
                    </div>
                    <div>
                      <h3 className="text-lg font-bold text-gray-900 dark:text-white">文件已就绪</h3>
                      <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">
                        已载入 {uploadedFiles.length} 个文件，将提取课时并按 {currentNamesList.length} 位标准教师名单汇总计算
                      </p>
                    </div>
                    <button
                      onClick={analyzeData}
                      className="w-full bg-blue-600 hover:bg-blue-700 active:scale-[0.99] text-white font-bold py-3.5 px-6 rounded-xl shadow-lg shadow-blue-500/25 flex items-center justify-center gap-2 text-base transition-all"
                    >
                      <FileSpreadsheet className="w-5 h-5" />
                      开始计算教师得分
                    </button>
                  </div>
                </div>
              )}

              {/* Analyzing Loader */}
              {isAnalyzing && (
                <div className="absolute inset-0 flex flex-col items-center justify-center z-20 bg-white/95 dark:bg-gray-850/95 backdrop-blur-xs p-6 text-center">
                  <Loader2 className="h-12 w-12 text-blue-600 dark:text-blue-400 animate-spin mb-4" />
                  <p className="text-base font-bold text-gray-900 dark:text-white">
                    正在分析 {uploadedFiles.length} 个文件...
                  </p>
                  <p className="text-xs text-gray-500 dark:text-gray-400 mt-1.5 max-w-xs">
                    正在识别教师与课时、匹配标准名单并按指定顺序计算得分
                  </p>
                </div>
              )}

              {/* Error Message */}
              {error && (
                <div className="mb-4 p-4 bg-red-50 dark:bg-red-950/40 border border-red-200 dark:border-red-800/80 rounded-xl text-red-700 dark:text-red-300 flex items-start gap-3">
                  <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
                  <div className="flex-1">
                    <p className="text-sm font-semibold">分析遇到问题</p>
                    <p className="text-xs mt-0.5">{error}</p>
                  </div>
                  <button
                    onClick={analyzeData}
                    className="text-xs font-semibold text-red-700 dark:text-red-300 underline hover:no-underline"
                  >
                    重试
                  </button>
                </div>
              )}

              {/* Results View */}
              {extractedData && (
                <div className="flex flex-col h-full w-full space-y-4">

                  {/* Results Header & Summary Stats */}
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-gray-100 dark:border-gray-750">
                    <div className="flex items-center gap-2">
                      <CheckCircle className="w-5 h-5 text-emerald-600 dark:text-emerald-400" />
                      <div>
                        <h2 className="text-sm font-bold text-gray-900 dark:text-white">
                          计算结果（共 {totalTeachers} 位教师）
                        </h2>
                        <div className="flex items-center gap-2 text-xs text-gray-500 dark:text-gray-400 mt-0.5">
                          <span>已匹配课时: <strong className="text-blue-600 dark:text-blue-400">{activeTeachers}</strong> 位</span>
                          <span>•</span>
                          <span>5分: <strong className="text-emerald-600 dark:text-emerald-400">{score5Teachers}</strong> 位</span>
                          <span>•</span>
                          <span>3分: <strong className="text-blue-600 dark:text-blue-400">{score3Teachers}</strong> 位</span>
                          <span>•</span>
                          <span>1分: <strong className="text-amber-600 dark:text-amber-400">{score1Teachers}</strong> 位</span>
                          <span>•</span>
                          <span>0分: <strong>{score0Teachers}</strong> 位</span>
                        </div>
                      </div>
                    </div>

                    <div className="flex items-center gap-2">
                      {unmatchedRows.length > 0 && (
                        <button
                          onClick={() => setShowUnmatched(!showUnmatched)}
                          className={`text-xs px-2.5 py-1.5 rounded-lg flex items-center gap-1 font-medium transition-colors ${
                            showUnmatched
                              ? 'bg-amber-100 text-amber-800 dark:bg-amber-900/60 dark:text-amber-200'
                              : 'bg-gray-100 text-gray-600 dark:bg-gray-800 dark:text-gray-300 hover:bg-gray-200 dark:hover:bg-gray-700'
                          }`}
                        >
                          {showUnmatched ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                          {unmatchedRows.length} 条未匹配
                        </button>
                      )}
                      <button
                        onClick={analyzeData}
                        className="text-xs px-2.5 py-1.5 rounded-lg bg-gray-100 hover:bg-gray-200 dark:bg-gray-800 dark:hover:bg-gray-700 text-gray-700 dark:text-gray-300 font-medium flex items-center gap-1"
                        title="重新计算"
                      >
                        <RefreshCw className="w-3.5 h-3.5" /> 重新计算
                      </button>
                    </div>
                  </div>

                  {/* Unmatched Rows Warning */}
                  {showUnmatched && unmatchedRows.length > 0 && (
                    <div className="bg-amber-50/80 dark:bg-amber-950/30 border border-amber-200 dark:border-amber-800/60 rounded-xl p-3 text-xs overflow-auto max-h-[140px]">
                      <div className="font-semibold text-amber-900 dark:text-amber-300 mb-1.5 flex items-center justify-between">
                        <span>未在标准名单中匹配到的原始数据项：</span>
                        <span className="text-[11px] text-amber-700 dark:text-amber-400 font-normal">可检查标准名单中是否有拼写差异</span>
                      </div>
                      <table className="w-full text-left text-amber-950 dark:text-amber-200">
                        <thead>
                          <tr className="border-b border-amber-200/60 dark:border-amber-800/40">
                            <th className="pb-1">识别到姓名</th>
                            <th className="pb-1">课时</th>
                            <th className="pb-1">来源文件</th>
                          </tr>
                        </thead>
                        <tbody>
                          {unmatchedRows.map((r, i) => (
                            <tr key={i} className="border-t border-amber-100 dark:border-amber-900/30">
                              <td className="py-1 pr-2 font-mono">{r.teacher}</td>
                              <td className="py-1 font-bold">{r.count}</td>
                              <td className="py-1 text-amber-600 dark:text-amber-400/80 text-[10px] truncate max-w-[120px]">{r.sourceFile}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}

                  {/* Teacher Scores Table */}
                  <div className="flex-1 overflow-auto border border-gray-200 dark:border-gray-750 rounded-xl bg-gray-50/40 dark:bg-gray-900/40 max-h-[420px] scrollbar-hide">
                    <table className="min-w-full divide-y divide-gray-200 dark:divide-gray-750 text-xs sm:text-sm">
                      <thead className="bg-gray-100 dark:bg-gray-800 sticky top-0 z-10 select-none">
                        <tr>
                          <th className="px-3 py-2.5 text-center text-gray-500 dark:text-gray-400 font-semibold w-12">序号</th>
                          <th className="px-3 py-2.5 text-left text-gray-600 dark:text-gray-300 font-semibold">教师姓名</th>
                          <th className="px-3 py-2.5 text-center text-gray-600 dark:text-gray-300 font-semibold">课时 (Count)</th>
                          <th className="px-3 py-2.5 text-center text-gray-600 dark:text-gray-300 font-semibold">得分 (Score)</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-200 dark:divide-gray-750 bg-white dark:bg-gray-850">
                        {extractedData.map((row, idx) => {
                          const isScore5 = row["得分"] === 5;
                          const isScore3 = row["得分"] === 3;
                          const isScore1 = row["得分"] === 1;
                          const isZero = row["课时"] === 0;

                          return (
                            <tr
                              key={row.id}
                              className={`hover:bg-blue-50/40 dark:hover:bg-blue-950/20 transition-colors ${
                                isZero ? "opacity-60" : ""
                              }`}
                            >
                              <td className="px-3 py-2 text-center text-gray-400 dark:text-gray-500 font-mono text-xs">
                                {row["序号"]}
                              </td>
                              <td className="px-3 py-2 whitespace-nowrap font-medium text-gray-900 dark:text-gray-100">
                                {row["姓名"]}
                              </td>
                              <td className="px-3 py-2 text-center">
                                {editingRowIndex === idx ? (
                                  <div className="inline-flex items-center gap-1">
                                    <input
                                      type="number"
                                      value={editingCountValue}
                                      onChange={(e) => setEditingCountValue(e.target.value)}
                                      onKeyDown={(e) => {
                                        if (e.key === 'Enter') saveEditing(idx);
                                        if (e.key === 'Escape') setEditingRowIndex(null);
                                      }}
                                      className="w-16 px-1.5 py-0.5 text-xs text-center border border-blue-500 rounded bg-white dark:bg-gray-900 font-mono text-gray-900 dark:text-white"
                                      autoFocus
                                    />
                                    <button
                                      onClick={() => saveEditing(idx)}
                                      className="p-1 text-emerald-600 hover:text-emerald-700"
                                      title="保存"
                                    >
                                      <Check className="w-3.5 h-3.5" />
                                    </button>
                                  </div>
                                ) : (
                                  <button
                                    onClick={() => startEditing(idx, row["课时"])}
                                    className="font-mono text-gray-700 dark:text-gray-300 hover:text-blue-600 dark:hover:text-blue-400 group inline-flex items-center gap-1"
                                    title="点击修改课时"
                                  >
                                    <span>{row["课时"]}</span>
                                    <Edit2 className="w-3 h-3 opacity-0 group-hover:opacity-60 text-gray-400" />
                                  </button>
                                )}
                              </td>
                              <td className="px-3 py-2 text-center">
                                <span
                                  className={`inline-flex items-center justify-center min-w-[28px] px-2 py-0.5 rounded-full text-xs font-bold ${
                                    isScore5
                                      ? "bg-emerald-100 text-emerald-800 dark:bg-emerald-900/60 dark:text-emerald-300 ring-1 ring-emerald-300 dark:ring-emerald-700"
                                      : isScore3
                                      ? "bg-blue-100 text-blue-800 dark:bg-blue-900/60 dark:text-blue-300 ring-1 ring-blue-300 dark:ring-blue-700"
                                      : isScore1
                                      ? "bg-amber-100 text-amber-800 dark:bg-amber-900/50 dark:text-amber-300 ring-1 ring-amber-300 dark:ring-amber-700"
                                      : "bg-gray-100 text-gray-500 dark:bg-gray-800 dark:text-gray-400"
                                  }`}
                                >
                                  {row["得分"]} 分
                                </span>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>

                  {/* Export and Action Buttons */}
                  <div className="flex flex-col sm:flex-row items-center gap-3 pt-2">
                    <button
                      onClick={downloadExcel}
                      className="w-full sm:flex-1 bg-emerald-600 hover:bg-emerald-700 active:scale-[0.99] text-white font-bold py-3 px-5 rounded-xl shadow-md shadow-emerald-600/20 flex items-center justify-center gap-2 text-sm transition-all"
                    >
                      <Download className="w-4 h-4" />
                      生成并下载 Excel 表格 (.xlsx)
                    </button>
                    <button
                      onClick={copyTableToClipboard}
                      className="w-full sm:w-auto px-5 py-3 rounded-xl border border-gray-300 dark:border-gray-700 hover:bg-gray-100 dark:hover:bg-gray-800 font-medium text-xs sm:text-sm text-gray-700 dark:text-gray-300 flex items-center justify-center gap-2 transition-colors"
                      title="复制表格内容（可直接粘贴到 Excel 或 WPS 中）"
                    >
                      {copied ? (
                        <>
                          <Check className="w-4 h-4 text-emerald-600 dark:text-emerald-400" />
                          <span className="text-emerald-600 dark:text-emerald-400 font-semibold">已复制到剪贴板</span>
                        </>
                      ) : (
                        <>
                          <Copy className="w-4 h-4" />
                          <span>复制表格数据</span>
                        </>
                      )}
                    </button>
                  </div>
                </div>
              )}

              {/* Initial Empty Placeholder when no files uploaded */}
              {uploadedFiles.length === 0 && (
                <div className="flex-1 flex flex-col items-center justify-center text-center p-8 text-gray-400 dark:text-gray-600">
                  <div className="w-16 h-16 rounded-2xl bg-gray-100 dark:bg-gray-800 flex items-center justify-center mb-4">
                    <FileSpreadsheet className="w-8 h-8 text-gray-400 dark:text-gray-600" />
                  </div>
                  <h3 className="text-base font-semibold text-gray-700 dark:text-gray-300 mb-1">
                    暂未载入数据
                  </h3>
                  <p className="text-xs text-gray-500 dark:text-gray-500 max-w-xs">
                    请在左侧上传或粘贴教师课时图片、PDF 或 Excel 表格，系统将自动汇总并按指定顺序计算得分
                  </p>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

const root = createRoot(document.getElementById('root')!);
root.render(<App />);
