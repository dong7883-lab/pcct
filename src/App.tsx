/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Users, Calendar, DoorOpen, FileSpreadsheet, FileText, AlertCircle, CheckCircle2, Filter, ArrowUpDown, Info, Type, Moon, Sun, ChevronLeft, ChevronRight } from 'lucide-react';
import { Invigilator, AssignmentResult, generateSchedule, compareVietnameseName } from './utils/scheduler';
import { exportToExcel, exportToDocx, exportM3ToDocx, exportM14ToDocx } from './utils/export';
import { ExportPreviewModal } from './components/ExportPreviewModal';

export default function App() {
  const loadSavedState = <T,>(key: string, defaultValue: T): T => {
    try {
      const saved = localStorage.getItem(key);
      if (saved) {
        return JSON.parse(saved);
      }
    } catch (e) {
      console.error(`Error loading ${key} from localStorage`, e);
    }
    return defaultValue;
  };

  const [invigilators, setInvigilators] = useState<Invigilator[]>(() => loadSavedState('invigilators', []));
  const [numShifts, setNumShifts] = useState<number>(() => loadSavedState('numShifts', 4));
  const [shiftNames, setShiftNames] = useState<string[]>(() => loadSavedState('shiftNames', Array.from({ length: 4 }, (_, i) => `Ca thi ${i + 1}`)));
  const [numRooms, setNumRooms] = useState<number>(() => loadSavedState('numRooms', 10));
  const [roomPrefix, setRoomPrefix] = useState<string>(() => loadSavedState('roomPrefix', 'Phòng '));
  const [roomStartNumber, setRoomStartNumber] = useState<number>(() => loadSavedState('roomStartNumber', 1));
  const [roomNamesText, setRoomNamesText] = useState<string>(() => loadSavedState('roomNamesText', ''));
  const [invigilatorsPerRoom, setInvigilatorsPerRoom] = useState<number>(() => loadSavedState('invigilatorsPerRoom', 2));
  const [rules, setRules] = useState(() => loadSavedState('rules', {
    minimizeConsecutive: false,
    fixedPairs: false,
    balanceExperience: false,
    avoidPairs: [] as [string, string][]
  }));
  const [result, setResult] = useState<AssignmentResult | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [fileName, setFileName] = useState<string>('');
  
  // Pagination state
  const [pageByShift, setPageByShift] = useState<Record<number, number>>({});
  const ITEMS_PER_PAGE = 10;
  
  // Input method states
  const [inputType, setInputType] = useState<'excel' | 'text' | 'manual'>(() => loadSavedState('inputType', 'excel'));
  const [manualText, setManualText] = useState<string>(() => loadSavedState('manualText', ''));
  const [newInvigilatorId, setNewInvigilatorId] = useState('');
  const [newInvigilatorName, setNewInvigilatorName] = useState('');
  const [avoidInv1, setAvoidInv1] = useState('');
  const [avoidInv2, setAvoidInv2] = useState('');
  const [isDarkMode, setIsDarkMode] = useState(() => {
    const savedTheme = localStorage.getItem('theme');
    if (savedTheme) {
      return savedTheme === 'dark';
    }
    return window.matchMedia('(prefers-color-scheme: dark)').matches;
  });
  
  // Preview Modal State
  const [previewConfig, setPreviewConfig] = useState<{
    isOpen: boolean;
    type: 'schedule' | 'm3' | 'm14' | null;
  }>({ isOpen: false, type: null });

  // Save states to localStorage
  useEffect(() => { localStorage.setItem('invigilators', JSON.stringify(invigilators)); }, [invigilators]);
  useEffect(() => { localStorage.setItem('numShifts', JSON.stringify(numShifts)); }, [numShifts]);
  useEffect(() => { localStorage.setItem('shiftNames', JSON.stringify(shiftNames)); }, [shiftNames]);
  useEffect(() => { localStorage.setItem('numRooms', JSON.stringify(numRooms)); }, [numRooms]);
  useEffect(() => { localStorage.setItem('roomPrefix', JSON.stringify(roomPrefix)); }, [roomPrefix]);
  useEffect(() => { localStorage.setItem('roomStartNumber', JSON.stringify(roomStartNumber)); }, [roomStartNumber]);
  useEffect(() => { localStorage.setItem('roomNamesText', JSON.stringify(roomNamesText)); }, [roomNamesText]);
  useEffect(() => { localStorage.setItem('invigilatorsPerRoom', JSON.stringify(invigilatorsPerRoom)); }, [invigilatorsPerRoom]);
  useEffect(() => { localStorage.setItem('rules', JSON.stringify(rules)); }, [rules]);
  useEffect(() => { localStorage.setItem('inputType', JSON.stringify(inputType)); }, [inputType]);
  useEffect(() => { localStorage.setItem('manualText', JSON.stringify(manualText)); }, [manualText]);

  useEffect(() => {
    if (isDarkMode) {
      document.documentElement.classList.add('dark');
      localStorage.setItem('theme', 'dark');
    } else {
      document.documentElement.classList.remove('dark');
      localStorage.setItem('theme', 'light');
    }
  }, [isDarkMode]);

  const handleAddManualInvigilator = () => {
    if (!newInvigilatorName.trim()) return;
    const id = newInvigilatorId.trim() || `GT${String(invigilators.length + 1).padStart(3, '0')}`;
    const newInv = { id, name: newInvigilatorName.trim() };
    setInvigilators([...invigilators, newInv]);
    setNewInvigilatorId('');
    setNewInvigilatorName('');
    setResult(null);
  };

  const handleRemoveManualInvigilator = (idToRemove: string) => {
    const updated = invigilators.filter(inv => inv.id !== idToRemove);
    setInvigilators(updated);
    setResult(null);
  };

  const handleAddAvoidPair = () => {
    if (!avoidInv1 || !avoidInv2 || avoidInv1 === avoidInv2) return;
    // Check if pair already exists
    const exists = rules.avoidPairs.some(([id1, id2]) => 
      (id1 === avoidInv1 && id2 === avoidInv2) || (id1 === avoidInv2 && id2 === avoidInv1)
    );
    if (!exists) {
      setRules({
        ...rules,
        avoidPairs: [...rules.avoidPairs, [avoidInv1, avoidInv2]]
      });
    }
    setAvoidInv1('');
    setAvoidInv2('');
  };

  const handleRemoveAvoidPair = (index: number) => {
    const newPairs = [...rules.avoidPairs];
    newPairs.splice(index, 1);
    setRules({
      ...rules,
      avoidPairs: newPairs
    });
  };

  // Filter and Sort states
  const [filterInvigilator, setFilterInvigilator] = useState<string>('all');
  const [filterRoom, setFilterRoom] = useState<string>('all');
  const [filterShift, setFilterShift] = useState<string>('all');
  const [sortBy, setSortBy] = useState<'room' | 'invigilator'>('room');

  useEffect(() => {
    setPageByShift({});
  }, [filterInvigilator, filterRoom, filterShift, sortBy]);

  const handleNumShiftsChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const val = parseInt(e.target.value) || 0;
    setNumShifts(val);
    setShiftNames(prev => {
      if (val > prev.length) {
        return [...prev, ...Array.from({ length: val - prev.length }, (_, i) => `Ca thi ${prev.length + i + 1}`)];
      } else {
        return prev.slice(0, val);
      }
    });
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json<any>(ws, { header: 1 });
      
      const parsedInvigilators: Invigilator[] = [];
      
      if (data.length > 1) {
        let nameColIdx = 0;
        let idColIdx = -1;
        let expColIdx = -1;
        const headerRow = data[0] as string[];
        
        headerRow.forEach((col, idx) => {
          const colLower = String(col).toLowerCase();
          if (colLower.includes('tên') || colLower.includes('name')) {
            nameColIdx = idx;
          }
          if (colLower.includes('mã') || colLower.includes('id')) {
            idColIdx = idx;
          }
          if (colLower.includes('kinh nghiệm') || colLower.includes('exp') || colLower.includes('năm')) {
            expColIdx = idx;
          }
        });

        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (row && row[nameColIdx]) {
            parsedInvigilators.push({
              id: idColIdx !== -1 && row[idColIdx] ? String(row[idColIdx]) : `GT${String(i).padStart(3, '0')}`,
              name: String(row[nameColIdx]).trim(),
              experience: expColIdx !== -1 && row[expColIdx] ? parseInt(row[expColIdx]) || 0 : 0
            });
          }
        }
      }
      
      setInvigilators(parsedInvigilators);
      setResult(null);
      setFilterInvigilator('all');
      setFilterRoom('all');
      setFilterShift('all');
    };
    reader.readAsBinaryString(file);
  };

  const handleRandomizeRooms = (shiftIdx: number) => {
    if (!result || !result.schedule) return;

    const newSchedule = [...result.schedule];
    const shift = [...newSchedule[shiftIdx]];
    
    // Extract all invigilators from this shift
    const invigilatorsInShift = shift.map(a => ({
      inv1: a.invigilator1,
      inv2: a.invigilator2
    }));

    // Shuffle the array of invigilator pairs/singles
    for (let i = invigilatorsInShift.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [invigilatorsInShift[i], invigilatorsInShift[j]] = [invigilatorsInShift[j], invigilatorsInShift[i]];
    }

    // Reassign to rooms
    shift.forEach((assignment, idx) => {
      assignment.invigilator1 = invigilatorsInShift[idx].inv1;
      assignment.invigilator2 = invigilatorsInShift[idx].inv2;
    });

    newSchedule[shiftIdx] = shift;
    setResult({ ...result, schedule: newSchedule });
  };

  const handleManualTextChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    const text = e.target.value;
    setManualText(text);
    
    if (!text.trim()) {
      setInvigilators([]);
      setResult(null);
      return;
    }

    const lines = text.split('\n').filter(line => line.trim() !== '');
    const parsedInvigilators: Invigilator[] = [];
    
    lines.forEach((line, i) => {
      const parts = line.split(/\t|,/).map(p => p.trim()).filter(p => p !== '');
      if (parts.length >= 3) {
        parsedInvigilators.push({
          id: parts[0],
          name: parts[1],
          experience: parseInt(parts[2]) || 0
        });
      } else if (parts.length === 2) {
        parsedInvigilators.push({
          id: parts[0],
          name: parts[1],
          experience: 0
        });
      } else if (parts.length === 1) {
        parsedInvigilators.push({
          id: `GT${String(i + 1).padStart(3, '0')}`,
          name: parts[0],
          experience: 0
        });
      }
    });

    setInvigilators(parsedInvigilators);
    setResult(null);
    setFilterInvigilator('all');
    setFilterRoom('all');
    setFilterShift('all');
  };

  const currentRoomNames = useMemo(() => {
    const customNames = roomNamesText.split(/[\n,;]+/).map(s => s.trim()).filter(s => s);
    const padLength = numRooms.toString().length;
    return Array.from({ length: Math.max(numRooms, customNames.length) }, (_, i) => {
      if (customNames[i]) return customNames[i];
      return `${roomPrefix}${(roomStartNumber + i).toString().padStart(padLength, '0')}`;
    });
  }, [numRooms, roomNamesText, roomPrefix, roomStartNumber]);

  const handleGenerate = () => {
    if (invigilators.length === 0) {
      setResult({ error: 'Vui lòng nhập danh sách giám thị.' });
      return;
    }
    
    setIsGenerating(true);
    setResult(null);
    setFilterInvigilator('all');
    setFilterRoom('all');
    setFilterShift('all');
    
    setTimeout(() => {
      const res = generateSchedule(invigilators, numShifts, numRooms, invigilatorsPerRoom, rules);
      setResult(res);
      setIsGenerating(false);
    }, 100);
  };

  const allRooms = useMemo(() => {
    if (!result?.schedule) return [];
    const rooms = new Set<number>();
    result.schedule.forEach(shift => shift.forEach(a => rooms.add(a.room)));
    return Array.from(rooms).sort((a, b) => a - b);
  }, [result]);

  const sortedInvigilators = useMemo(() => {
    return [...invigilators].sort((a, b) => compareVietnameseName(a.name, b.name));
  }, [invigilators]);

  const getUnassignedInvigilators = (shift: any[]) => {
    const assignedIds = new Set<string>();
    shift.forEach(a => {
      assignedIds.add(a.invigilator1.id);
      if (a.invigilator2) {
        assignedIds.add(a.invigilator2.id);
      }
    });
    return sortedInvigilators.filter(inv => !assignedIds.has(inv.id));
  };

  const validationWarning = useMemo(() => {
    if (invigilators.length === 0) return null;
    const N = invigilators.length;
    const R = numRooms;
    
    if (N < invigilatorsPerRoom * R) {
      return `Thiếu giám thị! Cần ít nhất ${invigilatorsPerRoom * R} giám thị cho ${R} phòng thi, nhưng chỉ có ${N} người.`;
    }
    
    return null;
  }, [invigilators, numRooms, invigilatorsPerRoom]);

  const { idealMax, idealMin } = useMemo(() => {
    if (invigilators.length === 0) return { idealMax: 0, idealMin: 0 };
    const totalSlots = invigilatorsPerRoom * numRooms * numShifts;
    return {
      idealMax: Math.ceil(totalSlots / invigilators.length),
      idealMin: Math.floor(totalSlots / invigilators.length)
    };
  }, [invigilators.length, numRooms, numShifts, invigilatorsPerRoom]);

  return (
    <div className="min-h-screen bg-slate-50 dark:bg-slate-900 text-slate-900 dark:text-slate-100 font-sans transition-colors duration-200">
      <header className="bg-white dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700 px-6 py-4 sticky top-0 z-10 transition-colors duration-200">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-600 dark:bg-indigo-500 p-2 rounded-lg">
              <Users className="w-6 h-6 text-white" />
            </div>
            <h1 className="text-xl font-semibold text-slate-800 dark:text-white">Hệ thống Phân công Giám thị</h1>
          </div>
          <button
            onClick={() => setIsDarkMode(!isDarkMode)}
            className="p-2 rounded-lg bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600 transition-colors"
            title={isDarkMode ? "Chuyển sang giao diện sáng" : "Chuyển sang giao diện tối"}
          >
            {isDarkMode ? <Sun className="w-5 h-5" /> : <Moon className="w-5 h-5" />}
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 py-8">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
          
          <div className="lg:col-span-4 space-y-6">
            <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-700 p-6 transition-colors duration-200">
              <h2 className="text-lg font-medium mb-4 flex items-center gap-2 text-slate-800 dark:text-white">
                <Calendar className="w-5 h-5 text-slate-500 dark:text-slate-400" />
                Cấu hình kỳ thi
              </h2>
              
              <div className="space-y-4">
                <div>
                  <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-2">
                    Phương thức nhập danh sách
                  </label>
                  <div className="flex bg-slate-100 dark:bg-slate-900/50 p-1 rounded-lg mb-4">
                    <button
                      onClick={() => setInputType('excel')}
                      className={`flex-1 py-2 text-sm font-medium rounded-md transition-colors ${inputType === 'excel' ? 'bg-white dark:bg-slate-800 text-indigo-600 dark:text-indigo-400 shadow-sm' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'}`}
                    >
                      <div className="flex items-center justify-center gap-2">
                        <FileSpreadsheet className="w-4 h-4" />
                        File Excel
                      </div>
                    </button>
                    <button
                      onClick={() => setInputType('text')}
                      className={`flex-1 py-2 text-sm font-medium rounded-md transition-colors ${inputType === 'text' ? 'bg-white dark:bg-slate-800 text-indigo-600 dark:text-indigo-400 shadow-sm' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'}`}
                    >
                      <div className="flex items-center justify-center gap-2">
                        <Type className="w-4 h-4" />
                        Dán văn bản
                      </div>
                    </button>
                    <button
                      onClick={() => setInputType('manual')}
                      className={`flex-1 py-2 text-sm font-medium rounded-md transition-colors ${inputType === 'manual' ? 'bg-white dark:bg-slate-800 text-indigo-600 dark:text-indigo-400 shadow-sm' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'}`}
                    >
                      <div className="flex items-center justify-center gap-2">
                        <Users className="w-4 h-4" />
                        Nhập thủ công
                      </div>
                    </button>
                  </div>

                  {inputType === 'excel' ? (
                    <label className="flex items-center justify-center w-full h-32 px-4 transition bg-white dark:bg-slate-800 border-2 border-slate-300 dark:border-slate-600 border-dashed rounded-xl appearance-none cursor-pointer hover:border-indigo-400 dark:hover:border-indigo-500 focus:outline-none">
                      <span className="flex items-center space-x-2">
                        <Upload className="w-6 h-6 text-slate-400 dark:text-slate-500" />
                        <span className="font-medium text-slate-600 dark:text-slate-300">
                          {fileName ? fileName : 'Chọn file Excel'}
                        </span>
                      </span>
                      <input type="file" name="file_upload" className="hidden" accept=".xlsx, .xls" onChange={handleFileUpload} />
                    </label>
                  ) : inputType === 'text' ? (
                    <textarea
                      value={manualText}
                      onChange={handleManualTextChange}
                      placeholder="Dán danh sách giám thị vào đây...&#10;Mỗi dòng 1 người (Họ tên) hoặc (Mã GT [tab/phẩy] Họ tên)"
                      className="w-full h-32 p-3 text-sm border border-slate-300 dark:border-slate-600 rounded-xl focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 resize-none"
                    />
                  ) : (
                    <div className="space-y-3">
                      <div className="flex gap-2">
                        <input
                          type="text"
                          value={newInvigilatorId}
                          onChange={(e) => setNewInvigilatorId(e.target.value)}
                          placeholder="Mã GT (Tùy chọn)"
                          className="w-1/3 px-3 py-2 text-sm border border-slate-300 dark:border-slate-600 rounded-lg focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100"
                        />
                        <input
                          type="text"
                          value={newInvigilatorName}
                          onChange={(e) => setNewInvigilatorName(e.target.value)}
                          onKeyDown={(e) => e.key === 'Enter' && handleAddManualInvigilator()}
                          placeholder="Họ và tên"
                          className="flex-1 px-3 py-2 text-sm border border-slate-300 dark:border-slate-600 rounded-lg focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100"
                        />
                        <button
                          onClick={handleAddManualInvigilator}
                          disabled={!newInvigilatorName.trim()}
                          className="px-3 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 disabled:bg-indigo-300 dark:disabled:bg-indigo-800/50"
                        >
                          Thêm
                        </button>
                      </div>
                      
                      {sortedInvigilators.length > 0 && (
                        <div className="flex justify-end mb-1">
                          <button
                            onClick={() => {
                              if (window.confirm('Bạn có chắc chắn muốn xóa toàn bộ danh sách giám thị không?')) {
                                setInvigilators([]);
                              }
                            }}
                            className="text-xs text-red-600 hover:text-red-700 dark:text-red-400 dark:hover:text-red-300 font-medium"
                          >
                            Xóa tất cả
                          </button>
                        </div>
                      )}
                      
                      {sortedInvigilators.length > 0 && (
                        <div className="max-h-40 overflow-y-auto border border-slate-200 dark:border-slate-700 rounded-lg divide-y divide-slate-100 dark:divide-slate-700/50">
                          {sortedInvigilators.map((inv) => (
                            <div key={inv.id} className="flex items-center justify-between px-3 py-2 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700/50">
                              <div className="flex items-center gap-3">
                                <span className="text-xs font-medium text-slate-500 dark:text-slate-400 w-12">{inv.id}</span>
                                <span className="text-sm text-slate-700 dark:text-slate-300">{inv.name}</span>
                              </div>
                              <button
                                onClick={() => handleRemoveManualInvigilator(inv.id)}
                                className="text-red-500 hover:text-red-700 dark:hover:text-red-400 p-1"
                              >
                                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M18 6 6 18"/><path d="m6 6 12 12"/></svg>
                              </button>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                  
                  {invigilators.length > 0 && (
                    <p className="mt-2 text-sm text-emerald-600 dark:text-emerald-400 flex items-center gap-1">
                      <CheckCircle2 className="w-4 h-4" />
                      Đã tải {invigilators.length} giám thị
                    </p>
                  )}
                </div>

                <div>
                  <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                    Số ca thi
                  </label>
                  <div className="relative">
                    <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                      <Calendar className="h-5 w-5 text-slate-400 dark:text-slate-500" />
                    </div>
                    <input
                      type="number"
                      min="1"
                      value={numShifts}
                      onChange={handleNumShiftsChange}
                      className="block w-full pl-10 pr-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 sm:text-sm"
                    />
                  </div>
                </div>

                {numShifts > 0 && (
                  <div>
                    <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                      Tên các ca thi (Tùy chọn)
                    </label>
                    <div className="space-y-2 mt-2 max-h-48 overflow-y-auto pr-2">
                      {shiftNames.map((name, idx) => (
                        <input
                          key={idx}
                          type="text"
                          value={name}
                          onChange={(e) => {
                            const newNames = [...shiftNames];
                            newNames[idx] = e.target.value;
                            setShiftNames(newNames);
                          }}
                          placeholder={`Ca thi ${idx + 1}`}
                          className="block w-full px-3 py-1.5 text-sm border border-slate-300 dark:border-slate-600 rounded-md focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100"
                        />
                      ))}
                    </div>
                  </div>
                )}

                <div>
                  <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                    Số phòng thi
                  </label>
                  <div className="relative">
                    <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                      <DoorOpen className="h-5 w-5 text-slate-400 dark:text-slate-500" />
                    </div>
                    <input
                      type="number"
                      min="1"
                      value={numRooms}
                      onChange={(e) => setNumRooms(parseInt(e.target.value) || 0)}
                      className="block w-full pl-10 pr-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 sm:text-sm"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                      Tiền tố tên phòng
                    </label>
                    <input
                      type="text"
                      value={roomPrefix}
                      onChange={(e) => setRoomPrefix(e.target.value)}
                      placeholder="VD: Phòng , P."
                      className="block w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 sm:text-sm"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                      Đánh số từ
                    </label>
                    <input
                      type="number"
                      value={roomStartNumber}
                      onChange={(e) => setRoomStartNumber(parseInt(e.target.value) || 0)}
                      className="block w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 sm:text-sm"
                    />
                  </div>
                </div>

                {numRooms > 0 && (
                  <div>
                    <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                      Tên phòng thi tùy chỉnh (Tùy chọn)
                    </label>
                    <textarea
                      value={roomNamesText}
                      onChange={(e) => setRoomNamesText(e.target.value)}
                      placeholder="Nhập tên phòng (mỗi phòng 1 dòng hoặc cách nhau bởi dấu phẩy)...&#10;Ví dụ: P.101, P.102, P.103"
                      className="w-full h-24 p-3 text-sm border border-slate-300 dark:border-slate-600 rounded-xl focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 resize-none"
                    />
                  </div>
                )}

                <div>
                  <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-2">
                    Số giám thị mỗi phòng
                  </label>
                  <div className="flex bg-slate-100 dark:bg-slate-900/50 p-1 rounded-lg">
                    <button
                      onClick={() => setInvigilatorsPerRoom(1)}
                      className={`flex-1 py-2 text-sm font-medium rounded-md transition-colors ${invigilatorsPerRoom === 1 ? 'bg-white dark:bg-slate-800 text-indigo-600 dark:text-indigo-400 shadow-sm' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'}`}
                    >
                      1 Giám thị
                    </button>
                    <button
                      onClick={() => setInvigilatorsPerRoom(2)}
                      className={`flex-1 py-2 text-sm font-medium rounded-md transition-colors ${invigilatorsPerRoom === 2 ? 'bg-white dark:bg-slate-800 text-indigo-600 dark:text-indigo-400 shadow-sm' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'}`}
                    >
                      2 Giám thị
                    </button>
                  </div>
                </div>

                <div>
                  <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-2">
                    Tùy chỉnh quy tắc phân công
                  </label>
                  <div className="flex flex-col gap-2">
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={rules.minimizeConsecutive}
                        onChange={(e) => setRules({...rules, minimizeConsecutive: e.target.checked})}
                        className="rounded text-indigo-600 focus:ring-indigo-500 dark:bg-slate-800 dark:border-slate-600"
                      />
                      <span className="text-sm text-slate-700 dark:text-slate-300">Giảm thiểu gác liên tục nhiều ca</span>
                    </label>
                    {invigilatorsPerRoom > 1 && (
                      <label className="flex items-center gap-2 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={rules.fixedPairs}
                          onChange={(e) => setRules({...rules, fixedPairs: e.target.checked})}
                          className="rounded text-indigo-600 focus:ring-indigo-500 dark:bg-slate-800 dark:border-slate-600"
                        />
                        <span className="text-sm text-slate-700 dark:text-slate-300">Ưu tiên phân công theo cặp cố định</span>
                      </label>
                    )}
                    {invigilatorsPerRoom > 1 && (
                      <label className="flex items-center gap-2 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={rules.balanceExperience}
                          onChange={(e) => setRules({...rules, balanceExperience: e.target.checked})}
                          className="rounded text-indigo-600 focus:ring-indigo-500 dark:bg-slate-800 dark:border-slate-600"
                        />
                        <span className="text-sm text-slate-700 dark:text-slate-300">Ưu tiên ghép người có kinh nghiệm với người mới (Yêu cầu nhập cột Kinh nghiệm)</span>
                      </label>
                    )}
                  </div>
                </div>

                {invigilatorsPerRoom > 1 && invigilators.length > 0 && (
                  <div>
                    <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-2">
                      Giám thị không làm việc cùng nhau
                    </label>
                    <div className="flex flex-col sm:flex-row gap-2 mb-3">
                      <select
                        value={avoidInv1}
                        onChange={(e) => setAvoidInv1(e.target.value)}
                        className="flex-1 block w-full pl-3 pr-10 py-2 text-sm border border-slate-300 dark:border-slate-600 rounded-md focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100"
                      >
                        <option value="">Chọn giám thị 1</option>
                        {sortedInvigilators.map(inv => (
                          <option key={inv.id} value={inv.id}>{inv.name} ({inv.id})</option>
                        ))}
                      </select>
                      <select
                        value={avoidInv2}
                        onChange={(e) => setAvoidInv2(e.target.value)}
                        className="flex-1 block w-full pl-3 pr-10 py-2 text-sm border border-slate-300 dark:border-slate-600 rounded-md focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100"
                      >
                        <option value="">Chọn giám thị 2</option>
                        {sortedInvigilators.map(inv => (
                          <option key={inv.id} value={inv.id}>{inv.name} ({inv.id})</option>
                        ))}
                      </select>
                      <button
                        onClick={handleAddAvoidPair}
                        disabled={!avoidInv1 || !avoidInv2 || avoidInv1 === avoidInv2}
                        className="inline-flex justify-center items-center px-4 py-2 border border-transparent text-sm font-medium rounded-md text-white bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 disabled:bg-indigo-400 disabled:cursor-not-allowed"
                      >
                        Thêm
                      </button>
                    </div>
                    
                    {rules.avoidPairs && rules.avoidPairs.length > 0 && (
                      <ul className="space-y-2 max-h-40 overflow-y-auto">
                        {rules.avoidPairs.map((pair, idx) => {
                          const inv1 = invigilators.find(i => i.id === pair[0]);
                          const inv2 = invigilators.find(i => i.id === pair[1]);
                          return (
                            <li key={idx} className="flex justify-between items-center bg-slate-50 dark:bg-slate-800/50 px-3 py-2 rounded-md border border-slate-200 dark:border-slate-700">
                              <span className="text-sm text-slate-700 dark:text-slate-300">
                                {inv1?.name} <span className="text-slate-400 mx-1">và</span> {inv2?.name}
                              </span>
                              <button
                                onClick={() => handleRemoveAvoidPair(idx)}
                                className="text-red-500 hover:text-red-700 text-sm font-medium"
                              >
                                Xóa
                              </button>
                            </li>
                          );
                        })}
                      </ul>
                    )}
                  </div>
                )}

                {validationWarning && (
                  <div className="rounded-md bg-amber-50 dark:bg-amber-900/20 p-3 border border-amber-200 dark:border-amber-800">
                    <div className="flex">
                      <div className="flex-shrink-0">
                        <AlertCircle className="h-5 w-5 text-amber-500 dark:text-amber-400" aria-hidden="true" />
                      </div>
                      <div className="ml-3">
                        <h3 className="text-sm font-medium text-amber-800 dark:text-amber-300">Cảnh báo</h3>
                        <div className="mt-1 text-xs text-amber-700 dark:text-amber-400">
                          <p>{validationWarning}</p>
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                <div className="pt-4">
                  <button
                    onClick={handleGenerate}
                    disabled={isGenerating || invigilators.length === 0 || !!validationWarning}
                    className={`w-full flex justify-center py-3 px-4 border border-transparent rounded-lg shadow-sm text-sm font-medium text-white ${
                      isGenerating || invigilators.length === 0 || !!validationWarning
                        ? 'bg-indigo-400 dark:bg-indigo-800/50 cursor-not-allowed'
                        : 'bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500'
                    }`}
                  >
                    {isGenerating ? 'Đang phân công...' : 'Phân công giám thị'}
                  </button>
                </div>
              </div>
            </div>

            {invigilators.length > 0 && (
              <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-700 p-6 transition-colors duration-200">
                <h2 className="text-lg font-medium mb-4 flex items-center gap-2 text-slate-800 dark:text-white">
                  <Users className="w-5 h-5 text-slate-500 dark:text-slate-400" />
                  Danh sách giám thị ({invigilators.length})
                </h2>
                <div className="max-h-60 overflow-y-auto border border-slate-100 dark:border-slate-700 rounded-lg">
                  <table className="min-w-full divide-y divide-slate-200 dark:divide-slate-700">
                    <thead className="bg-slate-50 dark:bg-slate-800/50 sticky top-0">
                      <tr>
                        <th scope="col" className="px-4 py-2 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Mã GT</th>
                        <th scope="col" className="px-4 py-2 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Họ và tên</th>
                        <th scope="col" className="px-4 py-2 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Kinh nghiệm</th>
                      </tr>
                    </thead>
                    <tbody className="bg-white dark:bg-slate-800 divide-y divide-slate-100 dark:divide-slate-700">
                      {sortedInvigilators.map((inv) => (
                        <tr key={inv.id} className="hover:bg-slate-50 dark:hover:bg-slate-700/50">
                          <td className="px-4 py-2 whitespace-nowrap text-xs text-slate-500 dark:text-slate-400">{inv.id}</td>
                          <td className="px-4 py-2 whitespace-nowrap text-sm text-slate-700 dark:text-slate-300">{inv.name}</td>
                          <td className="px-4 py-2 whitespace-nowrap text-xs text-slate-500 dark:text-slate-400">{inv.experience || 0}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            <div className="bg-blue-50 dark:bg-blue-900/20 rounded-2xl p-5 border border-blue-100 dark:border-blue-800/50 transition-colors duration-200">
              <h3 className="text-sm font-medium text-blue-800 dark:text-blue-300 mb-2">Mục tiêu phân công:</h3>
              <ul className="text-sm text-blue-700 dark:text-blue-400 space-y-1 list-disc pl-4">
                <li>Cân bằng số ca: Mỗi người gác tối đa {idealMax} ca, tối thiểu {idealMin} ca.</li>
                <li>Hạn chế tối đa việc gác cùng 1 phòng quá 1 lần.</li>
                {!rules.fixedPairs && (
                  <li>Hạn chế tối đa việc gác cùng 1 người quá 1 lần.</li>
                )}
                {rules.minimizeConsecutive && (
                  <li><strong className="font-semibold">Ưu tiên:</strong> Giảm thiểu việc giám thị phải gác liên tục nhiều ca thi liền nhau.</li>
                )}
                {rules.fixedPairs && (
                  <li><strong className="font-semibold">Ưu tiên:</strong> Phân công theo cặp đôi cố định để tăng sự nhất quán.</li>
                )}
                {rules.balanceExperience && (
                  <li><strong className="font-semibold">Ưu tiên:</strong> Ghép người có kinh nghiệm với người mới.</li>
                )}
                {rules.avoidPairs && rules.avoidPairs.length > 0 && (
                  <li><strong className="font-semibold">Bắt buộc:</strong> Tránh xếp chung phòng các cặp giám thị đã chỉ định.</li>
                )}
              </ul>
            </div>
          </div>

          <div className="lg:col-span-8">
            <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-700 p-6 min-h-[600px] transition-colors duration-200">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between mb-6 gap-4">
                <h2 className="text-lg font-medium text-slate-800 dark:text-white">Kết quả phân công</h2>
                
                {result?.schedule && (
                  <div className="flex flex-wrap gap-2">
                    <button
                      onClick={() => exportToExcel(result, shiftNames, currentRoomNames)}
                      className="inline-flex items-center px-3 py-2 border border-slate-300 dark:border-slate-600 shadow-sm text-sm leading-4 font-medium rounded-md text-slate-700 dark:text-slate-200 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500"
                    >
                      <FileSpreadsheet className="w-4 h-4 mr-2 text-emerald-600 dark:text-emerald-400" />
                      Xuất Excel
                    </button>
                    <button
                      onClick={() => setPreviewConfig({ isOpen: true, type: 'schedule' })}
                      className="inline-flex items-center px-3 py-2 border border-slate-300 dark:border-slate-600 shadow-sm text-sm leading-4 font-medium rounded-md text-slate-700 dark:text-slate-200 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500"
                    >
                      <FileText className="w-4 h-4 mr-2 text-blue-600 dark:text-blue-400" />
                      Xem & Xuất Word
                    </button>
                    <button
                      onClick={() => setPreviewConfig({ isOpen: true, type: 'm3' })}
                      className="inline-flex items-center px-3 py-2 border border-slate-300 dark:border-slate-600 shadow-sm text-sm leading-4 font-medium rounded-md text-slate-700 dark:text-slate-200 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500"
                    >
                      <FileText className="w-4 h-4 mr-2 text-indigo-600 dark:text-indigo-400" />
                      Xem & Xuất Mẫu M3
                    </button>
                    <button
                      onClick={() => setPreviewConfig({ isOpen: true, type: 'm14' })}
                      className="inline-flex items-center px-3 py-2 border border-slate-300 dark:border-slate-600 shadow-sm text-sm leading-4 font-medium rounded-md text-slate-700 dark:text-slate-200 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500"
                    >
                      <FileText className="w-4 h-4 mr-2 text-violet-600 dark:text-violet-400" />
                      Xem & Xuất Mẫu M14
                    </button>
                  </div>
                )}
              </div>

              {result?.schedule && (
                <div className="mb-6 p-4 bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700 rounded-xl grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                  <div>
                    <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-1 flex items-center gap-1">
                      <Filter className="w-3 h-3" /> Lọc theo Ca thi
                    </label>
                    <select
                      value={filterShift}
                      onChange={(e) => setFilterShift(e.target.value)}
                      className="block w-full pl-3 pr-10 py-2 text-sm border border-slate-300 dark:border-slate-600 rounded-md focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100"
                    >
                      <option value="all">Tất cả ca thi</option>
                      {shiftNames.slice(0, numShifts).map((name, idx) => (
                        <option key={idx} value={idx.toString()}>{name || `Ca thi ${idx + 1}`}</option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-1 flex items-center gap-1">
                      <Filter className="w-3 h-3" /> Lọc theo Giám thị
                    </label>
                    <select
                      value={filterInvigilator}
                      onChange={(e) => setFilterInvigilator(e.target.value)}
                      className="block w-full pl-3 pr-10 py-2 text-sm border border-slate-300 dark:border-slate-600 rounded-md focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100"
                    >
                      <option value="all">Tất cả giám thị</option>
                      {sortedInvigilators.map(inv => (
                        <option key={inv.id} value={inv.id}>{inv.name} ({inv.id})</option>
                      ))}
                    </select>
                  </div>
                  
                  <div>
                    <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-1 flex items-center gap-1">
                      <Filter className="w-3 h-3" /> Lọc theo Phòng
                    </label>
                    <select
                      value={filterRoom}
                      onChange={(e) => setFilterRoom(e.target.value)}
                      className="block w-full pl-3 pr-10 py-2 text-sm border border-slate-300 dark:border-slate-600 rounded-md focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100"
                    >
                      <option value="all">Tất cả phòng thi</option>
                      {allRooms.map(room => (
                        <option key={room} value={room.toString()}>{currentRoomNames[room - 1] || `Phòng ${room}`}</option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-1 flex items-center gap-1">
                      <ArrowUpDown className="w-3 h-3" /> Sắp xếp theo
                    </label>
                    <select
                      value={sortBy}
                      onChange={(e) => setSortBy(e.target.value as 'room' | 'invigilator')}
                      className="block w-full pl-3 pr-10 py-2 text-sm border border-slate-300 dark:border-slate-600 rounded-md focus:ring-indigo-500 focus:border-indigo-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100"
                    >
                      <option value="room">Phòng thi</option>
                      <option value="invigilator">Tên giám thị</option>
                    </select>
                  </div>
                </div>
              )}

              {result?.error ? (
                <div className="rounded-md bg-red-50 dark:bg-red-900/20 p-4 border border-red-200 dark:border-red-800">
                  <div className="flex">
                    <div className="flex-shrink-0">
                      <AlertCircle className="h-5 w-5 text-red-400 dark:text-red-500" aria-hidden="true" />
                    </div>
                    <div className="ml-3">
                      <h3 className="text-sm font-medium text-red-800 dark:text-red-300">Lỗi phân công</h3>
                      <div className="mt-2 text-sm text-red-700 dark:text-red-400">
                        <p>{result.error}</p>
                      </div>
                    </div>
                  </div>
                </div>
              ) : result?.schedule ? (
                <div className="space-y-8">
                  {result.schedule.map((shift, shiftIdx) => {
                    if (filterShift !== 'all' && shiftIdx.toString() !== filterShift) return null;

                    const flattenedShift = shift.flatMap(a => {
                      const assignments = [{ invigilator: a.invigilator1, room: a.room }];
                      if (a.invigilator2) {
                        assignments.push({ invigilator: a.invigilator2, room: a.room });
                      }
                      return assignments;
                    });

                    const filteredShift = flattenedShift.filter(a => {
                      if (filterRoom !== 'all' && a.room.toString() !== filterRoom) return false;
                      if (filterInvigilator !== 'all' && a.invigilator.id !== filterInvigilator) return false;
                      return true;
                    });

                    const unassigned = getUnassignedInvigilators(shift);

                    if (filteredShift.length === 0 && filterRoom !== 'all') return null;

                    const sortedShift = [...filteredShift].sort((a, b) => {
                      if (sortBy === 'room') {
                        if (a.room === b.room) {
                          return compareVietnameseName(a.invigilator.name, b.invigilator.name);
                        }
                        return a.room - b.room;
                      } else {
                        return compareVietnameseName(a.invigilator.name, b.invigilator.name);
                      }
                    });
                    
                    const currentPage = pageByShift[shiftIdx] || 1;
                    const totalPages = Math.ceil(sortedShift.length / ITEMS_PER_PAGE);
                    const paginatedShift = sortedShift.slice((currentPage - 1) * ITEMS_PER_PAGE, currentPage * ITEMS_PER_PAGE);

                    return (
                      <div key={shiftIdx} className="border border-slate-200 dark:border-slate-700 rounded-xl overflow-hidden">
                        <div className="bg-slate-50 dark:bg-slate-800/50 px-4 py-3 border-b border-slate-200 dark:border-slate-700 flex justify-between items-center">
                          <h3 className="text-md font-medium text-slate-800 dark:text-slate-200">{shiftNames[shiftIdx] || `Ca thi ${shiftIdx + 1}`}</h3>
                          <div className="flex items-center gap-3">
                            <button
                              onClick={() => handleRandomizeRooms(shiftIdx)}
                              className="inline-flex items-center px-2.5 py-1.5 border border-indigo-200 dark:border-indigo-800 shadow-sm text-xs font-medium rounded text-indigo-700 dark:text-indigo-300 bg-indigo-50 dark:bg-indigo-900/30 hover:bg-indigo-100 dark:hover:bg-indigo-900/50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 transition-colors"
                              title="Bốc thăm lại phòng thi ngẫu nhiên cho các giám thị trong ca này"
                            >
                              🎲 Bốc thăm phòng
                            </button>
                            <span className="text-xs font-medium text-slate-500 dark:text-slate-400 bg-white dark:bg-slate-800 px-2 py-1 rounded-md border border-slate-200 dark:border-slate-700">
                              {shift.length} phòng
                            </span>
                          </div>
                        </div>
                        
                        {sortedShift.length > 0 && (
                          <div className="flex flex-col">
                            <div className="overflow-x-auto overflow-y-auto max-h-[450px] scrollbar-thin scrollbar-thumb-slate-300 dark:scrollbar-thumb-slate-600">
                              <table className="min-w-full divide-y divide-slate-200 dark:divide-slate-700 relative">
                                <thead className="sticky top-0 z-20 shadow-sm">
                                  <tr>
                                    <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider bg-slate-100 dark:bg-slate-900">
                                      Tên giám thị
                                    </th>
                                    <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider w-32 bg-slate-100 dark:bg-slate-900">
                                      Phòng thi
                                    </th>
                                    <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider w-48 bg-slate-100 dark:bg-slate-900">
                                      Ký tên
                                    </th>
                                  </tr>
                                </thead>
                                <tbody className="bg-white dark:bg-slate-800 divide-y divide-slate-200 dark:divide-slate-700">
                                  {paginatedShift.map((assignment, aIdx) => {
                                    const highlight = filterInvigilator !== 'all' && assignment.invigilator.id === filterInvigilator;
                                    
                                    return (
                                      <tr key={aIdx} className={`hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors ${aIdx % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-slate-50/50 dark:bg-slate-800/30'}`}>
                                        <td className={`px-6 py-4 whitespace-nowrap text-sm ${highlight ? 'font-bold text-indigo-600 dark:text-indigo-400 bg-indigo-50/50 dark:bg-indigo-900/20' : 'font-medium text-slate-900 dark:text-slate-100'}`}>
                                          {assignment.invigilator.name} <span className="text-xs text-slate-500 dark:text-slate-400 font-normal ml-1">({assignment.invigilator.id})</span>
                                        </td>
                                        <td className="px-6 py-4 whitespace-nowrap text-sm text-slate-600 dark:text-slate-300">
                                          <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-blue-100 text-blue-800 dark:bg-blue-900/30 dark:text-blue-300">
                                            {currentRoomNames[assignment.room - 1] || assignment.room}
                                          </span>
                                        </td>
                                        <td className="px-6 py-4 whitespace-nowrap text-sm text-slate-500 dark:text-slate-400">
                                          <div className="border-b border-dashed border-slate-300 dark:border-slate-600 w-full h-6"></div>
                                        </td>
                                      </tr>
                                    );
                                  })}
                                </tbody>
                              </table>
                            </div>
                            
                            {totalPages > 1 && (
                              <div className="bg-slate-50 dark:bg-slate-800/80 backdrop-blur-sm px-4 py-3 border-t border-slate-200 dark:border-slate-700 flex items-center justify-between sm:px-6 sticky bottom-0 z-10">
                                <div className="hidden sm:flex-1 sm:flex sm:items-center sm:justify-between">
                                  <div>
                                    <p className="text-sm text-slate-700 dark:text-slate-300">
                                      Hiển thị <span className="font-medium">{(currentPage - 1) * ITEMS_PER_PAGE + 1}</span> đến <span className="font-medium">{Math.min(currentPage * ITEMS_PER_PAGE, sortedShift.length)}</span> trong số <span className="font-medium">{sortedShift.length}</span> kết quả
                                    </p>
                                  </div>
                                  <div>
                                    <nav className="relative z-0 inline-flex rounded-md shadow-sm -space-x-px" aria-label="Pagination">
                                      <button
                                        onClick={() => setPageByShift(prev => ({ ...prev, [shiftIdx]: Math.max(1, currentPage - 1) }))}
                                        disabled={currentPage === 1}
                                        className="relative inline-flex items-center px-2 py-2 rounded-l-md border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-sm font-medium text-slate-500 dark:text-slate-400 hover:bg-slate-50 dark:hover:bg-slate-700 disabled:opacity-50 disabled:cursor-not-allowed"
                                      >
                                        <span className="sr-only">Trước</span>
                                        <ChevronLeft className="h-4 w-4" aria-hidden="true" />
                                      </button>
                                      
                                      {/* Page numbers */}
                                      {Array.from({ length: totalPages }, (_, i) => i + 1).map(page => {
                                        // Show first, last, current, and adjacent pages
                                        if (
                                          page === 1 ||
                                          page === totalPages ||
                                          (page >= currentPage - 1 && page <= currentPage + 1)
                                        ) {
                                          return (
                                            <button
                                              key={page}
                                              onClick={() => setPageByShift(prev => ({ ...prev, [shiftIdx]: page }))}
                                              className={`relative inline-flex items-center px-4 py-2 border text-sm font-medium ${
                                                currentPage === page
                                                  ? 'z-10 bg-indigo-50 dark:bg-indigo-900/20 border-indigo-500 text-indigo-600 dark:text-indigo-400'
                                                  : 'bg-white dark:bg-slate-800 border-slate-300 dark:border-slate-600 text-slate-500 dark:text-slate-400 hover:bg-slate-50 dark:hover:bg-slate-700'
                                              }`}
                                            >
                                              {page}
                                            </button>
                                          );
                                        }
                                        
                                        // Show ellipsis for gaps
                                        if (
                                          (page === 2 && currentPage > 3) ||
                                          (page === totalPages - 1 && currentPage < totalPages - 2)
                                        ) {
                                          return (
                                            <span
                                              key={page}
                                              className="relative inline-flex items-center px-4 py-2 border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-sm font-medium text-slate-700 dark:text-slate-300"
                                            >
                                              ...
                                            </span>
                                          );
                                        }
                                        
                                        return null;
                                      })}
                                      
                                      <button
                                        onClick={() => setPageByShift(prev => ({ ...prev, [shiftIdx]: Math.min(totalPages, currentPage + 1) }))}
                                        disabled={currentPage === totalPages}
                                        className="relative inline-flex items-center px-2 py-2 rounded-r-md border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-sm font-medium text-slate-500 dark:text-slate-400 hover:bg-slate-50 dark:hover:bg-slate-700 disabled:opacity-50 disabled:cursor-not-allowed"
                                      >
                                        <span className="sr-only">Sau</span>
                                        <ChevronRight className="h-4 w-4" aria-hidden="true" />
                                      </button>
                                    </nav>
                                  </div>
                                </div>
                                <div className="flex items-center justify-between sm:hidden w-full">
                                  <button
                                    onClick={() => setPageByShift(prev => ({ ...prev, [shiftIdx]: Math.max(1, currentPage - 1) }))}
                                    disabled={currentPage === 1}
                                    className="relative inline-flex items-center px-4 py-2 border border-slate-300 dark:border-slate-600 text-sm font-medium rounded-md text-slate-700 dark:text-slate-200 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 disabled:opacity-50 disabled:cursor-not-allowed"
                                  >
                                    Trước
                                  </button>
                                  <span className="text-sm text-slate-700 dark:text-slate-300">
                                    Trang {currentPage} / {totalPages}
                                  </span>
                                  <button
                                    onClick={() => setPageByShift(prev => ({ ...prev, [shiftIdx]: Math.min(totalPages, currentPage + 1) }))}
                                    disabled={currentPage === totalPages}
                                    className="relative inline-flex items-center px-4 py-2 border border-slate-300 dark:border-slate-600 text-sm font-medium rounded-md text-slate-700 dark:text-slate-200 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 disabled:opacity-50 disabled:cursor-not-allowed"
                                  >
                                    Sau
                                  </button>
                                </div>
                              </div>
                            )}
                          </div>
                        )}
                        
                        {/* Unassigned Invigilators Section */}
                        {filterRoom === 'all' && unassigned.length > 0 && (
                          <div className="bg-amber-50/50 dark:bg-amber-900/10 px-6 py-4 border-t border-slate-200 dark:border-slate-700">
                            <h4 className="text-sm font-medium text-amber-800 dark:text-amber-300 mb-2 flex items-center gap-1">
                              <Info className="w-4 h-4" />
                              Giám thị nghỉ ca này ({unassigned.length}):
                            </h4>
                            <div className="flex flex-wrap gap-2">
                              {unassigned.map(inv => {
                                const isHighlighted = filterInvigilator !== 'all' && inv.id === filterInvigilator;
                                return (
                                  <span 
                                    key={inv.id} 
                                    className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium ${
                                      isHighlighted 
                                        ? 'bg-indigo-100 dark:bg-indigo-900/30 text-indigo-800 dark:text-indigo-300 border border-indigo-200 dark:border-indigo-800' 
                                        : 'bg-amber-100 dark:bg-amber-900/30 text-amber-800 dark:text-amber-300 border border-amber-200 dark:border-amber-800'
                                    }`}
                                  >
                                    {inv.name}
                                  </span>
                                );
                              })}
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })}
                  
                  {/* Show empty state if filters hide everything */}
                  {result.schedule.every((shift, shiftIdx) => {
                    if (filterShift !== 'all' && shiftIdx.toString() !== filterShift) return true;
                    return shift.filter(a => {
                      if (filterRoom !== 'all' && a.room.toString() !== filterRoom) return false;
                      if (filterInvigilator !== 'all' && a.invigilator1.id !== filterInvigilator && (!a.invigilator2 || a.invigilator2.id !== filterInvigilator)) return false;
                      return true;
                    }).length === 0;
                  }) && (
                    <div className="text-center py-12 text-slate-500 dark:text-slate-400">
                      Không tìm thấy kết quả phù hợp với bộ lọc hiện tại.
                    </div>
                  )}
                </div>
              ) : (
                <div className="flex flex-col items-center justify-center h-full text-slate-400 dark:text-slate-500 space-y-4 py-20">
                  <Calendar className="w-16 h-16 text-slate-200 dark:text-slate-700" />
                  <p>Tải lên danh sách hoặc dán văn bản và nhấn "Phân công giám thị" để xem kết quả</p>
                </div>
              )}
            </div>
          </div>
        </div>
      </main>

      <ExportPreviewModal
        isOpen={previewConfig.isOpen}
        onClose={() => setPreviewConfig({ isOpen: false, type: null })}
        onDownload={() => {
          if (previewConfig.type === 'schedule') exportToDocx(result!, shiftNames, currentRoomNames);
          else if (previewConfig.type === 'm3') exportM3ToDocx(result!, shiftNames);
          else if (previewConfig.type === 'm14') exportM14ToDocx(result!, shiftNames, currentRoomNames);
        }}
        type={previewConfig.type}
        result={result}
        shiftNames={shiftNames}
        roomNames={currentRoomNames}
      />
    </div>
  );
}