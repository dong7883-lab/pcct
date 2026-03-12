import React from 'react';
import { X, Download } from 'lucide-react';
import { AssignmentResult, compareVietnameseName } from '../utils/scheduler';

interface ExportPreviewModalProps {
  isOpen: boolean;
  onClose: () => void;
  onDownload: () => void;
  type: 'schedule' | 'm3' | 'm14' | null;
  result: AssignmentResult | null;
  shiftNames: string[];
  roomNames: string[];
}

export function ExportPreviewModal({ isOpen, onClose, onDownload, type, result, shiftNames, roomNames }: ExportPreviewModalProps) {
  if (!isOpen || !result || !type) return null;

  const allInvigilators = result.stats.map(s => s.invigilator);

  const renderSchedule = () => {
    const hasInvigilator2 = result.schedule[0]?.[0]?.invigilator2 !== undefined;
    
    return (
      <div className="space-y-8 font-serif text-black">
        <h2 className="text-2xl font-bold text-center mb-6">BẢNG PHÂN CÔNG GIÁM THỊ</h2>
        <table className="w-full border-collapse border border-black text-sm">
          <thead>
            <tr className="bg-gray-100">
              <th className="border border-black p-2">Ca thi</th>
              <th className="border border-black p-2">Phòng thi</th>
              <th className="border border-black p-2">Giám thị 1</th>
              {hasInvigilator2 && <th className="border border-black p-2">Giám thị 2</th>}
            </tr>
          </thead>
          <tbody>
            {result.schedule.map((shift, sIdx) => (
              shift.map((a, aIdx) => (
                <tr key={`${sIdx}-${aIdx}`}>
                  <td className="border border-black p-2 text-center">{shiftNames[sIdx] || `Ca ${sIdx + 1}`}</td>
                  <td className="border border-black p-2 text-center">{roomNames[a.room - 1] || `Phòng ${a.room}`}</td>
                  <td className="border border-black p-2">{a.invigilator1.name}</td>
                  {hasInvigilator2 && <td className="border border-black p-2">{a.invigilator2?.name || ''}</td>}
                </tr>
              ))
            ))}
          </tbody>
        </table>
      </div>
    );
  };

  const renderM3M14 = () => {
    return (
      <div className="space-y-16 font-serif text-black">
        {result.schedule.map((shift, index) => {
          const assignedIds = new Set<string>();
          const assignedRooms = new Map<string, string>();
          
          shift.forEach(a => {
            assignedIds.add(a.invigilator1.id);
            const roomName = roomNames[a.room - 1] || `Phòng ${a.room}`;
            assignedRooms.set(a.invigilator1.id, roomName);
            if (a.invigilator2) {
              assignedIds.add(a.invigilator2.id);
              assignedRooms.set(a.invigilator2.id, roomName);
            }
          });

          const shiftInvigilators = [...allInvigilators].sort((a, b) => compareVietnameseName(a.name, b.name));

          return (
            <div key={index} className="relative">
              {index > 0 && <div className="absolute -top-8 left-0 right-0 border-t-2 border-dashed border-gray-300"></div>}
              <div className="text-center mb-6">
                <h3 className="text-xl font-bold">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</h3>
                <h4 className="text-lg font-bold underline mb-4">Độc lập - Tự do - Hạnh phúc</h4>
                <h2 className="text-2xl font-bold mt-8 mb-2">
                  {type === 'm3' ? 'DANH SÁCH BỐC THĂM PHÂN CÔNG COI THI (MẪU M3)' : 'DANH SÁCH KÝ TÊN GIÁM THỊ (MẪU M14)'}
                </h2>
                <h3 className="text-xl font-bold">{shiftNames[index] || `Ca thi ${index + 1}`}</h3>
              </div>
              <table className="w-full border-collapse border border-black text-sm">
                <thead>
                  <tr className="bg-gray-100">
                    <th className="border border-black p-2 w-12">STT</th>
                    <th className="border border-black p-2 w-24">Mã GT</th>
                    <th className="border border-black p-2">Họ và tên</th>
                    <th className="border border-black p-2 w-32 whitespace-pre-line">
                      {type === 'm3' ? 'Phòng thi\n(Bốc thăm)' : 'Phòng thi'}
                    </th>
                    <th className="border border-black p-2 w-32">Ký tên</th>
                    <th className="border border-black p-2 w-24">Ghi chú</th>
                  </tr>
                </thead>
                <tbody>
                  {shiftInvigilators.map((inv, i) => {
                    const isResting = !assignedIds.has(inv.id);
                    const room = assignedRooms.get(inv.id);
                    return (
                      <tr key={inv.id}>
                        <td className="border border-black p-2 text-center">{i + 1}</td>
                        <td className="border border-black p-2 text-center">{inv.id}</td>
                        <td className="border border-black p-2">{inv.name}</td>
                        <td className="border border-black p-2 text-center">
                          {isResting ? 'Nghỉ' : (type === 'm14' ? room : '')}
                        </td>
                        <td className="border border-black p-2"></td>
                        <td className="border border-black p-2 text-center">{isResting ? 'Nghỉ' : ''}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          );
        })}
      </div>
    );
  };

  return (
    <div className="fixed inset-0 z-50 overflow-y-auto" aria-labelledby="modal-title" role="dialog" aria-modal="true">
      <div className="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div className="fixed inset-0 bg-slate-900/75 backdrop-blur-sm transition-opacity" aria-hidden="true" onClick={onClose}></div>
        <span className="hidden sm:inline-block sm:align-middle sm:h-screen" aria-hidden="true">&#8203;</span>
        
        <div className="inline-block align-bottom bg-white dark:bg-slate-800 rounded-xl text-left overflow-hidden shadow-2xl transform transition-all sm:my-8 sm:align-middle sm:max-w-5xl sm:w-full border border-slate-200 dark:border-slate-700">
          <div className="bg-white dark:bg-slate-800 px-4 pt-5 pb-4 sm:p-6 sm:pb-4">
            <div className="sm:flex sm:items-start">
              <div className="mt-3 text-center sm:mt-0 sm:text-left w-full">
                <h3 className="text-lg leading-6 font-medium text-slate-900 dark:text-white flex justify-between items-center mb-4" id="modal-title">
                  <span className="flex items-center gap-2">
                    <FileText className="w-5 h-5 text-indigo-500" />
                    Xem trước bản in ({type === 'schedule' ? 'Lịch phân công' : type === 'm3' ? 'Mẫu M3' : 'Mẫu M14'})
                  </span>
                  <button onClick={onClose} className="text-slate-400 hover:text-slate-500 dark:hover:text-slate-300 transition-colors p-1 rounded-md hover:bg-slate-100 dark:hover:bg-slate-700">
                    <X className="w-5 h-5" />
                  </button>
                </h3>
                
                <div className="mt-4 max-h-[65vh] overflow-y-auto border border-slate-200 dark:border-slate-700 rounded-lg p-8 bg-white shadow-inner">
                  {type === 'schedule' && renderSchedule()}
                  {(type === 'm3' || type === 'm14') && renderM3M14()}
                </div>
              </div>
            </div>
          </div>
          <div className="bg-slate-50 dark:bg-slate-800/80 px-4 py-3 sm:px-6 flex flex-col-reverse sm:flex-row sm:justify-end gap-2 border-t border-slate-200 dark:border-slate-700">
            <button
              type="button"
              onClick={onClose}
              className="w-full inline-flex justify-center rounded-lg border border-slate-300 dark:border-slate-600 shadow-sm px-4 py-2 bg-white dark:bg-slate-800 text-base font-medium text-slate-700 dark:text-slate-300 hover:bg-slate-50 dark:hover:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 sm:w-auto sm:text-sm transition-colors"
            >
              Đóng
            </button>
            <button
              type="button"
              onClick={() => { onDownload(); onClose(); }}
              className="w-full inline-flex justify-center items-center gap-2 rounded-lg border border-transparent shadow-sm px-4 py-2 bg-indigo-600 text-base font-medium text-white hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 sm:w-auto sm:text-sm transition-colors"
            >
              <Download className="w-4 h-4" />
              Tải xuống DOCX
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
