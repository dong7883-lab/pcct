import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType, AlignmentType, BorderStyle, VerticalAlign } from 'docx';
import { saveAs } from 'file-saver';
import { AssignmentResult, compareVietnameseName } from './scheduler';

export function exportToExcel(result: AssignmentResult, shiftNames: string[] = [], roomNames: string[] = []) {
  if (!result.schedule || !result.stats) return;

  const wb = XLSX.utils.book_new();

  // Determine if we have 1 or 2 invigilators per room
  const hasInvigilator2 = result.schedule[0]?.[0]?.invigilator2 !== undefined;

  // Sheet 1: Schedule
  const scheduleData = [];
  const header = ['Ca thi', 'Phòng thi', 'Giám thị 1'];
  if (hasInvigilator2) header.push('Giám thị 2');
  scheduleData.push(header);
  
  result.schedule.forEach((shift, index) => {
    shift.forEach(assignment => {
      const row = [
        shiftNames[index] || `Ca ${assignment.shift}`,
        roomNames[assignment.room - 1] || `Phòng ${assignment.room}`,
        assignment.invigilator1.name
      ];
      if (hasInvigilator2) {
        row.push(assignment.invigilator2 ? assignment.invigilator2.name : '');
      }
      scheduleData.push(row);
    });
  });

  const wsSchedule = XLSX.utils.aoa_to_sheet(scheduleData);
  XLSX.utils.book_append_sheet(wb, wsSchedule, 'Lịch phân công');

  // Sheet 2: Resting Invigilators
  const restingData = [];
  restingData.push(['Ca thi', 'Mã GT', 'Họ và tên']);
  
  const allInvigilators = result.stats.map(s => s.invigilator);
  
  result.schedule.forEach((shift, index) => {
    const shiftName = shiftNames[index] || `Ca ${index + 1}`;
    const assignedIds = new Set<string>();
    shift.forEach(a => {
      assignedIds.add(a.invigilator1.id);
      if (a.invigilator2) assignedIds.add(a.invigilator2.id);
    });
    
    const resting = allInvigilators.filter(inv => !assignedIds.has(inv.id));
    resting.sort((a, b) => compareVietnameseName(a.name, b.name));
    
    resting.forEach(inv => {
      restingData.push([shiftName, inv.id, inv.name]);
    });
  });

  if (restingData.length > 1) {
    const wsResting = XLSX.utils.aoa_to_sheet(restingData);
    XLSX.utils.book_append_sheet(wb, wsResting, 'Giám thị nghỉ');
  }

  // Sheet 3: Stats
  const statsData = [];
  statsData.push(['Mã GT', 'Họ và tên', 'Số ca coi thi']);
  result.stats.forEach(stat => {
    statsData.push([stat.invigilator.id, stat.invigilator.name, stat.count]);
  });

  const wsStats = XLSX.utils.aoa_to_sheet(statsData);
  XLSX.utils.book_append_sheet(wb, wsStats, 'Thống kê');

  XLSX.writeFile(wb, 'PhanCongGiamThi.xlsx');
}

export async function exportToDocx(result: AssignmentResult, shiftNames: string[] = [], roomNames: string[] = []) {
  if (!result.schedule) return;

  const hasInvigilator2 = result.schedule[0]?.[0]?.invigilator2 !== undefined;

  const children: any[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: "BẢNG PHÂN CÔNG GIÁM THỊ",
          bold: true,
          size: 32,
        }),
      ],
      spacing: { after: 400 }
    })
  ];

  result.schedule.forEach((shift, index) => {
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: shiftNames[index] || `Ca thi ${index + 1}`,
            bold: true,
            size: 28,
          }),
        ],
        spacing: { before: 400, after: 200 }
      })
    );

    const headerCells = [
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: "Phòng thi", bold: true })], alignment: AlignmentType.CENTER })],
        width: { size: 20, type: WidthType.PERCENTAGE },
      }),
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: hasInvigilator2 ? "Giám thị 1" : "Giám thị", bold: true })], alignment: AlignmentType.CENTER })],
        width: { size: hasInvigilator2 ? 40 : 80, type: WidthType.PERCENTAGE },
      })
    ];

    if (hasInvigilator2) {
      headerCells.push(
        new TableCell({
          children: [new Paragraph({ children: [new TextRun({ text: "Giám thị 2", bold: true })], alignment: AlignmentType.CENTER })],
          width: { size: 40, type: WidthType.PERCENTAGE },
        })
      );
    }

    const tableRows = [
      new TableRow({
        children: headerCells,
      })
    ];

    shift.forEach(assignment => {
      const rowCells = [
        new TableCell({
          children: [new Paragraph({ text: roomNames[assignment.room - 1] || `Phòng ${assignment.room}`, alignment: AlignmentType.CENTER })],
        }),
        new TableCell({
          children: [new Paragraph({ text: assignment.invigilator1.name })],
        })
      ];

      if (hasInvigilator2) {
        rowCells.push(
          new TableCell({
            children: [new Paragraph({ text: assignment.invigilator2 ? assignment.invigilator2.name : '' })],
          })
        );
      }

      tableRows.push(
        new TableRow({
          children: rowCells,
        })
      );
    });

    children.push(
      new Table({
        rows: tableRows,
        width: { size: 100, type: WidthType.PERCENTAGE },
      })
    );

    const allInvigilators = result.stats?.map(s => s.invigilator) || [];
    const assignedIds = new Set<string>();
    shift.forEach(a => {
      assignedIds.add(a.invigilator1.id);
      if (a.invigilator2) assignedIds.add(a.invigilator2.id);
    });
    
    const resting = allInvigilators.filter(inv => !assignedIds.has(inv.id));
    resting.sort((a, b) => compareVietnameseName(a.name, b.name));

    if (resting.length > 0) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `Giám thị nghỉ ca này (${resting.length}): `,
              bold: true,
              italics: true,
            }),
            new TextRun({
              text: resting.map(inv => inv.name).join(', '),
              italics: true,
            })
          ],
          spacing: { before: 200, after: 400 }
        })
      );
    } else {
      children.push(
        new Paragraph({
          text: "",
          spacing: { after: 400 }
        })
      );
    }
  });

  const doc = new Document({
    sections: [{
      properties: {},
      children: children,
    }],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, "PhanCongGiamThi.docx");
}

export async function exportM3ToDocx(result: AssignmentResult, shiftNames: string[] = []) {
  if (!result.schedule || !result.stats) return;

  const allInvigilators = result.stats.map(s => s.invigilator);

  const children: any[] = [];

  result.schedule.forEach((shift, index) => {
    // Collect all invigilators assigned to this shift to know who is resting
    const assignedIds = new Set<string>();
    shift.forEach(a => {
      assignedIds.add(a.invigilator1.id);
      if (a.invigilator2) {
        assignedIds.add(a.invigilator2.id);
      }
    });

    // Use all invigilators for the list
    const shiftInvigilators = [...allInvigilators];

    // Sort alphabetically by Vietnamese name
    shiftInvigilators.sort((a, b) => compareVietnameseName(a.name, b.name));

    // Add Header for the shift
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
            bold: true,
            size: 24,
          }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: "Độc lập - Tự do - Hạnh phúc",
            bold: true,
            size: 24,
          }),
        ],
        spacing: { after: 400 }
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: "DANH SÁCH BỐC THĂM PHÂN CÔNG COI THI (MẪU M3)",
            bold: true,
            size: 28,
          }),
        ],
        spacing: { after: 200 }
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: shiftNames[index] || `Ca thi ${index + 1}`,
            bold: true,
            size: 24,
          }),
        ],
        spacing: { after: 400 }
      })
    );

    // Create Table Header
    const headerCells = [
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: "STT", bold: true })], alignment: AlignmentType.CENTER })],
        width: { size: 8, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
      }),
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: "Mã GT", bold: true })], alignment: AlignmentType.CENTER })],
        width: { size: 15, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
      }),
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: "Họ và tên", bold: true })], alignment: AlignmentType.CENTER })],
        width: { size: 35, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
      }),
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: "Phòng thi\n(Bốc thăm)", bold: true })], alignment: AlignmentType.CENTER })],
        width: { size: 17, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
      }),
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: "Ký tên", bold: true })], alignment: AlignmentType.CENTER })],
        width: { size: 15, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
      }),
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: "Ghi chú", bold: true })], alignment: AlignmentType.CENTER })],
        width: { size: 10, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
      })
    ];

    const tableRows = [
      new TableRow({
        children: headerCells,
        tableHeader: true,
      })
    ];

    // Add rows for each invigilator
    shiftInvigilators.forEach((inv, i) => {
      const isResting = !assignedIds.has(inv.id);
      
      tableRows.push(
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ text: (i + 1).toString(), alignment: AlignmentType.CENTER })],
              verticalAlign: VerticalAlign.CENTER,
            }),
            new TableCell({
              children: [new Paragraph({ text: inv.id, alignment: AlignmentType.CENTER })],
              verticalAlign: VerticalAlign.CENTER,
            }),
            new TableCell({
              children: [new Paragraph({ text: inv.name })],
              verticalAlign: VerticalAlign.CENTER,
            }),
            new TableCell({
              children: [new Paragraph({ text: isResting ? "Nghỉ" : "", alignment: AlignmentType.CENTER })], // Empty for drawing or "Nghỉ"
              verticalAlign: VerticalAlign.CENTER,
            }),
            new TableCell({
              children: [new Paragraph({ text: "" })], // Empty for signature
              verticalAlign: VerticalAlign.CENTER,
            }),
            new TableCell({
              children: [new Paragraph({ text: isResting ? "Nghỉ" : "", alignment: AlignmentType.CENTER })], // Empty for notes or "Nghỉ"
              verticalAlign: VerticalAlign.CENTER,
            })
          ],
        })
      );
    });

    children.push(
      new Table({
        rows: tableRows,
        width: { size: 100, type: WidthType.PERCENTAGE },
      })
    );

    // Add page break after each shift except the last one
    if (index < result.schedule.length - 1) {
      children.push(
        new Paragraph({
          text: "",
          pageBreakBefore: true,
        })
      );
    }
  });

  const doc = new Document({
    sections: [{
      properties: {},
      children: children,
    }],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, "DanhSachBocTham_M3.docx");
}

export async function exportM14ToDocx(result: AssignmentResult, shiftNames: string[] = [], roomNames: string[] = []) {
  if (!result.schedule || !result.stats) return;

  const allInvigilators = result.stats.map(s => s.invigilator);

  const children: any[] = [];

  result.schedule.forEach((shift, index) => {
    // Map invigilator ID to their assigned room
    const assignedRooms = new Map<string, string>();
    shift.forEach(a => {
      const roomName = roomNames[a.room - 1] || `Phòng ${a.room}`;
      assignedRooms.set(a.invigilator1.id, roomName);
      if (a.invigilator2) {
        assignedRooms.set(a.invigilator2.id, roomName);
      }
    });

    // Use all invigilators for the list
    const shiftInvigilators = [...allInvigilators];

    // Sort alphabetically by Vietnamese name
    shiftInvigilators.sort((a, b) => compareVietnameseName(a.name, b.name));

    // Add Header for the shift
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
            bold: true,
            size: 24,
          }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: "Độc lập - Tự do - Hạnh phúc",
            bold: true,
            size: 24,
          }),
        ],
        spacing: { after: 400 }
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: "DANH SÁCH KÝ TÊN GIÁM THỊ (MẪU M14)",
            bold: true,
            size: 28,
          }),
        ],
        spacing: { after: 200 }
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: shiftNames[index] || `Ca thi ${index + 1}`,
            bold: true,
            size: 24,
          }),
        ],
        spacing: { after: 400 }
      })
    );

    // Create Table Header
    const headerCells = [
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: "STT", bold: true })], alignment: AlignmentType.CENTER })],
        width: { size: 8, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
      }),
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: "Mã GT", bold: true })], alignment: AlignmentType.CENTER })],
        width: { size: 15, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
      }),
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: "Họ và tên", bold: true })], alignment: AlignmentType.CENTER })],
        width: { size: 35, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
      }),
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: "Phòng thi", bold: true })], alignment: AlignmentType.CENTER })],
        width: { size: 17, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
      }),
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: "Ký tên", bold: true })], alignment: AlignmentType.CENTER })],
        width: { size: 15, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
      }),
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: "Ghi chú", bold: true })], alignment: AlignmentType.CENTER })],
        width: { size: 10, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
      })
    ];

    const tableRows = [
      new TableRow({
        children: headerCells,
        tableHeader: true,
      })
    ];

    // Add rows for each invigilator
    shiftInvigilators.forEach((inv, i) => {
      const room = assignedRooms.get(inv.id);
      const isResting = !room;
      
      tableRows.push(
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ text: (i + 1).toString(), alignment: AlignmentType.CENTER })],
              verticalAlign: VerticalAlign.CENTER,
            }),
            new TableCell({
              children: [new Paragraph({ text: inv.id, alignment: AlignmentType.CENTER })],
              verticalAlign: VerticalAlign.CENTER,
            }),
            new TableCell({
              children: [new Paragraph({ text: inv.name })],
              verticalAlign: VerticalAlign.CENTER,
            }),
            new TableCell({
              children: [new Paragraph({ text: isResting ? "Nghỉ" : room, alignment: AlignmentType.CENTER })],
              verticalAlign: VerticalAlign.CENTER,
            }),
            new TableCell({
              children: [new Paragraph({ text: "" })], // Empty for signature
              verticalAlign: VerticalAlign.CENTER,
            }),
            new TableCell({
              children: [new Paragraph({ text: isResting ? "Nghỉ" : "", alignment: AlignmentType.CENTER })],
              verticalAlign: VerticalAlign.CENTER,
            })
          ],
        })
      );
    });

    children.push(
      new Table({
        rows: tableRows,
        width: { size: 100, type: WidthType.PERCENTAGE },
      })
    );

    // Add page break after each shift except the last one
    if (index < result.schedule.length - 1) {
      children.push(
        new Paragraph({
          text: "",
          pageBreakBefore: true,
        })
      );
    }
  });

  const doc = new Document({
    sections: [{
      properties: {},
      children: children,
    }],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, "DanhSachKyTen_M14.docx");
}
