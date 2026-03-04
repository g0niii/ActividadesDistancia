import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export const downloadExcel = async (resData) => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(resData.type === 'planilla' ? "Mapa del Curso" : "Recurso");

    const UCC_BLUE = '003087';
    const LIGHT_BLUE = 'E7F3FF';

    if (resData.type === 'planilla') {
        const titleRow = sheet.addRow(["PLANILLA DE MONTAJE - MAPA DEL CURSO"]);
        sheet.mergeCells('A1:D1');
        titleRow.getCell(1).font = { size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
        titleRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: UCC_BLUE } };
        titleRow.getCell(1).alignment = { horizontal: 'center' };
        titleRow.height = 30;

        sheet.addRow([]);
        const meta = [
            ["Asignatura", resData.subjectName || "-"],
            ["Carrera", resData.career || "-"],
            ["Docente", resData.contentAuthor || "-"],
            ["Link Drive", resData.driveLink || "-"]
        ];
        meta.forEach(r => {
            const row = sheet.addRow(r);
            row.getCell(1).font = { bold: true };
            row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F5F5F5' } };
        });

        sheet.addRow([]);
        const headRow = sheet.addRow(["MÓDULO / SECCIÓN", "ÍTEM", "TIPO", "OBSERVACIONES"]);
        headRow.eachCell(cell => {
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: UCC_BLUE } };
            cell.alignment = { horizontal: 'center' };
        });

        resData.mountingModules?.forEach(m => {
            const modRow = sheet.addRow([m.title?.toUpperCase() || 'MÓDULO', "", "", ""]);
            sheet.mergeCells(`A${modRow.number}:D${modRow.number}`);
            modRow.getCell(1).font = { bold: true, color: { argb: UCC_BLUE } };
            modRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_BLUE } };

            m.items?.forEach(it => {
                const typeMap = { 'page': 'Página', 'assignment': 'Tarea', 'quiz': 'Evaluación', 'discussion': 'Foro', 'file': 'Archivo', 'link': 'Enlace' };
                sheet.addRow(["", it.title, typeMap[it.type] || it.type, it.comment]);
            });
        });

        sheet.getColumn(1).width = 25;
        sheet.getColumn(2).width = 45;
        sheet.getColumn(3).width = 15;
        sheet.getColumn(4).width = 60;

    } else {
        const titleRow = sheet.addRow(["DATOS DEL RECURSO"]);
        sheet.mergeCells('A1:B1');
        titleRow.getCell(1).font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
        titleRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: UCC_BLUE } };

        const meta = [
            ["Título", resData.title],
            ["Tipo", resData.type.toUpperCase()],
            ["Módulo", resData.module],
            ["Grupal", resData.isGroup ? "SÍ" : "NO"]
        ];
        meta.forEach(m => {
            const r = sheet.addRow(m);
            r.getCell(1).font = { bold: true };
        });

        sheet.addRow([]);
        sheet.addRow(["CONSIGNA (Vista Previa)"]).font = { bold: true };
        sheet.addRow([resData.consigna?.replace(/<[^>]*>?/gm, '') || '']).alignment = { wrapText: true };

        sheet.getColumn(1).width = 25;
        sheet.getColumn(2).width = 90;
    }

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `${resData.title || 'Recurso'}.xlsx`);
};
