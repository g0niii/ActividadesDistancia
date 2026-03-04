import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, HeadingLevel, VerticalAlign, PageBreak, ImageRun, ExternalHyperlink } from 'docx';

export const buildDocxBlob = async (resInput) => {
    const resourcesArray = Array.isArray(resInput) ? resInput : [resInput];
    const children = [];
    const TABLE_WIDTH = 9070; // ~16cm

    const htmlToParagraphs = (html) => {
        if (!html || html === '<p><br></p>') return [];
        const result = [];
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = html;

        const processNode = (node, isBullet = false) => {
            if (node.nodeType === 3) { // Text
                return [new TextRun({ text: node.textContent, font: "Arial", size: 22 })];
            }
            if (node.nodeType === 1) { // Element
                const tag = node.tagName.toLowerCase();
                const children = [];
                node.childNodes.forEach(child => {
                    children.push(...processNode(child));
                });

                if (tag === 'strong' || tag === 'b') {
                    children.forEach(c => { if (c instanceof TextRun) c.options.bold = true; });
                }
                if (tag === 'em' || tag === 'i') {
                    children.forEach(c => { if (c instanceof TextRun) c.options.italic = true; });
                }

                return children;
            }
            return [];
        };

        tempDiv.childNodes.forEach(node => {
            if (node.nodeType === 1) {
                const tag = node.tagName.toLowerCase();
                if (tag === 'p' || tag === 'div') {
                    const runs = processNode(node);
                    if (runs.length > 0) {
                        result.push(new Paragraph({ children: runs, spacing: { after: 120 } }));
                    }
                } else if (tag === 'ul' || tag === 'ol') {
                    node.childNodes.forEach(li => {
                        if (li.nodeType === 1 && li.tagName.toLowerCase() === 'li') {
                            result.push(new Paragraph({
                                children: processNode(li),
                                bullet: { level: 0 },
                                spacing: { after: 120 }
                            }));
                        }
                    });
                } else {
                    const runs = processNode(node);
                    if (runs.length > 0) {
                        result.push(new Paragraph({ children: runs, spacing: { after: 120 } }));
                    }
                }
            }
        });

        return result.length > 0 ? result : [new Paragraph({ children: [new TextRun({ text: tempDiv.innerText, font: "Arial", size: 22 })], spacing: { after: 120 } })];
    };

    for (const res of resourcesArray) {
        // Title
        children.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 400 },
            children: [
                new TextRun({
                    text: res.type === 'planilla' ? "PLANILLA DE MONTAJE" : (res.title || "ACTIVIDAD"),
                    bold: true,
                    size: 32,
                    color: "003087",
                    font: "Arial"
                })
            ]
        }));

        if (res.type === 'planilla') {
            // Planilla Table
            children.push(new Table({
                width: { size: TABLE_WIDTH, type: WidthType.DXA },
                rows: [
                    { label: 'Asignatura', value: res.subjectName },
                    { label: 'Carrera', value: res.career },
                    { label: 'Docente', value: res.contentAuthor },
                    { label: 'Drive', value: res.driveLink }
                ].map(item => new TableRow({
                    children: [
                        new TableCell({ width: { size: 2500, type: WidthType.DXA }, shading: { fill: "F5F5F5" }, children: [new Paragraph({ children: [new TextRun({ text: item.label, bold: true, font: "Arial", size: 20 })] })] }),
                        new TableCell({ width: { size: 6570, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: item.value || "-", font: "Arial", size: 20 })] })] })
                    ]
                }))
            }));

            res.mountingModules?.forEach(m => {
                children.push(new Paragraph({
                    spacing: { before: 400, after: 100 },
                    shading: { fill: "e7f3ff" },
                    children: [new TextRun({ text: (m.title || "Módulo").toUpperCase(), bold: true, size: 24, font: "Arial", color: "003087" })]
                }));

                const items = m.items?.map(it => {
                    const icons = { 'page': '📄', 'assignment': '📝', 'quiz': '🚀', 'discussion': '💬', 'file': '📎', 'link': '🔗' };
                    return new TableRow({
                        children: [new TableCell({
                            children: [
                                new Paragraph({ children: [new TextRun({ text: icons[it.type] || '📄', size: 22 }), new TextRun({ text: " " + (it.title || '-'), bold: true, font: "Arial", size: 22 })], spacing: { before: 80, after: 40 } }),
                                it.comment ? new Paragraph({ indent: { left: 400 }, children: [new TextRun({ text: "Observación: " + it.comment, italic: true, size: 18, font: "Arial", color: "666666" })], spacing: { after: 80 } }) : null
                            ].filter(Boolean)
                        })]
                    });
                });

                if (items?.length > 0) {
                    children.push(new Table({ width: { size: TABLE_WIDTH, type: WidthType.DXA }, rows: items, borders: { insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "E1E3E4" }, top: { style: BorderStyle.NIL }, bottom: { style: BorderStyle.NIL }, left: { style: BorderStyle.NIL }, right: { style: BorderStyle.NIL } } }));
                }
            });
        } else {
            // General sections
            const sections = [
                { title: "CONSIGNA / INSTRUCCIONES", html: res.consigna },
                { title: "OBJETIVOS DE APRENDIZAJE", html: res.objectives }
            ];

            if (res.type !== 'evaluacion') {
                sections.push({ title: res.type === 'foro' ? "PAUTAS DE PARTICIPACIÓN" : "PAUTAS DE PRESENTACIÓN", html: res.guidelines });
                sections.push({ title: "CRITERIOS DE EVALUACIÓN", html: res.criteria });
            }

            sections.forEach(s => {
                if (!s.html || s.html === '<p><br></p>') return;
                children.push(new Paragraph({
                    spacing: { before: 300, after: 150 },
                    children: [new TextRun({ text: s.title, bold: true, color: "003087", size: 24, font: "Arial" })]
                }));
                children.push(...htmlToParagraphs(s.html));
            });

            if (res.type === 'evaluacion' && res.questions && res.questions.length > 0) {
                children.push(new Paragraph({
                    spacing: { before: 400, after: 200 },
                    children: [new TextRun({ text: "CUESTIONARIO DE EVALUACIÓN", bold: true, color: "003087", size: 24, font: "Arial" })]
                }));

                res.questions.forEach((q, qIdx) => {
                    children.push(new Paragraph({
                        spacing: { before: 200, after: 100 },
                        children: [new TextRun({ text: `P${qIdx + 1}: ${q.prompt}`, bold: true, font: "Arial", size: 22 })]
                    }));

                    if (q.type === 'opcion-multiple' || q.type === 'respuestas-multiples') {
                        q.options?.forEach(opt => {
                            children.push(new Paragraph({
                                indent: { left: 400 },
                                children: [
                                    new TextRun({ text: opt.correct ? " (x) " : " ( ) ", font: "Courier New", size: 20 }),
                                    new TextRun({ text: opt.text, font: "Arial", size: 20 })
                                ]
                            }));
                        });
                    }

                    if (q.type === 'verdadero-falso') {
                        children.push(new Paragraph({ indent: { left: 400 }, children: [new TextRun({ text: `Respuesta correcta: ${q.tfAnswer}`, italic: true, size: 20, font: "Arial" })] }));
                    }

                    if (q.type === 'estimulo' && q.nestedQuestions) {
                        q.nestedQuestions.forEach((nq, nIdx) => {
                            children.push(new Paragraph({
                                indent: { left: 400 },
                                spacing: { before: 100 },
                                children: [new TextRun({ text: `${qIdx + 1}.${nIdx + 1}: ${nq.prompt}`, bold: true, font: "Arial", size: 20 })]
                            }));
                        });
                    }
                });
            }

            if (res.isGroup) {
                children.push(new Paragraph({ spacing: { before: 400 } }));
                children.push(new Table({
                    width: { size: TABLE_WIDTH, type: WidthType.DXA },
                    borders: { left: { style: BorderStyle.SINGLE, size: 60, color: "28a745" }, top: { style: BorderStyle.DOTTED, size: 1, color: "dddddd" }, bottom: { style: BorderStyle.DOTTED, size: 1, color: "dddddd" }, right: { style: BorderStyle.DOTTED, size: 1, color: "dddddd" } },
                    rows: [new TableRow({
                        children: [new TableCell({
                            margins: { top: 200, bottom: 200, left: 300, right: 300 },
                            children: [
                                new Paragraph({ children: [new TextRun({ text: "¿Cómo unirte a tu grupo para esta actividad?", bold: true, size: 26, color: "155724", font: "Arial" })], spacing: { after: 150 } }),
                                new Paragraph({ children: [new TextRun({ text: "Esta es una actividad grupal. Pasos para participar:", size: 20, font: "Arial" })], spacing: { after: 100 } }),
                                new Paragraph({ children: [new TextRun({ text: "• En el menú, haz clic en \"Personas\" y luego en \"Grupos\".", size: 20, font: "Arial" })], bullet: { level: 0 } }),
                                new Paragraph({ children: [new TextRun({ text: "• Elige un grupo disponible y haz clic en \"Unirse\".", size: 20, font: "Arial" })], bullet: { level: 0 } }),
                                new Paragraph({ children: [new TextRun({ text: "• Sólo una persona del grupo debe realizar la entrega final.", size: 20, font: "Arial", bold: true })] })
                            ]
                        })]
                    })]
                }));
            }
        }
    }

    const doc = new Document({ sections: [{ children }] });
    return await Packer.toBlob(doc);
};
