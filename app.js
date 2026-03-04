document.addEventListener('DOMContentLoaded', () => {
    let resources = [];
    let editingId = null;
    let questionIndex = 0;
    const questionsList = document.getElementById('questionsList');

    const isGroupCheckbox = document.getElementById('isGroup');
    const groupInstructionsField = document.getElementById('groupInstructionsField');
    const sidebarList = document.getElementById('resourcesList');
    const downloadAllBtn = document.getElementById('downloadAllBtn');

    // Quill Editors Initialization
    const editors = {};
    const quillOptions = {
        theme: 'snow',
        modules: {
            toolbar: [
                ['bold', 'italic', 'underline'],
                [{ 'list': 'ordered' }, { 'list': 'bullet' }],
                ['image', 'clean']
            ]
        }
    };

    ['consigna', 'objectives', 'guidelines', 'criteria'].forEach(id => {
        const container = document.getElementById(id + 'Editor');
        if (container) {
            editors[id] = new Quill(container, quillOptions);
        }
    });

    function getEditorHtml(id) {
        return editors[id] ? editors[id].root.innerHTML : '';
    }

    function setEditorHtml(id, html) {
        if (editors[id]) {
            editors[id].root.innerHTML = html || '';
        }
    }

    // Toggle Group Instructions display dynamically
    isGroupCheckbox.addEventListener('change', (e) => {
        if (e.target.checked) {
            groupInstructionsField.style.display = 'block';

            // Allow smooth UI rendering before scrolling
            setTimeout(() => {
                groupInstructionsField.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }, 100);
        } else {
            groupInstructionsField.style.display = 'none';
        }
    });

    // Nav logic
    const selectionScreen = document.getElementById('selectionScreen');
    const formScreen = document.getElementById('formScreen');
    const backBtn = document.getElementById('backBtn');
    const resourceTypeInput = document.getElementById('resourceType');
    const formTitle = document.getElementById('formTitle');
    const questionsSection = document.getElementById('questionsSection');
    const groupCheckboxContainer = document.getElementById('groupCheckboxContainer');

    // Labels for dynamic form changing
    const guidelinesLabel = document.getElementById('guidelinesLabel');
    const criteriaLabel = document.getElementById('criteriaLabel');

    const catSelect = document.getElementById('category');
    const modInput = document.getElementById('module');
    const titleInput = document.getElementById('title');

    function updateTitlePrefix() {
        const cat = catSelect.value;
        const mod = modInput.value;
        if (!cat || !mod) return;

        const newPrefix = `${cat} M${mod} - `;
        const oldVal = titleInput.value;

        const regex = /^(Actividad obligatoria|Actividad sugerida|Foro obligatorio|Foro sugerido|Evaluación|Autoevaluación|Montaje|Referencia bibliográfica)\s+M\d+\s+-\s*/i;
        if (regex.test(oldVal)) {
            titleInput.value = oldVal.replace(regex, newPrefix);
        } else if (oldVal === '') {
            titleInput.value = newPrefix;
        } else {
            // If it doesn't match and isn't empty, just prepend
            titleInput.value = newPrefix + oldVal;
        }
    }

    catSelect.addEventListener('change', updateTitlePrefix);
    modInput.addEventListener('input', updateTitlePrefix);

    const typeConfigs = {
        'foro': {
            title: 'Generar Foro',
            subtitle1: 'Pautas de participación',
            subtitle2: 'Criterios de evaluación',
            ph1: 'Instructivo y reglas para participar...',
            ph2: '¿Qué se evaluará en el debate?'
        },
        'tarea': {
            title: 'Generar Tarea',
            subtitle1: 'Pautas de presentación',
            subtitle2: 'Criterios de evaluación',
            ph1: 'Formato, extensión, fechas, tipo de documento, etc.',
            ph2: 'Rúbricas, puntajes a considerar, etc.'
        },
        'evaluacion': {
            title: 'Generar Evaluación',
            subtitle1: 'Instrucciones de la evaluación',
            subtitle2: 'Configuración de la evaluación',
            ph1: 'Reglas para el cuestionario...',
            ph2: 'Tiempo, cantidad de intentos permitidos, etc.'
        },
        'planilla': {
            title: 'Planilla de Montaje (Mapa del Curso)',
            subtitle1: 'Notas del Asesor Pedagógico',
            subtitle2: 'Observaciones para Maquetación',
            ph1: 'Escribe aquí consejos o pautas generales para el curso...',
            ph2: 'Detalles técnicos de configuración estructural...'
        }
    };

    // Auto-list formatter is no longer needed with Quill
    /*
    const addAutoListListener = (id) => { ... }
    */

    document.querySelectorAll('.type-card').forEach(card => {
        card.addEventListener('click', () => {
            const type = card.getAttribute('data-type');
            const config = typeConfigs[type];

            resourceTypeInput.value = type;
            formTitle.textContent = config.title;

            // Update category options based on type and set specific placeholder examples
            // 1. GLOBAL VISIBILITY RESET (Prevent glitches between types)
            document.getElementById('defaultHeader').style.display = 'block';
            document.getElementById('planillaHeader').style.display = 'none';
            document.getElementById('titleGroup').style.display = 'block';
            questionsSection.style.display = 'none';
            document.getElementById('mountingSection').style.display = 'none';
            document.getElementById('guidelinesEditor').parentElement.style.display = 'block';
            document.getElementById('criteriaEditor').parentElement.style.display = 'block';
            document.getElementById('consignaEditor').parentElement.style.display = 'block';
            document.getElementById('objectivesEditor').parentElement.style.display = 'block';
            groupCheckboxContainer.style.display = 'block';

            // 2. REQUIRED FIELDS RESET
            titleInput.required = true;
            modInput.required = true;
            catSelect.required = true;
            ['subjectName', 'career', 'contentAuthor', 'driveLink'].forEach(id => document.getElementById(id).required = false);

            // 3. APPLY TYPE CONFIGURATION
            catSelect.innerHTML = '<option value="" disabled selected>Seleccionar...</option>';
            if (type === 'foro') {
                catSelect.innerHTML += '<option value="Foro obligatorio">Foro obligatorio</option><option value="Foro sugerido">Foro sugerido</option>';
                titleInput.placeholder = 'Ej: Foro sugerido M1 - Debate sobre unidad 1';
                groupCheckboxContainer.style.display = 'none';
            } else if (type === 'evaluacion') {
                catSelect.innerHTML += '<option value="Evaluación">Evaluación</option><option value="Autoevaluación">Autoevaluación</option>';
                titleInput.placeholder = 'Ej: Autoevaluación M1 - Cuestionario inicial';
                questionsSection.style.display = 'block';
                document.getElementById('guidelinesEditor').parentElement.style.display = 'none';
                document.getElementById('criteriaEditor').parentElement.style.display = 'none';
                groupCheckboxContainer.style.display = 'none';
            } else if (type === 'planilla') {
                catSelect.innerHTML += '<option value="Montaje" selected>Planilla de Montaje</option>';
                titleInput.value = 'Planilla de Montaje';
                document.getElementById('defaultHeader').style.display = 'none';
                document.getElementById('planillaHeader').style.display = 'block';
                document.getElementById('titleGroup').style.display = 'none';
                document.getElementById('mountingSection').style.display = 'block';
                document.getElementById('consignaEditor').parentElement.style.display = 'none';
                document.getElementById('objectivesEditor').parentElement.style.display = 'none';

                titleInput.required = false;
                modInput.required = false;
                catSelect.required = false;
                ['subjectName', 'career', 'contentAuthor', 'driveLink'].forEach(id => document.getElementById(id).required = true);

                catSelect.value = "Montaje";
                modInput.value = "0";
                groupCheckboxContainer.style.display = 'none';
            } else if (type === 'pagina') {
                catSelect.innerHTML += '<option value="Actividad sugerida">Contenido M1</option>';
                titleInput.placeholder = 'Ej: 1.1. Nombre de la página';
                groupCheckboxContainer.style.display = 'none';
            } else {
                catSelect.innerHTML += '<option value="Actividad obligatoria">Obligatoria</option><option value="Actividad sugerida">Sugerida</option>';
                titleInput.placeholder = 'Ej: Actividad obligatoria M1 - Trabajo final';
            }

            // Update labels dynamically
            guidelinesLabel.innerHTML = `${config.subtitle1} <span class="required">*</span>`;
            criteriaLabel.innerHTML = `${config.subtitle2} <span class="required">*</span>`;

            selectionScreen.style.display = 'none';
            formScreen.style.display = 'block';
            window.scrollTo({ top: 0, behavior: 'smooth' });

            // clear form for fresh start if not editing
            if (!editingId) {
                document.getElementById('resourceForm').reset();
                setEditorHtml('consigna', '');
                setEditorHtml('objectives', '');
                setEditorHtml('criteria', '');
                questionsList.innerHTML = '';
                document.getElementById('canvasModuleContainer').innerHTML = ''; // reset worksheet
                groupInstructionsField.style.display = 'none';
                document.getElementById('saveResourceBtn').textContent = 'Guardar Recurso';
            }

            isGroupCheckbox.checked = false;
            if (type !== 'planilla' && type !== 'evaluacion' && type !== 'foro' && type !== 'pagina') {
                isGroupCheckbox.checked = false;
            }
        });
    });

    backBtn.addEventListener('click', () => {
        formScreen.style.display = 'none';
        selectionScreen.style.display = 'block';
        editingId = null;
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    function renderResourcesList() {
        sidebarList.innerHTML = '';
        if (resources.length === 0) {
            sidebarList.innerHTML = '<p class="empty-msg" style="color:#5C666F; font-size: 0.95rem;">No hay recursos guardados.</p>';
            downloadAllBtn.style.display = 'none';
            return;
        }

        resources.forEach(res => {
            const div = document.createElement('div');
            div.className = 'act-item';
            div.innerHTML = `
                <div class="act-type">${typeConfigs[res.type].title.replace('Generar ', '')}</div>
                <h4>${res.title || 'Sin Título'}</h4>
            `;
            const actions = document.createElement('div');
            actions.style.display = 'flex';
            actions.style.gap = '8px';
            actions.style.marginTop = '8px';

            const editBtn = document.createElement('button');
            editBtn.type = 'button';
            editBtn.className = 'btn-small';
            editBtn.textContent = 'Editar';
            editBtn.onclick = () => loadFormData(res);

            const previewBtn = document.createElement('button');
            previewBtn.type = 'button';
            previewBtn.className = 'btn-small';
            previewBtn.textContent = '👀 Ver';
            previewBtn.onclick = () => showPreview(res);

            const delBtn = document.createElement('button');
            delBtn.type = 'button';
            delBtn.className = 'btn-small';
            delBtn.style.color = 'var(--error-color)';
            delBtn.textContent = 'Borrar';
            delBtn.onclick = () => {
                resources = resources.filter(r => r.id !== res.id);
                if (editingId === res.id) {
                    formScreen.style.display = 'none';
                    selectionScreen.style.display = 'block';
                    editingId = null;
                }
                renderResourcesList();
            };

            const downloadBtn = document.createElement('button');
            downloadBtn.type = 'button';
            downloadBtn.className = 'btn-small';
            downloadBtn.style.background = '#e7f3ff';
            downloadBtn.style.color = '#0055a2';
            downloadBtn.textContent = '📥 Docx';
            downloadBtn.onclick = async () => {
                const orig = downloadBtn.textContent;
                downloadBtn.textContent = '...';
                try {
                    const blob = await buildDocxBlob(res);
                    const safeTitle = (res.title || 'Recurso').normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_\-]/g, '');
                    window.saveAs(blob, `${safeTitle}.docx`);
                } catch (e) {
                    console.error(e);
                    alert("Error: " + e.message);
                } finally {
                    downloadBtn.textContent = orig;
                }
            };

            const excelBtn = document.createElement('button');
            excelBtn.type = 'button';
            excelBtn.className = 'btn-small';
            excelBtn.style.background = '#e6f4ea';
            excelBtn.style.color = '#137333';
            excelBtn.textContent = '📊 Excel';
            excelBtn.onclick = async () => {
                try {
                    const safeTitle = (res.title || 'Recurso').normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_\-]/g, '');
                    await downloadExcel(res, safeTitle);
                } catch (e) {
                    alert("Error: " + e.message);
                }
            };

            actions.appendChild(previewBtn);
            actions.appendChild(editBtn);
            actions.appendChild(downloadBtn);
            actions.appendChild(excelBtn);
            actions.appendChild(delBtn);
            div.appendChild(actions);
            sidebarList.appendChild(div);
        });

        downloadAllBtn.style.display = 'block';
    }

    const modal = document.getElementById('previewModal');
    const closeBtn = document.querySelector('.close-modal');
    const copyBtn = document.getElementById('copyToClipboardBtn');

    closeBtn.addEventListener('click', () => { modal.style.display = 'none'; });
    window.addEventListener('click', (e) => {
        if (e.target === modal) { modal.style.display = 'none'; }
    });

    copyBtn.addEventListener('click', async () => {
        const content = document.getElementById('previewContent');

        // Create a temporary element to hold the styled HTML for copying
        const tempElement = document.createElement('div');
        tempElement.innerHTML = content.innerHTML;

        // Inject essential inline styles for Google Docs compatibility
        tempElement.style.fontFamily = 'Arial, sans-serif';
        tempElement.style.color = '#333';

        tempElement.querySelectorAll('h1').forEach(h => h.style.cssText = 'font-family: Arial; color: #2d3b45; font-size: 24pt; font-weight: bold; margin-bottom: 20px;');
        tempElement.querySelectorAll('h2').forEach(h => h.style.cssText = 'font-family: Arial; color: #1967d2; font-size: 16pt; font-weight: bold; margin-top: 30px; margin-bottom: 15px; border-bottom: 2px solid #e8eaed; padding-bottom: 5px;');

        tempElement.querySelectorAll('p, div, li, span').forEach(el => {
            if (!el.getAttribute('style') || !el.style.fontFamily) el.style.fontFamily = 'Arial, sans-serif';
            if (!el.style.fontSize) el.style.fontSize = '11pt';
            if (!el.style.lineHeight) el.style.lineHeight = '1.5';
        });

        tempElement.querySelectorAll('table').forEach(tbl => {
            tbl.style.width = '100%';
            tbl.style.borderCollapse = 'collapse';
            tbl.style.marginBottom = '20px';
            tbl.setAttribute('border', '1');
            tbl.style.border = '1px solid #ddd';
        });

        tempElement.querySelectorAll('td').forEach(td => {
            td.style.border = '1px solid #ddd';
            td.style.padding = '12px';
            td.style.fontFamily = 'Arial, sans-serif';
            td.style.fontSize = '10.5pt';
        });

        tempElement.querySelectorAll('strong').forEach(s => s.style.fontWeight = 'bold');
        tempElement.querySelectorAll('em').forEach(e => e.style.fontStyle = 'italic');

        // Special handling for lists to ensure they look okay
        tempElement.querySelectorAll('ul, ol').forEach(list => {
            list.style.paddingLeft = '30px';
            list.style.marginTop = '10px';
            list.style.marginBottom = '10px';
        });

        // Add to body briefly to copy
        document.body.appendChild(tempElement);
        const range = document.createRange();
        range.selectNode(tempElement);
        window.getSelection().removeAllRanges();
        window.getSelection().addRange(range);

        try {
            document.execCommand('copy');
            const origText = copyBtn.textContent;
            copyBtn.textContent = '✅ ¡Copiado!';
            copyBtn.style.background = '#e6ffed';
            copyBtn.style.color = '#22863a';
            setTimeout(() => {
                copyBtn.textContent = origText;
                copyBtn.style.background = '#e8f0fe';
                copyBtn.style.color = '#1967d2';
            }, 2000);
        } catch (err) {
            alert('No se pudo copiar el contenido.');
        }

        document.body.removeChild(tempElement);
        window.getSelection().removeAllRanges();
    });
    function showPreview(act) {
        const c = document.getElementById('previewContent');

        let html = `<h1>${act.title || 'Sin Título'}</h1>`;

        if (act.type === 'planilla') {
            html = `<h1 style="margin-bottom:0.5rem; font-family: Arial; color: #2D3B45;">Planilla de Montaje</h1>`;
            html += `
            <table style="width:100%; border-collapse: collapse; margin-bottom:1.5rem; font-family: Arial; font-size: 10pt;">
                <tr>
                    <td style="border: 1px solid #C7CDD1; padding: 10px; background: #f8f9fa;"><strong>Asignatura:</strong> ${act.subjectName || '-'}</td>
                    <td style="border: 1px solid #C7CDD1; padding: 10px; background: #f8f9fa;"><strong>Carrera:</strong> ${act.career || '-'}</td>
                </tr>
                <tr>
                    <td style="border: 1px solid #C7CDD1; padding: 10px; background: #f8f9fa;"><strong>Docente:</strong> ${act.contentAuthor || '-'}</td>
                    <td style="border: 1px solid #C7CDD1; padding: 10px; background: #f8f9fa;"><strong>Drive:</strong> <a href="${act.driveLink}" target="_blank" style="color:#0055a2;">Link al Drive</a></td>
                </tr>
            </table>`;

            html += `<h2 style="border-bottom: 2px solid #C7CDD1; padding-bottom:5px; font-family: Arial; color: #2D3B45;">Mapa del Curso</h2><div class="canvas-preview-flow">`;
            const tempFlow = document.createElement('div');
            tempFlow.innerHTML = act.mountingHTML || '';
            const modules = tempFlow.querySelectorAll('.canvas-module');

            modules.forEach(m => {
                const title = m.querySelector('.mod-title')?.value || 'Módulo sin título';
                html += `<div style="margin-bottom: 20px;">
                    <div style="background: #f1f3f4; padding: 10px; font-weight: bold; border: 1px solid #dadce0; font-family: Arial;">${title}</div>
                    <table style="width:100%; border-collapse: collapse; font-family: Arial;">`;

                const items = m.querySelectorAll('.canvas-item');
                items.forEach(it => {
                    const itTitle = it.querySelector('.item-title')?.value || 'Ítem sin título';
                    const itType = it.querySelector('.item-type-select')?.value || 'page';
                    const itComment = it.querySelector('.item-comment')?.value || '';

                    const iconMap = { 'page': '📄', 'assignment': '📝', 'quiz': '🚀', 'discussion': '💬', 'file': '📎', 'link': '🔗' };

                    html += `<tr>
                        <td style="border-bottom: 1px solid #e1e3e4; padding: 10px; vertical-align: middle;">
                            <div style="display:flex; align-items:center; gap:8px;">
                                <span style="font-size: 1.2rem;">${iconMap[itType] || '📄'}</span>
                                <span style="font-weight: bold;">${itTitle}</span>
                            </div>
                            ${itComment ? `<div style="margin-top:4px; font-size:0.85rem; color:#666; font-style:italic; padding-left:28px;">Observación: ${itComment}</div>` : ''}
                        </td>
                    </tr>`;
                });
                html += `</table></div>`;
            });
            html += `</div>`;
        } else {
            html += `<h2 style="font-family: Arial; color: #2D3B45;">Consigna</h2><div class="preview-text-block" style="font-family: Arial; line-height: 1.6;">${act.consigna || ''}</div>`;
            html += `<h2 style="font-family: Arial; color: #2D3B45;">Objetivos</h2><div class="preview-text-block" style="font-family: Arial; line-height: 1.6;">${act.objectives || ''}</div>`;
        }

        const config = typeConfigs[act.type] || typeConfigs['tarea'];
        if (act.type !== 'evaluacion') {
            html += `<h2>${config.subtitle1}</h2><div class="preview-text-block">${act.guidelines || ''}</div>`;
            html += `<h2>${config.subtitle2}</h2><div class="preview-text-block">${act.criteria || ''}</div>`;
        }

        if (act.type === 'evaluacion' && act.questionsHTML) {
            html += `<h2>Cuestionario de Evaluación</h2>`;
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = act.questionsHTML;
            const qs = tempDiv.querySelectorAll('.question-block');
            html += `<ol>`;
            qs.forEach(q => {
                const qType = q.querySelector('h4').textContent;
                const qTitle = q.querySelector('.q-title-input')?.value || qType;
                const prompt = q.querySelector('.q-prompt')?.value || '';
                html += `<li><strong>${qTitle}</strong><br><em>${prompt.replace(/\\n/g, '<br>')}</em></li>`;
            });
            html += `</ol>`;
        }

        c.innerHTML = html;
        modal.style.display = 'block';
    }

    function syncDOMValues(container) {
        container.querySelectorAll('input, textarea, select').forEach(el => {
            if (el.type === 'checkbox' || el.type === 'radio') {
                if (el.checked) el.setAttribute('checked', 'checked');
                else el.removeAttribute('checked');
            } else {
                el.setAttribute('value', el.value);
                if (el.tagName === 'TEXTAREA') el.textContent = el.value;
                if (el.tagName === 'SELECT') {
                    Array.from(el.options).forEach(opt => {
                        if (opt.selected) opt.setAttribute('selected', 'selected');
                        else opt.removeAttribute('selected');
                    });
                }
            }
        });
    }

    function getFormData() {
        if (resourceTypeInput.value === 'planilla') {
            document.getElementById('title').value = document.getElementById('subjectName').value || 'Planilla de Montaje';
        }
        syncDOMValues(questionsList);
        syncDOMValues(document.getElementById('canvasModuleContainer'));
        return {
            id: editingId || Date.now().toString(),
            type: resourceTypeInput.value,
            category: document.getElementById('category').value,
            module: document.getElementById('module').value,
            title: document.getElementById('title').value,
            subjectName: document.getElementById('subjectName').value,
            career: document.getElementById('career').value,
            contentAuthor: document.getElementById('contentAuthor').value,
            driveLink: document.getElementById('driveLink').value,
            consigna: getEditorHtml('consigna'),
            objectives: getEditorHtml('objectives'),
            guidelines: getEditorHtml('guidelines'),
            criteria: getEditorHtml('criteria'),
            isGroup: isGroupCheckbox.checked,
            questionsHTML: questionsList.innerHTML,
            mountingHTML: document.getElementById('canvasModuleContainer').innerHTML
        };
    }

    function loadFormData(res) {
        editingId = res.id;
        document.querySelector(`.type-card[data-type="${res.type}"]`).click();

        resourceTypeInput.value = res.type;
        document.getElementById('category').value = res.category || '';
        document.getElementById('module').value = res.module || '';
        document.getElementById('title').value = res.title;
        document.getElementById('subjectName').value = res.subjectName || '';
        document.getElementById('career').value = res.career || '';
        document.getElementById('contentAuthor').value = res.contentAuthor || '';
        document.getElementById('driveLink').value = res.driveLink || '';
        setEditorHtml('consigna', res.consigna);
        setEditorHtml('objectives', res.objectives);
        setEditorHtml('guidelines', res.guidelines || '');
        setEditorHtml('criteria', res.criteria || '');
        isGroupCheckbox.checked = res.isGroup;
        isGroupCheckbox.dispatchEvent(new Event('change'));
        questionsList.innerHTML = res.questionsHTML || '';
        document.getElementById('canvasModuleContainer').innerHTML = res.mountingHTML || '';

        document.getElementById('saveResourceBtn').textContent = 'Guardar Cambios';
    }

    // Question Builder Logic
    function getQuestionHTML(type, index) {
        let specificFields = '';
        let promptLabel = type === 'estimulo' ? 'Contenido del estímulo (Texto/Historia)' : 'Raíz de la pregunta <span class="required">*</span>';
        let baseFields = `<div class="form-group"><label>${promptLabel}</label><textarea class="q-prompt" rows="2" required></textarea></div>`;

        switch (type) {
            case 'opcion-multiple':
                specificFields = `<label>Respuestas posibles</label><div class="q-options" data-type="radio"><div class="q-option" style="display:flex; align-items:center; gap:8px; margin-bottom:5px;"><input type="radio" name="q${index}_correct" checked> <input style="flex:1;" type="text" class="opt-text" placeholder="Respuesta *" required></div><div class="q-option" style="display:flex; align-items:center; gap:8px; margin-bottom:5px;"><input type="radio" name="q${index}_correct"> <input style="flex:1;" type="text" class="opt-text" placeholder="Respuesta *" required></div></div><button type="button" class="add-opt-btn btn-small">+ Respuesta</button>`;
                break;
            case 'verdadero-falso':
                specificFields = `<label>Opciones</label><div class="q-tf-options"><label style="display:block; margin-bottom:5px;"><input type="radio" name="q${index}_tf" class="q-tf-answer" value="Verdadero" checked> Verdadero</label><label style="display:block;"><input type="radio" name="q${index}_tf" class="q-tf-answer" value="Falso"> Falso</label></div>`;
                break;
            case 'respuestas-multiples':
                specificFields = `<label>Respuestas posibles</label><div class="q-options" data-type="checkbox"><div class="q-option" style="display:flex; align-items:center; gap:8px; margin-bottom:5px;"><input type="checkbox" class="opt-correct" checked> <input style="flex:1;" type="text" class="opt-text" placeholder="Respuesta *" required></div><div class="q-option" style="display:flex; align-items:center; gap:8px; margin-bottom:5px;"><input type="checkbox" class="opt-correct"> <input style="flex:1;" type="text" class="opt-text" placeholder="Respuesta *" required></div></div><button type="button" class="add-opt-btn btn-small">+ Respuesta</button>`;
                break;
            case 'completar-espacio':
                specificFields = `<p class="group-hint" style="margin-bottom:0.6rem;">Introduce una frase y rodea una palabra con acentos graves para indicar el lugar en el que el estudiante tiene que escribir la respuesta. (p. ej., "Las rosas son \`rojas\`, las violetas son \`azules\`")</p><label>Tipo de respuesta</label><select class="q-fib-type"><option value="Entrada abierta">Entrada abierta</option><option value="Menú desplegable">Menú desplegable</option><option value="Banco de palabras">Banco de palabras</option></select><label style="margin-top:10px;">Respuestas aceptadas (separadas por coma)</label><input type="text" class="q-fib-answers" placeholder="Ej: rojas, rojas oscuro" required>`;
                break;
            case 'menu-desplegable':
                specificFields = `<label>Texto con espacios (ej: El [color] de la [flor])</label><textarea class="q-dd-text" rows="2" required></textarea><label>Define las opciones por espacio (ej: color -> rojo, *azul)</label><textarea class="q-dd-options" rows="3" placeholder="color -> rojo, *azul\\nflor -> *rosa, margarita"></textarea>`;
                break;
            case 'coincidencia':
                specificFields = `<label>Pares de pregunta/respuesta</label><div class="q-pairs"><div class="q-pair" style="display:flex; gap:8px; margin-bottom:10px;"><input style="flex:1;" type="text" placeholder="Pregunta *" class="pair-premise" required> <input style="flex:1;" type="text" placeholder="Respuesta *" class="pair-match" required></div><div class="q-pair" style="display:flex; gap:8px; margin-bottom:10px;"><input style="flex:1;" type="text" placeholder="Pregunta *" class="pair-premise" required> <input style="flex:1;" type="text" placeholder="Respuesta *" class="pair-match" required></div></div><button type="button" class="add-pair-btn btn-small">+ Par de pregunta/respuesta</button><label style="margin-top:10px;">Distractores adicionales</label><div class="q-distractors"><input type="text" class="q-match-distractors" placeholder="Distractor *" style="margin-bottom:5px;"></div><button type="button" class="add-distractor-btn btn-small">+ Distractor</button>`;
                break;
            case 'numerica':
                specificFields = `<label>Tipo de respuesta</label><select class="q-num-type"><option value="exacta">Respuesta exacta</option><option value="rango">Rango (Entre - Y)</option><option value="margen">Respuesta con margen extra</option></select><div class="q-num-value-container"><input type="text" class="q-num-value" placeholder="Valor numérico de respuesta" required></div>`;
                break;
            case 'formula':
                specificFields = `<label>Fórmula matemática</label><input type="text" class="q-formula-text" placeholder="ej: 5 * x + y" required><label>Variables y sus rangos (ej: x=1-10, y=5-20)</label><input type="text" class="q-formula-vars" placeholder="x=1-10" required>`;
                break;
            case 'ensayo':
                specificFields = `<label>Límite de palabras (opcional)</label><input type="number" class="q-essay-limit" placeholder="Ej: 500"><label><input type="checkbox" class="q-essay-rich" checked> Habilitar editor de texto enriquecido</label><br><label><input type="checkbox" class="q-essay-spellcheck" checked> Habilitar corrector ortográfico</label>`;
                break;
            case 'carga-archivo':
                specificFields = `<label>Tipos de archivos permitidos (ej: pdf, docx, jpg; dejar vacío para todos)</label><input type="text" class="q-file-types" placeholder="pdf, docx"><label>Cantidad máxima permitida</label><input type="number" class="q-file-limit" value="1" min="1">`;
                break;
            case 'ordenamiento':
                specificFields = `<label>Elementos para ordenar (ingresarlos en el orden final correcto)</label><div class="q-order-items"><div class="q-order-item"><input type="text" class="order-text" placeholder="Paso 1" required></div><div class="q-order-item"><input type="text" class="order-text" placeholder="Paso 2" required></div></div><button type="button" class="add-order-btn btn-small">+ Elemento</button><label>Etiqueta superior (opcional)</label><input type="text" class="q-order-top" placeholder="Ej: Más antiguo"><label>Etiqueta inferior (opcional)</label><input type="text" class="q-order-bot" placeholder="Ej: Más reciente">`;
                break;
            case 'categorizacion':
                specificFields = `<label>Categorías</label><div class="q-categories"><div class="q-category" style="margin-bottom:10px; border:1px solid #e1e4e8; padding:10px; border-radius:4px; position:relative;"><label>Descripción *</label><input type="text" class="cat-desc" required><div class="cat-answers" style="margin-top:10px;"><input type="text" class="cat-ans" placeholder="Respuesta *" required style="margin-bottom:5px;"></div><button type="button" class="add-cat-ans-btn btn-small">+ Respuesta</button></div></div><button type="button" class="add-cat-btn btn-small">+ Categoría</button><label style="margin-top:15px;">Distractores adicionales</label><div class="q-distractors-cat"><input type="text" class="q-cat-distractors" placeholder="Distractor *" style="margin-bottom:5px;"></div><button type="button" class="add-distractor-cat-btn btn-small">+ Distractor</button>`;
                break;
            case 'zona-activa':
                specificFields = `<label>Descripción o Subida visual de la imagen</label><input type="text" class="q-hotspot-img" placeholder="Mapa de Sudamérica" required><label>Detalle de la zona correcta de clic</label><input type="text" class="q-hotspot-zone" placeholder="Contorno de Argentina" required>`;
                break;
            case 'estimulo':
                specificFields = `
                <div class="stimulus-questions-container" style="margin-top: 15px; border-left: 3px solid #0055a2; padding-left: 15px;">
                    <!-- Nested questions will be appended here -->
                </div>
                <div class="add-nested-q-wrapper" style="margin-top: 15px; background: #fdfdfd; padding: 10px; border: 1px dashed #ccc;">
                    <label style="display:block; margin-bottom:5px; font-weight:bold; color:#0055a2;">Añadir pregunta a este estímulo:</label>
                    <div style="display:flex; gap:10px;">
                        <select class="nested-q-type" style="flex:1; padding:0.4rem;">
                            <option value="opcion-multiple">Opción múltiple</option>
                            <option value="verdadero-falso">Verdadero o falso</option>
                            <option value="respuestas-multiples">Respuesta múltiple</option>
                            <option value="completar-espacio">Rellenar el espacio en blanco</option>
                            <option value="menu-desplegable">Menús desplegables múltiples</option>
                            <option value="coincidencia">Coincidencia</option>
                            <option value="numerica">Numérico</option>
                            <option value="formula">Fórmula</option>
                            <option value="ensayo">Ensayo</option>
                            <option value="carga-archivo">Carga del archivo</option>
                            <option value="ordenamiento">Ordenar</option>
                            <option value="categorizacion">Categorización</option>
                            <option value="zona-activa">Zona activa</option>
                            <option value="bloque-texto">Bloque de texto</option>
                        </select>
                        <button type="button" class="btn-small add-nested-q-btn">Agregar</button>
                    </div>
                </div>`;
                break;
            case 'bloque-texto':
                specificFields = `<p class="group-hint">Nota: Utilice este bloque para insertar texto informativo entre preguntas sin esperar una respuesta calificada por parte de los estudiantes.</p>`;
                break;
        }

        const titleMap = {
            'opcion-multiple': 'Opción Múltiple',
            'verdadero-falso': 'Verdadero o falso',
            'respuestas-multiples': 'Respuesta múltiple',
            'completar-espacio': 'Rellenar el espacio en blanco',
            'menu-desplegable': 'Menús desplegables múltiples',
            'coincidencia': 'Coincidencia',
            'numerica': 'Numérico',
            'formula': 'Fórmula',
            'ensayo': 'Ensayo',
            'carga-archivo': 'Carga del archivo',
            'ordenamiento': 'Ordenar',
            'categorizacion': 'Categorización',
            'zona-activa': 'Zona activa',
            'estimulo': 'Estímulo',
            'bloque-texto': 'Bloque de texto'
        };

        return `
                                                <div class="question-block" data-type="${type}">
                                                    <div class="question-header">
                                                        <h4>Pregunta: ${titleMap[type]}</h4>
                                                        <button type="button" class="remove-q-btn" title="Eliminar pregunta">
                                                            <svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
                                                        </button>
                                                    </div>
                                                    ${baseFields}
                                                    <div class="question-type-fields">
                                                        ${specificFields}
                                                    </div>
                                                </div>
                                                `;
    }

    document.addEventListener('click', (e) => {
        // Main add question
        if (e.target.closest('#addQuestionBtn')) {
            const type = document.getElementById('newQuestionType').value;
            const html = getQuestionHTML(type, questionIndex++);
            const wrapper = document.createElement('div');
            wrapper.innerHTML = html.trim();
            questionsList.appendChild(wrapper.firstChild);
            return;
        }

        // Handle Add block to Stimulus
        const addNestedBtn = e.target.closest('.add-nested-q-btn');
        if (addNestedBtn) {
            const wrapperElem = addNestedBtn.closest('.add-nested-q-wrapper');
            const type = wrapperElem.querySelector('.nested-q-type').value;
            const container = wrapperElem.previousElementSibling;
            const html = getQuestionHTML(type, questionIndex++);
            const wrapper = document.createElement('div');
            wrapper.innerHTML = html.trim();
            container.appendChild(wrapper.firstChild);
            return;
        }

        if (e.target.closest('.remove-q-btn')) {
            e.target.closest('.question-block').remove();
            return;
        }

        if (e.target.closest('.add-opt-btn')) {
            const block = e.target.closest('.question-block');
            const container = block.querySelector('.q-options');
            const qType = container.getAttribute('data-type');
            const newOpt = document.createElement('div');
            newOpt.className = 'q-option';
            newOpt.style = "display:flex; align-items:center; gap:8px; margin-bottom:5px;";
            newOpt.innerHTML = `<input type="${qType}" name="${qType === 'radio' ? 'q' + (questionIndex) + '_correct' : ''}"> <input style="flex:1;" type="text" class="opt-text" placeholder="Respuesta *" required> <button type="button" class="btn-small rm-opt" style="margin:0; padding:2px 6px;">X</button>`;
            container.appendChild(newOpt);
            return;
        }

        if (e.target.closest('.rm-opt')) {
            e.target.closest('.q-option, .q-pair, .q-order-item, .q-category, div').remove();
            return;
        }

        if (e.target.closest('.add-pair-btn')) {
            const container = e.target.closest('.question-block').querySelector('.q-pairs');
            const newPair = document.createElement('div');
            newPair.className = 'q-pair';
            newPair.style = "display:flex; gap:8px; margin-bottom:10px;";
            newPair.innerHTML = `<input style="flex:1;" type="text" placeholder="Pregunta *" class="pair-premise" required> <input style="flex:1;" type="text" placeholder="Respuesta *" class="pair-match" required> <button type="button" class="btn-small rm-opt" style="margin:0; padding:2px 6px;">X</button>`;
            container.appendChild(newPair);
            return;
        }

        if (e.target.closest('.add-distractor-btn')) {
            const container = e.target.closest('.question-block').querySelector('.q-distractors');
            const el = document.createElement('div');
            el.style = "display:flex; gap:8px; margin-bottom:5px;";
            el.innerHTML = `<input style="flex:1;" type="text" class="q-match-distractors" placeholder="Distractor *" required> <button type="button" class="btn-small rm-opt" style="margin:0; padding:2px 6px;">X</button>`;
            container.appendChild(el);
            return;
        }

        if (e.target.closest('.add-distractor-cat-btn')) {
            const container = e.target.closest('.question-block').querySelector('.q-distractors-cat');
            const el = document.createElement('div');
            el.style = "display:flex; gap:8px; margin-bottom:5px;";
            el.innerHTML = `<input style="flex:1;" type="text" class="q-cat-distractors" placeholder="Distractor *" required> <button type="button" class="btn-small rm-opt" style="margin:0; padding:2px 6px;">X</button>`;
            container.appendChild(el);
            return;
        }

        if (e.target.closest('.add-cat-btn')) {
            const container = e.target.closest('.question-block').querySelector('.q-categories');
            const newCat = document.createElement('div');
            newCat.className = 'q-category';
            newCat.style = "margin-bottom:10px; border:1px solid #e1e4e8; padding:10px; border-radius:4px; position:relative;";
            newCat.innerHTML = `<button type="button" class="btn-small rm-opt" style="position:absolute; top:10px; right:10px; margin:0; padding:2px 6px;">X</button><label>Descripción *</label><input type="text" class="cat-desc" required><div class="cat-answers" style="margin-top:10px;"><input type="text" class="cat-ans" placeholder="Respuesta *" required style="margin-bottom:5px;"></div><button type="button" class="add-cat-ans-btn btn-small">+ Respuesta</button>`;
            container.appendChild(newCat);
            return;
        }

        if (e.target.closest('.add-cat-ans-btn')) {
            const ansCont = e.target.closest('.q-category').querySelector('.cat-answers');
            const el = document.createElement('div');
            el.style = "display:flex; gap:8px; margin-bottom:5px;";
            el.innerHTML = `<input style="flex:1;" type="text" class="cat-ans" placeholder="Respuesta *" required> <button type="button" class="btn-small rm-opt" style="margin:0; padding:2px 6px;">X</button>`;
            ansCont.appendChild(el);
            return;
        }

        if (e.target.closest('.add-order-btn')) {
            const container = e.target.closest('.question-block').querySelector('.q-order-items');
            const newItem = document.createElement('div');
            newItem.className = 'q-order-item';
            newItem.style = "display:flex; gap:8px; margin-bottom:5px;";
            newItem.innerHTML = `<input style="flex:1;" type="text" class="order-text" placeholder="Siguiente paso" required> <button type="button" class="btn-small rm-opt" style="margin:0; padding:2px 6px;">X</button>`;
            container.appendChild(newItem);
            return;
        }

        // --- Canvas Module Builder Management ---
        if (e.target.closest('#addCanvasModuleBtn')) {
            const container = document.getElementById('canvasModuleContainer');
            const mod = document.createElement('div');
            mod.className = 'canvas-module';
            mod.innerHTML = `
                <div class="canvas-module-header">
                    <span class="drag-handle">⠿</span>
                    <input type="text" placeholder="Título del Módulo (ej: Módulo 1: Introducción)" class="mod-title">
                    <button type="button" class="rm-module-btn">Eliminar Módulo</button>
                </div>
                <div class="canvas-item-list"></div>
                <div class="canvas-module-footer">
                    <button type="button" class="add-canvas-item-btn secondary-btn">+ Agregar Ítem</button>
                </div>
            `;
            container.appendChild(mod);
            return;
        }

        if (e.target.closest('.add-canvas-item-btn')) {
            const itemList = e.target.closest('.canvas-module').querySelector('.canvas-item-list');
            const item = document.createElement('div');
            item.className = 'canvas-item';
            item.innerHTML = `
                <div class="canvas-item-main">
                    <span class="drag-handle">⠿</span>
                    <select class="item-type-select">
                        <option value="page">📄 Página</option>
                        <option value="assignment">📝 Tarea</option>
                        <option value="quiz">🚀 Evaluación</option>
                        <option value="discussion">💬 Foro</option>
                        <option value="file">📎 Archivo</option>
                        <option value="link">🔗 Enlace</option>
                    </select>
                    <input type="text" placeholder="1.1. Nombre de la página" class="item-title">
                    <div class="item-actions">
                        <button type="button" class="rm-item-btn btn-small">X</button>
                    </div>
                </div>
                <div class="canvas-item-details">
                    <input type="text" placeholder="Observaciones / Comentarios para maquetación..." class="item-comment">
                </div>
            `;
            itemList.appendChild(item);
            return;
        }

        if (e.target.closest('.rm-module-btn')) {
            e.target.closest('.canvas-module').remove();
            return;
        }

        if (e.target.closest('.rm-item-btn')) {
            e.target.closest('.canvas-item').remove();
            return;
        }
    });

    // Delegated Change listener for dynamic placeholders
    document.addEventListener('change', (e) => {
        if (e.target.closest('.item-type-select')) {
            const select = e.target.closest('.item-type-select');
            const item = select.closest('.canvas-item');
            const titleInput = item.querySelector('.item-title');
            const val = select.value;

            if (val === 'assignment') titleInput.placeholder = "Actividad sugerida/obligatoria M1 - Nombre";
            else if (val === 'quiz') titleInput.placeholder = "Evaluación/Autoevaluación M1";
            else if (val === 'discussion') titleInput.placeholder = "Foro M1 - Nombre";
            else if (val === 'link') titleInput.placeholder = "Enlace";
            else if (val === 'file') titleInput.placeholder = "Enlace del Archivo";
            else titleInput.placeholder = "1.1. Nombre de la página";
        }
    });

    document.getElementById('saveResourceBtn').addEventListener('click', () => {
        if (!document.getElementById('resourceForm').reportValidity()) return;
        const data = getFormData();
        const existingIdx = resources.findIndex(r => r.id === data.id);
        if (existingIdx >= 0) resources[existingIdx] = data;
        else resources.push(data);

        editingId = null;
        document.getElementById('resourceForm').reset();
        questionsList.innerHTML = '';
        document.getElementById('saveResourceBtn').textContent = 'Guardar Recurso';
        formScreen.style.display = 'none';
        selectionScreen.style.display = 'block';
        renderResourcesList();
    });

    document.getElementById('generateDocxBtn').addEventListener('click', async () => {
        if (!document.getElementById('resourceForm').reportValidity()) return;
        const btn = document.getElementById('generateDocxBtn');
        const orig = btn.innerHTML;
        btn.innerHTML = `<span>Generando...</span>`; btn.disabled = true;
        try {
            const resData = getFormData();
            const blob = await buildDocxBlob(resData);

            // FILENAME FORMATING FIX: 
            // Chrome running on file:/// breaks the automatic download entirely (gives UUIDs) if it sees spaces or weird chars
            const safeTitle = resData.title.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_\-]/g, '') || 'Recurso';

            window.saveAs(blob, `${safeTitle}.docx`);
        } catch (e) {
            console.error(e);
            alert('Error al generar Docx: ' + (e.message || 'Desconocido'));
        } finally {
            btn.innerHTML = orig; btn.disabled = false;
        }
    });

    document.getElementById('generateExcelBtn').addEventListener('click', async () => {
        if (!document.getElementById('resourceForm').reportValidity()) return;
        const resData = getFormData();
        try {
            const safeTitle = resData.title.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_\-]/g, '') || 'Recurso';
            await downloadExcel(resData, safeTitle);
        } catch (e) {
            console.error(e);
            alert('Error al generar Excel: ' + (e.message || 'Desconocido'));
        }
    });

    downloadAllBtn.addEventListener('click', async () => {
        if (resources.length === 0) return;
        const btn = downloadAllBtn;
        const orig = btn.innerHTML;
        btn.innerHTML = `Generando Word...`; btn.disabled = true;
        try {
            const blob = await buildDocxBlob(resources);
            window.saveAs(blob, `Todos_Los_Recursos.docx`);
        } catch (e) {
            console.error(e);
            alert('Error generando documento consolidado: ' + (e.message || 'Desconocido'));
        } finally {
            btn.innerHTML = orig; btn.disabled = false;
        }
    });

    // Expects an array of resources.
    // If a single resource is passed, wrap it in an array to match logic.
    async function buildDocxBlob(resInput) {
        const resourcesArray = Array.isArray(resInput) ? resInput : [resInput];
        if (!window.docx) throw new Error('Docx library not loaded');
        // Destructure definitions from window.docx
        const { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, HeadingLevel, Table, TableRow, TableCell, BorderStyle, WidthType, VerticalAlign } = window.docx;

        // Prepare children paragraphs for the document body
        const children = [];

        // Constants for document styling
        const TABLE_WIDTH = 9031; // ~16cm in DXA
        const HALF_TABLE = 4515;

        for (let i = 0; i < resourcesArray.length; i++) {
            const resData = resourcesArray[i];

            if (i > 0) {
                children.push(new Paragraph({
                    pageBreakBefore: true,
                    children: [new TextRun("")]
                }));
            }

            const title = resData.title;
            const consigna = resData.consigna;
            const objectives = resData.objectives;
            const guidelines = resData.guidelines;
            const criteria = resData.criteria;
            const isGroup = resData.isGroup;
            const type = resData.type;

            const addSection = (subtitle, htmlContent) => {
                children.push(
                    new Paragraph({
                        heading: HeadingLevel.HEADING_3,
                        spacing: { before: 200, after: 100 },
                        children: [
                            new TextRun({
                                text: subtitle,
                                bold: true,
                                size: 24, // 12pt
                                color: "2D3B45",
                                font: "Arial"
                            }),
                        ],
                    })
                );
                const parsedParagraphs = parseHTMLToDocx(htmlContent);
                children.push(...parsedParagraphs);
            };

            if (type === 'planilla') {
                children.push(
                    new Paragraph({
                        heading: HeadingLevel.HEADING_2,
                        spacing: { after: 200 },
                        children: [new TextRun({ text: "Planilla de Montaje - Mapa del Curso", bold: true, size: 32, font: "Arial", color: "2D3B45" })],
                    })
                );

                children.push(
                    new Table({
                        width: { size: TABLE_WIDTH, type: WidthType.DXA },
                        columnWidths: [HALF_TABLE, HALF_TABLE],
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({ width: { size: HALF_TABLE, type: WidthType.DXA }, children: [new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun({ text: "Asignatura: ", bold: true, font: "Arial", size: 18 }), new TextRun({ text: resData.subjectName || "-", font: "Arial", size: 18 })] })] }),
                                    new TableCell({ width: { size: HALF_TABLE, type: WidthType.DXA }, children: [new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun({ text: "Carrera: ", bold: true, font: "Arial", size: 18 }), new TextRun({ text: resData.career || "-", font: "Arial", size: 18 })] })] }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({ width: { size: HALF_TABLE, type: WidthType.DXA }, children: [new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun({ text: "Docente: ", bold: true, font: "Arial", size: 18 }), new TextRun({ text: resData.contentAuthor || "-", font: "Arial", size: 18 })] })] }),
                                    new TableCell({ width: { size: HALF_TABLE, type: WidthType.DXA }, children: [new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun({ text: "Drive del curso: ", bold: true, font: "Arial", size: 18 }), new TextRun({ text: resData.driveLink || "-", color: "0000FF", underline: {}, font: "Arial", size: 18 })] })] }),
                                ]
                            })
                        ],
                        borders: {
                            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "C7CDD1" },
                            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "C7CDD1" },
                            top: { style: BorderStyle.SINGLE, size: 1, color: "C7CDD1" },
                            bottom: { style: BorderStyle.SINGLE, size: 1, color: "C7CDD1" },
                            left: { style: BorderStyle.SINGLE, size: 1, color: "C7CDD1" },
                            right: { style: BorderStyle.SINGLE, size: 1, color: "C7CDD1" },
                        }
                    })
                );
            } else {
                children.push(
                    new Paragraph({
                        heading: HeadingLevel.HEADING_2,
                        spacing: { before: 100, after: 200 },
                        children: [
                            new TextRun({
                                text: title,
                                bold: true,
                                size: 32, // 16pt
                                font: "Arial"
                            }),
                        ],
                    })
                );
            }

            // --- 2. CONSIGNA ---
            addSection("Consigna", consigna);

            // --- 3. OBJECTIVES ---
            addSection("Objetivos", objectives);

            // --- 4. PRESENTATION GUIDELINES (DYNAMIC) ---
            const config = typeConfigs[type] || typeConfigs['tarea'];

            if (type !== 'evaluacion') {
                addSection(config.subtitle1, guidelines);
            }

            // --- 5. EVALUATION CRITERIA (DYNAMIC) ---
            if (type !== 'evaluacion' && type !== 'planilla') {
                addSection(config.subtitle2, criteria);
            }

            // --- 5b. CANVAS MODULES EXPORT (DYNAMIC) ---
            if (type === 'planilla') {
                const tempFlow = document.createElement('div');
                tempFlow.innerHTML = resData.mountingHTML || '';
                const modules = tempFlow.querySelectorAll('.canvas-module');

                modules.forEach(m => {
                    const modTitle = m.querySelector('.mod-title')?.value || 'Módulo';

                    // Module Header Row
                    children.push(new Paragraph({
                        heading: HeadingLevel.HEADING_3,
                        spacing: { before: 250, after: 100 },
                        children: [new TextRun({ text: modTitle, bold: true, size: 24, color: "2D3B45", font: "Arial" })],
                        shading: { fill: "f5f5f5" }
                    }));

                    const items = m.querySelectorAll('.canvas-item');
                    const tableRows = [];

                    items.forEach(it => {
                        const itTitle = it.querySelector('.item-title')?.value || '-';
                        const itType = it.querySelector('.item-type-select')?.value || 'page';
                        const itComment = it.querySelector('.item-comment')?.value || '';

                        const iconMap = { 'page': '📄', 'assignment': '📝', 'quiz': '🚀', 'discussion': '💬', 'file': '📎', 'link': '🔗' };

                        const cellContents = [
                            new Paragraph({
                                spacing: { before: 40, after: itComment ? 40 : 40 },
                                children: [
                                    new TextRun({ text: (iconMap[itType] || '📄') + " ", size: 20 }),
                                    new TextRun({ text: itTitle, size: 20, font: "Arial", bold: true })
                                ]
                            })
                        ];

                        if (itComment) {
                            cellContents.push(
                                new Paragraph({
                                    indent: { left: 350 },
                                    spacing: { after: 40 },
                                    children: [
                                        new TextRun({ text: "Observación: ", bold: true, size: 16, font: "Arial", color: "444444" }),
                                        new TextRun({ text: itComment, size: 16, font: "Arial", color: "555555", italic: true })
                                    ]
                                })
                            );
                        }

                        tableRows.push(new TableRow({
                            children: [
                                new TableCell({
                                    width: { size: TABLE_WIDTH, type: WidthType.DXA },
                                    verticalAlign: VerticalAlign.CENTER,
                                    children: cellContents,
                                    margins: { top: 60, bottom: 60, left: 100, right: 100 },
                                    borders: {
                                        bottom: { style: BorderStyle.SINGLE, size: 1, color: "E1E3E4" },
                                        top: { style: BorderStyle.NIL },
                                        left: { style: BorderStyle.NIL },
                                        right: { style: BorderStyle.NIL }
                                    }
                                })
                            ]
                        }));
                    });

                    if (tableRows.length > 0) {
                        children.push(new Table({
                            width: { size: TABLE_WIDTH, type: WidthType.DXA },
                            columnWidths: [TABLE_WIDTH],
                            rows: tableRows,
                            borders: {
                                top: { style: BorderStyle.NIL }, bottom: { style: BorderStyle.NIL },
                                left: { style: BorderStyle.NIL }, right: { style: BorderStyle.NIL },
                                insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "E1E3E4" },
                                insideVertical: { style: BorderStyle.NIL }
                            }
                        }));
                    }
                });

                if (guidelines && guidelines.length > 10) {
                    addSection("Notas del Asesor", guidelines);
                }
                if (criteria && criteria.length > 10) {
                    addSection("Observaciones de Maquetación", criteria);
                }
            }

            // --- 6. EVALUATION NEW QUIZ QUESTIONS PARSER ---
            if (type === 'evaluacion') {
                const tempDiv = document.createElement('div');
                if (resData.questionsHTML) tempDiv.innerHTML = resData.questionsHTML;
                const questionBlocks = tempDiv.querySelectorAll('.question-block');
                if (questionBlocks.length > 0) {
                    addSection("Cuestionario de Evaluación", "A continuación se desarrolla el banco de preguntas.");

                    questionBlocks.forEach((block, idx) => {
                        const qType = block.dataset.type;
                        const qTitle = block.querySelector('h4').textContent;

                        const prompt = block.querySelector('.q-prompt')?.value || '';
                        const qUserTitle = block.querySelector('.q-title-input')?.value || qTitle;

                        const isNested = !!block.closest('.stimulus-questions-container');
                        let titleIndent = isNested ? 400 : 0;
                        let textIndent = isNested ? 800 : 400;

                        children.push(
                            new Paragraph({
                                spacing: { before: 300, after: 120 },
                                indent: { left: titleIndent },
                                children: [
                                    new TextRun({ text: (isNested ? "↳ " : "") + qUserTitle, bold: true, size: 28, color: "003087", font: "Arial" })
                                ],
                            })
                        );

                        if (prompt) {
                            const lines = prompt.split('\n');
                            lines.forEach(line => {
                                const trimmed = line.trim();
                                if (trimmed !== '') {
                                    children.push(
                                        new Paragraph({
                                            spacing: { after: 120 },
                                            indent: { left: textIndent },
                                            children: [new TextRun({ text: trimmed, size: 24, font: "Arial", color: "333333" })]
                                        })
                                    );
                                }
                            });
                        }

                        // Parse the specifics based on logic.
                        let detailData = [];
                        switch (qType) {
                            case 'opcion-multiple':
                            case 'respuestas-multiples':
                                const opts = block.querySelectorAll('.q-option');
                                opts.forEach(opt => {
                                    const isChecked = opt.querySelector('input[type="radio"], input[type="checkbox"]').checked;
                                    const text = opt.querySelector('.opt-text').value;
                                    if (text) detailData.push((isChecked ? "✅ " : "⚪ ") + text);
                                });
                                break;
                            case 'verdadero-falso':
                                const checkedTF = block.querySelector('input[type="radio"]:checked');
                                detailData.push(`Respuesta correcta: ${checkedTF ? checkedTF.value : 'No seleccionada'}`);
                                break;
                            case 'completar-espacio':
                                detailData.push(`Tipo de respuesta: ${block.querySelector('.q-fib-type').value}`);
                                detailData.push(`Respuestas aceptadas: ${block.querySelector('.q-fib-answers').value}`);
                                break;
                            case 'menu-desplegable':
                                detailData.push(`Texto configurado: ${block.querySelector('.q-dd-text').value}`);
                                detailData.push(`Opciones: ${block.querySelector('.q-dd-options').value.split('\n').join(' | ')}`);
                                break;
                            case 'coincidencia':
                                const pairs = block.querySelectorAll('.q-pair');
                                pairs.forEach(p => {
                                    const prem = p.querySelector('.pair-premise').value;
                                    const mat = p.querySelector('.pair-match').value;
                                    if (prem || mat) detailData.push(`- ${prem} -> ${mat}`);
                                });
                                const distractors = Array.from(block.querySelectorAll('.q-match-distractors')).map(n => n.value).filter(Boolean);
                                if (distractors.length > 0) detailData.push(`Distractores adicionales: ${distractors.join(', ')}`);
                                break;
                            case 'categorizacion':
                                const cats = block.querySelectorAll('.q-category');
                                cats.forEach(c => {
                                    const desc = c.querySelector('.cat-desc').value;
                                    const answers = Array.from(c.querySelectorAll('.cat-ans')).map(n => n.value).filter(Boolean);
                                    if (desc) detailData.push(`Categoría: ${desc} -> ${answers.join(', ')}`);
                                });
                                const catDistractors = Array.from(block.querySelectorAll('.q-cat-distractors')).map(n => n.value).filter(Boolean);
                                if (catDistractors.length > 0) detailData.push(`Distractores adicionales: ${catDistractors.join(', ')}`);
                                break;
                            case 'numerica':
                                detailData.push(`Tipo: ${block.querySelector('.q-num-type').value}`);
                                detailData.push(`Valor/Rango: ${block.querySelector('.q-num-value').value}`);
                                break;
                            case 'formula':
                                detailData.push(`Fórmula: ${block.querySelector('.q-formula-text').value}`);
                                detailData.push(`Variables: ${block.querySelector('.q-formula-vars').value}`);
                                break;
                            case 'ensayo':
                                const limit = block.querySelector('.q-essay-limit').value;
                                const rich = block.querySelector('.q-essay-rich').checked;
                                const spell = block.querySelector('.q-essay-spellcheck').checked;
                                detailData.push(`Límite de palabras: ${limit || 'Ninguno'}`);
                                detailData.push(`Controles adicionales: ${rich ? 'Editor enriquecido ' : ''}${spell ? 'Corrector ortográfico' : ''}`);
                                break;
                            case 'carga-archivo':
                                detailData.push(`Extensiones/Tipos permitidos: ${block.querySelector('.q-file-types').value || 'Cualquiera'}`);
                                detailData.push(`Cantidad máxima permitida: ${block.querySelector('.q-file-limit').value}`);
                                break;
                            case 'ordenamiento':
                                let oIdx = 1;
                                block.querySelectorAll('.order-text').forEach(oi => {
                                    if (oi.value) detailData.push(`${oIdx++}. ${oi.value}`);
                                });
                                detailData.push(`Etiqueta Superior: ${block.querySelector('.q-order-top').value || '-'}`);
                                detailData.push(`Etiqueta Inferior: ${block.querySelector('.q-order-bot').value || '-'}`);
                                break;
                            case 'categorizacion':
                                detailData.push(`Estructura configurada:`);
                                block.querySelector('.q-cat-items').value.split('\n').forEach(l => {
                                    if (l.trim()) detailData.push(` - ${l}`);
                                });
                                const cd = block.querySelector('.q-cat-distractors').value;
                                if (cd) detailData.push(`Elementos adicionales/distractores: ${cd}`);
                                break;
                            case 'zona-activa':
                                detailData.push(`Imagen de referencia: ${block.querySelector('.q-hotspot-img').value}`);
                                detailData.push(`Configuración de zona activa de clic: ${block.querySelector('.q-hotspot-zone').value}`);
                                break;
                            case 'estimulo':
                                break;
                        }

                        // Append details elegantly to Word logic.
                        detailData.forEach(textLine => {
                            children.push(
                                new Paragraph({
                                    spacing: { after: 100 },
                                    indent: { left: textIndent * 2 },
                                    children: [new TextRun({ text: textLine, size: 22, font: "Courier New", color: "111111" })]
                                })
                            );
                        });
                    });
                }
            }

            // --- 7. GROUP REGISTRATION INSTRUCTIONS ---
            if (isGroup) {
                const cellChildren = [
                    new Paragraph({
                        spacing: { after: 200 },
                        children: [
                            new TextRun({
                                text: "¿Cómo unirte a tu grupo para esta actividad?",
                                bold: true,
                                size: 30, // 15pt
                                color: "007D40", // Canvas Green from image
                                font: "Arial"
                            })
                        ]
                    }),
                    new Paragraph({
                        spacing: { after: 120 },
                        children: [
                            new TextRun({
                                text: "Esta es una ",
                                size: 24, // 12pt
                                font: "Arial",
                                color: "333333"
                            }),
                            new TextRun({
                                text: "actividad grupal",
                                bold: true,
                                size: 24,
                                font: "Arial",
                                color: "333333"
                            }),
                            new TextRun({
                                text: ". Para participar, seguí estos pasos:",
                                size: 24,
                                font: "Arial",
                                color: "333333"
                            })
                        ]
                    })
                ];

                const listItems = [
                    "1. En el menú del curso, hacé clic en \"Personas\" y luego en la pestaña \"Grupos\"",
                    "2. Elegí un grupo con lugar disponible y hacé clic en \"Unirse\".",
                    "3. Pueden usar el entorno del grupo para comunicarse, organizarse y trabajar en conjunto.",
                    "4. La entrega se realiza desde esta misma actividad, no desde el entorno del grupo. Solo una persona del grupo debe entregar"
                ];

                listItems.forEach(line => {
                    let textRuns = [];
                    // Bold part of 4th step as per UI logic
                    if (line.startsWith("4.")) {
                        textRuns.push(
                            new TextRun({
                                text: "4. ",
                                size: 24,
                                font: "Arial",
                                color: "333333"
                            }),
                            new TextRun({
                                text: "La entrega se realiza desde esta misma actividad, no desde el entorno del grupo. ",
                                bold: true,
                                size: 24,
                                font: "Arial",
                                color: "333333"
                            }),
                            new TextRun({
                                text: "Solo una persona del grupo debe entregar",
                                size: 24,
                                font: "Arial",
                                color: "333333"
                            })
                        );
                    } else {
                        textRuns.push(
                            new TextRun({
                                text: line,
                                size: 24,
                                font: "Arial",
                                color: "333333"
                            })
                        );
                    }

                    cellChildren.push(
                        new Paragraph({
                            spacing: { after: 120 },
                            indent: { left: 400 }, // Small indentation for pseudo list
                            children: textRuns
                        })
                    );
                });

                cellChildren.push(
                    new Paragraph({
                        spacing: { before: 120, after: 120 },
                        children: [
                            new TextRun({
                                text: "⚠️ Importante: ",
                                bold: true,
                                size: 24,
                                font: "Arial",
                                color: "8A5214" // Warning text
                            }),
                            new TextRun({
                                text: "Si no ves los grupos o tenés problemas, contactá al docente.",
                                size: 24,
                                font: "Arial",
                                color: "555555"
                            })
                        ]
                    })
                );

                // Spacing immediately before the table
                children.push(
                    new Paragraph({
                        spacing: { before: 100, after: 50 },
                        children: [new TextRun({ text: "" })]
                    })
                );

                // Creates the Border Box using a single-cell Docx Table to enclose instructions visually like a card.
                children.push(
                    new Table({
                        width: {
                            size: TABLE_WIDTH,
                            type: WidthType.DXA,
                        },
                        columnWidths: [TABLE_WIDTH],
                        borders: {
                            top: { style: BorderStyle.DOTTED, size: 8, color: "C7CDD1" },
                            bottom: { style: BorderStyle.DOTTED, size: 8, color: "C7CDD1" },
                            left: { style: BorderStyle.SINGLE, size: 64, color: "2D8A41" },
                            right: { style: BorderStyle.DOTTED, size: 8, color: "C7CDD1" },
                        },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        width: { size: TABLE_WIDTH, type: WidthType.DXA },
                                        margins: {
                                            top: 200,
                                            bottom: 200,
                                            left: 250,
                                            right: 250,
                                        },
                                        shading: {
                                            fill: "FFFFFF"
                                        },
                                        children: cellChildren
                                    })
                                ]
                            })
                        ]
                    })
                );
            }
        } // End of activities loop

        // Create docx Object Structure
        const doc = new Document({
            creator: "Estandarizador Canvas LMS",
            title: resourcesArray.length === 1 ? resourcesArray[0].title : "Actividades Compiladas",
            description: "Actividades estandarizadas para Canvas",
            sections: [{ properties: {}, children: children }],
        });

        // Convert to BLOB and return
        return await Packer.toBlob(doc);
    }

    async function downloadExcel(resData, filename) {
        if (!window.ExcelJS) throw new Error('La librería ExcelJS no se ha cargado.');
        const workbook = new window.ExcelJS.Workbook();
        const sheetName = resData.type === 'planilla' ? "Mapa del Curso" : "Recurso";
        const sheet = workbook.addWorksheet(sheetName);

        const UCC_BLUE = 'FF0055A2';
        const LIGHT_BLUE = 'FFE7F3FF';

        if (resData.type === 'planilla') {
            // Title Header
            const titleRow = sheet.addRow(["PLANILLA DE MONTAJE - MAPA DEL CURSO"]);
            sheet.mergeCells('A1:D1');
            titleRow.getCell(1).font = { size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
            titleRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: UCC_BLUE } };
            titleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
            titleRow.height = 30;

            sheet.addRow([]); // Spacer

            // Metadata
            const metaRows = [
                ["Asignatura", resData.subjectName || "-"],
                ["Carrera", resData.career || "-"],
                ["Docente", resData.contentAuthor || "-"],
                ["Drive", resData.driveLink || "-"]
            ];

            metaRows.forEach(row => {
                const r = sheet.addRow(row);
                r.getCell(1).font = { bold: true, color: { argb: 'FF333333' } };
                r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
                r.getCell(1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                r.getCell(2).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            });

            sheet.addRow([]); // Spacer

            // Table Header
            const headRow = sheet.addRow(["MÓDULO / SECCIÓN", "ÍTEM", "TIPO", "OBSERVACIONES"]);
            headRow.eachCell(cell => {
                cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: UCC_BLUE } };
                cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            });
            headRow.height = 25;

            // Data
            const tempFlow = document.createElement('div');
            tempFlow.innerHTML = resData.mountingHTML || '';
            const modules = tempFlow.querySelectorAll('.canvas-module');

            modules.forEach(m => {
                const modTitle = m.querySelector('.mod-title')?.value || 'Módulo';
                const modRow = sheet.addRow([modTitle.toUpperCase(), "", "", ""]);
                sheet.mergeCells(`A${modRow.number}:D${modRow.number}`);
                modRow.getCell(1).font = { bold: true, color: { argb: 'FF0055A2' } };
                modRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_BLUE } };
                modRow.getCell(1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                modRow.height = 20;

                const items = m.querySelectorAll('.canvas-item');
                items.forEach(it => {
                    const itTitle = it.querySelector('.item-title')?.value || '-';
                    const itTypeRaw = it.querySelector('.item-type-select')?.value || 'page';
                    const itComment = it.querySelector('.item-comment')?.value || '';

                    const typeMap = {
                        'page': 'Página',
                        'assignment': 'Tarea / Actividad',
                        'quiz': 'Evaluación / Quiz',
                        'discussion': 'Foro',
                        'file': 'Archivo / Documento',
                        'link': 'Enlace Externo'
                    };
                    const itType = typeMap[itTypeRaw] || itTypeRaw;

                    const row = sheet.addRow(["", itTitle, itType, itComment]);
                    row.eachCell(cell => {
                        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                        cell.alignment = { vertical: 'middle', wrapText: true };
                    });
                });
            });

            sheet.getColumn(1).width = 25;
            sheet.getColumn(2).width = 45;
            sheet.getColumn(3).width = 15;
            sheet.getColumn(4).width = 60;

        } else {
            // Tarea / Foro / Evaluacion
            const titleRow = sheet.addRow(["DATOS DEL RECURSO"]);
            sheet.mergeCells('A1:B1');
            titleRow.getCell(1).font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
            titleRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: UCC_BLUE } };
            titleRow.height = 25;

            const meta = [
                ["Título", resData.title],
                ["Tipo", resData.type],
                ["Categoría", resData.category],
                ["Módulo", resData.module],
                ["Grupal", resData.isGroup ? "Sí" : "No"]
            ];

            meta.forEach(m => {
                const r = sheet.addRow(m);
                r.getCell(1).font = { bold: true };
                r.getCell(1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                r.getCell(2).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            });

            sheet.addRow([]); // Spacer
            const contHead = sheet.addRow(["CAMPO", "CONTENIDO (Texto Limpio)"]);
            contHead.eachCell(c => {
                c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: UCC_BLUE } };
                c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            });

            const content = [
                ["Consigna", stripHtml(resData.consigna)],
                ["Objetivos", stripHtml(resData.objectives)],
                ["Pautas 1", stripHtml(resData.guidelines)],
                ["Pautas 2", stripHtml(resData.criteria)]
            ];

            content.forEach(c => {
                const r = sheet.addRow(c);
                r.getCell(1).font = { bold: true };
                r.getCell(2).alignment = { wrapText: true, vertical: 'top' };
                r.eachCell(cell => cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } });
            });

            sheet.getColumn(1).width = 25;
            sheet.getColumn(2).width = 90;
        }

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        window.saveAs(blob, `${filename}.xlsx`);
    }

    function stripHtml(html) {
        if (!html) return "";
        const tmp = document.createElement("DIV");
        tmp.innerHTML = html;
        return tmp.textContent || tmp.innerText || "";
    }

    function parseHTMLToDocx(html) {
        if (!html) return [];

        const { Paragraph, TextRun, ImageRun } = window.docx;
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = html;

        const results = [];

        function processNode(node, currentRuns = [], inheritedFormatting = {}) {
            if (node.nodeType === 3) { // Text node
                if (node.textContent.trim() || node.textContent === ' ') {
                    currentRuns.push(new TextRun({
                        text: node.textContent,
                        size: 22, // 11pt
                        font: "Arial",
                        ...inheritedFormatting
                    }));
                }
            } else if (node.nodeType === 1) { // Element node
                const tagName = node.tagName.toLowerCase();
                const newFormatting = { ...inheritedFormatting };

                if (['strong', 'b'].includes(tagName)) newFormatting.bold = true;
                if (['em', 'i'].includes(tagName)) newFormatting.italic = true;
                if (tagName === 'u') newFormatting.underline = {};

                if (['p', 'div', 'li'].includes(tagName)) {
                    const childRuns = [];
                    node.childNodes.forEach(child => processNode(child, childRuns, newFormatting));

                    if (childRuns.length > 0) {
                        const paraOptions = {
                            children: childRuns,
                            spacing: { after: 80 }
                        };
                        if (tagName === 'li') paraOptions.bullet = { level: 0 };
                        results.push(new Paragraph(paraOptions));
                    }
                } else if (tagName === 'br') {
                    currentRuns.push(new TextRun({ break: 1 }));
                } else if (tagName === 'img') {
                    const src = node.getAttribute('src');
                    if (src && src.startsWith('data:image')) {
                        const base64Data = src.split(',')[1];
                        try {
                            const binaryData = Uint8Array.from(atob(base64Data), c => c.charCodeAt(0));
                            currentRuns.push(new ImageRun({
                                data: binaryData,
                                transformation: { width: 550, height: 350 },
                            }));
                        } catch (err) {
                            console.error("Error processing image for docx:", err);
                        }
                    }
                } else {
                    node.childNodes.forEach(child => processNode(child, currentRuns, newFormatting));
                }
            }
        }

        tempDiv.childNodes.forEach(node => {
            const tagName = node.nodeType === 1 ? node.tagName.toLowerCase() : null;
            if (['p', 'li', 'div', 'ol', 'ul'].includes(tagName)) {
                if (tagName === 'ol' || tagName === 'ul') {
                    node.childNodes.forEach(li => {
                        if (li.nodeType === 1 && li.tagName.toLowerCase() === 'li') {
                            processNode(li);
                        }
                    });
                } else {
                    processNode(node);
                }
            } else {
                const strayRuns = [];
                processNode(node, strayRuns);
                if (strayRuns.length > 0) {
                    results.push(new Paragraph({
                        children: strayRuns,
                        spacing: { after: 80 }
                    }));
                }
            }
        });

        return results;
    }

    renderResourcesList();
});
