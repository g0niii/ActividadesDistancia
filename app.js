document.addEventListener('DOMContentLoaded', () => {
    let activities = [];
    let editingId = null;
    let questionIndex = 0;
    const questionsList = document.getElementById('questionsList');

    const isGroupCheckbox = document.getElementById('isGroup');
    const groupInstructionsField = document.getElementById('groupInstructionsField');
    const sidebarList = document.getElementById('activitiesList');
    const downloadAllBtn = document.getElementById('downloadAllBtn');

    // Toggle Group Instructions display dynamically
    isGroupCheckbox.addEventListener('change', (e) => {
        if (e.target.checked) {
            groupInstructionsField.style.display = 'block';

            // Allow smooth UI rendering before scrolling
            setTimeout(() => {
                groupInstructionsField.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }, 100);
        } else {
            groupInstructionsField.style.display = 'none';
        }
    });

    // Nav logic
    const selectionScreen = document.getElementById('selectionScreen');
    const formScreen = document.getElementById('formScreen');
    const backBtn = document.getElementById('backBtn');
    const activityTypeInput = document.getElementById('activityType');
    const formTitle = document.getElementById('formTitle');
    const questionsSection = document.getElementById('questionsSection');
    const groupCheckboxContainer = document.getElementById('groupCheckboxContainer');

    // Labels for dynamic form changing
    const guidelinesLabel = document.getElementById('guidelinesLabel');
    const criteriaLabel = document.getElementById('criteriaLabel');
    const guidelinesInput = document.getElementById('guidelines');
    const criteriaInput = document.getElementById('criteria');

    const catSelect = document.getElementById('category');
    const modInput = document.getElementById('module');
    const titleInput = document.getElementById('title');

    function updateTitlePrefix() {
        const cat = catSelect.value;
        const mod = modInput.value;
        if (!cat || !mod) return;

        const newPrefix = `${cat} M${mod} - `;
        const oldVal = titleInput.value;

        const regex = /^(Actividad obligatoria|Actividad sugerida|Foro obligatorio|Foro sugerido|Evaluación|Autoevaluación)\s+M\d+\s+-\s*/i;
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
        }
    };

    // Auto-list formatter for specific text areas
    const addAutoListListener = (id) => {
        const el = document.getElementById(id);
        if (!el) return;

        el.addEventListener('focus', () => {
            if (el.value.trim() === '') {
                el.value = '1. ';
            }
        });

        el.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                const lines = el.value.substring(0, el.selectionStart).split('\n');
                const lastLine = lines[lines.length - 1];
                const match = lastLine.match(/^(\d+)\.\s*/);

                if (match) {
                    e.preventDefault();
                    if (/^\d+\.\s*$/.test(lastLine)) {
                        // If it's just an empty number, remove it and just add new line
                        const start = el.value.lastIndexOf(lastLine);
                        el.value = el.value.substring(0, start) + '\n' + el.value.substring(el.selectionEnd);
                        el.selectionStart = el.selectionEnd = start + 1;
                        return;
                    }
                    const nextNum = parseInt(match[1], 10) + 1;
                    const start = el.selectionStart;
                    const end = el.selectionEnd;
                    const insertText = '\n' + nextNum + '. ';
                    el.value = el.value.substring(0, start) + insertText + el.value.substring(end);
                    el.selectionStart = el.selectionEnd = start + insertText.length;
                }
            }
        });
    };

    addAutoListListener('objectives');
    addAutoListListener('guidelines');
    addAutoListListener('criteria');

    document.querySelectorAll('.type-card').forEach(card => {
        card.addEventListener('click', () => {
            const type = card.getAttribute('data-type');
            const config = typeConfigs[type];

            activityTypeInput.value = type;
            formTitle.textContent = config.title;

            // Update category options based on type and set specific placeholder examples
            catSelect.innerHTML = '<option value="" disabled selected>Seleccionar...</option>';
            if (type === 'foro') {
                catSelect.innerHTML += '<option value="Foro obligatorio">Foro obligatorio</option><option value="Foro sugerido">Foro sugerido</option>';
                titleInput.placeholder = 'Ej: Foro sugerido M1 - Debate sobre unidad 1';
            } else if (type === 'evaluacion') {
                catSelect.innerHTML += '<option value="Evaluación">Evaluación</option><option value="Autoevaluación">Autoevaluación</option>';
                titleInput.placeholder = 'Ej: Autoevaluación M1 - Cuestionario inicial';
            } else {
                catSelect.innerHTML += '<option value="Actividad obligatoria">Obligatoria</option><option value="Actividad sugerida">Sugerida</option>';
                titleInput.placeholder = 'Ej: Actividad obligatoria M1 - Trabajo final';
            }

            // Update labels dynamically
            guidelinesLabel.innerHTML = `${config.subtitle1} <span class="required">*</span>`;
            criteriaLabel.innerHTML = `${config.subtitle2} <span class="required">*</span>`;
            guidelinesInput.placeholder = config.ph1;
            criteriaInput.placeholder = config.ph2;

            selectionScreen.style.display = 'none';
            formScreen.style.display = 'block';

            // clear form for fresh start if not editing
            if (!editingId) {
                document.getElementById('activityForm').reset();
                questionsList.innerHTML = '';
                groupInstructionsField.style.display = 'none';
                document.getElementById('saveActivityBtn').textContent = 'Guardar Actividad';
            }

            if (type === 'evaluacion') {
                questionsSection.style.display = 'block';
                guidelinesInput.parentElement.style.display = 'none';
                guidelinesInput.required = false;
                criteriaInput.parentElement.style.display = 'none';
                criteriaInput.required = false;

                groupCheckboxContainer.style.display = 'none';
                isGroupCheckbox.checked = false;
            } else if (type === 'foro') {
                questionsSection.style.display = 'none';
                guidelinesInput.parentElement.style.display = 'block';
                guidelinesInput.required = true;
                criteriaInput.parentElement.style.display = 'block';
                criteriaInput.required = true;

                groupCheckboxContainer.style.display = 'none';
                isGroupCheckbox.checked = false;
            } else {
                questionsSection.style.display = 'none';
                guidelinesInput.parentElement.style.display = 'block';
                guidelinesInput.required = true;
                criteriaInput.parentElement.style.display = 'block';
                criteriaInput.required = true;

                groupCheckboxContainer.style.display = 'block';
            }
        });
    });

    backBtn.addEventListener('click', () => {
        formScreen.style.display = 'none';
        selectionScreen.style.display = 'block';
        editingId = null;
    });

    function renderActivitiesList() {
        sidebarList.innerHTML = '';
        if (activities.length === 0) {
            sidebarList.innerHTML = '<p class="empty-msg" style="color:#5C666F; font-size: 0.95rem;">No hay actividades guardadas.</p>';
            downloadAllBtn.style.display = 'none';
            return;
        }

        activities.forEach(act => {
            const div = document.createElement('div');
            div.className = 'act-item';
            div.innerHTML = `
                <div class="act-type">${typeConfigs[act.type].title.replace('Generar ', '')}</div>
                <h4>${act.title || 'Sin Título'}</h4>
            `;
            const actions = document.createElement('div');
            actions.style.display = 'flex';
            actions.style.gap = '8px';
            actions.style.marginTop = '8px';

            const editBtn = document.createElement('button');
            editBtn.type = 'button';
            editBtn.className = 'btn-small';
            editBtn.textContent = 'Editar';
            editBtn.onclick = () => loadFormData(act);

            const previewBtn = document.createElement('button');
            previewBtn.type = 'button';
            previewBtn.className = 'btn-small';
            previewBtn.textContent = '👀 Ver';
            previewBtn.onclick = () => showPreview(act);

            const delBtn = document.createElement('button');
            delBtn.type = 'button';
            delBtn.className = 'btn-small';
            delBtn.style.color = 'var(--error-color)';
            delBtn.textContent = 'Borrar';
            delBtn.onclick = () => {
                activities = activities.filter(a => a.id !== act.id);
                if (editingId === act.id) {
                    formScreen.style.display = 'none';
                    selectionScreen.style.display = 'block';
                    editingId = null;
                }
                renderActivitiesList();
            };

            actions.appendChild(previewBtn);
            actions.appendChild(editBtn);
            actions.appendChild(delBtn);
            div.appendChild(actions);
            sidebarList.appendChild(div);
        });

        downloadAllBtn.style.display = 'block';
    }

    const modal = document.getElementById('previewModal');
    const closeBtn = document.querySelector('.close-modal');

    closeBtn.addEventListener('click', () => { modal.style.display = 'none'; });
    window.addEventListener('click', (e) => {
        if (e.target === modal) { modal.style.display = 'none'; }
    });

    function showPreview(act) {
        const c = document.getElementById('previewContent');

        let html = `<h1>${act.title || 'Sin Título'}</h1>`;
        html += `<h2>Consigna</h2><p>${(act.consigna || '').replace(/\\n/g, '<br>')}</p>`;
        html += `<h2>Objetivos</h2><p>${(act.objectives || '').replace(/\\n/g, '<br>')}</p>`;

        const config = typeConfigs[act.type] || typeConfigs['tarea'];
        if (act.type !== 'evaluacion') {
            html += `<h2>${config.subtitle1}</h2><p>${(act.guidelines || '').replace(/\\n/g, '<br>')}</p>`;
            html += `<h2>${config.subtitle2}</h2><p>${(act.criteria || '').replace(/\\n/g, '<br>')}</p>`;
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
        syncDOMValues(questionsList);
        return {
            id: editingId || Date.now().toString(),
            type: activityTypeInput.value,
            category: document.getElementById('category').value,
            module: document.getElementById('module').value,
            title: document.getElementById('title').value,
            consigna: document.getElementById('consigna').value,
            objectives: document.getElementById('objectives').value,
            guidelines: document.getElementById('guidelines').value,
            criteria: document.getElementById('criteria').value,
            isGroup: isGroupCheckbox.checked,
            questionsHTML: questionsList.innerHTML
        };
    }

    function loadFormData(act) {
        editingId = act.id;
        document.querySelector(`.type-card[data-type="${act.type}"]`).click();

        activityTypeInput.value = act.type;
        document.getElementById('category').value = act.category || '';
        document.getElementById('module').value = act.module || '';
        document.getElementById('title').value = act.title;
        document.getElementById('consigna').value = act.consigna;
        document.getElementById('objectives').value = act.objectives;
        document.getElementById('guidelines').value = act.guidelines || '';
        document.getElementById('criteria').value = act.criteria || '';
        isGroupCheckbox.checked = act.isGroup;
        isGroupCheckbox.dispatchEvent(new Event('change'));
        questionsList.innerHTML = act.questionsHTML;

        document.getElementById('saveActivityBtn').textContent = 'Guardar Cambios';
    }

    // Question Builder Logic
    function getQuestionHTML(type, index) {
        let specificFields = '';
        let promptLabel = type === 'estimulo' ? 'Contenido del estímulo (Texto/Historia)' : 'Raíz de la pregunta <span class="required">*</span>';
        let baseFields = `<div class="form-group" style="margin-bottom:0.8rem;"><label>Título de la pregunta</label><input type="text" class="q-title-input" placeholder="Título de la pregunta"></div><div class="form-group"><label>${promptLabel}</label><textarea class="q-prompt" rows="2" required></textarea></div>`;

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
    });

    document.getElementById('saveActivityBtn').addEventListener('click', () => {
        if (!document.getElementById('activityForm').reportValidity()) return;
        const data = getFormData();
        const existingIdx = activities.findIndex(a => a.id === data.id);
        if (existingIdx >= 0) activities[existingIdx] = data;
        else activities.push(data);

        editingId = null;
        document.getElementById('activityForm').reset();
        questionsList.innerHTML = '';
        document.getElementById('saveActivityBtn').textContent = 'Guardar Actividad';
        formScreen.style.display = 'none';
        selectionScreen.style.display = 'block';
        renderActivitiesList();
    });

    document.getElementById('generateDocxBtn').addEventListener('click', async () => {
        if (!document.getElementById('activityForm').reportValidity()) return;
        const btn = document.getElementById('generateDocxBtn');
        const orig = btn.innerHTML;
        btn.innerHTML = `<span>Generando...</span>`; btn.disabled = true;
        try {
            const actData = getFormData();
            const blob = await buildDocxBlob(actData);

            // FILENAME FORMATING FIX: 
            // Chrome running on file:/// breaks the automatic download entirely (gives UUIDs) if it sees spaces or weird chars
            const safeTitle = actData.title.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_\-]/g, '') || 'Actividad';

            window.saveAs(blob, `${safeTitle}.docx`);
        } catch (e) {
            console.error(e); alert('Error al generar Docx.');
        } finally {
            btn.innerHTML = orig; btn.disabled = false;
        }
    });

    downloadAllBtn.addEventListener('click', async () => {
        if (activities.length === 0) return;
        const btn = downloadAllBtn;
        const orig = btn.innerHTML;
        btn.innerHTML = `Generando Word...`; btn.disabled = true;
        try {
            const blob = await buildDocxBlob(activities);
            window.saveAs(blob, `Todas_Las_Actividades.docx`);
        } catch (e) {
            console.error(e); alert('Error generando documento consolidado');
        } finally {
            btn.innerHTML = orig; btn.disabled = false;
        }
    });

    // Expects an array of activities.
    // If a single activity is passed, wrap it in an array to match logic.
    async function buildDocxBlob(actInput) {
        const acts = Array.isArray(actInput) ? actInput : [actInput];
        if (!window.docx) throw new Error('Docx library not loaded');
        // Destructure definitions from window.docx
        const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, Table, TableRow, TableCell, BorderStyle, WidthType } = window.docx;

        // Prepare children paragraphs for the document body
        const children = [];

        for (let i = 0; i < acts.length; i++) {
            const actData = acts[i];

            if (i > 0) {
                // Insert page break before the next activity
                children.push(new Paragraph({
                    pageBreakBefore: true,
                    children: [new TextRun("")]
                }));
            }

            // Retrieve info from object
            const title = actData.title;
            const consigna = actData.consigna;
            const objectives = actData.objectives;
            const guidelines = actData.guidelines;
            const criteria = actData.criteria;
            const isGroup = actData.isGroup;
            const type = actData.type;

            // Reusable section adder factory
            const addSection = (subtitle, textBody) => {
                // Section Title
                children.push(
                    new Paragraph({
                        heading: HeadingLevel.HEADING_3,
                        spacing: { before: 400, after: 150 },
                        children: [
                            new TextRun({
                                text: subtitle,
                                bold: true,
                                size: 28, // 14pt
                                color: "2D3B45",
                                font: "Arial"
                            }),
                        ],
                    })
                );

                // Section content processing line-breaks appropriately
                const lines = textBody.split('\n');
                lines.forEach(line => {
                    const trimmed = line.trim();
                    if (trimmed !== '') {
                        children.push(
                            new Paragraph({
                                spacing: { after: 120 },
                                children: [
                                    new TextRun({
                                        text: trimmed,
                                        size: 24, // 12pt
                                        font: "Arial",
                                        color: "333333"
                                    })
                                ]
                            })
                        );
                    }
                });
            };

            // --- 1. TITLE (Left Aligned and Bold) ---
            children.push(
                new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    spacing: { after: 400 },
                    children: [
                        new TextRun({
                            text: title,
                            bold: true,
                            size: 36, // 18pt (represented in half-points)
                            font: "Arial"
                        }),
                    ],
                })
            );

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
            if (type !== 'evaluacion') {
                addSection(config.subtitle2, criteria);
            }

            // --- 6. EVALUATION NEW QUIZ QUESTIONS PARSER ---
            if (type === 'evaluacion') {
                const tempDiv = document.createElement('div');
                if (actData.questionsHTML) tempDiv.innerHTML = actData.questionsHTML;
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
                        spacing: { before: 400, after: 100 },
                        children: [new TextRun({ text: "" })]
                    })
                );

                // Creates the Border Box using a single-cell Docx Table to enclose instructions visually like a card.
                children.push(
                    new Table({
                        width: {
                            size: 100,
                            type: WidthType.PERCENTAGE,
                        },
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
                                        margins: {
                                            top: 250,
                                            bottom: 250,
                                            left: 350,
                                            right: 300,
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
            title: acts.length === 1 ? acts[0].title : "Actividades Compiladas",
            description: "Actividades estandarizadas para Canvas",
            sections: [{ properties: {}, children: children }],
        });

        // Convert to BLOB and return
        return await Packer.toBlob(doc);
    }
});
