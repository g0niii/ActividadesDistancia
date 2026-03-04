import React from 'react';
import { Trash2, Plus, GripVertical, Info } from 'lucide-react';

const QUESTION_LABELS = {
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
    'bloque-texto': 'Bloque de texto',
    'estimulo': 'Estímulo'
};

const QuestionBuilder = ({ questions, setQuestions, isNested = false }) => {

    const addQuestion = (type) => {
        const newQ = {
            id: Date.now().toString() + Math.random(),
            type,
            prompt: '',
            options: type === 'opcion-multiple' || type === 'respuestas-multiples' ? [
                { id: '1', text: '', correct: true },
                { id: '2', text: '', correct: false }
            ] : [],
            tfAnswer: 'Verdadero',
            fibAnswers: '',
            fibType: 'Entrada abierta',
            ddText: '',
            ddOptions: '',
            pairs: type === 'coincidencia' ? [
                { id: '1', premise: '', match: '' },
                { id: '2', premise: '', match: '' }
            ] : [],
            distractors: [],
            numType: 'exacta',
            numValue: '',
            formulaText: '',
            formulaVars: '',
            essayLimit: '',
            essayRich: true,
            essaySpellcheck: true,
            fileTypes: '',
            fileLimit: 1,
            orderItems: type === 'ordenamiento' ? [
                { id: '1', text: '' },
                { id: '2', text: '' }
            ] : [],
            orderTop: '',
            orderBot: '',
            categories: type === 'categorizacion' ? [
                { id: '1', desc: '', answers: [{ id: '1a', text: '' }] }
            ] : [],
            hotspotImg: '',
            hotspotZone: '',
            nestedQuestions: []
        };
        setQuestions([...questions, newQ]);
    };

    const updateQuestion = (id, updates) => {
        setQuestions(questions.map(q => q.id === id ? { ...q, ...updates } : q));
    };

    const removeQuestion = (id) => {
        setQuestions(questions.filter(q => q.id !== id));
    };

    const addOption = (qId) => {
        const q = questions.find(curr => curr.id === qId);
        const newOpt = { id: Date.now().toString(), text: '', correct: false };
        updateQuestion(qId, { options: [...q.options, newOpt] });
    };

    const updateOption = (qId, optId, updates) => {
        const q = questions.find(curr => curr.id === qId);
        const newOptions = q.options.map(opt => {
            if (opt.id === optId) {
                return { ...opt, ...updates };
            }
            // For radio-like behavior in multiple choice
            if (q.type === 'opcion-multiple' && updates.correct === true) {
                return { ...opt, correct: false };
            }
            return opt;
        });
        updateQuestion(qId, { options: newOptions });
    };

    const removeOption = (qId, optId) => {
        const q = questions.find(curr => curr.id === qId);
        updateQuestion(qId, { options: q.options.filter(opt => opt.id !== optId) });
    };

    return (
        <div className={`question-builder ${isNested ? 'nested' : ''}`}>
            <div className="questions-list">
                {questions.map((q, idx) => (
                    <div key={q.id} className="question-block">
                        <div className="question-header">
                            <span className="q-number">#{idx + 1} - {QUESTION_LABELS[q.type]}</span>
                            <button onClick={() => removeQuestion(q.id)} className="remove-q-btn"><Trash2 size={16} /></button>
                        </div>

                        <div className="form-group">
                            <label>{q.type === 'estimulo' ? 'Contenido del estímulo (Texto/Historia)' : 'Raíz de la pregunta'}</label>
                            <textarea
                                value={q.prompt}
                                onChange={(e) => updateQuestion(q.id, { prompt: e.target.value })}
                                rows={2}
                                placeholder="Escribe aquí el enunciado..."
                            />
                        </div>

                        <div className="question-specifics">
                            {/* Opción Múltiple & Respuesta Múltiple */}
                            {(q.type === 'opcion-multiple' || q.type === 'respuestas-multiples') && (
                                <div className="optionsv-container">
                                    <label>Respuestas posibles</label>
                                    {q.options.map((opt) => (
                                        <div key={opt.id} className="option-row">
                                            <input
                                                type={q.type === 'opcion-multiple' ? 'radio' : 'checkbox'}
                                                checked={opt.correct}
                                                onChange={(e) => updateOption(q.id, opt.id, { correct: e.target.checked })}
                                                name={`q-${q.id}-correct`}
                                            />
                                            <input
                                                type="text"
                                                value={opt.text}
                                                onChange={(e) => updateOption(q.id, opt.id, { text: e.target.value })}
                                                placeholder="Respuesta..."
                                                className="flex-1"
                                            />
                                            <button onClick={() => removeOption(q.id, opt.id)} className="btn-icon"><Trash2 size={14} /></button>
                                        </div>
                                    ))}
                                    <button onClick={() => addOption(q.id)} className="btn-small-outline"><Plus size={14} /> Respuesta</button>
                                </div>
                            )}

                            {/* Verdadero o Falso */}
                            {q.type === 'verdadero-falso' && (
                                <div className="tf-container">
                                    <label><input type="radio" checked={q.tfAnswer === 'Verdadero'} onChange={() => updateQuestion(q.id, { tfAnswer: 'Verdadero' })} /> Verdadero</label>
                                    <label><input type="radio" checked={q.tfAnswer === 'Falso'} onChange={() => updateQuestion(q.id, { tfAnswer: 'Falso' })} /> Falso</label>
                                </div>
                            )}

                            {/* Rellenar espacio */}
                            {q.type === 'completar-espacio' && (
                                <>
                                    <p className="hint">Usa `acentos graves` para marcar los espacios. Ej: "El cielo es `azul`".</p>
                                    <div className="form-group">
                                        <label>Respuestas aceptadas (separadas por coma)</label>
                                        <input type="text" value={q.fibAnswers} onChange={(e) => updateQuestion(q.id, { fibAnswers: e.target.value })} placeholder="Ej: azul, celeste" />
                                    </div>
                                </>
                            )}

                            {/* Coincidencia */}
                            {q.type === 'coincidencia' && (
                                <div className="pairs-container">
                                    <label>Pares de pregunta/respuesta</label>
                                    {q.pairs.map((p, pIdx) => (
                                        <div key={p.id} className="option-row">
                                            <input type="text" value={p.premise} onChange={(e) => {
                                                const newPairs = [...q.pairs];
                                                newPairs[pIdx].premise = e.target.value;
                                                updateQuestion(q.id, { pairs: newPairs });
                                            }} placeholder="Criterio..." className="flex-1" />
                                            <input type="text" value={p.match} onChange={(e) => {
                                                const newPairs = [...q.pairs];
                                                newPairs[pIdx].match = e.target.value;
                                                updateQuestion(q.id, { pairs: newPairs });
                                            }} placeholder="Match..." className="flex-1" />
                                            <button onClick={() => {
                                                updateQuestion(q.id, { pairs: q.pairs.filter(pair => pair.id !== p.id) });
                                            }} className="btn-icon"><Trash2 size={14} /></button>
                                        </div>
                                    ))}
                                    <button onClick={() => {
                                        updateQuestion(q.id, { pairs: [...q.pairs, { id: Date.now().toString(), premise: '', match: '' }] });
                                    }} className="btn-small-outline"><Plus size={14} /> Par</button>
                                </div>
                            )}

                            {/* Categorización */}
                            {q.type === 'categorizacion' && (
                                <div className="categories-container">
                                    <label>Categorías y sus respuestas</label>
                                    {q.categories.map((cat, catIdx) => (
                                        <div key={cat.id} className="category-block" style={{ border: '1px solid #eee', padding: '10px', borderRadius: '8px', marginBottom: '10px' }}>
                                            <div style={{ display: 'flex', gap: '8px', marginBottom: '8px' }}>
                                                <input type="text" value={cat.desc} onChange={(e) => {
                                                    const newCats = [...q.categories];
                                                    newCats[catIdx].desc = e.target.value;
                                                    updateQuestion(q.id, { categories: newCats });
                                                }} placeholder="Nombre de Categoría..." className="flex-1 font-bold" />
                                                <button onClick={() => {
                                                    updateQuestion(q.id, { categories: q.categories.filter(c => c.id !== cat.id) });
                                                }} className="btn-icon"><Trash2 size={14} /></button>
                                            </div>
                                            <div className="cat-ans-list" style={{ paddingLeft: '20px' }}>
                                                {cat.answers.map((ans, ansIdx) => (
                                                    <div key={ans.id} className="option-row">
                                                        <input type="text" value={ans.text} onChange={(e) => {
                                                            const newCats = [...q.categories];
                                                            newCats[catIdx].answers[ansIdx].text = e.target.value;
                                                            updateQuestion(q.id, { categories: newCats });
                                                        }} placeholder="Respuesta..." className="flex-1" />
                                                        <button onClick={() => {
                                                            const newCats = [...q.categories];
                                                            newCats[catIdx].answers = cat.answers.filter(a => a.id !== ans.id);
                                                            updateQuestion(q.id, { categories: newCats });
                                                        }} className="btn-icon"><Trash2 size={12} /></button>
                                                    </div>
                                                ))}
                                                <button onClick={() => {
                                                    const newCats = [...q.categories];
                                                    newCats[catIdx].answers.push({ id: Date.now().toString(), text: '' });
                                                    updateQuestion(q.id, { categories: newCats });
                                                }} className="btn-small-outline" style={{ fontSize: '0.75rem', padding: '4px 8px' }}>+ Respuesta</button>
                                            </div>
                                        </div>
                                    ))}
                                    <button onClick={() => {
                                        updateQuestion(q.id, { categories: [...q.categories, { id: Date.now().toString(), desc: '', answers: [{ id: 'a1', text: '' }] }] });
                                    }} className="btn-small-outline"><Plus size={14} /> Categoría</button>
                                </div>
                            )}

                            {/* Estímulo (Nested Questions) */}
                            {q.type === 'estimulo' && (
                                <div className="nested-builder">
                                    <QuestionBuilder
                                        questions={q.nestedQuestions}
                                        setQuestions={(nested) => updateQuestion(q.id, { nestedQuestions: nested })}
                                        isNested={true}
                                    />
                                </div>
                            )}

                            {/* Bloque de texto */}
                            {q.type === 'bloque-texto' && (
                                <div className="info-box"><Info size={16} /> Este bloque es solo informativo y no requiere respuesta.</div>
                            )}

                            {/* (Se pueden agregar el resto de tipos según necesidad siguiendo el patrón) */}
                        </div>
                    </div>
                ))}
            </div>

            <div className={`add-q-footer ${isNested ? 'nested-footer' : ''}`}>
                <select id={isNested ? undefined : "newQuestionType"} className="q-select-type">
                    {Object.entries(QUESTION_LABELS).map(([val, label]) => (
                        <option key={val} value={val}>{label}</option>
                    ))}
                </select>
                <button
                    onClick={() => {
                        const sel = document.getElementById(isNested ? undefined : "newQuestionType") || { value: 'opcion-multiple' };
                        addQuestion(sel.value || 'opcion-multiple');
                    }}
                    className="secondary-btn"
                >
                    <Plus size={18} /> Agregar Pregunta
                </button>
            </div>
        </div>
    );
};

export default QuestionBuilder;
