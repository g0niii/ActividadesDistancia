import React from 'react';
import { Trash2, Plus, GripVertical, ChevronDown, ChevronUp } from 'lucide-react';

const ITEM_TYPES = {
    'page': '📄 Página',
    'assignment': '📝 Tarea',
    'quiz': '🚀 Evaluación',
    'discussion': '💬 Foro',
    'file': '📎 Archivo',
    'link': '🔗 Enlace'
};

const PlanillaBuilder = ({ modules, setModules }) => {

    const addModule = () => {
        const newMod = {
            id: Date.now().toString() + Math.random(),
            title: '',
            items: [{ id: Date.now().toString(), type: 'page', title: '', comment: '' }]
        };
        setModules([...modules, newMod]);
    };

    const updateModule = (id, updates) => {
        setModules(modules.map(m => m.id === id ? { ...m, ...updates } : m));
    };

    const removeModule = (id) => {
        setModules(modules.filter(m => m.id !== id));
    };

    const addItem = (mId) => {
        const m = modules.find(curr => curr.id === mId);
        const newItem = { id: Date.now().toString(), type: 'page', title: '', comment: '' };
        updateModule(mId, { items: [...m.items, newItem] });
    };

    const updateItem = (mId, iId, updates) => {
        const m = modules.find(curr => curr.id === mId);
        const newItems = m.items.map(it => it.id === iId ? { ...it, ...updates } : it);
        updateModule(mId, { items: newItems });
    };

    const removeItem = (mId, iId) => {
        const m = modules.find(curr => curr.id === mId);
        updateModule(mId, { items: m.items.filter(it => it.id !== iId) });
    };

    const moveModule = (idx, dir) => {
        const newModules = [...modules];
        const targetIdx = idx + dir;
        if (targetIdx >= 0 && targetIdx < newModules.length) {
            [newModules[idx], newModules[targetIdx]] = [newModules[targetIdx], newModules[idx]];
            setModules(newModules);
        }
    };

    const moveItem = (mId, idx, dir) => {
        const m = modules.find(curr => curr.id === mId);
        const newItems = [...m.items];
        const targetIdx = idx + dir;
        if (targetIdx >= 0 && targetIdx < newItems.length) {
            [newItems[idx], newItems[targetIdx]] = [newItems[targetIdx], newItems[idx]];
            updateModule(mId, { items: newItems });
        }
    };

    return (
        <div className="planilla-builder">
            <div className="builder-header">
                <h3>Estructura de Módulos (Canvas)</h3>
                <button onClick={addModule} className="primary-btn-sm">+ Agregar Módulo</button>
            </div>

            <div className="modules-list">
                {modules.map((m, mIdx) => (
                    <div key={m.id} className="module-group">
                        <div className="module-header">
                            <span className="drag-handle"><GripVertical size={16} /></span>
                            <input
                                type="text"
                                value={m.title}
                                onChange={(e) => updateModule(m.id, { title: e.target.value })}
                                placeholder="Nombre del Módulo (ej: Módulo 1: Introducción)"
                                className="mod-title-input"
                            />
                            <div className="module-actions">
                                <button onClick={() => moveModule(mIdx, -1)} disabled={mIdx === 0}><ChevronUp size={16} /></button>
                                <button onClick={() => moveModule(mIdx, 1)} disabled={mIdx === modules.length - 1}><ChevronDown size={16} /></button>
                                <button onClick={() => removeModule(m.id)} className="btn-icon"><Trash2 size={16} /></button>
                            </div>
                        </div>

                        <div className="items-list">
                            {m.items.map((it, iIdx) => (
                                <div key={it.id} className="item-row">
                                    <span className="drag-handle-sm">⠿</span>
                                    <select
                                        value={it.type}
                                        onChange={(e) => updateItem(m.id, it.id, { type: e.target.value })}
                                        className="item-type-select"
                                    >
                                        {Object.entries(ITEM_TYPES).map(([val, label]) => (
                                            <option key={val} value={val}>{label}</option>
                                        ))}
                                    </select>
                                    <div className="item-content">
                                        <input
                                            type="text"
                                            value={it.title}
                                            onChange={(e) => updateItem(m.id, it.id, { title: e.target.value })}
                                            placeholder="Título del componente..."
                                            className="item-title-input"
                                        />
                                        <input
                                            type="text"
                                            value={it.comment}
                                            onChange={(e) => updateItem(m.id, it.id, { comment: e.target.value })}
                                            placeholder="Observaciones (ej: 'Utilizar video de YouTube')"
                                            className="item-comment-input"
                                        />
                                    </div>
                                    <div className="item-actions">
                                        <button onClick={() => moveItem(m.id, iIdx, -1)} disabled={iIdx === 0}><ChevronUp size={14} /></button>
                                        <button onClick={() => moveItem(m.id, iIdx, 1)} disabled={iIdx === m.items.length - 1}><ChevronDown size={14} /></button>
                                        <button onClick={() => removeItem(m.id, it.id)} className="btn-icon"><Trash2 size={14} /></button>
                                    </div>
                                </div>
                            ))}
                        </div>
                        <div className="module-footer">
                            <button onClick={() => addItem(m.id)} className="btn-small-outline">+ Agregar Ítem</button>
                        </div>
                    </div>
                ))}
            </div>
            {modules.length === 0 && (
                <div className="empty-state">No hay módulos creados aún. Haz clic en el botón superior para empezar.</div>
            )}
        </div>
    );
};

export default PlanillaBuilder;
