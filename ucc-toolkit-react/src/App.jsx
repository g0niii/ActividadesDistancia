import React, { useState, useEffect } from 'react';
import ReactQuill from 'react-quill';
import 'react-quill/dist/quill.snow.css';
import {
  FileText, MessageSquare, ClipboardCheck, Layout, ChevronLeft,
  Send, Save, Trash2, LogOut, X, GripVertical, AlertCircle, Eye, Download, FileSpreadsheet, Plus
} from 'lucide-react';
import { saveAs } from 'file-saver';
import { downloadExcel } from './utils/exportUtils';
import { buildDocxBlob } from './utils/docxUtils';
import { GoogleDriveService } from './services/GoogleDriveService';
import QuestionBuilder from './components/QuestionBuilder';
import PlanillaBuilder from './components/PlanillaBuilder';

// --- CONFIGURACIÓN ---
const GOOGLE_CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID;
const ALLOWED_DOMAIN = 'ucc.edu.ar';

const TYPE_CONFIGS = {
  'foro': {
    title: 'Generar Foro',
    subtitle1: 'Pautas de participación',
    subtitle2: 'Criterios de evaluación',
    icon: <MessageSquare />,
    color: 'card-foro'
  },
  'tarea': {
    title: 'Generar Tarea',
    subtitle1: 'Pautas de presentación',
    subtitle2: 'Criterios de evaluación',
    icon: <FileText />,
    color: 'card-tarea'
  },
  'evaluacion': {
    title: 'Generar Evaluación',
    subtitle1: 'Instrucciones de la evaluación',
    subtitle2: 'Configuración de la evaluación',
    icon: <ClipboardCheck />,
    color: 'card-evaluacion'
  },
  'planilla': {
    title: 'Planilla de Montaje (Mapa del Curso)',
    subtitle1: 'Notas del Asesor Pedagógico',
    subtitle2: 'Observaciones para Maquetación',
    icon: <Layout />,
    color: 'card-planilla'
  }
};

const CATEGORIES = {
  'tarea': [
    { value: 'Actividad obligatoria', label: 'Obligatoria' },
    { value: 'Actividad sugerida', label: 'Sugerida' }
  ],
  'foro': [
    { value: 'Foro obligatorio', label: 'Foro obligatorio' },
    { value: 'Foro sugerido', label: 'Foro sugerido' }
  ],
  'evaluacion': [
    { value: 'Evaluación', label: 'Evaluación' },
    { value: 'Autoevaluación', label: 'Autoevaluación' }
  ],
  'planilla': [
    { value: 'Montaje', label: 'Planilla de Montaje' }
  ]
};

function App() {
  const [user, setUser] = useState(null);
  const [view, setView] = useState('selection');
  const [resources, setResources] = useState([]);
  const [currentResource, setCurrentResource] = useState(null);
  const [accessToken, setAccessToken] = useState(null);

  // 1. Inicialización de Google Identity Services
  useEffect(() => {
    const handleCredentialResponse = (response) => {
      const payload = JSON.parse(atob(response.credential.split('.')[1]));
      const email = payload.email;

      if (!email.endsWith(`@${ALLOWED_DOMAIN}`) && email !== 'tecnologia.sied@ucc.edu.ar') {
        alert(`Acceso denegado. Solo se permiten cuentas de ${ALLOWED_DOMAIN}`);
        return;
      }

      const userData = {
        name: payload.name, // El nombre que viene de Google
        email: payload.email,
        picture: payload.picture
      };

      setUser(userData);
      localStorage.setItem('ucc_user', JSON.stringify(userData));

      // Intentamos obtener el token para Drive inmediatamente después del login
      getTokenForDrive();
    };

    if (window.google) {
      window.google.accounts.id.initialize({
        client_id: GOOGLE_CLIENT_ID,
        callback: handleCredentialResponse,
        hosted_domain: ALLOWED_DOMAIN
      });
    }
  }, []);

  const getTokenForDrive = (manualTrigger = false) => {
    if (!window.google) return;

    // Si ya tenemos un token, podríamos intentar usarlo, pero requestAccessToken forzará el popup si es necesario.
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_CLIENT_ID,
      scope: 'https://www.googleapis.com/auth/drive.file',
      callback: (response) => {
        if (response.access_token) {
          console.log("Token de Google Drive obtenido correctamente");
          setAccessToken(response.access_token);
          localStorage.setItem('drive_token', response.access_token);
          if (manualTrigger) alert("✅ Google Drive vinculado correctamente.");
        }
        if (response.error) {
          console.error("Error al obtener token:", response.error);
        }
      },
    });

    // requestAccessToken DEBE ser llamado tras una acción del usuario si es manualTrigger para evitar bloqueo de popups
    client.requestAccessToken({ prompt: manualTrigger ? 'consent' : '' });
  };

  useEffect(() => {
    const saved = localStorage.getItem('ucc_resources');
    if (saved) setResources(JSON.parse(saved));
    const savedUser = localStorage.getItem('ucc_user');
    if (savedUser) setUser(JSON.parse(savedUser));
    const savedToken = localStorage.getItem('drive_token');
    if (savedToken) setAccessToken(savedToken);
  }, []);

  useEffect(() => {
    localStorage.setItem('ucc_resources', JSON.stringify(resources));
  }, [resources]);

  const handleLogout = () => {
    setUser(null);
    setAccessToken(null);
    localStorage.removeItem('drive_token');
    localStorage.removeItem('ucc_user');
    setView('selection');
  };

  const updateTitlePrefix = (res) => {
    if (res.type === 'planilla') return 'Planilla de Montaje';
    const cat = res.category;
    const mod = res.module;
    if (!cat || !mod) return res.title;
    const prefix = `${cat} M${mod} - `;
    const oldVal = res.title || '';
    const regex = /^(Actividad obligatoria|Actividad sugerida|Foro obligatorio|Foro sugerido|Evaluación|Autoevaluación|Montaje|Referencia bibliográfica)\s+M\d+\s+-\s*/i;
    if (regex.test(oldVal)) return oldVal.replace(regex, prefix);
    else if (oldVal === '') return prefix;
    else return prefix + oldVal;
  };

  const startNewResource = (type) => {
    const newRes = {
      id: Date.now().toString(),
      type,
      title: '',
      category: CATEGORIES[type][0].value,
      module: type === 'planilla' ? '0' : '1',
      consigna: '', objectives: '', guidelines: '', criteria: '',
      isGroup: false, subjectName: '', career: '', contentAuthor: '', driveLink: '',
      questions: [],
      mountingModules: type === 'planilla' ? [{ id: Date.now().toString(), title: '', items: [] }] : []
    };
    newRes.title = updateTitlePrefix(newRes);
    setCurrentResource(newRes);
    setView('form');
  };

  const saveResource = () => {
    const exists = resources.find(r => r.id === currentResource.id);
    if (exists) setResources(resources.map(r => r.id === currentResource.id ? currentResource : r));
    else setResources([...resources, currentResource]);
    setView('list');
  };

  const exportToDrive = async (resource) => {
    if (!accessToken) {
      if (confirm("Para guardar en Drive primero debes vincular tu cuenta. ¿Deseas vincularla ahora?")) {
        getTokenForDrive(true);
      }
      return;
    }

    try {
      // Indicador de "subiendo..." opcionalmente, pero por ahora el alert final basta
      const blob = await buildDocxBlob(resource);
      await GoogleDriveService.uploadFile(accessToken, blob, `${resource.title}.docx`);
      alert("✅ ¡Archivo guardado con éxito en tu Google Drive!");
    } catch (err) {
      console.error(err);
      if (err.message?.includes('401') || err.message?.includes('token')) {
        alert("Tu sesión de Drive ha expirado. Por favor, vuelve a conectar.");
        getTokenForDrive(true);
      } else {
        alert("Hubo un problema al subir el archivo: " + err.message);
      }
    }
  };

  useEffect(() => {
    if (!user && view === 'selection' && window.google) {
      window.google.accounts.id.renderButton(
        document.getElementById("googleBtnContainer"),
        { theme: "outline", size: "large", width: "250" }
      );
    }
  }, [user, view]);

  if (!user) {
    return (
      <div className="auth-container">
        <img src="/logo-sied.png" width="180" style={{ marginBottom: '20px' }} alt="Logo SIED" />
        <h1 style={{ color: 'var(--primary-color)', marginBottom: '10px' }}>Acceso Toolkit UCC</h1>
        <p style={{ color: '#666', marginBottom: '30px' }}>Iniciá sesión con tu cuenta <strong>@ucc.edu.ar</strong> para continuar.</p>
        <div id="googleBtnContainer" style={{ display: 'flex', justifyContent: 'center' }}></div>
      </div>
    );
  }

  return (
    <div className="App">
      <div className="container">
        <header>
          <div className="header-left">
            <img src="/logo-sied.png" className="sied-logo" alt="SIED Logo" />
            <div>
              <h1 style={{ fontSize: '1.2rem' }}>Estandarizador UCC</h1>
              {user && (
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginTop: '4px' }}>
                  {user.picture && <img src={user.picture} style={{ width: '18px', borderRadius: '50%' }} alt="Avatar" />}
                  <span style={{ fontSize: '11px', opacity: 0.9 }}>Hola, {user.name || 'Docente'}</span>
                  <div style={{
                    width: '8px',
                    height: '8px',
                    borderRadius: '50%',
                    backgroundColor: accessToken ? '#28a745' : '#ffc107',
                    marginLeft: '4px'
                  }} title={accessToken ? "Google Drive Vinculado" : "Drive no vinculado"}></div>
                </div>
              )}
            </div>
          </div>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <button onClick={() => setView('selection')} className="btn-small">Inicio</button>
            <button onClick={() => setView('list')} className="btn-small">Recursos ({resources.length})</button>
            {!accessToken && <button onClick={() => getTokenForDrive(true)} className="btn-small" style={{ background: '#fff9e6', borderColor: '#ffeebf' }}>Conectar Drive</button>}
            <button onClick={handleLogout} className="btn-small" style={{ color: 'red', border: '1px solid #ffcfcf' }}><LogOut size={14} /></button>
          </div>
        </header>

        <main style={{ padding: '2rem' }}>
          {view === 'selection' && (
            <div className="premium-cards-grid">
              {Object.entries(TYPE_CONFIGS).map(([key, config]) => (
                <div key={key} className={`premium-type-card ${config.color}`} onClick={() => startNewResource(key)}>
                  <div className="premium-card-icon">{config.icon}</div>
                  <h3>{config.title}</h3>
                </div>
              ))}
            </div>
          )}

          {view === 'form' && currentResource && (
            <div className="resource-form">
              <button className="back-btn" onClick={() => setView('selection')}><ChevronLeft size={18} /> Volver</button>

              <div className="resource-form-header">
                <h2 style={{ color: 'var(--primary-color)' }}>{TYPE_CONFIGS[currentResource.type].title}</h2>
                <p style={{ color: '#666', fontSize: '0.9rem' }}>Completa los campos para generar tu documento estandarizado.</p>
              </div>

              {currentResource.type === 'planilla' ? (
                <div className="planilla-content">
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1.5rem', marginBottom: '2rem' }}>
                    <div className="form-group"><label>Asignatura</label><input type="text" value={currentResource.subjectName} onChange={e => setCurrentResource({ ...currentResource, subjectName: e.target.value })} placeholder="Ej: Programación I" /></div>
                    <div className="form-group"><label>Carrera</label><input type="text" value={currentResource.career} onChange={e => setCurrentResource({ ...currentResource, career: e.target.value })} placeholder="Ej: Lic. en Informática" /></div>
                    <div className="form-group"><label>Docente Contenidista</label><input type="text" value={currentResource.contentAuthor} onChange={e => setCurrentResource({ ...currentResource, contentAuthor: e.target.value })} placeholder="Nombre completo" /></div>
                    <div className="form-group"><label>Link del Drive</label><input type="text" value={currentResource.driveLink} onChange={e => setCurrentResource({ ...currentResource, driveLink: e.target.value })} placeholder="https://drive.google.com/..." /></div>
                  </div>

                  <PlanillaBuilder
                    modules={currentResource.mountingModules}
                    setModules={(mods) => setCurrentResource({ ...currentResource, mountingModules: mods })}
                  />
                </div>
              ) : (
                <>
                  <div style={{ display: 'grid', gridTemplateColumns: '1.5fr 0.5fr', gap: '1.5rem', marginBottom: '1.5rem' }}>
                    <div className="form-group">
                      <label>Categoría</label>
                      <select value={currentResource.category} onChange={e => {
                        const n = { ...currentResource, category: e.target.value };
                        n.title = updateTitlePrefix(n);
                        setCurrentResource(n);
                      }}>
                        {CATEGORIES[currentResource.type].map(c => <option key={c.value} value={c.value}>{c.label}</option>)}
                      </select>
                    </div>
                    <div className="form-group">
                      <label>Módulo</label>
                      <input type="number" value={currentResource.module} onChange={e => {
                        const n = { ...currentResource, module: e.target.value };
                        n.title = updateTitlePrefix(n);
                        setCurrentResource(n);
                      }} />
                    </div>
                  </div>

                  <div className="form-group">
                    <label>Título Final</label>
                    <input type="text" value={currentResource.title} onChange={e => setCurrentResource({ ...currentResource, title: e.target.value })} />
                  </div>

                  <div className="form-group">
                    <label>Consigna / Introducción</label>
                    <ReactQuill theme="snow" value={currentResource.consigna} onChange={v => setCurrentResource({ ...currentResource, consigna: v })} />
                  </div>

                  <div className="form-group">
                    <label>Objetivos de Aprendizaje</label>
                    <ReactQuill theme="snow" value={currentResource.objectives} onChange={v => setCurrentResource({ ...currentResource, objectives: v })} />
                  </div>

                  {currentResource.type !== 'evaluacion' && (
                    <>
                      <div className="form-group">
                        <label>{TYPE_CONFIGS[currentResource.type].subtitle1}</label>
                        <ReactQuill theme="snow" value={currentResource.guidelines} onChange={v => setCurrentResource({ ...currentResource, guidelines: v })} />
                      </div>
                      <div className="form-group">
                        <label>{TYPE_CONFIGS[currentResource.type].subtitle2}</label>
                        <ReactQuill theme="snow" value={currentResource.criteria} onChange={v => setCurrentResource({ ...currentResource, criteria: v })} />
                      </div>

                      {currentResource.type === 'tarea' && (
                        <div className="group-logic-section" style={{ marginTop: '2rem' }}>
                          <label className="checkbox-container">
                            <input type="checkbox" checked={currentResource.isGroup} onChange={e => setCurrentResource({ ...currentResource, isGroup: e.target.checked })} />
                            <span className="checkbox-text">¿Es una actividad grupal?</span>
                          </label>

                          {currentResource.isGroup && (
                            <div className="group-preview">
                              <div className="preview-icon">
                                <GripVertical color="white" />
                              </div>
                              <div className="preview-content">
                                <h4>¿Cómo unirte a tu grupo para esta actividad?</h4>
                                <p style={{ fontSize: '0.9rem', marginBottom: '0.5rem' }}>Esta es una <strong>actividad grupal</strong>. Pasos para los alumnos:</p>
                                <ul style={{ paddingLeft: '20px' }}>
                                  <li>Ir a "Personas" &gt; "Grupos".</li>
                                  <li>Elegir un grupo y clic en "Unirse".</li>
                                  <li>Sólo una persona del grupo debe entregar.</li>
                                </ul>
                                <p className="warning-text">Se incluirá automáticamente el bloque de instrucciones en el documento.</p>
                              </div>
                            </div>
                          )}
                        </div>
                      )}
                    </>
                  )}

                  {currentResource.type === 'evaluacion' && (
                    <div style={{ marginTop: '2rem' }}>
                      <QuestionBuilder
                        questions={currentResource.questions}
                        setQuestions={(qs) => setCurrentResource({ ...currentResource, questions: qs })}
                      />
                    </div>
                  )}
                </>
              )}

              <div className="form-actions" style={{ marginTop: '3rem', display: 'flex', gap: '15px', borderTop: '1px solid #eee', paddingTop: '2rem' }}>
                <button className="primary-btn" style={{ flex: 2 }} onClick={saveResource}><Save size={20} /> Guardar Recurso</button>
                <button className="secondary-btn" onClick={() => buildDocxBlob(currentResource).then(blob => saveAs(blob, `${currentResource.title || 'Recurso'}.docx`))}><Download size={18} /> Word</button>
                <button className="secondary-btn" onClick={() => downloadExcel(currentResource)}><FileSpreadsheet size={18} /> Excel</button>
                <button className="secondary-btn" style={{ background: '#e7f3ff' }} onClick={() => exportToDrive(currentResource)}><Send size={18} /> Drive</button>
              </div>
            </div>
          )}

          {view === 'list' && (
            <div className="list-view">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem' }}>
                <h2>Mis Recursos Guardados</h2>
                <button className="primary-btn" onClick={() => setView('selection')}><Plus size={18} /> Nuevo</button>
              </div>

              {resources.length === 0 ? (
                <div style={{ textAlign: 'center', padding: '3rem', color: '#888', background: '#f9f9f9', borderRadius: '12px' }}>
                  No tienes recursos guardados todavía.
                </div>
              ) : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
                  {resources.map(res => (
                    <div key={res.id} className="act-item" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <div>
                        <span className="act-type" style={{ color: `var(--accent, ${TYPE_CONFIGS[res.type].color})` }}>{res.type}</span>
                        <h4 style={{ margin: '4px 0' }}>{res.title}</h4>
                      </div>
                      <div style={{ display: 'flex', gap: '8px' }}>
                        <button className="btn-small" onClick={() => { setCurrentResource(res); setView('form'); }}>Editar</button>
                        <button className="btn-small" onClick={() => buildDocxBlob(res).then(blob => saveAs(blob, `${res.title}.docx`))}>Docx</button>
                        <button className="btn-small" onClick={() => exportToDrive(res)}>Drive</button>
                        <button className="btn-small" style={{ color: 'red', borderColor: '#ffcfcf' }} onClick={() => setResources(resources.filter(r => r.id !== res.id))}><Trash2 size={14} /></button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </main>
      </div>
    </div>
  );
}

export default App;
