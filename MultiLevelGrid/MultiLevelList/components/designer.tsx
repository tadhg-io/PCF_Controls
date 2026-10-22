import * as React from 'react';
import { DefaultButton, PrimaryButton, TextField, Dropdown, MessageBar,
  MessageBarType } from 'office-ui-fabric-react';
import '../styles/designer.css';

export interface Survey { id: string; name: string; surveyTitle: string; }
export interface Section {
  id: string; surveyId: string; name: string; internalDescription: string;
  title: string; instructions: string; sortOrder: number;
}
export type DataType = 'text' | 'multiline' | 'number' | 'boolean' | 'date';
export interface Question {
  id: string; sectionId: string; name: string; internalDescription: string;
  dataType: DataType; title: string; instructions: string; sortOrder: number;
}
export interface SurveyTables { surveys: Survey[]; sections: Section[]; questions: Question[]; }

// Normalized JSON tables: IDs/FKs mirror the eventual Dataverse relationships.
export const SAMPLE_DATA: SurveyTables = {
  surveys: [
    { id: 'survey-fire', name: 'fire_risk_v1', surveyTitle: 'Fire risk assessment' },
    { id: 'survey-safety', name: 'workplace_safety_v1', surveyTitle: 'Workplace safety assessment' },
    { id: 'survey-habitat', name: 'habitat_v1', surveyTitle: 'Habitat survey' }
  ],
  sections: [
    { id: 'fire-property', surveyId: 'survey-fire', name: 'property_details', internalDescription: 'Identify the property and assessor.', title: 'Property details', instructions: 'Record the asset being assessed.', sortOrder: 0 },
    { id: 'fire-escape', surveyId: 'survey-fire', name: 'escape_routes', internalDescription: 'Review exits and evacuation routes.', title: 'Means of escape', instructions: 'Inspect each designated escape route.', sortOrder: 1 },
    { id: 'fire-equipment', surveyId: 'survey-fire', name: 'fire_equipment', internalDescription: 'Check protection equipment.', title: 'Fire protection', instructions: 'Check equipment and maintenance records.', sortOrder: 2 },
    { id: 'safety-site', surveyId: 'survey-safety', name: 'site_details', internalDescription: 'Site identification.', title: 'Site details', instructions: 'Identify the site and inspection date.', sortOrder: 0 },
    { id: 'safety-hazards', surveyId: 'survey-safety', name: 'site_hazards', internalDescription: 'Record observed hazards.', title: 'Hazards', instructions: 'Inspect the work area.', sortOrder: 1 },
    { id: 'habitat-site', surveyId: 'survey-habitat', name: 'habitat_site', internalDescription: 'Survey area metadata.', title: 'Survey area', instructions: 'Describe the area surveyed.', sortOrder: 0 },
    { id: 'habitat-species', surveyId: 'survey-habitat', name: 'species_observations', internalDescription: 'Species observed on site.', title: 'Species observations', instructions: 'Record observations made during this visit.', sortOrder: 1 }
  ],
  questions: [
    { id: 'q1', sectionId: 'fire-property', name: 'property_reference', internalDescription: 'Asset reference used internally.', dataType: 'text', title: 'Property reference', instructions: 'Enter the asset reference.', sortOrder: 0 },
    { id: 'q2', sectionId: 'fire-property', name: 'assessment_date', internalDescription: 'Date of inspection.', dataType: 'date', title: 'Assessment date', instructions: 'Select the inspection date.', sortOrder: 1 },
    { id: 'q3', sectionId: 'fire-escape', name: 'exits_clear', internalDescription: 'Check all final exits.', dataType: 'boolean', title: 'Are all emergency exits clear?', instructions: 'Check for obstructions.', sortOrder: 0 },
    { id: 'q4', sectionId: 'fire-escape', name: 'escape_notes', internalDescription: 'Record any issues.', dataType: 'multiline', title: 'Escape route observations', instructions: 'Describe any issues found.', sortOrder: 1 },
    { id: 'q5', sectionId: 'fire-equipment', name: 'extinguisher_count', internalDescription: 'Total available extinguishers.', dataType: 'number', title: 'Number of extinguishers', instructions: 'Count accessible units.', sortOrder: 0 },
    { id: 'q6', sectionId: 'fire-equipment', name: 'alarm_tested', internalDescription: 'Confirm alarm test.', dataType: 'boolean', title: 'Has the fire alarm been tested?', instructions: 'Review the test log.', sortOrder: 1 },
    { id: 'q7', sectionId: 'safety-site', name: 'site_name', internalDescription: 'Site name.', dataType: 'text', title: 'Site name', instructions: '', sortOrder: 0 },
    { id: 'q8', sectionId: 'safety-site', name: 'inspection_date', internalDescription: 'Visit date.', dataType: 'date', title: 'Inspection date', instructions: '', sortOrder: 1 },
    { id: 'q9', sectionId: 'safety-hazards', name: 'trip_hazards', internalDescription: 'Trip hazard check.', dataType: 'boolean', title: 'Are trip hazards present?', instructions: '', sortOrder: 0 },
    { id: 'q10', sectionId: 'safety-hazards', name: 'hazard_notes', internalDescription: 'Details for follow-up.', dataType: 'multiline', title: 'Describe the hazards', instructions: '', sortOrder: 1 },
    { id: 'q11', sectionId: 'habitat-site', name: 'site_reference', internalDescription: 'Survey site reference.', dataType: 'text', title: 'Site reference', instructions: '', sortOrder: 0 },
    { id: 'q12', sectionId: 'habitat-site', name: 'area_hectares', internalDescription: 'Estimated area.', dataType: 'number', title: 'Area in hectares', instructions: '', sortOrder: 1 },
    { id: 'q13', sectionId: 'habitat-species', name: 'species_name', internalDescription: 'Species identification.', dataType: 'text', title: 'Species name', instructions: '', sortOrder: 0 },
    { id: 'q14', sectionId: 'habitat-species', name: 'observation_notes', internalDescription: 'Field notes.', dataType: 'multiline', title: 'Observation notes', instructions: '', sortOrder: 1 }
  ]
};

interface Selection {
  kind: 'section' | 'question';
  id: string;
}
interface DragItem {
  kind: 'section' | 'question';
  id: string;
}
const byOrder = <T extends { sortOrder: number }>(rows: T[]): T[] =>
  rows.slice().sort((a, b) => a.sortOrder - b.sortOrder);
const DATA_TYPES = [
  { key: 'text', text: 'Short text' }, { key: 'multiline', text: 'Long text' },
  { key: 'number', text: 'Number' }, { key: 'boolean', text: 'Yes / No' }, { key: 'date', text: 'Date' }
];
const fieldStyles = { fieldGroup: { borderRadius: 7 }, root: { minWidth: 0 } };

export default function SurveyDesigner(): React.ReactElement {
  const [tables, setTables] = React.useState<SurveyTables>(() => JSON.parse(JSON.stringify(SAMPLE_DATA)) as SurveyTables);
  const [surveyId, setSurveyId] = React.useState<string | null>(null);
  const [selected, setSelected] = React.useState<Selection | null>(null);
  const [expanded, setExpanded] = React.useState<Record<string, boolean>>({});
  const [notice, setNotice] = React.useState('');
  const drag = React.useRef<DragItem | null>(null);
  const idCounter = React.useRef(0);
  const newId = (prefix: string): string => `${prefix}-${Date.now()}-${++idCounter.current}`;
  const survey = tables.surveys.find(x => x.id === surveyId);
  const sections = byOrder(tables.sections.filter(x => x.surveyId === surveyId));
  const questionsFor = (id: string): Question[] => byOrder(tables.questions.filter(x => x.sectionId === id));
  const section = selected?.kind === 'section' ? sections.find(x => x.id === selected.id) : undefined;
  const question = selected?.kind === 'question' ? tables.questions.find(x => x.id === selected.id) : undefined;
  const record = section || question;
  const sectionIds = sections.map(x => x.id);
  const totalQuestions = tables.questions.filter(x => sectionIds.indexOf(x.sectionId) >= 0).length;

  function openSurvey(id: string): void {
    setSurveyId(id);
    const first = byOrder(tables.sections.filter(x => x.surveyId === id))[0];
    setSelected(first ? { kind: 'section', id: first.id } : null);
    setExpanded(first ? { [first.id]: true } : {});
    setNotice('');
  }
  function patchSection(id: string, patch: Partial<Section>): void {
    setTables(t => ({ ...t, sections: t.sections.map(x => x.id === id ? { ...x, ...patch } : x) }));
  }
  function patchQuestion(id: string, patch: Partial<Question>): void {
    setTables(t => ({ ...t, questions: t.questions.map(x => x.id === id ? { ...x, ...patch } : x) }));
  }
  function editField(field: 'name' | 'internalDescription' | 'title' | 'instructions', value: string): void {
    if (section) patchSection(section.id, { [field]: value });
    if (question) patchQuestion(question.id, { [field]: value });
  }
  function addSection(): void {
    if (!surveyId) return;
    const id = newId('section');
    setTables(t => ({ ...t, sections: [...t.sections, { id, surveyId,
      name: `section_${id.replace(/-/g, '_')}`, internalDescription: '', title: 'New section', instructions: '',
      sortOrder: t.sections.filter(x => x.surveyId === surveyId).length }] }));
    setSelected({ kind: 'section', id });
    setExpanded(x => ({ ...x, [id]: true }));
    setNotice('Section added.');
  }
  function addQuestion(sectionId: string): void {
    const id = newId('question');
    setTables(t => ({ ...t, questions: [...t.questions, { id, sectionId,
      name: `question_${id.replace(/-/g, '_')}`, internalDescription: '', dataType: 'text',
      title: 'New question', instructions: '', sortOrder: t.questions.filter(x => x.sectionId === sectionId).length }] }));
    setExpanded(x => ({ ...x, [sectionId]: true }));
    setSelected({ kind: 'question', id });
    setNotice('Question added.');
  }
  // Insert before a target, or append when beforeId is omitted. Reindex affected siblings.
  function moveSection(id: string, beforeId?: string): void {
    if (id === beforeId) return;
    setTables(t => {
      const source = t.sections.find(x => x.id === id);
      if (!source || source.surveyId !== surveyId) return t;
      const siblings = byOrder(t.sections.filter(x => x.surveyId === surveyId && x.id !== id));
      const index = beforeId ? siblings.findIndex(x => x.id === beforeId) : siblings.length;
      if (index < 0) return t;
      siblings.splice(index, 0, source);
      return { ...t, sections: t.sections.map(x => {
        const order = siblings.findIndex(y => y.id === x.id);
        return order >= 0 ? { ...x, sortOrder: order } : x;
      }) };
    });
    setNotice('Section order updated.');
  }
  function moveQuestion(id: string, targetSectionId: string, beforeId?: string): void {
    if (id === beforeId) return;
    setTables(t => {
      const source = t.questions.find(x => x.id === id);
      const allowed = t.sections.filter(x => x.surveyId === surveyId).map(x => x.id);
      if (!source || allowed.indexOf(source.sectionId) < 0 || allowed.indexOf(targetSectionId) < 0) return t;
      const destination = byOrder(t.questions.filter(x => x.sectionId === targetSectionId && x.id !== id));
      const index = beforeId ? destination.findIndex(x => x.id === beforeId) : destination.length;
      if (index < 0) return t;
      destination.splice(index, 0, { ...source, sectionId: targetSectionId });
      const oldSiblings = byOrder(t.questions.filter(x => x.sectionId === source.sectionId && x.id !== id));
      return { ...t, questions: t.questions.map(x => {
        const targetIndex = destination.findIndex(y => y.id === x.id);
        if (targetIndex >= 0) return { ...x, sectionId: targetSectionId, sortOrder: targetIndex };
        if (source.sectionId !== targetSectionId && x.sectionId === source.sectionId)
          return { ...x, sortOrder: oldSiblings.findIndex(y => y.id === x.id) };
        return x;
      }) };
    });
    setExpanded(x => ({ ...x, [targetSectionId]: true }));
    setNotice('Question position updated.');
  }
  function startDrag(event: React.DragEvent<HTMLElement>, item: DragItem): void {
    drag.current = item;
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', item.id);
  }
  function allowDrop(event: React.DragEvent<HTMLElement>, kind: DragItem['kind']): void {
    if (drag.current?.kind === kind) { event.preventDefault(); event.stopPropagation(); event.dataTransfer.dropEffect = 'move'; }
  }
  function dropSection(event: React.DragEvent<HTMLElement>, beforeId?: string): void {
    event.preventDefault(); event.stopPropagation();
    if (drag.current?.kind === 'section') moveSection(drag.current.id, beforeId);
    drag.current = null;
  }
  function dropQuestion(event: React.DragEvent<HTMLElement>, sectionId: string, beforeId?: string): void {
    event.preventDefault(); event.stopPropagation();
    if (drag.current?.kind === 'question') moveQuestion(drag.current.id, sectionId, beforeId);
    drag.current = null;
  }
  function moveButtons(kind: DragItem['kind'], id: string, ids: string[], parentId?: string): React.ReactElement {
    const i = ids.indexOf(id);
    const shift = (direction: -1 | 1): void => {
      // Inserting before the item two places ahead moves this item down one position.
      const before = direction === -1 ? ids[i - 1] : ids[i + 2];
      if (kind === 'section') moveSection(id, before);
      else if (parentId) moveQuestion(id, parentId, before);
    };
    return <span style={{ display: 'flex' }}>
      <button type="button" className="survey-designer__tiny" aria-label={`Move ${kind} up`} title="Move up" disabled={i <= 0} onClick={() => shift(-1)}>↑</button>
      <button type="button" className="survey-designer__tiny" aria-label={`Move ${kind} down`} title="Move down" disabled={i >= ids.length - 1} onClick={() => shift(1)}>↓</button>
    </span>;
  }
  function field(label: string, key: 'name' | 'internalDescription' | 'title' | 'instructions', multiline = false): React.ReactElement {
    return <TextField label={label} value={record?.[key] || ''} multiline={multiline} rows={multiline ? 3 : undefined}
      styles={fieldStyles} onChange={(_, value) => editField(key, value || '')} />;
  }

  return <div className="survey-designer">
    <header className="survey-designer__header">
      <div><div className="survey-designer__eyebrow">Survey Designer</div>
        <h1 className="survey-designer__title">{survey ? survey.surveyTitle || 'Untitled survey' : 'Survey designer'}</h1>
        <br />
        <div className="survey-designer__muted">{survey ? `${sections.length} sections · ${totalQuestions} questions` : 'Build reusable assessments for your assets.'}</div>
      </div>
      {survey && <DefaultButton text="← All surveys" onClick={() => { setSurveyId(null); setSelected(null); }} />}
    </header>
    {!survey ? <div className="survey-designer__cards">
      {tables.surveys.map(s => {
        const children = tables.sections.filter(x => x.surveyId === s.id);
        return <button type="button" key={s.id} className="survey-designer__card" onClick={() => openSurvey(s.id)}>
          <span className="survey-designer__badge">Survey template</span>
          <h2 style={{ fontSize: 19, margin: '18px 0 8px' }}>{s.surveyTitle || 'Untitled survey'}</h2>
          <div className="survey-designer__muted">{s.name}</div>
          <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 24 }}>
            <span className="survey-designer__muted">{children.length} sections</span><strong style={{ color: '#365fc7', fontSize: 13 }}>Open designer →</strong>
          </div>
        </button>;
      })}
    </div> : <>
      <div className="survey-designer__survey-fields">
        <TextField label="Internal survey name" value={survey.name} styles={fieldStyles}
          onChange={(_, value) => setTables(t => ({ ...t, surveys: t.surveys.map(x => x.id === survey.id ? { ...x, name: value || '' } : x) }))} />
        <TextField label="Published survey title" value={survey.surveyTitle} styles={fieldStyles}
          onChange={(_, value) => setTables(t => ({ ...t, surveys: t.surveys.map(x => x.id === survey.id ? { ...x, surveyTitle: value || '' } : x) }))} />
      </div>
      <div className="survey-designer__layout">
        <aside className="survey-designer__panel" aria-label="Survey structure">
          <div className="survey-designer__panel-header"><strong style={{ fontSize: 14 }}>Survey structure</strong><PrimaryButton text="+ Section" onClick={addSection} /></div>
          <div className="survey-designer__tree">
            <div className="survey-designer__muted" style={{ padding: '0 8px 12px', fontSize: 11 }}>Drag ⋮⋮ to reorder, or use the arrow buttons.</div>
            {sections.map((s, index) => {
              const children = questionsFor(s.id);
              return <div key={s.id} className="survey-designer__section">
                <div className={`survey-designer__row ${selected?.kind === 'section' && selected.id === s.id ? 'survey-designer__row--selected' : ''}`}
                  onDragOver={e => allowDrop(e, 'section')} onDrop={e => dropSection(e, s.id)}>
                  <span className="survey-designer__grip" draggable title="Drag section" aria-hidden="true"
                    onDragStart={e => startDrag(e, { kind: 'section', id: s.id })} onDragEnd={() => { drag.current = null; }}>⋮⋮</span>
                  <button type="button" className="survey-designer__tiny" aria-label={`${expanded[s.id] ? 'Collapse' : 'Expand'} ${s.title || s.name}`}
                    aria-expanded={!!expanded[s.id]} aria-controls={`questions-${s.id}`} onClick={() => setExpanded(x => ({ ...x, [s.id]: !x[s.id] }))}>{expanded[s.id] ? '▾' : '▸'}</button>
                  <button type="button" className="survey-designer__select" aria-pressed={selected?.kind === 'section' && selected.id === s.id} onClick={() => setSelected({ kind: 'section', id: s.id })}>
                    <span className="survey-designer__name">{index + 1}. {s.title || 'Untitled section'}</span><span className="survey-designer__meta">{s.name} · {children.length} questions</span>
                  </button>
                  {moveButtons('section', s.id, sectionIds)}
                </div>
                {expanded[s.id] && <div className="survey-designer__questions" id={`questions-${s.id}`}>
                  {children.map((q, qi) => <div key={q.id} className={`survey-designer__question ${selected?.kind === 'question' && selected.id === q.id ? 'survey-designer__row--selected' : ''}`}
                    onDragOver={e => allowDrop(e, 'question')} onDrop={e => dropQuestion(e, s.id, q.id)}>
                    <span className="survey-designer__grip" draggable title="Drag question" aria-hidden="true" onDragStart={e => startDrag(e, { kind: 'question', id: q.id })} onDragEnd={() => { drag.current = null; }}>⋮⋮</span>
                    <button type="button" className="survey-designer__select" aria-pressed={selected?.kind === 'question' && selected.id === q.id} onClick={() => setSelected({ kind: 'question', id: q.id })}>
                      <span className="survey-designer__name">{index + 1}.{qi + 1} {q.title || 'Untitled question'}</span><span className="survey-designer__meta">{q.dataType}</span>
                    </button>{moveButtons('question', q.id, children.map(x => x.id), s.id)}
                  </div>)}
                  <div className="survey-designer__drop" onDragOver={e => allowDrop(e, 'question')} onDrop={e => dropQuestion(e, s.id)}>Drop a question here to place it last</div>
                  <DefaultButton text="+ Add question" styles={{ root: { marginTop: 8, borderRadius: 6, width: '100%' } }} onClick={() => addQuestion(s.id)} />
                </div>}
              </div>;
            })}
            <div className="survey-designer__drop" onDragOver={e => allowDrop(e, 'section')} onDrop={e => dropSection(e)}>Drop a section here to place it last</div>
          </div>
        </aside>
        <main className="survey-designer__panel" aria-label="Edit selected item">
          {record ? <>
            <div className="survey-designer__editor" key={record.id}>
              <div className="survey-designer__field-group"><div className="survey-designer__group-title">Internal details <span className="survey-designer__muted">· Designer only</span></div>
                {field('Internal name', 'name')}{field('Internal description', 'internalDescription', true)}
                {question && <>
                  <Dropdown label="Data type" selectedKey={question.dataType} options={DATA_TYPES}
                    onChange={(_, option) => { if (option) patchQuestion(question.id, { dataType: option.key as DataType }); }} />
                  <Dropdown label="Section" selectedKey={question.sectionId} options={sections.map(x => ({ key: x.id, text: x.title || x.name || 'Untitled section' }))}
                    onChange={(_, option) => { if (option && option.key !== question.sectionId) moveQuestion(question.id, String(option.key)); }} />
                </>}
              </div>
              <div className="survey-designer__field-group"><div className="survey-designer__group-title">Published content <span className="survey-designer__muted">· Visible to respondents</span></div>
                {field('Title', 'title')}{field('Instructions', 'instructions', true)}
              </div>
            </div>
          </> : <div className="survey-designer__empty"><h2>Select a section or question</h2><p>Add a section to start building this survey.</p><PrimaryButton text="Add section" onClick={addSection} /></div>}
        </main>
      </div>
      {notice && <div style={{ marginTop: 12 }}><MessageBar messageBarType={MessageBarType.info} onDismiss={() => setNotice('')} dismissButtonAriaLabel="Dismiss notification">{notice}</MessageBar></div>}
    </>}
  </div>;
}
