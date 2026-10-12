import * as React from 'react';
import {
  DefaultButton, Dropdown, Icon, IconButton, IDropdownOption,
  PrimaryButton, TextField, Toggle, initializeIcons,
} from 'office-ui-fabric-react';

// Standalone, locally stateful UI. No PCF inputs or Dataverse connection required.
initializeIcons();

type QuestionType = 'Text' | 'Number' | 'YesNo' | 'Date' | 'Choice';

interface Question {
  id: string;
  name: string;
  prompt: string;
  help: string;
  type: QuestionType;
  required: boolean;
}

interface Section {
  id: string;
  name: string;
  description: string;
  questions: Question[];
}

type Selection = { kind: 'section'; sectionId: string } |
  { kind: 'question'; sectionId: string; questionId: string };

const initialSections: Section[] = [
  {
    id: 'people', name: 'People at risk', description: 'Identify people who may need assistance during evacuation.',
    questions: [
      { id: 'sleep', name: 'Sleeping occupants', prompt: 'How many sleeping occupants are present?', help: 'Enter the maximum number normally present overnight.', type: 'Number', required: true },
      { id: 'access', name: 'Accessibility needs', prompt: 'Are any occupants likely to require evacuation assistance?', help: 'Include temporary and permanent mobility limitations.', type: 'YesNo', required: true },
      { id: 'lone', name: 'Lone working', prompt: 'Are there lone workers on the premises?', help: 'Consider out-of-hours and weekend work.', type: 'YesNo', required: false },
    ],
  },
  {
    id: 'ignition', name: 'Sources of ignition', description: 'Assess common sources of ignition and combustible materials.',
    questions: [
      { id: 'electrical', name: 'Electrical inspections', prompt: 'Are electrical installations inspected regularly?', help: 'Review inspection records.', type: 'YesNo', required: true },
      { id: 'heating', name: 'Heating systems', prompt: 'Describe the heating systems on site.', help: 'Include fuel types and proximity to combustible materials.', type: 'Text', required: false },
      { id: 'flammable', name: 'Flammable storage', prompt: 'Are flammable materials stored safely?', help: 'Consider liquids, gases and waste.', type: 'YesNo', required: true },
    ],
  },
  {
    id: 'escape', name: 'Escape routes', description: 'Review routes, exits and evacuation signage.',
    questions: [
      { id: 'clear', name: 'Clear routes', prompt: 'Are escape routes clear and unobstructed?', help: 'Check corridors, stairwells and final exits.', type: 'YesNo', required: true },
      { id: 'exits', name: 'Emergency exits', prompt: 'How many emergency exits are available?', help: 'Count usable exits.', type: 'Number', required: true },
      { id: 'signage', name: 'Exit signage', prompt: 'Is emergency exit signage adequate?', help: 'Include illuminated signs where applicable.', type: 'Choice', required: false },
    ],
  },
  {
    id: 'detection', name: 'Detection & warning', description: 'Record fire detection and alarm arrangements.',
    questions: [
      { id: 'alarm', name: 'Alarm system', prompt: 'Is an automatic fire alarm installed?', help: 'Record coverage and system type.', type: 'YesNo', required: true },
      { id: 'test', name: 'Last alarm test', prompt: 'When was the alarm last tested?', help: 'Use the fire logbook.', type: 'Date', required: false },
    ],
  },
];

const typeOptions: IDropdownOption[] = [
  { key: 'Text', text: 'Text response' }, { key: 'Number', text: 'Number' },
  { key: 'YesNo', text: 'Yes / No' }, { key: 'Date', text: 'Date' },
  { key: 'Choice', text: 'Multiple choice' },
];
const typeIcons: Record<QuestionType, string> = {
  Text: 'TextField', Number: 'NumberField', YesNo: 'CheckboxComposite',
  Date: 'Calendar', Choice: 'BulletedList',
};
const uid = () => `item-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;

const css = `
.sd-root, .sd-root * { box-sizing: border-box; }
.sd-root { --ink:#172339; --muted:#728097; --line:#e6eaf0; --blue:#4565dd; --pale:#edf1ff; font-family: 'Segoe UI',system-ui,sans-serif; color:var(--ink); background:#f5f7fa; height:100%; min-height:580px; width:100%; display:flex; flex-direction:column; font-size:13px; }
.sd-top { min-height:65px; padding:12px 22px; display:flex; align-items:center; justify-content:space-between; gap:16px; background:#fff; border-bottom:1px solid var(--line); }
.sd-brand { display:flex; align-items:center; gap:12px; min-width:0; }
.sd-logo { width:36px; height:36px; background:#1d2b48; border-radius:10px; color:#fff; display:grid; place-items:center; font-size:17px; }
.sd-eyebrow { color:var(--muted); font-size:10px; text-transform:uppercase; letter-spacing:1.2px; font-weight:700; }
.sd-title { font-size:17px; font-weight:700; letter-spacing:-.45px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.sd-top-actions { display:flex; align-items:center; gap:12px; }
.sd-status { font-size:12px; color:var(--muted); white-space:nowrap; }
.sd-status.unsaved { color:#a66c23; }
.sd-body { display:flex; min-height:0; flex:1; padding:16px; gap:16px; }
.sd-nav { flex:0 0 325px; min-width:245px; max-width:42%; background:#fff; border:1px solid var(--line); border-radius:12px; display:flex; flex-direction:column; min-height:0; overflow:hidden; box-shadow:0 3px 14px rgba(24,39,67,.025); }
.sd-nav-head { display:flex; align-items:center; justify-content:space-between; padding:16px 16px 12px; }
.sd-panel-title { font-size:13px; font-weight:700; letter-spacing:-.15px; }
.sd-count { font-size:11px; font-weight:600; background:#f0f2f6; border-radius:20px; padding:3px 8px; color:#67758a; }
.sd-nav-scroll { flex:1; overflow:auto; padding:2px 9px 14px; }
.sd-section { margin:5px 0; }
.sd-section-row { display:flex; align-items:center; min-height:43px; border-radius:8px; gap:5px; padding:3px 5px; cursor:pointer; border:1px solid transparent; }
.sd-section-row:hover, .sd-question:hover { background:#f6f8fc; }
.sd-section-row.active { background:var(--pale); border-color:#dce4ff; }
.sd-section-name { flex:1; min-width:0; font-weight:650; font-size:12.5px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.sd-section-index { color:#98a2b4; font-size:11px; width:17px; text-align:right; font-variant-numeric:tabular-nums; }
.sd-section-row.active .sd-section-name { color:#304bb5; }
.sd-question-list { padding:3px 0 5px 26px; margin-left:13px; border-left:1px solid #e9edf3; }
.sd-question { width:100%; display:flex; align-items:center; text-align:left; border:1px solid transparent; border-radius:7px; background:transparent; min-height:35px; padding:3px 5px 3px 10px; gap:8px; cursor:pointer; color:#58677d; font:inherit; }
.sd-question.active { background:var(--pale); border-color:#dce4ff; color:#304bb5; }
.sd-question-label { flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-size:12px; font-weight:500; }
.sd-question.active .sd-question-label { font-weight:650; }
.sd-move { display:flex; align-items:center; gap:0; opacity:0; }
.sd-section-row:hover .sd-move, .sd-section-row:focus-within .sd-move, .sd-question:hover .sd-move, .sd-question:focus-within .sd-move, .sd-section-row.active .sd-move, .sd-question.active .sd-move { opacity:1; }
.sd-nav-footer { padding:12px; border-top:1px solid var(--line); }
.sd-add-section { width:100%; }
.sd-editor { min-width:0; flex:1; background:#fff; border:1px solid var(--line); border-radius:12px; display:flex; flex-direction:column; overflow:hidden; box-shadow:0 3px 14px rgba(24,39,67,.025); }
.sd-editor-head { padding:19px 24px 16px; border-bottom:1px solid var(--line); display:flex; align-items:center; justify-content:space-between; gap:12px; }
.sd-breadcrumb { color:var(--muted); font-size:11px; margin-bottom:6px; }
.sd-editor-title { font-size:21px; font-weight:700; letter-spacing:-.6px; }
.sd-pill { background:#f0f3f8; color:#68778e; padding:5px 10px; border-radius:30px; font-size:11px; font-weight:600; white-space:nowrap; }
.sd-editor-content { padding:22px 24px; overflow:auto; flex:1; }
.sd-form { max-width:820px; }
.sd-group { margin-bottom:24px; }
.sd-group-head { display:flex; align-items:center; gap:9px; margin-bottom:14px; }
.sd-group-icon { width:27px; height:27px; background:#f0f3fb; border-radius:7px; display:grid; place-items:center; color:#5266ad; }
.sd-group-title { font-size:13px; font-weight:700; }
.sd-group-hint { font-size:11px; color:var(--muted); margin-top:2px; }
.sd-fields { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:15px; }
.sd-span2 { grid-column:1/-1; }
.sd-field .ms-Label { font-size:12px; color:#4a5a71; font-weight:600; padding-top:0; }
.sd-field .ms-TextField-fieldGroup, .sd-field .ms-Dropdown-title { border-radius:6px; border-color:#dce2eb; }
.sd-field .ms-TextField-fieldGroup:hover, .sd-field .ms-Dropdown:hover .ms-Dropdown-title { border-color:#aebad0; }
.sd-divider { height:1px; background:var(--line); margin:4px 0 22px; }
.sd-editor-foot { border-top:1px solid var(--line); padding:12px 24px; display:flex; justify-content:space-between; align-items:center; gap:10px; }
.sd-foot-note { color:var(--muted); font-size:11px; }
.sd-empty { padding:40px; color:var(--muted); }
.sd-root .ms-Button--primary { background:#405fda; border-color:#405fda; border-radius:6px; }
.sd-root .ms-Button--primary:hover { background:#304dc2; border-color:#304dc2; }
.sd-root .ms-Button--default { border-radius:6px; }
.sd-root .sd-icon-button { width:23px; height:25px; color:#8592a6; }
.sd-root .sd-icon-button:hover { background:#e8edf8; color:#344eab; }
@media(max-width:740px) { .sd-body { padding:8px; gap:8px; } .sd-nav { flex-basis:245px; } .sd-fields { grid-template-columns:1fr; } .sd-editor-head,.sd-editor-content { padding:15px; } .sd-editor-foot { padding:10px 15px; } .sd-top { padding:10px; } }
`;

export interface SurveyDesignerProps {
  onSave?: (sections: Section[]) => void;
}

export const SurveyDesigner: React.FC<SurveyDesignerProps> = ({ onSave }) => {
  const [sections, setSections] = React.useState<Section[]>(initialSections);
  const [selection, setSelection] = React.useState<Selection>({ kind: 'question', sectionId: 'people', questionId: 'sleep' });
  const [expanded, setExpanded] = React.useState<string[]>(['people', 'ignition', 'escape', 'detection']);
  const [dirty, setDirty] = React.useState(false);
  const section = sections.find(s => s.id === selection.sectionId);
  const question = selection.kind === 'question' ? section?.questions.find(q => q.id === selection.questionId) : undefined;
  const totalQuestions = sections.reduce((n, s) => n + s.questions.length, 0);

  const updateSection = (patch: Partial<Section>) => {
    setSections(prev => prev.map(s => s.id === selection.sectionId ? { ...s, ...patch } : s));
    setDirty(true);
  };
  const updateQuestion = (patch: Partial<Question>) => {
    if (selection.kind !== 'question') return;
    setSections(prev => prev.map(s => s.id === selection.sectionId
      ? { ...s, questions: s.questions.map(q => q.id === selection.questionId ? { ...q, ...patch } : q) }
      : s));
    setDirty(true);
  };
  const moveSection = (id: string, direction: number) => {
    setSections(prev => {
      const at = prev.findIndex(s => s.id === id), to = at + direction;
      if (at < 0 || to < 0 || to >= prev.length) return prev;
      const next = [...prev]; [next[at], next[to]] = [next[to], next[at]]; return next;
    });
    setDirty(true);
  };
  const moveQuestion = (sectionId: string, questionId: string, direction: number) => {
    setSections(prev => prev.map(s => {
      if (s.id !== sectionId) return s;
      const at = s.questions.findIndex(q => q.id === questionId), to = at + direction;
      if (at < 0 || to < 0 || to >= s.questions.length) return s;
      const questions = [...s.questions]; [questions[at], questions[to]] = [questions[to], questions[at]];
      return { ...s, questions };
    }));
    setDirty(true);
  };
  const addSection = () => {
    const id = uid();
    setSections(prev => [...prev, { id, name: 'Untitled section', description: '', questions: [] }]);
    setExpanded(prev => [...prev, id]);
    setSelection({ kind: 'section', sectionId: id });
    setDirty(true);
  };
  const addQuestion = (sectionId: string) => {
    const id = uid();
    setSections(prev => prev.map(s => s.id === sectionId ? {
      ...s, questions: [...s.questions, { id, name: 'Untitled question', prompt: '', help: '', type: 'Text', required: false }],
    } : s));
    setExpanded(prev => prev.includes(sectionId) ? prev : [...prev, sectionId]);
    setSelection({ kind: 'question', sectionId, questionId: id });
    setDirty(true);
  };
  const deleteSelected = () => {
    if (!section) return;
    if (selection.kind === 'question') {
      setSections(prev => prev.map(s => s.id === section.id ? { ...s, questions: s.questions.filter(q => q.id !== selection.questionId) } : s));
      setSelection({ kind: 'section', sectionId: section.id });
    } else {
      const remaining = sections.filter(s => s.id !== section.id);
      setSections(remaining);
      if (remaining.length) setSelection({ kind: 'section', sectionId: remaining[0].id });
    }
    setDirty(true);
  };
  const save = () => {
    if (onSave) onSave(sections);
    else console.log('Survey definition (local demo):', sections);
    // Production: clear dirty only after persistence succeeds.
    setDirty(false);
  };
  const iconButton = (name: string, label: string, click: (e: React.MouseEvent<HTMLButtonElement>) => void, disabled = false) => (
    <IconButton className="sd-icon-button" iconProps={{ iconName: name }} title={label} ariaLabel={label}
      disabled={disabled} onClick={click} />
  );
  const groupHeading = (icon: string, title: string, hint: string) => (
    <div className="sd-group-head"><span className="sd-group-icon"><Icon iconName={icon} /></span>
      <div><div className="sd-group-title">{title}</div><div className="sd-group-hint">{hint}</div></div></div>
  );

  return <div className="sd-root">
    <style>{css}</style>
    <header className="sd-top">
      <div className="sd-brand"><span className="sd-logo"><Icon iconName="ClipboardList" /></span>
        <div><div className="sd-eyebrow">Survey studio / Draft</div><div className="sd-title">Fire risk assessment</div></div>
      </div>
      <div className="sd-top-actions"><span className={`sd-status ${dirty ? 'unsaved' : ''}`}>{dirty ? '● Unsaved changes' : 'All changes saved locally'}</span>
        <PrimaryButton text="Save changes" iconProps={{ iconName: 'Save' }} onClick={save} disabled={!dirty} />
      </div>
    </header>
    <div className="sd-body">
      <nav className="sd-nav" aria-label="Survey structure">
        <div className="sd-nav-head"><span className="sd-panel-title">Structure</span><span className="sd-count">{sections.length} sections · {totalQuestions} questions</span></div>
        <div className="sd-nav-scroll">
          {sections.map((s, index) => {
            const isOpen = expanded.includes(s.id);
            const selected = selection.kind === 'section' && selection.sectionId === s.id;
            return <div className="sd-section" key={s.id}>
              <div className={`sd-section-row ${selected ? 'active' : ''}`}
                onClick={() => setSelection({ kind: 'section', sectionId: s.id })}>
                {iconButton(isOpen ? 'ChevronDown' : 'ChevronRight', isOpen ? 'Collapse section' : 'Expand section', e => {
                  e.stopPropagation(); setExpanded(prev => isOpen ? prev.filter(id => id !== s.id) : [...prev, s.id]);
                })}
                <span className="sd-section-index">{String(index + 1).padStart(2, '0')}</span>
                <span className="sd-section-name" title={s.name}>{s.name}</span>
                <span className="sd-move">
                  {iconButton('ChevronUp', 'Move section up', e => { e.stopPropagation(); moveSection(s.id, -1); }, index === 0)}
                  {iconButton('ChevronDown', 'Move section down', e => { e.stopPropagation(); moveSection(s.id, 1); }, index === sections.length - 1)}
                </span>
                {iconButton('Add', 'Add question', e => { e.stopPropagation(); addQuestion(s.id); })}
              </div>
              {isOpen && <div className="sd-question-list">
                {s.questions.map((q, qi) => <button type="button" key={q.id}
                  className={`sd-question ${selection.kind === 'question' && selection.questionId === q.id ? 'active' : ''}`}
                  onClick={() => setSelection({ kind: 'question', sectionId: s.id, questionId: q.id })}>
                  <Icon iconName={typeIcons[q.type]} style={{ fontSize: 13, flexShrink: 0 }} />
                  <span className="sd-question-label" title={q.name}>{q.name}</span>
                  <span className="sd-move">
                    {iconButton('ChevronUp', 'Move question up', e => { e.stopPropagation(); moveQuestion(s.id, q.id, -1); }, qi === 0)}
                    {iconButton('ChevronDown', 'Move question down', e => { e.stopPropagation(); moveQuestion(s.id, q.id, 1); }, qi === s.questions.length - 1)}
                  </span>
                </button>)}
              </div>}
            </div>;
          })}
        </div>
        <div className="sd-nav-footer"><DefaultButton className="sd-add-section" text="Add section" iconProps={{ iconName: 'Add' }} onClick={addSection} /></div>
      </nav>
      <main className="sd-editor">
        {section ? <>
          <div className="sd-editor-head"><div>
            <div className="sd-breadcrumb">Survey / {section.name}{question ? ` / ${question.name}` : ''}</div>
            <div className="sd-editor-title">{question ? 'Edit question' : 'Edit section'}</div>
          </div><span className="sd-pill">{question ? 'QUESTION' : 'SECTION'}</span></div>
          <div className="sd-editor-content"><div className="sd-form">
            {question ? <>
              <div className="sd-group">{groupHeading('EditNote', 'Question details', 'How this question appears in the survey')}
                <div className="sd-fields">
                  <TextField className="sd-field" label="Question name" value={question.name} onChange={(_, v) => updateQuestion({ name: v || '' })} />
                  <Dropdown className="sd-field" label="Answer type" options={typeOptions} selectedKey={question.type}
                    onChange={(_, option) => { if (option) updateQuestion({ type: option.key as QuestionType }); }} />
                  <TextField className="sd-field sd-span2" label="Question text" value={question.prompt} onChange={(_, v) => updateQuestion({ prompt: v || '' })} />
                  <TextField className="sd-field sd-span2" label="Help text" multiline rows={3} value={question.help} onChange={(_, v) => updateQuestion({ help: v || '' })} />
                </div>
              </div><div className="sd-divider" />
              <div className="sd-group">{groupHeading('Settings', 'Behaviour', 'Response settings')}
                <Toggle label="Response required" checked={question.required} onChange={(_, v) => updateQuestion({ required: !!v })} />
              </div>
            </> : <div className="sd-group">{groupHeading('Folder', 'Section details', 'Organise questions into meaningful groups')}
              <div className="sd-fields">
                <TextField className="sd-field sd-span2" label="Section title" value={section.name} onChange={(_, v) => updateSection({ name: v || '' })} />
                <TextField className="sd-field sd-span2" label="Description / instructions" multiline rows={4} value={section.description} onChange={(_, v) => updateSection({ description: v || '' })} />
              </div>
            </div>}
          </div></div>
          <div className="sd-editor-foot"><span className="sd-foot-note">Changes are held locally until saved.</span>
            <DefaultButton text={question ? 'Delete question' : 'Delete section'} iconProps={{ iconName: 'Delete' }} onClick={deleteSelected}
              styles={{ root: { color: '#a44747' } }} /></div>
        </> : <div className="sd-empty">Add a section to begin designing your survey.</div>}
      </main>
    </div>
  </div>;
};

export default SurveyDesigner;
