import * as React from 'react';
import '../css/designer.css';
import {
  DefaultButton, Dropdown, Icon, IconButton, IDropdownOption,
  PrimaryButton, TextField, Toggle, initializeIcons,
} from 'office-ui-fabric-react';
import { ApiService } from "../services/apiservice";
import { ISection } from '../models/section';

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

export interface SurveyDesignerProps {
  onSave?: (sections: Section[]) => void;
  apiService?: ApiService;
}

interface SurveyDesignerState {
  sections: Section[];
  selection: Selection;
  expanded: string[];
  dirty: boolean;
}

export class SurveyDesign extends React.Component<SurveyDesignerProps, SurveyDesignerState> {

  private _apiService?: ApiService;
  private _sections: string[] = [];

  constructor(props: SurveyDesignerProps) {
    super(props);
    this.state = {
      sections: initialSections,
      selection: { kind: 'question', sectionId: 'people', questionId: 'sleep' },
      expanded: ['people', 'ignition', 'escape', 'detection'],
      dirty: false,
    };
    // get the API service from the props
    this._apiService = props.apiService;  

    this._apiService?.getSections().then((result: any) => {

          console.log("result", result);
          // loop through the rsults and build the array
          result.entities.forEach((t:any) => {
              let name = t.toc_name;
              this._sections.push(name);
          });

        }, (error: any) => {
          console.log("_apiService error", error);
        });
  }

  private get section(): Section | undefined {
    return this.state.sections.find(s => s.id === this.state.selection.sectionId);
  }

  private get question(): Question | undefined {
    const section = this.section;
    const selection = this.state.selection;

    if (selection.kind !== 'question') {
      return undefined;
    }

    return section?.questions.find(q => q.id === selection.questionId);
  }

  private get totalQuestions(): number {
    return this.state.sections.reduce((n, s) => n + s.questions.length, 0);
  }

  private updateSection = (patch: Partial<Section>) => {
    const { sectionId } = this.state.selection;
    this.setState(prev => ({
      sections: prev.sections.map(s => s.id === sectionId ? { ...s, ...patch } : s),
      dirty: true,
    }));
  };

  private updateQuestion = (patch: Partial<Question>) => {
    const { selection } = this.state;
    if (selection.kind !== 'question') return;
    this.setState(prev => ({
      sections: prev.sections.map(s => s.id === selection.sectionId
        ? { ...s, questions: s.questions.map(q => q.id === selection.questionId ? { ...q, ...patch } : q) }
        : s),
      dirty: true,
    }));
  };

  private moveSection = (id: string, direction: number) => {
    this.setState(prev => {
      const at = prev.sections.findIndex(s => s.id === id);
      const to = at + direction;
      if (at < 0 || to < 0 || to >= prev.sections.length) return prev;
      const next = [...prev.sections];
      [next[at], next[to]] = [next[to], next[at]];
      return { ...prev, sections: next, dirty: true };
    });
  };

  private moveQuestion = (sectionId: string, questionId: string, direction: number) => {
    this.setState(prev => ({
      sections: prev.sections.map(s => {
        if (s.id !== sectionId) return s;
        const at = s.questions.findIndex(q => q.id === questionId);
        const to = at + direction;
        if (at < 0 || to < 0 || to >= s.questions.length) return s;
        const questions = [...s.questions];
        [questions[at], questions[to]] = [questions[to], questions[at]];
        return { ...s, questions };
      }),
      dirty: true,
    }));
  };

  private addSection = () => {
    const id = uid();
    this.setState(prev => ({
      sections: [...prev.sections, { id, name: 'Untitled section', description: '', questions: [] }],
      expanded: [...prev.expanded, id],
      selection: { kind: 'section', sectionId: id },
      dirty: true,
    }));
  };

  private addQuestion = (sectionId: string) => {
    const id = uid();
    this.setState(prev => ({
      sections: prev.sections.map(s => s.id === sectionId ? {
        ...s,
        questions: [...s.questions, { id, name: 'Untitled question', prompt: '', help: '', type: 'Text', required: false }],
      } : s),
      expanded: prev.expanded.includes(sectionId) ? prev.expanded : [...prev.expanded, sectionId],
      selection: { kind: 'question', sectionId, questionId: id },
      dirty: true,
    }));
  };

  private deleteSelected = () => {
    const selection = this.state.selection;
    const section = this.section;
    if (!section) return;

    if (selection.kind === 'question') {
      this.setState(prev => ({
        sections: prev.sections.map(s => s.id === section.id ? {
          ...s,
          questions: s.questions.filter(q => q.id !== selection.questionId),
        } : s),
        selection: { kind: 'section', sectionId: section.id },
        dirty: true,
      }));
      return;
    }

    const remaining = this.state.sections.filter(s => s.id !== section.id);
    this.setState({
      sections: remaining,
      selection: remaining.length ? { kind: 'section', sectionId: remaining[0].id } : { kind: 'section', sectionId: '' },
      dirty: true,
    });
  };

  private save = () => {
    const { onSave } = this.props;
    const { sections } = this.state;

    if (onSave) {
      onSave(sections);
    } else {
      console.log('Survey definition (local demo):', sections);
    }

    this.setState({ dirty: false });
  };

  private iconButton = (name: string, label: string, click: (e: React.MouseEvent<HTMLButtonElement>) => void, disabled = false) => (
    <IconButton
      className="sd-icon-button"
      iconProps={{ iconName: name }}
      title={label}
      ariaLabel={label}
      disabled={disabled}
      onClick={click}
    />
  );

  private groupHeading = (icon: string, title: string, hint: string) => (
    <div className="sd-group-head">
      <span className="sd-group-icon"><Icon iconName={icon} /></span>
      <div>
        <div className="sd-group-title">{title}</div>
        <div className="sd-group-hint">{hint}</div>
      </div>
    </div>
  );

  render() {
    const { sections, selection, expanded, dirty } = this.state;
    const section = this.section;
    const question = this.question;
    const totalQuestions = this.totalQuestions;

    return <div className="sd-root">
      <header className="sd-top">
        <div className="sd-brand"><span className="sd-logo"><Icon iconName="ClipboardList" /></span>
          <div><div className="sd-eyebrow">Survey studio /+ {this._sections.length}</div><div className="sd-title">Fire risk assessment</div></div>
        </div>
        <div className="sd-top-actions"><span className={`sd-status ${dirty ? 'unsaved' : ''}`}>{dirty ? '● Unsaved changes' : 'All changes saved locally'}</span>
          <PrimaryButton text="Save changes" iconProps={{ iconName: 'Save' }} onClick={this.save} disabled={!dirty} />
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
                  onClick={() => this.setState({ selection: { kind: 'section', sectionId: s.id } })}>
                  {this.iconButton(isOpen ? 'ChevronDown' : 'ChevronRight', isOpen ? 'Collapse section' : 'Expand section', e => {
                    e.stopPropagation();
                    this.setState(prev => ({
                      expanded: isOpen ? prev.expanded.filter(id => id !== s.id) : [...prev.expanded, s.id],
                    }));
                  })}
                  <span className="sd-section-index">{String(index + 1).padStart(2, '0')}</span>
                  <span className="sd-section-name" title={s.name}>{s.name}</span>
                  <span className="sd-move">
                    {this.iconButton('ChevronUp', 'Move section up', e => { e.stopPropagation(); this.moveSection(s.id, -1); }, index === 0)}
                    {this.iconButton('ChevronDown', 'Move section down', e => { e.stopPropagation(); this.moveSection(s.id, 1); }, index === sections.length - 1)}
                  </span>
                  {this.iconButton('Add', 'Add question', e => { e.stopPropagation(); this.addQuestion(s.id); })}
                </div>
                {isOpen && <div className="sd-question-list">
                  {s.questions.map((q, qi) => <button type="button" key={q.id}
                    className={`sd-question ${selection.kind === 'question' && selection.questionId === q.id ? 'active' : ''}`}
                    onClick={() => this.setState({ selection: { kind: 'question', sectionId: s.id, questionId: q.id } })}>
                    <Icon iconName={typeIcons[q.type]} style={{ fontSize: 13, flexShrink: 0 }} />
                    <span className="sd-question-label" title={q.name}>{q.name}</span>
                    <span className="sd-move">
                      {this.iconButton('ChevronUp', 'Move question up', e => { e.stopPropagation(); this.moveQuestion(s.id, q.id, -1); }, qi === 0)}
                      {this.iconButton('ChevronDown', 'Move question down', e => { e.stopPropagation(); this.moveQuestion(s.id, q.id, 1); }, qi === s.questions.length - 1)}
                    </span>
                  </button>)}
                </div>}
              </div>;
            })}
          </div>
          <div className="sd-nav-footer"><DefaultButton className="sd-add-section" text="Add section" iconProps={{ iconName: 'Add' }} onClick={this.addSection} /></div>
        </nav>
        <main className="sd-editor">
          {section ? <>
            <div className="sd-editor-head"><div>
              <div className="sd-breadcrumb">Survey / {section.name}{question ? ` / ${question.name}` : ''}</div>
              <div className="sd-editor-title">{question ? 'Edit question' : 'Edit section'}</div>
            </div><span className="sd-pill">{question ? 'QUESTION' : 'SECTION'}</span></div>
            <div className="sd-editor-content"><div className="sd-form">
              {question ? <>
                <div className="sd-group">{this.groupHeading('EditNote', 'Question details', 'How this question appears in the survey')}
                  <div className="sd-fields">
                    <TextField className="sd-field" label="Question name" value={question.name} onChange={(_, v) => this.updateQuestion({ name: v || '' })} />
                    <Dropdown className="sd-field" label="Answer type" options={typeOptions} selectedKey={question.type}
                      onChange={(_, option) => { if (option) this.updateQuestion({ type: option.key as QuestionType }); }} />
                    <TextField className="sd-field sd-span2" label="Question text" value={question.prompt} onChange={(_, v) => this.updateQuestion({ prompt: v || '' })} />
                    <TextField className="sd-field sd-span2" label="Help text" multiline rows={3} value={question.help} onChange={(_, v) => this.updateQuestion({ help: v || '' })} />
                  </div>
                </div><div className="sd-divider" />
                <div className="sd-group">{this.groupHeading('Settings', 'Behaviour', 'Response settings')}
                  <Toggle label="Response required" checked={question.required} onChange={(_, v) => this.updateQuestion({ required: !!v })} />
                </div>
              </> : <div className="sd-group">{this.groupHeading('Folder', 'Section details', 'Organise questions into meaningful groups')}
                <div className="sd-fields">
                  <TextField className="sd-field sd-span2" label="Section title" value={section.name} onChange={(_, v) => this.updateSection({ name: v || '' })} />
                  <TextField className="sd-field sd-span2" label="Description / instructions" multiline rows={4} value={section.description} onChange={(_, v) => this.updateSection({ description: v || '' })} />
                </div>
              </div>}
            </div></div>
            <div className="sd-editor-foot"><span className="sd-foot-note">Changes are held locally until saved.</span>
              <DefaultButton text={question ? 'Delete question' : 'Delete section'} iconProps={{ iconName: 'Delete' }} onClick={this.deleteSelected}
                styles={{ root: { color: '#a44747' } }} /></div>
          </> : <div className="sd-empty">Add a section to begin designing your survey.</div>}
        </main>
      </div>
    </div>;
  }
}
export default SurveyDesign;
