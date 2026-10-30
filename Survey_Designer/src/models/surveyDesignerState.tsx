import { ApiService } from "../services/apiservice";
import { IInputs } from "../../SurveyDesigner/generated/ManifestTypes";
import { ISection } from "./section";
import { IQuestion } from "./question";

export interface ISurveyDesignerState {  

    // // inputs
    // lookupType: lookupTypes,
    // targetField: string,
    // tableName: string,
    
    // // local variables
    sections: ISection[];
    questions?: IQuestion[];
    selectedSectionId?: string | null;
    selectedQuestionId?: string | null;
    selection: Selection;
    expanded?: string[];
    dirty?: boolean;

    // // outputs
    // selectedValue? : string,

    // // UX control
    // isLoading?: boolean,
    // successMessage?: string,
    // errorMessage?: string,

}

export interface ISurveyDesignerProps extends ISurveyDesignerState { 
    context?: ComponentFramework.Context<IInputs>;
    onSave?: (sections: ISection[]) => void;
    apiService?: ApiService;
}

type Selection = { kind?: 'section'; sectionId?: string } |
  { kind?: 'question'; sectionId?: string; questionId?: string };