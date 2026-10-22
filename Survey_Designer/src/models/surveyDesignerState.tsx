import { ApiService } from "../services/apiservice";
import { IInputs } from "../../SurveyDesigner/generated/ManifestTypes";

export interface ISurveyDesignerState {  

    // // inputs
    // lookupType: lookupTypes,
    // targetField: string,
    // tableName: string,
    
    // // local variables
    // tables?: ITable[],
    // columns?: string[],

    // // outputs
    // selectedValue? : string,

    // // UX control
    // isLoading?: boolean,
    // successMessage?: string,
    // errorMessage?: string,

}

export interface ISurveyDesignerProps extends ISurveyDesignerState { 
    context?: ComponentFramework.Context<IInputs>;
    apiService?: ApiService;
}