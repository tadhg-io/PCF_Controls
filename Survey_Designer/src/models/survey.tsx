import { ISection } from "./section";

export interface ISurvey {
    id: string;
    name: string;
    surveyTitle: string;
    Sections: ISection[];
}