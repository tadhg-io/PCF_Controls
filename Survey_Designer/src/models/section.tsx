import { IQuestion } from "./question";

export interface ISection {
    id: string;
    name: string;
    sectionTitle: string;
    internalDescription: string;
    instructions: string;
    surveyId: string;
    Questions: IQuestion[];
}