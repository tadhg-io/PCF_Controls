export interface IQuestion {
  id: string;
  sectionId: string;
  name: string;
  internalDescription: string;
  title: string;
  instructions: string;
  dataType: string;
  required: boolean;
  options?: string[];
}


type QuestionType = 'Text' | 'Number' | 'YesNo' | 'Date' | 'Choice';