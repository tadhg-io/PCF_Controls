export interface IQuestion {
  id: string;
  sectionId: string;
  name: string;
  internalDescription: string;
  title: string;
  instructions: string;
  dataType: QuestionDataType;
  required: boolean;
  options?: string[];
}

// enum for data type
export enum QuestionDataType {
    Text = 932400000,
    Number = 932400001,
    Boolean = 932400002,
    Option = 932400003,
}

