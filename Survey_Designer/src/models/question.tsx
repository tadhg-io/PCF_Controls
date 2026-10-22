export interface IQuestion {
  id: string;
  sectionId: string;
  questionText: string;
  internalDescription: string;
  instructions: string;
  questionType: string;
  options?: string[];
}