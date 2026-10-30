import { IInputs } from "../../SurveyDesigner/generated/ManifestTypes";
import { IQuestion } from "../models/question";
import { ISection } from "../models/section";

export class ApiService {
    private _context: ComponentFramework.Context<IInputs>;

    constructor(context: ComponentFramework.Context<IInputs>) {
        this._context = context;
    }

    public getSections = async (): Promise<ISection[]> => {
        var fetchXml = [
            "<fetch>",
            "  <entity name='toc_section'>",
            "    <attribute name='toc_name' />",
            "    <attribute name='toc_title' />",
            "    <attribute name='toc_internaldescription' />",
            "    <attribute name='toc_instructions' />",
            "    <attribute name='toc_assessmentdefinition' />",
            "  </entity>",
            "</fetch>",
        ].join("");

        let odataUrl = `?fetchXml=${fetchXml}`;
        return await this._context.webAPI.retrieveMultipleRecords("toc_section", odataUrl).then(

            async (result) => {
                let sections: ISection[] = result.entities.map((entity) => {
                    return {
                        id: entity.toc_sectionid,
                        name: entity.toc_name,
                        sectionTitle: entity.toc_title,
                        internalDescription: entity.toc_internaldescription,
                        instructions: entity.toc_instructions,
                        surveyId: entity.toc_assessmentdefinition,
                        questions: [],
                    };
                });
                return sections;
            },

            (error) => {
                console.log(error.message);
                return Promise.reject(error.message);
            }
        );;


    };

    public getQuestions = async (): Promise<IQuestion[]> => {

        const fetchXml = [
            "<fetch>",
            "  <entity name='toc_question'>",
            "    <attribute name='toc_questionid' />",
            "    <attribute name='toc_name' />",
            "    <attribute name='toc_title' />",
            "    <attribute name='toc_internaldescription' />",
            "    <attribute name='toc_instructions' />",
            "    <attribute name='toc_datatype' />",
            "    <attribute name='toc_section' />",
            "  </entity>",
            "</fetch>"
        ].join("");

        const result = await this._context.webAPI.retrieveMultipleRecords(
            "toc_question",
            `?fetchXml=${fetchXml}`
        );

        return result.entities.map((entity): IQuestion => ({
            id: entity.toc_questionid,
            sectionId: entity._toc_section_value ?? "",
            name: entity.toc_name ?? "",
            title: entity.toc_title ?? "",
            internalDescription: entity.toc_internaldescription ?? "",
            instructions: entity.toc_instructions ?? "",
            dataType: entity.toc_datatype,
            required: false
        }));
    };

}