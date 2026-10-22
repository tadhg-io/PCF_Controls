import { IInputs } from "../../SurveyDesigner/generated/ManifestTypes";

export class ApiService {
    private _context: ComponentFramework.Context<IInputs>;

    constructor(context: ComponentFramework.Context<IInputs>) {
        this._context = context;
    }

    public getSections = async (): Promise<any> => { 

        
        var fetchXml = [
            "<fetch>",
            "  <entity name='toc_section'>",
            "    <attribute name='toc_name' />",
            "  </entity>",
            "</fetch>",
        ].join("");

        let odataUrl = `?fetchXml=${fetchXml}`;   
        return this._context.webAPI.retrieveMultipleRecords("toc_section", odataUrl);
        

    };

}