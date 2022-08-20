import React = require("react");
import ReactDOM = require('react-dom');
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { d365notesfilepreviewControl, Id365notesfilepreviewControlProps } from './d365notesfilepreviewControl';

export class d365notesfilepreview implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _context: ComponentFramework.Context<IInputs>;
    private _container: HTMLDivElement;
    private _themeMainColor: string;
    /**
     * Empty constructor.
     */
    constructor()
    {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
     public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
     {
         // Add control initialization code
         this._context = context;
         this._container = container;
     }
 
 
     /**
      * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
      * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
      */
     public updateView(context: ComponentFramework.Context<IInputs>): void
     {
         this._context = context;
 
         if (this._themeMainColor) {
             ReactDOM.render(React.createElement(d365notesfilepreviewControl, this.getControlProps()), this._container);
         }
         else {
             this.getMainThemeColor().then(color => {
                 this._themeMainColor = color;
 
                 ReactDOM.render(React.createElement(d365notesfilepreviewControl, this.getControlProps()), this._container);
             });
         }
     }
 
     /** 
      * It is called by the framework prior to a control receiving new data. 
      * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
      */
     public getOutputs(): IOutputs
     {
         return {};
     }
 
     /** 
      * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
      * i.e. cancelling any pending remote calls, removing listeners, etc.
      */
     public destroy(): void
     {
         // Add code to cleanup control if necessary
     }
 
     private getControlProps(): Id365notesfilepreviewControlProps {
         let entityId = (<any>this._context).page.entityId;
         let entityName = (<any>this._context).page.entityTypeName;
 
         return {
             context: this._context,
             themeMainColor: this._themeMainColor,
             entityId: entityId,
             entityName: entityName
         };
     };
 
     private getMainThemeColor(): Promise<string> {
         return this._context.webAPI.retrieveMultipleRecords('theme', '?$select=maincolor&$filter=isdefaulttheme eq true').then((result) => {
             return result.entities[0]['maincolor'].toString();
         }, (err) => {
             this._context.navigation.openErrorDialog({ message: this._context.resources.getString('errorThemeRetrieval') });
         });
     };
 }