import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration, IPropertyPaneDropdownOption,
  PropertyPaneDropdown, PropertyPaneButton, PropertyPaneButtonType
} from '@microsoft/sp-webpart-base';

import * as strings from 'FileuploaderWebPartStrings';
import Fileuploader from './components/Fileuploader';
import IFileuploaderProps from './components/IFileuploaderProps';

// export interface IFileuploaderWebPartProps {
//   description: string;
//   targetlib: any;
//   required_fields: any;
//   target_fields: string;
// }

export interface IFileuploaderWebPartProps {
  description: string;
  pr_targetlibrary: string;
  pr_addbutton: string;
}

import {OperatorService} from "../../services/operator.service";
import {forIn} from "@microsoft/sp-lodash-subset";
// import {IFileuploaderWebPartProps} from "../../../lib/webparts/fileuploader/FileuploaderWebPart";


export default class FileuploaderWebPart extends BaseClientSideWebPart<IFileuploaderWebPartProps> {
  public ops: OperatorService;
  public fieldKeyTemplate = 'targetField_';
  public fieldKeyRegExp = new RegExp(this.fieldKeyTemplate, 'g');

  // Master list of all property pane elements
  public propPaneList: Array<any> = [];

  // Lists for dropdown menu options
  public reqFieldOptions = [];
  public libraryOptions = [];

  // Data passed to child webparts
  public targetLib: string = undefined;
  public reqFieldMetaData = [];
  public fieldSchema = {};

  protected onInit(): Promise<void>{
    window['webPartContext'] = this.context;
    this.ops = new OperatorService(this.context);


    return super.onInit();
  }

  public render(): void {
    Promise.all(this.scanProperties())
      .then(val => {
        this.plugPaneElement([ this.createLibraryDropDown(), this.createAddButton()]);
        this.getPropKeys().forEach(key => {
          if (key.match(this.fieldKeyRegExp)){
            this.plugPaneElement([this.addRequiredField(key)]);
          }
        });
        return(this.setFieldSchema());
      })
      .then(val => {
        const element: React.ReactElement<IFileuploaderProps> = React.createElement(
          Fileuploader,
          {
            description: this.properties.description,
            target_library: this.targetLib,
            required_fields_schema: this.fieldSchema,
          }
        );
        ReactDom.render(element, this.domElement);
      });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'pr_targetlibrary') { // only runs when the targetLib property is changed.
      this.targetLib = newValue;
      this.resetProperties();
      this.resetPropPaneList();
      this.context.propertyPane.refresh();
    }
    else if(propertyPath === 'pr_addbutton') {
      this.plugPaneElement([this.addRequiredField()]);
    }
    return super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: this.propPaneList,
            }
          ]
        }
      ]
    };
  }

  //**************** Helper functions ********************
  private scanProperties(){
    /* Scans the web part properties and updates relevant variables */
    let prArray = [this.getLibraryOptions()];
    this.getPropKeys().forEach(key => {
      if(key == 'pr_targetlibrary'){
        this.targetLib = this.properties[key];
        prArray.push(this.setFieldData());
      }
    });
    return prArray;
  }

  private getLibraryOptions() {
    return Promise.resolve(this.ops.getAllLibraries())
      .then((results: IPropertyPaneDropdownOption[]) => {
        this.libraryOptions = results;
      });
  }

  private setFieldData(){
    /*Updates the 'required field' info which is used throughout this webpart and its child components */
    return Promise.resolve(this.ops.getListFields(this.targetLib))
      .then(fielddata => {
        this.reqFieldOptions = fielddata['dropdownoptions'];
        this.reqFieldMetaData = fielddata['rawdata']['value'];
      });
  }

  private setFieldSchema(): void {
    /* Check if any 'required fields' are lookup types. If so, then get the lookup table's data (title,Id)
    * and set it in this.fieldSchema.
    */
    let field_metadata = this.reqFieldMetaData;
    let selected_fileds = [];
    this.getPropKeys().forEach(key => {
      if(key.match(this.fieldKeyRegExp)){
        selected_fileds.push(this.properties[key]);
      }
    });

    this.fieldSchema = {};
    selected_fileds.forEach((fieldname: string) =>{
      this.fieldSchema[fieldname] = {};
      field_metadata.forEach(item => {
        if(fieldname == item['InternalName'] && item['@odata.type'] == '#SP.FieldLookup'){
          let strippedGUID = item['LookupList'].replace(new RegExp('[{}]','g'), '');
          Promise.resolve(this.ops.getItemsByGUID(strippedGUID))
            .then(results => {
              this.fieldSchema[fieldname] = results;
            });
        }
      });
    });
  }

  private resetProperties(): void {
    /* Removes all required fields from the webpart's properties. */
    this.getPropKeys().map(key => {
      if(key.match(this.fieldKeyRegExp)){
        delete this.properties[key];
      }
    });
  }

  private resetPropPaneList(): void{
    /*Removes all 'required field' dropdown elements from <this.propPaneList> */
    const delete_list = this.propPaneList.filter(element =>
      element['targetProperty'].match(this.fieldKeyRegExp) == null
    );
    this.propPaneList = delete_list;
  }

  //**************** Property Pane Element Functions ********************

  private createAddButton() {
    /* Creates add field button */
    return PropertyPaneButton('pr_addbutton', {
      text: "Add Field",
      buttonType: PropertyPaneButtonType.Primary,
      onClick: ()=>{},
    });
  }

  private createLibraryDropDown(){
    /* Creates 'Target Library' dropdown menu */
    return PropertyPaneDropdown('pr_targetlibrary', {
      label: 'Target Library',
      options: this.libraryOptions
    });
  }

  private plugPaneElement(pPaneElements: Array<any>): void{
    /*Checks to see if element already exists before adding it to <this.propPaneList> */
    let existing_fields = this.propPaneList.map(ele => {
      return ele['targetProperty'];
    });
    pPaneElements.forEach(element => {
      let eleId = element['targetProperty'];
      if(existing_fields.indexOf(eleId) == -1){
        this.propPaneList.push(element);
      }
    });
  }

  private addRequiredField(fieldname?){
    /* Adds a 'required field' dropdown menu to the property pane. */
    let property_label = '';
    (fieldname != undefined) ? property_label = fieldname : property_label = this.fieldKeyTemplate + (this.getReqFieldCount() + 1).toString();

    return PropertyPaneDropdown(property_label, {
      label: 'Required Field',
      options: this.reqFieldOptions
    });
  }

  //**************** Misc stuff ********************
  private getReqFieldCount() {
    /*Counts number of web part elements that contain <this.fieldKeyTemplate> in their 'targetProperty' property. */
    let property_list = this.getAllProperties();
    let existing_rfields = property_list.filter(prop_name => {
      return prop_name.match(this.fieldKeyRegExp);
    });
    return existing_rfields.length;
  }

  private getAllProperties(): Array<string> {
    /* Gets all elements in the propPaneList, returns list of each element's 'targetProperty' value. */
    const list: Array<string> =[];
    this.propPaneList.forEach(item => {
      list.push(item.targetProperty);
    });
    return list;
  }

  private getPropKeys(){
    /* Return all keys in <this.properties> */
    return Object.keys(this.properties);
  }
}
