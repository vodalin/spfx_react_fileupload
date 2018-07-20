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
import {OperatorService} from "../../services/operator.service";

export interface IFileuploaderWebPartProps {
  pr_targetlibrary: string;
  pr_addbutton: string;
}

export default class FileuploaderWebPart extends BaseClientSideWebPart<IFileuploaderWebPartProps> {
  public ops: OperatorService;
  public fieldKeyTemplate = 'targetField_';
  public fieldKeyRegExp = new RegExp(this.fieldKeyTemplate, 'g');
  public reqFieldMetaData = [];

  // Master list of all property pane elements
  public propPaneList: Array<any> = [];

  // Lists for dropdown menu options
  public reqFieldOptions = [];
  public libraryOptions = [];

  // Data passed to child webparts
  public targetLib: string = undefined;
  public fieldSchema = {};

  protected onInit(): Promise<void>{
    window['webPartContext'] = this.context;
    this.ops = new OperatorService(this.context);
    return super.onInit();
  }

  public render(): void {
    /* For each render, scan the webpart's properties, then set/create property pane elements accordingly.
    * Afterwords, set the <fieldSchema> variable.*/
    Promise.all(this.initPPaneOptions())
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

  /**************** Helper functions ********************/
  private initPPaneOptions(){
    /* Returns list of promises that create options for the property pane dropdown menus.*/
    let targetLibName = this.properties['pr_targetlibrary'];
    let prArray = [this.setLibraryOptions()];
    if(targetLibName != undefined){
      this.targetLib = targetLibName;
      prArray.push(this.setReqFieldData());
    }
    return prArray;
  }

  private setLibraryOptions() {
    /* Gets names of all document libraries in the site. Then sets options for the "Target Library" dropdown. */
    return Promise.resolve(this.ops.getAllLibraries())
      .then((results: IPropertyPaneDropdownOption[]) => {
        this.libraryOptions = results;
      });
  }

  private setReqFieldData(){
    /* When a target library is defined, this sets the options for all "Required Field" dropdowns
    <reqFieldOptions> in property pane. Also stores the metadata for each field in <reqFieldMetaData> */
    return Promise.resolve(this.ops.getListFields(this.targetLib))
      .then(fielddata => {
        this.reqFieldOptions = fielddata['dropdownoptions'];
        this.reqFieldMetaData = fielddata['rawdata']['value'];
      });
  }

  private setFieldSchema(): void {
    /* Compares <reqFieldMetaData> and required fields in webpart properties. Then, populates
     * <fieldSchema>, which is used to define dropdowns in miniselector component. */
    let selected_fileds = [];
    this.getPropKeys().forEach(key => {
      if(key.match(this.fieldKeyRegExp)){
        selected_fileds.push(this.properties[key]);
      }
    });

    this.fieldSchema = {};
    selected_fileds.forEach((fieldname: string) =>{
      this.fieldSchema[fieldname] = {header_text: '', field_data: {}};
      this.reqFieldMetaData.forEach(item => {

        if(fieldname == item['InternalName']) {
          this.fieldSchema[fieldname].header_text = item['Title'];

          if(item['@odata.type'] == '#SP.FieldLookup'){
            let strippedGUID = item['LookupList'].replace(new RegExp('[{}]','g'), '');
            Promise.resolve(this.ops.getItemsByGUID(strippedGUID))
              .then(results => {
                this.fieldSchema[fieldname].field_data = results;
              });
          }
          else if(item['@odata.type'] == '#SP.FieldChoice'){
            let choiceList = item['Choices'].map(c => { return {Id: undefined, Title: c}; });
            this.fieldSchema[fieldname].field_data = {
              value: choiceList
            };
          }
        }

      });
    });
  }

  private resetProperties(): void {
    /* Removes any required fields from the webpart's properties. */
    this.getPropKeys().map(key => {
      if(key.match(this.fieldKeyRegExp)){
        delete this.properties[key];
      }
    });
  }

  private resetPropPaneList(): void{
    /* Removes 'Required Field' dropdown elements from <propPaneList> */
    const delete_list = this.propPaneList.filter(element =>
      element['targetProperty'].match(this.fieldKeyRegExp) == null
    );
    this.propPaneList = delete_list;
  }

  /**************** Property Pane Generator functions ********************/

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

  /**************** Misc stuff ********************/
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
