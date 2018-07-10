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

export interface IFileuploaderWebPartProps {
  description: string;
  targetlib: any;
  required_fields: any;
  target_fields: string;
}

export interface IFileuploaderWebPartProps2 {
  description: string;
  pr_targetlibrary: string;
  pr_addbutton: string;
}

import {OperatorService} from "../../services/operator.service";
import {IFileuploaderWebPartProps} from "../../../lib/webparts/fileuploader/FileuploaderWebPart";

export default class FileuploaderWebPart extends BaseClientSideWebPart<IFileuploaderWebPartProps2> {
  /*this.propertes = {
   pr_targetlibrary : "2017 Compliance Documents",
   targetField_1 : "FileLeafRef"
  }*/
  public ops: OperatorService;
  public fieldKeyTemplate = 'targetField_';
  public fieldKeyRegExp = new RegExp(this.fieldKeyTemplate, 'g');
  public propPaneList: Array<any> = []; // Master list of all property pane elements
  // Lists for dropdown menu options
  public reqFieldOptions = [];
  // Flags to signal data retrieval + panel refresh
  public refetch = false;
  // public addedfield = false;
  // Data passed to webparts
  public targetLib: string = undefined;
  public selectedFields = [];
  public fieldMetaData = [];
  public fieldSchema = {};


  //******************* NOW THE PROPERTIES IN FILEUPLOADER AREN'T UPDATING
  protected onInit(): Promise<void>{
    window['webPartContext'] = this.context;
    this.ops = new OperatorService(this.context);
    this.targetLib = this.properties['pr_targetlibrary'];

    /*If <this.properties> contains old data, generate the appropriate property pane elements.*/
    Promise.all([this.initialSetup()])
      .then(val => {
        if(this.targetLib != undefined){
          return this.setFieldData();
        }
      })
      .then(val => {
        this.getPropKeys().forEach(key => {
          if(key.match(this.fieldKeyRegExp)){
            this.addRequiredField(key);
          }
        });
        //this.refetch = true;
        //  this.setSelectedFields();
        //  this.getLookupData();
      });
    return super.onInit();
  }

  public render(): void {
    this.setSelectedFields();
    Promise.all(this.getLookupData())
      .then(val =>{
        const element: React.ReactElement<IFileuploaderWebPartProps> = React.createElement(
          Fileuploader,
          {
            description: this.properties.description,
            target_library: this.targetLib,
            required_fields: this.selectedFields,
            //required_fields_metadata: this.fieldMetaData,
            required_fields_schema: this.fieldSchema,
          }
        );
        ReactDom.render(element, this.domElement);
      });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'pr_targetlibrary') { // only runs when the targetLib property is changed.
      this.targetLib = newValue;
      Promise.resolve(this.setFieldData())
        .then(val => {
          this.resetProperties();
          this.resetPropPaneList();
          this.context.propertyPane.refresh();
        });
    }
    else if(propertyPath === 'pr_addbutton') {
      this.addRequiredField();
    }
    else {
      //this.refetch = true;
      //  this.setSelectedFields();
      //  this.getLookupData();
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
  private initialSetup(){
    /*On startup, create 'add field' button and target library dropdown.*/
    let pBtn_AddField = PropertyPaneButton('pr_addbutton', {
      text: "Add Field",
      buttonType: PropertyPaneButtonType.Primary,
      onClick: ()=>{},
    });

    return Promise.resolve(this.ops.getAllLibraries())
      .then((results: IPropertyPaneDropdownOption[]) => {
        let pDD_DocumentLibraries = PropertyPaneDropdown('pr_targetlibrary', {
          label: 'Target Library',
          options: results
        });
        this.propPaneList.push(pDD_DocumentLibraries, pBtn_AddField);
      })
  }

  private setFieldData(){
    /*Updates the field metadata used throughout this webpart and its child components */
    return Promise.resolve(this.ops.getListFields(this.targetLib))
      .then(fielddata => {
        this.reqFieldOptions = fielddata['dropdownoptions'];
        this.fieldMetaData = fielddata['rawdata']['value'];
      });
  }

  private setSelectedFields() {
    /* Uses <this.selectedFileds> to track which required fields were selected. */
    let selectedFieldsArray = [];
    this.getPropKeys().forEach((key: string) => {
      if (key.match(this.fieldKeyRegExp)){
        selectedFieldsArray.push(this.properties[key]);
      }
    });
    this.selectedFields = selectedFieldsArray;
  }

  private addRequiredField(fieldname?){
    /* Adds a 'required field' dropdown menu to the property pane. */
    let property_label = '';
    (fieldname != undefined) ? property_label = fieldname : property_label = this.fieldKeyTemplate + (this.getReqFieldCount() + 1).toString();

    let pDD_LibraryFields = PropertyPaneDropdown(property_label, {
      label: 'Required Field',
      options: this.reqFieldOptions
    });

    this.propPaneList.push(pDD_LibraryFields);
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

  private getLookupData() {
    /* Check if any 'required fields' are lookup types. If so, then get the lookup table's data (title,Id)
    * and set it in this.fieldSchema.
    */
    let target_fileds = this.selectedFields;
    let field_metadata = this.fieldMetaData;
    let prarray = [];
    this.fieldSchema = {};

    target_fileds.forEach((fieldname: string) => {
      this.fieldSchema[fieldname] = {};
      field_metadata.forEach(item => {
        if(fieldname == item['InternalName'] && item['@odata.type'] == '#SP.FieldLookup'){
          let strippedGUID = item['LookupList'].replace(new RegExp('[{}]','g'), '');
          prarray.push(
            Promise.resolve(this.ops.getItemsByGUID(strippedGUID))
              .then(results => {
                this.fieldSchema[fieldname] = results;
              })
          );
        }
      });
    });
    return prarray;
  }

  //**************** Misc stuff ********************
  private getReqFieldCount() {
    /*Counts number of web part elements that contain <this.fieldKeyTemplate> in their 'targetProperty' property. */
    let property_list = this.getAllProperties();
    let existing_rfields = property_list.filter(prop_name => {
      return prop_name.match(this.fieldKeyRegExp)
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
