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
import { IFileuploaderProps } from './components/IFileuploaderProps';
//***
export interface IFileuploaderWebPartProps {
  description: string;
  targetlib: any;
  required_fields: any;
  target_fields: string;
}
import {OperatorService} from "../../services/operator.service";

export default class FileuploaderWebPart extends BaseClientSideWebPart<IFileuploaderWebPartProps> {
  public ops: OperatorService;
  public fieldKeyTemplate = 'targetField_';
  public propPaneList: Array<any> = []; // Master list of all property pane elements
  // Lists for dropdown menu options
  public docLibNames = []; // List of all document libraries
  public fieldOptions = [];
  // Flags to signal data retrieval + panel refresh
  public startfetch = false;
  public addedfield = false;
  // Data passed to webparts
  //public targetLib: string = 'default'; // Target document library name
  public targetLib: string = undefined;
  public selectedFields = [];
  public fieldMetaData = [];
  // public fieldSchema = [];
  public fieldSchema = {};

  public render(): void {
    window['webPartContext'] = this.context;
    this.ops = new OperatorService(this.context);
    this.startfetch = true;

    Promise.resolve(this.initParams())
      .then(val => {
        return Promise.all(this.getLookupData());
      })
      .then(val => {
        const element: React.ReactElement<IFileuploaderWebPartProps> = React.createElement(
          Fileuploader,
          {
            description: this.properties.description,
            target_library: this.targetLib,
            required_fields: this.selectedFields,
            required_fields_metadata: this.fieldMetaData,
            required_fields_schema: this.fieldSchema
          }
        );
        //console.log(this.properties, this.selectedFields, this.fieldMetaData);
        ReactDom.render(element, this.domElement);
      });

  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  // Sets flags / variables to for when the getOptions() method is fired.
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'targetlib') { // only runs when the targetLib property is changed.
      this.fieldSchema = {};
      this.propPaneList = [];
      this.deleteOldFileds();
      this.startfetch = true;
      this.targetLib = newValue;
    }
    else if(propertyPath === 'addbutton') {
      if(this.targetLib !== undefined){
        this.addedfield = true;
      }
    }
    else{
      this.fieldSchema = {};
      this.getSelectedFields();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    Promise.resolve()
      .then(() => {
        if (this.startfetch) {
          return Promise.all([this.initialFetch(), this.updateListFields()])
            .then(results =>{
              this.startfetch = false;
              this.checkForFields();
              this.context.propertyPane.refresh(); // Forces property pane to refresh again.
            })
            .catch(err => {
              throw new Error(err.stack);
            });
        }

        if(this.addedfield) {
          return Promise.resolve(this.addRequiredField())
            .then(results => {
              this.addedfield = false;
              this.context.propertyPane.refresh(); // Forces property pane to refresh again.
            });
        }
      })
      .catch(err => {
        console.log(err.stack);
        alert('Error loading Property Pane.');
      });
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: this.propPaneList
            }
          ]
        }
      ]
    };
  }

  //*** Helper functions
  private initParams(){
    //console.log(this.properties);
    let pr = Promise.resolve()
      .then(val =>{
        let keylist = Object.keys(this.properties);
        keylist.forEach((key: string) => {
          if (key == 'targetlib'){
            this.targetLib = this.properties[key];
          }
        });
        return this.updateListFields();
      });
    this.getSelectedFields();
    return pr;
  }

  private initialFetch(){
    let pBtn_AddField = PropertyPaneButton('addbutton', {
      text: "Add Field",
      buttonType: PropertyPaneButtonType.Primary,
      onClick: ()=>{},
    });

    // Populates the Target Library dropdown menu in webpart config.
    return Promise.resolve(this.ops.getAllLibraries())
      .then((results: IPropertyPaneDropdownOption[]) => {
        let pDD_DocumentLibraries = PropertyPaneDropdown('targetlib', {
          label: 'Target Library',
          options: results
        });
        this.pushProp([pDD_DocumentLibraries, pBtn_AddField]);
      })
      .catch(err => {
        throw new Error('initialFetch() failed \n' + err.stack);
      });
  }

  //Things that need to be updated when startfetch is set to true
  private updateListFields(){
    let fieldPr = Promise.resolve('Empty');
    if(this.targetLib != undefined){
      fieldPr = Promise.resolve(this.ops.getListFields(this.targetLib))
        .then(fielddata => {
          this.fieldOptions = fielddata['dropdownoptions'];
          this.fieldMetaData = fielddata['rawdata']['value'];
          return 'Done';
        })
        .catch(err => {
          throw new Error('updateListField failed' + '\n' + err.stack);
        });
    }
    return fieldPr;
  }

  private checkForFields() {
    let keylist = Object.keys(this.properties);
    keylist.forEach((key: string) => {
      if (key.search(this.fieldKeyTemplate) != -1){
        this.addRequiredField(key);
      }
    });
  }

  private addRequiredField(fieldname?: string){
    let button_label = this.fieldKeyTemplate + (this.getHighestTFI() + 1).toString();
    if(fieldname != undefined) {
      button_label = fieldname;
    }

    let pDD_LibraryFields = PropertyPaneDropdown(button_label, {
      label: 'Required Field',
      options: this.fieldOptions
    });
    this.pushProp([pDD_LibraryFields]);
  }

  //**** Misc stuff
  private pushProp(propFields: Array<any>){
    propFields.forEach(prop => {
      if((this.getAllProperties()).indexOf(prop.targetProperty) == -1){
        this.propPaneList.push(prop);
      }
    });
  }

  private getAllProperties(): Array<string> {
    const list: Array<string> =[];
    this.propPaneList.forEach(item => {
      list.push(item.targetProperty);
    });
    return list;
  }

  private deleteOldFileds(): void {
    //Remove required fields from this.properties
    let keylist = Object.keys(this.properties);
    keylist.forEach((key: string) => {
      if (key.search(this.fieldKeyTemplate) != -1){
        delete this.properties[key];
      }
    });
  }

  private getSelectedFields() {
    let selectedFieldsArray = [];
    let keylist = Object.keys(this.properties);
    keylist.forEach((key: string) => {
      if (key.search(this.fieldKeyTemplate) != -1){
        selectedFieldsArray.push(this.properties[key]);
      }
    });
    this.selectedFields = selectedFieldsArray;
  }

  private getHighestTFI() {
    let highest_index = 0;
    let keylist = Object.keys(this.properties);
    keylist.forEach((key: string) => {
      if (key.search(this.fieldKeyTemplate) != -1){
        highest_index = parseInt(key.split('_')[1]);
      }
    });
    return highest_index;
  }

  private getLookupData() {
    let target_fileds = this.selectedFields;
    let field_metadata = this.fieldMetaData;
    let prarray = [];

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
}
