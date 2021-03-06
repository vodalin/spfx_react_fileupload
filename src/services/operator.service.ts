import {SPAPIhelperService} from './SPAPIhelper.service';
import {IPropertyPaneDropdownOption} from '@microsoft/sp-webpart-base';
import {filter} from "minimatch";
import resolve = Promise.resolve;
import {obj} from "through2";

export class OperatorService {
  private spCaller: SPAPIhelperService;
  private curContext: any;
  private byTitleSect = '/lists/getByTitle(\'%\')';
  private byServerRelativeSect = '/getfolderbyserverrelativeurl(\'%\')';
  private byGUIDSelect = '/Lists(guid\'%\')';
  private subsiteRef: string;

  constructor(context){
    this.curContext = context;
    this.spCaller = new SPAPIhelperService(context);
    this.subsiteRef = this.curContext.pageContext.web.serverRelativeUrl; // /sites/devsite
  }

  // These functions make dropdown options for the property pane.
  public getAllLibraries(){
    /*Gets all doc library names in site, returns them as drop down options.*/
    return Promise.resolve(this.spCaller.getSPData(
      '/Lists',
      {
        Filter: 'Hidden eq false and BaseTemplate eq 101',
        Select: 'Title'
      })
    ).then(results =>{
      return this.makeDropOptions(results, 'Title', 'Title');
    }).catch(err => {
      throw new Error('getAllLibraries() failed' + '\n' + err);
    });
  }

  public getListFields(targetList: string) {
    /**/
    let fieldMetaData = {};
    return Promise.resolve(this.spCaller.getSPData(
      this.byTitleSect.replace('%', targetList) + '/Fields',
      {
        Filter: 'Group eq \'Custom Columns\' and Hidden eq false',
        Select: 'Title,Id,LookupList,InternalName,Choices',
        Top: '2000'
      })
    ).then(results =>{
      fieldMetaData['dropdownoptions'] = this.makeDropOptions(results, 'InternalName','Title');
      fieldMetaData['rawdata'] = results;
      return fieldMetaData;
    });
  }

  private makeDropOptions (optionData: Array<any>, keySelector: string, textSelector: string): Array<IPropertyPaneDropdownOption>{
    const ddOptions: IPropertyPaneDropdownOption[] = [];
    optionData['value'].forEach((list) => {
      ddOptions.push({
        key: list[keySelector],
        text: list[textSelector]
      });
    });
    return ddOptions;
  }

  public getItemsByGUID(guid: string) {
    return Promise.resolve(this.spCaller.getSPData(
      this.byGUIDSelect.replace('%', guid) + '/Items',
      {
        Select: 'Id,Title',
        Top: 1000
      }
    ));
  }

  public startUploads(submit_data, target_folder, target_library){
    let upload_pr_list = [];
    let edit_pr_list = [];
    let subfolder_path = target_library + '/' + target_folder;
    let processed_sub_data = stripDictKeys(submit_data);

    let submit_data_keys = Object.keys(processed_sub_data);
    Object.keys(processed_sub_data).forEach(filekey => {
      let raw_file = processed_sub_data[filekey]['raw_file'];
      upload_pr_list.push(this.spCaller.uploadFiles(raw_file, subfolder_path));
    });

    return Promise.all(upload_pr_list)
      .then(val => {
        return Promise.resolve(this.spCaller.getSPData(
          this.byTitleSect.replace('%', target_library) + '/items',
          {
            Select: '*,LinkFilename,FileDirRef',
            filter: encodeURI('FileDirRef eq \'' + this.subsiteRef + '/' + subfolder_path + '\''),
            Top: 1000,
          }
        ));
      })
      .then(results => {
        let filelist = results['value'];
        filelist.forEach(file => {
          let fileId = file['Id'];
          let fileName = file['LinkFilename'];
          if(submit_data_keys.indexOf(fileName) > -1){
            let coldata = {};
            let subdata = processed_sub_data[fileName];
            Object.keys(subdata).forEach(key =>{
              if(key != 'raw_file'){
                let finalvalue = undefined;
                let colkey = '';
                //When id != undefined, it indicates a lookup column
                //Lookup column REST names always end with an 'Id'
                subdata[key]['Id'] == undefined ? (
                  finalvalue = subdata[key]['Text'],
                  colkey=key
                ) : (
                  finalvalue = subdata[key]['Id'],
                  colkey=key + 'Id');
                coldata[colkey] = String(finalvalue);
              }
            });
            edit_pr_list.push(this.spCaller.setItemColumn(target_library, fileId, coldata));
          }
        });
        return edit_pr_list;
      })
      .then(pr_list =>{
        return Promise.all(pr_list)
          .then(val =>{
            console.log('All done');
            return Promise.resolve();
          });
      });
  }

}

function stripDictKeys(input_dict) {
  Object.keys(input_dict).forEach(key => {
    let dictData = input_dict[key];
    let strippedKey = key.replace(/['!&*?=\/|\\":<>]/g,'');
    delete input_dict[key];
    input_dict[strippedKey] = dictData;
  });
  return input_dict;
}

