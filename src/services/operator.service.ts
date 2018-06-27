import {SPAPIhelperService} from './SPAPIhelper.service';
import {IPropertyPaneDropdownOption} from '@microsoft/sp-webpart-base';
import {filter} from "minimatch";

export class OperatorService {
  private spCaller: SPAPIhelperService;
  private curContext: any;
  private byTitleSect = '/lists/getByTitle(\'%\')';
  private byServerRelativeSect = '/getfolderbyserverrelativeurl(\'%\')';
  private byGUIDSelect = '/Lists(guid\'%\')';

  constructor(context){
    this.curContext = context;
    this.spCaller = new SPAPIhelperService(context);
  }

  public getAllLibraries(){
    return Promise.resolve(this.spCaller.getSPData(
      '/Lists',
      {
        Filter: 'Hidden eq false and BaseTemplate eq 101',
        Select: 'Title'
      })
    ).then(results =>{
      return this.makeDropOptions(results, 'Title');
    }).catch(err => {
      throw new Error('getAllLibraries() failed' + '\n' + err);
    });
  }

  public getListFields(targetList: string) {
    let fieldMetaData = {};
    return Promise.resolve(this.spCaller.getSPData(
      this.byTitleSect.replace('%', targetList) + '/Fields',
      {
        Filter: 'Group eq \'Custom Columns\' and Hidden eq false',
        Select: 'Title,Id,LookupList,InternalName',
        Top: '2000'
      })
    ).then(results =>{
      fieldMetaData['dropdownoptions'] = this.makeDropOptions(results, 'InternalName');
      fieldMetaData['rawdata'] = results;
      return fieldMetaData;
    });
  }

  private makeDropOptions (spDataList: Array<any>, keyselector: string): Array<IPropertyPaneDropdownOption>{
    const ddOptions: IPropertyPaneDropdownOption[] = [];
    spDataList['value'].forEach((list) => {
      ddOptions.push({
        key: list[keyselector],
        text: list[keyselector]
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


  // // Helper functions that specify filter object for getItemsFrom()
  // public fetchSubOptions(listname: string) {
  //   return this.getItemsFrom(listname, {select: 'Title,Id'});
  // }
  //
  // public fetchSubfolderItems(targetFolder: string, subfoldername: string) {
  //   const filterobject = {
  //     select: '*,LinkFilename,FileDirRef',
  //     filter: encodeURI('FileDirRef eq \'' + this.curContext.pageContext.web.serverRelativeUrl + subfoldername + '\'')
  //   };
  //   return this.getItemsFrom(targetFolder, filterobject);
  // }
  //
  // // Main caller that uses SPAPIhelper serivce to get items
  // public getItemsFrom(targetList: string, filterObject?: Object): Promise<any> {
  //   const subFolList = [];
  //   return Promise.resolve(this.spCaller.getListItems(targetList, filterObject))
  //     .then((json: Object) => {
  //       json['value'].forEach(entry => {
  //         subFolList.push(entry);
  //       });
  //       return subFolList;
  //     })
  //     .catch(err =>{
  //       console.log('Cannot build options lists. \n' + err.stack);
  //     });
  // }
  //
  // // Helper functions for filteredSearch().
  // public getAllLibs() {
  //   return this.filteredSearch('getLists' ,{
  //     Filter: 'Hidden eq false and BaseTemplate eq 101',
  //     Select: 'Title'
  //   });
  // }
  //
  // public getAllLists() {
  //   return this.filteredSearch('getLists',{
  //     Filter: 'Hidden eq false and BaseTemplate eq 100',
  //     Select: 'Title'
  //   });
  // }
  //
  // public getAllFields(targetList: string) {
  //   return this.filteredSearch('getListColumnNames', {
  //     listName: targetList,
  //     filterParams: {
  //       Filter: 'Group eq \'Custom Columns\' and Hidden eq false',
  //       Select: 'Title'
  //     }
  //   });
  // }
  //
  // // Returns object: any[{key: '', text: ''}], then appwebparts casts the results to type IPropertyPaneDropdownOption[]
  // private filteredSearch (spFnName: string, param: string | object): Promise<any>{
  //   const selector = param['filterParams'] === undefined ? param['Select'] : param['filterParams']['Select'];
  //   const ddOptions: IPropertyPaneDropdownOption[] = [];
  //   return Promise.resolve(this.spCaller[spFnName](param))
  //     .then(results => {
  //       results['value'].forEach((list) => {
  //         ddOptions.push({
  //           key: list[selector],
  //           text: list[selector]
  //         });
  //       });
  //       return ddOptions;
  //     })
  //     .catch(err => {
  //       console.log('Lists not found for webpart property panel \n' + err.stack); // ** Final error log
  //     });
  // }
  //
  // public addDropDivEvents(element, highlightclass?){
  //   // Adds an event listener that prevents the default behaviors for the listed events.
  //   ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
  //     element.addEventListener(eventName, preventDefaults, false);
  //   });
  //
  //   function preventDefaults (e) {
  //     e.preventDefault();
  //     e.stopPropagation();
  //   }
  //
  //   // Adds the highlight class when events are fired.
  //   ['dragenter', 'dragover'].forEach(eventName => {
  //     element.addEventListener(eventName, highlight, false);
  //   });
  //
  //   function highlight(e) {
  //     //element.classList.add('highlight');
  //     element.classList.add(highlightclass);
  //   }
  //
  //   // Removes the highlight class when events are fired.
  //   ['dragleave', 'drop'].forEach(eventName => {
  //     element.addEventListener(eventName, unhighlight, false);
  //   });
  //
  //   function unhighlight(e) {
  //     element.classList.remove(highlightclass);
  //   }
  // }

}
//
// function makeQueryString(queryobj: object) {
//   const template = '?$';
//   const result = [];
//   Object.keys(queryobj).forEach(key => {
//     const tempstring = key + '=' + queryobj[key];
//     result.push(tempstring);
//   });
//   return template + result.join('&$');
// }
