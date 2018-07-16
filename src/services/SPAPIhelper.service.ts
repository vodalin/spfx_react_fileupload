import {HttpcallerService} from './httpcaller.service';

import {
  IDigestCache,
  DigestCache,
  ISPHttpClientOptions
} from '@microsoft/sp-http';
import * as $ from 'jquery';

export class SPAPIhelperService {
  private curContext: any;
  private cs: HttpcallerService;
  private siteurl: string;
  private subsiteRef: string;
  private baseURL: string;
  private urlSegments: string[];

  private byTitleSect = '/lists/getByTitle(\'%\')';
  private byServerRelSect = '/getfolderbyserverrelativeurl(\'%\')';

  constructor(context) {
    this.curContext = context;
    this.cs = new HttpcallerService(this.curContext);
    this.siteurl = this.curContext.pageContext.web.absoluteUrl; // https://xoom.sharepoint.com/sites/devsite
    this.subsiteRef = this.curContext.pageContext.web.serverRelativeUrl; // /sites/devsite
    this.baseURL = this.siteurl + '/_api/web'; // https://xoom.sharepoint.com/sites/devsite/_api/web
  }

  // Functions that interact with API
  public getSPData(resourcePath: string, filterObj?: object){
    this.urlSegments = [resourcePath];
    if(filterObj !== undefined){
      this.urlSegments.push(this.makeQueryString(filterObj));
    }
    const callUrl = this.buildUrl(this.urlSegments);
    return this.makeCall(this.cs.getCall(callUrl), 'getSPResource() failed');
  }

  // Write item ID to a specific column for a file. Assuming target column is a SP picklist.
  public setItemColumn(target_library: string, fileId, coldata: Object, ) {
    this.urlSegments = [this.byTitleSect.replace('%', target_library), '/items', '(\'' + fileId + '\')'];
    const callUrl = this.buildUrl(this.urlSegments);
    const opt: ISPHttpClientOptions = {
      headers:{
        'IF-MATCH': '*',
        'content-type': 'application/json',
        'X-HTTP-METHOD': 'MERGE'
      },
      body: JSON.stringify(coldata)
    };
    return this.makeCall(this.cs.postCall(callUrl, opt), 'setItemColumn() failed');
  }

  public uploadFiles(raw_file: any, target_folder: string) {
    let stripped_file_name = raw_file.name.replace(/['!&*?=\/|\\":<>]/g,'');
    this.urlSegments = [
      this.byServerRelSect.replace('%', target_folder),
      '/files/add(overwrite=true,url=\'' + stripped_file_name + '\')'
    ];
    const callUrl = this.buildUrl(this.urlSegments);
    const digestCache: IDigestCache = this.curContext.serviceScope.consume(DigestCache.serviceKey);
    return Promise.all([
      this.getFileBuffer(raw_file),
      digestCache.fetchDigest(this.subsiteRef)
    ])
      .then(results => {
        const digest = results[1];
        const fbuffer = results[0];
        const header = {
          'accept': 'application/json;odata=verbose',
          'X-RequestDigest': digest,
          'content-length': fbuffer.byteLength
        };
        //console.log(callUrl, fbuffer, header);
        return this.cs.ajaxPost(callUrl, fbuffer, header);
      })
      .catch(err =>{
        throw new Error('uploadFiles() failed. \n' + err.stack);
      });
  }

  // URL building functions.
  private buildUrl(segments: string[]): string {
    let fullurl = this.baseURL;
    segments.forEach(segment =>{
      fullurl = fullurl + segment;
    });
    return fullurl;
  }

  private makeQueryString(queryobj: object) {
    const template = '?$';
    const result = [];
    Object.keys(queryobj).forEach(key => {
      const tempstring = key + '=' + queryobj[key];
      result.push(tempstring);
    });
    return template + result.join('&$');
  }

  // MISC functions
  private getFileBuffer(file){
    const deferred = $.Deferred();
    const reader = new FileReader();
    reader.onloadend = (e) => {
      deferred.resolve(e.target['result']);
    };
    reader.readAsArrayBuffer(file);
    return deferred.promise();
  }

  private makeCall(callfn?: Promise<Object | void>, errorstring?: string) {
    return Promise.resolve(callfn)
      .then(result => {
        return result;
      })
      .catch(err => {
        throw new Error( errorstring + '\n' + err.stack);
      });
  }
}

