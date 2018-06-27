import * as $ from 'jquery';

import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from '@microsoft/sp-http';

export class HttpcallerService {
  private curContext: any;

  constructor(context) {
    this.curContext = context;
  }

  public getCall(url: string) {
    return Promise.resolve(this.curContext.spHttpClient.get(url, SPHttpClient.configurations.v1))
      .then((reply: SPHttpClientResponse) => {
        return reply.json();
      })
      .then((finaldata: Object) => {
        return finaldata;
      })
      .catch (err => {
        throw new Error('getCall() failed. \n' + err);
      });
  }

  public postCall(url: string, opt: ISPHttpClientOptions) {
    return Promise.resolve(this.curContext.spHttpClient.post(url, SPHttpClient.configurations.v1, opt))
      .then((reply: SPHttpClientResponse) => {
        return reply;
      })
      .catch (err => {
        throw new Error('postCall() failed. \n' + err);
      });
  }

  public ajaxPost(url: string, data: any, headers: {}) {
    const ajcall = $.ajax({
      url: url,
      type: 'POST',
      data: data,
      processData: false,
      headers: headers
    });
    return Promise.resolve(ajcall)
      .then(val =>{
        console.log('ajaxPost() done.');
      })
      .catch(err => {
        throw new Error('ajaxPost() failed. \n' + err);
      });
  }
}
