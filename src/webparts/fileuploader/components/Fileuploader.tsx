import * as React from 'react';
import styles from './Fileuploader.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { Filetile } from "./filetile/filetile";
import {IFileuploaderProps} from "./IFileuploaderProps";

//**
import {OperatorService} from "../../../services/operator.service";
import {IFieldData} from "../../../../lib/webparts/fileuploader/components/filetile/filetile";

export interface ISubmit_Data{
  filedata: any;
  fieldpayload: IFieldData;
}

export default class RctUploader extends React.Component<IFileuploaderProps, {}> {
  public dropDiv: HTMLElement;
  public os: OperatorService;
  public RootFolder: string;
  public max_file_amount = 2;

  constructor(props) {
    super(props);
    this.state = {
      filetile_list: [],
      rootfolder: '',
      submit_data: {},  // {
                        //  File1.doc: {raw_file:{}, col1:{id:1,Text:"wah"}, col2:{id:undefined,Text:"hah"},
                        //  File2.csv: {raw_file:{}, col1:{id:42,Text:"some"}, col2:{id:99,Text:"thing"}
                        // }
    };
    this.os = new OperatorService(window['webPartContext']);
    this.handleDrop.bind(this);
    this.handleSubmit.bind(this);
  }

  public componentDidMount() {
    const url = new URL(window.location.href);
    this.setState({rootfolder: url.searchParams.get('RootFolder')});

    if(this.dropDiv != undefined){
      addDropDivEvents(this.dropDiv, styles.highlight);
    }
  }

  //*******************************
  public handleSubmit(){
    let submit_data = this.state['submit_data'];
    let target_library = this.props.target_library;
    //let target_folder = target_library + '/' + this.state['rootfolder'];
    let target_folder = this.state['rootfolder'];
    //let target_folder = this.props.target_library;
    this.os.startUploads(submit_data, target_folder, target_library);

    // let upload_queue = [];
    // Object.keys(submit_data).forEach(filekey => {
    //   let datapayload = submit_data[filekey];
    //   let raw_file = undefined;
    //   let columndata = [];
    //   Object.keys(datapayload).forEach(datakey => {
    //     let cur_item = datapayload[datakey];
    //     if(datakey == 'raw_file'){
    //       raw_file = cur_item;
    //     }
    //     else{
    //       columndata.push(cur_item)
    //     }
    //   });
    //   console.log(raw_file, columndata);
    // });
  }

  public getFieldData(child_data) {
    let filename = Object.keys(child_data)[0];
    let field_data = child_data[filename];
    let state_data = this.state['submit_data'];
    state_data[filename] = {...state_data[filename], ...field_data};

    this.setState({'submit_data': state_data});
  }

  public makeHeaders() {
    let headers = [<th>Title</th>];
    Object.keys(this.props.required_fields_schema).sort().forEach((key,index) => {
      headers.push(<th key={index.toString()}>{key}</th>);
    });
    let header_row = (<tr>{headers}</tr>);
    return header_row;
  }

  public handleDrop(event){
    try{
      // Error checking after dropping files.
      if(this.state['rootfolder'] != null) {
        event.stopPropagation();
        const files = event.dataTransfer.files;
        let currentTiles: Array<any> = this.state['filetile_list'];
        let newTileList: Array<any> = [];

        for(let i=0; i < files.length; i++){
          let current_file = files[i];
          if((currentTiles.length + newTileList.length) >= this.max_file_amount){
            this.setState({filetile_list: currentTiles.concat(newTileList)});
            throw EvalError('Exceeded file limit: ' + this.max_file_amount);
          }
          else{
            //###
            let subdata = this.state['submit_data'];
            subdata[current_file['name']] = {raw_file: current_file};

            newTileList.push((
              <Filetile
                file={current_file}
                fieldschema={this.props.required_fields_schema}
                getFieldData={this.getFieldData.bind(this)}
              />
            ));
          }
          this.setState({filetile_list: currentTiles.concat(newTileList)});
        }
      }
      else {
        throw EvalError('Rootfolder is empty.');
      }
    }
    catch (e) {
      if(e instanceof EvalError){
        alert(e.message);
      }
      else{
        alert('Error occured.');
        console.log(e);
      }
    }
  }

  public componentDidUpdate(){
    //console.log(this.state['submit_data']);
  }

  public render(): React.ReactElement<IFileuploaderProps> {
    return (
      <div className={styles.rctUploader}>
        {/*{Object.keys(this.state['submit_data']).length != 0 && <span>{JSON.stringify(this.state['submit_data'])}</span>}*/}
        {
          this.props.target_library != undefined ?
            <div>
              <span>Uploading to: {this.props.target_library + '/' + this.state['rootfolder']}</span>
              <br/>
              <div ref={elem => this.dropDiv = elem} className={styles.uploadbin} onDrop={this.handleDrop.bind(this)}>
                <p className={styles.droptext}>Drop Files Here!</p>
              </div>
              <button className={styles.submitBtn} onClick={this.handleSubmit.bind(this)}>SUBMIT</button>
              <div className='filelist'>
                <table>
                  {this.makeHeaders()}
                  {this.state['filetile_list']}
                </table>
              </div>
            </div> : <div>Please select a target library.</div>
        }
      </div>
    );
  }

}

function addDropDivEvents(element, highlightclass?){
  // Adds an event listener that prevents the default behaviors for the listed events.
  ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    element.addEventListener(eventName, preventDefaults, false);
  });

  function preventDefaults (e) {
    e.preventDefault();
  }

  // Adds the highlight class when events are fired.
  ['dragenter', 'dragover'].forEach(eventName => {
    element.addEventListener(eventName, highlight, false);
  });

  function highlight(e) {
    //element.classList.add('highlight');
    element.classList.add(highlightclass);
  }

  // Removes the highlight class when events are fired.
  ['dragleave', 'drop'].forEach(eventName => {
    element.addEventListener(eventName, unhighlight, false);
  });

  function unhighlight(e) {
    element.classList.remove(highlightclass);
  }
}
