import * as React from 'react';
import styles from './Fileuploader.module.scss';
import { Filetile } from "./filetile/filetile";
import IFileuploaderProps from "./IFileuploaderProps";

import {OperatorService} from "../../../services/operator.service";
import {IFieldData} from "../../../../lib/webparts/fileuploader/components/filetile/filetile";
import {ObjectIterator} from "lodash";

export interface ISubmit_Data{
  filedata: any;
  fieldpayload: IFieldData;
}

export default class RctUploader extends React.Component<IFileuploaderProps, {}> {
  public dropDiv: HTMLElement;
  public os: OperatorService;
  public RootFolder: string;
  public max_file_amount = 100;

  constructor(props) {
    super(props);
    this.state = {
      filetile_list: [],
      rootfolder: '',
      submit_data: {},  // {
                        //  File1.doc: {raw_file:{}, col1:{id:1,Text:"wah"}, col2:{id:undefined,Text:"hah"},
                        //  File2.csv: {raw_file:{}, col1:{id:42,Text:"some"}, col2:{id:99,Text:"thing"}
                        // }
      runningUpload: false,
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

  public componentDidUpdate(){
    //console.log(this.props.required_fields, this.props.required_fields_schema);
  }

  public render(): React.ReactElement<IFileuploaderProps> {
    return (
      <div className={styles.rctUploader}>
        {
          (this.props.target_library != undefined) ? this.getUploaderTemplate() : this.getPromptTemplate()
        }
      </div>
    );
  }


  /************* Helper functions *************/
  public handleReset(){
    this.setState({submit_data: {}, filetile_list: []});
  }

  public handleSubmit(){
    /* Action for submit button; starts uploading all files in <state.submit_data> */
    let target_library = this.props.target_library;
    let submit_data = this.state['submit_data'];
    let target_folder = this.state['rootfolder'];
    let allPr = this.os.startUploads(submit_data, target_folder, target_library);

    this.setState({runningUpload: true});
    Promise.resolve(allPr)
      .then(val =>{
        this.setState({runningUpload: false});
      });
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
    /* Process files dropped into this.dropdiv */
    let tiles: Array<any> = this.state['filetile_list'];
    let evt_files = event.dataTransfer.files;
    let filelist = Object.keys(evt_files).map(key => {
      return evt_files[key];
    });

    try {
      if(this.state['rootfolder'] != null){
        filelist.forEach(file => {
          if((tiles.length + 1) <= this.max_file_amount){
            this.addToSubmitData(file);
            this.addFileTile(file);
          }
          else {
            throw new EvalError('Exceeded file limit: ' + this.max_file_amount);
          }
        });
      }
      else{
        throw new EvalError('Missing ?RootFolder query parameter.');
      }
    } catch (e){
      (e instanceof EvalError) ? alert(e.message) : console.log(e);
    }
  }

  private addToSubmitData(file){
    /* Adds to or updates the <submit_data> state variable with <file> */
    let filename = file['name'];
    let sub_data = this.state['submit_data'];
    let sub_item = sub_data[filename];
    (sub_item == undefined) ? sub_data[filename] = {'raw_file': file} : sub_item['raw_file'] = file;

    this.setState({'submit_data': sub_data});
  }

  private addFileTile(file){
    /* Checks the key(filename) of each tile in <state[filetile_list]>
    *  If a key already exists, then do not add another tile.
    * */
    let filename = file['name'];
    let tile_list = this.state['filetile_list'];
    let matching_tiles = tile_list.filter(tileObj => {
      if(tileObj['key'] == filename){ return tileObj; }
    });

    if(matching_tiles.length == 0){
      let newTile = (
        <Filetile
          key = {filename}
          file={file}
          fieldschema={this.props.required_fields_schema}
          getFieldData={this.getFieldData.bind(this)}
        />
      );
      tile_list.push(newTile);

      this.setState({'filetile_list': tile_list});
    }
  }

  /************* Template functions *************/
  private getUploaderTemplate() {
    /* Generates main template */
    return (
      <div>
        <span>Uploading to: {this.props.target_library + '/' + this.state['rootfolder']}</span>
        <br/>
        <div className={'DropDiv'}>
          <div ref={elem => this.dropDiv = elem} className={styles.uploadbin} onDrop={this.handleDrop.bind(this)}>
            <p className={styles.droptext}>Drop Files Here!</p>
          </div>
        </div>
        {
          this.state['runningUpload'] ?
            <button className={styles.loading} disabled>Working on it....</button>
            :
            <button className={styles.submitBtn} onClick={this.handleSubmit.bind(this)}>SUBMIT</button>
        }
        <div className='filelist'>
          <table>
            {this.makeHeaders()}
            {this.state['filetile_list']}
          </table>
        </div>
        <button className={styles.reset} onClick={this.handleReset.bind(this)}>RESET</button>
      </div>
    )
  }

  private getPromptTemplate(){
    /* Generates template used when no library is selected */
    return (<div>Please select a target library.</div>)
  }
}


/************* Drop div default event handlers *************/
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
