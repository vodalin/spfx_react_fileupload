import * as React from 'react';
import styles from './Fileuploader.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { Filetile } from "./filetile/filetile";
import {IFileuploaderProps} from "./IFileuploaderProps";

//**
import {OperatorService} from "../../../services/operator.service";

export default class RctUploader extends React.Component<IFileuploaderProps, {}> {
  public dropDiv: HTMLElement;
  public os: OperatorService;
  public RootFolder: string;
  public max_file_amount = 2;
  constructor(props) {
    super(props);
    this.state = {
      filetile_list: [],
      rootfolder: ''
    };
    this.os = new OperatorService(window['webPartContext']);
    this.handleDrop.bind(this);
  }

  public componentDidMount() {
    const url = new URL(window.location.href);
    this.setState({rootfolder: url.searchParams.get('RootFolder')});

    if(this.dropDiv != undefined){
      addDropDivEvents(this.dropDiv, styles.highlight);
    }
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

        Object.keys(files).sort().forEach((key: string, index) => {
          if((currentTiles.length + newTileList.length) >= this.max_file_amount){
            this.setState({filetile_list: currentTiles.concat(newTileList)});
            throw EvalError('Exceeded file limit: ' + this.max_file_amount);
          }
          else{
            newTileList.push((
              <Filetile file={files[key]} fieldschema={this.props.required_fields_schema}/>
            ));
          }
        });
        this.setState({filetile_list: currentTiles.concat(newTileList)});
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

  public render(): React.ReactElement<IFileuploaderProps> {
    // console.log(this.props.target_library);
    // console.log(this.props.required_fields);
    // console.log(this.props.required_fields_metadata);
    // console.log(this.props.required_fields_schema);
    return (
      <div className={styles.rctUploader}>
        {
          this.props.target_library != undefined ?
            <div>
              <span>Uploading to: {this.props.target_library + '/' + this.state['rootfolder']}</span>
              <br/>
              <div ref={elem => this.dropDiv = elem} className={styles.uploadbin} onDrop={this.handleDrop.bind(this)}>
                <p className={styles.droptext}>Drop Files Here!</p>
              </div>
              <button className={styles.submitBtn}>SUBMIT</button>
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

  //Helper functions

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
