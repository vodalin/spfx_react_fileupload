import * as React from 'react';
import styles from './filetile.module.scss';
import {Miniselector} from "./miniselector/miniselector";
import {IFieldData} from "../../../../../lib/webparts/fileuploader/components/filetile/filetile";


export interface IFileTileProps {
  file: any;
  fieldschema: Array<any>;  // {
                            //   Author: {},
                            //   MNMCorrespondent: {
                            //      value:[{Id: 1, Title: bank1}, {Id: 2, Title: bank2}]
                            //   }
                            // }
  getFieldData: Function;
}

export interface IFieldData{
  FieldName: string;
  Id: number;
  Text: string;
}

export class Filetile extends React.Component<IFileTileProps> {
  public mainfile = this.props.file;
  constructor(props){
    super(props);
    this.state = {
      //File: this.props.file,
      FieldObjects: {},
    };
    this.handleChange = this.handleChange.bind(this);
    this.textChange = this.textChange.bind(this);
  }

  public handleChange(child_data: IFieldData) {
    let fobjProperty = this.state['FieldObjects'];
    fobjProperty[child_data.FieldName] = {Id: child_data.Id, Text: child_data.Text};
    this.setState({fobjProperty});
  }

  public textChange(event){
    let newtext = event.target.value;
    let fieldname = event.target.accessKey;
    let fobjProperty = this.state['FieldObjects'];
    fobjProperty[fieldname] = {Id: undefined, Text: newtext};
    this.setState({fobjProperty});
  }

  public makeInfoRow() {
    let rowschema = this.props.fieldschema;
    let infoColumns = [<td>{this.mainfile['name']}</td>];

    Object.keys(rowschema).sort().forEach((key,index) =>{
      let column_element = undefined;
      if(Object.keys(rowschema[key]).length == 0) {
        column_element = (
          <td key={index.toString()}>
            <input data-fieldname={key} onChange={this.textChange} accessKey={key}/>
          </td>);
      }
      else{
        column_element = (
          <td key={index.toString()}>
            <Miniselector
              fieldname={key}
              options={rowschema[key]['value']}
              parentcallback = {this.handleChange}
            />
          </td>);
      }
      infoColumns.push(column_element);
    });
    return infoColumns;
  }

  public componentDidUpdate(){
    this.props.getFieldData({[this.mainfile['name']]: this.state['FieldObjects']});
  }

  // public render(): React.ReactElement<IFileTileProps>{
  //   return (
  //     <tr className={styles.ftile}>
  //       <div>
  //         <span>{JSON.stringify(this.state['FieldObjects'])}</span><br/>
  //         {this.makeInfoRow()}
  //       </div>
  //     </tr>
  //   );
  // }

  public render(): React.ReactElement<IFileTileProps>{
    return (<tr className={styles.ftile}>{this.makeInfoRow()}</tr>);
  }
}
