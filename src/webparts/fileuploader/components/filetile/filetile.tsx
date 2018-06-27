import * as React from 'react';
import styles from './filetile.module.scss';
import {Miniselector} from "../miniselector/miniselector";


export interface IFileTileProps {
  file: any;
  fieldschema: {};
}

export class Filetile extends React.Component<IFileTileProps> {
  constructor(props){
    super(props);
    this.state = {
      columList: []
    };
  }

  public makeInfoRow() {
    let cur_file = this.props.file;
    let rowschema = this.props.fieldschema;
    let infoColumns = [<td>{cur_file['name']}</td>];
    Object.keys(rowschema).sort().forEach((key,index) =>{
      let column_element = undefined;
      if(Object.keys(rowschema[key]).length == 0) {
        column_element = (<td key={index.toString()}><input /></td>);
      }
      else{
        column_element = (<td key={index.toString()}><Miniselector options={rowschema[key]['value']}/></td>);
      }
      infoColumns.push(column_element);
    });
    return infoColumns;
  }

  public render(): React.ReactElement<IFileTileProps>{
    return (<tr className={styles.ftile}>{this.makeInfoRow()}</tr>);
  }
}
