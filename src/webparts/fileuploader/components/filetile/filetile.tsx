import * as React from 'react';
import styles from './filetile.module.scss';
import {Miniselector} from "./miniselector/miniselector";
import {IFieldData} from "../../../../../lib/webparts/fileuploader/components/filetile/filetile";

export interface IFileTileProps {
  file: any;
  fieldschema: Array<any>;  // {
                            //   Column1: {},
                            //   Column2: {
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
      FieldObjects: {},
    };
    this.handleChange = this.handleChange.bind(this);
  }

  public handleChange(child_data: IFieldData) {
    let fobjProperty = this.state['FieldObjects'];
    fobjProperty[child_data.FieldName] = {Id: child_data.Id, Text: child_data.Text};
    this.setState({fobjProperty});
  }

  public makeInfoRow() {
    let fieldschema = this.props.fieldschema;
    let infoColumns = [<td>{this.mainfile['name']}</td>];

    Object.keys(fieldschema).sort().forEach((key,index) =>{
      let column_element = (
        <td key={index.toString()}>
          <Miniselector
            fieldname={key}
            options={fieldschema[key]['value']}
            parentcallback = {this.handleChange}
            columndata={fieldschema[key]}
          />
        </td>);

      infoColumns.push(column_element);
    });
    return infoColumns;
  }

  public componentDidUpdate(){
    this.props.getFieldData({[this.mainfile['name']]: this.state['FieldObjects']});
  }

  public render(): React.ReactElement<IFileTileProps>{
    return (<tr className={styles.ftile}>{this.makeInfoRow()}</tr>);
  }
}
