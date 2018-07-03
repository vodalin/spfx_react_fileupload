import * as React from 'react';
import styles from './miniselector.module.scss';
import classNames from 'classnames/bind';
//***********
export interface IMiniSelectProps {
  options: Array<any>;
  fieldname: string;
  parentcallback: Function;
}

import {IFieldData} from "../filetile";

export class Miniselector extends React.Component<IMiniSelectProps> {
  public textInput: HTMLElement;
  public optionDiv: HTMLElement;
  public minidiv: HTMLElement;
  public classlist: Array<any>;
  public classObj: Object;
  public cx: any;

  constructor(props){
    super(props);
    this.cx = classNames.bind(styles);
    this.state = {
      isValid: false,
      searchtext: '',
      matches: this.props.options,
      classObj: {
        hidden: true,
        revealed: true,
      },
    };

    this.hideDiv = this.hideDiv.bind(this);
    this.showDiv = this.showDiv.bind(this);
    this.isSelected = this.isSelected.bind(this);
    this.unSelected = this.unSelected.bind(this);
    this.isClicked = this.isClicked.bind(this);
    this.txtChanged = this.txtChanged.bind(this);
  }

  public async hideDiv() {
    await this.sleep(200);
    this.addClass('classObj','hidden');
  }

  public showDiv() {
    this.deleteClass('classObj','hidden');
  }

  public isSelected(event) {
    event.target['className'] = styles.selected;
  }

  public unSelected(event) {
    event.target['className'] = '';
  }

  //****************
  public isClicked(event){
    let text = event.target.innerText;
    let inputdata = this.props.options;
    let fielddata: IFieldData = {
      FieldName: this.props.fieldname,
      Id: 0,
      Text: ''
    };
    let re_spec_characters = new RegExp('(?=[&^$()])','ig');
    let fieldtext = text.replace(re_spec_characters, '\\');
    let re = new RegExp('^' + fieldtext + '$', 'g');
    let matchlist = inputdata.filter(item => { return item['Title'].match(re); });
    if(matchlist.length !== 0){
      this.setState({isValid: true});
      fielddata.Id = matchlist[0]['Id'];
      fielddata.Text =  matchlist[0]['Title'];
      this.props.parentcallback(fielddata);
    }
    else {
      this.setState({isValid: false});
    }
    this.setState({searchtext: text});
  }

  public txtChanged(event){
    this.verifyInput(event.target.value);
  }

  public sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve,ms));
  }

  public verifyInput(text){
    let inputdata = this.props.options;
    let itemId = undefined;
    let itemText = text;
    let fielddata: IFieldData = {
      FieldName: this.props.fieldname,
      Id: 0,
      Text: ''
    };

    //This bit escapes special characters and defines the escape sequence to find the text.
    let re_spec_characters = new RegExp('(?=[&^$()])','ig');
    let fieldtext = text.replace(re_spec_characters, '\\');
    let re = new RegExp(fieldtext,'i');

    let matchlist = inputdata.filter(item => { return item['Title'].match(re); });
    (matchlist.length === 0) ? this.setState({matches: inputdata}) : this.setState({matches: matchlist});
    //(matchlist.length === 1 && matchlist[0]['Title']) === text ? this.setState({isValid: true}) : this.setState({isValid: false});
    if(matchlist.length === 1 && matchlist[0]['Title'] === text) {
      this.setState({isValid: true});
      fielddata.Id = matchlist[0]['Id'];
      fielddata.Text = text;
      this.props.parentcallback(fielddata);
      //this.props.parentcallback({fieldName: this.props.fieldname, Id: itemId, Text: itemText});
    }
    else{
      this.setState({isValid: false});
    }
    this.setState({searchtext: text});
  }


  public makeDropDiv() {
    let optionRows = [];
    let propOptions = this.state['matches'].sort(this.compare);
    propOptions.forEach((option, index) => {
      optionRows.push(
        <tr>
          <td>
            <div key={option['Id']}
                 onMouseEnter={this.isSelected}
                 onMouseLeave={this.unSelected}
                 onClick={this.isClicked}
            >
              <span className={'object_title'}>{option['Title']}</span>
            </div>
          </td>
        </tr>);
    });
    return optionRows;
  }


  public addClass(classobject: string, classname: string){
    let class_object = this.state[classobject];
    if((Object.keys(class_object)).indexOf(classname) == -1){
      class_object[classname] = true;
      this.setState({[classobject]: class_object});
    }
  }

  public deleteClass(classobject: string, classname: string){
    let class_object = this.state[classobject];
    if((Object.keys(class_object)).indexOf(classname) != -1){
      delete class_object[classname];
      this.setState({[classobject]: class_object});
    }
  }

  public compare(a, b) {
    // During makeDropDiv(), use this to alphabetize all choices by their titles.
    const TitleA = a['Title'].toLowerCase();
    const TitleB = b['Title'].toLowerCase();

    let comparison = 0;
    if (TitleA > TitleB) {
      comparison = 1;
    } else if (TitleA < TitleB) {
      comparison = -1;
    }
    return comparison;
  }


  public render(): React.ReactElement<IMiniSelectProps>{
    let className = this.cx(this.state['classObj']);
    return (
      <div className={styles.minselect}>
        <div className={styles.restricted_div}>
          <input ref={(elem) => this.textInput = elem}
                 onFocus={this.showDiv} onBlur={this.hideDiv}
                 onChange={this.txtChanged}
                 value={this.state['searchtext']}
                 className = {this.state['isValid'] ? styles.isValid : styles.notValid}
          />
          <div ref={(elem) => this.optionDiv = elem} className={className}>
            <table>
              <tbody>
                {this.makeDropDiv()}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  }

}
