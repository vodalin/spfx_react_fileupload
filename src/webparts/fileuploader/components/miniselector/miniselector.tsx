import * as React from 'react';
import styles from './miniselector.module.scss';
import classNames from 'classnames/bind';
//***********
export interface IMiniSelectProps {
  options: Array<any>;
}

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
      columList: [],
      div_class_str: 'styles.hidden',
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
    console.log('ah');
  }

  public sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve,ms));
  }

  public componentDidMount() {
    // this.textInput.addEventListener('focus', () => {
    //   this.deleteClass('classObj','hidden');
    //   this.addClass('classObj','revealed');
    // });
    // this.textInput.addEventListener('blur', () => {
    //   this.addClass('classObj','hidden');
    //   this.deleteClass('classObj','revealed');
    // });
    //
    // this.minidiv.addEventListener('mouseenter', () => {console.log('tag!');});
    // this.optionDiv.addEventListener('mouseleave', () => {});
  }

  public makeDropDiv() {
    let propOptions = this.props.options;
    let optionRows = [];
    propOptions.forEach((option, index) => {
      if(index < 20){
        optionRows.push(
          <tr>
            <td>
              {/*<div ref={(elem) => this.addDDEvents(elem)} key={option['Id']} >*/}
                {/*<span>{option['Title']}</span>*/}
              {/*</div>*/}
              <div key={option['Id']}
                   onMouseEnter={this.isSelected}
                   onMouseLeave={this.unSelected}
                   onClick={this.isClicked}
              >
                <span className={'object_title'}>{option['Title']}</span><span className={'object_id'} style={{display:'none'}}>{option['Id']}</span>
              </div>
            </td>
          </tr>);
      }
    });
    return optionRows;
  }

  public testButtons() {
    return(<tr><td><button onClick={this.isClicked}>TAP</button></td></tr>);
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

  public addDDEvents(e: HTMLElement){
    if(e){
      e.addEventListener('mouseenter', ()=> { e.className = styles.selected; });
      e.addEventListener('mouseleave', ()=> { e.className = ''; });
      e.addEventListener('click', ()=> { console.log('ah'); });
    }
  }

  public render(): React.ReactElement<IMiniSelectProps>{
    let className = this.cx(this.state['classObj']);
    return (
      <div className={styles.minselect}>
        <div className={styles.restricted_div}>
          <input ref={(elem) => this.textInput = elem} onFocus={this.showDiv} onBlur={this.hideDiv} type="text"/>
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
