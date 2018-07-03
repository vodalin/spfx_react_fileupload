import * as React from 'react';
import styles from './textbox.module.scss';
import classNames from 'classnames/bind';
//***********
export interface IMiniSelectProps {
  options: Array<any>;
  fieldname: string;
}

export class Textbox extends React.Component<IMiniSelectProps> {

  public render(): React.ReactElement<IMiniSelectProps>{
    return (
      <div>
        <input/>
      </div>
    );
  }
}
