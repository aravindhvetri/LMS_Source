import * as React from 'react';
//import styles from './Lms.module.scss';
import type { ILmsProps } from './ILmsProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import {ISPFXContext, sp } from "@pnp/pnpjs";
import MainComponent from './MainComponent/MainComponent';

export default class Lms extends React.Component<ILmsProps, {}> {
  constructor(prop: ILmsProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context as unknown as ISPFXContext,
    });
  }
  public render(): React.ReactElement<ILmsProps> {
    
    return <div><MainComponent context={this.props.context}/></div>
    
  }
}
