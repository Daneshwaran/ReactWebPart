import * as React from 'react';
import styles from './ReactWebPart.module.scss';
import { IReactWebPartProps } from './IReactWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class ReactWebPart extends React.Component<IReactWebPartProps, {name:any,email:any}> {



  public componentDidMount(): void{
    this.props.graphClient
    .api("https://graph.microsoft.com/v1.0/sites/danesh96.sharepoint.com,1e8f08be-d4db-43d8-a398-198099a9378b,a89d96e0-bb32-4cc1-a664-2d63c703214b/drive/root:/new.xlsx:/workbook/tables('1')/rows")
    .get((error:any,user:any,rawResponse?:any) => {
      this.setState({
        name: user.displayName,
        email:user.mail
      });
    });
  }




  public render(): React.ReactElement<IReactWebPartProps> {
    return (
      <div className={ styles.reactWebPart }>

      </div>
    );
  }
}
