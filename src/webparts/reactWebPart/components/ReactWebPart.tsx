import * as React from 'react';
import styles from './ReactWebPart.module.scss';
import { IReactWebPartProps } from './IReactWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class ReactWebPart extends React.Component<IReactWebPartProps, {name:any,email:any}> {



  public componentDidMount(): void{
    this.props.graphClient
    .api('me')
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
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!! Graph: {this.state.name}  </span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
                
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
