import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './GraphCallTwoWebPart.module.scss';
import pnp from 'sp-pnp-js';
import * as strings from 'GraphCallTwoWebPartStrings';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { GraphError } from '@microsoft/microsoft-graph-client';

export interface IGraphCallTwoWebPartProps {
  description: string;
}

export default class GraphCallTwoWebPart extends BaseClientSideWebPart<IGraphCallTwoWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.graphCallTwo }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
      let usrMail = this.context.pageContext.user.email;
      let usrPrincipalName = this.context.pageContext.user.loginName;
      let foobar: any = {
        "securityEnabledOnly": false
      };
      this.getUserId(usrMail).then(usrIdResult => {
          this.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
                        
            /* POST /groups/{id}/getMemberGroups */ 
            
            //  client.api('/groups/{ '+ usrIdResult + ' |  ' + usrPrincipalName + '}/getMemberGroups').post((error : Error , response : Response) => {
            
            //  client.api('/groups/{ '+ usrIdResult + '}/getMemberGroups').post((error : Error , response : Response) => {
            
            //  client.api('/me').get((error : GraphError, response : Response) => {
            
            //  client.api('/me/getMemberGroups').post((error : Error , response : Response) => {
            
            //  client.api('/users/{'+ usrIdResult +' | '+ usrPrincipalName +'}/memberOf').get((error : GraphError , response : Response) => {  
            
              client.api('/me/getMemberGroups').post(foobar)
              .then((content: any) => {
                  console.log("Groups ", content);
            }).catch((err)=>{
              console.log("err ", err);
            })
        });
      });
    }

  public getUserId(email: string): Promise<number> {
    return pnp.sp.site.rootWeb.ensureUser(email).then(result => {
        return result.data.Id;
    });
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
