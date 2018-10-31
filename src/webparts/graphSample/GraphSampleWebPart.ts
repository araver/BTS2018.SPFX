import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GraphSampleWebPart.module.scss';
import * as strings from 'GraphSampleWebPartStrings';
import * as jquery from "jquery";

import { BaseDialog, IDialogConfiguration, Dialog } from '@microsoft/sp-dialog';
import { UserDialog } from './dialog/UserDialog';

export interface IGraphSampleWebPartProps {
  description: string;
}

export default class GraphSampleWebPart extends BaseClientSideWebPart<IGraphSampleWebPartProps> {
  private users = [];
  public render(): void {
    this.domElement.innerHTML = `Loading...`;

    if(!this.renderedOnce){
      jquery(this.domElement).on('click', '.user-row', (event, element)=>{
        console.log(event);
        var userId = jquery(event.currentTarget).data().user;
        var user = this.users.filter(i => i.id === userId)[0];
        if(user){
          UserDialog.Show(user);
        }
      });
    }

    this._render();
  }

  public async _render(){
    var gClient = await this.context.msGraphClientFactory.getClient();

    var results = await gClient.api("/users").get();
    var resultsArray = [];
    resultsArray = resultsArray.concat(results.value);
    while(results["@odata.nextLink"]){
      results = await gClient.api(results["@odata.nextLink"]).get();
      resultsArray = resultsArray.concat(results.value);
    }
    this.users = resultsArray;
    console.log("REsults", resultsArray);

    var usersHtmlArr = resultsArray.map((value) =>{
      return `<div class="user-row" title="${value.id}" data-user="${value.id}">
    <div>${value.displayName}</div>
    <div>${value.mail || ""}</div>
</div>`;
    });
    this.domElement.innerHTML = `<div class="${styles.graphSample}">
      <div class="user-row header-row">
        <div>Name</div>
        <div>Email</div>
      </div>
      ${usersHtmlArr.join("\n")}
    </div>`;


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
