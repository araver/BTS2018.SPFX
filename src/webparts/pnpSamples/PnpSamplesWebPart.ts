import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PnpSamplesWebPart.module.scss';
import * as strings from 'PnpSamplesWebPartStrings';
import * as spfx from "@pnp/spfx-property-controls";
import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
import { PropertyFieldPeoplePicker, PrincipalType, IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { PropertyFieldDateTimePicker, DateConvention, TimeConvention, IDateTimeFieldValue,  } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';

export interface IPnpSamplesWebPartProps {
  description: string;
  listId: string;
  listIds: string[];
  color: string;
  backgroundColor: string;
  htmlCode: string;
  cssCode: string;
  jsCode: string;
  datetime: IDateTimeFieldValue;
  people: IPropertyFieldGroupOrPerson[];

}

export default class PnpSamplesWebPart extends BaseClientSideWebPart<IPnpSamplesWebPartProps> {
  public render(): void {
    console.log(this.properties);
    var ilStyles = [];
    if(this.properties.color){
      ilStyles.push("color:" + this.properties.color);
    }

    if(this.properties.backgroundColor){
      ilStyles.push("background-color:" + this.properties.backgroundColor);
    }

    var styleString = ilStyles.join(";");
    this.domElement.innerHTML = `<div class="${ styles.pnpSamples }">
  <div class="${ styles.container }">
    <div class="${ styles.row }" style="${styleString}">
      <div class="${styles.section}">
        <h4>List ID</h4>
        <span>${this.properties.listId || ""}</span>
      </div>
      <div class="${styles.section}">
        <h4>List IDs</h4>
        <span>${this.properties.listIds ? this.properties.listIds.join(", ") : ""}</span>
      </div>
      <div class="${styles.section}">
        <h4>Date Time</h4>
        <span>${this.properties.datetime ? this.properties.datetime.displayValue : ""}</span>
      </div>
      <div class="${styles.section}">
        <h4>People</h4>
        <span>${this.properties.people ? this.properties.people.map(i => i.fullName).join(", ") : ""}</span>
      </div class="${styles.section}">
      <div class="${styles.section}">
        <h4>Font Color</h4>
        <span>${this.properties.color ? this.properties.color : ""}</span>
      </div>
      <div class="${styles.section}">
        <h4>Background Color</h4>
        <span>${this.properties.backgroundColor ? this.properties.backgroundColor : ""}</span>
      </div>
      <div class="${styles.section}">
        <h4>Html</h4>
        <div>${this.properties.htmlCode || ""}</div>
      </div>
      <div class="${styles.section}">
        <h4>JavaScript</h4>
        <div class="${styles.pre}">${this.properties.jsCode || ""}</div>
      </div>

      <div>
        <h4>Css</h4>
        <style>${this.properties.cssCode}</style>
        <div class="${styles.pre}">${this.properties.cssCode || ""}</div>
      </div>
    </div>
  </div>
</div>`;

      if(this.properties.jsCode){
        eval(this.properties.jsCode);
      }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyChanged(){
    console.log(arguments);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: "Lists",
              isCollapsed: true,
              groupFields: [
                spfx.PropertyFieldListPicker("listId", {
                  context: this.context,
                  label: "List Picker 1",
                  selectedList: this.properties.listId,
                  multiSelect: false,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyChanged.bind(this),
                  key: "listId"
                }),
                spfx.PropertyFieldListPicker("listIds", {
                  context: this.context,
                  label: "List Picker 2",
                  selectedList: this.properties.listIds,
                  multiSelect: true,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyChanged.bind(this),
                  key: "listIds"
                }),



                
              ]
            },

            {
              groupName: "Miscellaneous",
              isCollapsed: true,
              groupFields: [
                PropertyFieldDateTimePicker('datetime', {
                  label: 'Select the date and time',
                  initialDate: this.properties.datetime,
                  dateConvention: DateConvention.DateTime,
                  timeConvention: TimeConvention.Hours12,
                  onPropertyChange: this.onPropertyChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'dateTimeFieldId'
                }),
                PropertyFieldPeoplePicker('people', {
                  label: 'PropertyFieldPeoplePicker',
                  initialData: this.properties.people,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'
                }),

                spfx.PropertyFieldColorPicker("color", {
                  label: "Font Color",
                  key: "color",
                  onPropertyChange: this.onPropertyChanged.bind(this),
                  properties: this.properties,
                  selectedColor: this.properties.color,
                }),
                spfx.PropertyFieldColorPicker("backgroundColor", {
                  label: "Background Color",
                  key: "backgroundColor",
                  onPropertyChange: this.onPropertyChanged.bind(this),
                  properties: this.properties,
                  selectedColor: this.properties.backgroundColor,
                }),
              ]
            },
            {
              groupName: "Code",
              isCollapsed: true,
              groupFields: [
                PropertyFieldCodeEditor("htmlCode", {
                  label: "Html Editor",
                  panelTitle:"Html Editor",
                  key: "htmlCode",
                  initialValue: this.properties.htmlCode,
                  language: PropertyFieldCodeEditorLanguages.HTML,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyChanged.bind(this)

                }),
                PropertyFieldCodeEditor("jsCode", {
                  label: "JS Editor",
                  panelTitle:"JS Editor",
                  key: "jsCode",
                  initialValue: this.properties.jsCode,
                  language: PropertyFieldCodeEditorLanguages.JavaScript,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyChanged.bind(this)

                }),

                PropertyFieldCodeEditor("cssCode", {
                  label: "Css Editor",
                  panelTitle:"Css Editor",
                  key: "cssCode",
                  initialValue: this.properties.cssCode,
                  language: PropertyFieldCodeEditorLanguages.css,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyChanged.bind(this)

                }),
              ]
            },
          ]
        }
      ]
    };
  }
}
