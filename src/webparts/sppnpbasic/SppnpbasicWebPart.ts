import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SppnpbasicWebPart.module.scss';
import * as strings from 'SppnpbasicWebPartStrings';

import { sp } from '@pnp/sp'

export interface ISppnpbasicWebPartProps {
  description: string;
}

export default class SppnpbasicWebPart extends BaseClientSideWebPart<ISppnpbasicWebPartProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      // other init code may be present
  
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {

    this.domElement.innerHTML = "Loading ...";
    setTimeout(async () => {

      const data = await sp.web.select("Title", "Description").get();
      this.domElement.innerHTML = `<pre>${JSON.stringify(data, null, 2)}</pre>`;}, 1000);
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
