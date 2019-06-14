import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CustomerInformationWebPart.module.scss';
import * as strings from 'CustomerInformationWebPartStrings';

export interface ICustomerInformationWebPartProps {
  description: string;
}


  export interface ISPListCustomers{
  value:ISPListCustomerItem[];
  }

  export interface ISPListCustomerItem{
    Title:string;
    CustomerID:string;
    CustomerName:string;
    CustomerAddress:string;
    CustomerType:string;
    
    }

export default class CustomerInformationWebPart extends BaseClientSideWebPart<ICustomerInformationWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.customerInformation }">
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
