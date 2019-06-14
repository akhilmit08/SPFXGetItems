import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CustomerInformationWebPart.module.scss';
import * as strings from 'CustomerInformationWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

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
                <div class="${ styles.container }">
                  <div id="spListContainer"/>
                  </div>
                `;
this._renderListDataAsync();
  }

  // Get List Items Method
  private _getListCustomerData():Promise<ISPListCustomers>
{
return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl+
`/_api/web/lists/GetByTitle('Customers')/Items`,SPHttpClient.configurations.v1).
then((responseListCustomer:SPHttpClientResponse)=>{
       return responseListCustomer.json();
});
}

private _renderListCustomer(items:ISPListCustomerItem[]):void
{
let html:string=`<table width='100%' border=2>`;
html+=`<thead><th>ID</th><th>Name</th><th>Address</th><th>Type</th><th>Author</th>
<th>Lookup</th>`+
`</thead><tbody>`;
items.forEach((item:ISPListCustomerItem)=>
{
html+= `<tr><td>${item.CustomerID}</td>
<td>${item.CustomerName}</td>
<td>${item.CustomerAddress}</td>
<td>${item.CustomerType}</td>

</tr>`;
});
html+=`</tbody></table>`;
const listContainer:Element=this.domElement.querySelector("#spListContainer");
listContainer.innerHTML=html;
}

private _renderListDataAsync():void
{
this._getListCustomerData().then((response)=>
{
this._renderListCustomer(response.value);
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
