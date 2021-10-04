import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneButton,
  PropertyPaneLabel,
  PropertyPaneHorizontalRule
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'GlossaryWpWebPartStrings';
import GlossaryWp from './components/GlossaryWp';
import { IGlossaryWpProps } from './components/IGlossaryWpProps';
import { PropertyPaneAsyncDropdown } from './controls/PropertyPaneAsyncDropdown';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { update, get } from '@microsoft/sp-lodash-subset';
import { SPHttpClient , SPHttpClientResponse, ISPHttpClientOptions} from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';  

export interface IGlossaryWpWebPartProps {
  listName: string;
  definitions: string;
  createListTextField: string;
  context: WebPartContext;
  addNewTerm: string;
  addNewDesc: string;
  addNewDef: string;
  addNewDefDesc: string;
}

export default class GlossaryWpWebPart extends BaseClientSideWebPart<IGlossaryWpWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGlossaryWpProps> = React.createElement(
      GlossaryWp,
      {
        listName: this.properties.listName,
        context: this.context,
        createListTextField: this.properties.createListTextField,
        definitions: this.properties.definitions,
        addNewTerm: this.properties.addNewTerm,
        addNewDesc: this.properties.addNewDesc,
        addNewDef: this.properties.addNewDef,
        addNewDefDesc: this.properties.addNewDefDesc
      }
    );

    ReactDom.render(element, this.domElement);
  }
  private newTermDisabled: boolean = true;
  private newDefDisabled: boolean = true;

  private async loadLists(): Promise<IDropdownOption[]> {
    return new Promise<IDropdownOption[]>(async(resolve: (options: IDropdownOption[]) => void, reject: (error: any) => void) => {
      
      setTimeout(async ()=> {
        var a = new Array()
        const restApi = `${this.context.pageContext.web.absoluteUrl}/_api/lists`;
        await this.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
          .then(resp => { return resp.json(); })
          .then(items => {
            items.value.forEach(element => {
              if(element.Description.includes("Glossary List Items")) {
                a.push({key: element.Title, text: element.Title})
              }
            });
            resolve(a);
          });
      }, 1500)
    });
  }

  private async loadListsDef(): Promise<IDropdownOption[]> {
    return new Promise<IDropdownOption[]>(async(resolve: (options: IDropdownOption[]) => void, reject: (error: any) => void) => {
      
      setTimeout(async ()=> {
        var a = new Array()
        const restApi = `${this.context.pageContext.web.absoluteUrl}/_api/lists`;
        await this.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
          .then(resp => { return resp.json(); })
          .then(items => {
            items.value.forEach(element => {
              if(element.Description.includes("Definitions List")) {
                a.push({key: element.Title, text: element.Title})
              }
            });
            resolve(a);
          });
      }, 1500)
    });
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.newTermDisabled = !this.properties.listName;
    this.newDefDisabled = !this.properties.definitions;
    this.properties.createListTextField = "";
    this.context.propertyPane.refresh();
  };

  
  public componentDidUpdate(prevProps, prevState): void {
    console.log(prevProps, this)
  }

  private onListChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });
    this.newTermDisabled = false;
    this.context.propertyPane.refresh();
    // refresh web part
    this.render();
  }
  
  private onListDefChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });
    this.context.propertyPane.refresh();
    this.newDefDisabled = false;
    // refresh web part
    this.render();
  }
  
  private handleCreateList() {
    console.log(this.properties.createListTextField)
    const getListUrl = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${this.properties.createListTextField}')`;
    const getListUrlFields = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${this.properties.createListTextField}')/fields`;
    this.context.spHttpClient.get(getListUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      if (response.status === 200) {
      alert("List already exists.");
      return; // list already exists
      }
      if (response.status === 404) {
        const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists";
        const listDefinition : any = {
          "Title": this.properties.createListTextField,
          "Description": this.properties.createListTextField + "Glossary List Items",
          "AllowContentTypes": true,
          "BaseTemplate": 100,
          "ContentTypesEnabled": true,
        };
        const spHttpClientOptions: ISPHttpClientOptions = {
          "body": JSON.stringify(listDefinition)
        };
        this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
        .then((response: SPHttpClientResponse) => {
          if (response.status === 201) {
            alert("List created successfully");
            // Create columns
            const column1: ISPHttpClientOptions = {
              "body": JSON.stringify({
                 '@odata.type': '#SP.Field', 
                 'FieldTypeKind': 2, 
                 'Title':'GlossaryDesc'
              })
            };
            this.context.spHttpClient.post(getListUrlFields, SPHttpClient.configurations.v1, column1)
              .then((response: SPHttpClientResponse) => {
                if (response.status === 201) {
                  alert("Columns created successfully");
                } else {
                  alert("Response status "+response.status+" - "+response.statusText);
                }
              });

          } else {
            //alert("Response status "+response.status+" - "+response.statusText);
          }
        });
      } else {
        //alert("Something went wrong. "+response.status+" "+response.statusText);
      }
    });
    
    this.loadLists();
  }
  
  private handleCreateDefList() {
    const getListUrl = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${this.properties.definitions}')`;
    const getListUrlFields = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${this.properties.definitions}')/fields`;
    this.context.spHttpClient.get(getListUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      if (response.status === 200) {
      alert("List already exists.");
      return; // list already exists
      }
      if (response.status === 404) {
        const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists";
        const listDefinition : any = {
          "Title": this.properties.definitions,
          "Description": this.properties.definitions + "Definitions List",
          "AllowContentTypes": true,
          "BaseTemplate": 100,
          "ContentTypesEnabled": true,
        };
        const spHttpClientOptions: ISPHttpClientOptions = {
          "body": JSON.stringify(listDefinition)
        };
        this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
        .then((response: SPHttpClientResponse) => {
          if (response.status === 201) {
            alert("List created successfully");
            // Create columns
            const column1: ISPHttpClientOptions = {
              "body": JSON.stringify({
                 '@odata.type': '#SP.Field', 
                 'FieldTypeKind': 2, 
                 'Title':'DefinitionDescription'
              })
            };
            this.context.spHttpClient.post(getListUrlFields, SPHttpClient.configurations.v1, column1)
              .then((response: SPHttpClientResponse) => {
                if (response.status === 201) {
                  alert("Columns created successfully");
                } else {
                  alert("Response status "+response.status+" - "+response.statusText);
                }
              });

          } else {
            //alert("Response status "+response.status+" - "+response.statusText);
          }
        });
      } else {
        //alert("Something went wrong. "+response.status+" "+response.statusText);
      }
    });
  }
  
  private handleCreateTerm() {
    const getListUrl = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${this.properties.listName}')/items?$filter=Title eq '${this.properties.addNewTerm}'`;
    this.context.spHttpClient.get(getListUrl, SPHttpClient.configurations.v1).then(resp => { return resp.json();
    }).then(item => {
      if (item.value.length > 0) {
      alert("Term already exists.");
      return; // term already exists
      } else {
        const url: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${this.properties.listName}')/items`;
        const listDefinition : any = {
          '@odata.type': `#SP.Data.${(this.properties.listName).replace(" ", "_x0020_")}ListItem`, 
          'Title': this.properties.addNewTerm,
          'GlossaryDesc': this.properties.addNewDesc
        };
        const spHttpClientOptions: ISPHttpClientOptions = {
          "body": JSON.stringify(listDefinition)
        };
        this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
        .then((response: SPHttpClientResponse) => {
          if (response.status === 201) {
            alert("Term created successfully");
          } else {
            alert("Response status "+response.status+" - "+response.statusText);
          }
        });
      }
    });
  }
  
  private handleCreateDef() {
    const getListUrl = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${this.properties.definitions}')/items?$filter=Title eq '${this.properties.addNewDef}'`;
    this.context.spHttpClient.get(getListUrl, SPHttpClient.configurations.v1).then(resp => { return resp.json();
    }).then(item => {
      if (item.value.length > 0) {
      alert("Definition already exists.");
      return; // term already exists
      } else {
        const url: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${this.properties.definitions}')/items`;
        const listDefinition : any = {
          '@odata.type': `#SP.Data.${(this.properties.definitions).replace(" ", "_x0020_")}ListItem`, 
          'Title': this.properties.addNewDef,
          'DefinitionDescription': this.properties.addNewDefDesc
        };
        const spHttpClientOptions: ISPHttpClientOptions = {
          "body": JSON.stringify(listDefinition)
        };
        this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
        .then((response: SPHttpClientResponse) => {
          if (response.status === 201) {
            alert("Definition created successfully");
          } else {
            alert("Response status "+response.status+" - "+response.statusText);
          }
        });
      }
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('createListTextField', {
                  label: strings.CreateListLabel
                }),
                PropertyPaneButton('createButton', {
                  text: strings.ButtonLabel,
                  onClick: this.handleCreateList.bind(this)
                }),
                PropertyPaneLabel('Orlabel', {
                  text: "OR"
                }),
                new PropertyPaneAsyncDropdown('listName', {
                  label: strings.ListFieldLabel,
                  loadOptions: this.loadLists.bind(this),
                  onPropertyChange: this.onListChange.bind(this),
                  selectedKey: this.properties.listName
                })
              ]
            },
            {
              groupFields: [
                PropertyPaneLabel('newTermLabel', {
                  text: "Add New Term"
                }),
                PropertyPaneTextField('addNewTerm', {
                  label: "Term Title",
                  disabled: this.newTermDisabled
                }),
                PropertyPaneTextField('addNewDesc', {
                  label: "Term Description",
                  disabled: this.newTermDisabled
                }),
                PropertyPaneButton('addTermButton', {
                  text: "Add Term",
                  onClick: this.handleCreateTerm.bind(this),
                  disabled: this.newTermDisabled
                })
              ]
            },
            {
              groupFields: [
                PropertyPaneHorizontalRule()
              ]
            },
            {
              groupFields: [
                PropertyPaneTextField('definitions', {
                  label: "Definitions"
                }),
                PropertyPaneButton('createDefinitions', {
                  text: "Create new Definitions",
                  onClick: this.handleCreateDefList.bind(this)
                }),
                PropertyPaneLabel('OrlabelDef', {
                  text: "OR"
                }),
                new PropertyPaneAsyncDropdown('definitions', {
                  label: "Select Definitions List",
                  loadOptions: this.loadListsDef.bind(this),
                  onPropertyChange: this.onListDefChange.bind(this),
                  selectedKey: this.properties.definitions
                })
              ]
            },
            {
              groupFields: [
                PropertyPaneLabel('newDefLabel', {
                  text: "Add New Definition"
                }),
                PropertyPaneTextField('addNewDef', {
                  label: "Definition",
                  disabled: this.newDefDisabled
                }),
                PropertyPaneTextField('addNewDefDesc', {
                  label: "Definition Description",
                  disabled: this.newDefDisabled
                }),
                PropertyPaneButton('addDefinitionButton', {
                  text: "Add Definition",
                  disabled: this.newDefDisabled,
                  onClick: this.handleCreateDef.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
