import { WebPartContext } from '@microsoft/sp-webpart-base';  
export interface IGlossaryWpProps {
  listName: string;
  definitions: string;
  createListTextField: string;
  context: WebPartContext;  
  addNewTerm: string;
  addNewDesc: string;
  addNewDef: string;
  addNewDefDesc: string;
}

