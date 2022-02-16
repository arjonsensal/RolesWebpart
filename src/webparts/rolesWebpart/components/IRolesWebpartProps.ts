
import { WebPartContext } from '@microsoft/sp-webpart-base';  
export interface IRolesWebpartProps {
  description: string;
  context: WebPartContext;  
  listName: string;
  unique: string;
  filterList: string;
  uniqueFilter: string;
}
  