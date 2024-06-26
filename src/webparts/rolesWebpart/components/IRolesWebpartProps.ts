
import { WebPartContext } from '@microsoft/sp-webpart-base';  
export interface IRolesWebpartProps {
  description: string;
  context: WebPartContext;  
  listName: string;
  unique: string;
  columns: string;
  filterList: string;
  uniqueFilter: string;
  optionalColumnFilter: string;
  optionalColumnFilterValue: string;
  removeColumns: string;
}
  