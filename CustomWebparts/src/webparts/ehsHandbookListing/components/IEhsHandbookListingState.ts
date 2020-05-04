import { IPickerTerms } from '@pnp/spfx-controls-react/lib/TaxonomyPicker';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { HandbookComposite } from '../../../BAL/HandbookComposite';

export interface IEhsHandbookListingState {
  siblingMetadata: HandbookComposite[];
  showPanel: boolean;
  hideDialog: boolean;
  pageName: string;
  handBookCompositeMetadata: HandbookComposite[];
  selectedKey: string;
  sortText: string;
  sortIcon: string;
  sortAsc: boolean;
  filterDisplay: string;
  statusText: string;
  filterState: { name: string, internalName: string, termsetId: string, id: string, allowMultipleValues: boolean, value: IPickerTerms }[];
  options: IDropdownOption[];
  sortByField: string;
  sortByType: string;
  pageContentType: string;
}