import * as React from 'react';
import styles from './SppPeopleFinder.module.scss';
import * as strings from 'SppPeopleFinderWebPartStrings';
import { escape } from '@microsoft/sp-lodash-subset';
import { Stack, IStackProps, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react';
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Log } from "@microsoft/sp-core-library";

/*Componentes*/
import { ByFirstName } from "./Search/ByFirstName";

/* State & Props */
import { ISppPeopleFinderProps } from './ISppPeopleFinderProps';
import { ISppPeopleFinderState } from "./ISppPeopleFinderState";

/*Stack properties*/
const stackTokens = { childrenGap: 20 };
const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 700 } },
};
const LOG_SOURCE = "SPPPeopleFinder";

export default class SppPeopleFinder extends React.Component<ISppPeopleFinderProps, ISppPeopleFinderState> {
  private headers = [
    { label:'Name', key: 'displayName' },
    { label:'Email', key: 'email' },
    { label:'Mobile Phone',key:'mobilePhone'},
    { label:'JobTitle',key:'JobTitle'},
    { label:'OfficeLocation',key:'OfficeLocation'},
    { label:'Business Phone',key:'businessPhone'
  }];
  
  constructor(props:ISppPeopleFinderProps){
    super(props);

    /*definir columnas de grilla*/
    const  columns: IColumn[] = [
      {
        key: 'column1',
        name: strings.DisplayName,
        isRowHeader: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        fieldName: 'displayName',
        minWidth: 100,
        maxWidth: 300,
        isResizable: false
      },
      {
        key: 'column2',
        name: strings.Email,
        fieldName: 'email',
        isSorted: true,
        isSortedDescending: false,
        minWidth: 200,
        maxWidth: 300,
        isResizable: false
      },
      {
        key: 'column3',
        name: strings.MobilePhone,
        fieldName: 'mobilePhone',
        isSorted: true,
        isSortedDescending: false,
        minWidth: 200,
        maxWidth: 300,
        isResizable: false
      },
      {
        key: 'column5',
        name: strings.JobTitle,
        fieldName: 'JobTitle',
        isSorted: true,
        isSortedDescending: false,
        minWidth: 200,
        maxWidth: 300,
        isResizable: false
      },
      {
        key: 'column6',
        name: strings.OfficeLocation,
        fieldName: 'OfficeLocation',
        isSorted: true,
        isSortedDescending: false,
        minWidth: 100,
        maxWidth: 300,
        isResizable: true
      },
      {
        key: 'column7',
        name: strings.businessPhone,
        fieldName: 'businessPhone',
        isSorted: true,
        isSortedDescending: false,
        minWidth: 100,
        maxWidth: 300,
        isResizable: true
      }
    ];
    
    this.state={
      loading:false,
      selectedKey:"byFirstName",
      columns:columns,
      searchFor:'',
      userProperties:[]
    };

    this.searchUsersButton = this.searchUsersButton.bind(this);
    this._handleChange = this._handleChange.bind(this);
  }

  componentDidMount() {
    this.getUsers(this.state.searchFor);
  }

  private _handleChange(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void
  {
    try {
      this.setState({
       searchFor: newValue,
      });
    } catch (error) 
    {
      Log.error(LOG_SOURCE,error);
    }
  }

  //button
  private searchUsersButton(): void {
    try {
     
      this.getUsers(this.state.searchFor);
    } 
    catch (error) 
    {
      Log.error(LOG_SOURCE,error);
    }
  }

  private async getUsers(searchby:string) : Promise<any>{
    this.setState({loading:true},async()=>{
       await this.props.MSGraphServiceInstance
       .getUsersByAllProperties(searchby,this.props.MsGraphClient)
       // tslint:disable-next-line: no-shadowed-variable
       .then((users)=>{
         if(users.length !== 0){
           this.setState({
             userProperties:users,
             loading:false
           });
         }
         else
         {
           this.setState({
             userProperties:[],
             loading:false
           });
         }
       });
     });
   }

   private searchUsersError(value: string): string {
    // The search for text cannot contain spaces
      return (value == null || value.length == 0 || value.indexOf(" ") < 0)
      ? ''
      : 'Nothing matched';
  }

  public render(): React.ReactElement<ISppPeopleFinderProps> {
    return (
      <div className={ styles.sppPeopleFinder }>
            <div>
              <div>
                <div>
                <WebPartTitle displayMode={this.props.DisplayMode}
                  title={this.props.WebpartTitle}
                  updateProperty={this.props.updateProperty} />
                  <Stack horizontal tokens={stackTokens} styles={stackStyles}>
                    <Stack {...columnProps}>
                      <TextField
                        label={strings.SearchUserByFirstName}
                        required={false}
                        value={this.state.searchFor}
                        onChange={this._handleChange}
                        onGetErrorMessage={this.searchUsersError}
                      />
                      <DefaultButton text="Search" onClick={this.searchUsersButton} />
                    </Stack>
                  </Stack>
                 <br/>
                  {this.state.selectedKey === "byFirstName" &&
                    <ByFirstName 
                      userProperties={this.state.userProperties}
                      Columns={this.state.columns}
                      loading={this.state.loading}>                      
                    </ByFirstName>
                  }
                  
              </div>
             </div>
           </div>
       </div>
    );
  }
}
