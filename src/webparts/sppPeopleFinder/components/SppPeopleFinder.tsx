import * as React from 'react';
import styles from './SppPeopleFinder.module.scss';
import * as strings from 'SppPeopleFinderWebPartStrings';
import { escape } from '@microsoft/sp-lodash-subset';
import { Pivot, PivotItem,PivotLinkFormat, PivotLinkSize } from 'office-ui-fabric-react/lib/Pivot';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";

/*Componentes*/
import { ByFirstName } from "./Search/ByFirstName";

/* State & Props */
import { ISppPeopleFinderProps } from './ISppPeopleFinderProps';
import { ISppPeopleFinderState } from "./ISppPeopleFinderState";

export default class SppPeopleFinder extends React.Component<ISppPeopleFinderProps, ISppPeopleFinderState> {
  private headers = [
    { label: 'Name', key: 'displayName' },
    { label: 'Email', key: 'email' },
    { label:'Mobile Phone',key:'mobilePhone'},
    { label:'JobTitle',key:'JobTitle'},
    { label:'OfficeLocation',key:'OfficeLocation'},
    { label:'Business Phone',key:'businessPhone'
  }];
  
  constructor(props:ISppPeopleFinderProps){
    super(props);

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
      columns:columns
    };
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
                
                 <br/>
                  {this.state.selectedKey === "byFirstName" &&
                    <ByFirstName 
                      MSGraphClient={this.props.MsGraphClient} 
                      MSGraphServiceInstance={this.props.MSGraphServiceInstance}
                      context={this.props.context}
                      Columns={this.state.columns}></ByFirstName>
                  }
                  
              </div>
             </div>
           </div>
       </div>
    );
  }
}
