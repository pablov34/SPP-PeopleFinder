import * as React from "react";
import { ByFirstNameProps } from "./ByFirstNameProps";
import styles from '../SppPeopleFinder.module.scss';
import { ByFirstNameState } from "./ByFirstNameState";
import { autobind } from "office-ui-fabric-react/lib/Utilities";
import * as strings from 'SppPeopleFinderWebPartStrings';

import { Log } from "@microsoft/sp-core-library";
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';

const LOG_SOURCE = "ByFirstName";

export class ByFirstName extends React.Component<ByFirstNameProps,ByFirstNameState>{
    constructor(props:ByFirstNameProps){
        super(props);
        this.state={
            loading:false,
            isDataFound:true,
          };
    }

  public render(): React.ReactElement<ByFirstNameProps> 
  {    
        if(!this.props.loading){
            return (<div className={styles.sppPeopleFinder}> 
                    <div>
                      <div id='detailedList'>
                        {this.props.userProperties.length !== 0 && 
                          <DetailsList
                            items={this.props.userProperties}
                            columns={this.props.Columns}
                            isHeaderVisible={true}
                            layoutMode={DetailsListLayoutMode.justified}
                          />
                        }
                      </div>
                     </div>
                   </div>
            )
          }
          else
          {
              return(
                <div>
                <Spinner label="Loading..." />
              </div>
              )
          } 
    
  }
}

