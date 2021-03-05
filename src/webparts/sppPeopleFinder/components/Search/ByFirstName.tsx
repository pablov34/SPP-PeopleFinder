import * as React from "react";
import { ByFirstNameProps } from "./ByFirstNameProps";
import styles from '../SppPeopleFinder.module.scss';
import { ByFirstNameState } from "./ByFirstNameState";
import { autobind } from "office-ui-fabric-react/lib/Utilities";
import * as strings from 'SppPeopleFinderWebPartStrings';

import { Log } from "@microsoft/sp-core-library";
import { Stack, IStackProps, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react';
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';

const LOG_SOURCE = "ByFirstName";
const stackTokens = { childrenGap: 20 };
const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 700 } },
};
export class ByFirstName extends React.Component<ByFirstNameProps,ByFirstNameState>{
    constructor(props:ByFirstNameProps){
        super(props);
        this.state={
            loading:false,
            searchFor: '',
            userProperties:[],
            isDataFound:true,
          };
          this.searchUsersButton = this.searchUsersButton.bind(this);
          this._handleChange = this._handleChange.bind(this);
    }

  componentDidMount() {
    this.getUsers(this.state.searchFor);
  }

  @autobind
  private async getUsers(searchby:string) : Promise<any>{
   this.setState({loading:true},async()=>{
      await this.props.MSGraphServiceInstance
      .getUsersByAllProperties(searchby,this.props.MSGraphClient)
      // tslint:disable-next-line: no-shadowed-variable
      .then((users)=>{
        if(users.length !== 0){
          this.setState({
            userProperties:users,
            isDataFound:true,
            loading:false
          });
        }
        else
        {
          this.setState({
            userProperties:[],
            isDataFound:false,
            loading:false
          });
        }
      });
    });
  }

  @autobind
  private searchUsersError(value: string): string {
    // The search for text cannot contain spaces
      return (value == null || value.length == 0 || value.indexOf(" ") < 0)
      ? ''
      : 'Nothing matched';
  }

  //handle change, set state
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


  public render(): React.ReactElement<ByFirstNameProps> 
  {    
        if(!this.state.loading){
            return (<div className={styles.sppPeopleFinder}> 
                <div>
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
                  <div>
                  </div>
                    <div id='detailedList'>
                      {this.state.userProperties.length !== 0 && 
                        <DetailsList
                          items={this.state.userProperties}
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

