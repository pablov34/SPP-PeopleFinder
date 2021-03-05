import { IUserProperties } from "../../../Services/IUserProperties";
import { IColumn } from "office-ui-fabric-react/lib/DetailsList";

export interface ISppPeopleFinderState{
    loading:boolean;
    columns:IColumn[];
    selectedKey:string;
    searchFor: string;
    userProperties:IUserProperties[];
}