import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphService } from "../../../../Services/MSGraphService";
import { MSGraphClient } from "@microsoft/sp-http";
import { IColumn } from "office-ui-fabric-react/lib/DetailsList";
import { IUserProperties } from "../../../../Services/IUserProperties";

export interface ByFirstNameProps{
    Columns:IColumn[];
    userProperties:IUserProperties[];
    loading:boolean;
}