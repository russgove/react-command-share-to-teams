import { IconButton } from "@microsoft/office-ui-fabric-react-bundle";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { graphfi, SPFx as SPFxGR } from "@pnp/graph";
import "@pnp/graph/";
import "@pnp/graph/groups";
import "@pnp/graph/onedrive";
import "@pnp/graph/sites";
import "@pnp/graph/sites/types";
import "@pnp/graph/teams";
import "@pnp/graph/users";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/folders";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/security";
import { IRoleDefinitionInfo } from "@pnp/sp/security";
import "@pnp/sp/security/web";
import { ISiteUserProps } from "@pnp/sp/site-users/types";
import "@pnp/sp/site-users/web";
import "@pnp/sp/views";
import { IViewInfo } from "@pnp/sp/views";
import "@pnp/sp/webs";
import { PrimaryButton } from "office-ui-fabric-react";
import { DetailsList, SelectionMode } from "office-ui-fabric-react/lib/DetailsList";
import { Label } from "office-ui-fabric-react/lib/Label";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import { ITag } from "office-ui-fabric-react/lib/Pickers";
import * as React from "react";
import { ShareType } from "../model/model";
import  * as Utilities from "../utilities";


// import "@pnp/graph/onedrive";
export interface IExistingShareProps {
  removeRoleAssignment:(roleDefId: number,PrincipalId:number)=>void;

  onClose: () => void;
  existingShare: any;
  title: string;
  context: BaseComponentContext;
  shareType: ShareType, sp: SPFI,
  listId:string
}
export function ExistingShare(props: IExistingShareProps) {

  debugger;
  return (

    <Panel
      type={PanelType.medium}
      isOpen={true}
      onDismiss={props.onClose}
      headerText={props.title}
    >
    <div>
      {props.existingShare["Member"]["Title"]}
      <Label >Permissions</Label>
      <DetailsList items={props.existingShare.RoleDefinitionBindings} selectionMode={SelectionMode.none}
        columns={[

          {
            key: "cmd",
            minWidth: 20, name: "",
            onRender: (item?, index?, column?) => {
              return <IconButton iconProps={{ iconName: "Delete" }} onClick={e => {
                debugger;
             Utilities.removeRoleAssignmentFromList({listId:props.listId,ra:props.existingShare,roleDefId:item.Id,sp:props.sp})
             props.removeRoleAssignment(item.Id,props.existingShare.PrincipalId)
              }}></IconButton>
            }
          },
          {
            key: "Name",
            minWidth: 90, name: "Name", isResizable: true,
            onRender: (item?, index?, column?) => {
              return item.Name
            }
          },
          {
            key: "Description", isResizable: true, isMultiline:true,
            minWidth: 400, name: "Description",
            onRender: (item?, index?, column?) => {
              return item.Description
            }
          }
        ]}
      />
    <PrimaryButton onClick={props.onClose}>Done</PrimaryButton>
    </div>
    </Panel>
  );
}
