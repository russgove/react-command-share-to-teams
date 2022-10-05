import { IconButton } from "@microsoft/office-ui-fabric-react-bundle";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { GraphFI, graphfi, SPFx as SPFxGR } from "@pnp/graph";
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
import { filter } from "lodash";
import { PrimaryButton } from "office-ui-fabric-react";
import { DetailsList, SelectionMode } from "office-ui-fabric-react/lib/DetailsList";
import { Label } from "office-ui-fabric-react/lib/Label";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import { ITag } from "office-ui-fabric-react/lib/Pickers";
import * as react from "react";
import * as React from "react";
import { useEffect } from "react";
import { ShareType } from "../model/model";
import * as Utilities from "../utilities";


// import "@pnp/graph/onedrive";
export interface IExistingShareProps {
  removeRoleAssignment: (roleDefId: number, PrincipalId: number) => void;
  onClose: () => void;
  existingShare: any;
  title: string;
  context: BaseComponentContext;
  shareType: ShareType,
  sp: SPFI,
  graph: GraphFI,
  listId: string
  contentUrl: string;

}
export function ExistingShare(props: IExistingShareProps) {
  const [teamsTabs, setTeamsTabs] = react.useState<any[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  debugger;
  useEffect(() => {
    async function asyncStartup() {
      debugger;
      setTeamsTabs(await Utilities.getTeamTabs({ graph: props.graph, teamId: Utilities.getTeamIdFromLoginName(props.existingShare.Member.LoginName), contentUrl: props.contentUrl }));
    }
    // declare the data fetching function
    setIsLoading(true);

    asyncStartup().then(() => {
      setIsLoading(false)
    });
  }, []);

  debugger;
  return (

    <Panel
      type={PanelType.medium}
      isOpen={true}
      onDismiss={props.onClose}
      headerText={props.title}
    >
      <div>
        {props.contentUrl}<br />
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
                  Utilities.removeRoleAssignmentFromList({ listId: props.listId, ra: props.existingShare, roleDefId: item.Id, sp: props.sp })
                  props.removeRoleAssignment(item.Id, props.existingShare.PrincipalId)
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
              key: "Description", isResizable: true, isMultiline: true,
              minWidth: 400, name: "Description",
              onRender: (item?, index?, column?) => {
                return item.Description
              }
            }
          ]}
        />
        <Label >TeamsTabs</Label>
        <DetailsList items={teamsTabs} selectionMode={SelectionMode.none}
          columns={[

            {
              key: "cmd",
              minWidth: 20, name: "",
              onRender: (item?, index?, column?) => {
                return <IconButton iconProps={{ iconName: "Delete" }} onClick={e => {
                  debugger;
                  debugger;
                  //remove selected tab
                  Utilities.removeTeamsTab(props.graph, Utilities.getTeamIdFromLoginName(props.existingShare.Member.LoginName), item.channelId, item.id)
                  const temptabs = filter(teamsTabs, (tt) => { return tt.id !== item.id })
                  setTeamsTabs(temptabs);
                  debugger;
                }}></IconButton>
              }
            },
            {
              key: "Channel",
              minWidth: 90, name: "Channel Name", fieldName: "displayName", isResizable: true,
              onRender: (item?, index?, column?) => {
                return item.channelName
              }
            },
            {
              key: "Name",
              minWidth: 90, name: "Tab Name", fieldName: "displayName", isResizable: true,
              onRender: (item?, index?, column?) => {
                return item.displayName
              }
            },
            {
              key: "contentUrl",
              minWidth: 90, name: "Url", fieldName: "displayName", isResizable: true,
              onRender: (item?, index?, column?) => {
                return item.configuration.contentUrl
              }
            },

          ]}
        />
        <PrimaryButton onClick={props.onClose}>Done</PrimaryButton>
      </div>
    </Panel>
  );
}
