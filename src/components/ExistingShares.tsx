
import { IconButton } from "@microsoft/office-ui-fabric-react-bundle";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { ChatMessage, Drive, DriveItem, TeamsTab } from "@microsoft/microsoft-graph-types";
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
import { ExistingShare } from "./ExistingShare";

import "@pnp/sp/lists";

import "@pnp/sp/security";
import "@pnp/sp/security/web";
import "@pnp/sp/site-users/web";
import "@pnp/sp/views";
import "@pnp/sp/webs";
import { DetailsList, SelectionMode } from "office-ui-fabric-react/lib/DetailsList";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import * as React from "react";
import { ShareType } from "../model/model";
import { filter, map } from "lodash";
import { useEffect } from "react";
import { Spinner } from "office-ui-fabric-react";


// import "@pnp/graph/onedrive";
export interface IExistingSharesProps {
  setExistingShares: React.Dispatch<React.SetStateAction<any[]>>;
  onClose: () => void;
  existingShares: any[];
  title: string;
  context: BaseComponentContext;
  shareType: ShareType,
  sp: SPFI,
  graph: GraphFI,
  listId: string
  getTeamsTabConfig: Promise<[TeamsTab, string]>;
  itemId?:number;
}
export function ExistingShares(props: IExistingSharesProps) {

  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [selectedTeamShare, setSelectedTeamShare] = React.useState<any>(null);
  const [contentUrl, setContentUrl] = React.useState<any>(null);
  useEffect(() => {
    async function asyncStartup() {
      debugger;
      let [teamsTab, appUrl] = await props.getTeamsTabConfig
      setContentUrl(teamsTab.configuration.contentUrl);
      debugger;
    }
    // declare the data fetching function
    setIsLoading(true);

    asyncStartup().then(() => { 
      setIsLoading(false)
     });
  },[]);

  debugger;
  if (isLoading) {
    return (
      <Panel
        isOpen={true}
        onDismiss={props.onClose}
        headerText={props.title}

      ><Spinner label="Loading..."></Spinner></Panel>

    )
  }
  return (

    <Panel
      // type={PanelType.medium}
      isOpen={true}
      onDismiss={props.onClose}
      headerText={props.title}
    >
      <div>
        {contentUrl}
        {selectedTeamShare === null &&
          <DetailsList items={props.existingShares} selectionMode={SelectionMode.none}
            columns={[

              {
                key: "cmd",
                minWidth: 20, name: "", isResizable: true,
                onRender: (item?, index?, column?) => {
                  return <IconButton iconProps={{ iconName: "Edit" }} onClick={e => {
                    debugger;
                    setSelectedTeamShare(item)
                  }}></IconButton>
                }
              },
              {
                key: "Title",
                minWidth: 180, name: "Team", isResizable: true,
                onRender: (item?, index?, column?) => {
                  return item.Member.Title
                }
              }
            ]}
          />
        }
        {selectedTeamShare !== null &&
          <ExistingShare
          contentUrl={contentUrl}
            existingShare={selectedTeamShare}
            context={props.context}
            onClose={() => { debugger; setSelectedTeamShare(null) }}
            shareType={props.shareType}
            title={props.title}
            sp={props.sp} graph={props.graph}
            listId={props.listId}
            itemId={props.itemId}
            removeRoleAssignment={(roleDefId, principalId) => {
              debugger;
              //remove selected role
              var tempExistingShares = map(props.existingShares, ((es) => {
                if (es.PrincipalId === principalId) {
                  es.RoleDefinitionBindings = filter(es.RoleDefinitionBindings, (rdb) => { return rdb.Id !== roleDefId })
                }
                return es;
              }
              ));
              // if no roles left remove the item
              tempExistingShares = filter(tempExistingShares, (es) => { return es.RoleDefinitionBindings.length > 0 })
              props.setExistingShares(tempExistingShares);
              debugger;

            }}
       


          />
        }
      </div>




    </Panel>
  );
}
