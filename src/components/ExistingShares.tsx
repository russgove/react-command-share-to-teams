
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


// import "@pnp/graph/onedrive";
export interface IExistingSharesProps {

  onClose: () => void;
  existingShares: any[];
  title: string;
  context: BaseComponentContext;
  shareType: ShareType, sp: SPFI,listId:string
}
export function ExistingShares(props: IExistingSharesProps) {
  const sp = spfi().using(SPFx(props.context));
  const graph = graphfi().using(SPFxGR(props.context));

  //const [shareMethod, setShareMethod] = React.useState<ShareMethod>(0);
  const [selectedTeamShare, setSelectedTeamShare] = React.useState<any>(null);


  debugger;
  return (

    <Panel
      type={PanelType.medium}
      isOpen={true}
      onDismiss={props.onClose}
      headerText={props.title}
    >
      <div>
        {selectedTeamShare === null &&
          <DetailsList items={props.existingShares} selectionMode={SelectionMode.none}
            columns={[

              {
                key: "cmd",
                minWidth: 50, name: "",
                onRender: (item?, index?, column?) => {
                  return <IconButton iconProps={{ iconName: "Edit" }} onClick={e => {
                    debugger;
                    setSelectedTeamShare(item)
                  }}></IconButton>
                }
              },
              {
                key: "Title",
                minWidth: 400, name: "Team",
                onRender: (item?, index?, column?) => {
                  return item.Member.Title
                }
              }
            ]}
          />
        }
        {selectedTeamShare !== null &&
          <ExistingShare
            existingShare={selectedTeamShare}
            context={props.context}
            onClose={() => setSelectedTeamShare(null)}
            shareType={props.shareType}
            title={props.title}
            sp={props.sp}
            listId={props.listId}


          />
        }
      </div>




    </Panel>
  );
}
