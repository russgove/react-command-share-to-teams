import { TeamsTab } from "@microsoft/microsoft-graph-types";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { BaseDialog, IDialogConfiguration } from "@microsoft/sp-dialog";
import { AadHttpClient } from "@microsoft/sp-http";
import { IListViewCommandSetExecuteEventParameters } from "@microsoft/sp-listview-extensibility";
import { graphfi, SPFx as SPFxGR } from "@pnp/graph";
import "@pnp/graph/teams";
import "@pnp/graph/users";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/folders";
import "@pnp/sp/items";
import { IItem, Item } from "@pnp/sp/items";
import "@pnp/sp/lists";
import { IListInfo } from "@pnp/sp/lists";
import "@pnp/sp/security";
import { IBasePermissions, IRoleDefinition, IRoleDefinitionInfo, PermissionKind, RoleDefinitions } from "@pnp/sp/security";
import "@pnp/sp/security/web";
import "@pnp/sp/views";
import { IViewInfo } from "@pnp/sp/views";
import "@pnp/sp/webs";
import { TeamChannelPicker } from "@pnp/spfx-controls-react/lib/TeamChannelPicker";
import { TeamPicker } from "@pnp/spfx-controls-react/lib/TeamPicker";
import { find } from "lodash";
import { ChoiceGroup, PrimaryButton } from "office-ui-fabric-react";
import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';
import { DetailsList, SelectionMode } from "office-ui-fabric-react/lib/DetailsList";
import { DialogContent } from "office-ui-fabric-react/lib/Dialog";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import { ITag } from "office-ui-fabric-react/lib/Pickers";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import * as React from "react";
import { useEffect } from "react";
import * as ReactDOM from "react-dom";
import { ShareType } from "../model/model";
interface IShareToTeamsProps {
  title: string;
  close: () => void;
  aadHttpClient: AadHttpClient;
  context: BaseComponentContext;
  event: IListViewCommandSetExecuteEventParameters;
}
function ShareToTeamsContent(props: IShareToTeamsProps) {
  function _onSelectedTeams(tagList: ITag[]) {
    setSelectedTeams(tagList);
  };
  function _onSelectedTeamChannels(tagList: ITag[]) {
    setSelectedTeamChannels(tagList);
  }
  async function addTab() {
    const graph = graphfi().using(SPFxGR(props.context));
    debugger;

    const teamId: string = selectedTeams[0].key as string;
    const channelId: string = selectedTeamChannels[0].key as string;
    console.log(`TEAM ID is ${teamId}. CHANNEL ID is ${channelId}`);
    const team = await graph.teams.getById(teamId)();
    console.log(team);
    const channel = await graph.teams.getById(teamId).channels.getById(channelId);
    console.log(channel);

    const tabs = await graph.teams.getById(teamId).channels.getById(channelId).tabs;
    console.log(tabs);






    const teamsTab: TeamsTab = {} as TeamsTab;
    teamsTab.displayName = tabName;

    teamsTab.configuration = {
      contentUrl: "https://russellwgove.sharepoint.com/sites/CR-EU-Manufacturing/Shared%20Documents/Forms/AllItems.aspx",
    }

    const newTab = tabs.add('Tab', 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/2a527703-1f6f-4559-a332-d8a7d288cd88', teamsTab)
      .then((t) => {
        debugger;
        console.log(t);
      })
      .catch(err => {
        debugger;
        console.log(err);
        alert(err.message);
      });

    // const newTab = await graph.teams.getById(teamId)
    //   .channels
    //   .getById(channelId)
    //   .tabs
    //   .add('Tab', 'https://www.google.com', teamsTab);
    console.log(newTab);
    debugger;
  }
  const [shareType, setShareType] = React.useState<ShareType>(null);
  const [list, setList] = React.useState<IListInfo>(null);
  const [item, setItem] = React.useState<any>(null);
  const [selectedTeams, setSelectedTeams] = React.useState<ITag[]>([]);
  const [selectedTeamChannels, setSelectedTeamChannels] = React.useState<ITag[]>([]);
  const [roleDefinitionInfos, setRoleDefinitionInfos] = React.useState<IRoleDefinitionInfo[]>([]);
  const [folderServerRelativePath, setFolderServerRelativePath] = React.useState<string>(null);
  const [userCanManagePermissions, setUserCanManagePermissions] = React.useState<boolean>(false);
  const [allViews, setAllViews] = React.useState<IViewInfo[]>([]);
  const [viewId, setViewId] = React.useState<string>(null);
  const [tabName, setTabName] = React.useState<string>("");
  const [libraryName, setLibraryName] = React.useState<string>("");
  const [permissionsOnSP, setPermissionsOnSP] = React.useState<IBasePermissions>(null);
  async function getRoleDefs(sp) {
    // get the role definitions for the current web -- now full condtrol or designer
    await sp.web.roleDefinitions
      .filter("BasePermissions ne null and Hidden eq false and RoleTypeKind ne 4 and RoleTypeKind ne 5 and RoleTypeKind ne 6")  // 4 is designer, 5 is admin, 6 is editor
      .orderBy("Order", true)
      ().then((roleDefs: IRoleDefinitionInfo[]) => {
        console.log(roleDefs);
        setRoleDefinitionInfos(roleDefs);
      }).catch(err => {

        console.log(err);
      });
  }
  async function getListViews(sp, viewId: string) {
    await sp.web.lists
      .getById(props.context.pageContext.list.id.toString())
      .views().then(views => {

        setAllViews(views.filter(v => v.Hidden === false));
        if (!viewId) {
          const viewFromPageUrl = find(views, (v) => {
            return v.ServerRelativeUrl === decodeURIComponent(document.location.pathname);
          });
          if (viewFromPageUrl) {
            setViewId(viewFromPageUrl.Id);
          }

          // dunno what view to use, so use the first one
          else {
            setViewId(views[0].Id);
          }
        }
      });
  }
  useEffect(() => {
    // declare the data fetching function
    const fetchData = async () => {
      let locShareType: ShareType;
      const sp = spfi().using(SPFx(props.context));
      const urlParams = new URLSearchParams(window.location.search);
      debugger;
      let locFolderServerRelativePath = urlParams.get("id")
      const locViewId = urlParams.get("viewid");
      const locListId = props.context.pageContext.list.id.toString();
      let locItemId: number;

      //  figure out what type of share we are dealing with
      if (props.event.selectedRows.length === 1) {
        locItemId = parseInt(props.event.selectedRows[0].getValueByName("ID"))
        // they selected an item. Need to see if its a folder or a documnent
        debugger;
        let locItem: IItem = await sp.web.lists
          .getById(locListId)
          .items.getById(locItemId)
          .expand("File", "Folder")
          .select("Title", "EffectiveBasePermissions", "FileSystemObjectType", "File/ServerRelativeUrl", "Folder/ServerRelativeUrl")
          .expand("File", "Folder")
          ();
        setUserCanManagePermissions(sp.web.hasPermissions(locItem["EffectiveBasePermissions"], PermissionKind.ManagePermissions));
        debugger;
        if (locItem["FileSystemObjectType"] == 1) {
          // its a folder

          setShareType(ShareType.Folder);
          setFolderServerRelativePath(locItem["Folder"]["ServerRelativeUrl"]);
          // see if user has permissions to share this folder

        } else {
          // its a document
          setShareType(ShareType.File);
        }
      } else {
        debugger;
        if (locFolderServerRelativePath) {
          // they selected a folder
          setFolderServerRelativePath(locFolderServerRelativePath);
          setShareType(ShareType.Folder);
          sp.web.getFolderByServerRelativePath(locFolderServerRelativePath).select("Title", "EffectiveBasePermissions")()
            .then(folder => {
              setUserCanManagePermissions(sp.web.hasPermissions(folder["EffectiveBasePermissions"], PermissionKind.ManagePermissions));
            });
        } else {
          // they selected a document
          setShareType(ShareType.Library)
          sp.web.lists.getById(locListId).select("Title", "EffectiveBasePermissions")()
            .then(folder => {
              setUserCanManagePermissions(sp.web.hasPermissions(folder["EffectiveBasePermissions"], PermissionKind.ManagePermissions));
            });
        }

      }

      setLibraryName(props.context.pageContext.list.title);
      setTabName(props.context.pageContext.list.title);
      setViewId(locViewId);
      await getListViews(sp, locViewId);
      await getRoleDefs(sp);
      // if (props.event.selectedRows.length === 1) {
      //   // they selected an item. Need to see if its a folder or a documnent
      //   debugger;
      //   const locList: IListInfo = await sp.web.lists
      //     .getById(props.context.pageContext.list.id.toString())();
      //   const locItem = await sp.web.lists
      //     .getById(props.context.pageContext.list.id.toString())
      //     .items.getById(parseInt(props.event.selectedRows[0].getValueByName("ID")))
      //     .expand("File")();
      //   debugger;
      //   setList(locList);

      //   if (locItem.FileSystemObjectType === 1) {
      //     // its a folder
      //     locShareType = ShareType.Folder;



      //   } else {
      //     // its a document
      //     locShareType = ShareType.File;
      //   }
      // } else {
      //   if (locFolderServerRelativePath) {
      //     // they selected a folder
      //     locShareType = ShareType.Folder;
      //   } else {
      //     // they selected a document
      //     locShareType = ShareType.Library;
      //   }
    }

    // ok now that we figured out what we are sharing, lets see if they have permissions to share it
    // switch (locShareType) {
    //   case ShareType.Library:
    //     await sp.web.lists
    //       .getById(props.context.pageContext.list.id.toString())
    //       .currentUserHasPermissions(PermissionKind.ManagePermissions).then((hasPermissions) => {
    //         setUserCanManagePermissions(hasPermissions);
    //       });
    //     // await sp.web.lists.getById(props.context.pageContext.list.id.toString()).
    //     //   effectiveBasePermissions()
    //     //   .then(permissions => {
    //     //     setUserCanManagePermissions(permissions.has(PermissionKind.ManagePermissions));
    //     //     setPermissionsOnSP(permissions);

    //     //   }).catch(err => {
    //     //     debugger;
    //     //     console.log(err);
    //     //   });
    //     break;
    //   case ShareType.Folder:

    //     const locFolder = await sp.web.getFolderByServerRelativePath(folderServerRelativePath).getItem();
    //     locFolder.effectiveBasePermissions().then((permissions) => {
    //       setPermissionsOnSP(permissions);
    //       setUserCanManagePermissions(permissions.has(PermissionKind.ManagePermissions));
    //     }).catch(err => {
    //       debugger;
    //       console.log(err);
    //     });
    //     break;
    //   case ShareType.File:
    //     break;

    // }


    // call the function
    fetchData()
      // make sure to catch any error
      .catch(console.error);
  }, []);
  return (
    <DialogContent
      title={"Share to Teams"}
      onDismiss={props.close}
      showCloseButton={true}
    >

      <div>
        ShareType is {ShareType[shareType]}<br />
        Library  is {libraryName}<br />
        folderServerRelativePath is {folderServerRelativePath}<br />
        ViewId is {viewId}<br />
        userCanManagePermissions is {userCanManagePermissions ? "true" : "false"}<br />
        <TeamPicker label={`What Team would you like to share this ${ShareType[shareType]} to?`}
          selectedTeams={selectedTeams}
          appcontext={props.context}
          itemLimit={1}
          onSelectedTeams={_onSelectedTeams} />

        <TeamChannelPicker label={`What Channel would you like to share this ${ShareType[shareType]}  to?`}
          teamId={selectedTeams.length > 0 ? selectedTeams[0].key : null}
          selectedChannels={selectedTeamChannels}
          appcontext={props.context}
          itemLimit={1}
          onSelectedChannels={_onSelectedTeamChannels} />
        <ChoiceGroup
          label="Which view would you like to show in the Teams Tab?"
          title="View"
          options={allViews.map(view => { return { key: view.Id, text: view.Title } })}
          defaultSelectedKey={viewId}
          selectedKey={viewId}
          onChange={(e, o) => { setViewId(o.key) }}
        />
        <ChoiceGroup
          label={`What permission like give to the members of the ${selectedTeams.length == 0 ? "" : selectedTeams[0].name} team to this ${ShareType[shareType]} ?`}
          title="View"
          options={roleDefinitionInfos.map((rd) => {
            return { key: rd.Id.toString(), text: `${rd.Name} (${rd.Description})` };
          })}
          defaultSelectedKey={viewId}
          selectedKey={viewId}
          onChange={(e, o) => { setViewId(o.key) }}
        />
        <TextField label="What would you like the text in the Teams Tab to say?" onChange={(e, newValue) => { setTabName(newValue) }} value={tabName} />
        <PrimaryButton onClick={addTab}> Add Tab to Team</PrimaryButton>
      </div>




    </DialogContent>
  );

}

export default class ShareToTeamsDialog extends BaseDialog {

  public title: string;
  public event: IListViewCommandSetExecuteEventParameters;
  public aadHttpClient: AadHttpClient;
  public context: BaseComponentContext;
  public render(): void {
    ReactDOM.render(
      <ShareToTeamsContent
        event={this.event}
        aadHttpClient={this.aadHttpClient}
        title="SS"
        context={this.context}
        close={this.close}
      />,
      this.domElement
    );
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: true,
    };
  }

  protected onAfterClose(): void {
    super.onAfterClose();
    // Clean up the element for the next dialog
    ReactDOM.unmountComponentAtNode(this.domElement);
  }
}


