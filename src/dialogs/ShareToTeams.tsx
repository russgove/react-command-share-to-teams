import { TeamsTab } from "@microsoft/microsoft-graph-types";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { BaseDialog, IDialogConfiguration } from "@microsoft/sp-dialog";
import { AadHttpClient } from "@microsoft/sp-http";
import { IListViewCommandSetExecuteEventParameters } from "@microsoft/sp-listview-extensibility";
import { graphfi, SPFx as SPFxGR } from "@pnp/graph";
import "@pnp/graph/teams";
import "@pnp/graph/";
import "@pnp/graph/users";
import "@pnp/graph/groups";
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
import { find, forEach } from "lodash";
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
  const graph = graphfi().using(SPFxGR(props.context));
  const [shareType, setShareType] = React.useState<ShareType>(null);
  const [list, setList] = React.useState<IListInfo>(null);
  const [item, setItem] = React.useState<any>(null);
  const [canManageTabs, setCanManageTabs] = React.useState<boolean>(false);
  const [selectedTeam, setSelectedTeam] = React.useState<ITag[]>([]);
  const [selectedTeamChannels, setSelectedTeamChannels] = React.useState<ITag[]>([]);
  const [roleDefinitionInfos, setRoleDefinitionInfos] = React.useState<IRoleDefinitionInfo[]>([]);
  const [selectedRoleDefinitionId, setSelectedRoleDefinitionId] = React.useState<number>(null);
  const [folderServerRelativePath, setFolderServerRelativePath] = React.useState<string>(null);
  const [userCanManagePermissions, setUserCanManagePermissions] = React.useState<boolean>(false);
  const [allViews, setAllViews] = React.useState<IViewInfo[]>([]);
  const [selectedViewId, setSelectedViewId] = React.useState<string>(null);
  const [tabName, setTabName] = React.useState<string>("");
  const [libraryName, setLibraryName] = React.useState<string>("");
  const [permissionsOnSP, setPermissionsOnSP] = React.useState<IBasePermissions>(null);
  async function addTab() {

    debugger;

    const teamId: string = selectedTeam[0].key as string;
    const channelId: string = selectedTeamChannels[0].key as string;
    console.log(`TEAM ID is ${teamId}. CHANNEL ID is ${channelId}`);
    const team = await graph.teams.getById(teamId)();
    console.log(team);
    const channel = await graph.teams.getById(teamId).channels.getById(channelId);
    console.log(channel);
    const channelTabs = await graph.teams.getById(teamId).channels.getById(channelId).tabs;
    console.log(channelTabs);
    const teamsTab: TeamsTab = {} as TeamsTab;
    teamsTab.displayName = tabName;
    let contentUrl = "";
    switch (shareType) {
      case ShareType.Library:
        let lView = find(allViews, (view) => view.Id === selectedViewId)
        contentUrl = `${document.location.origin}${lView.ServerRelativeUrl}`;
        //grantTeamMembersAcessToLibrary(teamId, channelId, contentUrl);
        debugger
        break;
      case ShareType.Folder:
        let fview = find(allViews, (view) => view.Id === selectedViewId)
        //grantTeamMembersAcessToFolder(teamId, channelId, folderServerRelativePath, fview.ServerRelativeUrl);
        contentUrl = `${document.location.origin}${fview.ServerRelativeUrl}?id=${folderServerRelativePath}`;

        break;
      case ShareType.File:

        const sp = spfi().using(SPFx(props.context));

        contentUrl = await sp.web.lists.getById(props.context.pageContext.list.id.toString())
          .items.getById(item["Id"]).getWopiFrameUrl(0);//read only in word
        contentUrl = await sp.web.lists.getById(props.context.pageContext.list.id.toString())
          .items.getById(item["Id"]).getWopiFrameUrl(1);//update mode in word
        contentUrl = await sp.web.lists.getById(props.context.pageContext.list.id.toString())
          .items.getById(item["Id"]).getWopiFrameUrl(2);//read only in word
        contentUrl = await sp.web.lists.getById(props.context.pageContext.list.id.toString())
          .items.getById(item["Id"]).getWopiFrameUrl(3);
        //https://graph.microsoft.com/v1.0/sites/russellwgove.sharepoint.com:/sites/CR-EU-Manufacturing:/drives
        break;

    }

    teamsTab.configuration = {
      contentUrl: contentUrl,
    }
    const newTab = channelTabs.add('Tab', 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/2a527703-1f6f-4559-a332-d8a7d288cd88', teamsTab)
      .then((t) => {
        ;
        console.log(t);
        channel.messages({body: {content:`New tab ${tabName} has been added to channel ${channelId}`}});
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
            setSelectedViewId(viewFromPageUrl.Id);
          }

          // dunno what view to use, so use the first one
          else {
            setSelectedViewId(views[0].Id);
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
      //TODO: save view enhancements to state and reapply isAscending=true sortField=LinkFilenameFilterFields1=testcol1 FilterValues1=a%3B%23b FilterTypes1=Text       let locFolderServerRelativePath = urlParams.get("id")

      let locFolderServerRelativePath = urlParams.get("id")
      const locViewId = urlParams.get("viewid");
      const locListId = props.context.pageContext.list.id.toString();
      let locItemId: number;

      //  figure out what type of share we are dealing with
      if (props.event.selectedRows.length === 1) {
        locItemId = parseInt(props.event.selectedRows[0].getValueByName("ID"))
        // they selected an item. Need to see if its a folder or a documnent
        let locItem: IItem = await sp.web.lists
          .getById(locListId)
          .items.getById(locItemId)
          .expand("File", "Folder")
          .select("Id", "Title", "EffectiveBasePermissions", "FileSystemObjectType", "ServerRedirectedEmbedUrl", "File/LinkingUrl", "File/ServerRelativeUrl", "Folder/ServerRelativeUrl")
          .expand("File", "Folder")
          ();
        setUserCanManagePermissions(sp.web.hasPermissions(locItem["EffectiveBasePermissions"], PermissionKind.ManagePermissions));

        if (locItem["FileSystemObjectType"] == 1) {
          // its a folder

          setShareType(ShareType.Folder);
          setFolderServerRelativePath(locItem["Folder"]["ServerRelativeUrl"]);
          setTabName(props.context.pageContext.list.title);// see if user has permissions to share this folder

        } else {
          // its a document
          setItem(locItem);
          setShareType(ShareType.File);
          setTabName(locItem["File"]["Name"]);
        }
      } else {

        if (locFolderServerRelativePath) {
          // they are within a folder
          setFolderServerRelativePath(locFolderServerRelativePath);
          setShareType(ShareType.Folder);
          sp.web.getFolderByServerRelativePath(locFolderServerRelativePath)
            .expand("ListItemAllFields/EffectiveBasePermissions")()
            .then(folder => {
              setUserCanManagePermissions(sp.web.hasPermissions(folder["ListItemAllFields"]["EffectiveBasePermissions"], PermissionKind.ManagePermissions));
            });
        } else {
          // they are at the root of the list
          setShareType(ShareType.Library)
          sp.web.lists.getById(locListId).select("Title", "EffectiveBasePermissions")()
            .then(folder => {
              setUserCanManagePermissions(sp.web.hasPermissions(folder["EffectiveBasePermissions"], PermissionKind.ManagePermissions));
            });
        }

      }

      setLibraryName(props.context.pageContext.list.title);

      setSelectedViewId(locViewId);
      await getListViews(sp, locViewId);
      await getRoleDefs(sp);

    }
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
        ViewId is {selectedViewId}<br />
        userCanManagePermissions is {userCanManagePermissions ? "true" : "false"}<br />
        selectedRoleDefinitionId is {selectedRoleDefinitionId}<br />
        canManageTabs is {canManageTabs ? "true" : "false"}<br />
        <TeamPicker label={`What Team would you like to share this ${ShareType[shareType]} to?`}
          selectedTeams={selectedTeam}
          appcontext={props.context}
          itemLimit={1}
          onSelectedTeams={(tagList: ITag[]) => {
            setSelectedTeam(tagList);
            graph.teams.getById(tagList[0].key.toString())()
            .then(team => {
              debugger;
              if(team.memberSettings.allowCreateUpdateRemoveTabs){
                setCanManageTabs(true);
              }
              else{
                graph.groups.getById(tagList[0].key.toString()).expand("owners").select("owners")()
                .then(group => {
                  debugger;
                  // if user is owner of the group, then they can manage tabs
                  for (const owner of group.owners) {
                    if(owner["userPrincipalName"].toLowerCase() === props.context.pageContext.user.loginName.toLowerCase()){
                      setCanManageTabs(true);
                      return;
                    }
                  }
                  setCanManageTabs(false);
                })
                
               
                .catch(err => {
                  console.log(err)
                });
             
              }
            })
            .catch(err => {
              console.log(err);
            });
           
          }}
        />

        <TeamChannelPicker label={`What Channel would you like to share this ${ShareType[shareType]}  to?`}
          teamId={selectedTeam.length > 0 ? selectedTeam[0].key : null}
          selectedChannels={selectedTeamChannels}
          appcontext={props.context}
          itemLimit={1}
          onSelectedChannels={(tagList: ITag[]) => {
            setSelectedTeamChannels(tagList);
          }} />
        <ChoiceGroup
          label="Which view would you like to show in the Teams Tab?"
          title="View"
          options={allViews.map(view => { return { key: view.Id, text: view.Title } })}
          defaultSelectedKey={selectedViewId}
          selectedKey={selectedViewId}
          onChange={(e, o) => { setSelectedViewId(o.key) }}
        />
        <ChoiceGroup
          label={`What permission like give to the members of the ${selectedTeam.length == 0 ? "" : selectedTeam[0].name} team to this ${ShareType[shareType]} ?`}
          title="View"
          options={roleDefinitionInfos.map((rd) => {
            return { key: rd.Id.toString(), text: `${rd.Name} (${rd.Description})` };
          })}
          defaultSelectedKey={selectedRoleDefinitionId}
          selectedKey={selectedRoleDefinitionId ? selectedRoleDefinitionId.toString() : null}
          onChange={(e, o) => {
            debugger;
            setSelectedRoleDefinitionId(parseInt(o.key))
          }}
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


