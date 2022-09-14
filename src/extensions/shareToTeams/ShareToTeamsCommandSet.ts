import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
//import ShareToTeamsDialog from "../../dialogs/ShareToTeams";
import {
  ShareToTeamsContent,
  IShareToTeamsProps,
} from "../../components/ShareToTeams";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
} from "@microsoft/sp-listview-extensibility";
import { Dialog } from "@microsoft/sp-dialog";
import {
  AadHttpClient,
  HttpClientResponse,
  MSGraphClient,
  AadHttpClientConfiguration,
} from "@microsoft/sp-http";

import "@pnp/graph/users";
import { spfi, SPFx } from "@pnp/sp";

import * as strings from "ShareToTeamsCommandSetStrings";
import { graphfi } from "@pnp/graph";
import { SPFx as SPFxgr } from "@pnp/graph";
import * as ReactDOM from "react-dom";
import * as React from "react";
import { assign } from "lodash";
import { BaseComponentContext } from "@microsoft/sp-component-base";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IShareToTeamsCommandSetProperties {
  supportedFileTypes: string; //tenantproperties?
  allowListSharing: boolean;
  allowFolderSharing: boolean;
  allowFileSharing: boolean;
  librarySharingMethod: string; // "native" attempts to use the navis teams app. "page" just opens a sharepoint page
  folderSharingMethod: string;
  fileSharingMethod: string;
}

const LOG_SOURCE: string = "ShareToTeamsCommandSet";

export default class ShareToTeamsCommandSet extends BaseListViewCommandSet<IShareToTeamsCommandSetProperties> {
  private msGraphClient: MSGraphClient;
  private panelPlaceHolder: HTMLDivElement = null;
  private panelProps:IShareToTeamsProps;
  
  @override
  public async onInit(): Promise<void> {
    await super.onInit();
    await this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        this.msGraphClient = client;
      });
    // Create the container for our React component
    this.panelPlaceHolder = document.body.appendChild(
      document.createElement("div")
    );
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(
    event: IListViewCommandSetListViewUpdatedParameters
  ): void {
    const shareToTeamsCommand: Command = this.tryGetCommand(
      "COMMAND_SHARE_TO_TEAMS"
    );
   
    if (shareToTeamsCommand) {
      if (event.selectedRows.length == 1) {
        //
        switch (event.selectedRows[0].getValueByName("FSObjType")) {
          //one row selected
          case "0":
            //its a file
            if (
              this.properties.supportedFileTypes.indexOf(
                event.selectedRows[0].getValueByName("File_x0020_Type")
              ) !== -1 &&
              this.properties.allowFileSharing
            ) {
              shareToTeamsCommand.visible = true;
            } else {
              shareToTeamsCommand.visible = false;
            }
            break;
          case "1":
            //its a folder
            shareToTeamsCommand.visible = this.properties.allowFolderSharing;
            break;
          default:
            shareToTeamsCommand.visible = false;
        }
      } else {
        if (event.selectedRows.length > 1 || event.selectedRows.length < 0) {
          shareToTeamsCommand.visible = false;
        } else {
          //no rows selected are they at the top or in a folder
          const urlParams = new URLSearchParams(window.location.search);
          if (urlParams.get("id")) {
            // in a folder
            shareToTeamsCommand.visible = this.properties.allowFolderSharing;
          } else {
            // at root
            shareToTeamsCommand.visible = this.properties.allowListSharing;
          }
        }
      }
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case "COMMAND_SHARE_TO_TEAMS":
        this.cmdShareToTeams(event);
        break;
      default:
        throw new Error("Unknown command");
    }
  }

  private cmdShareToTeams(event: IListViewCommandSetExecuteEventParameters) {
    debugger;
    this.panelProps = {
      event: event,
      msGraphClient: this.msGraphClient,
      settings: this.properties,
      context: this.context,
      onClose: this._dismissPanel.bind(this),
      isOpen:true
    };
    this._showPanel();
  }
  private _showPanel() {
    debugger;
    this._renderPanelComponent();
  }

  private _dismissPanel() {
    debugger;
    this.panelProps.isOpen=false;
    this._renderPanelComponent();
  }

  private _renderPanelComponent() {
    debugger;
    const element: React.ReactElement<IShareToTeamsProps> = React.createElement(
      ShareToTeamsContent,
      this.panelProps
    );
    ReactDOM.render(element, this.panelPlaceHolder);
  }
}
