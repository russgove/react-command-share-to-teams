import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import ShareToTeamsDialog, {} from "../../dialogs/ShareToTeams";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import {
  AadHttpClient,
  HttpClientResponse,
  AadHttpClientConfiguration,
} from "@microsoft/sp-http";

import "@pnp/graph/users";
import { spfi, SPFx } from "@pnp/sp";

import * as strings from 'ShareToTeamsCommandSetStrings';
import { graphfi } from '@pnp/graph';
import { SPFx as SPFxgr } from '@pnp/graph';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IShareToTeamsCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'ShareToTeamsCommandSet';

export default class ShareToTeamsCommandSet extends BaseListViewCommandSet<IShareToTeamsCommandSetProperties> {
 // private aadHttpClient: AadHttpClient;
  @override
  public async onInit(): Promise<void> {
    debugger;
   
    await super.onInit();
   
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const shareToTeamsCommand: Command = this.tryGetCommand('COMMAND_SHARE_TO_TEAMS');
    if (shareToTeamsCommand) {
      // This command should be hidden unless exactly one row is selected.
      shareToTeamsCommand.visible = event.selectedRows.length <= 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_SHARE_TO_TEAMS':
       this.cmdShareToTeams(event);
         break;
      default:
        throw new Error('Unknown command');
    }
  }

  private cmdShareToTeams(event: IListViewCommandSetExecuteEventParameters) {
    const dialog: ShareToTeamsDialog = new ShareToTeamsDialog();
    dialog.title = `CHECK STATUS`;
    dialog.context = this.context;
    dialog.event = event;
    dialog.show();
  }
}
