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
import { spfi, SPFx } from "@pnp/sp";
import * as strings from 'ShareToTeamsCommandSetStrings';

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
  private aadHttpClient: AadHttpClient;
  @override
  public onInit(): Promise<void> {
    debugger;
    return super.onInit().then((_) => {
      debugger;
      const sp = spfi().using(SPFx(this.context));

      return this.context.aadHttpClientFactory
        //.getClient(this.properties.hellosignFunctionClientID)
        .getClient("becb3efa-2875-4eda-8fb8-40d03f7cb4e7")
        .then((client): void => {
          // connect to the API
          debugger;
          this.aadHttpClient = client;
      
        })
        .catch((e) => {
          debugger;
        });
    });
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
       debugger;
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private cmdShareToTeams(event: IListViewCommandSetExecuteEventParameters) {
    const dialog: ShareToTeamsDialog = new ShareToTeamsDialog();
    dialog.title = `CHECK STATUS`;
    dialog.aadHttpClient = this.aadHttpClient;
    dialog.context = this.context;
    dialog.event = event;
    dialog.show();
  }
}
