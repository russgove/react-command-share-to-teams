import { BaseDialog, Dialog, IDialogConfiguration } from "@microsoft/sp-dialog";
import { AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";



import { DetailsList, SelectionMode, Selection } from "office-ui-fabric-react/lib/DetailsList";
import { spfi, SPFx } from "@pnp/sp";
import { DialogContent } from "office-ui-fabric-react/lib/Dialog";
import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';

import { Link } from "office-ui-fabric-react/lib/Link";
import { MessageBar, MessageBarType, } from "office-ui-fabric-react/lib/MessageBar";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";

import { TextField } from "office-ui-fabric-react/lib/TextField";
import * as React from "react";
import * as ReactDOM from "react-dom";


import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { IListViewCommandSetExecuteEventParameters } from "@microsoft/sp-listview-extensibility";

interface IShareToTeamsProps {

  title: string;
  close: () => void;
  aadHttpClient: AadHttpClient;
  context: BaseComponentContext;
  event: IListViewCommandSetExecuteEventParameters;
}
interface IShareToTeamsState {
  errors: Array<string>;
  selectionDetails: string;
}

class ShareToTeamsContent extends React.Component<IShareToTeamsProps, IShareToTeamsState> {
  private _selection: Selection;
  constructor(props: IShareToTeamsProps) {
    super(props);
    this._selection = new Selection({
    });
    this.state = {
      errors: [],

      selectionDetails: "",
    };
  }

  public componentDidMount() {


  }
  public render(): JSX.Element {
    const _items: ICommandBarItemProps[] = [

      {
        key: 'View',
        text: 'View on HelloSign',
        iconProps: { iconName: 'View' }

      },

    ];
    return (
      <DialogContent
        title={"Share to Teams"}

        showCloseButton={true}
      >
        <div>
          {this.state.errors.map((error, i) => {
            console.log("Entered");
            // Return the errors
            return (
              <MessageBar messageBarType={MessageBarType.error}>
                {error}{" "}
              </MessageBar>
            );
          })}
        </div>



        <Panel
          type={PanelType.large} headerText="HelloSign Status"
          onDismiss={(e) => {

            this.setState((current) => ({
              ...current,
              signatureRequestFromHS: null,
            }));
          }}
        >
          <CommandBar items={_items} />




          <TextField
            label="Requester"
            value={"[]"}
            borderless={true}
          />
          <DetailsList selection={this._selection}
            items={[]}
            // layoutMode={DetailsListLayoutMode.fixedColumns}
            selectionMode={SelectionMode.single}
            columns={[
              {
                key: "signerName",
                name: "Name",
                fieldName: "signerName",
                minWidth: 200,

                isResizable: true,
              },


            ]}
          ></DetailsList>
        </Panel>


      </DialogContent>
    );
  }
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
