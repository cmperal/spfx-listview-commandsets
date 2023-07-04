import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  ListViewAccessorStateChanges,
  ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICommandSetButtonsCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'CommandSetButtonsCommandSet';

export default class CommandSetButtonsCommandSet extends BaseListViewCommandSet<ICommandSetButtonsCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized CommandSetButtonsCommandSet');

    // initial state of the command's visibility
    const buttonAlwaysOn: Command = this.tryGetCommand('ALWAYS_ON');
    buttonAlwaysOn.visible = true;

    const buttonOne: Command = this.tryGetCommand('ONE_ITEM_SELECTED');
    buttonOne.visible = false;

    const buttonTwo: Command = this.tryGetCommand('TWO_ITEM_SELECTED');
    buttonTwo.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    switch (event.itemId) {
      case 'ONE_ITEM_SELECTED':
        Dialog.alert(`${this.properties.sampleTextOne}`).catch(() => {
          /* handle error */
        });
        break;
      case 'TWO_ITEM_SELECTED':
        Dialog.alert(`${this.properties.sampleTextTwo}`).catch(() => {
          /* handle error */
        });
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    if (args.stateChanges !== ListViewAccessorStateChanges.SelectedRows) {
      return;
    }

    const oneItemSelected: Command = this.tryGetCommand('ONE_ITEM_SELECTED');
    oneItemSelected.visible = this.context.listView.selectedRows!.length === 1;

    const twoItemSelected: Command = this.tryGetCommand('TWO_ITEM_SELECTED');
    twoItemSelected.visible = this.context.listView.selectedRows!.length > 1;



    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
