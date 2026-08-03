import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { DialogFooter, PrimaryButton, DefaultButton, DialogContent } from '@fluentui/react';

interface IConfirmationDialogContentProps {
  message: string;
  title?: string;
  close: () => void;
  submit: () => void;
}

class ConfirmationDialogContent extends React.Component<IConfirmationDialogContentProps, {}> {
  public render(): JSX.Element {
    return (
      <DialogContent
        title={this.props.title}
        subText={this.props.message}
        onDismiss={this.props.close}
        showCloseButton={true}
      >
        <DialogFooter>
          <PrimaryButton text="Yes, Cancel Request" onClick={this.props.submit} />
          <DefaultButton text="No, Keep It" onClick={this.props.close} />
        </DialogFooter>
      </DialogContent>
    );
  }
}

export default class ConfirmationDialog extends BaseDialog {
  public message: string;
  public title: string;
  public isConfirmed: boolean = false;

  constructor(message: string, title: string = "Confirm Action") {
    super();
    this.message = message;
    this.title = title;
  }

  public render(): void {
    ReactDOM.render(
      <ConfirmationDialogContent
        message={this.message}
        title={this.title}
        close={this._close}
        submit={this._submit}
      />,
      this.domElement
    );
  }

  public getConfig(): IDialogConfiguration {
    return { isBlocking: false };
  }

  private _close = (): void => {
    this.isConfirmed = false;
    this.close().catch(console.error);
  }

  private _submit = (): void => {
    this.isConfirmed = true;
    this.close().catch(console.error);
  }

  protected onAfterClose(): void {
    super.onAfterClose();
    // Clean up
    ReactDOM.unmountComponentAtNode(this.domElement);
  }
}
