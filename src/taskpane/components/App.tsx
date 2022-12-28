import * as React from "react";
import { ActionButton, PartialTheme, Separator, Stack, TextField } from "@fluentui/react";
import Progress from "./Progress";

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
  itemChangedRegister: any;
  theme: PartialTheme;
}

export interface AppState {
  saveDisabled: boolean,
  customProps: any,
  notes: string,
  fatalError: string,
  logs: string[],
}

const DEBUG = false;

// notes:
// The maximum length of a CustomProperties JSON object is 2500 characters.
// Outlook on Mac doesn't cache custom properties. If the user's network goes down, mail add-ins can't access their custom properties.


export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    if (props.itemChangedRegister) {
      props.itemChangedRegister(this.onItemChanged);
    }
    this.state = {
      saveDisabled: true,
      customProps: null,
      notes: '',
      fatalError: '',
      logs: [],
    };
  }

  onItemChanged = () => {
    this.log("item changed");
    this.setState((prevState) => ({
      ...prevState,
      customProps: null,
      notes: null
    }));
    this.loadNotes();
  }

  log = (msg) => {
    if (DEBUG) {
      console.log(msg);
      this.setState((prevState) => ({ ...prevState, logs: [...prevState.logs, msg] }))
    }
  }

  componentDidUpdate(prevProps: Readonly<AppProps>, _prevState: Readonly<AppState>, _snapshot?: any): void {
    if (!prevProps.isOfficeInitialized && this.props.isOfficeInitialized) {
      this.loadNotes();
    }
  }

  componentDidMount(): void {
    if (this.props.isOfficeInitialized) this.loadNotes();
  }

  save = () => {
    this.state.customProps.set("notes", this.state.notes);
    this.state.customProps.saveAsync((result) => {
      this.log(JSON.stringify(result.status));
      if (result.status == Office.AsyncResultStatus.Failed) {
        this.log('could not save notes');
      } else {
        this.log('notes saved');
        this.setState((prevState) => ({ ...prevState, saveDisabled: true }))
      }
    })
  };

  loadNotes = () => {
    Office.context.mailbox.item.loadCustomPropertiesAsync((result: Office.AsyncResult<Office.CustomProperties>) => {
      if (result.status == Office.AsyncResultStatus.Failed) {
        this.setState((prevState) => ({ ...prevState, fatalError: 'Failed to load notes' }))
      } else {
        let customProps = result.value;
        let notes = customProps.get("notes");
        this.log("loaded notes: " + notes);
        this.setState((prevState) => ({ ...prevState, customProps, notes }))
      }
    })
  }

  onChange = (_ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    this.setState((prevState) => ({ ...prevState, notes: newText, saveDisabled: false }));
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (this.state.fatalError) {
      return (
        <div className="main">
          <p className="error-message">
            {this.state.fatalError}
          </p>
        </div>
      );
    }

    if (!isOfficeInitialized || !this.state.customProps) {
      return (
        <div className="progress_center">
          <Progress
            title={title}
            logo={require("./../../../assets/logo-filled.png")}
            message=""
          />
        </div>
      );
    }

    return (
      <div>
        {
          DEBUG ? (
            <div>
              {Office.context.mailbox.item.itemId}
              <br></br>
              {Office.context.mailbox.item.subject}
              <br></br>
              {Office.context.mailbox.item.conversationId}
              <br></br>
              {Office.context.displayLanguage}
              <br></br>
            </div>
          ) : null
        }
        <Stack horizontalAlign="end">
          <ActionButton iconProps={{ iconName: 'Save' }} allowDisabledFocus disabled={this.state.saveDisabled} onClick={this.save}>
            Save
          </ActionButton>
        </Stack>
        <Separator styles={{root: { padding: 0 }}}/>
        <div>
          <TextField placeholder="Add notes.." multiline autoAdjustHeight value={this.state.notes} rows={10} onChange={this.onChange}
          borderless
          />
        </div>
        {this.state.logs.length > 0 ? (
          <div>
            {this.state.logs.map(log => (<p>{JSON.stringify(log)}</p>))}
          </div>
        ) : null}
      </div>
    );
  }
}

