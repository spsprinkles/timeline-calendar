import * as React from 'react';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { MonacoEditor } from "@pnp/spfx-controls-react/lib/MonacoEditor";

interface IMonacoPanelEditorProps {
  key: string;
  disabled: boolean;
  buttonText?: string;
  headerText: string;
  description?: string;
  value: string;
  language: "html" | "css" | "json" | "typescript";
  onValueChanged: (newValue: string) => void;
}

export interface IMonacoPanelEditorState {
  isOpen: boolean;
  //saveDisabled: boolean;
}
 /*
 When instantiating this within the PropPaneConfig...
  return (  
    React.createElement(MonacoPanelEditor, {

...you can define initial value to load in editor...
  if (value == null && item.Name != null)
    value = "<b>" + item.Name + "</b>";

...but it's not saved unless onUpdate called, and even if the user opens the editor
and clicks Save, it doesn't populate newValue if they don't type in anything.
      onValueChanged: (newValue: string) =>
*/

export default class MonacoPanelEditor extends React.Component<IMonacoPanelEditorProps, IMonacoPanelEditorState> {
    constructor(props: IMonacoPanelEditorProps, state: IMonacoPanelEditorState) {
      super(props);
      this.state = {
        isOpen: false,
        //saveDisabled: false
      };

      //To make this.setState work correctly
      this.btnHtmlHandler = this.btnHtmlHandler.bind(this);
    }
  
    //When component first loads
    public componentDidMount(): void {
      //Nothing
    }
  
    public componentDidUpdate(prevProps: IMonacoPanelEditorProps, prevState: IMonacoPanelEditorState): void {
      //Nothing
    }
  
    //Temporarily store value of what's changed in the editor (before saving back to main object)
    private tempValue: string;

    public render(): React.ReactElement<IMonacoPanelEditorProps> {//was using JSX.Element {
      const styles = { paddingTop: "20px" };
      return (
        <div>
          <DefaultButton //was PrimaryButton which had color
            text={this.props.buttonText || "Configure"}
            disabled={this.props.disabled}
            onClick={this.btnHtmlHandler} />
          <Panel
            headerText={this.props.headerText}
            type={PanelType.custom} //was medium
            customWidth="800px"
            isOpen={this.state.isOpen}
            onDismiss={() => this.closePanel(false)}
            isLightDismiss={false}
            isBlocking={true}
            closeButtonAriaLabel='Close'
            onRenderFooterContent={this.onRenderFooterContent}
            isFooterAtBottom={true}>
            <div style={styles}>{this.props.description}</div>
            <MonacoEditor
                key={this.props.key}
                value={this.props.value}
                showLineNumbers={true}
                showMiniMap={false}
                onValueChange={(newValue: string) => {
                    //Value is only temp right now; see "closePanel" function
                    this.tempValue = newValue;
                    //Tried checking for invalid JSON input to disable the Save button,
                    //but onValueChanged is *not* called if user changes JSON back to original input after changing
                }}
                language={this.props.language}
                theme='vs-dark' />
        </Panel>
        </div>
      );
    }
  
    private btnHtmlHandler(): void {
        this.setState({isOpen: true});
    }

    // private saveBtnDisabled(value:boolean): void {
    //   if (this.state.saveDisabled != value)
    //     this.setState({saveDisabled: value});
    // }

    //Checks if content should be saved (only when Save button is clicked)
    private closePanel = (doSave: boolean) => {
        this.setState({isOpen: false});

        //Should the value be saved?
        if (doSave)
          this.props.onValueChanged(this.tempValue);
    }

    //<PrimaryButton onClick={(e) => this.closePanel(e, true)} styles={this.buttonStyles}>Save</PrimaryButton>
    //<PrimaryButton onClick={e => this.closePanel(e, true)} styles={this.buttonStyles}>Save</PrimaryButton>
    private onRenderFooterContent = () => {
      const buttonStyles = { root: {marginRight: 8 }};
      return (
        <div>
          <PrimaryButton onClick={() => this.closePanel(true)} styles={buttonStyles}>Save</PrimaryButton>
          <DefaultButton onClick={() => this.closePanel(false)}>Cancel</DefaultButton>
        </div>
      )
    }
  }