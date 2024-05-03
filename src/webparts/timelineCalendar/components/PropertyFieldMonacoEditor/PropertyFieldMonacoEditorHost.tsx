import * as React from 'react';

import {DefaultButton, mergeStyles, mergeStyleSets, Panel, PanelType, PrimaryButton, Stack, TextField} from '@fluentui/react';
//import strings from 'PropertyControlStrings';
//import * as telemetry from '../../common/telemetry';
import {IPropertyFieldMonacoEditorHostProps, IPropertyFieldMonacoEditorHostState} from './IPropertyFieldMonacoEditorHost';
//import { MonacoEditor } from './monacoEditorControl';
import { MonacoEditor} from '../MonacoEditor';

const DEFAULT_PANEL_WIDTH = "800px";

export default class PropertyFieldMonacoEditorHost extends React.Component<
  IPropertyFieldMonacoEditorHostProps,
  IPropertyFieldMonacoEditorHostState
> {
  constructor(props: IPropertyFieldMonacoEditorHostProps) {
    super(props);
    //telemetry.track("PropertyFieldOrder", {});
    this.state = {
      value: this.props.value,
      validationErrors: [],
      showPanel: false,
    };
  }

  public componentDidUpdate(
    prevProps: IPropertyFieldMonacoEditorHostProps,
    prevState: IPropertyFieldMonacoEditorHostState
  ): void {
    if (prevProps.value !== this.props.value) {
      this.setState({ value: this.props.value });
    }
  }

  protected showPanel = (indicator: boolean): void => {
    this.setState({ showPanel: indicator });
  }

  private controlClasses = mergeStyleSets({
    headerTitle: mergeStyles({
      paddingTop: 20,
    }),
    textFieldStyles: mergeStyles({
      paddingBottom: 5,
    }),
  });

  protected _onValueChange = (newValue: string, errors: string[]): void => {
    this.setState({ value: newValue, validationErrors: errors });
  }

  //strings -> https://github.com/pnp/sp-dev-fx-property-controls/blob/master/src/loc/mystrings.d.ts
  //"Save" -> {strings.MonacoEditorSaveButtonLabel}
  //"Cancel" -> {strings.MonacoEditorCancelButtonLabel}
  protected onRenderFooterContent = (): JSX.Element => {
    return (
      <Stack horizontal horizontalAlign="start" tokens={{ childrenGap: 5 }}>
        <PrimaryButton
          onClick={(ev) => {
            ev.preventDefault();
            this.props.onPropertyChange(this.state.value);
            this.showPanel(false);
          }}
        >
          Save
        </PrimaryButton>
        <DefaultButton onClick={(ev) => {
           ev.preventDefault();
           this.props.onPropertyChange(this.props.value);
          this.showPanel(false);
          }}>Cancel</DefaultButton>
      </Stack>
    );
  }

  //"Open editor" -> {strings.MonacoEditorOpenButtonLabel}
  //"Edit Template" -> {strings.MonacoEditorPanelTitle}
  public render(): React.ReactElement<IPropertyFieldMonacoEditorHostProps> {
    const { panelWidth } = this.props;
    const _panelWidth = panelWidth ? `${panelWidth}px` : DEFAULT_PANEL_WIDTH;
    return (
      <>
        <TextField value={this.props.value} readOnly className={this.controlClasses.textFieldStyles} />
        <PrimaryButton
          text="Open editor"
          onClick={(ev) => {
            ev.preventDefault();
            this.showPanel(true);
          }}/>
        <Panel
          type={PanelType.custom}
          customWidth={_panelWidth}
          isOpen={this.state.showPanel}
          onDismiss={() => {
            this.showPanel(false);
          }}
          headerText="Edit Template"
          onRenderFooterContent={this.onRenderFooterContent}
          isFooterAtBottom={true}
          layerProps={{ eventBubblingEnabled: true }}
        >
          <div className={this.controlClasses.headerTitle}>
            <MonacoEditor {...this.props} onValueChange={ this._onValueChange}/>
          </div>
        </Panel>
      </>
    );
  }
}