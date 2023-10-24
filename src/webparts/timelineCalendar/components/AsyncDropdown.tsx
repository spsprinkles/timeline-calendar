import * as React from 'react';
//import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
//import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
//remove /components
import { Dropdown, IDropdownOption, IDropdownStyles } from 'office-ui-fabric-react/lib/components/Dropdown';
//import { Spinner } from 'office-ui-fabric-react';
//import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
//import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';

//Reference: https://github.com/pnp/sp-dev-fx-aces/blob/main/samples/PrimaryTextCard-PlannerTasks/src/controls/PropertyPaneAsyncDropdown/components/AsyncDropdown.tsx

interface IAsyncDropdownProps {
  label: string;
  loadOptions: () => Promise<IDropdownOption[]>;
  onChange: (event: Event, option: IDropdownOption) => void;
  selectedKey: string | number;
  disabled: boolean;
  stateKey: string; //Can contain invalid data (bad Site URL) that doesn't match that field's validation
  //isNewRow: boolean; //TODO: //From PropPaneConfig onCustomRender(), item.sortIdx evaluating true means it's an *existing* row (not a *new* row)
}

export interface IAsyncDropdownState {
    loading: boolean;
    options: IDropdownOption[];
    error: string;
}

export default class AsyncDropdown extends React.Component<IAsyncDropdownProps, IAsyncDropdownState> {
    private selectedKey: React.ReactText;
    
    constructor(props: IAsyncDropdownProps, state: IAsyncDropdownState) {
      super(props);
      this.selectedKey = props.selectedKey;
            
      this.state = {
        loading: false,
        options: undefined,
        error: undefined
      };
    }

    //When component first loads (not fired again for row re-used as the "new row")
    public componentDidMount(): void {
      //if (this.props.item.Site == null) //this is always null for initial/new creation
      if (this.props.stateKey == null)
         return;
      else {
        //console.log("AsyncDropdown componentDidMount stateKey: " + this.props.stateKey + " // selectedKey: " + this.props.selectedKey);
        this.loadOptions(true);
      }
    }
  
    public componentDidUpdate(prevProps: IAsyncDropdownProps, prevState: IAsyncDropdownState): void {
      //Address case where field is re-used for the "new row"
      //TODO: For calendar list selecting start & end, this.selectedKey is null and not showing
      if (this.props.selectedKey == null && this.selectedKey != null && this.selectedKey != "temp")
        this.selectedKey = null;

      if (this.props.stateKey == null)
        return;

      if (this.props.disabled !== prevProps.disabled || this.props.stateKey !== prevProps.stateKey) {
        //console.log("Async Update current stateKey:" + this.props.stateKey + " // prevProps.stateKey: " + prevProps.stateKey);
        this.loadOptions();
      }
    }

    private loadOptions(noDelay?:boolean): void {
      //Was selecting item even when URL was changed and should have loaded new lists
      /*if (this.selectedKey != undefined && this.selectedKey != "temp") {
        this.setState({
          options: [{key: this.selectedKey, text: "the text here"}]
        });
        return;
      }*/

      const currentStateKey = this.props.stateKey; //save for checking later to prevent queries for each character input
      const timeoutTime = (noDelay ? 0 : 1500);//(this.props.isNewRow ? 10 : 1500); //TODO

      //Prevent running again if already set to loading
      // if (this.state.loading == false) {
      //   this.setState({
      //     loading: true,
      //     error: undefined,
      //     options: [{key:"temp", text:"Loading..."}]
      //   });
      //   this.selectedKey = "temp";
      //   //Had above moved under the below if statement
      //   //TODO: rather check if this is the very first load for new row
      // }

      setTimeout(() => {
        if (currentStateKey != this.props.stateKey)
          return; //don't call loadOptions for "temporary" keys such as in case of Site being typed in and each letter change was a call
        
        this.props.loadOptions()
          .then((options: IDropdownOption[]): void => {
            //TODO: Need to add selected if not within options?

            this.setState({
                loading: false,
                error: undefined,
                options: options
            });
            //this.selectedKey = null; //needed?
            this.selectedKey = this.props.selectedKey;

          }, (error: any): void => {
              // this.setState((prevState: IAsyncDropdownState, props: IAsyncDropdownProps): IAsyncDropdownState => {
              //     prevState.loading = false;
              //     prevState.error = error;
              //     return prevState;
              // });
              this.setState({
                loading: false,
                error: error,
                options: []
              });
              this.selectedKey = null; //needed?

              if (this.props.onChange)
                this.props.onChange(null, null);
          });
      }, timeoutTime);
    }
  
    public render(): JSX.Element {
      const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { minWidth: 125 } }; //, maxWidth:200 doesn't seem to do anything
        return (
            <Dropdown label={this.props.label}
                disabled={this.props.disabled || this.props.stateKey == null || this.state.loading || this.state.error !== undefined}
                onChange={this.onChange.bind(this)}
                selectedKey={this.selectedKey}
                styles={dropdownStyles}
                //placeholder='Loading items...'
                options={this.state.options} />
        )
    }

    private onChange(event:Event, option: IDropdownOption): void {
        this.selectedKey = option.key;
        //Reset previously selected options (not needed?)
        const options: IDropdownOption[] = this.state.options;
        // options.forEach((o: IDropdownOption): void => {
        //   if (o.key !== option.key) {
        //     o.selected = false;
        //   }
        // });
        this.setState((prevState: IAsyncDropdownState, props: IAsyncDropdownProps): IAsyncDropdownState => {
          prevState.options = options;
          return prevState;
        });
        if (this.props.onChange)
          this.props.onChange(event, option);
      }
  }