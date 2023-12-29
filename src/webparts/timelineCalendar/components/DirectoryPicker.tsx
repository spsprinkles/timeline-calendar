import * as React from 'react';
//import { IPersonaSharedProps } from '@fluentui/react/lib/Persona'; //IPersonaProps
import { NormalPeoplePicker, IBasePickerSuggestionsProps, IBasePickerStyles, //ValidationState
    IPeoplePickerItemSelectedProps, PeoplePickerItem } from '@fluentui/react/lib/Pickers';
import { IPersonaProps } from './IConfigurationItems';
import { MSGraphClientV3,} from '@microsoft/sp-http';
import { GraphError } from '@microsoft/microsoft-graph-client'
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { DefaultButton } from '@fluentui/react/lib/Button'; //PrimaryButton

interface IDirectoryPickerProps {
  onChange: (items:IPersonaProps[]) => void;
  selectedPersonas: IPersonaProps[];
  initialSuggestions: IPersonaProps[];
  graphClient: Promise<MSGraphClientV3>;
  disabled: boolean;
  //uniqueId: string;
  //sortIdx?: number;
}

export interface IDirectoryPickerState {
    //loading: boolean;
    //options: IDropdownOption[];
    //recentlyUsed: any[];
    selected: IPersonaProps[];
    directoryListing: any[];
    error: string;
    graphPermissionsErrorHeader: string;
    graphPermissionsError: string;
    graphPermissionsSubError: string;
    hideMissingPermissionsDialog: boolean;
}

export default class AsyncDropdown extends React.Component<IDirectoryPickerProps, IDirectoryPickerState> {
    constructor(props: IDirectoryPickerProps, state: IDirectoryPickerState) {
      super(props);      
      this.state = {
        //options: undefined,
        //recentlyUsed: undefined,
        directoryListing: undefined,
        error: undefined,
        selected: props.selectedPersonas,
        graphPermissionsErrorHeader: undefined,
        graphPermissionsError: undefined,
        graphPermissionsSubError: undefined,
        hideMissingPermissionsDialog: true
      };
    }

    //When component first loads (not fired again for row re-used as the "new row")
    public componentDidMount(): void {
      //Nothing needed
    }
  
    public componentDidUpdate(prevProps: IDirectoryPickerProps, prevState: IDirectoryPickerState): void {
      //Address case where field is re-used for the "new row"
      //Also, if unique ID is passed in, you'll see this.props.uniqueId != prevProps.uniqueId because the field is "updated" for the "new row"
      if (this.props.selectedPersonas == null && this.state.selected != null && this.state.selected.length != 0)
        this.setState({
          selected: []
        })
    }

    public render(): JSX.Element {
        const suggestionProps: IBasePickerSuggestionsProps = {
            suggestionsHeaderText: 'Suggested Results',
            mostRecentlyUsedHeaderText: 'Suggested Items',
            noResultsFoundText: 'No results found',
            loadingText: 'Loading...',
            showRemoveButtons: true,
            suggestionsAvailableAlertText: 'PickersSuggestions available',
            suggestionsContainerAriaLabel: 'Suggested items',
        };

        const toggleHideMissingPermissionsDialog = (): void => {
          this.setState({
            hideMissingPermissionsDialog: !this.state.hideMissingPermissionsDialog
          });
        }

        const onInputChanged = (filterText: string, currentPersonas: IPersonaProps[], limitResults?: number): IPersonaProps[] | Promise<IPersonaProps[]> => {
          if (filterText) {
            //*Filter* for users instead of search
            //const filterQuery = `(startswith(displayName,'${filterText}') or startswith(mailNickname,'${filterText}') or startswith(mail,'${filterText}'))`;
            //https://graph.microsoft.com/v1.0/users?$select=id,givenName,jobTitle,mail,mailNickname,surname,displayName,userPrincipalName&$filter=(userType eq 'Member') AND (startswith(displayName,'jo') OR startswith(mail,'doe'))
            //.orderby("displayName")

            const encodedQuery = `${filterText.replace(/#/g, '%2523')}`; //for groups only?

            //TODO: Check if user is a guest user and change queries (Graph to "special" SP API)
            //TODO: Can guest users even query calendars? If not, show dialog that they cannot use this

            //Build users promise and query
            const usersPromise = new Promise<string | any[]>((resolve, reject) => {
              //*Search* for users
              this.props.graphClient.then((client:MSGraphClientV3): void => {
                client.api("/users").select("id,givenName,surname,displayName,mail,userPrincipalName") //userType & jobTitle are null with just basic perms
                .filter("userType eq 'Member'") //cannot filter on userType with just User.ReadBasic.All perms
                .search(`"displayName:${encodedQuery}" OR "mail:${encodedQuery}"`)
                .header('ConsistencyLevel', 'eventual')
                //.orderby ???
                //.count(true)
                .get((error:GraphError, response:any, rawResponse?:any) => {
                  //Look for and save Graph scope information
                  /*
                  for (const key in sessionStorage) {
                    //@ts-ignore (for startsWith)
                    if (key && typeof key == "string" && key.startsWith('{"authority":')) {
                      try {
                        const keyObj = JSON.parse(key);
                        //Find the correct Graph results entry
                        if (keyObj.scopes && keyObj.scopes.indexOf('profile openid') != -1) {
                          const scopes = keyObj.scopes.split(" ");
                          this._graphScopes = scopes.map((s:string, i:Number) => {
                            if (s.indexOf('graph.microsoft') != -1)
                              return s.split("/")[3]; //https://graph.microsoft.com/User.Read (just get "User.Read" portion)
                            else
                              return s;
                          });
                        }
                      }
                      catch (e) {}
                    }
                  }*/
                  
                  if (error)
                    resolve(error.message);
                  else {
                    const users:MicrosoftGraph.User[] = response.value;
                    const personas:IPersonaProps[] = users.map((user:MicrosoftGraph.User, index:number) => {
                        return {
                            key: user.id,
                            mail: user.mail,
                            //imageUrl: './images/persona-male.png',
                            //imageInitials: 'AL', //leaving out or set to null still renders initials
                            text: user.displayName,
                            secondaryText: user.mail,
                            //tertiaryText: 'In a meeting',
                            //optionalText: 'Available at 4:00pm',
                            //isValid: true, ??????
                            //presence: 0 //"none"
                            personaType: "user"
                        }
                    });
                    resolve(personas);
                  }
                });
              });
            });

            //Build groups promise and query
            const groupsPromise = new Promise<string | any[]>((resolve, reject) => {
              //*Filter* for groups
              //https://graph.microsoft.com/v1.0/groups?$select=id,displayName,mail&$filter=groupTypes/any(c:c eq 'Unified') AND (startswith(displayName,'some') OR startswith(mail,'some'))
              //orderBy not supported!
              this.props.graphClient.then((client:MSGraphClientV3): void => {
                client.api("/groups").select("id,displayName,mail,visibility")
                .filter(`groupTypes/any(c:c eq 'Unified')`)//had this: AND (startswith(displayName,'${encodedQuery}') OR startswith(mailNickname,'${encodedQuery}') OR startswith(mail,'${encodedQuery}'))`)
                //Switch to search instead to match portion of group name
                .search(`"displayName:${encodedQuery}" OR "mail:${encodedQuery}"`)
                .header('ConsistencyLevel', 'eventual')
                .get((error:GraphError, response:any, rawResponse?:any) => {
                  if (error)
                    resolve(error.message);
                  else {
                    const groups:MicrosoftGraph.Group[] = response.value;
                    const personas:IPersonaProps[] = groups.map((group:MicrosoftGraph.Group, index:number) => {
                        return {
                            key: group.id,
                            mail: group.mail,
                            //imageInitials: "G", //can force specific initials
                            text: group.displayName,
                            secondaryText: group.visibility + " group",
                            //canExpand: true,
                            //isValid: true
                            personaType: "group"
                        }
                    });
                    resolve(personas);
                  }
                });
              });
            });

            //When both are finished, combine the results
            //const finalPromise = 
            return Promise.all([usersPromise, groupsPromise]).then(values => {
              //When errors are returned, they look like this (depending on if users or groups query failed)
              //values[0] == "Insufficient privileges to complete the operation."

              //Check if *array* objects were returned instead of *string* error messages
              let personas:IPersonaProps[] = [];
              if (Array.isArray(values[0])) //typeof values[0] == "object"
                personas = personas.concat(values[0] as IPersonaProps[]);
              if (Array.isArray(values[1])) //typeof values[1] == "object"
                personas = personas.concat(values[1] as IPersonaProps[]);

              //Check for errors to show dialog message
              let hideDialog = true;
              let errorHeader:string = undefined;
              let errorMsg:string = undefined;
              let errorSubMsg:string = undefined;

              //Look for CAS policy error (should be in both queries but just check the first)
              //@ts-ignore (for startsWith)
              if (typeof values[0] == "string" && values[0].startsWith("AADSTS53003:")) { //"AADSTS53003: Access has been blocked by Conditional Access policies. The access policy does not allow token issuance. Trace ID: 53f94e25-27a1-4f11-8318-b4c794570800 Correlation ID: 003f1386-5365-4a10-9f5b-f762e1788619 Timestamp: 2023-12-28 13:51:19Z"
                //error.code == "InteractionRequiredAuthError" // error.statusCode == -1
                hideDialog = false;
                errorHeader = "Graph API token cannot be generated";
                errorMsg = "Your current sign-in may be from a \"location\" that is restricted. Please connect to your organization's network or VPN and re-sign in before trying again.";
                errorSubMsg = values[0];
              }
              else {
                //Look for permission error messages (due to lack of approved Graph scopes)
                if (values[0] == "Insufficient privileges to complete the operation.")
                  errorMsg = "users" + 
                  //Add trailing . character if there's no permission issues with groups
                  (values[1] == "Insufficient privileges to complete the operation." ? "" : ".");
                //Check if groups call had error
                if (values[1] == "Insufficient privileges to complete the operation.") {
                  //Check if users message is already present
                  if (errorMsg)
                    errorMsg += " and groups."
                  else
                    errorMsg = "groups."
                }

                if (errorMsg) {
                  errorMsg = "The SharePoint Online admins for your tenant have not approved the permissions needed to search for " 
                    + errorMsg;
                  errorHeader = "Graph API permissions not approved";
                  errorSubMsg = "Please contact them (submit a ticket) and point them to the documentation links provided in the last page of the editing panel within this web part.";
  
                  //Check if message should be displayed
                  const sessionVar = sessionStorage["TCWP-GraphPermsDirSearch"] as string;
                  const now = new Date();
                  if (sessionVar) {
                    //Check if it's been over 12 hours
                    if (((new Date(sessionVar)).getTime() - now.getTime()) / (1000*60*60) >= 12) {
                      hideDialog = false;
                      sessionStorage["TCWP-GraphPermsDirSearch"] = now.toISOString();
                    }
                  }
                  else {
                    //No prior check in storage
                    hideDialog = false;
                    sessionStorage["TCWP-GraphPermsDirSearch"] = now.toISOString();
                  }
                }
              }

              //Set received personas from search and/or show error dialog if applicable
              this.setState({
                directoryListing: personas,
                graphPermissionsErrorHeader: errorHeader,
                graphPermissionsError: errorMsg,
                graphPermissionsSubError: errorSubMsg,
                hideMissingPermissionsDialog: hideDialog
              });
              return this.state.directoryListing;
            });
          } else //no valid filterText
            return [];
        };
        
        //Show the suggestions
        const returnMostRecentlyUsed = (currentPersonas: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> => {
          return this.props.initialSuggestions;
          //return filterPromise(this.removeDuplicates(this.props.initialSuggestions, currentPersonas));
          //this.state.recentlyUsed
        };
      
        const onChange = (items: any[]): void => {  
          this.setState({
            selected: items
          });
  
          if (this.props.onChange)
            this.props.onChange(items);
        }

        const onBeforeRenderItem = (props: IPeoplePickerItemSelectedProps) => { //IPickerItemProps<IPersonaProps>
          const newProps = {
            ...props,
            item: {
              ...props.item,
              //@ts-ignore (for mail prop)
              secondaryText: props.item.mail,
              //title: "Some specified title", to force a browser tooltip
              //ValidationState: ValidationState.valid, //doesn't seem to be needed
              showSecondaryText: true
            },
          };
      
          return <PeoplePickerItem {...newProps} />;
        };

        const pickerStyles: Partial<IBasePickerStyles> = { root : { "backgroundColor": "white" }};
        return (
          <>
            <NormalPeoplePicker
                //key?
                onResolveSuggestions={onInputChanged}
                selectedItems={this.state.selected}
                itemLimit={1}
                onEmptyResolveSuggestions={returnMostRecentlyUsed}
                //getTextFromItem={this.getTextFromItem}
                pickerSuggestionsProps={suggestionProps}
                className={'ms-PeoplePicker'}
                removeButtonAriaLabel={'Remove'}
                //onValidateInput={this.validateInput} //not applicable
                //onItemSelected={} ????
                onRenderItem={onBeforeRenderItem} //adjust how selected items appear (show their email)
                onChange={onChange}
                inputProps={{
                    //onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
                    //onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
                    'aria-label': 'People Picker',
                }}
                //componentRef={picker}
                resolveDelay={1000}
                disabled={false}
                styles={pickerStyles}
            />
            <Dialog
              hidden={this.state.hideMissingPermissionsDialog}
              onDismiss={toggleHideMissingPermissionsDialog}
              minWidth={415}
              maxWidth={415}
              dialogContentProps={{
                type: DialogType.largeHeader,
                title: this.state.graphPermissionsErrorHeader,
                //subText: this.state.graphPermissionsError //shows in slightly different color
              }}
              modalProps={{
                isBlocking: true,
                //isModeless: true,
              }}
            >
              <div>{this.state.graphPermissionsError}</div>
              <div style={{marginTop: 20}}>{this.state.graphPermissionsSubError}</div>
              <DialogFooter>
                {/* <PrimaryButton onClick={toggleHideMissingPermissionsDialog} text="Save" /> */}
                <DefaultButton onClick={toggleHideMissingPermissionsDialog} text="OK" />
              </DialogFooter>
            </Dialog>
          </>
        )
    }
    
    /*private getTextFromItem(persona: IPersonaProps): string {
        console.log("inside getTextFromItem");
        return persona.text as string;
    }*/
    
    /*private validateInput(input: string): ValidationState {
      //console.log("inside validateInput");
        if (input.indexOf('@') !== -1) {
            return ValidationState.valid;
        } else if (input.length > 1) {
            return ValidationState.warning;
        } else {
            return ValidationState.invalid;
        }
    }*/
}