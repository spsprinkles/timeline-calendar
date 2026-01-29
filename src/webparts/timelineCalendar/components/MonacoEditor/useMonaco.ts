import { useState, useEffect } from "react";
import loader, { Monaco } from "@monaco-editor/loader";
import AcquireContext from "../AcquireContext";

export enum EStatus {
  LOADING,
  LOADED,
  ERROR,
}

export const useMonaco = (): {
  monaco: Monaco;
  status: EStatus;
  error: Error;
} => {
  const [monaco, setMonaco] = useState<Monaco>(undefined);
  const [status, setStatus] = useState<EStatus>(EStatus.LOADING);
  const [error, setError] = useState<Error>(undefined);

  let CDN_PATH_TO_MONACO_EDITOR = "https://cdnjs.cloudflare.com/ajax/libs/monaco-editor/0.47.0/min/vs";
  
  //Try to load the monaco editor
  useEffect(() => {
    (async () => {
      const loadMonaco = async () => {
        try {
          loader.config({ paths: { vs: CDN_PATH_TO_MONACO_EDITOR } });
          const monacoObj = await loader.init();
          setStatus(EStatus.LOADED);
          setMonaco(monacoObj);
        } catch (error) { //error instanceOf Event
          setStatus(EStatus.ERROR);
          setMonaco(undefined);
          if (error instanceof Error) {
            if (error.message !== null && error.message !== "")
              setError(error);
            else
              setError(new Error("There was an unknown error loading the Monaco Editor."));
          }
          else {
            //likely: (error instanceof Event)
            setError(new Error("There was an unknown error loading the Monaco Editor."));
          }
        }
      }

      //Tests for network path errors and CSP enforcement blocks
      const testLoadingScript = (src:string):Promise<string> => {
        const promise = new Promise<string>((resolve, reject) => {
          const s = document.createElement('script');
          s.src = src;
          s.async = true;
          //Successful load
          s.onload = () => resolve('');
          //Error due to Network/DNS/HTTP errors, blocked by CORS, integrity mismatch, etc.
          s.onerror = (ev) => reject('');
          //Append to <head>
          document.head.appendChild(s);
        });

        return promise;
      }
      
      const testUrlForMonaco = (url:string, cdnCheck:boolean):Promise<string> => {
        const promise = new Promise<string>((resolve, reject) => {
          fetch(url, {
            //method: 'HEAD',
            mode: 'cors'
            //redirect: 'follow'
          })
          .then(response => {
              return response.text();
          }).then(data => {
              if (data == null) {
                reject('data was null');
              }
              else
                if (data.indexOf('Version: 0.47.0(69991d66135e4a1fc1cf0b1ac4ad25d429866a0d)') == -1) {
                  reject('Version: 0.47.0 was not found');
                }
                else {
                  if (cdnCheck) {
                    resolve('');
                  }
                  else { //custom check (for user specified path)
                    if (data.indexOf('//# sourceMappingURL=../../min-maps/vs/loader.js.map') == -1)
                      reject('');
                    else
                      resolve('');
                  }
                }
          }).catch(error => {
              reject('promise catch');
          })
        });

        return promise;
      }

      const context = AcquireContext.getContext();
      /*context.pageContext.aadInfo {}
        .instanceUrl: "https://login.microsoftonline.com"
        .tenantId: {
            _guid: 'c5807244-0000-0000-adb9-357387d2a1de'
          }
          userId: {
            _guid: '7de78987-0000-0000-8d17-d19edeab9e41'
          }
      */
      //aadInstanceUrl = 'https://login.microsoftonline.com' / 'https://login.microsoftonline.us' / 'https://login.microsoftonline.microsoft.scloud'
      //cloudType = 'prod' / 'dod' / 'ag09'
      //msGraphEndpointUrl = 'https://graph.microsoft.com' / 'https://dod-graph.microsoft.us' / 'https://dod-graph.microsoft.scloud'
      //webDomain = 'sharepoint.com' / 'sharepoint-mil.us' (even *.dps.mil sites) / 'spo.microsoft.scloud'
      //substrateEndpointUrl = 'https://substrate.office.com' / 'https://substrate-dod.office365.us' / 'https://substrate.exo.microsoft.scloud'
      if (context.pageContext.legacyPageContext.aadInstanceUrl && 
            context.pageContext.legacyPageContext.aadInstanceUrl.endsWith("microsoft.scloud")) {
        //Override the path
        CDN_PATH_TO_MONACO_EDITOR = "https://dod365sec.spo.microsoft.scloud/sites/USAF-TipsToolsApps/code/monaco/min/vs";
        loadMonaco();
      }
      else {
        //Check if monaco is already loaded?
        //@ts-ignore
        //if (window.monaco)
        //Would need to store URL value outside/above and then read-in during initalization

        //Loading /loader.js twice causes a console error, even tho the edit appears to still load
        //loader.js:1 Uncaught SyntaxError: Identifier '_amdLoaderGlobal' has already been declared
        /*Alt options include:
          /basic-languages/html/html.js
          /language/html/htmlMode.min.js
          /basic-languages/handlebars/handlebars.min.js";
        */
        testLoadingScript(CDN_PATH_TO_MONACO_EDITOR + "/editor/editor.main.js")
          .then(() => loadMonaco())
          //Error
          .catch(() => {
            const params1 = new URLSearchParams(window.location.search);
            if (params1.has("loadMonaco")) {
                loadMonaco();
                return;
            }

            //Try another CDN
            CDN_PATH_TO_MONACO_EDITOR = "https://cdn.jsdelivr.net/npm/monaco-editor@0.47.0/min/vs";
            testLoadingScript(CDN_PATH_TO_MONACO_EDITOR + "/editor/editor.main.js")
              .then(() => loadMonaco())
              //Error
              .catch(() => {
                //Check if query string was found
                const params = new URLSearchParams(window.location.search);
                if (params.has("monaco")) {
                  let folderPath = decodeURIComponent(params.get("monaco"));
                  //Removal the ending file part if found
                  folderPath = folderPath.split("/loader.js")[0];
                  //Remove trailing slash if user provided one
                  if (folderPath.lastIndexOf("/") + 1 === folderPath.length)
                    folderPath = folderPath.substring(0, folderPath.length-1);

                  //Test to see if this is the expected file
                  CDN_PATH_TO_MONACO_EDITOR = folderPath;
                  testUrlForMonaco(CDN_PATH_TO_MONACO_EDITOR + "/loader.js", false)
                    .then(() => loadMonaco())
                    .catch(() => {
                      setStatus(EStatus.ERROR);
                      setMonaco(undefined);
                      setError(new Error("CDN paths to load Monaco Editor are blocked and alt provided path was invalid."));
                    });
                }
                else {
                  //No local monaco path provided
                  setStatus(EStatus.ERROR);
                  setMonaco(undefined);
                  setError(new Error("CDN paths to load Monaco Editor are blocked and no local path was provided."));
                }
              });
          });

        return;

        //Since domains can be blocked at network level, check for access first (this doesn't test for CSP blocks)
        testUrlForMonaco(CDN_PATH_TO_MONACO_EDITOR + "/loader.js", true)
        .then(() => {
          loadMonaco();
        })
        //Error
        .catch(() => {
          //Try another one
          testUrlForMonaco("https://cdn.jsdelivr.net/npm/monaco-editor@0.47.0/min/vs/loader.js", true)
          .then(() => {
            CDN_PATH_TO_MONACO_EDITOR = "https://cdn.jsdelivr.net/npm/monaco-editor@0.47.0/min/vs";
            loadMonaco();
          })
          //Error
          .catch(() => {
            //Check if query string was found
            const params = new URLSearchParams(window.location.search);
            if (params.has("monaco")) {
              let folderPath = decodeURIComponent(params.get("monaco"));
              //Removal the ending file part if found
              folderPath = folderPath.split("/loader.js")[0];
              //Remove trailing slash if user provided one
              if (folderPath.lastIndexOf("/") + 1 === folderPath.length)
                folderPath = folderPath.substring(0, folderPath.length-1);

              //Test to see if this is the expected file
              testUrlForMonaco(folderPath + "/loader.js", false)
              .then(() => {
                CDN_PATH_TO_MONACO_EDITOR = folderPath;
                loadMonaco();
              })
              .catch(() => {
                //Attempt the load anyway so that at least a failure msg is shown if applicable
                loadMonaco();
                //TODO: Change above
              });
            }
            else {
              //No local monaco path provided
              setStatus(EStatus.ERROR);
              setMonaco(undefined);
              setError(new Error("CDN paths to load Monaco editor are blocked and no local path was provided."));
            }
          });
        });
      }
    })().then(() => { /* no-op; */ }).catch(() => { /* no-op; */ });
  }, []);

  return {
    monaco,
    status,
    error,
  };
};