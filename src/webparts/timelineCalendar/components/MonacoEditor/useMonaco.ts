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

  let CDN_PATH_TO_MONACO_EDITOR = "https://cdnjs.cloudflare.com/ajax/libs/monaco-editor/0.47.0/min/vs"; //orig was 0.32.1 "https://cdn.jsdelivr.net/npm/monaco-editor@0.47.0/min/vs"

  //Try to load the monaco editor
  useEffect(() => {
    (async () => {
      const loadMonaco = async () => {
        try {
          loader.config({ paths: { vs: CDN_PATH_TO_MONACO_EDITOR } });
          const monacoObj = await loader.init();
          setStatus(EStatus.LOADED);
          setMonaco(monacoObj);
        } catch (error) {
          setStatus(EStatus.ERROR);
          setMonaco(undefined);
          setError(error);
        }
      }
      
      const testUrlNetworkAccess = (url:string, defaultCheck:boolean):Promise<string> => {
        const promise = new Promise<string>((resolve, reject) => {
          fetch(url, {
            //method: 'HEAD',
            mode: 'cors'
            //redirect: 'follow'
          })
          .then(response => {
              return response.text();
          }).then(data => {
              if (data == null)
                reject('');
              else
                if (defaultCheck) {
                  if (data.indexOf('Version: 0.47.0(69991d66135e4a1fc1cf0b1ac4ad25d429866a0d)') == -1)
                    reject('');
                  else
                    resolve('');
                }
                else { //custom check (for user specified path)
                  if (data.indexOf('(AMDLoader || (AMDLoader = {})') == -1 && data.indexOf('//# sourceMappingURL=../../min-maps/vs/loader.js.map') == -1)
                    reject('');
                  else
                    resolve('');
                }
          }).catch(error => {
              reject('');
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
        //Check if monaco is already loaded
        //@ts-ignore
        //if (window.monaco)
        //Would need to store URL value outside/above and then read-in during initalization

        //Since domains can be blocked, check for access first
        testUrlNetworkAccess(CDN_PATH_TO_MONACO_EDITOR + "/loader.js", true)
        .then(() => {
          loadMonaco();
        })
        //Error
        .catch(() => {
          //Try another one
          testUrlNetworkAccess("https://cdn.jsdelivr.net/npm/monaco-editor@0.47.0/min/vs/loader.js", true)
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
              testUrlNetworkAccess(folderPath + "/loader.js", false)
              .then(() => {
                CDN_PATH_TO_MONACO_EDITOR = folderPath;
                loadMonaco();
              })
              .catch(() => {
                //Attempt the load anyway so that at least a failure msg is shown if applicable
                loadMonaco();
              });
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