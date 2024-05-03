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

  let CDN_PATH_TO_MONACO_EDITOR = "https://cdn.jsdelivr.net/npm/monaco-editor@0.47.0/min/vs"; //orig was 0.32.1
  // CDN_PATH_TO_MONACO_EDITOR = "https://cdnjs.cloudflare.com/ajax/libs/monaco-editor/0.47.0/min/vs";

  //aadInstanceUrl = 'https://login.microsoftonline.com' / 'https://login.microsoftonline.us' / 'https://login.microsoftonline.microsoft.scloud'
  //cloudType = 'prod' / 'dod' / 'ag09'
  //msGraphEndpointUrl = 'https://graph.microsoft.com' / 'https://dod-graph.microsoft.us' / 'https://dod-graph.microsoft.scloud'
  //webDomain = 'sharepoint.com' / 'sharepoint-mil.us' (even *.dps.mil sites) / 'spo.microsoft.scloud'
  //substrateEndpointUrl = 'https://substrate.office.com' / 'https://substrate-dod.office365.us' / 'https://substrate.exo.microsoft.scloud'

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
  if (context.pageContext.legacyPageContext.aadInstanceUrl && 
        context.pageContext.legacyPageContext.aadInstanceUrl.endsWith("microsoft.scloud"))
    //Override the path
    CDN_PATH_TO_MONACO_EDITOR = "https://dod365sec.spo.microsoft.scloud/sites/USAF-TipsToolsApps/code/monaco/min/vs";

  //Try to load the monaco editor
  useEffect(() => {
    (async () => {
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
    })().then(() => { /* no-op; */ }).catch(() => { /* no-op; */ });
  }, []);

  return {
    monaco,
    status,
    error,
  };
};