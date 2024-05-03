import { WebPartContext } from "@microsoft/sp-webpart-base";

export default class AcquireContext {
    private static context: WebPartContext = null;
    
    //"private" before initContext
    constructor(initContext?:WebPartContext) {
        if (initContext !== null)
            AcquireContext.context = initContext;
    }

    public static setContext(initContext:WebPartContext):void {
        AcquireContext.context = initContext;
    }

    public static getContext(): WebPartContext {
        return AcquireContext.context;
    }
}