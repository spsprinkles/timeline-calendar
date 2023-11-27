//Resolve issue with structuredClone() not being in Node version
//https://stackoverflow.com/questions/70661021/structuredclone-not-available-in-typescript
interface WindowOrWorkerGlobalScope {
    structuredClone(value: any, options?: StructuredSerializeOptions): any;
}
declare function structuredClone( value: any, options?: StructuredSerializeOptions): any;