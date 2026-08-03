// shared/pnpjsConfig.ts
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { graphfi, GraphFI, SPFx as GraphSPFx } from "@pnp/graph";
import "@pnp/graph/users";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export const getSP = (context: any, baseUrl?: string): SPFI => {
  return baseUrl
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    ? spfi(baseUrl).using(SPFx(context))
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    : spfi().using(SPFx(context));
};

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export const getGraph = (context: any): GraphFI => {
  return graphfi().using(GraphSPFx(context));
};
