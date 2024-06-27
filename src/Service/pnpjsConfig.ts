/* eslint-disable no-var */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { LogLevel, PnPLogging } from "@pnp/logging";
import { spfi, SPFI, SPFx, } from "@pnp/sp";
import { graphfi, GraphFI, SPFx as graphSPFx } from "@pnp/graph";
import { MSGraphClientV3 } from '@microsoft/sp-http';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/comments";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/files";

var _sp: SPFI;
var _graph: GraphFI;
var _graphClient: MSGraphClientV3;

export const getSP = (context?: WebPartContext): SPFI => {
    if (!!context) {
        _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Error));
    }
    return _sp;
}

export const getGraph = (context?: WebPartContext): GraphFI => {
    if (_graph === undefined && context !== undefined) {
        _graph = graphfi().using(graphSPFx(context)).using(PnPLogging(LogLevel.Error));
    }
    return _graph;
}

export const getGraphClient = async (context?: WebPartContext): Promise<MSGraphClientV3> => {
    if (context) {
        const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');
        _graphClient = client;
    }
    return _graphClient;
}