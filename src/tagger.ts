import * as azdev from "azure-devops-node-api";
import {ReleaseApi} from "azure-devops-node-api/ReleaseApi";
import {WorkItemTrackingApi} from "azure-devops-node-api/WorkItemTrackingApi";
import {IRequestHandler} from "azure-devops-node-api/interfaces/common/VsoBaseInterfaces";
import {WebApi} from "azure-devops-node-api";

export default class Tagger {
    releaseClient: ReleaseApi;
    workItemClient: WorkItemTrackingApi;
    
    public static async build(accessToken: string, organizationUri: string): Promise<Tagger> {
        let tagger = await new Tagger().init(accessToken, organizationUri);
        return tagger;
    }
    
    public async init(accessToken: string, organizationUri: string): Promise<Tagger> {
        let authHandler: IRequestHandler = azdev.getPersonalAccessTokenHandler(accessToken);
        let connection: WebApi = new azdev.WebApi(organizationUri, authHandler);
        this.releaseClient = await connection.getReleaseApi();
        this.workItemClient = await connection.getWorkItemTrackingApi();
        return this;
    }
    
    public async tagItemsInRelease(projectName: string, releaseId: number): Promise<void>{
        console.log(`Project name: ${projectName}, release id: ${releaseId}`);
        const workItemIds = await this.getWorkItemIds(projectName, releaseId);
        const year = await this.getReleaseYear(projectName, releaseId);
        await this.addReleaseTagToWorkItem(projectName, workItemIds, year);
    }
    
    private async getWorkItemIds(projectName: string, releaseId: number): Promise<number[]> {
        let workItems = await this.releaseClient.getReleaseWorkItemsRefs(projectName, releaseId);
        if (!workItems) {
            throw "No work items found";
        }
        
        let workItemIds = workItems.map(wi => {
            if (!wi.id) throw "No work item reference id";
            return parseInt(wi.id);
        });
        
        return workItemIds;
    }

    private async getReleaseYear(projectName: string, releaseId: number): Promise<number> {
        let release = await this.releaseClient.getRelease(projectName, releaseId);
        if (!release) {
            throw "No release found";
        }
        
        if (!release.createdOn) {
            throw "No release date found";
        }
        
        return release.createdOn.getFullYear();
    }

    private async addReleaseTagToWorkItem(projectName: string, workItemIds: number[], year: number): Promise<void> {
        const tagName = `ProdRelease_${year}`;
        const patchDocument = [{
            "op": "add",
            "path": "/fields/System.Tags",
            "value": tagName
        }];
        
        for (const id of workItemIds) {
            await this.workItemClient.updateWorkItem(null, patchDocument, id, projectName);
        }
    }
}