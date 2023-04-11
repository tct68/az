import * as azdev from "azure-devops-node-api";
import { config } from "dotenv";
import { writeFileSync, existsSync, mkdirSync, readFileSync } from "fs";
import _ from "lodash";
import xlsx from "xlsx";
config();

class Az {
    private authHandler = azdev.getPersonalAccessTokenHandler(`${process.env.AZURE_PERSONAL_ACCESS_TOKEN}`);
    private connection = new azdev.WebApi(`${process.env.ORG_URL}`, this.authHandler);
    private readonly DATA_PATH = "data";
    private readonly INPUT_PATH = "input";
    private readonly OUTPUT_PATH = "output";
    private readonly TIME_LOG_PATH = `${this.INPUT_PATH}/timelog.csv`;
    private readonly OUTPUT_FILE_PATH = `${this.OUTPUT_PATH}/output.xlsx`;

    constructor() {
        if (!existsSync(this.DATA_PATH)) {
            mkdirSync(this.DATA_PATH);
        }
    }

    convertExeclDate = (excelSerialDate: any) => {
        const unixTimestamp = (excelSerialDate - 25569) * 86400;
        const dateObj = new Date(unixTimestamp * 1000);

        const month = dateObj.getMonth() + 1;
        const day = dateObj.getDate();
        const year = dateObj.getFullYear();

        const formattedDate = `${month}/${day}/${year}`;
        return formattedDate;
    }

    getWorkItemsFromTimelog = async (): Promise<TimeLogWorkItem[]> => {
        const timelog = xlsx.readFile(this.TIME_LOG_PATH)
        const data = xlsx.utils.sheet_to_json(timelog.Sheets['Sheet1'])
        const wIds: TimeLogWorkItem[] = []
        console.log(`Có ${data.length} workItem trong timelog`)
        for (let i = 0; i < data.length; i++) {
            const wiInfo: any = data[i];
            console.log(`Get thông tin workitem ${wiInfo.title}`)
            const wi = await this.getWorkItemInfo(wiInfo.workItemId)
            if (wi) {
                wIds.push({
                    workitem: {
                        id: wi.id ?? 0,
                        wiTitle: wi.fields!['System.Title'],
                        wiUrl: `https://symphonyvsts.visualstudio.com/VSA/_workitems/edit/${wi.id}`,
                        podLead: wi.fields!['System.CreatedBy'].displayName,
                        fields: wi.fields,
                    },
                    date: this.convertExeclDate(wiInfo.date),
                    type: wiInfo.type
                })
            }
        }
        return wIds;
    }

    writeData = (path: String, data: any) => {
        writeFileSync(`data/${path}`, JSON.stringify(data, null, 2));
    }

    getQuarterFromDate(dateStr: string) {
        const date = new Date(dateStr);
        const month = date.getMonth();
        const quarter = Math.floor(month / 3) + 1;
        return `Q${quarter}`;
    }

    getWorkItemInfo = async (wiId: Number): Promise<WorkItem | null> => {
        const workItemTracking = await this.connection.getWorkItemTrackingApi();
        const wi = await workItemTracking.getWorkItem(Number(wiId));
        if (wi) {
            return {
                id: wi.id ?? 0,
                wiTitle: wi.fields!['System.Title'],
                wiUrl: `https://symphonyvsts.visualstudio.com/VSA/_workitems/edit/${wi.id}`,
                podLead: wi.fields!['System.CreatedBy'].displayName,
                fields: wi.fields
            }
        }

        return null;
    }

    getWorkItemsInfo = async (workItemIds: (string | undefined)[]): Promise<WorkItem[]> => {
        const workItems: WorkItem[] = [];
        for (let i = 0; i < workItemIds.length; i++) {
            const wId = workItemIds[i];
            const wi = await this.getWorkItemInfo(Number(wId));
            if (wi && wi.fields!['System.AssignedTo'].id === process.env.USER_ID) {
                workItems.push(wi)
            }
        }

        return workItems;
    }

    getWorkItemPullRequest = (workItemId: Number, pullRequests: PullRequest[]) => {
        return pullRequests.find(x => x.workItems.some((w: WorkItem) => w.id == workItemId))
    }

    getWorkItemTracking = async () => {
        const gitApi = await this.connection.getGitApi();
        const workItemTimeLog = await this.getWorkItemsFromTimelog();
        const pullRequests = await gitApi.getPullRequests("VSA.Application", {
            creatorId: process.env.USER_ID,
            status: 3
        }, "VSA");

        console.log(`Có ${pullRequests.length} pull request`)

        const pullRequestWorkItems: PullRequest[] = [];
        for (let index = 0; index < pullRequests.length; index++) {
            const pr = pullRequests[index];
            const pullRequestWorkItemRefs = await gitApi.getPullRequestWorkItemRefs("VSA.Application", pr.pullRequestId ?? 0, "VSA");
            const pullRequestWorkItemIds = pullRequestWorkItemRefs.map(x => x.id);
            const workItems = await this.getWorkItemsInfo(pullRequestWorkItemIds);
            console.log(`Get thông tin pulll request ${pr.title}`);

            pullRequestWorkItems.push({
                title: pr.title ?? "",
                pullRequestId: pr.pullRequestId ?? 0,
                pullRequestUrl: `https://symphonyvsts.visualstudio.com/VSA/_git/VSA.Application/pullrequest/${pr.pullRequestId}`,
                workItems
            });
        }

        const taskSummaries: TaskSummary[] = [];
        for (const w of workItemTimeLog) {
            const pr = this.getWorkItemPullRequest(w.workitem.id, pullRequestWorkItems);

            taskSummaries.push({
                date: w.date,
                channelName: "",
                podlead: w.workitem?.podLead ?? "",
                quarter: this.getQuarterFromDate(w.date),
                ticket: w.workitem.wiTitle,
                workItemType: w.type,
                pr: !!pr ? pr.pullRequestUrl : "N/A"
            });
        }

        this.writeData('finalData.json', JSON.stringify(this.sort(JSON.stringify(taskSummaries)), null, 2));
    }

    sort = (taskSummaries: string) => {
        let json = JSON.parse(taskSummaries)
        json = _.map(json, x => {
            return {
                date: x.date,
                fullDate: new Date(x.date),
                ...x
            }
        })
        json = _.sortBy(json, 'fullDate')
        json = _.map(json, x => {
            delete x.fullDate
            return {
                ...x
            }
        })

        return json
    }
}

const az = new Az();
az.getWorkItemTracking()
