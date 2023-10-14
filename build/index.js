"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const azdev = __importStar(require("azure-devops-node-api"));
const dotenv_1 = require("dotenv");
const fs_1 = require("fs");
const lodash_1 = __importDefault(require("lodash"));
const xlsx_1 = __importDefault(require("xlsx"));
const json_as_xlsx_1 = __importDefault(require("json-as-xlsx"));
const moment_1 = __importDefault(require("moment"));
const perf_hooks_1 = require("perf_hooks");
const helpers_1 = __importDefault(require("./helpers"));
(0, dotenv_1.config)();
class Az {
    constructor(empCode, leader) {
        this.authHandler = azdev.getPersonalAccessTokenHandler(`${process.env.AZURE_PERSONAL_ACCESS_TOKEN}`);
        this.connection = new azdev.WebApi(`${process.env.ORG_URL}`, this.authHandler);
        this.DATA_PATH = "data";
        this.INPUT_PATH = "input";
        this.OUTPUT_PATH = "output";
        this.TIME_LOG_PATH = `${this.INPUT_PATH}/timelog.csv`;
        this.FINAL_DATA = `${this.DATA_PATH}/finalData.json`;
        this.DATA = `${this.DATA_PATH}/data.json`;
        this.TASK_FILE_PATH = `${this.INPUT_PATH}/task.xlsx`;
        this.getWorkItemsFromTimelog = () => __awaiter(this, void 0, void 0, function* () {
            var _a;
            const timelog = xlsx_1.default.readFile(this.TIME_LOG_PATH);
            const data = xlsx_1.default.utils.sheet_to_json(timelog.Sheets["Sheet1"]);
            const wIds = [];
            console.log(`Có ${data.length} workItem trong timelog`);
            for (let i = 0; i < data.length; i++) {
                const wiInfo = data[i];
                console.log(`Get thông tin workitem ${wiInfo.title}`);
                const wi = yield this.getWorkItemInfo(wiInfo.workItemId);
                if (wi) {
                    wIds.push({
                        workitem: {
                            id: (_a = wi.id) !== null && _a !== void 0 ? _a : 0,
                            wiTitle: wi.fields["System.Title"],
                            wiUrl: `https://symphonyvsts.visualstudio.com/VSA/_workitems/edit/${wi.id}`,
                            podLead: wi.fields["System.CreatedBy"].displayName,
                            fields: wi.fields,
                            channelName: "",
                        },
                        date: helpers_1.default.convertExeclDate(wiInfo.date),
                        type: wiInfo.type,
                        quarter: helpers_1.default.convertExeclDate(wiInfo.date),
                    });
                }
            }
            return wIds;
        });
        this.getWorkItemInfo = (wiId) => __awaiter(this, void 0, void 0, function* () {
            var _b;
            console.log(wiId);
            const workItemTracking = yield this.connection.getWorkItemTrackingApi();
            const wi = yield workItemTracking.getWorkItem(wiId);
            console.log(wi);
            if (wi) {
                helpers_1.default.writeData(`workItems/${wiId}.json`, wi);
                return {
                    id: (_b = wi.id) !== null && _b !== void 0 ? _b : 0,
                    wiTitle: wi.fields["System.Title"],
                    wiUrl: `https://symphonyvsts.visualstudio.com/VSA/_workitems/edit/${wi.id}`,
                    podLead: wi.fields["System.CreatedBy"].displayName,
                    fields: wi.fields,
                    channelName: "",
                };
            }
            return null;
        });
        this.getWorkItemsInfo = (workItemIds) => __awaiter(this, void 0, void 0, function* () {
            const workItems = [];
            for (let i = 0; i < workItemIds.length; i++) {
                const wId = workItemIds[i];
                const wi = yield this.getWorkItemInfo(Number(wId));
                if (wi && wi.fields["System.AssignedTo"].id === process.env.USER_ID) {
                    workItems.push(wi);
                }
            }
            return workItems;
        });
        this.getWorkItemPullRequest = (workItemId, pullRequests) => {
            return pullRequests.find((x) => x.workItems.some((w) => w.id == workItemId));
        };
        this.getWorkItemTracking = (workItems) => __awaiter(this, void 0, void 0, function* () {
            var _c, _d, _e, _f, _g;
            const gitApi = yield this.connection.getGitApi();
            let workItemTimeLog;
            if (!workItems || workItems.length == 0) {
                workItemTimeLog = yield this.getWorkItemsFromTimelog();
            }
            else {
                workItemTimeLog = workItems;
            }
            if (!lodash_1.default.isEmpty(workItemTimeLog)) {
                const pullRequests = helpers_1.default.readFileJson("pullRequests.json");
                console.log(`Có ${pullRequests.length} pull request`);
                const pullRequestWorkItems = [];
                for (let index = 0; index < pullRequests.length; index++) {
                    const pr = pullRequests[index];
                    console.log(`Get thông tin pull request ${pr.title}`);
                    const pullRequestWorkItemRefs = yield gitApi.getPullRequestWorkItemRefs("VSA.Application", (_c = pr.pullRequestId) !== null && _c !== void 0 ? _c : 0, "VSA");
                    const pullRequestWorkItemIds = pullRequestWorkItemRefs.map((x) => x.id);
                    const workItems = yield this.getWorkItemsInfo(pullRequestWorkItemIds);
                    pullRequestWorkItems.push({
                        title: (_d = pr.title) !== null && _d !== void 0 ? _d : "",
                        pullRequestId: (_e = pr.pullRequestId) !== null && _e !== void 0 ? _e : 0,
                        pullRequestUrl: `https://symphonyvsts.visualstudio.com/VSA/_git/VSA.Application/pullrequest/${pr.pullRequestId}`,
                        workItems,
                    });
                }
                const taskSummaries = [];
                for (const w of workItemTimeLog) {
                    const pr = this.getWorkItemPullRequest(w.workitem.id, pullRequestWorkItems);
                    taskSummaries.push({
                        date: w.date,
                        channelName: w.workitem.channelName,
                        podlead: (_g = (_f = w.workitem) === null || _f === void 0 ? void 0 : _f.podLead) !== null && _g !== void 0 ? _g : "",
                        quarter: w.quarter,
                        ticket: w.workitem.wiTitle,
                        workItemType: w.type,
                        workItemId: w.workitem.id,
                        pr: !!pr ? pr.pullRequestUrl : "N/A",
                    });
                }
                helpers_1.default.writeData("finalData.json", helpers_1.default.sort(taskSummaries));
            }
        });
        this.exportXls = () => {
            let jsData = (0, fs_1.readFileSync)(this.DATA).toString();
            const json = JSON.parse(jsData);
            let data = [
                {
                    sheet: "Summary",
                    columns: [
                        { label: "Date", value: "date" },
                        { label: "Work Item Type", value: "workItemType" },
                        { label: "Podlead", value: "podlead" },
                        { label: "Ticket", value: "title" },
                        { label: "Pr", value: "pr" },
                        { label: "workItemId", value: "workItemId" },
                        { label: "quarter", value: "quarter" },
                        { label: "Channel Name", value: "channelName" },
                    ],
                    content: json,
                },
            ];
            let settings = {
                fileName: "MySpreadsheet",
                extraLength: 3,
                writeMode: "writeFile",
                writeOptions: {
                    type: "file",
                },
            };
            (0, json_as_xlsx_1.default)(data, settings);
        };
        this.getEmpData = () => {
            const date = new Date();
            const quarter = Math.floor(date.getMonth() / 3);
            const quarter2Dates = helpers_1.default.getQuarterDates(date, quarter);
            const allSheets = [];
            quarter2Dates.forEach((date) => {
                const shortDate = (0, moment_1.default)(date).format("ll");
                const dateArr = shortDate.split(",");
                const mday = dateArr[0];
                let sheetName = this.getSheetName(mday);
                const data = xlsx_1.default.utils.sheet_to_json(this.tasks.Sheets[sheetName]);
                if (data.length > 0) {
                    allSheets.push(sheetName);
                }
            });
            if (lodash_1.default.isEmpty(allSheets)) {
                allSheets.push("Master");
            }
            let empData = [];
            allSheets.forEach((x, i) => {
                const data = xlsx_1.default.utils.sheet_to_json(this.tasks.Sheets[x]);
                const empes = lodash_1.default.filter(data, (x) => x["Emp Code"] === this.empCode);
                lodash_1.default.forEach(empes, (emp) => {
                    const ticket = emp["Ticket URL"];
                    if (ticket && ticket.indexOf("OFF") === -1) {
                        empData.push({
                            channelName: emp["OFF AM/PM /// Teams Channel Name"],
                            date: helpers_1.default.convertExeclDate(emp["Date Created"]),
                            podlead: this.leader,
                            quarter: `Q${quarter}`,
                            ticket: ticket,
                            workItemId: this.getWorkItemIdFromTicketUrl(ticket),
                            workItemType: "",
                        });
                    }
                });
            });
            empData = lodash_1.default.orderBy(empData, (x) => new Date(x.date), "asc");
            helpers_1.default.writeData("empData.json", empData);
            helpers_1.default.writeData("allSheets.json", allSheets);
            return empData;
        };
        this.getUserWorkItems = (workItemIds) => __awaiter(this, void 0, void 0, function* () {
            const workItemTrackingApi = yield this.connection.getWorkItemTrackingApi();
            const uniqWorkitemIds = lodash_1.default.uniq(workItemIds);
            const workItems = yield workItemTrackingApi.getWorkItems(uniqWorkitemIds);
            helpers_1.default.writeData("workItems.json", workItems);
        });
        this.getPullRequestWorkItemRefs = () => __awaiter(this, void 0, void 0, function* () {
            const gitApi = yield this.connection.getGitApi();
            const pullRequests = helpers_1.default.readFileJson("pullRequests.json");
            const pullRequestWorkItemRefs = [];
            for (let index = 0; index < pullRequests.length; index++) {
                const pullrequest = pullRequests[index];
                console.log(`Get ${pullrequest.title}...`);
                const pullRequestWorkItemRef = yield gitApi.getPullRequestWorkItemRefs("VSA.Application", pullrequest.pullRequestId, "VSA");
                if (pullRequestWorkItemRef) {
                    lodash_1.default.forEach(pullRequestWorkItemRef, (x) => {
                        pullRequestWorkItemRefs.push({
                            workItemId: x.id,
                            pullRequestId: pullrequest.pullRequestId,
                        });
                    });
                }
            }
            helpers_1.default.writeData("pullRequestWorkItemRefs.json", lodash_1.default.groupBy(pullRequestWorkItemRefs, "workItemId"));
        });
        this.init = (first) => __awaiter(this, void 0, void 0, function* () {
            if (first) {
                const tempData = this.getEmpData();
                yield this.getUserPullRequest();
                yield this.getUserWorkItems(lodash_1.default.map(tempData, (x) => x.workItemId));
                yield this.getPullRequestWorkItemRefs();
            }
            return true;
        });
        this.run = () => {
            const empData = helpers_1.default.readFileJson("empData.json");
            const pullRequests = helpers_1.default.readFileJson("pullRequests.json");
            const workItems = helpers_1.default.readFileJson("workItems.json");
            const pullRequestWorkItemRefs = helpers_1.default.readFileJson("pullRequestWorkItemRefs.json");
            const finalData = [];
            for (let index = 0; index < empData.length; index++) {
                const data = empData[index];
                const workItem = lodash_1.default.find(workItems, (x) => x.id === data.workItemId);
                if (workItem) {
                    data.workItemType = lodash_1.default.get(workItem.fields, "System.WorkItemType");
                    data.title = lodash_1.default.get(workItem.fields, "System.Title");
                    const bugId = helpers_1.default.getBugIdFromTicket(data.title);
                    const pullRequestWorkItemRef = pullRequestWorkItemRefs[bugId];
                    if (pullRequestWorkItemRef) {
                        const _prs = lodash_1.default.uniq(lodash_1.default.map(pullRequestWorkItemRef, (x) => x.pullRequestId));
                        data.pr = [];
                        if (_prs) {
                            lodash_1.default.forEach(_prs, (prId) => {
                                const pr = lodash_1.default.find(pullRequests, (p) => p.pullRequestId === prId);
                                if (pr) {
                                    data.pr.push({
                                        url: `https://symphonyvsts.visualstudio.com/VSA/_git/VSA.Application/pullrequest/${pr.pullRequestId}`,
                                        title: pr.title,
                                    });
                                }
                            });
                        }
                    }
                }
                finalData.push(data);
            }
            helpers_1.default.writeData("finalData.json", finalData);
        };
        this.flatData = () => {
            const finalData = helpers_1.default.readFileJson("finalData.json");
            const data = [];
            lodash_1.default.forEach(finalData, (d) => {
                if (lodash_1.default.isEmpty(d.pr)) {
                    data.push(Object.assign(Object.assign({}, d), { pr: "N/A" }));
                }
                else {
                    lodash_1.default.forEach(d.pr, (pr) => {
                        data.push(Object.assign(Object.assign({}, d), { pr: pr.url }));
                    });
                }
            });
            helpers_1.default.writeData("data.json", data);
        };
        if (!(0, fs_1.existsSync)(this.DATA_PATH)) {
            (0, fs_1.mkdirSync)(this.DATA_PATH);
        }
        this.tasks = xlsx_1.default.readFile(this.TASK_FILE_PATH);
        this.empCode = empCode;
        this.leader = leader;
    }
    getWorkItemIdFromTicketUrl(url) {
        const urlArr = url.split("/");
        const urlArrLength = urlArr.length;
        let lastItem = urlArr[urlArrLength - 1];
        if (!lastItem) {
            lastItem = urlArr[urlArrLength - 1 - 1];
        }
        return Number(`${lastItem}`);
    }
    getSheetName(x) {
        const sheetNameArr = x.split(" ");
        let date = sheetNameArr[1];
        if (date.length === 1) {
            date = `0${date}`;
        }
        return [sheetNameArr[0], date].join(" ");
    }
    getUserPullRequest() {
        return __awaiter(this, void 0, void 0, function* () {
            const t0 = perf_hooks_1.performance.now();
            var gitApi = yield this.connection.getGitApi();
            const pullRequests = yield gitApi.getPullRequests("VSA.Application", {
                creatorId: process.env.USER_ID,
                status: 3,
            }, "VSA");
            helpers_1.default.writeData("pullRequests.json", pullRequests);
            const t1 = perf_hooks_1.performance.now();
            console.log(`Call to getUserPullRequest took ${t1 - t0} milliseconds.`);
            return pullRequests;
        });
    }
}
const az = new Az(411, "Duy Ba Nguyen");
az.init(false)
    .catch(console.log)
    .then((x) => {
    if (x) {
        az.run();
        az.flatData();
        az.exportXls();
    }
});
