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
const path_1 = require("path");
const perf_hooks_1 = require("perf_hooks");
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
        this.TASK_FILE_PATH = `${this.INPUT_PATH}/task.xlsx`;
        this.convertExeclDate = (excelSerialDate) => {
            const unixTimestamp = (excelSerialDate - 25569) * 86400;
            const dateObj = new Date(unixTimestamp * 1000);
            const month = dateObj.getMonth() + 1;
            const day = dateObj.getDate();
            const year = dateObj.getFullYear();
            const formattedDate = `${month}/${day}/${year}`;
            return formattedDate;
        };
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
                        date: this.convertExeclDate(wiInfo.date),
                        type: wiInfo.type,
                        quarter: this.convertExeclDate(wiInfo.date),
                    });
                }
            }
            return wIds;
        });
        this.writeData = (path, data) => {
            (0, fs_1.writeFileSync)(`data/${path}`, JSON.stringify(data, null, 2));
        };
        this.getWorkItemInfo = (wiId) => __awaiter(this, void 0, void 0, function* () {
            var _b;
            console.log(wiId);
            const workItemTracking = yield this.connection.getWorkItemTrackingApi();
            const wi = yield workItemTracking.getWorkItem(wiId);
            console.log(wi);
            if (wi) {
                this.writeData(`workItems/${wiId}.json`, wi);
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
                const pullRequests = yield this.getUserPullRequest(gitApi);
                console.log(`Có ${pullRequests.length} pull request`);
                const pullRequestWorkItems = [];
                for (let index = 0; index < pullRequests.length; index++) {
                    const pr = pullRequests[index];
                    const pullRequestWorkItemRefs = yield gitApi.getPullRequestWorkItemRefs("VSA.Application", (_c = pr.pullRequestId) !== null && _c !== void 0 ? _c : 0, "VSA");
                    const pullRequestWorkItemIds = pullRequestWorkItemRefs.map((x) => x.id);
                    const workItems = yield this.getWorkItemsInfo(pullRequestWorkItemIds);
                    console.log(`Get thông tin pull request ${pr.title}`);
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
                this.writeData("finalData.json", this.sort(taskSummaries));
            }
        });
        this.sort = (taskSummaries) => {
            let json = taskSummaries;
            json = lodash_1.default.map(json, (x) => {
                return Object.assign({ date: x.date, fullDate: new Date(x.date) }, x);
            });
            json = lodash_1.default.sortBy(json, "fullDate");
            json = lodash_1.default.map(json, (x) => {
                delete x.fullDate;
                return Object.assign({}, x);
            });
            return JSON.stringify(json, null, 2);
        };
        this.parse = () => {
            let jsData = (0, fs_1.readFileSync)(this.FINAL_DATA).toString();
            jsData = JSON.parse(jsData);
            (0, fs_1.writeFileSync)(this.FINAL_DATA, JSON.stringify(jsData, null, 2));
        };
        this.exportXls = () => {
            let jsData = (0, fs_1.readFileSync)(this.FINAL_DATA).toString();
            const json = JSON.parse(jsData);
            let data = [
                {
                    sheet: "Summary",
                    columns: [
                        { label: "Date", value: "date" },
                        { label: "Work Item Type", value: "workItemType" },
                        { label: "Podlead", value: "podlead" },
                        { label: "Ticket", value: "ticket" },
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
        this.removePrs = () => __awaiter(this, void 0, void 0, function* () {
            const gitApi = yield this.connection.getTfvcApi();
            let branches = yield gitApi.getBranches("VSA");
            let myBranches = branches.map((x) => {
                return Object.assign({}, x);
            });
            (0, fs_1.writeFileSync)(`${this.DATA_PATH}/branches.json`, JSON.stringify(myBranches, null, 2));
        });
        this.getEmpData = () => {
            const date = new Date();
            const quarter = Math.floor(date.getMonth() / 3);
            const quarter2Dates = this.getQuarterDates(date, quarter);
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
                            date: this.convertExeclDate(emp["Date Created"]),
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
            this.writeData("empData.json", empData);
            this.writeData("allSheets.json", allSheets);
            return empData;
        };
        this.getBugIdFromTicket = (ticket) => {
            const ticketArr = lodash_1.default.split(ticket, " ");
            const bugIndex = lodash_1.default.indexOf(ticketArr, "Bug");
            if (bugIndex > -1) {
                return +`${lodash_1.default.replace(ticketArr[bugIndex + 1], ":", "")}`;
            }
            return 0;
        };
        this.mappingPullRequest = () => __awaiter(this, void 0, void 0, function* () {
            const finalData = this.readFileJson("finalData.json");
            const devTaskDefect = lodash_1.default.filter(finalData, (x) => x.workItemType === "Dev Task_Defect");
            for (let index = 0; index < devTaskDefect.length; index++) {
                const devTask = devTaskDefect[index];
                const { ticket } = devTask;
                const bugId = this.getBugIdFromTicket(ticket);
                if (bugId) {
                    const wi = yield this.getWorkItemInfo(bugId);
                    break;
                }
            }
        });
        if (!(0, fs_1.existsSync)(this.DATA_PATH)) {
            (0, fs_1.mkdirSync)(this.DATA_PATH);
        }
        this.tasks = xlsx_1.default.readFile(this.TASK_FILE_PATH);
        this.empCode = empCode;
        this.leader = leader;
    }
    getQuarterFromDate(dateStr) {
        const date = new Date(dateStr);
        const month = date.getMonth();
        const quarter = Math.floor(month / 3) + 1;
        return `Q${quarter}`;
    }
    getQuarterDates(date, quarter) {
        const year = date.getFullYear();
        const quarterStartMonth = 3 * quarter - 2;
        const startDate = new Date(year, quarterStartMonth - 1, 1);
        const endDate = new Date(year, quarterStartMonth + 2, 0);
        const dates = [];
        let currentDate = startDate;
        while (currentDate <= endDate) {
            dates.push(new Date(currentDate));
            currentDate.setDate(currentDate.getDate() + 1);
        }
        return dates;
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
    getWorkItemFromDailyTask() {
        return __awaiter(this, void 0, void 0, function* () {
            const workItemJson = this.readFileJson("workItems.json");
            let workItems = [];
            if (lodash_1.default.size(workItemJson) > 0) {
                workItems = workItemJson;
            }
            else {
                const empData = this.getEmpData();
                for (let index = 0; index < empData.length; index++) {
                    const element = empData[index];
                    const wi = yield this.getWorkItemInfo(element.workItemId);
                    if (wi) {
                        console.log(`Retrieving ${wi.wiTitle}...`);
                        workItems.push({
                            workitem: {
                                id: element.workItemId,
                                podLead: element.podlead,
                                wiTitle: wi.wiTitle,
                                wiUrl: element.ticket,
                                fields: wi.fields,
                                channelName: element.channelName,
                            },
                            quarter: element.quarter,
                            date: element.date,
                            type: wi.fields ? wi.fields["System.WorkItemType"] : "",
                        });
                    }
                    else {
                        console.log(element.workItemId);
                        console.log(element.date);
                    }
                }
                this.writeData("workItems.json", workItems);
            }
            yield this.getWorkItemTracking(workItems);
            return "Ok!";
        });
    }
    readFileJson(fileName) {
        const filePath = (0, path_1.join)(__dirname, "../", this.DATA_PATH, fileName);
        if (!(0, fs_1.existsSync)(filePath)) {
            return null;
        }
        const str = (0, fs_1.readFileSync)(filePath).toString();
        try {
            const data = JSON.parse(str);
            return data;
        }
        catch (error) {
            return null;
        }
    }
    remapPullRequest() {
        return __awaiter(this, void 0, void 0, function* () {
            const data = this.readFileJson("finalData.json");
            const missingPrTask = lodash_1.default.filter(data, (x) => x.pr === "N/A");
            for (let index = 0; index < missingPrTask.length; index++) {
                const element = missingPrTask[index];
                this.getWorkItemPullRequest;
            }
        });
    }
    getUserPullRequest(gitApi = null) {
        return __awaiter(this, void 0, void 0, function* () {
            const t0 = perf_hooks_1.performance.now();
            if (!gitApi) {
                gitApi = yield this.connection.getGitApi();
            }
            let pullRequests = this.readFileJson("pullRequests.json");
            if (lodash_1.default.isEmpty(pullRequests)) {
                pullRequests = yield gitApi.getPullRequests("VSA.Application", {
                    creatorId: process.env.USER_ID,
                    status: 3,
                }, "VSA");
                this.writeData("pullRequests.json", pullRequests);
            }
            const t1 = perf_hooks_1.performance.now();
            console.log(`Call to getUserPullRequest took ${t1 - t0} milliseconds.`);
            return pullRequests;
        });
    }
    test() {
        return __awaiter(this, void 0, void 0, function* () { });
    }
}
const az = new Az(411, "Duy Ba Nguyen");
// az.getWorkItemFromDailyTask()
//   .catch((x) => console.log(x.message))
//   .then((x) => {
//     az.parse()
//   })
az.mappingPullRequest().catch(console.log);
