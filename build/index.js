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
(0, dotenv_1.config)();
class Az {
    constructor() {
        this.authHandler = azdev.getPersonalAccessTokenHandler(`${process.env.AZURE_PERSONAL_ACCESS_TOKEN}`);
        this.connection = new azdev.WebApi(`${process.env.ORG_URL}`, this.authHandler);
        this.DATA_PATH = "data";
        this.INPUT_PATH = "input";
        this.OUTPUT_PATH = "output";
        this.TIME_LOG_PATH = `${this.INPUT_PATH}/timelog.csv`;
        this.OUTPUT_FILE_PATH = `${this.OUTPUT_PATH}/output.xlsx`;
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
            const data = xlsx_1.default.utils.sheet_to_json(timelog.Sheets['Sheet1']);
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
                            wiTitle: wi.fields['System.Title'],
                            wiUrl: `https://symphonyvsts.visualstudio.com/VSA/_workitems/edit/${wi.id}`,
                            podLead: wi.fields['System.CreatedBy'].displayName,
                            fields: wi.fields,
                        },
                        date: this.convertExeclDate(wiInfo.date),
                        type: wiInfo.type
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
            const workItemTracking = yield this.connection.getWorkItemTrackingApi();
            const wi = yield workItemTracking.getWorkItem(Number(wiId));
            if (wi) {
                return {
                    id: (_b = wi.id) !== null && _b !== void 0 ? _b : 0,
                    wiTitle: wi.fields['System.Title'],
                    wiUrl: `https://symphonyvsts.visualstudio.com/VSA/_workitems/edit/${wi.id}`,
                    podLead: wi.fields['System.CreatedBy'].displayName,
                    fields: wi.fields
                };
            }
            return null;
        });
        this.getWorkItemsInfo = (workItemIds) => __awaiter(this, void 0, void 0, function* () {
            const workItems = [];
            for (let i = 0; i < workItemIds.length; i++) {
                const wId = workItemIds[i];
                const wi = yield this.getWorkItemInfo(Number(wId));
                if (wi && wi.fields['System.AssignedTo'].id === process.env.USER_ID) {
                    workItems.push(wi);
                }
            }
            return workItems;
        });
        this.getWorkItemPullRequest = (workItemId, pullRequests) => {
            return pullRequests.find(x => x.workItems.some((w) => w.id == workItemId));
        };
        this.getWorkItemTracking = () => __awaiter(this, void 0, void 0, function* () {
            var _c, _d, _e, _f, _g;
            const gitApi = yield this.connection.getGitApi();
            const workItemTimeLog = yield this.getWorkItemsFromTimelog();
            const pullRequests = yield gitApi.getPullRequests("VSA.Application", {
                creatorId: process.env.USER_ID,
                status: 3
            }, "VSA");
            console.log(`Có ${pullRequests.length} pull request`);
            const pullRequestWorkItems = [];
            for (let index = 0; index < pullRequests.length; index++) {
                const pr = pullRequests[index];
                const pullRequestWorkItemRefs = yield gitApi.getPullRequestWorkItemRefs("VSA.Application", (_c = pr.pullRequestId) !== null && _c !== void 0 ? _c : 0, "VSA");
                const pullRequestWorkItemIds = pullRequestWorkItemRefs.map(x => x.id);
                const workItems = yield this.getWorkItemsInfo(pullRequestWorkItemIds);
                console.log(`Get thông tin pulll request ${pr.title}`);
                pullRequestWorkItems.push({
                    title: (_d = pr.title) !== null && _d !== void 0 ? _d : "",
                    pullRequestId: (_e = pr.pullRequestId) !== null && _e !== void 0 ? _e : 0,
                    pullRequestUrl: `https://symphonyvsts.visualstudio.com/VSA/_git/VSA.Application/pullrequest/${pr.pullRequestId}`,
                    workItems
                });
            }
            const taskSummaries = [];
            for (const w of workItemTimeLog) {
                const pr = this.getWorkItemPullRequest(w.workitem.id, pullRequestWorkItems);
                taskSummaries.push({
                    date: w.date,
                    channelName: "",
                    podlead: (_g = (_f = w.workitem) === null || _f === void 0 ? void 0 : _f.podLead) !== null && _g !== void 0 ? _g : "",
                    quarter: this.getQuarterFromDate(w.date),
                    ticket: w.workitem.wiTitle,
                    workItemType: w.type,
                    pr: !!pr ? pr.pullRequestUrl : "N/A"
                });
            }
            this.writeData('finalData.json', JSON.stringify(this.sort(JSON.stringify(taskSummaries)), null, 2));
        });
        this.sort = (taskSummaries) => {
            let json = JSON.parse(taskSummaries);
            json = lodash_1.default.map(json, x => {
                return Object.assign({ date: x.date, fullDate: new Date(x.date) }, x);
            });
            json = lodash_1.default.sortBy(json, 'fullDate');
            json = lodash_1.default.map(json, x => {
                delete x.fullDate;
                return Object.assign({}, x);
            });
            return json;
        };
        if (!(0, fs_1.existsSync)(this.DATA_PATH)) {
            (0, fs_1.mkdirSync)(this.DATA_PATH);
        }
    }
    getQuarterFromDate(dateStr) {
        const date = new Date(dateStr);
        const month = date.getMonth();
        const quarter = Math.floor(month / 3) + 1;
        return `Q${quarter}`;
    }
}
const az = new Az();
az.getWorkItemTracking();
