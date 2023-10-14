"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const fs_1 = require("fs");
const lodash_1 = __importDefault(require("lodash"));
const path_1 = require("path");
class Helpers {
    static readFileJson(fileName) {
        const filePath = (0, path_1.join)(__dirname, "../data", fileName);
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
    static getQuarterDates(date, quarter) {
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
    static getQuarterFromDate(dateStr) {
        const date = new Date(dateStr);
        const month = date.getMonth();
        const quarter = Math.floor(month / 3) + 1;
        return `Q${quarter}`;
    }
}
Helpers.writeData = (path, data) => {
    (0, fs_1.writeFileSync)(`data/${path}`, JSON.stringify(data, null, 2));
};
Helpers.sort = (taskSummaries) => {
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
Helpers.getBugIdFromTicket = (ticket) => {
    const ticketArr = lodash_1.default.split(ticket, " ");
    const bugIndex = lodash_1.default.indexOf(ticketArr, "Bug");
    if (bugIndex > -1) {
        return +`${lodash_1.default.replace(ticketArr[bugIndex + 1], ":", "")}`;
    }
    return 0;
};
Helpers.convertExeclDate = (excelSerialDate) => {
    const unixTimestamp = (excelSerialDate - 25569) * 86400;
    const dateObj = new Date(unixTimestamp * 1000);
    const month = dateObj.getMonth() + 1;
    const day = dateObj.getDate();
    const year = dateObj.getFullYear();
    const formattedDate = `${month}/${day}/${year}`;
    return formattedDate;
};
exports.default = Helpers;
