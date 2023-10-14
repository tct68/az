import * as azdev from "azure-devops-node-api"
import { GitPullRequest } from "azure-devops-node-api/interfaces/GitInterfaces"
import WorkItemTrackingInterfaces from "azure-devops-node-api/interfaces/WorkItemTrackingInterfaces"
import VSSInterfaces from "azure-devops-node-api/interfaces/common/VSSInterfaces"
import { config } from "dotenv"
import { existsSync, mkdirSync, readFileSync } from "fs"
import _ from "lodash"
import xlsx from "xlsx"
import jxlsx, { IJsonSheet, ISettings } from "json-as-xlsx"
import moment from "moment"
import { performance } from "perf_hooks"
import Helpers from "./helpers"
import { PullRequest, PullRequestWorkItemRefs, TimeLogWorkItem, WorkItem } from "./models/pullRequestModel"
config()

class Az {
  private authHandler = azdev.getPersonalAccessTokenHandler(`${process.env.AZURE_PERSONAL_ACCESS_TOKEN}`)
  private connection = new azdev.WebApi(`${process.env.ORG_URL}`, this.authHandler)
  private readonly DATA_PATH = "data"
  private readonly INPUT_PATH = "input"
  private readonly OUTPUT_PATH = "output"
  private readonly TIME_LOG_PATH = `${this.INPUT_PATH}/timelog.csv`
  private readonly FINAL_DATA = `${this.DATA_PATH}/finalData.json`
  private readonly DATA = `${this.DATA_PATH}/data.json`
  private readonly TASK_FILE_PATH = `${this.INPUT_PATH}/task.xlsx`
  private readonly tasks: xlsx.WorkBook
  private readonly empCode: any
  private readonly leader: string

  constructor(empCode: any, leader: string) {
    if (!existsSync(this.DATA_PATH)) {
      mkdirSync(this.DATA_PATH)
    }
    this.tasks = xlsx.readFile(this.TASK_FILE_PATH)
    this.empCode = empCode
    this.leader = leader
  }

  getWorkItemsFromTimelog = async (): Promise<TimeLogWorkItem[]> => {
    const timelog = xlsx.readFile(this.TIME_LOG_PATH)
    const data = xlsx.utils.sheet_to_json(timelog.Sheets["Sheet1"])
    const wIds: TimeLogWorkItem[] = []
    console.log(`Có ${data.length} workItem trong timelog`)
    for (let i = 0; i < data.length; i++) {
      const wiInfo: any = data[i]
      console.log(`Get thông tin workitem ${wiInfo.title}`)
      const wi = await this.getWorkItemInfo(wiInfo.workItemId)
      if (wi) {
        wIds.push({
          workitem: {
            id: wi.id ?? 0,
            wiTitle: wi.fields!["System.Title"],
            wiUrl: `https://symphonyvsts.visualstudio.com/VSA/_workitems/edit/${wi.id}`,
            podLead: wi.fields!["System.CreatedBy"].displayName,
            fields: wi.fields,
            channelName: "",
          },
          date: Helpers.convertExeclDate(wiInfo.date),
          type: wiInfo.type,
          quarter: Helpers.convertExeclDate(wiInfo.date),
        })
      }
    }
    return wIds
  }

  getWorkItemInfo = async (wiId: number): Promise<WorkItem | null> => {
    console.log(wiId)
    const workItemTracking = await this.connection.getWorkItemTrackingApi()
    const wi = await workItemTracking.getWorkItem(wiId)
    console.log(wi)
    if (wi) {
      Helpers.writeData(`workItems/${wiId}.json`, wi)
      return {
        id: wi.id ?? 0,
        wiTitle: wi.fields!["System.Title"],
        wiUrl: `https://symphonyvsts.visualstudio.com/VSA/_workitems/edit/${wi.id}`,
        podLead: wi.fields!["System.CreatedBy"].displayName,
        fields: wi.fields,
        channelName: "",
      }
    }

    return null
  }

  getWorkItemsInfo = async (workItemIds: (string | undefined)[]): Promise<WorkItem[]> => {
    const workItems: WorkItem[] = []
    for (let i = 0; i < workItemIds.length; i++) {
      const wId = workItemIds[i]
      const wi = await this.getWorkItemInfo(Number(wId))
      if (wi && wi.fields!["System.AssignedTo"].id === process.env.USER_ID) {
        workItems.push(wi)
      }
    }

    return workItems
  }

  getWorkItemPullRequest = (workItemId: Number, pullRequests: PullRequest[]) => {
    return pullRequests.find((x) => x.workItems.some((w: WorkItem) => w.id == workItemId))
  }

  getWorkItemTracking = async (workItems: TimeLogWorkItem[]) => {
    const gitApi = await this.connection.getGitApi()
    let workItemTimeLog: TimeLogWorkItem[]

    if (!workItems || workItems.length == 0) {
      workItemTimeLog = await this.getWorkItemsFromTimelog()
    } else {
      workItemTimeLog = workItems
    }

    if (!_.isEmpty(workItemTimeLog)) {
      const pullRequests = Helpers.readFileJson("pullRequests.json")
      console.log(`Có ${pullRequests.length} pull request`)

      const pullRequestWorkItems: PullRequest[] = []
      for (let index = 0; index < pullRequests.length; index++) {
        const pr = pullRequests[index]
        console.log(`Get thông tin pull request ${pr.title}`)
        const pullRequestWorkItemRefs = await gitApi.getPullRequestWorkItemRefs("VSA.Application", pr.pullRequestId ?? 0, "VSA")
        const pullRequestWorkItemIds = pullRequestWorkItemRefs.map((x) => x.id)
        const workItems = await this.getWorkItemsInfo(pullRequestWorkItemIds)

        pullRequestWorkItems.push({
          title: pr.title ?? "",
          pullRequestId: pr.pullRequestId ?? 0,
          pullRequestUrl: `https://symphonyvsts.visualstudio.com/VSA/_git/VSA.Application/pullrequest/${pr.pullRequestId}`,
          workItems,
        })
      }

      const taskSummaries: TaskSummary[] = []
      for (const w of workItemTimeLog) {
        const pr = this.getWorkItemPullRequest(w.workitem.id, pullRequestWorkItems)

        taskSummaries.push({
          date: w.date,
          channelName: w.workitem.channelName,
          podlead: w.workitem?.podLead ?? "",
          quarter: w.quarter,
          ticket: w.workitem.wiTitle,
          workItemType: w.type,
          workItemId: w.workitem.id,
          pr: !!pr ? pr.pullRequestUrl : "N/A",
        })
      }

      Helpers.writeData("finalData.json", Helpers.sort(taskSummaries))
    }
  }

  exportXls = () => {
    console.log(this.exportXls.name)
    let jsData = readFileSync(this.DATA).toString()
    const json = JSON.parse(jsData)

    let data: IJsonSheet[] = [
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
    ]

    let settings: ISettings = {
      fileName: "MySpreadsheet",
      extraLength: 3,
      writeMode: "writeFile",
      writeOptions: {
        type: "file",
      },
    }

    jxlsx(data, settings)
  }

  getWorkItemIdFromTicketUrl(url: string) {
    const urlArr: string[] = url.split("/")
    const urlArrLength = urlArr.length
    let lastItem = urlArr[urlArrLength - 1]
    if (!lastItem) {
      lastItem = urlArr[urlArrLength - 1 - 1]
    }

    return Number(`${lastItem}`)
  }

  getSheetName(x: string) {
    const sheetNameArr = x.split(" ")
    let date = sheetNameArr[1]
    if (date.length === 1) {
      date = `0${date}`
    }

    return [sheetNameArr[0], date].join(" ")
  }

  getEmpData = () => {
    console.log(this.getEmpData.name)
    const date = new Date()
    const quarter = Math.floor(date.getMonth() / 3)
    const quarter2Dates = Helpers.getQuarterDates(date, quarter)
    const allSheets: string[] = []

    quarter2Dates.forEach((date) => {
      const shortDate = moment(date).format("ll")
      const dateArr = shortDate.split(",")
      const mday = dateArr[0]
      let sheetName = this.getSheetName(mday)

      const data = xlsx.utils.sheet_to_json(this.tasks.Sheets[sheetName])
      if (data.length > 0) {
        allSheets.push(sheetName)
      }
    })

    if (_.isEmpty(allSheets)) {
      allSheets.push("Master")
    }

    let empData: TaskSummary[] = []
    allSheets.forEach((x, i) => {
      const data = xlsx.utils.sheet_to_json(this.tasks.Sheets[x])
      const empes: any = _.filter(data, (x: any) => x["Emp Code"] === this.empCode)
      _.forEach(empes, (emp) => {
        const ticket = emp["Ticket URL"]
        if (ticket && ticket.indexOf("OFF") === -1) {
          empData.push({
            channelName: emp["OFF AM/PM /// Teams Channel Name"],
            date: Helpers.convertExeclDate(emp["Date Created"]),
            podlead: this.leader,
            quarter: `Q${quarter}`,
            ticket: ticket,
            workItemId: this.getWorkItemIdFromTicketUrl(ticket),
            workItemType: "",
          })
        }
      })
    })

    empData = _.orderBy(empData, (x) => new Date(x.date), "asc")

    Helpers.writeData("empData.json", empData)
    Helpers.writeData("allSheets.json", allSheets)
    return empData
  }

  async getUserPullRequest() {
    console.log(this.getUserPullRequest.name)
    const t0 = performance.now()
    var gitApi = await this.connection.getGitApi()

    const pullRequests = await gitApi.getPullRequests(
      "VSA.Application",
      {
        creatorId: process.env.USER_ID,
        status: 3,
      },
      "VSA"
    )

    Helpers.writeData("pullRequests.json", pullRequests)

    const t1 = performance.now()
    console.log(`Call to getUserPullRequest took ${t1 - t0} milliseconds.`)

    return pullRequests
  }

  getUserWorkItems = async (workItemIds: any[]) => {
    console.log(this.getUserWorkItems.name)
    const workItemTrackingApi = await this.connection.getWorkItemTrackingApi()
    const uniqWorkitemIds = _.uniq(workItemIds)
    const workItems = await workItemTrackingApi.getWorkItems(uniqWorkitemIds)
    Helpers.writeData("workItems.json", workItems)
  }

  getPullRequestWorkItemRefs = async () => {
    console.log(this.getPullRequestWorkItemRefs.name)
    const gitApi = await this.connection.getGitApi()
    const pullRequests = Helpers.readFileJson("pullRequests.json")
    const pullRequestWorkItemRefs: any[] = []
    for (let index = 0; index < pullRequests.length; index++) {
      const pullrequest = pullRequests[index]
      console.log(`Get ${pullrequest.title}...`)

      const pullRequestWorkItemRef = await gitApi.getPullRequestWorkItemRefs("VSA.Application", pullrequest.pullRequestId, "VSA")

      if (pullRequestWorkItemRef) {
        _.forEach(pullRequestWorkItemRef, (x: VSSInterfaces.ResourceRef) => {
          pullRequestWorkItemRefs.push({
            workItemId: x.id,
            pullRequestId: pullrequest.pullRequestId,
          })
        })
      }
    }

    Helpers.writeData("pullRequestWorkItemRefs.json", _.groupBy(pullRequestWorkItemRefs, "workItemId"))
  }

  init = async (first: boolean) => {
    if (first) {
      const tempData = this.getEmpData()
      await this.getUserPullRequest()
      await this.getUserWorkItems(_.map(tempData, (x) => x.workItemId))
      await this.getPullRequestWorkItemRefs()
      return true
    }

    return true
  }

  run = () => {
    console.log(this.run.name)
    const empData: TaskSummary[] = Helpers.readFileJson("empData.json")
    const pullRequests: GitPullRequest[] = Helpers.readFileJson("pullRequests.json")
    const workItems: WorkItemTrackingInterfaces.WorkItem[] = Helpers.readFileJson("workItems.json")
    const pullRequestWorkItemRefs: any = Helpers.readFileJson("pullRequestWorkItemRefs.json")

    const finalData: any[] = []
    for (let index = 0; index < empData.length; index++) {
      const data: any = empData[index]
      const workItem = _.find(workItems, (x: WorkItemTrackingInterfaces.WorkItem) => x.id === data.workItemId)
      if (workItem) {
        data.workItemType = _.get(workItem.fields, "System.WorkItemType")
        data.title = _.get(workItem.fields, "System.Title")

        const bugId = Helpers.getBugIdFromTicket(data.title)
        const pullRequestWorkItemRef = pullRequestWorkItemRefs[bugId]
        if (pullRequestWorkItemRef) {
          const _prs = _.uniq(_.map(pullRequestWorkItemRef, (x: any) => x.pullRequestId))
          data.pr = []
          if (_prs) {
            _.forEach(_prs, (prId: any) => {
              const pr = _.find(pullRequests, (p: GitPullRequest) => p.pullRequestId === prId)
              if (pr) {
                data.pr.push({
                  url: `https://symphonyvsts.visualstudio.com/VSA/_git/VSA.Application/pullrequest/${pr.pullRequestId}`,
                  title: pr.title,
                })
              }
            })
          }
        }
      }

      finalData.push(data)
    }

    Helpers.writeData("finalData.json", finalData)
  }

  flatData = () => {
    console.log(this.flatData.name)
    const finalData = Helpers.readFileJson("finalData.json")
    const data: any[] = []
    _.forEach(finalData, (d) => {
      if (_.isEmpty(d.pr)) {
        data.push({
          ...d,
          pr: "N/A",
        })
      } else {
        _.forEach(d.pr, (pr) => {
          data.push({
            ...d,
            pr: pr.url,
          })
        })
      }
    })

    Helpers.writeData("data.json", data)
  }
}

const az = new Az(413, "Duy Ba Nguyen")
az.init(true)
  .catch(console.log)
  .then((x) => {
    if (x) {
      az.run()
      az.flatData()
      az.exportXls()
    }
  })
