import * as azdev from "azure-devops-node-api"
import { config } from "dotenv"
import { writeFileSync, existsSync, mkdirSync, readFileSync } from "fs"
import _ from "lodash"
import xlsx from "xlsx"
import jxlsx, { IJsonSheet, ISettings } from "json-as-xlsx"
import moment from "moment"
import { join } from "path"
import { IGitApi } from "azure-devops-node-api/GitApi"
import { performance } from "perf_hooks"
config()

class Az {
  private authHandler = azdev.getPersonalAccessTokenHandler(
    `${process.env.AZURE_PERSONAL_ACCESS_TOKEN}`
  )
  private connection = new azdev.WebApi(
    `${process.env.ORG_URL}`,
    this.authHandler
  )
  private readonly DATA_PATH = "data"
  private readonly INPUT_PATH = "input"
  private readonly OUTPUT_PATH = "output"
  private readonly TIME_LOG_PATH = `${this.INPUT_PATH}/timelog.csv`
  private readonly FINAL_DATA = `${this.DATA_PATH}/finalData.json`
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

  convertExeclDate = (excelSerialDate: any) => {
    const unixTimestamp = (excelSerialDate - 25569) * 86400
    const dateObj = new Date(unixTimestamp * 1000)

    const month = dateObj.getMonth() + 1
    const day = dateObj.getDate()
    const year = dateObj.getFullYear()

    const formattedDate = `${month}/${day}/${year}`
    return formattedDate
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
          date: this.convertExeclDate(wiInfo.date),
          type: wiInfo.type,
          quarter: this.convertExeclDate(wiInfo.date),
        })
      }
    }
    return wIds
  }

  writeData = (path: String, data: any) => {
    writeFileSync(`data/${path}`, JSON.stringify(data, null, 2))
  }

  getQuarterFromDate(dateStr: string) {
    const date = new Date(dateStr)
    const month = date.getMonth()
    const quarter = Math.floor(month / 3) + 1
    return `Q${quarter}`
  }

  getWorkItemInfo = async (wiId: number): Promise<WorkItem | null> => {
    console.log(wiId)
    const workItemTracking = await this.connection.getWorkItemTrackingApi()
    const wi = await workItemTracking.getWorkItem(wiId)
    console.log(wi)
    if (wi) {
      this.writeData(`workItems/${wiId}.json`, wi)
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

  getWorkItemsInfo = async (
    workItemIds: (string | undefined)[]
  ): Promise<WorkItem[]> => {
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

  getWorkItemPullRequest = (
    workItemId: Number,
    pullRequests: PullRequest[]
  ) => {
    return pullRequests.find((x) =>
      x.workItems.some((w: WorkItem) => w.id == workItemId)
    )
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
      const pullRequests = await this.getUserPullRequest(gitApi)
      console.log(`Có ${pullRequests.length} pull request`)

      const pullRequestWorkItems: PullRequest[] = []
      for (let index = 0; index < pullRequests.length; index++) {
        const pr = pullRequests[index]
        const pullRequestWorkItemRefs = await gitApi.getPullRequestWorkItemRefs(
          "VSA.Application",
          pr.pullRequestId ?? 0,
          "VSA"
        )
        const pullRequestWorkItemIds = pullRequestWorkItemRefs.map((x) => x.id)
        const workItems = await this.getWorkItemsInfo(pullRequestWorkItemIds)
        console.log(`Get thông tin pull request ${pr.title}`)

        pullRequestWorkItems.push({
          title: pr.title ?? "",
          pullRequestId: pr.pullRequestId ?? 0,
          pullRequestUrl: `https://symphonyvsts.visualstudio.com/VSA/_git/VSA.Application/pullrequest/${pr.pullRequestId}`,
          workItems,
        })
      }

      const taskSummaries: TaskSummary[] = []
      for (const w of workItemTimeLog) {
        const pr = this.getWorkItemPullRequest(
          w.workitem.id,
          pullRequestWorkItems
        )

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

      this.writeData("finalData.json", this.sort(taskSummaries))
    }
  }

  sort = (taskSummaries: any[]) => {
    let json = taskSummaries
    json = _.map(json, (x) => {
      return {
        date: x.date,
        fullDate: new Date(x.date),
        ...x,
      }
    })
    json = _.sortBy(json, "fullDate")
    json = _.map(json, (x) => {
      delete x.fullDate
      return {
        ...x,
      }
    })

    return JSON.stringify(json, null, 2)
  }

  parse = () => {
    let jsData = readFileSync(this.FINAL_DATA).toString()
    jsData = JSON.parse(jsData)
    writeFileSync(this.FINAL_DATA, JSON.stringify(jsData, null, 2))
  }

  exportXls = () => {
    let jsData = readFileSync(this.FINAL_DATA).toString()
    const json = JSON.parse(jsData)

    let data: IJsonSheet[] = [
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

  removePrs = async () => {
    const gitApi = await this.connection.getTfvcApi()
    let branches = await gitApi.getBranches("VSA")

    let myBranches = branches.map((x) => {
      return {
        ...x,
      }
    })

    writeFileSync(
      `${this.DATA_PATH}/branches.json`,
      JSON.stringify(myBranches, null, 2)
    )
  }

  getQuarterDates(date: Date, quarter: number) {
    const year = date.getFullYear()
    const quarterStartMonth = 3 * quarter - 2
    const startDate = new Date(year, quarterStartMonth - 1, 1)
    const endDate = new Date(year, quarterStartMonth + 2, 0)

    const dates = []
    let currentDate = startDate

    while (currentDate <= endDate) {
      dates.push(new Date(currentDate))
      currentDate.setDate(currentDate.getDate() + 1)
    }

    return dates
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
    const date = new Date()
    const quarter = Math.floor(date.getMonth() / 3)
    const quarter2Dates = this.getQuarterDates(date, quarter)
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
      const empes: any = _.filter(
        data,
        (x: any) => x["Emp Code"] === this.empCode
      )
      _.forEach(empes, (emp) => {
        const ticket = emp["Ticket URL"]
        if (ticket && ticket.indexOf("OFF") === -1) {
          empData.push({
            channelName: emp["OFF AM/PM /// Teams Channel Name"],
            date: this.convertExeclDate(emp["Date Created"]),
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

    this.writeData("empData.json", empData)
    this.writeData("allSheets.json", allSheets)
    return empData
  }

  async getWorkItemFromDailyTask() {
    const workItemJson = this.readFileJson("workItems.json")

    let workItems: TimeLogWorkItem[] = []
    if (_.size(workItemJson) > 0) {
      workItems = workItemJson
    } else {
      const empData = this.getEmpData()
      for (let index = 0; index < empData.length; index++) {
        const element = empData[index]
        const wi = await this.getWorkItemInfo(element.workItemId)
        if (wi) {
          console.log(`Retrieving ${wi.wiTitle}...`)

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
          })
        } else {
          console.log(element.workItemId)
          console.log(element.date)
        }
      }

      this.writeData("workItems.json", workItems)
    }
    await this.getWorkItemTracking(workItems)

    return "Ok!"
  }

  readFileJson(fileName: string) {
    const filePath = join(__dirname, "../", this.DATA_PATH, fileName)
    if (!existsSync(filePath)) {
      return null
    }

    const str = readFileSync(filePath).toString()
    try {
      const data = JSON.parse(str)
      return data
    } catch (error) {
      return null
    }
  }

  async remapPullRequest() {
    const data: TaskSummary[] = this.readFileJson("finalData.json")
    const missingPrTask = _.filter(data, (x) => x.pr === "N/A")
    for (let index = 0; index < missingPrTask.length; index++) {
      const element = missingPrTask[index]
      this.getWorkItemPullRequest
    }
  }

  async getUserPullRequest(gitApi: IGitApi | undefined | null = null) {
    const t0 = performance.now()
    if (!gitApi) {
      gitApi = await this.connection.getGitApi()
    }

    let pullRequests = this.readFileJson("pullRequests.json")
    if (_.isEmpty(pullRequests)) {
      pullRequests = await gitApi.getPullRequests(
        "VSA.Application",
        {
          creatorId: process.env.USER_ID,
          status: 3,
        },
        "VSA"
      )

      this.writeData("pullRequests.json", pullRequests)
    }
    const t1 = performance.now()
    console.log(`Call to getUserPullRequest took ${t1 - t0} milliseconds.`)

    return pullRequests
  }

  async test() {}

  getBugIdFromTicket = (ticket: string) => {
    const ticketArr = _.split(ticket, " ")
    const bugIndex = _.indexOf(ticketArr, "Bug")
    if (bugIndex > -1) {
      return +`${_.replace(ticketArr[bugIndex + 1], ":", "")}`
    }

    return 0
  }

  mappingPullRequest = async () => {
    const finalData = this.readFileJson("finalData.json")
    const devTaskDefect = _.filter(
      finalData,
      (x) => x.workItemType === "Dev Task_Defect"
    )

    for (let index = 0; index < devTaskDefect.length; index++) {
      const devTask = devTaskDefect[index]
      const { ticket } = devTask
      const bugId = this.getBugIdFromTicket(ticket)

      if (bugId) {
        const wi = await this.getWorkItemInfo(bugId)

        break
      }
    }
  }
}

const az = new Az(411, "Duy Ba Nguyen")
// az.getWorkItemFromDailyTask()
//   .catch((x) => console.log(x.message))
//   .then((x) => {
//     az.parse()
//   })

az.mappingPullRequest().catch(console.log)
