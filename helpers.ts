import { existsSync, readFileSync, writeFileSync } from "fs"
import _ from "lodash"
import { join } from "path"
class Helpers {
  static writeData = (path: String, data: any) => {
    writeFileSync(`data/${path}`, JSON.stringify(data, null, 2))
  }

  static sort = (taskSummaries: any[]) => {
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

  static readFileJson(fileName: string) {
    const filePath = join(__dirname, "../data", fileName)
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

  static getBugIdFromTicket = (ticket: string) => {
    const ticketArr = _.split(ticket, " ")
    const bugIndex = _.indexOf(ticketArr, "Bug")
    if (bugIndex > -1) {
      return +`${_.replace(ticketArr[bugIndex + 1], ":", "")}`
    }

    return 0
  }

  static convertExeclDate = (excelSerialDate: any) => {
    const unixTimestamp = (excelSerialDate - 25569) * 86400
    const dateObj = new Date(unixTimestamp * 1000)

    const month = dateObj.getMonth() + 1
    const day = dateObj.getDate()
    const year = dateObj.getFullYear()

    const formattedDate = `${month}/${day}/${year}`
    return formattedDate
  }

  static getQuarterDates(date: Date, quarter: number) {
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

  static getQuarterFromDate(dateStr: string) {
    const date = new Date(dateStr)
    const month = date.getMonth()
    const quarter = Math.floor(month / 3) + 1
    return `Q${quarter}`
  }
}

export default Helpers
