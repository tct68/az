import VSSInterfaces from "azure-devops-node-api/interfaces/common/VSSInterfaces"
interface WorkItem {
  id: number
  wiTitle: string
  wiUrl: string
  podLead: string
  channelName: string
  fields?: {
    [key: string]: any
  }
}

interface PullRequest {
  title: string
  pullRequestId: number
  pullRequestUrl: string
  workItems: WorkItem[]
}

interface TimeLogWorkItem {
  workitem: WorkItem
  date: string
  type: string
  quarter: string
}

interface PullRequestWorkItemRefs {
  pullRequestId: number
  workItems: VSSInterfaces.ResourceRef[]
}

export { PullRequest, WorkItem, TimeLogWorkItem, PullRequestWorkItemRefs }
