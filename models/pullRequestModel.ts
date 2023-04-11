interface WorkItem {
    id: number;
    wiTitle: string;
    wiUrl: string;
    podLead: string;
    fields?: {
        [key: string]: any;
    };
}

interface PullRequest {
    title: string;
    pullRequestId: number;
    pullRequestUrl: string;
    workItems: WorkItem[];
}

interface TimeLogWorkItem {
    workitem: WorkItem;
    date: string;
    type: string;
}