import WITClient = require("TFS/WorkItemTracking/RestClient");
import Models = require("TFS/WorkItemTracking/Contracts");
import moment = require("moment");
import Q = require("q");

const extensionContext = VSS.getExtensionContext();
const dataService = VSS.getService(VSS.ServiceIds.ExtensionData);
const vssContext = VSS.getWebContext();
const client = WITClient.getClient();

const fields: Models.WorkItemField[] = [];

interface IQuery {
  id: string;
  isPublic: boolean;
  name: string;
  path: string;
  wiql: string;
}

interface IActionContext {
  id?: number;
  workItemId?: number;
  query?: IQuery;
  queryText?: string;
  ids?: number[];
  workItemIds?: number[];
  columns?: string[];
}

const dummy = [
  { name: "Assigned To", referenceName: "System.AssignedTo" },
  { name: "State", referenceName: "System.State" },
  { name: "Created Date", referenceName: "System.CreatedDate" },
  { name: "Description", referenceName: "System.Description" },
  {
    name: "Acceptance Criteria",
    referenceName: "Microsoft.VSTS.Common.AcceptanceCriteria"
  },
  { name: "History", referenceName: "System.History" }
];

// Utilities
declare global {
  interface String {
    sanitize(): string;
    htmlize(): string;
  }
}

const localeTime = "L LT";

String.prototype.sanitize = function (this: string) {
  return this.replace(/\s/g, "-").replace(/[^a-z0-9\-]/gi, "");
};

String.prototype.htmlize = function (this: string) {
  return this.replace(/<\/*(step|param|desc|comp)(.*?)>/g, "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, `"`)
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
};

const printWorkItems = {
  getMenuItems: (context: any) => {
    let menuItemText = "Print";
    if (context.workItemIds && context.workItemIds.length > 1) {
      menuItemText = "Print Selection";
    }

    return [
      {
        action: (actionContext: IActionContext) => {
          const wids = actionContext.workItemIds ||
            actionContext.ids || [actionContext.workItemId || actionContext.id];

          return getWorkItems(wids)
            .then(workItems => prepare(workItems))
            .then(pages => {
              return Q.all(pages);
            })
            .then(pages => {
              let page = pages.reduce((p, c) => p + c);
              document.getElementById("workitems").innerHTML += page;
              console.log(page);
              window.focus(); // needed for IE
              let ieprint = document.execCommand("print", false, null);
              if (!ieprint) {
                window.print();
              }
              document.getElementById("workitems").innerHTML = "";
            });
        },
        icon: "static/img/print14.png",
        text: menuItemText,
        title: menuItemText
      } as IContributedMenuItem
    ];
  }
};

const printQueryToolbar = {
  getMenuItems: (context: any) => {
    return [
      {
        action: (actionContext: IActionContext) => {
          return client
            .queryByWiql(
            { query: actionContext.query.wiql },
            vssContext.project.name,
            vssContext.team.name
            )
            .then(result => {
              if (result.workItemRelations) {
                return result.workItemRelations.map(wi => wi.target.id);
              } else {
                return result.workItems.map(wi => wi.id);
              }
            })
            .then(wids => {
              return getWorkItems(wids)
                .then(workItems => prepare(workItems))
                .then(pages => {
                  return Q.all(pages);
                })
                .then(pages => {
                  let page = pages.reduce((p, c) => p + c);
                  document.getElementById("workitems").innerHTML += page;
                  window.focus(); // needed for IE
                  let ieprint = document.execCommand("print", false, null);
                  if (!ieprint) {
                    window.print();
                  }
                  document.getElementById("workitems").innerHTML = "";
                });
            });
        },
        icon: "static/img/print16.png",
        text: "Print All",
        title: "Print All"
      } as IContributedMenuItem
    ];
  }
};

// Promises
function getWorkItems(wids: number[]): IPromise<Models.WorkItem[]> {
  return client.getWorkItems(
    wids,
    undefined,
    undefined,
    Models.WorkItemExpand.Fields
  );
}

function getWorkItemFields() {
  return client.getFields();
}

function getFields(workItem: Models.WorkItem) {
  return dataService.then((service: IExtensionDataService) => {
    return service
      .getValue(
      `wiprint-${workItem.fields["System.WorkItemType"].sanitize()}`,
      {
        scopeType: "user",
        defaultValue: dummy as Models.WorkItemTypeFieldInstance[]
      }
      )
      .then(
      (data: Models.WorkItemTypeFieldInstance[]) =>
        data.length > 0 ? data : dummy
      );
  });
}

function getHistory(workItem: Models.WorkItem) {
  if (vssContext.account.name === "TEAM FOUNDATION") {
    return client.getHistory(workItem.id).then(comments => {
      return comments.map(comment => {
        return {
          revisedBy: comment.revisedBy,
          revisedDate: comment.revisedDate,
          revision: comment.rev,
          text: comment.value
        } as Models.WorkItemComment;
      });
    });
  }

  return client.getComments(workItem.id).then(comments => comments.comments);
}

function prepare(workItems: Models.WorkItem[]) {
  return workItems.map((item, index) => {
    return Q.all([getFields(item), getHistory(item), getWorkItemFields()])
      .then(results => {
        return results;
      })
      .spread(
      (
        fields: Models.WorkItemTypeFieldInstance[],
        history: Models.WorkItemComment[],
        allFields: Models.WorkItemField[]
      ) => {
        let insertText = "";

        if (index % 6 === 0) {
          insertText += `<div class="row mb-4 row-height">`;
        }
        if (index % 6 === 3) {
          insertText += `<div class="row row-height">`;
        }
        let desc = item.fields["System.Description"] || "";
        insertText += `<div class="col-sm-4">
                                <div class="border">
                                    <div class="pl-3 pr-3 pt-3 clearfix">
                                        <div class="float-left">
                                            <div>
                                                <strong>${item.fields["System.Id"] || ""}</strong>
                                            </div>
                                        </div>
                                        <div class="float-right">
                                            <div>
                                                <em>${item.fields["System.IterationPath"] || ""}</em>
                                            </div>
                                        </div>
                                    </div>
                                    <hr />
                                    <h5 class="pl-2 pr-2 text-center">${item.fields["System.Title"] || ""}</h4>
                                    <div class="pl-2 pr-2 text-justify row-desc-height" style="font-size: 10px;">${desc}</div>
                                    <hr />
                                    <div class="pl-3 pr-3 pb-3 clearfix">
                                        <div class="float-left">
                                            <div> Point: ${item.fields["Microsoft.VSTS.Scheduling.Effort"] || ""} </div>
                                        </div>
                                        <div class="float-right">
                                            <div> Priority: ${item.fields["Microsoft.VSTS.Common.Priority"] || ""} </div>
                                        </div>
                                    </div>
                                </div>
                              </div>`;
        if (index % 6 === 2 || (index === workItems.length - 1 && index % 6 < 2)) {
          insertText += `</div>`;
        }
        if (index % 6 === 5 || (index === workItems.length - 1 && index % 6 >= 3 && index % 6 < 5)) {
          insertText += `</div><div class="page-break"></div>`;
        }
        return insertText;
      });
  });
}
// VSTS/2017
VSS.register(
  `${extensionContext.publisherId}.${extensionContext.extensionId}.print-work-item`,
  printWorkItems
);
VSS.register(
  `${extensionContext.publisherId}.${extensionContext.extensionId}.print-query-toolbar`,
  printQueryToolbar
);
VSS.register(
  `${extensionContext.publisherId}.${extensionContext.extensionId}.print-query-menu`,
  printQueryToolbar
);

// 2015
VSS.register(`print-work-item`, printWorkItems);
VSS.register(`print-query-menu`, printQueryToolbar);
