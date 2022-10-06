import { BaseComponentContext } from "@microsoft/sp-component-base";
import { SPFI } from "@pnp/sp";
import { GraphFI, graphfi, SPFx as SPFxGR } from "@pnp/graph";
import "@pnp/graph/";
import "@pnp/graph/groups";
import "@pnp/graph/onedrive";
import "@pnp/graph/sites";
import { Site } from "@pnp/graph/sites";
import "@pnp/graph/sites/types";
import "@pnp/graph/teams";
import "@pnp/graph/users";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/folders";
import "@pnp/sp/items";
import { IItem } from "@pnp/sp/items";
import "@pnp/sp/lists";
import { IList } from "@pnp/sp/lists";
import "@pnp/sp/security";
import { TimelinePipe } from "@pnp/core";
import { Queryable } from "@pnp/queryable";
import { HttpsProxyAgent } from "https-proxy-agent";
import "@pnp/sp/security/web";
import { ISiteUserProps } from "@pnp/sp/site-users/types";

import "@pnp/sp/site-users/web";
import "@pnp/sp/views";
import { IViewInfo } from "@pnp/sp/views";
import "@pnp/sp/webs";
import {
  IBasePermissions,
  IRoleDefinitionInfo,
  PermissionKind,
  IRoleAssignmentInfo,
  IRoleDefinition,
  IRoleDefinitions,
} from "@pnp/sp/security";
import {
  IChannel,
  ITab,
  ITeam,
  ITeams,
  Team,
  Teams,
  Channel,
} from "@pnp/graph/teams";
import { find } from "lodash";

export async function getItemsInListWithUniqueRoleAssignments(
  sp: SPFI,
  listId: string
) {
  debugger;
  const r = await sp.web.lists
    .getById(listId)
    .getItemsByCAMLQuery({
      ViewXml: `<View><Query><Where><IsNotNull><FieldRef Name='SharedWithDetails' /></IsNotNull></Where></Query></View>`,
    })
    .then((e) => {
      debugger;
    })
    .catch((e) => {
      debugger;
    });
  return r;
}

export function CacheBust(): TimelinePipe<Queryable> {
  // see https://pnp.github.io/pnpjs/core/behavior-recipes/#add-querystring-to-bypass-request-caching
  return (instance: Queryable) => {
    instance.on.pre(async (url, init, result) => {
      url += url.indexOf("?") > -1 ? "&" : "?";
      url += "nonce=" + encodeURIComponent(new Date().toISOString());
      return [url, init, result];
    });
    return instance;
  };
}
export async function grantTeamMembersAcessToItem(
  teamId: string,
  roleDefinitionId: number,
  sp: SPFI,
  roleDefinitionInfos: IRoleDefinitionInfo[],
  listId: string,
  itemId: number
) {
  const siteUser = await ensureTeamsUser(sp, teamId);
  const roledefinition = find(
    roleDefinitionInfos,
    (x) => x.Id === roleDefinitionId
  );
  const selectedItem = await sp.web.lists.getById(listId).items.getById(itemId);

  const teamPermissions = await selectedItem.getUserEffectivePermissions(
    siteUser.LoginName
  );
  debugger;
  const teamHasPermissions = await sp.web.hasPermissions(
    teamPermissions,
    roledefinition.RoleTypeKind
  );

 // if (!teamHasPermissions) {
    await selectedItem.breakRoleInheritance(true, false);
    await selectedItem.roleAssignments.add(siteUser.Id, roleDefinitionId);
 // }
}

export async function grantTeamMembersAcessToFolder(
  teamId: string,
  roleDefinitionId: number,
  sp: SPFI,
  roleDefinitionInfos: IRoleDefinitionInfo[],
  folderServerRelativePath: string
) {
  //const sp = spfi().using(SPFx(props.context));
  const siteUser = await ensureTeamsUser(sp, teamId);
  const roledefinition = find(
    roleDefinitionInfos,
    (x) => x.Id === roleDefinitionId
  );
  const folder = await sp.web
    .getFolderByServerRelativePath(folderServerRelativePath)
    .getItem();
  const teamPermissions = await folder.getUserEffectivePermissions(
    siteUser.LoginName
  );
  const teamHasPermissions = await sp.web.hasPermissions(
    teamPermissions,
    roledefinition.RoleTypeKind
  );

  // if (!teamHasPermissions) {
    await folder.breakRoleInheritance(true, false);
    await folder.roleAssignments.add(siteUser.Id, roleDefinitionId);
  //}
}

export async function grantTeamMembersAcessToLibrary(
  teamId: string,
  roleDefinitionId: number,
  sp: SPFI,
  roleDefinitionInfos: IRoleDefinitionInfo[],
  listId: string
) {
  debugger;
  const siteUser = await ensureTeamsUser(sp, teamId);
  const roledefinition = find(
    roleDefinitionInfos,
    (x) => x.Id === roleDefinitionId
  );
  const teamPermissions = await sp.web.lists
    .getById(listId)
    .getUserEffectivePermissions(siteUser.LoginName);

  const hasem = hasPermissions(teamPermissions, roledefinition.BasePermissions);
  debugger;
 // if (!hasem) {
    await await sp.web.lists.getById(listId).breakRoleInheritance(true, false);
    await sp.web.lists
      .getById(listId)
      .roleAssignments.add(siteUser.Id, roleDefinitionId);
 // }
}
export function hasPermissions(
  existingPermissions: any,
  requiredPermissions: any
) {
  // hig and low are strings , not numbers!
  // see : https://www.w3schools.com/js/js_bitwise.asp
  // and  : https://www.darraghoriordan.com/2019/07/29/bitwise-mask-typescript/
  const eHi = parseInt(existingPermissions["High"], 10);
  const eLo = parseInt(existingPermissions["Low"], 10);
  const rHi = parseInt(requiredPermissions["High"], 10);
  const rLo = parseInt(requiredPermissions["Low"], 10);

  const hasPerms = (rHi & eHi) === rHi && (rLo & eLo) === rLo;

  return hasPerms;
}
export async function ensureTeamsUser(
  sp: SPFI,
  teamId: string
): Promise<ISiteUserProps> {
  // const group = await graph.groups.getById(teamId)();
  const user = await sp.web.ensureUser(getTeamLoginNameFromTeamId(teamId));
  console.dir(user);
  return user.data;
}
export function getTeamLoginNameFromTeamId(teamId: string): string {
  return `c:0o.c|federateddirectoryclaimprovider|${teamId}`;
}
export async function getRoleDefs(sp): Promise<IRoleDefinitionInfo[]> {
  // get the role definitions for the current web -- now full condtrol or designer
  return await sp.web.roleDefinitions
    .filter(
      "BasePermissions ne null and Hidden eq false and RoleTypeKind ne 4 and RoleTypeKind ne 5 and RoleTypeKind ne 6"
    ) // 4 is designer, 5 is admin, 6 is editor
    .orderBy("Order", true)()
    .then((roleDefs: IRoleDefinitionInfo[]) => {
      return roleDefs;
    })
    .catch((err) => {
      alert(err);
      console.log(err);
      return [];
    });
}
export async function removeRoleAssignmentFromItem(parms: {
  sp: SPFI;
  listId: string;
  ra: IRoleAssignmentInfo;
  roleDefId: number;
  itemId: number;
}) {
  debugger;
  await parms.sp.web.lists
    .getById(parms.listId)
    .items.getById(parms.itemId)
    .roleAssignments.remove(parms.ra.PrincipalId, parms.roleDefId);
}
export async function removeRoleAssignmentFromList(parms: {
  sp: SPFI;
  listId: string;
  ra: IRoleAssignmentInfo;
  roleDefId: number;
}) {
  debugger;
  await parms.sp.web.lists
    .getById(parms.listId)
    .roleAssignments.remove(parms.ra.PrincipalId, parms.roleDefId);
}
export async function getJoinedTeams(parms: {
  graph: GraphFI;
}): Promise<any[]> {
  debugger;
  return parms.graph.me.joinedTeams();
}
export function RemoveLimitedAccess(roleAssignments: Array<any>): Array<any> {
  debugger;
  //   {
  //     "High": "48",
  //     "Low": "134287360"
  // }
  var temp = roleAssignments
    .map((ra) => {
      return {
        ...ra,
        RoleDefinitionBindings: ra.RoleDefinitionBindings.filter((rdb) => {
          return !(
            rdb.BasePermissions.High === "48" &&
            rdb.BasePermissions.Low === "134287360"
          );
        }),
      };
    })
    .filter((ra) => {
      return ra.RoleDefinitionBindings.length !== 0;
    });

  return temp;
}
export async function getTeamTabs(parms: {
  graph: GraphFI;
  teamId: string;
  contentUrl: string;
}): Promise<any[]> {
  var channels = await parms.graph.teams.getById(parms.teamId).channels();
  var promsies = new Array<Promise<any>>();
  const teamTabs = [];
  for (var channel of channels) {
    promsies.push(
      parms.graph.teams
        .getById(parms.teamId)
        .channels.getById(channel.id)
        .tabs()
        .then((tabs) => {
          debugger;
          for (const tab of tabs) {
            if (tab.configuration.contentUrl === parms.contentUrl) {
              teamTabs.push({
                ...tab,
                channelId: channel.id,
                channelName: channel.displayName,
              }); // add the channel id so we can delete it
            }
          }
        })
    );
  }
  return Promise.all(promsies)
    .then(() => {
      return teamTabs;
    })
    .catch((e) => {
      alert(e);
      return [];
    });
}
export function getTeamIdFromLoginName(loginName: string): string {
  return loginName.split("|")[2];
}

export function removeTeamsTab(
  graph: GraphFI,
  teamId: string,
  channelId: string,
  tabId: string
) {
  graph.teams
    .getById(teamId)
    .channels.getById(channelId)
    .tabs.getById(tabId)
    .delete();
}
//   export async function getTabs(parms: {
//     graph:GraphFI;
//     teams:any[];

//   }):Promise<IChannel> {
//     debugger;
//     const p = new Array<Promise>;
//     for (var team of parms.teams){
//        parms.graph.teams.getById(team.id)
//     }
//     return await parms.graph.me.joinedTeams()

//   }
