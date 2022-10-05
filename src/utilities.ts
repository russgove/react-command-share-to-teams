import { BaseComponentContext } from "@microsoft/sp-component-base";
import { SPFI } from "@pnp/sp";
import { graphfi, SPFx as SPFxGR } from "@pnp/graph";
import "@pnp/graph/";
import "@pnp/graph/groups";
import "@pnp/graph/onedrive";
import "@pnp/graph/sites";
import { Site } from "@pnp/graph/sites";
import "@pnp/graph/sites/types";
import "@pnp/graph/teams";
import "@pnp/graph/users";
import {  spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/folders";
import "@pnp/sp/items";
import { IItem } from "@pnp/sp/items";
import "@pnp/sp/lists";
import { IList } from "@pnp/sp/lists";
import "@pnp/sp/security";

import "@pnp/sp/security/web";
import { ISiteUserProps } from "@pnp/sp/site-users/types";
import "@pnp/sp/site-users/web";
import "@pnp/sp/views";
import { IViewInfo } from "@pnp/sp/views";
import "@pnp/sp/webs";
import { IBasePermissions, IRoleDefinitionInfo, PermissionKind ,IRoleAssignmentInfo,IRoleDefinition,IRoleDefinitions} from "@pnp/sp/security";

async function getRoleDefs(sp): Promise< IRoleDefinitionInfo[]> {
    // get the role definitions for the current web -- now full condtrol or designer
   return await sp.web.roleDefinitions
      .filter("BasePermissions ne null and Hidden eq false and RoleTypeKind ne 4 and RoleTypeKind ne 5 and RoleTypeKind ne 6")  // 4 is designer, 5 is admin, 6 is editor
      .orderBy("Order", true)
      ().then((roleDefs: IRoleDefinitionInfo[]) => {

        return roleDefs;
      }).catch(err => {
alert(err);
        console.log(err);
        return [];
      });
  }
export async function removeRoleAssignmentFromList( parms:{sp:SPFI,listId:string,ra:IRoleAssignmentInfo,roleDefId:number}) {
    debugger;
   await parms.sp.web.lists
        .getById(parms.listId)
        .roleAssignments.remove(parms.ra.PrincipalId, parms.roleDefId);


}
