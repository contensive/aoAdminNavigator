using System;
using Contensive.BaseClasses;

namespace Contensive.AdminNavigator {
    public sealed class NodeManageAddonsController {
        // 
        // ====================================================================================================
        // 
        public static string getNode(CPBaseClass cp, ApplicationEnvironmentModel env, ref string Return_NavigatorJS, ref string nodeIDString) {
            try {
                nodeIDString = "";
                string result = "";
                string addonManagerVerionGuid = "{4DD876E7-BCC4-4AF0-B32E-59FFAB816478}";
                var addon = Models.Db.DbBaseModel.create<Models.Db.AddonModel>(cp, addonManagerVerionGuid);
                if (addon is not null) {
                    result += NodeController.getNode(cp, env, 0, 0, 0, 0, 0, "?addonguid=" + addonManagerVerionGuid, addon, 0, "Add-on Manager", env.emptyNodeList, 0, common.NavIconTypeAddon, "Add-on Manager", env.autoManageAddons, common.NodeTypeEnum.NodeTypeAddon, false, false, env.openNodeList, nodeIDString, ref Return_NavigatorJS, "");
                }
                // 
                // List Collections
                string fieldList = "Name,0 as id,ccaddoncollections.id as collectionid,0 as AddonID,0 as NewWindow,0 as ContentID,'' as LinkPage," + common.NavIconTypeFolder + " as NavIconType,Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as contentcontrolid,blockNavigatorNode,system";
                string criteria = "((system=0)or(system is null))";
                // 
                // -- reduce confusion
                criteria += "and((blockNavigatorNode=0)or(blockNavigatorNode is null))";
                // If Not env.isDeveloper Then
                // criteria &= "and((blockNavigatorNode=0)or(blockNavigatorNode is null))"
                // End If
                using (var cs3 = cp.CSNew()) {
                    var NodeType = common.NodeTypeEnum.NodeTypeCollection;
                    bool BlockSubNodes = false;
                    if (cs3.Open("Add-on Collections", criteria, "name", true, fieldList, 9999, 1)) {
                        do {
                            string Name = cs3.GetText("name").Trim();
                            string NavIconTitle = Name;
                            int CollectionID = cs3.GetInteger("collectionid");
                            nodeIDString = common.NodeIDManageAddonsCollectionPrefix + "." + CollectionID;
                            string NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(NavIconTitle);
                            string linkSuffixList = "";
                            // 
                            // -- reduce confusion
                            // If env.isDeveloper Then
                            // linkSuffixList &= "<a href=""" & env.addonEditCollectionUrlPrefix & CollectionID & """>edit</a>"
                            // If cs3.GetBoolean("system") Then
                            // linkSuffixList &= ",sys"
                            // End If
                            // If cs3.GetBoolean("blockNavigatorNode") Then
                            // linkSuffixList &= ",dev"
                            // End If
                            // linkSuffixList = "&nbsp;(" & linkSuffixList & ")"
                            // End If
                            result += NodeController.getNode(cp, env, CollectionID, 0, 0, 0, 0, "", null, 0, Name, env.emptyNodeList, 0, common.NavIconTypeAddon, NavIconTitleHtmlEncoded, env.autoManageAddons, common.NodeTypeEnum.NodeTypeCollection, false, false, env.openNodeList, nodeIDString, ref Return_NavigatorJS, linkSuffixList);
                            Return_NavigatorJS += Return_NavigatorJS;
                            cs3.GoNext();
                        }
                        while (cs3.OK());
                    }
                    cs3.Close();
                }
                // 
                // Advanced folder to contain edit links to create addons and collections
                // 
                nodeIDString = common.NodeIDManageAddonsAdvanced;
                result += NodeController.getNode(cp, env, 0, 0, 0, 0, 0, "", null, 0, "Advanced", env.emptyNodeList, 0, common.NavIconTypeFolder, "Add-ons With No Collection", env.autoManageAddons, common.NodeTypeEnum.NodeTypeEntry, false, false, env.openNodeList, nodeIDString, ref Return_NavigatorJS, "");
                Return_NavigatorJS += Return_NavigatorJS;
                return result;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}