using System;
using System.Collections.Generic;
using Contensive.BaseClasses;
using Contensive.Models.Db;

namespace Contensive.AdminNavigator {
    public sealed class NodeListMixedController {
        // 
        // ====================================================================================================
        /// <summary>
        /// list mixed nodes (settings/reports/tools)   
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="env"></param>
        /// <param name="xEmptyNodeList"></param>
        /// <param name="TopParentNode"></param>
        /// <param name="AutoManageAddons"></param>
        /// <param name="NodeType"></param>
        /// <param name="xOpenNodeList"></param>
        /// <param name="AddonNavTypeID"></param>
        /// <param name="MenuParentNodeID"></param>
        /// <param name="AdminNavIconTypeSetting"></param>
        /// <param name="Return_DraggableJS"></param>
        /// <returns></returns>
        internal static string getNodeListMixed(CPBaseClass cp, ApplicationEnvironmentModel env, string xEmptyNodeList, string TopParentNode, bool AutoManageAddons, common.NodeTypeEnum NodeType, string xOpenNodeList, int AddonNavTypeID, int MenuParentNodeID, int AdminNavIconTypeSetting, ref string Return_DraggableJS) {
            string returnNav = "";
            try {
                string lastName;
                var sortedNodes = new List<KeyValuePair<string, common.SortNodeType>>();
                var SortPtr = default(int);
                string NodeIDString;
                string Criteria;
                var BlockSubNodes = default(bool);
                // 
                // list mixed nodes (settings/reports/tools), includes menu nodes and addons with type='setting' sorted in
                // 
                if ((env.emptyNodeList + ",").IndexOf("," + TopParentNode + ",", StringComparison.Ordinal) >= 0) {
                    // 
                } else {
                    // 
                    // Add addons to node list
                    // 
                    NodeIDString = "";
                    Criteria = "(navtypeid=" + AddonNavTypeID + ")";
                    if (AddonNavTypeID == 2 | AddonNavTypeID == 3 | AddonNavTypeID == 4) {
                        // 
                        // if setting, report or tool, "admin" is not needed
                        // 
                        // 
                        // for Manage Addons node, admin is needed for non-developer
                        // 
                    } else if (!cp.User.IsDeveloper) {
                        Criteria += "and(admin<>0)";
                    }
                    using (var cs10 = cp.CSNew()) {
                        if (cs10.Open("add-ons", Criteria, "name")) {
                            do {
                                string nodeName = cs10.GetText("name").Trim();
                                var sortNode = new common.SortNodeType() {
                                    Name = cs10.GetText("name").Trim(),
                                    addonid = cs10.GetInteger("ID"),
                                    NavIconTitle = cs10.GetText("name").Trim(),
                                    ContentControlID = cs10.GetInteger("ContentControlID"),
                                    NavIconType = AdminNavIconTypeSetting
                                };
                                sortedNodes.Add(new KeyValuePair<string, common.SortNodeType>(nodeName, sortNode));
                                SortPtr += 1;
                                cs10.GoNext();
                            }
                            while (cs10.OK());
                        }
                        cs10.Close();
                    }
                    // 
                    // Add real navigator nodes to node list
                    // 
                    using (var cs11 = cp.CSNew()) {
                        if (cs11.OpenSQL(MenuSqlController.getMenuSQL(cp, "parentid=" + MenuParentNodeID, ""))) {
                            do {
                                string nodeName = cs11.GetText("name").Trim();
                                string navIconTitle = cs11.GetText("NavIconTitle");
                                var sortNode = new common.SortNodeType() {
                                    Name = cs11.GetText("name").Trim(),
                                    NavigatorID = cs11.GetInteger("ID"),
                                    CollectionID = cs11.GetInteger("CollectionID"),
                                    NewWindow = cs11.GetBoolean("newwindow"),
                                    contentID = cs11.GetInteger("ContentID"),
                                    Link = cs11.GetText("LinkPage").Trim(),
                                    addonid = cs11.GetInteger("AddonID"),
                                    NavIconType = cs11.GetInteger("NavIconType"),
                                    NavIconTitle = string.IsNullOrEmpty(navIconTitle) ? nodeName : navIconTitle,
                                    HelpAddonID = cs11.GetInteger("HelpAddonID"),
                                    helpCollectionID = cs11.GetInteger("HelpCollectionID"),
                                    ContentControlID = cs11.GetInteger("ContentControlID"),
                                    NodeIDString = nodeName.ToLowerInvariant() == "all content" ? common.NodeIDAllContentList : common.NodeIDAllContentList
                                };
                                if (nodeName.ToLowerInvariant() == "all content") {
                                    sortNode.NodeIDString = common.NodeIDAllContentList;
                                } else {
                                    sortNode.NodeIDString = sortNode.NavigatorID.ToString();
                                    sortNode.NavIconTitle = nodeName;
                                }
                                sortedNodes.Add(new KeyValuePair<string, common.SortNodeType>(nodeName, sortNode));
                                SortPtr += 1;
                                cs11.GoNext();
                            }
                            while (cs11.OK());
                        }
                        cs11.Close();
                    }
                    // 
                    sortedNodes.Sort((valueA, valueB) => valueA.Key.CompareTo(valueB.Key));
                    // 
                    if (sortedNodes.Count == 0) {
                        // 
                        // Empty list, add to env.EmptyNodeList
                        // 
                        env.emptyNodeList += "," + TopParentNode;
                    } else {
                        lastName = "";
                        string contentNodeList = "";
                        foreach (var kvp in sortedNodes) {
                            string nodeName = kvp.Key.ToLowerInvariant();
                            if ((nodeName ?? "") != (lastName ?? "")) {
                                var sortedNode = kvp.Value;
                                string NodeDraggableJS = "";
                                if (!sortedNode.contentID.Equals(0)) {
                                    // 
                                    // -- content nodes last
                                    contentNodeList += NodeController.getNode(cp, env, sortedNode.CollectionID, sortedNode.ContentControlID, sortedNode.helpCollectionID, sortedNode.HelpAddonID, sortedNode.contentID, sortedNode.Link, null, 0, sortedNode.Name, env.emptyNodeList, sortedNode.NavigatorID, sortedNode.NavIconType, cp.Utils.EncodeHTML(sortedNode.NavIconTitle), AutoManageAddons, NodeType, sortedNode.NewWindow, BlockSubNodes, env.openNodeList, sortedNode.NodeIDString, ref NodeDraggableJS, "");
                                } else if (sortedNode.addonid.Equals(0)) {
                                    // 
                                    // -- hard-coded nodes
                                    returnNav += NodeController.getNode(cp, env, sortedNode.CollectionID, sortedNode.ContentControlID, sortedNode.helpCollectionID, sortedNode.HelpAddonID, sortedNode.contentID, sortedNode.Link, null, 0, sortedNode.Name, env.emptyNodeList, sortedNode.NavigatorID, sortedNode.NavIconType, cp.Utils.EncodeHTML(sortedNode.NavIconTitle), AutoManageAddons, NodeType, sortedNode.NewWindow, BlockSubNodes, env.openNodeList, sortedNode.NodeIDString, ref NodeDraggableJS, "");
                                } else {
                                    // 
                                    // -- addon nodes next
                                    var addon = DbBaseModel.create<AddonModel>(cp, sortedNode.addonid);
                                    returnNav += NodeController.getNode(cp, env, sortedNode.CollectionID, sortedNode.ContentControlID, sortedNode.helpCollectionID, sortedNode.HelpAddonID, sortedNode.contentID, sortedNode.Link, addon, 0, sortedNode.Name, env.emptyNodeList, sortedNode.NavigatorID, sortedNode.NavIconType, cp.Utils.EncodeHTML(sortedNode.NavIconTitle), AutoManageAddons, NodeType, sortedNode.NewWindow, BlockSubNodes, env.openNodeList, sortedNode.NodeIDString, ref NodeDraggableJS, "");
                                }
                                Return_DraggableJS += NodeDraggableJS;
                                lastName = nodeName;
                            }
                        }
                        returnNav += contentNodeList;
                    }
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
            return returnNav;
        }
    }
}