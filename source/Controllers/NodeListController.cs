using System;
using System.Collections.Generic;
using Contensive.BaseClasses;
using Contensive.Models.Db;

namespace Contensive.AdminNavigator {
    public sealed class NodeListController {
        // 
        // ====================================================================================================
        /// <summary>
        /// get Node list?
        /// </summary>
        internal static string getNodeList(CPBaseClass cp, ApplicationEnvironmentModel env, string ParentNode, ref string Return_NavigatorJS) {
            string returnNav = "";
            int hint = 0;
            try {
                string TopParentNode = ParentNode;
                string[] parentNodeStack;
                if (string.IsNullOrEmpty(TopParentNode)) {
                    // 
                    // bad call
                    // 
                    parentNodeStack = new string[1];
                    parentNodeStack[0] = "";
                } else {
                    // 
                    // load ParentNodes with argument
                    // 
                    parentNodeStack = TopParentNode.Split('.');
                }
                var NodeNavigatorJS = default(string);
                string ATag;
                string FieldList;
                string IconNoSubNodes;
                int NavigatorID;
                int CollectionID;
                string Name;
                int ContentID;
                string Link;
                string Criteria;
                var BlockSubNodes = default(bool);
                var NodeType = default(common.NodeTypeEnum);
                int NavIconType;
                string NavIconTitle = "";
                string NavIconTitleHtmlEncoded;
                int ContentControlID;
                var csChildList = cp.CSNew();

                const bool AutoManageAddons = true;
                string nodeIDString = "";
                switch (parentNodeStack[0] ?? "") {
                    // 
                    // Open CS so:
                    // Name = the caption that is displayed for the entry
                    // ID (NavigatorID) = the NavigatorEntry the record represents
                    // if the node has no navigation entry, NavigatorID=0 if there are no child nodes
                    // this number is used for the open/close javascript, as well as stored to remember open/close state
                    // during future hits, as the menu is built, this is checked in the open/close list for a match
                    // NavigatorID=0 will only work if the node has not child nodes
                    // AddonID = the ID of the addon that should be run if this entry is selected, 0 otherwise
                    // CollectionID, if this is the manage add-ons section, this is the collection node
                    // NewWindow = 0 or 1, if the link opens in new window
                    // ContentID = the id of the content to be opened in list mode if the entry is clicked
                    // Link = URL to link the menu entry
                    // NodeIDString = unique string that represents this node
                    // Navigator Entries - 'n'+EntryID
                    // Collections = 'c'+CollectionID
                    // Add-ons = 'a'+AddonID
                    // CDefs = 'd'+ContentID
                    // 
                    case common.NodeIDManageAddons: {
                            // 
                            // Special Case: clicked on Manage Add-ons ("manageaddons")
                            // Link to Add-on Manager
                            returnNav += NodeManageAddonsController.getNode(cp, env, ref Return_NavigatorJS, ref nodeIDString);
                            break;
                        }
                    case common.NodeIDManageAddonsCollectionPrefix: {
                            // 
                            // Special Case: clicked on Manage Add-ons.collection
                            // ParentNode(1) is the id of the collection they clicked on
                            // List all add-ons
                            // List all CDef
                            // Add Collection Help
                            // Add Layouts associated with collection
                            // 
                            string nodeHtml = "";
                            string cacheName;
                            CollectionID = 0;
                            if (parentNodeStack.Length - 1 > 0) {
                                CollectionID = cp.Utils.EncodeInteger(parentNodeStack[1]);
                            }
                            cacheName = "addonNav." + common.NodeIDManageAddonsCollectionPrefix + "." + CollectionID + "." + cp.User.Id.ToString();
                            nodeHtml = cp.Cache.GetText(cacheName);
                            if (string.IsNullOrEmpty(nodeHtml)) {
                                // 
                                // Help Icon
                                // 
                                Name = "Help";
                                nodeIDString = "";
                                NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(NavIconTitle);
                                nodeHtml += NodeController.getNode(cp, env, 0, 0, CollectionID, 0, 0, "", null, 0, Name, env.emptyNodeList, 0, common.NavIconTypeHelp, NavIconTitleHtmlEncoded, AutoManageAddons, common.NodeTypeEnum.NodeTypeEntry, false, true, env.openNodeList, nodeIDString, ref NodeNavigatorJS, "");
                                Return_NavigatorJS += NodeNavigatorJS;
                                // 
                                // List out add-ons in this collection
                                // 
                                nodeIDString = "";
                                FieldList = "*";
                                // 
                                // -- block all addons not marked to be on the admin-nav (admin boolean)
                                // -- big change, developers do not need this long list.
                                Criteria = "(collectionid=" + CollectionID + ")and(admin>0)";
                                string NameSuffix;
                                foreach (AddonModel addon in DbBaseModel.createList<AddonModel>(cp, Criteria, "name")) {
                                    Name = (addon.name ?? "").Trim();
                                    NameSuffix = "";
                                    string linkSuffixList = "";
                                    if (env.isDeveloper) {
                                        linkSuffixList += "<a href=\"" + env.addonEditAddonUrlPrefix + addon.id + "\">edit</a>";
                                        if (!addon.admin) {
                                            linkSuffixList += ",dev";
                                        }
                                        linkSuffixList = "&nbsp;(" + linkSuffixList + ")";
                                    }
                                    int addonidx = addon.id;
                                    NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(Name);
                                    ContentControlID = addon.contentControlId;
                                    switch (addon.navTypeId) {
                                        case 2: {
                                                NavIconType = common.NavIconTypeReport;
                                                break;
                                            }
                                        case 3: {
                                                NavIconType = common.NavIconTypeSetting;
                                                break;
                                            }
                                        case 4: {
                                                NavIconType = common.NavIconTypeTool;
                                                break;
                                            }

                                        default: {
                                                NavIconType = common.NavIconTypeAddon;
                                                break;
                                            }
                                    }
                                    nodeHtml += NodeController.getNode(cp, env, 0, ContentControlID, 0, 0, 0, "", addon, 0, Name, env.emptyNodeList, 0, NavIconType, NavIconTitleHtmlEncoded, AutoManageAddons, common.NodeTypeEnum.NodeTypeAddon, false, false, env.openNodeList, nodeIDString, ref NodeNavigatorJS, linkSuffixList);
                                    Return_NavigatorJS += NodeNavigatorJS;

                                }
                                // 
                                // List out cdefs connected to this collection
                                // 
                                nodeIDString = "";
                                Criteria = "(collectionid=" + CollectionID + ")";
                                if (env.isDeveloper) {
                                } else if (cp.User.IsAdmin) {
                                    Criteria += "and(developeronly=0)";
                                } else {
                                    Criteria += "and(developeronly=0)and(adminonly=0)";
                                }
                                int LastContentID;
                                bool DupsFound;
                                LastContentID = -1;
                                DupsFound = false;
                                string SQL = "select c.id,c.name,c.contentcontrolid,c.developeronly,c.adminonly from ccContent c left join ccAddonCollectionCDefRules r on r.contentid=c.id where " + Criteria + " order by c.name";
                                var cs7 = cp.CSNew();
                                if (cs7.OpenSQL(SQL)) {
                                    do {
                                        Name = cs7.GetText("name").Trim();
                                        ContentID = cs7.GetInteger("id");
                                        if (ContentID == LastContentID) {
                                            DupsFound = true;
                                        } else {
                                            string linkSuffixList = "";
                                            if (env.isDeveloper) {
                                                linkSuffixList = "<a href=\"" + env.contentFieldEditToolPrefix + ContentID + "\">edit</a>";
                                                if (cs7.GetBoolean("developeronly")) {
                                                    linkSuffixList += ",dev";
                                                } else if (cs7.GetBoolean("adminonly")) {
                                                    linkSuffixList += ",adm";
                                                }
                                                if (!string.IsNullOrEmpty(linkSuffixList)) {
                                                    linkSuffixList = "&nbsp;(" + linkSuffixList + ")";
                                                }
                                            }
                                            NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(Name);
                                            ContentControlID = cs7.GetInteger("ContentControlID");
                                            NameSuffix = "";
                                            nodeHtml += NodeController.getNode(cp, env, 0, ContentControlID, 0, 0, ContentID, "", null, 0, Name, env.emptyNodeList, 0, common.NavIconTypeContent, NavIconTitleHtmlEncoded, AutoManageAddons, common.NodeTypeEnum.NodeTypeContent, false, true, env.openNodeList, nodeIDString, ref NodeNavigatorJS, linkSuffixList);
                                            Return_NavigatorJS += NodeNavigatorJS;
                                        }
                                        LastContentID = ContentID;
                                        cs7.GoNext();
                                    }
                                    while (cs7.OK());
                                }
                                cs7.Close();
                                // 
                                // list all data records associated to this collection
                                // 
                                string dataRecordList = "";
                                var dataRecords = new List<string>();
                                string[] dataRecordParts;
                                // Dim dataRecordId As Integer
                                // Dim dataRecordGuid As String
                                string dataRecordName;
                                string dataRecordCdefName;
                                int dataRecordCdefID;

                                if (cs7.Open("add-on collections", "id=" + CollectionID)) {
                                    dataRecordList = cs7.GetText("dataRecordList");
                                }
                                cs7.Close();
                                hint = 1100;
                                if (!string.IsNullOrEmpty(dataRecordList)) {
                                    try {
                                        dataRecords.AddRange(dataRecordList.Split(new[] { "\r\n" }, StringSplitOptions.None));
                                        foreach (string dataRecord in dataRecords) {
                                            dataRecordParts = dataRecord.Split(",".ToCharArray());
                                            dataRecordCdefName = dataRecordParts[0];
                                            dataRecordCdefID = cp.Content.GetID(dataRecordCdefName);
                                            if (dataRecordCdefID == 0) {
                                                // -- invalid content, skip it
                                            } else {
                                                // -- export records from this content
                                                string sqlCriteria = "";
                                                if (string.IsNullOrEmpty(dataRecordCdefName)) {
                                                    // -- no comma, export all records
                                                } else if (dataRecordParts.Length >= 2) {
                                                    hint = 1110;
                                                    // 
                                                    // contentname,(name or guid)
                                                    if (string.IsNullOrEmpty(dataRecordParts[1])) {
                                                        // -- no data given, list all records in export
                                                    } else if (dataRecordParts[1].Substring(0, 1) == "{") {
                                                        sqlCriteria = "ccguid=" + cp.Db.EncodeSQLText(dataRecordParts[1]);
                                                    } else {
                                                        sqlCriteria = "name=" + cp.Db.EncodeSQLText(dataRecordParts[1]);
                                                    }
                                                    hint = 1120;
                                                }
                                                if (cs7.Open(dataRecordCdefName, sqlCriteria)) {
                                                    do {
                                                        dataRecordName = cs7.GetText("name");


                                                        NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML("Edit '" + dataRecordName + "' in '" + dataRecordCdefName + "'");
                                                        IconNoSubNodes = common.IconRecord;
                                                        IconNoSubNodes = IconNoSubNodes.Replace("{title}", NavIconTitleHtmlEncoded);
                                                        Link = "?id=" + cs7.GetInteger("id").ToString() + "&cid=" + dataRecordCdefID.ToString() + "&af=4";
                                                        ATag = "<a href=\"" + Link + "\" title=\"" + NavIconTitleHtmlEncoded + "\">";
                                                        nodeHtml += common.cr + "<div class=\"ccNavLink ccNavLinkEmpty\">" + ATag + IconNoSubNodes + "</a>&nbsp;" + ATag + dataRecordCdefName + ":" + dataRecordName + "</a></div>";
                                                        // nodeHtml &= GetNode(cp, env, 0, 0, 0, 0, 0, "/admin?af=4&cid=" & dataRecordCdefID.ToString & "&id=" & cs7.GetInteger("id"), 0, 0, dataRecordCdefName & ":" & cs7.GetText("name"), env.EmptyNodeList, 0, NavIconTypeContent, "Data", AutoManageAddons, NodeTypeEnum.NodeTypeContent, False, True, env.OpenNodeList, NodeIDString, NodeNavigatorJS, "")
                                                        cs7.GoNext();
                                                    }
                                                    while (cs7.OK());

                                                }
                                                cs7.Close();
                                            }
                                        }
                                        hint = 1100;
                                        // 
                                        if (DupsFound) {
                                            SQL = "select b.id from ccAddonCollectionCDefRules a,ccAddonCollectionCDefRules b where (a.id<b.id) and (a.contentid=b.contentid) and (a.collectionid=b.collectionid)";
                                            SQL = "delete from ccAddonCollectionCDefRules where id in (" + SQL + ")";
                                            cp.Db.ExecuteNonQuery(SQL);
                                        }
                                        cp.Cache.Store(cacheName, nodeHtml, DateTime.Now.AddHours(1d), env.cacheDependencyList);
                                    } catch (Exception ex) {
                                        cp.Site.ErrorReport(ex, $"Error creating nodes for Data Records, continuing");
                                    }
                                }
                            }
                            returnNav += nodeHtml;
                            break;
                        }
                    case common.NodeIDManageAddonsAdvanced: {
                            // 
                            // Special Case: clicked on Manage Add-ons.advanced
                            // edit links for Add-ons, Add-on Collections
                            // 
                            // Folder to Add-ons without Collections
                            // 
                            nodeIDString = common.NodeIDAddonsNoCollection;
                            returnNav += NodeController.getNode(cp, env, 0, 0, 0, 0, 0, "", null, 0, "Add-ons With No Collection", env.emptyNodeList, 0, common.NavIconTypeAddon, "Add-ons With No Collection", AutoManageAddons, common.NodeTypeEnum.NodeTypeEntry, false, false, env.openNodeList, nodeIDString, ref NodeNavigatorJS, "");
                            Return_NavigatorJS += NodeNavigatorJS;
                            // 
                            Name = "Add-ons";
                            returnNav += NodeController.getNode(cp, env, 0, 0, 0, 0, cp.Content.GetID(Name), "", null, 0, Name, env.emptyNodeList, 0, common.NavIconTypeContent, Name, AutoManageAddons, common.NodeTypeEnum.NodeTypeEntry, false, false, env.openNodeList, "", ref NodeNavigatorJS, "");
                            Return_NavigatorJS += NodeNavigatorJS;
                            // 
                            Name = "Add-on Collections";
                            returnNav += NodeController.getNode(cp, env, 0, 0, 0, 0, cp.Content.GetID(Name), "", null, 0, Name, env.emptyNodeList, 0, common.NavIconTypeContent, Name, AutoManageAddons, common.NodeTypeEnum.NodeTypeEntry, false, false, env.openNodeList, "", ref NodeNavigatorJS, "");
                            Return_NavigatorJS += NodeNavigatorJS;
                            // 
                            Name = "Add-on Categories";
                            returnNav += NodeController.getNode(cp, env, 0, 0, 0, 0, cp.Content.GetID(Name), "", null, 0, Name, env.emptyNodeList, 0, common.NavIconTypeContent, Name, AutoManageAddons, common.NodeTypeEnum.NodeTypeEntry, false, false, env.openNodeList, "", ref NodeNavigatorJS, "");
                            Return_NavigatorJS += NodeNavigatorJS;
                            break;
                        }
                    // 
                    case common.NodeIDAddonsNoCollection: {
                            // 
                            // special case: Add-on List that do not have collections
                            // 
                            CollectionID = 0;

                            FieldList = "0 as ContentControlID,A.Name as Name,A.ID as ID,A.ID as AddonID,0 as NewWindow,0 as ContentID,'' as LinkPage," + common.NavIconTypeAddon + " as NavIconType,A.Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as collectionid";
                            string SQL = "select" + " " + FieldList + " from ccAggregateFunctions A" + " left join ccAddonCollections C on C.ID=A.CollectionID" + " where C.ID is null" + " order by A.Name";




                            NodeType = common.NodeTypeEnum.NodeTypeAddon;
                            BlockSubNodes = true;
                            nodeIDString = "";
                            foreach (AddonModel addon in DbBaseModel.createList<AddonModel>(cp, "id in (select a.id from ccAggregateFunctions a left join ccAddonCollections C on C.ID=A.CollectionID where c.id is null)")) {
                                returnNav += NodeController.getNode(cp, env, 0, addon.contentControlId, 0, 0, 0, "", addon, 0, (addon.name ?? "").Trim(), env.emptyNodeList, 0, common.NavIconTypeAddon, cp.Utils.EncodeHTML(addon.name), AutoManageAddons, common.NodeTypeEnum.NodeTypeAddon, false, false, env.openNodeList, nodeIDString, ref NodeNavigatorJS, "");
                                Return_NavigatorJS += NodeNavigatorJS;
                            }

                            break;
                        }
                    case common.NodeIDLegacyMenu: {
                            // 
                            // Special Case: build old top menus under this Navigator entry
                            // 
                            BlockSubNodes = false;
                            string SQL = MenuSqlController.getMenuSQL(cp, "(parentid=0)or(parentid is null)", common.LegacyMenuContentName);
                            if (!csChildList.OpenSQL(SQL)) {
                                // 
                                // Empty list, add to env.EmptyNodeList
                                // 
                                env.emptyNodeList += "," + TopParentNode;
                            }

                            break;
                        }
                    case common.NodeIDAllContentList: {
                            // 
                            // special case: all content
                            // 
                            FieldList = "Name,ID,0 as AddonID,0 as NewWindow,ID as ContentID,'' as LinkPage," + common.NavIconTypeContent + " as NavIconType,Name as NavIconTitle,0 as SettingPageID,0 as HelpAddonID,0 as HelpCollectionID,0 as contentcontrolid,0 as collectionid";
                            string SQL = "select " + FieldList + " from cccontent order by name";
                            csChildList.OpenSQL(SQL);
                            NodeType = common.NodeTypeEnum.NodeTypeContent;
                            BlockSubNodes = true;
                            break;
                        }
                    case "0":
                    case var @case when @case == "": {
                            // 
                            // Navigator Entries, list home(s) plus all roots
                            // 
                            NodeType = common.NodeTypeEnum.NodeTypeEntry;
                            BlockSubNodes = false;
                            Link = cp.Utils.EncodeHTML("http://" + cp.Site.Domain);
                            returnNav += common.cr + "<div class=ccNavLink><A href=\"" + Link + "\">" + common.IconPublicHome + "</A>&nbsp;<A href=\"" + Link + "\">Public Home</A></div>";
                            Link = cp.Utils.EncodeHTML(env.adminUrl);
                            returnNav += common.cr + "<div class=ccNavLink><A href=\"" + Link + "\">" + common.IconAdminHome + "</A>&nbsp;<A href=\"" + Link + "\">Admin Home</A></div>";
                            var cs8 = cp.CSNew();
                            if (cs8.OpenSQL(MenuSqlController.getMenuSQL(cp, "((Parentid=0)or(Parentid is null))", common.NavigatorContentName))) {
                                do {
                                    Name = cs8.GetText("name").Trim();
                                    NavigatorID = cs8.GetInteger("ID");
                                    NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(Name);
                                    nodeIDString = NavigatorID.ToString();
                                    if (AutoManageAddons) {
                                        // 
                                        // special cases - root nodes that do not just deliver menu entries
                                        // 
                                        switch (Name.ToLowerInvariant()) {
                                            case "manage add-ons": {
                                                    nodeIDString = common.NodeIDManageAddons;
                                                    break;
                                                }
                                            case "settings": {
                                                    nodeIDString = common.NodeIDSettings;
                                                    break;
                                                }
                                            case "tools": {
                                                    nodeIDString = common.NodeIDTools;
                                                    break;
                                                }
                                            case "reports": {
                                                    nodeIDString = common.NodeIDReports;
                                                    break;
                                                }
                                        }
                                    }
                                    returnNav += NodeController.getNode(cp, env, 0, 0, 0, 0, 0, "", null, 0, Name, env.emptyNodeList, NavigatorID, common.NavIconTypeFolder, NavIconTitleHtmlEncoded, AutoManageAddons, common.NodeTypeEnum.NodeTypeEntry, false, false, env.openNodeList, nodeIDString, ref NodeNavigatorJS, "");
                                    Return_NavigatorJS += NodeNavigatorJS;
                                    cs8.GoNext();
                                }
                                while (cs8.OK());
                            }
                            cs8.Close();
                            // 
                            // Add a Legacy Menu node to the top-most parent menu at the very end
                            // 
                            if (cp.Utils.EncodeBoolean(cp.Site.GetText("AllowNavigatorLegacyEntry", "0"))) {
                                Name = "Legacy Menu";
                                NavIconTitleHtmlEncoded = "Legacy Menu";
                                nodeIDString = common.NodeIDLegacyMenu;
                                returnNav += NodeController.getNode(cp, env, 0, 0, 0, 0, 0, "", null, 0, Name, env.emptyNodeList, 0, common.NavIconTypeFolder, NavIconTitleHtmlEncoded, AutoManageAddons, common.NodeTypeEnum.NodeTypeEntry, false, BlockSubNodes, env.openNodeList, nodeIDString, ref NodeNavigatorJS, "");
                                Return_NavigatorJS += NodeNavigatorJS;
                            }

                            break;
                        }
                    case common.NodeIDSettings: {
                            // 
                            // list setting nodes, includes menu nodes with setting parents, and addons with type=setting sorted in
                            // 
                            returnNav += NodeListMixedController.getNodeListMixed(cp, env, env.emptyNodeList, TopParentNode, AutoManageAddons, NodeType, env.openNodeList, 3, cp.Content.GetRecordID("navigator entries", "settings"), common.NavIconTypeSetting, ref NodeNavigatorJS);
                            Return_NavigatorJS += NodeNavigatorJS;
                            break;
                        }
                    case common.NodeIDTools: {
                            // 
                            // list setting nodes, includes menu nodes with setting parents, and addons with type=setting sorted in
                            // 
                            returnNav += NodeListMixedController.getNodeListMixed(cp, env, env.emptyNodeList, TopParentNode, AutoManageAddons, NodeType, env.openNodeList, 4, cp.Content.GetRecordID("navigator entries", "tools"), common.NavIconTypeTool, ref NodeNavigatorJS);
                            Return_NavigatorJS += NodeNavigatorJS;
                            break;
                        }
                    case common.NodeIDReports: {
                            // 
                            // list setting nodes, includes menu nodes with setting parents, and addons with type=setting sorted in
                            // 
                            returnNav += NodeListMixedController.getNodeListMixed(cp, env, env.emptyNodeList, TopParentNode, AutoManageAddons, NodeType, env.openNodeList, 2, cp.Content.GetRecordID("navigator entries", "reports"), common.NavIconTypeReport, ref NodeNavigatorJS);
                            Return_NavigatorJS += NodeNavigatorJS;
                            break;
                        }

                    default: {
                            // 
                            // numeric node (default case) - list navigator records with parent=TopParentNode
                            // 
                            int CS = -1;
                            if (int.TryParse(TopParentNode, out _)) {
                                if ((env.emptyNodeList + ",").IndexOf("," + TopParentNode + ",", StringComparison.Ordinal) >= 0) {
                                    // 
                                } else {
                                    // 
                                    // Navigator Entries, child under TopParentNode
                                    // 
                                    string SQL = MenuSqlController.getMenuSQL(cp, "parentid=" + TopParentNode, "");
                                    BlockSubNodes = false;
                                    if (!csChildList.OpenSQL(SQL)) {
                                        // 
                                        // Empty list, add to env.EmptyNodeList
                                        // 
                                        env.emptyNodeList += "," + TopParentNode;
                                    }
                                }
                            }

                            break;
                        }
                }
                // 
                // ----- List Navigator Nodes, if not already displayed
                // 
                if (!csChildList.OK() & NodeType == common.NodeTypeEnum.NodeTypeEntry) {
                    // 
                    // No child nodes, if this node includes a CID, list the first 20 content records with a 'more'
                    // 
                    ContentID = 0;
                    if (int.TryParse(TopParentNode, out _)) {
                        csChildList.Close();
                        if (csChildList.Open(common.NavigatorContentName, "id=" + TopParentNode)) {
                            ContentID = csChildList.GetInteger("ContentID");
                        }
                        if (ContentID != 0) {
                            ContentID = ContentID;
                            string ContentName = cp.Content.GetRecordName("content", ContentID);
                            if (!string.IsNullOrEmpty(ContentName)) {
                                csChildList.Close();
                                int Ptr = 0;
                                if (csChildList.Open(ContentName, "", "name", true, "ID,Name,ContentControlID", 20, 1)) {
                                    env.emptyNodeList = env.emptyNodeList.Replace("," + TopParentNode, "");
                                    do {
                                        NavigatorID = csChildList.GetInteger("ID");
                                        string RecordName = csChildList.GetText("Name");
                                        if (string.IsNullOrEmpty(RecordName)) {
                                            RecordName = "Record " + NavigatorID;
                                        }
                                        // 
                                        if (RecordName.Length > 53) {
                                            RecordName = RecordName.Substring(0, 25) + "..." + RecordName.Substring(RecordName.Length - 25);
                                        }
                                        NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML("Edit '" + RecordName + "' in '" + ContentName + "'");
                                        IconNoSubNodes = common.IconRecord;
                                        IconNoSubNodes = IconNoSubNodes.Replace("{title}", NavIconTitleHtmlEncoded);
                                        Link = "?id=" + NavigatorID + "&cid=" + csChildList.GetInteger("ContentControlID") + "&af=4";
                                        ATag = "<a href=\"" + Link + "\" title=\"" + NavIconTitleHtmlEncoded + "\">";
                                        returnNav += common.cr + "<div class=\"ccNavLink ccNavLinkEmpty\">" + ATag + IconNoSubNodes + "</a>&nbsp;" + ATag + RecordName + "</a></div>";
                                        Ptr += 1;
                                        csChildList.GoNext();
                                    }
                                    while (csChildList.OK() & Ptr < 20);
                                    if (Ptr == 20) {
                                        NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML("Open All '" + common.NavigatorContentName + "'");
                                        Link = "?cid=" + ContentID;
                                        var IconClosed = default(string);
                                        returnNav += common.cr + "<div class=\"ccNavLink ccNavLinkEmpty\">" + IconClosed + "&nbsp;<a href=\"" + Link + "\" title=\"" + NavIconTitleHtmlEncoded + "\">more...</a></div>";
                                    }
                                }
                            }
                        }
                    }
                } else if (csChildList.OK()) {
                    // 
                    // List out child menus
                    // 
                    var SettingPageID = default(int);
                    do {
                        var addon = DbBaseModel.create<AddonModel>(cp, csChildList.GetInteger("AddonID"));
                        // 
                        CollectionID = csChildList.GetInteger("CollectionID");
                        NavigatorID = csChildList.GetInteger("ID");
                        Name = csChildList.GetText("name").Trim();
                        bool NewWindow = csChildList.GetBoolean("newwindow");
                        ContentID = csChildList.GetInteger("ContentID");
                        Link = csChildList.GetText("LinkPage").Trim();
                        NavIconType = csChildList.GetInteger("NavIconType");
                        NavIconTitle = csChildList.GetText("NavIconTitle");
                        int HelpAddonID = csChildList.GetInteger("HelpAddonID");
                        if (HelpAddonID != 0) {
                            HelpAddonID = HelpAddonID;
                        }
                        int helpCollectionID = csChildList.GetInteger("HelpCollectionID");
                        if (string.IsNullOrEmpty(NavIconTitle)) {
                            NavIconTitle = Name;
                        }
                        ContentControlID = csChildList.GetInteger("ContentControlID");
                        if (Name.ToLowerInvariant() == "all content") {
                            // 
                            // special case: any Navigator Entry named 'all content' returns the content list
                            // 
                            nodeIDString = common.NodeIDAllContentList;
                        } else {
                            nodeIDString = NavigatorID.ToString();
                        }
                        string linkSuffixList = "";
                        if (ContentID != 0 & env.isDeveloper) {
                            linkSuffixList = "<a href=\"" + env.contentFieldEditToolPrefix + ContentID + "\">edit</a>";
                            if (!string.IsNullOrEmpty(linkSuffixList)) {
                                linkSuffixList = "&nbsp;(" + linkSuffixList + ")";
                            }
                        }
                        NavIconTitleHtmlEncoded = cp.Utils.EncodeHTML(NavIconTitle);
                        returnNav += NodeController.getNode(cp, env, CollectionID, ContentControlID, helpCollectionID, HelpAddonID, ContentID, Link, addon, SettingPageID, Name, env.emptyNodeList, NavigatorID, NavIconType, NavIconTitleHtmlEncoded, AutoManageAddons, NodeType, NewWindow, BlockSubNodes, env.openNodeList, nodeIDString, ref NodeNavigatorJS, linkSuffixList);
                        Return_NavigatorJS += NodeNavigatorJS;
                        csChildList.GoNext();
                    }
                    while (csChildList.OK());
                    csChildList.Close();
                }
                return returnNav;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex, $"hint [{hint}]");
                throw;
            }
        }
    }
}