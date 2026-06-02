using System;
using Contensive.BaseClasses;

namespace Contensive.AdminNavigator {
    public sealed class NodeController {
        // 
        // ====================================================================================================
        /// <summary>
        /// get a single node and its children
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="env"></param>
        /// <param name="CollectionID"></param>
        /// <param name="ContentControlID"></param>
        /// <param name="helpCollectionID"></param>
        /// <param name="HelpAddonID"></param>
        /// <param name="ContentID"></param>
        /// <param name="link"></param>
        /// <param name="addon"></param>
        /// <param name="ignore"></param>
        /// <param name="Name"></param>
        /// <param name="xEmptyNodeList"></param>
        /// <param name="NavigatorID"></param>
        /// <param name="NavIconType"></param>
        /// <param name="NavIconTitleHtmlEncoded"></param>
        /// <param name="AutoManageAddons"></param>
        /// <param name="NodeType"></param>
        /// <param name="NewWindow"></param>
        /// <param name="BlockSubNodes"></param>
        /// <param name="xOpenNodeList"></param>
        /// <param name="NodeIDString"></param>
        /// <param name="Return_NavigatorJS"></param>
        /// <param name="linkSuffixList"></param>
        /// <returns></returns>
        internal static string getNode(CPBaseClass cp, ApplicationEnvironmentModel env, int CollectionID, int ContentControlID, int helpCollectionID, int HelpAddonID, int ContentID, string link, Models.Db.AddonModel addon, int ignore, string Name, string xEmptyNodeList, int NavigatorID, int NavIconType, string NavIconTitleHtmlEncoded, bool AutoManageAddons, common.NodeTypeEnum NodeType, bool NewWindow, bool BlockSubNodes, string xOpenNodeList, string NodeIDString, ref string Return_NavigatorJS, string linkSuffixList) {
            int hint = 0;
            try {
                string SubNav;
                string DivIDClosed;
                string DivIDOpened;
                string DivIDContent;
                string DivIDEmpty;
                string result = "";
                string DivIDBase;
                string AddonGuid;
                string AddonName;
                bool IsVisible;
                string WorkingName;
                // 
                // Determine if it has either a function, or child entries
                // 
                bool BlockNode = false;
                Return_NavigatorJS = "";
                WorkingName = Name;
                IsVisible = CollectionID != 0 | helpCollectionID != 0 | HelpAddonID != 0 | ContentID != 0 | !string.IsNullOrEmpty(link) | addon is not null | WorkingName.ToLowerInvariant() == "all content" | WorkingName.ToLowerInvariant() == "add-ons with no collection";
                if (!IsVisible) {
                    // 
                    // IsVisible if it is not in the EmptyNodeList (has child entries)
                    // 
                    IsVisible = ("," + env.emptyNodeList + ",").IndexOf("," + NavigatorID + ",", StringComparison.Ordinal) < 0;
                }
                if (IsVisible) {
                    // 
                    // hide the legacy node 'switch to navigator'
                    // 
                    IsVisible = WorkingName.ToLowerInvariant() != "switch to navigator";
                }
                if (IsVisible) {
                    string IconOpened = "";
                    string IconClosed = "";
                    string IconNoSubNodes = "";
                    // 
                    // Setup Icons
                    // 
                    switch (NavIconType) {
                        case common.NavIconTypeCustom: {
                                break;
                            }
                        // 
                        // reserved for future addition of a custom Icon field
                        // not done now because there is no facility now to import files during collection build
                        // 
                        case common.NavIconTypeAdvanced: {
                                IconOpened = common.IconAdvancedOpened;
                                IconClosed = common.IconAdvancedClosed;
                                IconNoSubNodes = common.IconAdvanced;
                                break;
                            }
                        case common.NavIconTypeContent: {
                                IconOpened = common.IconContentOpened;
                                IconClosed = common.IconContentClosed;
                                IconNoSubNodes = common.IconContent;
                                break;
                            }
                        case common.NavIconTypeEmail: {
                                IconOpened = common.IconEmailOpened;
                                IconClosed = common.IconEmailClosed;
                                IconNoSubNodes = common.IconEmail;
                                break;
                            }
                        case common.NavIconTypeUser: {
                                IconOpened = common.IconUsersOpened;
                                IconClosed = common.IconUsersClosed;
                                IconNoSubNodes = common.IconUsers;
                                break;
                            }
                        case common.NavIconTypeReport: {
                                IconOpened = common.IconReportsOpened;
                                IconClosed = common.IconReportsClosed;
                                IconNoSubNodes = common.IconReports;
                                break;
                            }
                        case common.NavIconTypeSetting: {
                                IconOpened = common.IconSettingsOpened;
                                IconClosed = common.IconSettingsClosed;
                                IconNoSubNodes = common.IconSettings;
                                break;
                            }
                        case common.NavIconTypeTool: {
                                IconOpened = common.IconToolsOpened;
                                IconClosed = common.IconToolsClosed;
                                IconNoSubNodes = common.IconTools;
                                break;
                            }
                        case common.NavIconTypeRecord: {
                                IconOpened = common.IconRecordOpened;
                                IconClosed = common.IconRecordClosed;
                                IconNoSubNodes = common.IconRecord;
                                break;
                            }
                        case common.NavIconTypeAddon: {
                                IconOpened = common.IconAddonsOpened;
                                IconClosed = common.IconAddonsClosed;
                                IconNoSubNodes = common.IconAddons;
                                break;
                            }
                        case common.NavIconTypeHelp: {
                                IconOpened = common.IconHelp;
                                IconClosed = common.IconHelp;
                                IconNoSubNodes = common.IconHelp; // NavIconTypeFolder
                                break;
                            }

                        default: {
                                IconOpened = common.IconFolderOpened;
                                IconClosed = common.IconFolderClosed;
                                IconNoSubNodes = common.IconFolderNoSubNodes;
                                break;
                            }
                    }
                    IconOpened = IconOpened.Replace("{title}", $"Close {NavIconTitleHtmlEncoded}");
                    IconClosed = IconClosed.Replace("{title}", $"Open {NavIconTitleHtmlEncoded}");
                    IconNoSubNodes = IconNoSubNodes.Replace("{title}", NavIconTitleHtmlEncoded);
                    // 
                    // NodeIDString - the unique string that is passed by here as ParentNode to get all the child nodes
                    // is always the navigator entry ID, unless it is a hardcoded subsection
                    // DIVID must be unique for this entire menu, but does not need to be recognized in a call back to the server
                    // 
                    // 
                    // set flag for 'hardcoded' lists - like add-ons
                    // 
                    if (AutoManageAddons & WorkingName.ToLowerInvariant() == "manage add-ons") {
                        // 
                        // test special case - replace manage add-ons branch
                        // 
                        DivIDBase = NavigatorID.ToString();
                        // NodeIDString = NodeIDManageAddons
                    } else {
                        switch (WorkingName.ToLowerInvariant()) {
                            case "legacy menu": {
                                    // 
                                    // special case - if node has this name, a click to it calls back with NodeIDLegacyMenu
                                    // 

                                    DivIDBase = NavigatorID.ToString();
                                    break;
                                }
                            case "add-ons with no collection": {
                                    // 
                                    // any Navigator Entry named 'all content' returns the content list
                                    // 
                                    DivIDBase = NavigatorID.ToString();
                                    break;
                                }
                            case "all content": {
                                    // 
                                    // any Navigator Entry named 'all content' returns the content list
                                    // 
                                    DivIDBase = NavigatorID.ToString();
                                    break;
                                }

                            default: {
                                    // 
                                    // This entry is made from a navigator entry record
                                    // 
                                    switch (NodeType) {
                                        case common.NodeTypeEnum.NodeTypeAddon: {
                                                // 
                                                // List of Addons
                                                // 
                                                DivIDBase = "a" + NavigatorID;
                                                break;
                                            }
                                        case common.NodeTypeEnum.NodeTypeCollection: {
                                                // 
                                                // List of Collections
                                                // 
                                                DivIDBase = "c" + CollectionID;
                                                break;
                                            }
                                        case common.NodeTypeEnum.NodeTypeContent: {
                                                // 
                                                // List of content
                                                // 
                                                DivIDBase = "d" + NavigatorID;
                                                break;
                                            }

                                        default: {
                                                // 
                                                // List of Navigator Entries
                                                // 
                                                DivIDBase = "n" + NavigatorID;
                                                break;
                                            }
                                    }

                                    break;
                                }
                        }
                    }
                    // 
                    // check name for length
                    // 
                    if (WorkingName.Length > 53) {
                        WorkingName = WorkingName.Substring(0, 25) + "..." + WorkingName.Substring(WorkingName.Length - 25);
                    }
                    string NavLinkHTMLId = "";
                    string workingNameHtmlEncoded = "";
                    // 
                    // setup link
                    // 
                    if (!string.IsNullOrEmpty(link)) {
                        NavLinkHTMLId = "n" + NavigatorID;
                        workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName);
                        if (NewWindow) {
                            workingNameHtmlEncoded = "<a class=navDrag name=\"navLink\" id=\"" + NavLinkHTMLId + "\" href=\"" + link + "\" target=\"_blank\" title=\"Open '" + workingNameHtmlEncoded + "'\">" + workingNameHtmlEncoded + "</a>";
                        } else {
                            workingNameHtmlEncoded = "<a class=navDrag name=navLink id=\"" + NavLinkHTMLId + "\" href=\"" + link + "\" title=\"Open '" + workingNameHtmlEncoded + "'\">" + workingNameHtmlEncoded + "</a>";
                        }
                    } else {
                        hint = 1000;
                        // 
                        // If link page, use this
                        // 
                        if (addon is not null) {
                            hint = 1010;
                            // 
                            // link to addon
                            // 
                            link = "";
                            AddonGuid = addon.ccguid;
                            AddonName = addon.name;
                            if (addon.remoteMethod) {
                                hint = 1020;
                                NewWindow = true;
                                // 
                                // -- url encode, but encode space as %20, not plus
                                foreach (var linkPart in AddonName.Split(' '))
                                    link += "%20" + cp.Utils.EncodeUrl(linkPart);
                                link = link.Substring(3);
                                hint = 1030;
                                // 
                                if (!string.IsNullOrEmpty(link)) {
                                    link = link.Replace(@"\", "/");
                                    if (link.Substring(0, 1) != "/") {
                                        link = "/" + link;
                                    }
                                }
                                hint = 1040;
                                // link = env.adminUrl & "?" & RequestNameRemoteMethodAddon & "=" & cp.Utils.EncodeRequestVariable(AddonName)
                            }
                            if (string.IsNullOrEmpty(link)) {
                                if (!string.IsNullOrEmpty(AddonGuid)) {
                                    link = cp.Utils.ModifyLinkQueryString(env.adminUrl, "addonguid", AddonGuid, true);
                                } else {
                                    link = cp.Utils.ModifyLinkQueryString(env.adminUrl, "addonid", addon.id.ToString(), true);
                                }
                            }
                            NavLinkHTMLId = "a" + addon.id;
                            workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName);
                            if (NewWindow) {
                                workingNameHtmlEncoded = "<a class=navDrag name=navLink id=\"" + NavLinkHTMLId + "\" href=\"" + link + "\" target=\"_blank\" title=\"Run '" + workingNameHtmlEncoded + "'\">" + workingNameHtmlEncoded + "</a>";
                            } else {
                                workingNameHtmlEncoded = "<a class=navDrag name=navLink id=\"" + NavLinkHTMLId + "\" href=\"" + link + "\" title=\"Run '" + workingNameHtmlEncoded + "'\">" + workingNameHtmlEncoded + "</a>";
                            }
                            // 
                            if (env.allowToolTip & (addon.template | addon.content | addon.navTypeId > 1)) {
                                // 
                                // -- add tooltip for addons
                                workingNameHtmlEncoded += "<i data-tooltip=\"" + AddonGuid + "\" class=\"contensiveToolTip fas fa-info-circle\" href=\"#\" data-toggle=\"tooltip\" data-html=\"true\" rel=\"tooltip\" title=\"Click for details.\"></i>";
                            }
                        } else if (ContentID != 0) {
                            hint = 2000;
                            // 
                            // go edit the content
                            // 
                            link = cp.Utils.ModifyLinkQueryString(env.adminUrl, "cid", ContentID.ToString(), true);
                            NavLinkHTMLId = "c" + ContentID;
                            workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName);
                            if (NewWindow) {
                                workingNameHtmlEncoded = "<a class=navDrag name=navLink id=\"" + NavLinkHTMLId + "\" href=\"" + link + "\" target=\"_blank\" title=\"List All '" + NavIconTitleHtmlEncoded + "'\">" + NavIconTitleHtmlEncoded + "</a>";
                            } else {
                                workingNameHtmlEncoded = "<a class=navDrag name=navLink id=\"" + NavLinkHTMLId + "\" href=\"" + link + "\" title=\"List All '" + NavIconTitleHtmlEncoded + "'\">" + NavIconTitleHtmlEncoded + "</a>";
                            }
                        } else if (HelpAddonID != 0) {
                            hint = 3000;
                            // 
                            // go to Addon Help
                            // 
                            link = cp.Utils.ModifyLinkQueryString(env.adminUrl, "helpaddonid", HelpAddonID.ToString(), true);
                            workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName);
                            if (NewWindow) {
                                workingNameHtmlEncoded = "<a href=\"" + link + "\" target=\"_blank\" title=\"Help for Add-on '" + NavIconTitleHtmlEncoded + "'\">" + NavIconTitleHtmlEncoded + "</a>";
                            } else {
                                workingNameHtmlEncoded = "<a href=\"" + link + "\" title=\"Help for Add-on '" + NavIconTitleHtmlEncoded + "'\">" + NavIconTitleHtmlEncoded + "</a>";
                            }
                        } else if (helpCollectionID != 0) {
                            hint = 4000;
                            // 
                            // go to Collection Help
                            // 
                            using (var cs13 = cp.CSNew()) {
                                BlockNode = true;
                                if (cs13.Open("add-on collections", "id=" + helpCollectionID, "name", true, "name,helpLink,help")) {

                                    string collectionName = cs13.GetText("name");
                                    string collectionHelpLink = cs13.GetText("helpLink");
                                    string collectionHelp = cs13.GetText("help");
                                    // 
                                    WorkingName = collectionName;
                                    link = collectionHelpLink;
                                    workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName);
                                    if (!string.IsNullOrEmpty(link)) {
                                        BlockNode = false;
                                        NewWindow = true;
                                    } else if (!string.IsNullOrEmpty(collectionHelp)) {
                                        BlockNode = false;
                                        link = cp.Utils.ModifyLinkQueryString(env.adminUrl, "helpcollectionid", helpCollectionID.ToString(), true);
                                    }
                                    if (!BlockNode) {
                                        if (NewWindow) {
                                            workingNameHtmlEncoded = "<a href=\"" + link + "\" target=\"_blank\" title=\"Help for Collection '" + workingNameHtmlEncoded + "'\">Help</a>";
                                        } else {
                                            workingNameHtmlEncoded = "<a href=\"" + link + "\" title=\"Help for Collection '" + workingNameHtmlEncoded + "'\">Help</a>";
                                        }
                                    }

                                }
                                cs13.Close();
                            }
                        } else {
                            hint = 5000;
                            workingNameHtmlEncoded = cp.Utils.EncodeHTML(WorkingName);
                        }
                    }
                    hint = 6000;
                    // 
                    if (!BlockNode) {
                        hint = 6010;
                        if (BlockSubNodes | addon is not null) {
                            hint = 6020;
                            // 
                            // This is a hardcoded item (like Add-on), it has no subnodes
                            // 
                            result += common.cr + "<div class=\"ccNavLink ccNavLinkEmpty\">" + IconNoSubNodes + workingNameHtmlEncoded + linkSuffixList + "</div>";
                        } else {
                            hint = 6030;
                            // 
                            DivIDClosed = DivIDBase + "a";
                            DivIDOpened = DivIDBase + "b";
                            DivIDContent = DivIDBase + "c";
                            DivIDEmpty = DivIDBase + "d";

                            if (ContentID == 0 & (env.emptyNodeList + ",").IndexOf("," + NodeIDString + ",", StringComparison.Ordinal) >= 0) {
                                hint = 6040;
                                // 
                                // In EmptyNodeList
                                // 
                                result += common.cr + "<div class=\"ccNavLink ccNavLinkEmpty\">" + IconNoSubNodes + workingNameHtmlEncoded + linkSuffixList + "</div>";
                            } else if ((env.openNodeList + ",").IndexOf("," + NodeIDString + ",", StringComparison.Ordinal) >= 0) {
                                hint = 6050;
                                string NodeNavigatorJS = "";
                                // 
                                // This node is open
                                // 
                                SubNav = NodeListController.getNodeList(cp, env, NodeIDString, ref NodeNavigatorJS);
                                Return_NavigatorJS += NodeNavigatorJS;
                                if (!string.IsNullOrEmpty(SubNav)) {
                                    // 
                                    // display the subnav
                                    // 
                                    result += common.cr + "<div class=ccNavLink ID=" + DivIDClosed + " style=\"display:none;\"><A class=\"ccNavClosed\" href=\"#\" onclick=\"AdminNavOpenClick('" + DivIDClosed + "','" + DivIDOpened + "','" + DivIDContent + "','" + NodeIDString + "','" + DivIDEmpty + "');return false;\">" + IconClosed + "</A>&nbsp;" + workingNameHtmlEncoded + linkSuffixList + "</div>";
                                    result += common.cr + "<div class=ccNavLink ID=" + DivIDOpened + "><A class=\"ccNavOpened\" href=\"#\" onclick=\"AdminNavCloseClick('" + DivIDOpened + "','" + DivIDClosed + "','" + DivIDContent + "','" + NodeIDString + "');return false;\">" + IconOpened + "</A>&nbsp;" + workingNameHtmlEncoded + linkSuffixList + "</div>";
                                    result = result + common.cr + "<div class=ccNavLinkChild ID=" + DivIDContent + ">" + common.kmaIndent(SubNav) + common.cr + "</div>";


                                } else {
                                    // 
                                    // it has a NO subnav
                                    // 
                                    // -- what happens if we don't display empty nodes
                                    // result &= cr & "<div class=""ccNavLink ccNavLinkEmpty"">" & IconNoSubNodes & workingNameHtmlEncoded & linkSuffixList & "</div>"
                                }
                            } else {
                                hint = 6060;
                                // 
                                // This node is closed
                                // 
                                result += common.cr + "<div class=ccNavLink ID=" + DivIDClosed + " ><A class=\"ccNavClosed\" href=\"#\" onclick=\"AdminNavOpenClick('" + DivIDClosed + "','" + DivIDOpened + "','" + DivIDContent + "','" + NodeIDString + "','','" + DivIDContent + "');return false;\">" + IconClosed + "</A>&nbsp;" + workingNameHtmlEncoded + linkSuffixList + "</div>";
                                result += common.cr + "<div class=ccNavLink ID=" + DivIDOpened + " style=\"display:none;\"><A class=\"ccNavOpened\" href=\"#\" onclick=\"AdminNavCloseClick('" + DivIDOpened + "','" + DivIDClosed + "','" + DivIDContent + "','" + NodeIDString + "');return false;\">" + IconOpened + "</A>&nbsp;" + workingNameHtmlEncoded + linkSuffixList + "</div>";
                                result += common.cr + "<div class=\"ccNavLink ccNavLinkEmpty\" ID=" + DivIDEmpty + " style=\"display:none;\">" + IconNoSubNodes + workingNameHtmlEncoded + linkSuffixList + "</div>";
                                result += common.cr + "<div class=ccNavLinkChild ID=" + DivIDContent + " style=\"display:none;margin-left:20px;\">&nbsp;&nbsp;&nbsp;&nbsp;<img src=\"https://s3.amazonaws.com/cdn.contensive.com/assets/20190729/images/ajax-loader-small.gif\" width=\"16\" height=\"16\"></div>";
                            }
                        }
                    }
                }
                // 
                return result;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex, $"NodeController.getNode, hint [{hint}]");
                throw;
            }
        }
    }
}