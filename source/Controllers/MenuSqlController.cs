using System;
using Contensive.BaseClasses;

namespace Contensive.AdminNavigator {
    public sealed class MenuSqlController {
        // 
        // ====================================================================================================
        /// <summary>
        /// Get sql for menu, if MenuContentName is blank, it will select values from either cdef
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="ParentCriteria"></param>
        /// <param name="menuContentName"></param>
        /// <returns></returns>
        public static string getMenuSQL(CPBaseClass cp, string ParentCriteria, string menuContentName) {
            string getMenuSQLRet = default;
            try {
                string criteria = "(Active<>0)";
                if (!string.IsNullOrEmpty(menuContentName)) {
                    criteria += "AND" + cp.Content.GetContentControlCriteria(menuContentName);
                }
                if (cp.User.IsDeveloper) {
                    // 
                    // -- Developer
                } else if (cp.User.IsAdmin) {
                    // 
                    // -- Administrator
                    criteria = criteria + "AND((DeveloperOnly is null)or(DeveloperOnly=0))" + "AND(ID in (" + " SELECT AllowedEntries.ID" + " FROM CCMenuEntries AllowedEntries LEFT JOIN ccContent ON AllowedEntries.ContentID = ccContent.ID" + " Where ((ccContent.Active<>0)And((ccContent.DeveloperOnly is null)or(ccContent.DeveloperOnly=0)))" + "OR(ccContent.ID Is Null)" + "))";






                } else {
                    // 
                    // ----- Content Manager
                    // 
                    string ContentManagementList = getContentManagementList(cp);
                    // 
                    string CMCriteria;
                    if (string.IsNullOrEmpty(ContentManagementList)) {
                        CMCriteria = "(1=0)";
                        // ContentManagementList = Mid(ContentManagementList, 2, Len(ContentManagementList) - 2)
                    } else if (!ContentManagementList.Contains(",")) {
                        CMCriteria = "(ccContent.ID=" + ContentManagementList + ")";
                    } else {
                        CMCriteria = "(ccContent.ID in (" + ContentManagementList + "))";
                    }

                    criteria = criteria + "AND((DeveloperOnly is null)or(DeveloperOnly=0))" + "AND((AdminOnly is null)or(AdminOnly=0))" + "AND(ID in (" + " SELECT AllowedEntries.ID" + " FROM CCMenuEntries AllowedEntries LEFT JOIN ccContent ON AllowedEntries.ContentID = ccContent.ID" + " Where (" + CMCriteria + "and(ccContent.Active<>0)And((ccContent.DeveloperOnly is null)or(ccContent.DeveloperOnly=0))And((ccContent.AdminOnly is null)or(ccContent.AdminOnly=0)))" + "OR(ccContent.ID Is Null)" + "))";







                }
                if (!string.IsNullOrEmpty(ParentCriteria)) {
                    criteria = "(" + ParentCriteria + ")AND" + criteria;
                }
                string SelectList = "ccMenuEntries.contentcontrolid, ccMenuEntries.Name, ccMenuEntries.ID, ccMenuEntries.LinkPage, ccMenuEntries.ContentID, ccMenuEntries.NewWindow, ccMenuEntries.ParentID, ccMenuEntries.AddonID, ccMenuEntries.NavIconType, ccMenuEntries.NavIconTitle, HelpAddonID,HelpCollectionID,0 as collectionid";
                getMenuSQLRet = "select " + SelectList + " from ccMenuEntries where " + criteria + " order by ccMenuEntries.Name";
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }

            return getMenuSQLRet;
        }
        // 
        // ===========================================================================
        /// <summary>
        /// Get Authoring List, returns a comma delimited list of ContentIDs that the Member can author
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        private static string getContentManagementList(CPBaseClass cp) {
            try {
                string result = "";
                string sqlTablememberRules = cp.Content.GetTable("Member Rules");
                string SQL = "Select ccGroupRules.ContentID as ID" + " FROM ((" + sqlTablememberRules + " Left Join ccGroupRules on " + sqlTablememberRules + ".GroupID=ccGroupRules.GroupID)" + " Left Join ccContent on ccGroupRules.ContentID=ccContent.ID)" + " WHERE" + " (" + sqlTablememberRules + ".MemberID=" + cp.User.Id + ")" + " AND(ccGroupRules.Active<>0)" + " AND(ccContent.Active<>0)" + " AND(" + sqlTablememberRules + ".Active<>0)";







                using (var cs = cp.CSNew()) {
                    if (cs.OpenSQL(SQL)) {
                        do {
                            int ContentID = cs.GetInteger("id");
                            result += "," + ContentID.ToString();
                            string contentName = cp.Content.GetRecordName("content", ContentID);
                            if (!string.IsNullOrEmpty(contentName)) {
                                string ChildIDList = cp.Content.GetProperty(contentName, "ChildIDList");
                                if (!string.IsNullOrEmpty(ChildIDList)) {
                                    result += "," + ChildIDList;
                                }
                            }
                            cs.GoNext();
                        }
                        while (cs.OK());
                    }
                    cs.Close();
                }
                if (result.Length > 1) {
                    result = result.Substring(1);
                }
                // 
                return result;
            } catch (Exception ex) {
                cp.Site.ErrorReport("");
                throw;
            }
        }
    }
}