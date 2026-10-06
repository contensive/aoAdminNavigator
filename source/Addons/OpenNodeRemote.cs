using System;
using System.Linq;
using Contensive.BaseClasses;

namespace Contensive.AdminNavigator {
    public class OpenNodeRemote : AddonBaseClass {
        //
        public override object Execute(CPBaseClass CP) {
            try {
                // Authentication required - admin navigation is admin-only
                if (!CP.User.IsAdmin) {
                    return string.Empty;
                }
                string nodeId = CP.Doc.GetText("nodeid");
                if (!string.IsNullOrWhiteSpace(nodeId)) {
                    var nodeList = CP.Visit.GetText("AdminNavOpenNodeList", "").Split(',').ToList();
                    if (!nodeList.Contains(nodeId)) {
                        nodeList.Add(nodeId);
                        CP.Visit.SetProperty("AdminNavOpenNodeList", string.Join(",", nodeList));
                    }
                }
                return string.Empty;
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}