using System;
using System.Linq;
using Contensive.BaseClasses;

namespace Contensive.AdminNavigator {
    public class CloseNodeRemote : AddonBaseClass {
        // 
        public override object Execute(CPBaseClass CP) {
            try {
                string nodeId = CP.Doc.GetText("nodeid");
                if (!string.IsNullOrWhiteSpace(nodeId)) {
                    var nodeList = CP.Visit.GetText("AdminNavOpenNodeList", "").Split(',').ToList();
                    if (nodeList.Contains(nodeId)) {
                        nodeList.Remove(nodeId);
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