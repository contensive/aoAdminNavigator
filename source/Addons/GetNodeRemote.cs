using System;
using Contensive.BaseClasses;

namespace Contensive.AdminNavigator {
    public class GetNodeRemote : AddonBaseClass {
        // 
        // ====================================================================================================
        /// <summary>
        /// Return the full navigator
        /// </summary>
        /// <param name="CP"></param>
        /// <returns></returns>
        public override object Execute(CPBaseClass CP) {
            return getNode(CP, new ApplicationEnvironmentModel(CP));
        }
        // 
        // ====================================================================================================
        // 
        public static string getNode(CPBaseClass CP, ApplicationEnvironmentModel env) {
            try {
                string parentNode = CP.Doc.GetText("nodeid");
                if (!string.IsNullOrEmpty(parentNode)) {
                    if (string.IsNullOrEmpty(env.openNodeList)) {
                        env.openNodeList = "," + parentNode;
                    } else if ((env.openNodeList + ",").IndexOf("," + parentNode.ToString() + ",") < 0) {
                        env.openNodeList = env.openNodeList + "," + parentNode;
                    }
                }
                string returnJs = "";
                string returnHtml = NodeListController.getNodeList(CP, env, parentNode, ref returnJs);
                if (!string.IsNullOrEmpty(returnJs)) {
                    returnJs = "if(window.navDrop) {" + returnJs + "};";
                    CP.Doc.AddBodyEnd("<script language=\"javascript\" type=\"text/javascript/\">jQuery(document).ready(function(){" + returnJs + "})</script>");
                }
                return returnHtml;
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}