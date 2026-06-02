using System;
using Contensive.BaseClasses;

namespace Contensive.AdminNavigator {
    public class AdminNavigatorAddon : AddonBaseClass {
        // 
        public override object Execute(CPBaseClass CP) {
            try {
                using (var env = new ApplicationEnvironmentModel(CP)) {
                    string AdminNavContent;
                    string AdminNavHead;
                    if (env.adminNavOpen & env.allowAdminNavSaveState) {
                        // 
                        // draw the page with Nav Opened
                        CP.Doc.SetProperty("nodeid", "0");
                        string AdminNavContentOpened = "" + common.cr + "<div id=\"AdminNavContentOpened\" class=\"opened\">" + common.kmaIndent(GetNodeRemote.getNode(CP, env) + "<img alt=\"space\" src=\"https://s3.amazonaws.com/cdn.contensive.com/assets/20190729/images/spacer.gif\" width=\"200\" height=\"1\" style=\"clear:both\">") + common.cr + "</div>" + "";




                        string AdminNavContentClosed = "" + common.cr + "<div id=\"AdminNavContentClosed\" class=\"closed\" style=\"display:none;\">" + common.NavigatorClosedLabel + "</div>" + "";

                        // 
                        string AdminNavHeadOpened = "" + common.cr + "<div id=\"AdminNavHeadOpened\" class=\"opened\">" + common.cr + "\t" + "<table border=0 cellpadding=0 cellspacing=0 width=\"100%\"><tr>" + common.cr + "\t" + "\t" + "<td valign=Middle class=\"left\">Navigator</td>" + common.cr + "\t" + "\t" + "<td valign=Middle class=\"right\"><a href=\"#\" onClick=\"closeAdminNav();return false\">" + common.iconCloseWindow + "</a></td>" + common.cr + "\t" + "</tr></table>" + common.cr + "</div>" + "";






                        string AdminNavHeadClosed = "" + common.cr + "<div id=\"AdminNavHeadClosed\" class=\"closed\" style=\"display:none;\">" + common.cr + "\t" + "<a href=\"#\" onClick=\"reOpenAdminNav();return false\">" + common.iconOpenWindow + "</a>" + common.cr + "</div>" + "";



                        AdminNavHead = "" + common.cr + "<div class=\"ccHeaderCon\">" + common.kmaIndent(AdminNavHeadOpened) + common.kmaIndent(AdminNavHeadClosed) + common.cr + "</div>" + "";




                        AdminNavContent = "" + common.cr + "<div class=\"ccContentCon\">" + common.kmaIndent(AdminNavContentOpened) + common.kmaIndent(AdminNavContentClosed) + common.cr + "</div>" + "";




                    } else {
                        // 
                        // draw the page with Nav Closed
                        // 
                        string AdminNavHeadOpened = "" + common.cr + "<div id=\"AdminNavHeadOpened\" class=\"opened\" style=\"display:none;\">" + common.cr + "\t" + "<table border=0 cellpadding=0 cellspacing=0 width=\"100%\"><tr>" + common.cr + "\t" + "\t" + "<td valign=Middle class=\"left\">Navigator</td>" + common.cr + "\t" + "\t" + "<td valign=Middle class=\"right\"><a href=\"#\" onClick=\"reCloseAdminNav();return false\">" + common.iconCloseWindow + "</a></td>" + common.cr + "\t" + "</tr></table>" + common.cr + "</div>" + "";






                        string AdminNavHeadClosed = "" + common.cr + "<div id=\"AdminNavHeadClosed\" class=\"closed\">" + common.cr + "\t" + "<a href=\"#\" onClick=\"OpenAdminNav();return false\">" + common.iconOpenWindow + "</a>" + common.cr + "</div>" + "";



                        AdminNavHead = "" + common.cr + "<div class=\"ccHeaderCon\">" + common.kmaIndent(AdminNavHeadOpened) + common.kmaIndent(AdminNavHeadClosed) + common.cr + "</div>" + "";




                        AdminNavContent = "" + common.cr + "<div class=\"ccContentCon\">" + common.cr + "<div id=\"AdminNavContentOpened\" class=\"opened\" style=\"display:none;\"><div style=\"text-align:center;\"><img src=\"https://s3.amazonaws.com/cdn.contensive.com/assets/20190729/images/ajax-loader-small.gif\" width=16 height=16></div></div>" + common.cr + "<div id=\"AdminNavContentClosed\" class=\"closed\">" + common.NavigatorClosedLabel + "</div>" + common.cr + "<div id=\"AdminNavContentMinWidth\" style=\"display:none;\"><img alt=\"space\" src=\"https://s3.amazonaws.com/cdn.contensive.com/assets/20190729/images/spacer.gif\" width=\"200\" height=\"1\" style=\"clear:both\"></div>" + common.cr + "</div>";




                    }
                    string returnHtml = "" + common.cr + "<div>" + common.kmaIndent(AdminNavHead) + common.kmaIndent(AdminNavContent) + common.cr + "</div>" + "<input type=\"hidden\" name=\"allowAdminNavSaveState\" value=\"" + (env.allowAdminNavSaveState ? "true" : "false") + "\">";




                    return returnHtml;
                }
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}