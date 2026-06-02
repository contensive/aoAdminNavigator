using System;
using Contensive.BaseClasses;

namespace Contensive.AdminNavigator {
    // 
    public class ApplicationEnvironmentModel : IDisposable {
        // 
        private readonly CPBaseClass cp;
        // 
        // ==========================================================================================
        /// <summary>
        /// normalized admin url, leading slash, no trailing slash
        /// </summary>
        /// <returns></returns>
        public string adminUrl {
            get {
                try {
                    if (adminUrl_local is not null)
                        return adminUrl_local;
                    adminUrl_local = cp.Site.GetText("adminurl");
                    if (string.IsNullOrEmpty(adminUrl_local))
                        return "";
                    if (adminUrl_local.Substring(0, 1) != "/")
                        adminUrl_local = "/" + adminUrl_local;
                    if (adminUrl_local.Length == 1)
                        return adminUrl_local;
                    while (adminUrl_local.Length > 0 && adminUrl_local.Substring(adminUrl_local.Length - 1, 1) == "/")
                        adminUrl_local = adminUrl_local.Substring(0, adminUrl_local.Length - 1);
                    return adminUrl_local;

                } catch (Exception ex) {
                    cp.Site.ErrorReport(ex, $"hint adminUrl");
                    throw;
                }
            }
        }
        private string adminUrl_local = null;
        // 
        // ==========================================================================================
        /// <summary>
        /// The user. this was added because contensive returns cp.user.developer true for admin (fixed 11/11/2021)
        /// </summary>
        /// <returns></returns>
        public Models.Db.PersonModel user {
            get {
                if (user_local is not null)
                    return user_local;
                user_local = Models.Db.DbBaseModel.create<Models.Db.PersonModel>(cp, cp.User.Id);
                return user_local;
            }
        }
        private Models.Db.PersonModel user_local = null;
        // 
        // ==========================================================================================
        /// <summary>
        /// is the current user a administrator
        /// </summary>
        /// <returns></returns>
        public bool isAdministrator {
            get {
                return user.admin | user.developer;
            }
        }
        // 
        // ==========================================================================================
        /// <summary>
        /// is the current user a developer
        /// </summary>
        /// <returns></returns>
        public bool isDeveloper {
            get {
                return user.developer;
            }
        }
        // 
        // ==========================================================================================
        public string buildVersion {
            get {
                if (buildVersion_local is not null)
                    return buildVersion_local;
                buildVersion_local = cp.Site.GetText("buildversion");
                return buildVersion_local;
            }
        }
        private string buildVersion_local = null;
        // 
        // ==========================================================================================
        public string addonEditAddonUrlPrefix {
            get {
                if (addonEditAddonUrlPrefix_local is not null)
                    return addonEditAddonUrlPrefix_local;
                addonEditAddonUrlPrefix_local = adminUrl + "?cid=" + cp.Content.GetID("add-ons") + "&af=4&id=";
                return addonEditAddonUrlPrefix_local;
            }
        }
        private string addonEditAddonUrlPrefix_local = null;
        // 
        // ==========================================================================================
        public string addonEditCollectionUrlPrefix {
            get {
                if (addonEditCollectionUrlPrefix_local is not null)
                    return addonEditCollectionUrlPrefix_local;
                addonEditCollectionUrlPrefix_local = adminUrl + "?cid=" + cp.Content.GetID("add-on collections") + "&af=4&id=";
                return addonEditCollectionUrlPrefix_local;
            }
        }
        private string addonEditCollectionUrlPrefix_local = null;
        // 
        // ==========================================================================================
        public string contentFieldEditToolPrefix {
            get {
                if (contentFieldEditToolPrefix_local is not null)
                    return contentFieldEditToolPrefix_local;
                contentFieldEditToolPrefix_local = adminUrl + "?af=105&contentid=";
                return contentFieldEditToolPrefix_local;
            }
        }
        private string contentFieldEditToolPrefix_local = null;
        // 
        // ==========================================================================================
        public string cacheDependencyList {
            get {
                if (cacheDependencyList_local is not null)
                    return cacheDependencyList_local;
                addonEditCollectionUrlPrefix_local = "";
                cacheDependencyList_local = cp.Cache.CreateTableDependencyKey("ccMenuEntries", "default") + "," + cp.Cache.CreateTableDependencyKey("ccaggregatefunctions", "default");
                return cacheDependencyList_local;
            }
        }
        private string cacheDependencyList_local = null;
        // 
        // ==========================================================================================
        public bool allowToolTip {
            get {
                return cp.Site.GetBoolean("Admin Navigator Allow Tooltip");
            }
        }
        // 
        private string emptyNodeListCacheKey {
            get {
                if (isDeveloper)
                    return "AdminNav EmptyNodeList Dev";
                if (isAdministrator)
                    return "AdminNav EmptyNodeList Admin";
                return "AdminNav EmptyNodeList CM" + cp.User.Id;
            }
        }
        // 
        // ==========================================================================================
        /// <summary>
        /// 
        /// </summary>
        public string emptyNodeList {
            get {
                try {
                    if (emptyNodeList_local is not null)
                        return emptyNodeList_local;
                    // 
                    // -- read cache version, return if not empty
                    emptyNodeList_local = cp.Cache.GetText(emptyNodeListCacheKey);
                    if (!string.IsNullOrEmpty(emptyNodeList_local))
                        return emptyNodeList_local;
                    // 
                    // -- build a new string, store in cache and return
                    string SQL = "select n.ID from ccMenuEntries n left join ccMenuEntries c on c.parentid=n.id Where c.ID Is Null group by n.id";
                    var cs2 = cp.CSNew();
                    if (cs2.OpenSQL(SQL)) {
                        do {
                            emptyNodeList_local += "," + cs2.GetText("id");
                            cs2.GoNext();
                        }
                        while (cs2.OK());
                        emptyNodeList_local = emptyNodeList_local.Substring(1);
                    }
                    cs2.Close();
                    cp.Cache.Store(emptyNodeListCacheKey, emptyNodeList_local, cacheDependencyList);
                    return emptyNodeList_local;
                } catch (Exception ex) {
                    cp.Site.ErrorReport(ex, $"hint emptyNodeList");
                    throw;
                }
            }
            set {
                cp.Cache.Store(emptyNodeListCacheKey, value, cacheDependencyList);
            }
        }
        private string emptyNodeList_local = null;
        // 
        // ==========================================================================================
        // 
        public string openNodeList {
            get {
                if (openNodeList_local is not null)
                    return openNodeList_local;
                openNodeList_local = cp.Visit.GetText("AdminNavOpenNodeList", "");
                return openNodeList_local;
            }
            set {
                cp.Visit.SetProperty("AdminNavOpenNodeList", value);
            }
        }
        private string openNodeList_local = null;
        // 
        // ==========================================================================================
        // 
        public bool autoManageAddons { get; set; } = false;
        // 
        // ==========================================================================================
        // 
        public bool adminNavOpen {
            get {
                return cp.Utils.EncodeBoolean(cp.Visit.GetText("AdminNavOpen", "1"));
            }
        }
        // 
        public bool allowAdminNavSaveState {
            get {
                return cp.Site.GetBoolean("adminNav allowSaveState", false);
            }
        }
        // 
        // ==========================================================================================
        // 
        public ApplicationEnvironmentModel(CPBaseClass cp) {
            this.cp = cp;
        }
        // 
        #region  IDisposable Support 
        protected bool disposed = false;
        // 
        // ==========================================================================================
        /// <summary>
        /// dispose
        /// </summary>
        /// <param name="disposing"></param>
        /// <remarks></remarks>
        protected virtual void Dispose(bool disposing) {
            if (!disposed) {
                if (disposing) {
                    // 
                    // ----- call .dispose for managed objects
                    // 
                }
                // 
                // Add code here to release the unmanaged resource.
                // 
            }
            disposed = true;
        }
        // Do not change or add Overridable to these methods.
        // Put cleanup code in Dispose(ByVal disposing As Boolean).
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        ~ApplicationEnvironmentModel() {
            Dispose(false);
        }
        #endregion
    }
}